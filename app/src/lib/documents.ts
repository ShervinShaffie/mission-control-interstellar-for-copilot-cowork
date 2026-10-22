import * as XLSX from "xlsx";
import { chipName, cleanDoc, plainText, type Block } from "@/lib/citations";

export interface DocSection { heading: string; text: string }
export interface DocContent { title: string; sections: DocSection[] }
export type DocFormat = "docx" | "xlsx" | "pdf";
export const FORMAT_LABELS: Record<DocFormat, string> = { docx: "Word", xlsx: "Excel", pdf: "PDF" };

// ---------- minimal store-only ZIP writer (for .docx) ----------
const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();
export function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
export function zipStore(files: { name: string; data: Uint8Array }[]): Uint8Array {
  const encoder = new TextEncoder();
  const chunks: Uint8Array[] = [], central: Uint8Array[] = [];
  let offset = 0;
  for (const file of files) {
    const name = encoder.encode(file.name), crc = crc32(file.data), size = file.data.length;
    const local = new Uint8Array(30 + name.length), view = new DataView(local.buffer);
    view.setUint32(0, 0x04034b50, true); view.setUint16(4, 20, true); view.setUint16(6, 0x0800, true); view.setUint16(8, 0, true);
    view.setUint16(10, 0, true); view.setUint16(12, 0x21, true); view.setUint32(14, crc, true); view.setUint32(18, size, true); view.setUint32(22, size, true);
    view.setUint16(26, name.length, true); view.setUint16(28, 0, true); local.set(name, 30);
    const entry = new Uint8Array(46 + name.length), cv = new DataView(entry.buffer);
    cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true); cv.setUint16(8, 0x0800, true); cv.setUint16(10, 0, true);
    cv.setUint16(12, 0, true); cv.setUint16(14, 0x21, true); cv.setUint32(16, crc, true); cv.setUint32(20, size, true); cv.setUint32(24, size, true);
    cv.setUint16(28, name.length, true); cv.setUint32(42, offset, true); entry.set(name, 46);
    chunks.push(local, file.data); central.push(entry);
    offset += local.length + size;
  }
  const centralSize = central.reduce((sum, item) => sum + item.length, 0);
  const end = new Uint8Array(22), ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, files.length, true); ev.setUint16(10, files.length, true);
  ev.setUint32(12, centralSize, true); ev.setUint32(16, offset, true);
  const all = [...chunks, ...central, end], out = new Uint8Array(all.reduce((sum, item) => sum + item.length, 0));
  let at = 0; for (const part of all) { out.set(part, at); at += part.length; }
  return out;
}

function xml(text: string) {
  // Drop characters XML 1.0 does not allow (control characters other than tab, newline, return).
  const allowed = Array.from(text).filter(char => { const code = char.charCodeAt(0); return code >= 32 || code === 9 || code === 10 || code === 13; }).join("");
  return allowed.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function paragraph(text: string, style?: string) {
  const runs = text.split("\n").map((line, index) => `${index ? "<w:br/>" : ""}<w:t xml:space="preserve">${xml(line)}</w:t>`).join("");
  return `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}<w:r>${runs}</w:r></w:p>`;
}
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:b/><w:color w:val="005A9E"/><w:sz w:val="48"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="005A9E"/><w:sz w:val="30"/></w:rPr></w:style></w:styles>`;
const LINK_RUN = `<w:rPr><w:color w:val="0563C1"/><w:sz w:val="18"/><w:u w:val="single"/></w:rPr>`;
function textRun(text: string) { return `<w:r><w:t xml:space="preserve">${xml(text)}</w:t></w:r>`; }
function linkRun(text: string, rel: string) { return `<w:hyperlink r:id="${rel}" w:history="1"><w:r>${LINK_RUN}<w:t xml:space="preserve">${xml(text)}</w:t></w:r></w:hyperlink>`; }
/** A real Word (.docx) document: title, headings, paragraphs and bullets, with short clickable source names and a Sources list. */
export function buildDocx(content: DocContent): Uint8Array {
  const encoder = new TextEncoder();
  const doc = cleanDoc(content);
  const rel = (n: number) => `rIdS${n}`;
  const block = (b: Block) => {
    const runs = b.inlines.map(item => item.kind === "text" ? textRun(item.text)
      : item.cites.map((cite, i) => `${i ? textRun(" ") : textRun(" ")}${linkRun(chipName(cite.name), rel(cite.n))}`).join("")).join("");
    return `<w:p>${b.kind === "bullet" ? `<w:pPr><w:ind w:left="360" w:hanging="216"/></w:pPr>${textRun("•\u00a0")}` : ""}${runs}</w:p>`;
  };
  const body = [paragraph(doc.title, "Title"), ...doc.sections.flatMap(section => [
    ...(section.heading ? [paragraph(section.heading, "Heading1")] : []),
    ...section.blocks.map(block),
  ]), ...(doc.sources.length ? [paragraph("Sources", "Heading1"), ...doc.sources.map(cite => `<w:p>${textRun(`${cite.n}. `)}<w:hyperlink r:id="${rel(cite.n)}" w:history="1"><w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">${xml(cite.name)}</w:t></w:r></w:hyperlink></w:p>`)] : [])].join("");
  const links = doc.sources.map(cite => `<Relationship Id="${rel(cite.n)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xml(cite.url)}" TargetMode="External"/>`).join("");
  const files = {
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${links}</Relationships>`,
    "word/styles.xml": STYLES,
    "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  };
  return zipStore(Object.entries(files).map(([name, text]) => ({ name, data: encoder.encode(text) })));
}

/** An Excel workbook: one sheet per table (first row = headers). */
export function buildXlsxBase64(sheets: { name: string; rows: (string | number)[][] }[]): string {
  const book = XLSX.utils.book_new();
  for (const sheet of sheets) {
    const ws = XLSX.utils.aoa_to_sheet(sheet.rows);
    ws["!cols"] = (sheet.rows[0] ?? []).map((_, column) => ({ wch: Math.min(60, Math.max(10, ...sheet.rows.map(row => String(row[column] ?? "").length))) }));
    XLSX.utils.book_append_sheet(book, ws, sheet.name.replace(/[[\]*?/\\:]/g, " ").slice(0, 31) || "Sheet1");
  }
  return XLSX.write(book, { bookType: "xlsx", type: "base64" }) as string;
}
/** Sections as a single Excel sheet (Heading | Content). */
export function sectionsToSheet(content: DocContent): (string | number)[][] {
  const doc = cleanDoc(content);
  return [["Section", "Content"], ...doc.sections.map(section => [section.heading, plainText(section.blocks)]), ...(doc.sources.length ? [["", ""], ["Sources", ""], ...doc.sources.map(cite => [`${cite.n}. ${cite.name}`, cite.url])] : [])];
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let index = 0; index < bytes.length; index += 0x8000) binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
  return btoa(binary);
}
export async function blobToBase64(blob: Blob): Promise<string> {
  return bytesToBase64(new Uint8Array(await blob.arrayBuffer()));
}
/** A safe file name: letters, numbers, spaces, dashes; with the date and extension. */
export function fileName(title: string, format: DocFormat): string {
  const base = title.replace(/[^\p{L}\p{N} _-]+/gu, " ").replace(/\s+/g, " ").trim().slice(0, 80) || "Document";
  return `${base} ${new Date().toISOString().slice(0, 10)}.${format}`;
}
