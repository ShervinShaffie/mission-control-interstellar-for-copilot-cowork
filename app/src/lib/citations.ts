/**
 * Turns Copilot-written text (markdown-ish, with [1](https://very-long-url) citations) into clean blocks
 * with short, named source links, plus one numbered "Sources" list per document.
 */
import { sourceName } from "@/lib/source-links";
import type { DocContent } from "@/lib/documents";

export interface Cite { n: number; name: string; url: string }
export type Inline = { kind: "text"; text: string } | { kind: "cites"; cites: Cite[] };
export interface Block { kind: "para" | "bullet"; inlines: Inline[] }
export interface CleanSection { heading: string; blocks: Block[] }
export interface CleanDoc { title: string; sections: CleanSection[]; sources: Cite[] }

const LINK = /\[([^\]]*)\]\((https?:\/\/[^\s)]+(?:\([^\s)]*\)[^\s)]*)*)\)|(https?:\/\/[^\s<>"')\]]+)/gi;
const MAX_CHIP = 34;

/** Removes markdown noise Copilot sometimes adds: escaped brackets, heading hashes, bold markers. */
export function cleanText(value: string): string {
  return (value ?? "")
    .replace(/\\([[\]()_*#])/g, "$1")
    .replace(/^\s{0,3}#{1,6}\s+/gm, "")
    .replace(/\*\*([^*]+)\*\*/g, "$1").replace(/__([^_]+)__/g, "$1")
    .replace(/\r\n/g, "\n");
}
export function chipName(name: string) { return name.length > MAX_CHIP ? `${name.slice(0, MAX_CHIP - 1).trimEnd()}…` : name; }

/** Shared numbering for one document: the same URL always gets the same number and name. */
export class SourceIndex {
  readonly list: Cite[] = [];
  private byUrl = new Map<string, Cite>();
  add(url: string, label: string): Cite {
    const key = url.replace(/[.,;:]+$/, "");
    const existing = this.byUrl.get(key);
    const name = sourceName(key, label);
    if (existing) { if (/^(Teams message|Email|Source|SharePoint|OneDrive file|Teams chat)$/.test(existing.name) && !/^\d+$/.test(label.trim()) && label.trim()) existing.name = name; return existing; }
    const cite = { n: this.list.length + 1, name, url: key };
    this.byUrl.set(key, cite); this.list.push(cite);
    return cite;
  }
}

function inlines(line: string, index: SourceIndex): Inline[] {
  const out: Inline[] = [];
  let at = 0;
  for (const match of line.matchAll(LINK)) {
    const start = match.index ?? 0;
    const before = line.slice(at, start);
    const cite = index.add(match[2] ?? match[3], match[2] ? match[1] : "");
    const last = out.at(-1);
    // Citations next to each other (only spaces or commas between) form one group; duplicates collapse.
    if (last?.kind === "cites" && /^[\s,;]*$/.test(before)) { if (!last.cites.some(item => item.url === cite.url)) last.cites.push(cite); }
    else {
      if (before) out.push({ kind: "text", text: before });
      out.push({ kind: "cites", cites: [cite] });
    }
    at = start + match[0].length;
  }
  if (at < line.length) out.push({ kind: "text", text: line.slice(at) });
  // Tidy spaces before a citation group and before punctuation that follows it.
  return out.map((item, i) => item.kind === "text" && out[i + 1]?.kind === "cites" ? { ...item, text: item.text.replace(/\s+$/, " ") } : item)
    .filter(item => item.kind !== "text" || item.text.length > 0);
}

export function parseBlocks(text: string, index: SourceIndex): Block[] {
  const blocks: Block[] = [];
  for (const raw of cleanText(text).split(/\n/)) {
    const line = raw.trim();
    if (!line) continue;
    const bullet = line.match(/^(?:[-*•]|\d+[.)])\s+(.*)$/);
    blocks.push({ kind: bullet ? "bullet" : "para", inlines: inlines(bullet ? bullet[1] : line, index) });
  }
  return blocks;
}

const norm = (value: string) => value.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
export function cleanDoc(doc: DocContent): CleanDoc {
  const index = new SourceIndex();
  const sections = doc.sections.map((section, i) => ({
    // Drop a first heading that just repeats the title.
    heading: i === 0 && norm(cleanText(section.heading)) === norm(doc.title) ? "" : cleanText(section.heading).trim(),
    blocks: parseBlocks(section.text, index),
  })).filter(section => section.heading || section.blocks.length);
  return { title: cleanText(doc.title).trim(), sections, sources: index.list };
}
/** Plain text with citations as "(Name)" — for Excel cells and copying. */
export function plainText(blocks: Block[]): string {
  return blocks.map(block => (block.kind === "bullet" ? "• " : "") + block.inlines.map(item => item.kind === "text" ? item.text : ` (${item.cites.map(cite => cite.name).join("; ")})`).join("").replace(/\s+\(/g, " (").trim()).join("\n");
}
