/**
 * In-dashboard previews for Google Drive and Box files. Bytes come only through the app's own connections and
 * stay in memory. Images are shown from checked bytes as data: URLs (the only image source the app may use);
 * PDFs go to the built-in PDF viewer; plain text is decoded and shown as text. Everything else opens at the source.
 */
import { GoogleDriveService } from "../../generated/services/GoogleDriveService";
import { BoxService } from "../../generated/services/BoxService";

export type FileSourceKey = "googledrive" | "box";
export type PreviewKind = "image" | "pdf" | "text";

export const THUMB_MAX_BYTES = 3 * 1024 * 1024;
export const IMAGE_MAX_BYTES = 15 * 1024 * 1024;
export const TEXT_MAX_BYTES = 1024 * 1024;
export const TEXT_SHOWN_CHARS = 8000;

const IMAGE_EXT = /\.(png|jpe?g|gif|webp|bmp)$/i;
const TEXT_EXT = /\.(txt|csv|tsv|md|markdown|json|log|xml|ya?ml|ini|cfg)$/i;

export function previewKindOf(file: { name: string; mediaType?: string; isFolder: boolean }): PreviewKind | null {
  if (file.isFolder) return null;
  const type = (file.mediaType ?? "").toLowerCase();
  if (type === "application/pdf" || /\.pdf$/i.test(file.name)) return "pdf";
  if (/^image\/(png|jpe?g|gif|webp|bmp)$/.test(type) || IMAGE_EXT.test(file.name)) return "image";
  if (type.startsWith("text/") || type === "application/json" || TEXT_EXT.test(file.name)) return "text";
  return null;
}

/** Raw file content from the owning connection, unconverted. */
export async function loadFileContent(source: FileSourceKey, id: string): Promise<unknown> {
  const result = source === "box" ? await BoxService.GetFileContent(id, false) : await GoogleDriveService.GetFileContent(id, false);
  if (!result.success) {
    console.error("File preview load failed");
    throw new Error("Could not load the file.");
  }
  return result.data;
}

/** Normalizes the connector's content shapes (base64 text, a { $content } envelope, a binary string, bytes) to bytes. */
export function toBytes(value: unknown): Uint8Array | null {
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  if (typeof value === "object" && value !== null && "$content" in value) return toBytes((value as { $content: unknown }).$content);
  if (typeof value !== "string" || !value) return null;
  const compact = value.replace(/\s+/g, "");
  if (compact.length % 4 === 0 && /^[A-Za-z0-9+/]+={0,2}$/.test(compact)) {
    try {
      const binary = atob(compact);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      return bytes;
    } catch { /* not base64 after all */ }
  }
  const bytes = new Uint8Array(value.length);
  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i);
    if (code > 255) return null; // already decoded text, not bytes
    bytes[i] = code;
  }
  return bytes;
}

function imageType(bytes: Uint8Array): string | null {
  const b = bytes;
  if (b.length > 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return "image/png";
  if (b.length > 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return "image/jpeg";
  if (b.length > 6 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38) return "image/gif";
  if (b.length > 12 && b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return "image/webp";
  if (b.length > 2 && b[0] === 0x42 && b[1] === 0x4d) return "image/bmp";
  return null;
}

/** A data: URL for a recognised raster image, or null (SVG and anything unrecognised are never shown). */
export function imageDataUrl(value: unknown, maxBytes = IMAGE_MAX_BYTES): string | null {
  const bytes = toBytes(value);
  if (!bytes || bytes.length === 0 || bytes.length > maxBytes) return null;
  const type = imageType(bytes);
  if (!type) return null;
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  return `data:${type};base64,${btoa(binary)}`;
}

/** Plain-text content, or null when it isn't readable text. */
export function textOf(value: unknown): { text: string; truncated: boolean } | null {
  if (typeof value === "string" && [...value].some(ch => ch.charCodeAt(0) > 255)) {
    return { text: value.slice(0, TEXT_SHOWN_CHARS), truncated: value.length > TEXT_SHOWN_CHARS };
  }
  const bytes = toBytes(value);
  if (!bytes || bytes.length > TEXT_MAX_BYTES) return null;
  let text: string;
  try { text = new TextDecoder("utf-8", { fatal: true }).decode(bytes); } catch { return null; }
  if (text.includes(String.fromCharCode(0))) return null;
  return { text: text.slice(0, TEXT_SHOWN_CHARS), truncated: text.length > TEXT_SHOWN_CHARS };
}

export const EDIT_MAX_BYTES = 1024 * 1024;
/** The whole text of a file for editing, or null when it isn't editable text (or is over 1 MB). */
export function editableText(value: unknown): string | null {
  if (typeof value === "string" && [...value].some(ch => ch.charCodeAt(0) > 255)) return value.length <= EDIT_MAX_BYTES ? value : null;
  const bytes = toBytes(value);
  if (!bytes || bytes.length > EDIT_MAX_BYTES) return null;
  try {
    const text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    return text.includes(String.fromCharCode(0)) ? null : text;
  } catch { return null; }
}
