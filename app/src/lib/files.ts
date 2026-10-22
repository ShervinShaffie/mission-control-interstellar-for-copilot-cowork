/**
 * Files from data the dashboard already reads (no extra services): links and attachments in your email,
 * Teams posts and chats, and meeting invites, plus documents trending around you. Ranked by what needs you.
 */
import { hideEmails } from "@/lib/privacy";

export interface FileRef { name: string; url: string }
export type FileKind = "word" | "excel" | "powerpoint" | "pdf" | "onenote" | "loop" | "other";
export type FileOrigin = "reply" | "chat" | "meeting" | "mention" | "team" | "email" | "trending";
export interface FileHit { file: FileRef; origin: FileOrigin; who?: string; when?: Date; context: string; contextUrl?: string; meetingAt?: Date }
export interface FileItem {
  key: string; name: string; url: string; kind: FileKind; score: number;
  reasons: string[]; contexts: { label: string; url?: string }[]; who?: string; when?: Date; origins: FileOrigin[];
}

const FILE_HOST = /(^|\.)sharepoint\.com$|(^|\.)1drv\.ms$|(^|\.)onedrive\.live\.com$|(^|\.)loop\.microsoft\.com$/i;
const SHARE_KIND: Record<string, FileKind> = { w: "word", x: "excel", p: "powerpoint", b: "pdf", o: "onenote", l: "loop", f: "other" };
const EXT_KIND: [RegExp, FileKind][] = [[/\.(docx?|docm|dotx|rtf|odt)$/i, "word"], [/\.(xlsx?|xlsm|xlsb|csv)$/i, "excel"], [/\.(pptx?|pptm|ppsx)$/i, "powerpoint"], [/\.pdf$/i, "pdf"], [/\.one(note)?$/i, "onenote"], [/\.(loop|fluid)$/i, "loop"]];
const KIND_LABEL: Record<FileKind, string> = { word: "Word document", excel: "Excel workbook", powerpoint: "PowerPoint deck", pdf: "PDF", onenote: "OneNote", loop: "Loop page", other: "File" };

function safeHttps(url: string): URL | undefined {
  try { const parsed = new URL(url.replace(/&amp;/g, "&")); return parsed.protocol === "https:" ? parsed : undefined; } catch { return undefined; }
}
function decode(value: string) { try { return decodeURIComponent(value); } catch { return value; } }

/** What kind of file a name or link points to. */
export function fileKind(name: string, url = ""): FileKind {
  for (const [pattern, kind] of EXT_KIND) if (pattern.test(name)) return kind;
  const parsed = safeHttps(url);
  const share = parsed?.pathname.match(/^\/:([a-z]):\//i)?.[1]?.toLowerCase();
  if (share && SHARE_KIND[share]) return SHARE_KIND[share];
  const file = parsed?.searchParams.get("file") ?? parsed?.pathname.split("/").pop() ?? "";
  for (const [pattern, kind] of EXT_KIND) if (pattern.test(decode(file))) return kind;
  return "other";
}
export function kindLabel(kind: FileKind) { return KIND_LABEL[kind]; }

/** True when a link points at a document (not a site home page, a folder view or a Teams link). */
export function isFileLink(url: string): boolean {
  const parsed = safeHttps(url);
  if (!parsed || !FILE_HOST.test(parsed.hostname)) return false;
  if (/^\/:([wxpbol]):\//i.test(parsed.pathname)) return true;
  if (parsed.hostname.endsWith("1drv.ms") || parsed.hostname.endsWith("loop.microsoft.com")) return true;
  if (/\/Doc\.aspx$|\/WopiFrame/i.test(parsed.pathname) && (parsed.searchParams.get("file") || parsed.searchParams.get("sourcedoc"))) return true;
  return EXT_KIND.some(([pattern]) => pattern.test(decode(parsed.pathname)));
}

/** A readable name for a file link when the link text is just the address. */
export function nameFromUrl(url: string): string {
  const parsed = safeHttps(url);
  if (!parsed) return "File";
  const file = parsed.searchParams.get("file");
  if (file) return decode(file);
  const last = decode(parsed.pathname.split("/").filter(Boolean).pop() ?? "");
  if (last && /\.[a-z0-9]{2,5}$/i.test(last)) return last;
  return kindLabel(fileKind("", url));
}

function cleanName(text: string, url: string) {
  const name = hideEmails(text.replace(/<[^>]+>/g, " ").replace(/&amp;/g, "&").replace(/&nbsp;/g, " ").replace(/\s+/g, " ").trim());
  return !name || /^https?:\/\//i.test(name) || name.length > 160 ? nameFromUrl(url) : name;
}

/** File links in an email, invite or post body (HTML or plain text). */
export function filesFromHtml(html?: string): FileRef[] {
  if (!html) return [];
  const out: FileRef[] = [];
  for (const match of html.matchAll(/<a\b[^>]*href\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi)) {
    const url = match[1].replace(/&amp;/g, "&");
    if (isFileLink(url)) out.push({ url, name: cleanName(match[2], url) });
  }
  const text = html.replace(/<a\b[\s\S]*?<\/a>/gi, " ").replace(/<[^>]+>/g, " ");
  for (const match of text.matchAll(/https:\/\/[^\s"'<>)\]]+/g)) {
    const url = match[0].replace(/&amp;/g, "&").replace(/[.,;:]+$/, "");
    if (isFileLink(url)) out.push({ url, name: nameFromUrl(url) });
  }
  return dedupeRefs(out);
}

/** Files attached to a Teams message (contentType "reference"). */
export function filesFromAttachments(attachments: unknown): FileRef[] {
  if (!Array.isArray(attachments)) return [];
  return dedupeRefs(attachments.flatMap(entry => {
    const item = entry && typeof entry === "object" ? entry as Record<string, unknown> : {};
    const url = typeof item.contentUrl === "string" ? item.contentUrl : "";
    const name = typeof item.name === "string" ? item.name : "";
    return item.contentType === "reference" && safeHttps(url) ? [{ url, name: cleanName(name, url) }] : [];
  }));
}

function dedupeRefs(refs: FileRef[]) {
  const seen = new Set<string>();
  return refs.filter(ref => { const key = fileKey(ref); if (seen.has(key)) return false; seen.add(key); return true; }).slice(0, 10);
}
/** Same file, even when shared through different links. */
export function fileKey(ref: FileRef) {
  const parsed = safeHttps(ref.url);
  const id = parsed?.searchParams.get("sourcedoc") ?? parsed?.searchParams.get("sourceDoc");
  if (id) return `doc:${id.toLowerCase()}`;
  return `name:${ref.name.trim().toLowerCase()}`;
}

const ORIGIN: Record<FileOrigin, { score: number; reason: (hit: FileHit) => string }> = {
  reply: { score: 45, reason: hit => `In an email waiting on your reply${hit.who ? ` from ${hit.who}` : ""}` },
  chat: { score: 40, reason: hit => `Sent in a chat waiting on you${hit.who ? ` (${hit.who})` : ""}` },
  meeting: { score: 38, reason: hit => `Linked in the invite for ${hit.context}` },
  mention: { score: 35, reason: hit => `Shared in a post that mentions you${hit.who ? ` (${hit.who})` : ""}` },
  team: { score: 22, reason: hit => `Shared in ${hit.context}${hit.who ? ` by ${hit.who}` : ""}` },
  email: { score: 18, reason: hit => `In an email${hit.who ? ` from ${hit.who}` : ""}` },
  trending: { score: 12, reason: hit => `Trending around you${hit.context ? ` · ${hit.context}` : ""}` },
};

/** One ranked list: the same file found in several places becomes one row with every reason. */
export function rankFiles(hits: FileHit[], now = new Date()): FileItem[] {
  const byKey = new Map<string, FileItem>();
  const weights = new Map<string, Map<string, number>>();
  for (const hit of hits) {
    const key = fileKey(hit.file);
    const origin = ORIGIN[hit.origin];
    const age = hit.when ? (now.getTime() - hit.when.getTime()) / 864e5 : 30;
    let score = origin.score + Math.max(0, 10 - Math.max(0, age));
    if (hit.meetingAt) { const hours = (hit.meetingAt.getTime() - now.getTime()) / 36e5; if (hours >= 0 && hours < 24) score += 8; }
    const reason = origin.reason(hit);
    const existing = byKey.get(key);
    if (!existing) {
      weights.set(key, new Map([[reason, origin.score]]));
      byKey.set(key, { key, name: hit.file.name, url: hit.file.url, kind: fileKind(hit.file.name, hit.file.url), score, reasons: [reason], contexts: [{ label: hit.context, url: hit.contextUrl }], who: hit.who, when: hit.when, origins: [hit.origin] });
      continue;
    }
    const best = Math.max(existing.score, score);
    existing.score = best + 4;
    if (!existing.reasons.includes(reason)) {
      existing.reasons.push(reason);
      // Most important reason first.
      const weight = weights.get(key)!; weight.set(reason, origin.score);
      existing.reasons.sort((a, b) => (weight.get(b) ?? 0) - (weight.get(a) ?? 0));
    }
    if (!existing.contexts.some(context => context.label === hit.context)) existing.contexts.push({ label: hit.context, url: hit.contextUrl });
    if (!existing.origins.includes(hit.origin)) existing.origins.push(hit.origin);
    if (hit.when && (!existing.when || hit.when > existing.when)) { existing.when = hit.when; existing.who = hit.who ?? existing.who; }
    if (existing.name === kindLabel(existing.kind) && hit.file.name !== kindLabel(existing.kind)) existing.name = hit.file.name;
  }
  return [...byKey.values()].sort((a, b) => b.score - a.score || (b.when?.getTime() ?? 0) - (a.when?.getTime() ?? 0));
}

/** Markdown links for a short list of files ("[Plan.docx](url), [Budget.xlsx](url)"). */
export function fileLinks(refs: FileRef[], max = 3) {
  const list = refs.slice(0, max).map(ref => `[${ref.name.replace(/[[\]]/g, "")}](${ref.url})`);
  return refs.length > max ? `${list.join(", ")} and ${refs.length - max} more` : list.join(", ");
}
