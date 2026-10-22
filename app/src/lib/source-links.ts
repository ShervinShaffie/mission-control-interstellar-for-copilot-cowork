// Turns citation text such as "metadata [15](https://…/Doc.aspx?file=Plan.xlsx…)"
// into short, readable source names instead of long URLs.
export type RichSegment = { kind: "text"; text: string } | { kind: "link"; text: string; url: string };

const MARKDOWN_LINK = /\[([^\]]*)\]\((https?:\/\/[^\s)]+)\)/gi;
const BARE_URL = /https?:\/\/[^\s<>"')\]]+/gi;
const MAX_NAME = 60;

function unescapeMarkdown(value: string) {
  return value.replace(/\\([_*[\]()~`#+\-.!\\])/g, "$1");
}
function decode(value: string) {
  try { return decodeURIComponent(value.replace(/\+/g, " ")); } catch { return value; }
}
function shorten(value: string) {
  const text = value.replace(/\s+/g, " ").trim();
  return text.length > MAX_NAME ? `${text.slice(0, MAX_NAME - 1).trimEnd()}…` : text;
}
function stripExtension(name: string) {
  return name.replace(/\.(docx?|xlsx?|pptx?|pdf|txt|csv|one|loop|vsdx?|md|aspx)$/i, "");
}

/** A short human name for a source link. */
export function sourceName(url: string, label = ""): string {
  const cleanLabel = unescapeMarkdown(label).trim();
  if (cleanLabel && !/^\d+$/.test(cleanLabel) && !/^https?:\/\//i.test(cleanLabel) && cleanLabel.length <= MAX_NAME) return cleanLabel;
  let parsed: URL;
  try { parsed = new URL(unescapeMarkdown(url)); } catch { return "Source"; }
  const host = parsed.hostname.toLowerCase();
  const file = parsed.searchParams.get("file");
  if (file) return shorten(stripExtension(decode(file)));
  if (host.startsWith("outlook.") || host.endsWith("outlook.com")) return host.includes("calendar") || parsed.pathname.includes("calendar") || /calendar|event/i.test(parsed.searchParams.get("viewmodel") ?? "") ? "Calendar event" : "Email";
  if (host === "teams.microsoft.com" || host === "teams.cloud.microsoft") {
    const team = parsed.searchParams.get("teamName"), channel = parsed.searchParams.get("channelName");
    if (team || channel) return shorten([team, channel].filter(Boolean).map(value => decode(value!)).join(" › "));
    if (/meetup-join|\/l\/meeting\//.test(parsed.pathname)) return "Meeting";
    if (/contextType.{0,8}chat/i.test(decode(parsed.search)) || /@thread\.v2|@unq\.gbl/.test(decode(parsed.pathname))) return "Teams chat";
    return "Teams message";
  }
  const segments = parsed.pathname.split("/").filter(Boolean).map(decode).filter(part => !/^(_layouts|15|Doc\.aspx|r|s|personal|sites|Forms|AllItems\.aspx|:[a-z]:)$/i.test(part));
  const last = segments.at(-1);
  if (last && !/^[0-9a-f-]{20,}$/i.test(last)) return shorten(stripExtension(last));
  if (host.includes("sharepoint.com")) return host.includes("-my.") ? "OneDrive file" : "SharePoint";
  return shorten(host.replace(/^www\./, ""));
}

/** Splits text into plain text and short-named links. */
export function parseRichText(value: string): RichSegment[] {
  // Copilot sometimes escapes the brackets of its citations (\\[1\\](url)); unescape them so they become links.
  const text = (value ?? "").replace(/\\([[\]()])/g, "$1");
  const segments: RichSegment[] = [];
  const pattern = new RegExp(`${MARKDOWN_LINK.source}|${BARE_URL.source}`, "gi");
  let index = 0;
  for (const match of text.matchAll(pattern)) {
    const start = match.index ?? 0;
    if (start > index) segments.push({ kind: "text", text: text.slice(index, start) });
    const url = unescapeMarkdown(match[2] ?? match[0]).replace(/[.,;:]+$/, "");
    segments.push({ kind: "link", url, text: sourceName(url, match[2] ? match[1] : "") });
    index = start + match[0].length;
  }
  if (index < text.length) segments.push({ kind: "text", text: text.slice(index) });
  return segments.map(segment => segment.kind === "text" ? { ...segment, text: unescapeMarkdown(segment.text) } : segment);
}

/** The first link in the text, used when a record has no URL of its own. */
export function firstLinkUrl(value: string): string | undefined {
  return parseRichText(value).find((segment): segment is Extract<RichSegment, { kind: "link" }> => segment.kind === "link")?.url;
}
