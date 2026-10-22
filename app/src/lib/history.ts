/**
 * History: every action you run, every Ask Copilot exchange and every document Copilot writes, so nothing is
 * lost on Refresh. Kept per user in this browser, plus a free copy in your OneDrive
 * (Mission Control Interstellar/Activity history.json, via the OneDrive connector — no Copilot Credits) so it survives
 * other browsers and devices. Only content you produced from the dashboard is kept.
 */
import { useSyncExternalStore } from "react";
import { OneDriveforBusinessService } from "../../generated/services/OneDriveforBusinessService";
import type { ActionLink } from "@/lib/actions";
import type { DocContent } from "@/lib/documents";

export type HistoryKind = "action" | "ask" | "document" | "created";
export interface HistoryEntry {
  id: string; at: number; kind: HistoryKind; title: string; section?: string; cost?: "free" | "credits";
  status: "done" | "failed"; notes?: string; choice?: string; message?: string; error?: string;
  links?: ActionLink[]; text?: string; doc?: DocContent; request?: string;
}
export type SyncState = "off" | "loading" | "saving" | "saved" | "local";
interface Store { entries: HistoryEntry[]; sync: SyncState; fileUrl?: string; folderUrl?: string }

const MAX_ENTRIES = 300, MAX_TEXT = 20000;
export const HISTORY_FOLDER = "/Mission Control Interstellar", HISTORY_FILE = "Activity history.json";
const HISTORY_PATH = `${HISTORY_FOLDER}/${HISTORY_FILE}`;
let owner: string | undefined;
let oneDriveHome: string | undefined;
/** A browser link to the history file in OneDrive (built from your OneDrive address; the connector returns no link). */
function fileLink(path?: string) {
  if (!oneDriveHome) return undefined;
  const clean = (path || HISTORY_PATH).split("/").filter(Boolean).map(encodeURIComponent).join("/");
  return `${oneDriveHome.replace(/\/+$/, "")}/Documents/${clean}`;
}
let store: Store = { entries: [], sync: "off" };
let fileId: string | undefined;
let syncBroken = false;
let saveTimer: ReturnType<typeof setTimeout> | undefined;
const listeners = new Set<() => void>();
const emit = () => listeners.forEach(listener => listener());
const set = (next: Partial<Store>) => { store = { ...store, ...next }; emit(); };

function storageKey(id?: string) { return id && /^[0-9a-f-]{36}$/i.test(id) ? `mission-control-interstellar:history:v1:${id.toLowerCase()}` : undefined; }
const clip = (value?: string) => value && value.length > MAX_TEXT ? `${value.slice(0, MAX_TEXT)}…` : value;

function isEntry(value: unknown): value is HistoryEntry {
  if (!value || typeof value !== "object") return false;
  const row = value as Record<string, unknown>;
  return typeof row.id === "string" && typeof row.at === "number" && typeof row.title === "string" && typeof row.kind === "string";
}
export function parseHistory(raw: unknown): HistoryEntry[] {
  let value = raw;
  if (typeof value === "string") { try { value = JSON.parse(value); } catch { return []; } }
  const list = Array.isArray(value) ? value : value && typeof value === "object" && Array.isArray((value as { entries?: unknown }).entries) ? (value as { entries: unknown[] }).entries : [];
  return list.filter(isEntry);
}
/** Newest first, one copy per id (the newer copy wins), capped. */
export function mergeHistory(...lists: HistoryEntry[][]): HistoryEntry[] {
  const byId = new Map<string, HistoryEntry>();
  for (const entry of lists.flat()) { const existing = byId.get(entry.id); if (!existing || entry.at >= existing.at) byId.set(entry.id, entry); }
  return [...byId.values()].sort((a, b) => b.at - a.at).slice(0, MAX_ENTRIES);
}
function saveLocal() {
  try { const key = storageKey(owner); if (key) localStorage.setItem(key, JSON.stringify(store.entries)); }
  catch {
    // Browser storage full: keep the newest half locally (OneDrive keeps the full copy).
    try { const key = storageKey(owner); if (key) localStorage.setItem(key, JSON.stringify(store.entries.slice(0, Math.floor(store.entries.length / 2)))); } catch { /* memory only */ }
  }
}

// ---------- OneDrive copy (free) ----------
function contentText(data: unknown): string | undefined {
  if (typeof data === "string") return data;
  if (data && typeof data === "object") {
    const record = data as Record<string, unknown>;
    if (typeof record.$content === "string") { try { return new TextDecoder().decode(Uint8Array.from(atob(record.$content), char => char.charCodeAt(0))); } catch { return undefined; } }
    if (Array.isArray(record.entries)) return JSON.stringify(record);
  }
  return undefined;
}
async function loadRemote(): Promise<HistoryEntry[]> {
  const meta = await OneDriveforBusinessService.GetFileMetadataByPath(HISTORY_PATH);
  if (!meta.success || !meta.data?.Id) return [];
  fileId = meta.data.Id;
  set({ fileUrl: fileLink(meta.data.Path) });
  const content = await OneDriveforBusinessService.GetFileContentByPath(HISTORY_PATH, false);
  if (!content.success) throw new Error("History file unreadable");
  return parseHistory(contentText(content.data) ?? "[]");
}
async function saveRemote() {
  if (syncBroken || !owner) return;
  set({ sync: "saving" });
  const body = JSON.stringify({ app: "Mission Control Interstellar", version: 1, savedAt: new Date().toISOString(), entries: store.entries }, null, 1);
  try {
    if (fileId) {
      const updated = await OneDriveforBusinessService.UpdateFile(fileId, body);
      if (!updated.success) throw updated.error ?? new Error("update failed");
    } else {
      const created = await OneDriveforBusinessService.CreateFile(HISTORY_FOLDER, HISTORY_FILE, body);
      if (!created.success || !created.data?.Id) throw created.error ?? new Error("create failed");
      fileId = created.data.Id;
      set({ fileUrl: fileLink(created.data.Path) });
    }
    set({ sync: "saved" });
  } catch {
    syncBroken = true; // stop trying this session; the browser copy still works
    set({ sync: "local" });
  }
}
function scheduleSave() {
  saveLocal();
  if (saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(() => { void saveRemote(); }, 2500);
}

/** Called once the signed-in user's profile id is known: loads this browser's copy, then merges the OneDrive copy. */
export async function initHistory(id?: string, mySite?: string) {
  if (mySite) oneDriveHome = /^https:\/\/[^/]+-my\.sharepoint\.com\//i.test(mySite) ? mySite : undefined;
  if (!id || id === owner) return;
  owner = id; fileId = undefined; syncBroken = false;
  let local: HistoryEntry[] = [];
  try { const key = storageKey(id); local = parseHistory(key ? localStorage.getItem(key) : null); } catch { /* none */ }
  set({ entries: mergeHistory(local, store.entries), sync: "loading" });
  try {
    const remote = await loadRemote();
    const merged = mergeHistory(store.entries, remote);
    const changed = merged.length !== remote.length || merged.some((entry, i) => entry.id !== remote[i]?.id);
    set({ entries: merged, sync: "saved" });
    saveLocal();
    if (changed && merged.length) scheduleSave();
  } catch {
    syncBroken = true;
    set({ sync: "local" });
  }
}

export function recordHistory(entry: Omit<HistoryEntry, "id" | "at"> & { at?: number }) {
  const full: HistoryEntry = { ...entry, id: `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`, at: entry.at ?? Date.now(), text: clip(entry.text), message: clip(entry.message) };
  set({ entries: mergeHistory([full], store.entries) });
  scheduleSave();
  return full.id;
}
export function deleteHistory(id: string) { set({ entries: store.entries.filter(entry => entry.id !== id) }); scheduleSave(); }
export function clearHistory() { set({ entries: [] }); scheduleSave(); }
export function retryHistorySync() { syncBroken = false; void saveRemote(); }
export function useHistory() {
  return useSyncExternalStore(listener => { listeners.add(listener); return () => listeners.delete(listener); }, () => store);
}

/** Plain text of an entry, for searching. */
export function entryText(entry: HistoryEntry): string {
  const doc = entry.doc ? [entry.doc.title, ...entry.doc.sections.flatMap(section => [section.heading, section.text])].join(" ") : "";
  return [entry.title, entry.section, entry.notes, entry.request, entry.message, entry.error, entry.text, doc, ...(entry.links ?? []).map(link => link.label)].filter(Boolean).join(" ").toLowerCase();
}
export function searchHistory(entries: HistoryEntry[], query: string, kind: "all" | HistoryKind | "failed" = "all"): HistoryEntry[] {
  const terms = query.toLowerCase().split(/\s+/).filter(Boolean);
  return entries.filter(entry => (kind === "all" || (kind === "failed" ? entry.status === "failed" : entry.kind === kind)) && terms.every(term => entryText(entry).includes(term)));
}
/** For tests only. */
export function resetHistoryForTests() { owner = undefined; fileId = undefined; syncBroken = false; store = { entries: [], sync: "off" }; }
