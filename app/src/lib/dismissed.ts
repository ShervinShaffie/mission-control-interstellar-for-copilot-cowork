/**
 * Dismissed suggestions, kept per user in this browser (no business data beyond the suggestion's title).
 * A dismissed item stays in a "Dismissed" list for the chosen number of days so it can be restored; after
 * that it is removed for good (it stays hidden, and no longer appears in the list).
 */
import { useSyncExternalStore } from "react";

export type RetentionDays = 1 | 3 | 7;
export const RETENTION_CHOICES: RetentionDays[] = [1, 3, 7];
export interface DismissedItem { key: string; title: string; section: string; at: number }
interface Store { retention: RetentionDays; items: DismissedItem[]; hidden: string[] }

const DAY = 864e5, MAX_HIDDEN = 400, MAX_ITEMS = 200;
const EMPTY: Store = { retention: 7, items: [], hidden: [] };
let owner: string | undefined;
let store: Store = EMPTY;
const listeners = new Set<() => void>();

/** The same suggestion can appear on Today, in Projects and in its own component: one key covers all. */
export function dismissKey(id: string) { return id.replace(/^(pa-|att-)+/, ""); }

function storageKey(id?: string) { return id && /^[0-9a-f-]{36}$/i.test(id) ? `mission-control-interstellar:dismissed:v1:${id.toLowerCase()}` : undefined; }
function parse(raw: string | null): Store {
  try {
    const value: unknown = raw ? JSON.parse(raw) : undefined;
    if (!value || typeof value !== "object") return { ...EMPTY };
    const record = value as Record<string, unknown>;
    const retention = RETENTION_CHOICES.includes(record.retention as RetentionDays) ? record.retention as RetentionDays : 7;
    const items = Array.isArray(record.items) ? record.items.flatMap((item: unknown): DismissedItem[] => {
      if (!item || typeof item !== "object") return [];
      const row = item as Record<string, unknown>;
      return typeof row.key === "string" && typeof row.title === "string" && typeof row.section === "string" && typeof row.at === "number" ? [{ key: row.key, title: row.title.slice(0, 200), section: row.section, at: row.at }] : [];
    }) : [];
    const hidden = Array.isArray(record.hidden) ? record.hidden.filter((key): key is string => typeof key === "string") : [];
    return { retention, items, hidden };
  } catch { return { ...EMPTY }; }
}
/** Moves items older than the retention window out of the list (they stay hidden for good). */
function expire(value: Store, now = Date.now()): Store {
  const cutoff = now - value.retention * DAY;
  const keep = value.items.filter(item => item.at >= cutoff), gone = value.items.filter(item => item.at < cutoff);
  if (!gone.length) return value;
  return { ...value, items: keep, hidden: [...new Set([...value.hidden, ...gone.map(item => item.key)])].slice(-MAX_HIDDEN) };
}
function save(next: Store) {
  store = next;
  try { const key = storageKey(owner); if (key) localStorage.setItem(key, JSON.stringify(next)); } catch { /* keep in memory */ }
  listeners.forEach(listener => listener());
}

/** Called once the signed-in user's profile id is known. */
export function setDismissOwner(id?: string) {
  if (id === owner) return;
  owner = id;
  let raw: string | null = null;
  try { const key = storageKey(id); raw = key ? localStorage.getItem(key) : null; } catch { /* in memory only */ }
  const loaded = parse(raw), expired = expire(loaded);
  store = expired;
  if (expired !== loaded) save(expired); else listeners.forEach(listener => listener());
}
export function dismiss(id: string, title: string, section: string) {
  const key = dismissKey(id);
  save({ ...store, items: [{ key, title, section, at: Date.now() }, ...store.items.filter(item => item.key !== key)].slice(0, MAX_ITEMS), hidden: store.hidden.filter(hidden => hidden !== key) });
}
export function restore(key: string) {
  save({ ...store, items: store.items.filter(item => item.key !== key), hidden: store.hidden.filter(hidden => hidden !== key) });
}
export function restoreAll() { save({ ...store, items: [], hidden: [] }); }
export function setRetention(days: RetentionDays) { save(expire({ ...store, retention: days })); }

export function useDismissed() {
  const value = useSyncExternalStore(listener => { listeners.add(listener); return () => listeners.delete(listener); }, () => store);
  const keys = new Set([...value.items.map(item => item.key), ...value.hidden]);
  return {
    retention: value.retention, items: value.items, hiddenCount: value.hidden.length,
    isDismissed: (id: string) => keys.has(dismissKey(id)),
    daysLeft: (item: DismissedItem, now = Date.now()) => Math.max(0, Math.ceil((item.at + value.retention * DAY - now) / DAY)),
  };
}
/** For tests only. */
export function resetDismissedForTests() { owner = undefined; store = { ...EMPTY, items: [], hidden: [] }; }
export { expire as expireDismissed };
