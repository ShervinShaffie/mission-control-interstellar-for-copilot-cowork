import { useSyncExternalStore } from "react";
import type { DocContent } from "@/lib/documents";
import type { BriefKind } from "@/lib/copilot-prompts";

export type ActionCost = "free" | "credits";
export interface ActionLink { label: string; url: string }
export interface ActionResult {
  message: string;
  links: ActionLink[];
  /** Draft text the user can copy (for example a Teams reply). */
  text?: string;
  /** Generated document content, shown as a preview the user can save. */
  doc?: DocContent;
}
export interface ActionInput { notes: string; choice?: string }
export interface SuggestedAction {
  id: string;
  title: string;
  detail?: string;
  /** Higher runs first: overdue 100, due today 90, prep soon 85, invites 80, replies 75, overlaps 70 … nice-to-have 20. */
  priority: number;
  cost: ActionCost;
  choices?: { value: string; label: string }[];
  /** Shown as the notes box label. Credit actions always get a notes box. */
  notesLabel?: string;
  notesDefault?: string;
  runLabel?: string;
  /** Omitted for attention items whose only step is opening the source (free). */
  run?: (input: ActionInput) => Promise<ActionResult>;
  /** Optional "Draft with Copilot": a ready-made request for the person's own Microsoft 365 Copilot (never sent by the app). */
  copilot?: (input: ActionInput) => string;
  /** This action's run only builds a document; "Draft with Copilot" hands that document to Copilot as material. */
  copilotBrief?: BriefKind;
  /** Where the signal came from (shown as source badges on Today). */
  sources?: SourceKey[];
  /** Direct link to the source item (free). */
  open?: ActionLink;
}
export type SourceKey = "calendar" | "outlook" | "teams" | "todo" | "onenote" | "copilot" | "sharepoint" | "onedrive" | "loop";
export const SOURCE_LABELS: Record<SourceKey, string> = { calendar: "Calendar", outlook: "Outlook", teams: "Teams", todo: "To Do", onenote: "OneNote", copilot: "Team activity", sharepoint: "SharePoint", onedrive: "OneDrive", loop: "Loop" };
export const PRIORITY = { overdue: 100, dueToday: 90, prepSoon: 85, invite: 80, reply: 75, overlap: 70, awayConflict: 65, followUp: 60, focus: 50, tasks: 45, insight: 40, document: 30, niceToHave: 20 } as const;
export const CREDIT_NOTE = "Nothing runs until you select Run.";
export const SAVE_NOTE = "";

export function rankActions(actions: SuggestedAction[]): SuggestedAction[] {
  return [...actions].sort((a, b) => b.priority - a.priority || (a.cost === b.cost ? 0 : a.cost === "free" ? -1 : 1) || a.title.localeCompare(b.title));
}

// ---------- Everything created or changed this session (shown in Create) ----------
export interface CreatedItem { id: string; title: string; kind: string; at: number; links: ActionLink[] }
let created: CreatedItem[] = [];
const listeners = new Set<() => void>();
export function recordCreated(item: Omit<CreatedItem, "id" | "at">) {
  created = [{ ...item, id: `${Date.now()}-${created.length}`, at: Date.now() }, ...created].slice(0, 50);
  listeners.forEach(listener => listener());
}
export function useCreatedItems(): CreatedItem[] {
  return useSyncExternalStore(listener => { listeners.add(listener); return () => listeners.delete(listener); }, () => created);
}
