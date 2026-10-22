export const componentCatalog = [
  { id: "today", title: "Today", detail: "Your plan for the day, what needs attention now, and the most important actions across everything.", wide: true },
  { id: "ask", title: "Quick actions", detail: "Draft an email, send a meeting invite, add a To Do task, create a OneNote page or a document — plus your History.", wide: true },
  { id: "projects", title: "Projects", detail: "Your project teams, most in need of attention first, with status, deadlines, meetings, To Do tasks, briefings, a weekly status report and a tracker.", wide: true },
  { id: "calendar", title: "Calendar", detail: "Next seven days, day agenda and who is away, with invites, overlaps, focus time and meeting prep.", wide: true },
  { id: "email", title: "Email", detail: "Emails needing a reply, unread, flagged and recent mail, with reply drafts and automatic replies.", wide: false },
  { id: "teams", title: "Teams", detail: "Who is waiting on you and your recent chats, with a catch-up document.", wide: false },
  { id: "files", title: "Files", detail: "Files that need you now — linked in your email, Teams and meeting invites, or trending around you — plus your OneDrive.", wide: false },
  { id: "onenote", title: "OneNote", detail: "Recent pages and open to-dos, with day pages, meeting notes and To Do.", wide: false },
  { id: "planner", title: "Planner", detail: "Your Planner tasks across every plan — overdue, due soon and in progress — with progress, labels and checklists.", wide: false },
  { id: "plugins", title: "Plugins", detail: "GitHub, Google Drive, Box and Google Calendar in one place, each with links straight to the source.", wide: true },
  { id: "insights", title: "Time insights", detail: "Meeting load and time by your calendar categories, with targets and suggestions.", wide: true },
] as const;
export type ComponentId = typeof componentCatalog[number]["id"];
export type DensityView = "bar" | "line" | "table";
export type AllocationView = "bar" | "pie" | "table";
export type CalendarView = "graphic" | "table";
export type SearchGroup = "priorities" | "meetings" | "replies" | "collaboration" | "pinned";
export type SearchSwitches = Record<SearchGroup, boolean>;
export interface DashboardPreferences {
  version: 2;
  visible: ComponentId[];
  densityView: DensityView;
  allocationView: AllocationView;
  calendarView: CalendarView;
  theme: "system" | "light" | "dark";
  searches: SearchSwitches;
  /** Target share (0-100) of scheduled time per calendar category name. */
  targets: Record<string, number>;
  /** Components added after a layout was saved, already offered once (so removing them later sticks). */
  introduced: ComponentId[];
  /** True once the user has edited time targets; until then suggested targets are shown. */
  targetsSet: boolean;
  /** Version of the default arrangement this layout was last aligned with. */
  layoutRevision: number;
  /** Master switch for the automatic Copilot (Work IQ) searches on open and Refresh. */
  copilot: boolean;
  /** Master switch for everything that uses Work IQ SharePoint (search, pinned sites, pinned changes, Save to OneDrive). */
  sharepoint: boolean;
  /** Components shown collapsed to their title bar. */
  collapsed: ComponentId[];
  /** Order of the panels inside Plugins. */
  pluginOrder: PluginKey[];
  /** Show Google Calendar events alongside Outlook in the Calendar section. */
  combineGoogle: boolean;
  /** Days shown in the calendar graphic: 1, 3 or 7. */
  calendarDays: 1 | 3 | 7;
}
export const PLUGIN_KEYS = ["googledrive", "box", "googlecalendar", "github"] as const;
export type PluginKey = typeof PLUGIN_KEYS[number];
export const SEARCH_GROUPS = ["priorities", "meetings", "replies", "collaboration"] as const;
/** Every part of the dashboard reads directly from your Microsoft 365 apps; nothing needs switching off. */
export function effectiveSwitches(): SearchSwitches {
  // Every part of the dashboard now reads directly from your Microsoft 365 apps, with nothing to switch off.
  return { priorities: true, meetings: true, replies: true, collaboration: true, pinned: false };
}
/** Bump when the default arrangement changes: older saved layouts are re-ordered once (removed components stay removed). */
export const LAYOUT_REVISION = 2;
/** Components added in later versions: appended once to older saved layouts. */
const LATER: { id: ComponentId; after: ComponentId }[] = [{ id: "projects", after: "today" }, { id: "ask", after: "today" }, { id: "files", after: "teams" }, { id: "planner", after: "onenote" }, { id: "plugins", after: "planner" }];
export function defaultPreferences(): DashboardPreferences {
  return {
    version: 2, visible: componentCatalog.map(item => item.id), densityView: "bar", allocationView: "pie", calendarView: "graphic", theme: "system",
    searches: { priorities: true, meetings: true, replies: true, collaboration: true, pinned: true }, targets: {}, introduced: LATER.map(item => item.id), targetsSet: false, layoutRevision: LAYOUT_REVISION, copilot: true, sharepoint: true, collapsed: [], pluginOrder: [...PLUGIN_KEYS], combineGoogle: true, calendarDays: 7,
  };
}
export function isComponentId(value: unknown): value is ComponentId {
  return componentCatalog.some(item => item.id === value);
}
function bool(value: unknown, fallback: boolean) { return typeof value === "boolean" ? value : fallback; }
export function parsePreferences(value: unknown): DashboardPreferences {
  const fallback = defaultPreferences();
  if (typeof value !== "object" || value === null || !("version" in value) || value.version !== 2 || !("visible" in value) || !Array.isArray(value.visible)) return fallback;
  const record = value as Record<string, unknown>;
  const searches = typeof record.searches === "object" && record.searches !== null ? record.searches as Record<string, unknown> : {};
  const targets: Record<string, number> = {};
  if (typeof record.targets === "object" && record.targets !== null) {
    for (const [name, target] of Object.entries(record.targets as Record<string, unknown>)) {
      if (typeof target === "number" && Number.isFinite(target) && target >= 0 && target <= 100 && name.length <= 80) targets[name] = Math.round(target);
    }
  }
  const visible = [...new Set((record.visible as unknown[]).filter(isComponentId))];
  const introduced = Array.isArray(record.introduced) ? (record.introduced as unknown[]).filter(isComponentId) : [];
  for (const later of LATER) {
    if (introduced.includes(later.id) || visible.includes(later.id)) continue;
    const at = visible.indexOf(later.after);
    visible.splice(at >= 0 ? at + 1 : later.id === "files" ? visible.length : 0, 0, later.id);
  }
  const revision = typeof record.layoutRevision === "number" ? record.layoutRevision : 1;
  const arranged = revision < LAYOUT_REVISION ? componentCatalog.map(item => item.id).filter(id => visible.includes(id)) : visible;
  return {
    version: 2,
    visible: arranged,
    layoutRevision: LAYOUT_REVISION,
    copilot: bool(record.copilot, true),
    sharepoint: bool(record.sharepoint, bool(searches.pinned, true)),
    targetsSet: bool(record.targetsSet, Object.keys(targets).length > 0),
    introduced: [...new Set([...introduced, ...LATER.map(item => item.id)])],
    densityView: record.densityView === "bar" || record.densityView === "line" || record.densityView === "table" ? record.densityView : fallback.densityView,
    allocationView: record.allocationView === "bar" || record.allocationView === "pie" || record.allocationView === "table" ? record.allocationView : fallback.allocationView,
    calendarView: record.calendarView === "graphic" || record.calendarView === "table" ? record.calendarView : fallback.calendarView,
    theme: record.theme === "dark" || record.theme === "light" || record.theme === "system" ? record.theme : fallback.theme,
    searches: { priorities: bool(searches.priorities, true), meetings: bool(searches.meetings, true), replies: bool(searches.replies, true), collaboration: bool(searches.collaboration, true), pinned: bool(searches.pinned, true) },
    targets,
    combineGoogle: bool(record.combineGoogle, true),
    calendarDays: record.calendarDays === 1 || record.calendarDays === 3 ? record.calendarDays : 7,
    pluginOrder: Array.isArray(record.pluginOrder)
      ? [...new Set([...(record.pluginOrder as unknown[]).filter((key): key is PluginKey => PLUGIN_KEYS.includes(key as PluginKey)), ...PLUGIN_KEYS])]
      : [...PLUGIN_KEYS],
    collapsed: Array.isArray(record.collapsed) ? [...new Set((record.collapsed as unknown[]).filter(isComponentId))] : [],
  };
}
export function preferenceKey(playerId?: string): string | undefined {
  // Only a directory ID returned by a successful current-player profile read.
  return playerId && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(playerId)
    ? `mission-control-interstellar:layout:v2:${playerId.toLowerCase()}` : undefined;
}
export function moveComponent(visible: ComponentId[], id: ComponentId, direction: -1 | 1): ComponentId[] {
  const index = visible.indexOf(id), next = index + direction;
  if (index < 0 || next < 0 || next >= visible.length) return visible;
  const result = [...visible];
  [result[index], result[next]] = [result[next], result[index]];
  return result;
}
// Drop after a later box, or before an earlier box; never mutate the stored order.
export function reorderComponent(visible: ComponentId[], id: ComponentId, target: ComponentId): ComponentId[] {
  const from = visible.indexOf(id), to = visible.indexOf(target);
  if (from < 0 || to < 0 || from === to) return visible;
  const next = [...visible];
  next.splice(from, 1);
  next.splice(to, 0, id);
  return next;
}
