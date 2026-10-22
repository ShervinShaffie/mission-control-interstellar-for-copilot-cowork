import type { IOperationResult } from "@microsoft/managed-apps/data";
import { MicrosoftTeamsService } from "../../generated/services/MicrosoftTeamsService";
import { Office365OutlookService } from "../../generated/services/Office365OutlookService";
import { OneDriveforBusinessService } from "../../generated/services/OneDriveforBusinessService";
import { isRecord, parseChats, parseMessages, withJoinLinkEvidence, type CalendarEvent } from "@/lib/command-center-evidence";
import { readFailureCause, TeamsReadError, type ReadFailureCause } from "@/lib/command-center-failures";

export function unwrap<T>(result: IOperationResult<T>, label: string): T {
  if (!result.success) {
    console.error(`${label} request failed`, result.error);
    throw new Error(`${label} unavailable`);
  }
  return result.data;
}
/** One work signal shown on the dashboard (built from the direct connections). */
export interface WorkItem {
  kind: string; title: string; summary: string; next: string; status: string; severity: string;
  date: string; endDate: string; email: string; url: string; evidence: string; relatedMeeting: string;
}
/** Kept for older components: a failed direct read. */
export class WorkContextError extends Error {
  constructor(readonly stage: "search" | "validation", readonly failureCause: ReadFailureCause = "unknown") {
    super("Work data could not be read");
    this.name = "WorkContextError";
  }
}
export const INBOX_UNREAD_CAP = 100;
export interface InboxUnreadCount { count: number; capped: boolean }
/**
 * Counts unread Inbox messages with the standard mail read (the direct folder
 * request is blocked by organization policy in some tenants). At most
 * INBOX_UNREAD_CAP messages are checked; reaching the cap is shown as "100+".
 */
export async function loadInboxUnreadCount(): Promise<InboxUnreadCount> {
  const data: unknown = unwrap(await Office365OutlookService.GetEmailsV3("Inbox", undefined, undefined, undefined, undefined, undefined, false, undefined, true, false, undefined, false, undefined, INBOX_UNREAD_CAP), "Inbox count");
  if (!isRecord(data) || !Array.isArray(data.value)) throw new Error("Inbox unread messages were not returned as a list");
  const count = data.value.length;
  return { count, capped: count >= INBOX_UNREAD_CAP };
}
async function readTeams<T, U>(operation: () => Promise<IOperationResult<T>>, parse: (data: unknown) => U, stage: "chat-list" | "messages"): Promise<U> {
  let result: IOperationResult<T>;
  try {
    result = await operation();
    if (!result.success) throw result.error;
  } catch (error) {
    const failureCause = readFailureCause(error);
    console.error(`Teams ${stage} request failed`, { stage, cause: failureCause });
    throw new TeamsReadError(stage, failureCause);
  }
  try { return parse(result.data); }
  catch {
    console.error(`Teams ${stage} response validation failed`, { stage: "validation" });
    throw new TeamsReadError("validation", "response");
  }
}
// The connector's List chats operation accepts topic = all | isDefined | notDefined.
// "all" lists chats whether or not they have a title (an empty value was a 404).
export const TEAMS_CHAT_TOPIC_FILTER = "all";
export async function loadChats(chatType: string) {
  return readTeams(() => MicrosoftTeamsService.GetChats(chatType, TEAMS_CHAT_TOPIC_FILTER), parseChats, "chat-list");
}
export async function loadChatMessages(id: string) {
  return readTeams(() => MicrosoftTeamsService.GetMessagesFromChat(id, undefined, undefined, "10"), parseMessages, "messages");
}

export async function loadCalendarPage(calendarId: string, start: string, end: string, offset: number, pageSize: number, order?: string): Promise<CalendarEvent[]> {
  const data = unwrap(await Office365OutlookService.GetEventsCalendarViewV3(calendarId, start, end, undefined, order, pageSize, offset), "Calendar");
  if (!Array.isArray(data.value) || data.value.length > pageSize) throw new Error("Unsupported calendar page");
  // Sensitivity comes with the calendar page itself. Online status is read
  // from a join link in the invitation; no further per-event request is made.
  return data.value.map(event => withJoinLinkEvidence(event));
}
// Outlook's preset category colors (preset0-preset24), as shown in Outlook.
export const OUTLOOK_CATEGORY_COLORS: Record<string, string> = {
  preset0: "#e81123", preset1: "#f7630c", preset2: "#8e562e", preset3: "#fce100", preset4: "#13a10e",
  preset5: "#00b7c3", preset6: "#7a7574", preset7: "#0078d4", preset8: "#8764b8", preset9: "#c239b3",
  preset10: "#69797e", preset11: "#486860", preset12: "#a0aeb2", preset13: "#7a7574", preset14: "#393939",
  preset15: "#a4262c", preset16: "#ca5010", preset17: "#603d30", preset18: "#c19c00", preset19: "#0b6a0b",
  preset20: "#038387", preset21: "#5d5a58", preset22: "#004e8c", preset23: "#5c2e91", preset24: "#881798",
};
export interface CalendarCategory { name: string; color?: string }
export interface CalendarCategories { categories: CalendarCategory[]; colorsAvailable: boolean }
/**
 * The user's own Outlook categories, from the connector's standard category read. The app assigns the colors:
 * reading Outlook's own category colors needs the connector's open-ended "Send an HTTP request" action, which the
 * default policy for these apps blocks (and which tripped the "blocked by your organization's policy" banner).
 */
export async function loadCalendarCategories(): Promise<CalendarCategories> {
  const names = unwrap(await Office365OutlookService.GetOutlookCategoryNames(), "Calendar categories");
  if (!Array.isArray(names)) throw new Error("Calendar categories were not returned as a list");
  return { categories: names.flatMap(row => typeof row.displayName === "string" && row.displayName.trim() ? [{ name: row.displayName }] : []), colorsAvailable: false };
}
export async function loadRootFolders() {
  return unwrap(await OneDriveforBusinessService.ListRootFolder(), "Drive folders");
}
export async function loadFolder(id: string, token?: string) {
  const page = unwrap(await OneDriveforBusinessService.ListFolderV2(id, token, 20), "Folder contents");
  let next: string | undefined;
  if (page.nextLink) {
    next = new URL(page.nextLink, window.location.origin).searchParams.get("skipToken") || undefined;
    if (!next) throw new Error("Folder continuation could not be read");
  }
  return { rows: page.value ?? [], next };
}
