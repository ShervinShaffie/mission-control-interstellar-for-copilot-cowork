import { hideEmails } from "@/lib/privacy";
import { filesFromAttachments, filesFromHtml, type FileRef } from "@/lib/files";
import type { GraphCalendarEventClientReceive } from "../../generated/models/Office365OutlookModel";

export interface CalendarEvent extends GraphCalendarEventClientReceive {
  onlineStatus?: boolean;
  onlineProvider?: string;
  onlineMetadataUnavailable?: boolean;
}
export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
export function timestamp(value: unknown): string | undefined {
  if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$/i.test(value) || !Number.isFinite(Date.parse(value))) return undefined;
  const date = new Date(`${value.slice(0, 10)}T00:00:00Z`);
  return date.toISOString().slice(0, 10) === value.slice(0, 10) ? value : undefined;
}
function optionalString(value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== "string") throw new Error("Invalid source text field");
  return value;
}
export function hasAttendees(event: CalendarEvent) {
  return [event.requiredAttendees, event.optionalAttendees].some(value => value?.split(";").some(attendee => attendee.trim()));
}
export function isAcceptedMeeting(event: CalendarEvent) {
  return event.responseType === "accepted" && !event.isAllDay && hasAttendees(event);
}
export function isPrivateEvent(event: CalendarEvent) {
  return event.sensitivity === "private" || event.sensitivity === "confidential";
}
export function eventTitle(event: CalendarEvent) {
  if (event.sensitivity === "private") return "Private appointment";
  if (event.sensitivity === "confidential") return "Busy";
  return hideEmails(event.subject) || "Untitled event";
}
const joinLinkProviders: [RegExp, string][] = [
  [/https:\/\/teams\.microsoft\.com\/l\/meetup-join\//i, "Teams"],
  [/https:\/\/teams\.live\.com\/meet\//i, "Teams"],
  [/https:\/\/[\w.-]*zoom\.us\/(?:j|my|w)\//i, "Zoom"],
  [/https:\/\/meet\.google\.com\//i, "Google Meet"],
  [/https:\/\/[\w.-]*webex\.com\/(?:meet|join|[\w.-]+\/j\.php)/i, "Webex"],
];
/**
 * Marks an event as online only when its own invitation text or location holds
 * a recognized meeting join link. No link means "not shown", never "in person".
 * Private and confidential events are not inspected.
 */
export function withJoinLinkEvidence(source: CalendarEvent): CalendarEvent {
  const event: CalendarEvent = { ...source };
  if (isPrivateEvent(event)) return event;
  const text = `${event.location ?? ""}\n${event.body ?? ""}`;
  const match = joinLinkProviders.find(([pattern]) => pattern.test(text));
  if (match) return { ...event, onlineStatus: true, onlineProvider: match[1] };
  return event;
}
export function applyEventMetadata(event: CalendarEvent, metadata: unknown): CalendarEvent {
  if (!isRecord(metadata)) return { ...event };
  // Metadata from an individual event request must belong to that exact event.
  if (metadata !== event && metadata.id !== event.id) return { ...event };
  const sensitivity = metadata.sensitivity;
  return {
    ...event,
    sensitivity: sensitivity === "normal" || sensitivity === "personal" || sensitivity === "private" || sensitivity === "confidential" ? sensitivity : event.sensitivity,
    onlineStatus: typeof metadata.isOnlineMeeting === "boolean" ? metadata.isOnlineMeeting : event.onlineStatus,
    onlineProvider: typeof metadata.onlineMeetingProvider === "string" && metadata.onlineMeetingProvider.trim() ? metadata.onlineMeetingProvider : event.onlineProvider,
  };
}
export interface ChatRecord {
  id: string;
  topic: string;
  url?: string;
  readCutoff?: string;
  hidden?: boolean;
}
export interface ChatPage { rows: ChatRecord[]; hasMore: boolean }
export function parseChats(data: unknown): ChatPage {
  if (!isRecord(data) || !Array.isArray(data.value)) throw new Error("Unsupported chat-list response");
  const next = optionalString(data["@odata.nextLink"]);
  const rows = data.value.map((row: unknown): ChatRecord => {
    if (!isRecord(row) || typeof row.id !== "string" || !row.id) throw new Error("Invalid chat record");
    const topic = optionalString(row.topic);
    const url = optionalString(row.webUrl);
    if (row.viewpoint !== undefined && row.viewpoint !== null && !isRecord(row.viewpoint)) throw new Error("Invalid chat viewpoint");
    const viewpoint = isRecord(row.viewpoint) ? row.viewpoint : undefined;
    if (viewpoint?.isHidden !== undefined && typeof viewpoint.isHidden !== "boolean") throw new Error("Invalid chat visibility");
    return { id: row.id, topic: topic || "Chat without a title", url, readCutoff: timestamp(viewpoint?.lastMessageReadDateTime), hidden: typeof viewpoint?.isHidden === "boolean" ? viewpoint.isHidden : undefined };
  });
  return { rows, hasMore: !!next };
}
export interface MessagePreview {
  id: string;
  created?: string;
  senderId?: string;
  sender?: string;
  summary?: string;
  content?: string;
  messageType?: string;
  deleted?: boolean;
  /** Files attached to or linked in the message. */
  files?: FileRef[];
}
export interface MessagePage { rows: MessagePreview[]; hasMore: boolean }
export function parseMessages(data: unknown): MessagePage {
  if (!isRecord(data) || !Array.isArray(data.value) || data.value.length > 10) throw new Error("Unsupported message-list response");
  const next = optionalString(data["@odata.nextLink"]);
  const rows = data.value.map((row: unknown): MessagePreview => {
    if (!isRecord(row) || typeof row.id !== "string" || !row.id) throw new Error("Invalid message record");
    const from = isRecord(row.from) ? row.from : undefined;
    const user = isRecord(from?.user) ? from.user : undefined;
    const body = isRecord(row.body) ? row.body : undefined;
    if (row.deleted !== undefined && typeof row.deleted !== "boolean") throw new Error("Invalid message deletion state");
    if (row.deletedDateTime !== undefined && row.deletedDateTime !== null && !timestamp(row.deletedDateTime)) throw new Error("Invalid message deletion timestamp");
    return { id: row.id, created: timestamp(row.createdDateTime), senderId: optionalString(user?.id), sender: optionalString(user?.displayName), summary: optionalString(row.summary), content: optionalString(body?.content), messageType: optionalString(row.messageType), deleted: row.deleted === true || !!timestamp(row.deletedDateTime), files: [...filesFromAttachments(row.attachments), ...filesFromHtml(optionalString(body?.content))].slice(0, 6) };
  });
  return { rows, hasMore: !!next };
}
export type MessageReadState = "unread" | "read" | "own" | "excluded" | "unknown";
export function messageReadState(message: MessagePreview, chat: ChatRecord, currentUserId?: string): MessageReadState {
  if (message.deleted || message.messageType === "systemEventMessage") return "excluded";
  if (currentUserId && message.senderId === currentUserId) return "own";
  const cutoff = timestamp(chat.readCutoff), created = timestamp(message.created);
  if (!cutoff || !created) return "unknown";
  if (Date.parse(created) <= Date.parse(cutoff)) return "read";
  if (!currentUserId || !message.senderId || message.messageType !== "message") return "unknown";
  return "unread";
}
