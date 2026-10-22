/**
 * Changes the Plugins section can make: upload, replace and delete files in Google Drive and Box, and add, edit
 * and delete Google Calendar events — all through the person's own connections.
 *
 * Uploads are limited to text-based files. The app's connections send file content as text, so binary files
 * (images, PDFs, Office documents) would arrive damaged; those are refused here with a link to upload at the source.
 */
import { GoogleDriveService } from "../../generated/services/GoogleDriveService";
import { BoxService } from "../../generated/services/BoxService";
import { GoogleCalendarService } from "../../generated/services/GoogleCalendarService";
import type { FileSourceKey } from "@/lib/file-preview";

export const UPLOAD_MAX_BYTES = 4 * 1024 * 1024;
const TEXT_UPLOAD_EXT = /\.(txt|csv|tsv|md|markdown|json|xml|html?|css|js|ts|ya?ml|ini|cfg|log|svg)$/i;

function failed(label: string, error: unknown): never {
  console.error(`${label} failed`);
  throw error ?? new Error(`${label} failed`);
}

export type UploadCheck = { ok: true; text: string } | { ok: false; reason: string };
/** Reads a picked file and confirms it can be uploaded intact (text, UTF-8, within the size limit). */
export async function readUploadable(file: File): Promise<UploadCheck> {
  if (file.size > UPLOAD_MAX_BYTES) return { ok: false, reason: `${file.name} is larger than ${UPLOAD_MAX_BYTES / 1024 / 1024} MB.` };
  const textType = file.type.startsWith("text/") || /^application\/(json|xml)$/.test(file.type) || TEXT_UPLOAD_EXT.test(file.name);
  if (!textType) return { ok: false, reason: `${file.name} isn't a text file. Images, PDFs and Office files can't be uploaded from the dashboard without being damaged.` };
  try {
    const text = new TextDecoder("utf-8", { fatal: true }).decode(await file.arrayBuffer());
    return { ok: true, text };
  } catch {
    return { ok: false, reason: `${file.name} isn't plain UTF-8 text, so it can't be uploaded from the dashboard without being damaged.` };
  }
}

export async function uploadTextFile(source: FileSourceKey, folderPath: string, name: string, text: string) {
  const result = source === "box" ? await BoxService.CreateFile(folderPath, name, text) : await GoogleDriveService.CreateFile(folderPath, name, text);
  if (!result.success) failed("Upload", result.error);
  return result.data;
}
export async function replaceTextFile(source: FileSourceKey, id: string, text: string) {
  const result = source === "box" ? await BoxService.UpdateFile(id, text) : await GoogleDriveService.UpdateFile(id, text);
  if (!result.success) failed("Replace", result.error);
}
export async function deleteSourceFile(source: FileSourceKey, id: string) {
  const result = source === "box" ? await BoxService.DeleteFile(id) : await GoogleDriveService.DeleteFile(id);
  if (!result.success) failed("Delete", result.error);
}

export interface EventDraft { title: string; allDay: boolean; start: string; end: string; location: string; description: string; attendees: string }

/** Converts the form's local date / date-time values into what Google Calendar expects. */
function toEventTimes(draft: EventDraft): { start: string; end: string } {
  if (draft.allDay) {
    const [y, m, d] = draft.end.split("-").map(Number);
    const exclusive = new Date(y, m - 1, d + 1); // Google all-day events end the day after
    return { start: draft.start, end: exclusive.toLocaleDateString("en-CA") };
  }
  return { start: new Date(draft.start).toISOString(), end: new Date(draft.end).toISOString() };
}

export async function createGoogleEvent(calendarId: string, draft: EventDraft) {
  const times = toEventTimes(draft);
  const result = await GoogleCalendarService.CreateEvent({
    summary: draft.title.trim(), ...times, isAllDay: draft.allDay,
    location: draft.location.trim() || undefined, description: draft.description.trim() || undefined, attendees: draft.attendees.trim() || undefined,
  }, calendarId);
  if (!result.success) failed("Add event", result.error);
  return result.data;
}
export async function updateGoogleEvent(calendarId: string, eventId: string, draft: EventDraft) {
  const times = toEventTimes(draft);
  const result = await GoogleCalendarService.UpdateEvent({
    summary: draft.title.trim(), ...times, isAllDay: draft.allDay,
    location: draft.location.trim(), description: draft.description.trim(), attendees: draft.attendees.trim() || undefined,
  }, calendarId, eventId);
  if (!result.success) failed("Edit event", result.error);
  return result.data;
}
export async function deleteGoogleEvent(calendarId: string, eventId: string) {
  const result = await GoogleCalendarService.DeleteEvent(eventId, calendarId);
  if (!result.success) failed("Delete event", result.error);
}

/** Extracts a .zip (by path) into a folder (by path), inside the same service — works for any file type. */
export async function extractArchive(source: FileSourceKey, archivePath: string, destinationPath: string) {
  const result = source === "box" ? await BoxService.ExtractFolderV2(archivePath, destinationPath, false) : await GoogleDriveService.ExtractFolderV2(archivePath, destinationPath, false);
  if (!result.success) failed("Extract", result.error);
  return Array.isArray(result.data) ? result.data.length : undefined;
}

/** Saves text as a new file; used by New file, Save plan and similar. */
export const createTextFile = uploadTextFile;
