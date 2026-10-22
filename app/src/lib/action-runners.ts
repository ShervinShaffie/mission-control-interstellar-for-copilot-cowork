import { Office365OutlookService } from "../../generated/services/Office365OutlookService";
import { OneNote_Business_Service } from "../../generated/services/OneNote_Business_Service";
import { MicrosoftTo_Do_Business_Service } from "../../generated/services/MicrosoftTo_Do_Business_Service";
import { recordCreated, type ActionResult } from "@/lib/actions";
import type { CalendarEvent } from "@/lib/command-center-evidence";
import { buildDocx, buildXlsxBase64, bytesToBase64, blobToBase64, fileName, sectionsToSheet, type DocContent, type DocFormat } from "@/lib/documents";
import { createPdfBlob } from "@/lib/pdf-export";
import { hideEmails } from "@/lib/privacy";
import { cleanDoc, type Block } from "@/lib/citations";

const OUTLOOK_DRAFTS = "https://outlook.office.com/mail/drafts";
const OUTLOOK_CALENDAR = "https://outlook.office.com/calendar/view/week";
const TODO_HOME = "https://to-do.office.com/tasks/";
const ONENOTE_HOME = "https://www.office.com/launch/onenote";

export class ActionError extends Error {
  constructor(message: string) { super(message); this.name = "ActionError"; }
}
async function ok<T>(label: string, call: Promise<{ success: boolean; data: T; error?: unknown }>): Promise<T> {
  const result = await call;
  if (!result.success) {
    console.error(`${label} failed`);
    throw new ActionError(`${label} could not be completed. It may be blocked by your organization or need a fresh sign-in.`);
  }
  return result.data;
}
function escapeHtml(text: string) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function toHtml(text: string) { return escapeHtml(text).replace(/\r?\n/g, "<br>"); }
function utcLocal(date: Date) { return date.toISOString().slice(0, 19); }

// ---------- Calendar ----------
export async function respondToInvite(event: CalendarEvent, response: string, comment: string): Promise<ActionResult> {
  if (!event.id || !["accept", "tentativelyAccept", "decline"].includes(response)) throw new ActionError("This invitation cannot be answered from here.");
  await ok("Responding to the invitation", Office365OutlookService.RespondToEvent_V2(event.id, response, { Comment: comment.trim() || undefined, SendResponse: true }));
  const label = response === "accept" ? "Accepted" : response === "decline" ? "Declined" : "Tentatively accepted";
  const links = [{ label: "Open the event", url: event.webLink ?? OUTLOOK_CALENDAR }, { label: "Open your calendar", url: OUTLOOK_CALENDAR }];
  recordCreated({ title: `${label}: ${hideEmails(event.subject ?? "invitation")}`, kind: "Calendar response", links });
  return { message: `${label}. The organizer was notified.`, links };
}
export async function blockTime(calendarId: string, start: Date, end: Date, subject: string, notes: string, category?: string): Promise<ActionResult> {
  const created = await ok("Adding the calendar block", Office365OutlookService.V4CalendarPostItem(calendarId, {
    subject, start: utcLocal(start), end: utcLocal(end), timeZone: "(UTC) Coordinated Universal Time",
    showAs: "busy", sensitivity: "private", importance: "normal", isReminderOn: true, reminderMinutesBeforeStart: 5,
    body: notes.trim() ? toHtml(notes) : undefined, categories: category ? [category] : undefined,
  }));
  const links = [{ label: "Open the new event", url: created?.webLink ?? OUTLOOK_CALENDAR }, { label: "Open your calendar", url: OUTLOOK_CALENDAR }];
  recordCreated({ title: subject, kind: "Calendar block", links });
  return { message: `Added "${subject}" to your calendar, ${start.toLocaleString([], { weekday: "short", hour: "numeric", minute: "2-digit" })} – ${end.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}.`, links };
}

// ---------- Email ----------
export async function flagMessage(id: string, subject: string): Promise<ActionResult> {
  await ok("Flagging the email", Office365OutlookService.Flag_V2(id, undefined, { flag: { flagStatus: "flagged" } }));
  return { message: `Flagged "${hideEmails(subject)}" for follow-up.`, links: [{ label: "Open the email", url: `https://outlook.office.com/mail/deeplink/read/${encodeURIComponent(id)}` }, { label: "Open Flagged email", url: "https://outlook.office.com/mail/inbox" }] };
}
export async function markRead(id: string, subject: string): Promise<ActionResult> {
  await ok("Marking the email as read", Office365OutlookService.MarkAsRead_V3(id, undefined, { isRead: true }));
  return { message: `Marked "${hideEmails(subject)}" as read.`, links: [{ label: "Open the email", url: `https://outlook.office.com/mail/deeplink/read/${encodeURIComponent(id)}` }, { label: "Open your Inbox", url: "https://outlook.office.com/mail/inbox" }] };
}
export async function setAutoReplies(start: Date, end: Date, message: string): Promise<ActionResult> {
  await ok("Turning on automatic replies", Office365OutlookService.SetAutomaticRepliesSetting_V2({ automaticRepliesSetting: {
    status: "scheduled", externalAudience: "none",
    scheduledStartDateTime: { dateTime: utcLocal(start), timeZone: "UTC" }, scheduledEndDateTime: { dateTime: utcLocal(end), timeZone: "UTC" },
    internalReplyMessage: toHtml(message), externalReplyMessage: "",
  } }));
  const links = [{ label: "Open automatic replies settings", url: "https://outlook.office.com/mail/options/mail/automaticReplies" }];
  recordCreated({ title: "Automatic replies scheduled", kind: "Outlook setting", links });
  return { message: `Automatic replies are scheduled from ${start.toLocaleDateString()} to ${end.toLocaleDateString()} (inside your organization only).`, links };
}
interface FoundMessage { id: string; subject: string; from: string }
async function findMessage(subject: string): Promise<FoundMessage | undefined> {
  const clean = subject.replace(/["\\]/g, " ").trim().slice(0, 120);
  if (!clean) return undefined;
  const data = await ok("Finding the email", Office365OutlookService.GetEmailsV3("Inbox", undefined, undefined, undefined, undefined, undefined, false, undefined, false, false, undefined, false, `"${clean}"`, 5));
  const rows = data.value ?? [];
  const wanted = clean.toLowerCase().replace(/^(re|fw|fwd):\s*/i, "");
  const match = rows.find(row => (row.subject ?? "").toLowerCase().includes(wanted)) ?? rows[0];
  return match?.id && match.from ? { id: match.id, subject: match.subject ?? subject, from: match.from } : undefined;
}
async function saveDraft(to: string, subject: string, body: string, replyTo?: string): Promise<{ id?: string }> {
  const message = { To: to, Subject: subject, Body: toHtml(body), Importance: "Normal" as const };
  if (replyTo) {
    const threaded = await Office365OutlookService.DraftEmail(message, replyTo, "Reply");
    if (threaded.success) return { id: threaded.data?.Id };
  }
  const draft = await ok("Saving the draft", Office365OutlookService.DraftEmail(message));
  return { id: draft?.Id };
}
function draftLinks(id?: string) {
  return [...(id ? [{ label: "Open the draft", url: `https://outlook.office.com/mail/drafts/id/${encodeURIComponent(id)}` }] : []), { label: "Open your Drafts folder", url: OUTLOOK_DRAFTS }];
}
const greet = (address?: string) => { const name = (address ?? "").split("@")[0].split(/[._-]/)[0]; return name ? `Hi ${name.charAt(0).toUpperCase()}${name.slice(1)},` : "Hi,"; };
const idFromLink = (url?: string) => { const match = /\/read\/([^/?#]+)/.exec(url ?? ""); return match ? decodeURIComponent(match[1]) : undefined; };
/** Saves your reply (the text you typed) as a draft in the email's thread. Free; the app never sends. */
export async function draftReply(item: { title: string; url?: string; email?: string }, notes: string): Promise<ActionResult> {
  const id = idFromLink(item.url);
  const found = id ? { id, subject: item.title, from: item.email ?? "" } : await findMessage(item.title);
  if (!found) throw new ActionError("The original email could not be found in your Inbox, so no draft was made.");
  const body = notes.trim() || `${greet(found.from)}\n\n`;
  const draft = await saveDraft(found.from, /^re:/i.test(found.subject) ? found.subject : `RE: ${found.subject}`, body, found.id);
  const links = draftLinks(draft.id);
  recordCreated({ title: `Reply draft: ${hideEmails(found.subject)}`, kind: "Outlook draft", links });
  return { message: "Your reply draft is saved in the email's thread in Outlook Drafts. Finish it and send it from Outlook; the app never sends email.", links };
}
/** Saves a new email draft with the text you typed. Free. */
export async function draftNewEmail(to: string | undefined, body: string, subject: string): Promise<ActionResult> {
  if (!to || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to)) throw new ActionError("No email address was found for this person, so no draft was made.");
  const text = body.trim() || `${greet(to)}\n\n`;
  const draft = await saveDraft(to, subject, text);
  const links = draftLinks(draft.id);
  recordCreated({ title: `Draft: ${hideEmails(subject)}`, kind: "Outlook draft", links });
  return { message: "Your draft is saved in Outlook Drafts. Finish it and send it from Outlook; the app never sends email.", links };
}

// ---------- To Do and OneNote ----------
export async function addToDo(title: string, due: string | undefined, notes: string, sourceUrl?: string): Promise<ActionResult> {
  const lists = await ok("Opening Microsoft To Do", MicrosoftTo_Do_Business_Service.GetAllTodoListsV2());
  const list = lists?.find(item => item.wellknownListName === "defaultList") ?? lists?.[0];
  if (!list?.id) throw new ActionError("No Microsoft To Do list was found.");
  const day = due && /^\d{4}-\d{2}-\d{2}/.test(due) ? due.slice(0, 10) : undefined;
  const content = [notes.trim(), sourceUrl ? `Source: ${sourceUrl}` : ""].filter(Boolean).join("\n\n");
  await ok("Adding the task", MicrosoftTo_Do_Business_Service.CreateToDoV3(list.id, {
    title: hideEmails(title).slice(0, 255), importance: "normal", status: "notStarted",
    dueDateTime: day ? { dateTime: `${day}T00:00:00`, timeZone: "UTC" } : undefined,
    body: content ? { contentType: "text", content } : undefined,
  }));
  const links = [{ label: `Open your "${list.displayName ?? "Tasks"}" list`, url: `${TODO_HOME}${encodeURIComponent(list.id)}` }, { label: "Open Microsoft To Do", url: TODO_HOME }];
  recordCreated({ title: `Task: ${hideEmails(title)}`, kind: "To Do task", links });
  return { message: `Added "${hideEmails(title)}" to Microsoft To Do${day ? `, due ${new Date(`${day}T12:00:00`).toLocaleDateString()}` : ""}.`, links };
}
export async function oneNotePage(title: string, bodyHtml: string): Promise<ActionResult> {
  await ok("Creating the OneNote page", OneNote_Business_Service.CreatePageInQuickNotes(`<!DOCTYPE html><html><head><title>${escapeHtml(title)}</title></head><body>${bodyHtml}</body></html>`));
  const links = [{ label: "Open OneNote (Quick Notes)", url: ONENOTE_HOME }];
  recordCreated({ title, kind: "OneNote page", links });
  return { message: `Created "${title}" in your OneNote Quick Notes.`, links };
}
function blocksHtml(blocks: Block[]) {
  const inline = (block: Block) => block.inlines.map(item => item.kind === "text" ? escapeHtml(item.text) : item.cites.map(cite => ` <a href="${escapeHtml(cite.url)}">${escapeHtml(cite.name)}</a>`).join("")).join("");
  let html = "", list = false;
  for (const block of blocks) {
    if (block.kind === "bullet" && !list) { html += "<ul>"; list = true; }
    if (block.kind !== "bullet" && list) { html += "</ul>"; list = false; }
    html += block.kind === "bullet" ? `<li>${inline(block)}</li>` : `<p>${inline(block)}</p>`;
  }
  return list ? `${html}</ul>` : html;
}
/** Saves a brief, report or document as a OneNote page (formatting and clickable sources kept). OneNote lives in your OneDrive. */
export async function saveDocToOneNote(doc: DocContent): Promise<ActionResult> {
  const clean = cleanDoc(doc);
  const body = `<h1>${escapeHtml(clean.title)}</h1>${clean.sections.map(section => `${section.heading ? `<h2>${escapeHtml(section.heading)}</h2>` : ""}${blocksHtml(section.blocks)}`).join("")}${clean.sources.length ? `<h2>Sources</h2><ol>${clean.sources.map(source => `<li><a href="${escapeHtml(source.url)}">${escapeHtml(source.name)}</a></li>`).join("")}</ol>` : ""}`;
  return oneNotePage(clean.title || "Document", body);
}
export function listHtml(items: string[]) { return `<ul>${items.map(item => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`; }
export function todoHtml(items: string[]) { return items.map(item => `<p data-tag="to-do">${escapeHtml(item)}</p>`).join(""); }

// ---------- Copilot-written documents ----------
// ---------- Saving documents to OneDrive ----------
async function buildFile(content: DocContent, format: DocFormat, pdfTarget?: HTMLElement | null, rows?: (string | number)[][]) {
  const name = fileName(content.title, format);
  if (format === "docx") return { name, base64: bytesToBase64(buildDocx(content)) };
  if (format === "xlsx") return { name, base64: buildXlsxBase64([{ name: content.title, rows: rows ?? sectionsToSheet(content) }]) };
  if (!pdfTarget) throw new ActionError("The preview to turn into a PDF isn't on screen.");
  return { name, base64: await blobToBase64((await createPdfBlob(pdfTarget, name)).blob) };
}
const MIME: Record<DocFormat, string> = { docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", pdf: "application/pdf" };
/** Free: builds the file in the browser and downloads it. No connection or Copilot Credits needed. */
export async function downloadDocument(content: DocContent, format: DocFormat, pdfTarget?: HTMLElement | null, rows?: (string | number)[][]): Promise<ActionResult> {
  const { name, base64 } = await buildFile(content, format, pdfTarget, rows);
  const bytes = Uint8Array.from(atob(base64), char => char.charCodeAt(0));
  const url = URL.createObjectURL(new Blob([bytes], { type: MIME[format] }));
  const link = document.createElement("a");
  link.href = url; link.download = name; link.style.display = "none";
  document.body.appendChild(link); link.click(); link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
  return { message: `Downloaded ${name}. Check your browser's downloads.`, links: [] };
}

// ---------- Used by Ask (content already written; these calls use no Copilot Credits) ----------
const EMAIL = /^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$/;
export function validAddresses(list: string) { return list.split(/[;,\s]+/).map(item => item.trim()).filter(item => EMAIL.test(item)); }
export async function saveEmailDraft(to: string, subject: string, body: string): Promise<ActionResult> {
  const addresses = validAddresses(to);
  const draft = await saveDraft(addresses.join(";"), subject.trim() || "(no subject)", body);
  const links = draftLinks(draft.id);
  recordCreated({ title: `Draft: ${hideEmails(subject)}`, kind: "Outlook draft", links });
  return { message: `Saved in Outlook Drafts${addresses.length ? "" : " without a recipient — add one in Outlook"}. Review and send it from Outlook; the app never sends email.`, links };
}
/** Sends a meeting invitation from the user's primary calendar (attendees receive it immediately). */
export async function sendMeetingInvite(calendarId: string, invite: { subject: string; start: Date; end: Date; attendees: string; body: string; location: string }): Promise<ActionResult> {
  const attendees = validAddresses(invite.attendees);
  if (!attendees.length) throw new ActionError("No attendee addresses were found, so nothing was sent. Use “Open in Outlook” to add people yourself.");
  if (!(invite.end > invite.start)) throw new ActionError("The end time must be after the start time.");
  const created = await ok("Sending the meeting invite", Office365OutlookService.V4CalendarPostItem(calendarId, {
    subject: invite.subject.trim() || "Meeting", start: utcLocal(invite.start), end: utcLocal(invite.end), timeZone: "(UTC) Coordinated Universal Time",
    requiredAttendees: attendees.join(";"), body: invite.body.trim() ? toHtml(invite.body) : undefined, location: invite.location.trim() || undefined,
    showAs: "busy", importance: "normal", isReminderOn: true, reminderMinutesBeforeStart: 15, responseRequested: true,
  }));
  const links = [{ label: "Open the meeting", url: created?.webLink ?? OUTLOOK_CALENDAR }, { label: "Open your calendar", url: OUTLOOK_CALENDAR }];
  recordCreated({ title: `Meeting: ${hideEmails(invite.subject)}`, kind: "Meeting invite", links });
  return { message: `Invite sent to ${attendees.length} ${attendees.length === 1 ? "person" : "people"} for ${invite.start.toLocaleString([], { weekday: "short", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" })}. To add a Teams link, open it in Outlook.`, links };
}
/** Free fallback: opens a prefilled new-event form in Outlook on the web. */
export function outlookComposeLink(subject: string, start: Date, end: Date, body: string) {
  const params = new URLSearchParams({ subject, startdt: start.toISOString(), enddt: end.toISOString(), body: hideEmails(body).slice(0, 1500) });
  return `https://outlook.office.com/calendar/deeplink/compose?${params.toString()}`;
}
