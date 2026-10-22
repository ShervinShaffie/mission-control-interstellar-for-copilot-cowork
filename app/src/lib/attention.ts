/**
 * Today's holistic priority list. Free: built in the browser from data already loaded on the dashboard
 * (calendar, Outlook, Teams, To Do, OneNote and team activity). Nothing here calls a service;
 * each item's optional `run` only executes when the user selects Run.
 */
import type { ActionLink, SourceKey, SuggestedAction } from "@/lib/actions";
import { addToDo, blockTime, draftNewEmail, draftReply, flagMessage, markRead, respondToInvite } from "@/lib/action-runners";
import { eventTitle, hasAttendees, isPrivateEvent, type CalendarEvent } from "@/lib/command-center-evidence";
import type { WorkItem } from "@/lib/command-center-data";
import type { OneNotePage } from "@/lib/direct-sources";
import { firstName, hideEmails } from "@/lib/privacy";
import { firstLinkUrl } from "@/lib/source-links";
import { agoLabel, type WaitingChat } from "@/lib/teams-waiting";
import { TODO_HOME, taskState, type OpenTask } from "@/lib/todo";
import type { TeamFeed } from "@/lib/direct-signals";
import { meetingBrief, projectBrief } from "@/lib/briefs";
import { awayRequest, fileRequest, followUpRequest, replyRequest, teamsReplyRequest } from "@/lib/copilot-prompts";
import type { FileItem } from "@/lib/files";
import { endOf, freeGaps, startOf } from "@/lib/suggestions";

export type Horizon = "overdue" | "now" | "today" | "tomorrow" | "week" | "later";
export const HORIZONS: { id: Horizon; label: string }[] = [
  { id: "overdue", label: "Overdue" }, { id: "now", label: "Next 2 hours" }, { id: "today", label: "Today" },
  { id: "tomorrow", label: "Tomorrow" }, { id: "week", label: "This week" }, { id: "later", label: "When you can" },
];
export interface AttentionItem extends SuggestedAction { sources: SourceKey[]; horizon: Horizon }
export interface InboxMail { id?: string; subject?: string; from?: string; importance?: string; receivedDateTime?: string; bodyPreview?: string; isRead?: boolean }
export interface AttentionInput {
  now: Date; calendarId?: string;
  events: CalendarEvent[]; unread: InboxMail[]; flagged: InboxMail[]; waiting: WaitingChat[]; tasks: OpenTask[]; pages: OneNotePage[];
  replies: WorkItem[]; priorities: WorkItem[]; meetings: WorkItem[]; collaboration: WorkItem[];
  /** Team activity (projects). */
  feeds?: TeamFeed[];
  /** Files ranked by what needs you (Files → For you now). */
  files?: FileItem[];
}

const BOOST: Record<Horizon, number> = { overdue: 40, now: 35, today: 25, tomorrow: 12, week: 5, later: 0 };
const clock = (date: Date) => date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
const shortDay = (date: Date) => date.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" });
const sameDay = (a: Date, b: Date) => a.toDateString() === b.toDateString();
const RISK = /\b(red|at risk|amber|blocked|blocker|behind|delayed|slipp|escalat|overdue|missed)\b/i;

/** Which Microsoft 365 app a source link points to. */
export function sourceOfUrl(url?: string): SourceKey | undefined {
  if (!url) return undefined;
  let host = "";
  try { host = new URL(url).hostname.toLowerCase(); } catch { return undefined; }
  if (host.endsWith("-my.sharepoint.com")) return "onedrive";
  if (host.endsWith("sharepoint.com")) return "sharepoint";
  if (host.includes("teams.microsoft.com") || host.includes("teams.cloud.microsoft")) return "teams";
  if (host.includes("outlook.")) return "outlook";
  if (host.includes("loop.")) return "loop";
  if (host.includes("onenote")) return "onenote";
  if (host.includes("to-do.office.com")) return "todo";
  return undefined;
}
function itemLink(item: WorkItem) { return item.url || firstLinkUrl(item.evidence) || firstLinkUrl(item.summary) || undefined; }
/** Where to look when an item has no link: an Outlook search for the item (free). */
const outlookSearch = (text: string): ActionLink => ({ label: "Find it in Outlook", url: `https://outlook.office.com/mail/inbox?searchquery=${encodeURIComponent(text.slice(0, 120))}` });
function openFor(item: WorkItem, label = "Open the source"): ActionLink {
  const link = itemLink(item);
  if (!link) return outlookSearch(item.title);
  const source = sourceOfUrl(link);
  return { label: source === "teams" ? "Open in Teams" : source === "outlook" ? "Open in Outlook" : source === "sharepoint" || source === "onedrive" ? "Open the file" : source === "loop" ? "Open in Loop" : label, url: link };
}
function copilotSources(item: WorkItem): SourceKey[] {
  const from = sourceOfUrl(itemLink(item));
  return from ? [from] : ["teams"];
}
function horizonOfDate(value: string | undefined, now: Date, fallback: Horizon): Horizon {
  if (!value) return fallback;
  const date = /^\d{4}-\d{2}-\d{2}$/.test(value) ? new Date(`${value}T23:59:59`) : new Date(value);
  if (Number.isNaN(date.getTime())) return fallback;
  if (date < now && !sameDay(date, now)) return "overdue";
  if (sameDay(date, now)) return date.getTime() - now.getTime() <= 2 * 36e5 && date > now ? "now" : "today";
  const tomorrow = new Date(now); tomorrow.setDate(now.getDate() + 1);
  if (sameDay(date, tomorrow)) return "tomorrow";
  return date.getTime() - now.getTime() <= 7 * 864e5 ? "week" : "later";
}
function dueText(value: string | undefined, horizon: Horizon) {
  if (!value) return "";
  const date = new Date(/^\d{4}-\d{2}-\d{2}$/.test(value) ? `${value}T12:00:00` : value);
  if (Number.isNaN(date.getTime())) return "";
  return horizon === "overdue" ? `Overdue — was due ${shortDay(date)}` : horizon === "today" || horizon === "now" ? "Due today" : horizon === "tomorrow" ? "Due tomorrow" : `Due ${shortDay(date)}`;
}
const mailLink = (id?: string): ActionLink | undefined => id ? { label: "Open in Outlook", url: `https://outlook.office.com/mail/deeplink/read/${encodeURIComponent(id)}` } : undefined;
const key = (title: string) => title.toLowerCase().replace(/^(re|fw|fwd):\s*/g, "").replace(/[^a-z0-9]+/g, " ").trim().slice(0, 48);

export function buildAttention(input: AttentionInput): AttentionItem[] {
  const { now } = input, out: AttentionItem[] = [];
  const add = (item: Omit<AttentionItem, "priority"> & { base: number; extra?: number }) => {
    const { base, extra = 0, ...rest } = item;
    out.push({ ...rest, priority: base + BOOST[rest.horizon] + extra });
  };
  const future = input.events.filter(event => (startOf(event) ?? now) > now || ((endOf(event) ?? now) > now && !event.isAllDay));
  const timed = future.filter(event => !event.isAllDay && event.responseType !== "declined" && event.showAs !== "free" && startOf(event) && endOf(event))
    .sort((a, b) => startOf(a)!.getTime() - startOf(b)!.getTime());
  const feedFor = (text: string) => (input.feeds ?? []).find(feed => text.toLowerCase().includes(feed.name.toLowerCase()) || feed.name.toLowerCase().split(/\s+/).filter(word => word.length > 4).some(word => text.toLowerCase().includes(word)));
  const eventFor = (title: string) => input.events.find(event => (event.subject ?? "").trim().toLowerCase() === title.trim().toLowerCase());
  const prepFor = (event: CalendarEvent) => input.meetings.find(item => item.kind === "prep" && item.title.trim().toLowerCase() === (event.subject ?? "").trim().toLowerCase());
  const usedPrep = new Set<WorkItem>();

  // ---- Calendar ----
  const next = timed.find(event => hasAttendees(event) && startOf(event)! > now && startOf(event)!.getTime() - now.getTime() <= 2 * 36e5);
  if (next) {
    const start = startOf(next)!, prep = isPrivateEvent(next) ? undefined : prepFor(next);
    if (prep) usedPrep.add(prep);
    add({ id: `att-next-${next.id}`, title: `Next meeting: ${eventTitle(next)} at ${clock(start)}`, horizon: "now", base: 30, extra: prep ? 10 : 0, cost: "free",
      sources: ["calendar"], detail: prep ? `Prep note: ${prep.summary}${prep.next ? ` Next: ${prep.next}` : ""}` : `Starts ${agoUntil(start, now)}${next.organizer ? ` · organized by ${firstName(next.organizer) || "someone"}` : ""}.`,
      open: { label: "Open the meeting", url: next.webLink || "https://outlook.office.com/calendar/view/day" },
      runLabel: "Build a prep brief", notesLabel: "Your own notes to include (optional)", copilotBrief: "prep" as const, run: async ({ notes }) => ({ message: "Here's your prep brief.", links: [], doc: meetingBrief(next, prep, feedFor(next.subject ?? ""), [], notes) }) });
  }
  for (const event of future.filter(event => event.responseType === "notResponded" && hasAttendees(event) && event.id && startOf(event)!.getTime() - now.getTime() < 7 * 864e5).slice(0, 6)) {
    const start = startOf(event)!;
    add({ id: `att-invite-${event.id}`, title: `Answer the invite: ${eventTitle(event)}`, horizon: horizonOfDate(start.toISOString(), now, "week"), base: 40, cost: "free", sources: ["calendar"],
      detail: `${shortDay(start)}, ${clock(start)}${event.organizer ? ` · from ${firstName(event.organizer) || "the organizer"}` : ""}. You haven't responded yet.`,
      open: { label: "Open the invite", url: event.webLink || "https://outlook.office.com/calendar/view/week" },
      choices: [{ value: "accept", label: "Accept" }, { value: "tentativelyAccept", label: "Tentative" }, { value: "decline", label: "Decline" }], notesLabel: "Note to the organizer (optional)",
      run: ({ choice, notes }) => respondToInvite(event, choice ?? "accept", notes) });
  }
  const tomorrowEnd = new Date(now); tomorrowEnd.setDate(now.getDate() + 2); tomorrowEnd.setHours(0, 0, 0, 0);
  const soon = timed.filter(event => startOf(event)! < tomorrowEnd);
  let overlaps = 0;
  for (let i = 0; i < soon.length && overlaps < 3; i++) for (let j = i + 1; j < soon.length && overlaps < 3; j++) {
    const a = soon[i], b = soon[j];
    if (startOf(b)! >= endOf(a)!) break;
    const choices = [a, b].filter(event => event.responseType !== "organizer" && event.id).map(event => ({ value: event.id!, label: `Decline "${eventTitle(event)}"` }));
    if (!choices.length) continue;
    overlaps++;
    add({ id: `att-overlap-${a.id}-${b.id}`, title: `Double-booked: ${eventTitle(a)} and ${eventTitle(b)}`, horizon: horizonOfDate(startOf(b)!.toISOString(), now, "today"), base: 35, cost: "free", sources: ["calendar"],
      detail: `${shortDay(startOf(b)!)}, ${clock(startOf(b)!)}. Pick one to decline, with an optional note.`, choices, open: { label: "Open the day in Outlook", url: `https://outlook.office.com/calendar/view/day/${startOf(b)!.getFullYear()}/${startOf(b)!.getMonth() + 1}/${startOf(b)!.getDate()}` }, notesLabel: "Note to the organizer (optional)",
      run: ({ choice, notes }) => respondToInvite([a, b].find(event => event.id === choice) ?? a, "decline", notes) });
  }
  for (const away of input.collaboration.filter(item => item.kind === "oof" && item.relatedMeeting).slice(0, 3)) {
    const horizon = horizonOfDate(away.date, now, "week");
    if (horizon === "overdue" || horizon === "later") continue;
    add({ id: `att-away-${away.title}-${away.relatedMeeting}`, title: `${firstName(away.title) || away.title} is away for "${away.relatedMeeting}"`, horizon, base: 30, cost: "free", sources: ["outlook", "calendar"],
      detail: `Away ${away.date}${away.endDate ? ` to ${away.endDate}` : ""}. Plan around it or propose a new time.`, open: openFor(away),
      runLabel: "Save a draft", notesLabel: "Your message (saved as a draft, not sent)", copilot: ({ notes }) => awayRequest(away, notes), run: ({ notes }) => draftNewEmail(away.email, notes, `Rescheduling: ${away.relatedMeeting}`) });
  }

  // ---- Email ----
  for (const item of input.replies.slice(0, 8)) add({ id: `att-reply-${item.title}`, title: `Reply needed: ${item.title}`, horizon: "today", base: 42, cost: "free", sources: ["outlook"],
    detail: [item.summary, item.next && `Next: ${item.next}`].filter(Boolean).join(" "), open: itemLink(item) ? openFor(item) : outlookSearch(item.title),
    runLabel: "Save a reply draft", notesLabel: "Your reply (saved as a draft in the thread, not sent)", copilot: ({ notes }) => replyRequest(item, notes), run: ({ notes }) => draftReply(item, notes) });
  for (const mail of input.unread.filter(mail => mail.id && mail.importance === "high").slice(0, 4)) add({ id: `att-important-${mail.id}`, title: `High-importance email: ${hideEmails(mail.subject ?? "") || "No subject"}`, horizon: "today", base: 38, cost: "free", sources: ["outlook"],
    detail: `${firstName(mail.from) ? `From ${firstName(mail.from)}. ` : ""}${hideEmails(mail.bodyPreview ?? "").slice(0, 140)}`, open: mailLink(mail.id),
    choices: [{ value: "flag", label: "Flag for follow-up" }, { value: "read", label: "Mark as read" }], run: ({ choice }) => choice === "read" ? markRead(mail.id!, mail.subject ?? "") : flagMessage(mail.id!, mail.subject ?? "") });
  for (const mail of input.flagged.filter(mail => mail.id).slice(0, 4)) add({ id: `att-flagged-${mail.id}`, title: `Flagged email: ${hideEmails(mail.subject ?? "") || "No subject"}`, horizon: "week", base: 28, cost: "free", sources: ["outlook"],
    detail: `${firstName(mail.from) ? `From ${firstName(mail.from)}. ` : ""}You flagged this for follow-up.`, open: mailLink(mail.id) });

  // ---- Files that need you (linked in a reply you owe, a chat waiting on you, or a meeting in the next day) ----
  for (const file of (input.files ?? []).filter(file => file.origins.some(origin => origin === "reply" || origin === "chat" || origin === "meeting")).slice(0, 2)) add({
    id: `att-file-${file.key}`, title: `Read: ${file.name}`, horizon: file.origins.includes("meeting") ? "tomorrow" : "today", base: 24, cost: "free", sources: ["onedrive"],
    detail: file.reasons.slice(0, 2).join(" · "), open: { label: "Open the file", url: file.url },
    notesLabel: "Task notes (optional)", runLabel: "Add a review task to To Do", run: ({ notes }) => addToDo(`Review ${file.name}`.slice(0, 250), undefined, notes, file.url), copilot: ({ notes }) => fileRequest(file, notes) });
  // ---- Teams ----
  input.waiting.slice(0, 5).forEach(chat => add({ id: `att-teams-${chat.chatId}`, title: `${chat.person} is waiting on you in Teams`, horizon: "today", base: 36, extra: Date.parse(chat.created) < now.getTime() - 864e5 ? 5 : 0, cost: "free", sources: ["teams"],
    detail: `"${chat.preview}" · ${agoLabel(chat.created, now)}${chat.group ? ` · ${chat.chatTitle}` : ""}`, open: { label: "Open in Teams", url: chat.messageUrl }, notesLabel: "What you want to say (optional)", copilot: ({ notes }) => teamsReplyRequest(chat, notes) }));

  // ---- To Do ----
  for (const task of input.tasks) {
    const state = taskState(task, now);
    const horizon: Horizon | undefined = state === "overdue" ? "overdue" : state === "today" ? "today" : state === "tomorrow" ? "tomorrow" : state === "week" ? "week" : task.importance === "high" ? "later" : undefined;
    if (!horizon) continue;
    add({ id: `att-todo-${task.id}`, title: task.title, horizon, base: 34, extra: task.importance === "high" ? 8 : 0, cost: "free", sources: ["todo"],
      detail: [dueText(task.due, horizon) || "No due date", task.importance === "high" ? "High importance" : "", `List: ${task.list}`].filter(Boolean).join(" · "),
      open: { label: "Open in To Do", url: `${TODO_HOME}${encodeURIComponent(task.listId)}` } });
  }

  // ---- OneNote ----
  const recentPages = input.pages.filter(page => page.todos.length && (!page.modified || Date.parse(page.modified) > now.getTime() - 7 * 864e5));
  for (const page of recentPages.slice(0, 3)) for (const todo of page.todos.slice(0, 2)) add({ id: `att-onenote-${page.id}-${todo}`, title: todo, horizon: "later", base: 22, cost: "free", sources: ["onenote"],
    detail: `Unchecked to-do on "${page.title}" (${page.notebook}).`, open: { label: "Open the page", url: page.url || "https://www.office.com/launch/onenote" },
    runLabel: "Add to To Do", notesLabel: "Task notes (optional)", run: ({ notes }) => addToDo(todo, undefined, notes, page.url) });

  // ---- Team activity: urgent posts, deadlines, projects; prep, people, files ----
  for (const item of input.priorities) {
    const link = itemLink(item), open = openFor(item);
    if (item.kind === "urgent" || item.kind === "action") {
      const horizon = horizonOfDate(item.date, now, item.kind === "urgent" ? "today" : "later");
      if (item.kind === "action" && horizon === "later") continue;
      add({ id: `att-${item.kind}-${item.title}-${item.date}`, title: item.title, horizon, base: item.kind === "urgent" ? 45 : 33, extra: RISK.test(`${item.severity} ${item.status}`) ? 5 : 0, cost: "free", sources: copilotSources(item),
        detail: [dueText(item.date, horizon), item.summary, item.next && `Next: ${item.next}`].filter(Boolean).join(" · "), open,
        runLabel: "Add to To Do", notesLabel: "Task notes (optional)", notesDefault: item.next, run: ({ notes }) => addToDo(item.title, item.date, notes, link) });
    } else if (item.kind === "project" && RISK.test(`${item.status} ${item.severity} ${item.summary}`)) {
      add({ id: `att-project-${item.title}`, title: `Project needs attention: ${item.title}`, horizon: "week", base: 34, cost: "free", sources: copilotSources(item),
        detail: [item.status && `Status: ${item.status}`, item.summary, item.next && `Next: ${item.next}`].filter(Boolean).join(" · "), open,
        runLabel: "Build a project briefing", notesLabel: "Your own notes to include (optional)", copilotBrief: "project", run: async ({ notes }) => ({ message: "Here's your project briefing.", links: [], doc: projectBrief(item, feedFor(item.title), input.priorities.filter(row => row.kind === "action" && row.evidence.includes(item.title)), [], [], notes) }) });
    }
  }
  for (const prep of input.meetings.filter(item => item.kind === "prep" && !usedPrep.has(item)).slice(0, 3)) add({ id: `att-prep-${prep.title}`, title: `Prepare for: ${prep.title}`, horizon: "tomorrow", base: 30, cost: "free", sources: ["calendar"],
    detail: [prep.summary, prep.next && `Prepare: ${prep.next}`].filter(Boolean).join(" "), open: openFor(prep),
    ...(eventFor(prep.title) ? { runLabel: "Build a prep brief", notesLabel: "Your own notes to include (optional)", copilotBrief: "prep" as const, run: async ({ notes }: { notes: string }) => ({ message: "Here's your prep brief.", links: [], doc: meetingBrief(eventFor(prep.title)!, prep, feedFor(prep.title), [], notes) }) } : {}) });
  for (const person of input.collaboration.filter(item => item.kind === "person").slice(0, 4)) add({ id: `att-person-${person.title}`, title: `Follow up with ${firstName(person.title) || person.title}`, horizon: "week", base: 30, cost: "free", sources: ["outlook"],
    detail: [person.summary, person.next && `Next: ${person.next}`].filter(Boolean).join(" "), open: openFor(person),
    runLabel: "Save a follow-up draft", notesLabel: "Your follow-up (saved as a draft, not sent)", copilot: ({ notes }) => followUpRequest(person, notes), run: ({ notes }) => draftNewEmail(person.email, notes, `Following up: ${person.summary.match(/"([^"]+)"/)?.[1] ?? "my earlier email"}`) });
  for (const file of input.collaboration.filter(item => item.kind === "file").slice(0, 2)) add({ id: `att-file-${file.title}`, title: `Recently active: ${file.title}`, horizon: "later", base: 12, cost: "free", sources: copilotSources(file),
    detail: [file.date && `Changed ${shortDay(new Date(file.date))}`, file.summary].filter(Boolean).join(" · "), open: itemLink(file) ? { label: "Open the folder", url: itemLink(file)! } : undefined });

  // ---- Protect time when a lot is due ----
  const pressing = out.filter(item => item.horizon === "overdue" || item.horizon === "today").length;
  const gap = input.calendarId ? freeGaps(input.events, new Date(now.getFullYear(), now.getMonth(), now.getDate()), now, 60)[0] : undefined;
  if (gap && pressing >= 3) {
    const end = new Date(Math.min(gap.end.getTime(), gap.start.getTime() + 90 * 60000));
    add({ id: "att-focus", title: `Block ${clock(gap.start)}–${clock(end)} to clear ${pressing} pressing items`, horizon: "today", base: 26, cost: "free", sources: ["calendar"],
      detail: "Your largest free window today. Adds a private busy hold so meetings don't land there.", notesLabel: "Event title", notesDefault: "Focus: clear urgent items",
      runLabel: "Block the time", run: ({ notes }) => blockTime(input.calendarId!, gap.start, end, notes.trim() || "Focus time", "") });
  }

  return merge(out);
}

function agoUntil(start: Date, now: Date) {
  const minutes = Math.max(0, Math.round((start.getTime() - now.getTime()) / 60000));
  return minutes < 60 ? `in ${minutes} min` : `in ${Math.round(minutes / 6) / 10} h`;
}

/** The same task found in several places (for example a team post and To Do) becomes one item listing every source. */
function merge(items: AttentionItem[]): AttentionItem[] {
  const byKey = new Map<string, AttentionItem>();
  for (const item of items) {
    const name = key(item.title.replace(/^(reply needed|answer the invite|project needs attention|prepare for|follow up with|flagged email|high-importance email):\s*/i, ""));
    const existing = name ? byKey.get(name) : undefined;
    if (!existing) { byKey.set(name || item.id, item); continue; }
    const [keep, other] = existing.priority >= item.priority ? [existing, item] : [item, existing];
    const runner = keep.run ? keep : other;
    byKey.set(name, { ...keep, run: runner.run, runLabel: runner.runLabel, cost: runner.cost, choices: runner.choices, notesLabel: runner.notesLabel ?? other.notesLabel, notesDefault: runner.notesDefault, copilot: runner.copilot ?? other.copilot, copilotBrief: runner.copilotBrief, open: keep.open ?? other.open, priority: keep.priority + 3,
      sources: [...new Set([...keep.sources, ...other.sources])], horizon: HORIZONS.findIndex(h => h.id === other.horizon) < HORIZONS.findIndex(h => h.id === keep.horizon) ? other.horizon : keep.horizon });
  }
  return [...byKey.values()].sort((a, b) => b.priority - a.priority);
}
