import { PRIORITY, type SuggestedAction } from "@/lib/actions";
import { addToDo, blockTime, downloadDocument, draftNewEmail, draftReply, flagMessage, listHtml, markRead, oneNotePage, respondToInvite, setAutoReplies, todoHtml } from "@/lib/action-runners";
import { eventTitle, hasAttendees, isAcceptedMeeting, isPrivateEvent, type CalendarEvent } from "@/lib/command-center-evidence";
import type { WorkItem } from "@/lib/command-center-data";
import type { OneNotePage } from "@/lib/direct-sources";
import { firstName, hideEmails } from "@/lib/privacy";
import { agoLabel, type WaitingChat } from "@/lib/teams-waiting";
import type { TeamFeed } from "@/lib/direct-signals";
import type { OpenTask } from "@/lib/todo";
import { meetingBrief, statusReport, teamsCatchUp } from "@/lib/briefs";
import { awayRequest, fileRequest, followUpRequest, replyRequest, teamsReplyRequest } from "@/lib/copilot-prompts";
import type { FileItem } from "@/lib/files";

// ---------- time helpers ----------
export function instant(value?: string) {
  if (!value || !/(Z|[+-]\d{2}:\d{2})$/i.test(value)) return undefined;
  const date = new Date(value); return Number.isNaN(date.getTime()) ? undefined : date;
}
export const startOf = (event: CalendarEvent) => instant(event.startWithTimeZone ?? event.start);
export const endOf = (event: CalendarEvent) => instant(event.endWithTimeZone ?? event.end);
const clock = (date: Date) => date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
const dayName = (date: Date) => date.toLocaleDateString([], { weekday: "long", month: "short", day: "numeric" });
const sameDay = (a: Date, b: Date) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
function midnight(date: Date) { const value = new Date(date); value.setHours(0, 0, 0, 0); return value; }
function busy(event: CalendarEvent) { return !event.isAllDay && event.responseType !== "declined" && event.showAs !== "free"; }
function dueState(date: string, now: Date): "overdue" | "today" | "soon" | "later" | undefined {
  if (!date) return undefined;
  const due = /^\d{4}-\d{2}-\d{2}$/.test(date) ? new Date(`${date}T23:59:59`) : new Date(date);
  if (Number.isNaN(due.getTime())) return undefined;
  if (due < now && !sameDay(due, now)) return "overdue";
  if (sameDay(due, now)) return "today";
  return due.getTime() - now.getTime() < 7 * 864e5 ? "soon" : "later";
}
const dueLabel = (date: string) => new Date(/^\d{4}-\d{2}-\d{2}$/.test(date) ? `${date}T12:00:00` : date).toLocaleDateString([], { month: "short", day: "numeric" });

/** Free gaps of at least `minutes` between 9:00 and 17:00 on a day, after `from`. */
export function freeGaps(events: CalendarEvent[], day: Date, from: Date, minutes = 60) {
  const open = new Date(day); open.setHours(9, 0, 0, 0);
  const close = new Date(day); close.setHours(17, 0, 0, 0);
  let cursor = new Date(Math.max(open.getTime(), from.getTime()));
  cursor.setMinutes(Math.ceil(cursor.getMinutes() / 15) * 15, 0, 0);
  const blocks = events.filter(busy).map(event => [startOf(event), endOf(event)] as const)
    .filter((pair): pair is readonly [Date, Date] => !!pair[0] && !!pair[1] && pair[1] > open && pair[0] < close)
    .sort((a, b) => a[0].getTime() - b[0].getTime());
  const gaps: { start: Date; end: Date }[] = [];
  for (const [start, end] of blocks) {
    if (start.getTime() - cursor.getTime() >= minutes * 60000) gaps.push({ start: new Date(cursor), end: new Date(start) });
    if (end > cursor) cursor = new Date(end);
  }
  if (close.getTime() - cursor.getTime() >= minutes * 60000) gaps.push({ start: new Date(cursor), end: close });
  return gaps;
}

// ---------- Plan my day (free, from refreshed data) ----------
export interface PlanStep { time?: string; text: string; tone: "meeting" | "urgent" | "reply" | "prep" | "focus" | "task" }
export function buildDayPlan(events: CalendarEvent[], work: WorkItem[], replies: WorkItem[], now = new Date()): PlanStep[] {
  const today = events.filter(event => { const start = startOf(event); return start && sameDay(start, now) && busy(event); })
    .sort((a, b) => (startOf(a)?.getTime() ?? 0) - (startOf(b)?.getTime() ?? 0));
  const prepTitles = new Map(work.filter(item => item.kind === "prep").map(item => [item.title.trim().toLowerCase(), item]));
  const steps: PlanStep[] = [];
  const overdue = work.filter(item => (item.kind === "action" || item.kind === "urgent") && dueState(item.date, now) === "overdue");
  const dueToday = work.filter(item => (item.kind === "action" || item.kind === "urgent") && dueState(item.date, now) === "today");
  const urgent = work.filter(item => item.kind === "urgent" && !item.date);
  if (overdue.length) steps.push({ text: `Clear overdue: ${overdue.slice(0, 3).map(item => `${item.title} (due ${dueLabel(item.date)})`).join("; ")}`, tone: "urgent" });
  for (const item of [...dueToday, ...urgent].slice(0, 3)) steps.push({ text: `Today: ${item.title}${item.next ? ` — ${item.next}` : ""}`, tone: "urgent" });
  const upcoming = today.filter(event => (endOf(event) ?? now) > now);
  for (const event of upcoming.slice(0, 6)) {
    const start = startOf(event)!, prep = prepTitles.get((event.subject ?? "").trim().toLowerCase());
    if (prep && !isPrivateEvent(event) && start > now) steps.push({ time: `Before ${clock(start)}`, text: `Prep for ${eventTitle(event)}: ${prep.next || prep.summary}`, tone: "prep" });
    steps.push({ time: clock(start), text: eventTitle(event), tone: "meeting" });
  }
  const gaps = freeGaps(events, now, now, 30);
  if (replies.length && gaps[0]) steps.push({ time: `${clock(gaps[0].start)}`, text: `Reply to ${replies.length} email${replies.length === 1 ? "" : "s"} waiting on you (see Email → Actions for drafts)`, tone: "reply" });
  const longGap = gaps.find(gap => gap.end.getTime() - gap.start.getTime() >= 60 * 60000 && gap !== gaps[0]) ?? (gaps[0] && !replies.length ? gaps[0] : undefined);
  const project = work.find(item => item.kind === "project");
  if (longGap) steps.push({ time: `${clock(longGap.start)}–${clock(longGap.end)}`, text: `Focus block${project ? `: move ${project.title} forward${project.next ? ` (${project.next})` : ""}` : ""}`, tone: "focus" });
  return steps.sort((a, b) => {
    const order = (step: PlanStep) => step.tone === "urgent" ? -1 : 0;
    return order(a) - order(b);
  });
}

// ---------- per-component suggestions ----------
export interface Ctx { now: Date; calendarId?: string; mySite?: string; me?: string; feeds?: TeamFeed[]; tasks?: OpenTask[]; waiting?: WaitingChat[] }

export function calendarActions(ctx: Ctx, events: CalendarEvent[], meetings: WorkItem[], collaboration: WorkItem[]): SuggestedAction[] {
  const { now, calendarId } = ctx, out: SuggestedAction[] = [];
  const future = events.filter(event => (startOf(event) ?? now) > now);
  for (const event of future.filter(event => event.responseType === "notResponded" && hasAttendees(event) && event.id).slice(0, 5)) {
    const start = startOf(event)!;
    out.push({ id: `invite-${event.id}`, title: `Respond to invite: ${eventTitle(event)}`, detail: `${dayName(start)}, ${clock(start)}${event.organizer ? ` · from ${firstName(event.organizer) || "the organizer"}` : ""}`, priority: PRIORITY.invite, cost: "free",
      choices: [{ value: "accept", label: "Accept" }, { value: "tentativelyAccept", label: "Tentative" }, { value: "decline", label: "Decline" }], notesLabel: "Note to the organizer (optional)",
      run: ({ choice, notes }) => respondToInvite(event, choice ?? "accept", notes) });
  }
  const timed = future.filter(event => busy(event) && startOf(event) && endOf(event)).sort((a, b) => startOf(a)!.getTime() - startOf(b)!.getTime());
  const seen = new Set<string>();
  for (let i = 0; i < timed.length && out.filter(action => action.id.startsWith("overlap")).length < 3; i++) {
    for (let j = i + 1; j < timed.length; j++) {
      const a = timed[i], b = timed[j];
      if (startOf(b)! >= endOf(a)!) break;
      const key = `${a.id}|${b.id}`; if (seen.has(key)) continue; seen.add(key);
      const choices = [a, b].filter(event => event.responseType !== "organizer" && event.id).map(event => ({ value: event.id!, label: `Decline "${eventTitle(event)}"` }));
      if (!choices.length) continue;
      out.push({ id: `overlap-${key}`, title: `Overlap: ${eventTitle(a)} and ${eventTitle(b)}`, detail: `${dayName(startOf(b)!)}, ${clock(startOf(b)!)}. Pick one to decline (with an optional note).`, priority: PRIORITY.overlap, cost: "free", choices, notesLabel: "Note to the organizer (optional)",
        run: ({ choice, notes }) => respondToInvite([a, b].find(event => event.id === choice) ?? a, "decline", notes) });
    }
  }
  if (calendarId) for (const offset of [0, 1]) {
    const day = midnight(now); day.setDate(day.getDate() + offset);
    const gap = freeGaps(events, day, now, 60)[0];
    if (!gap) continue;
    const end = new Date(Math.min(gap.end.getTime(), gap.start.getTime() + 90 * 60000));
    out.push({ id: `focus-${day.toDateString()}`, title: `Block focus time ${offset ? "tomorrow" : "today"}, ${clock(gap.start)}–${clock(end)}`, detail: "Adds a private \"busy\" hold to your calendar so meetings don't land there.", priority: PRIORITY.focus + (offset ? 0 : 3), cost: "free",
      notesLabel: "What will you focus on? (used as the event title)", notesDefault: "Focus time",
      run: ({ notes }) => blockTime(calendarId, gap.start, end, notes.trim() ? (/^focus/i.test(notes.trim()) ? notes.trim() : `Focus: ${notes.trim()}`) : "Focus time", "") });
  }
  const next = timed.find(event => hasAttendees(event) && !isPrivateEvent(event) && startOf(event)!.getTime() - now.getTime() < 24 * 36e5);
  if (next) {
    const start = startOf(next)!, soon = start.getTime() - now.getTime() < 4 * 36e5;
    const prep = meetings.find(item => item.kind === "prep" && item.title.trim().toLowerCase() === (next.subject ?? "").trim().toLowerCase());
    out.push({ id: `brief-${next.id}`, title: `Prep brief for ${eventTitle(next)}`, detail: `${dayName(start)}, ${clock(start)}. A one-page brief from the invite, related emails and team posts — download it as Word or PDF.`, priority: soon ? PRIORITY.prepSoon : PRIORITY.document + 15, cost: "free", notesLabel: "Your own notes to include (optional)", runLabel: "Build the brief", copilotBrief: "prep",
      run: async ({ notes }) => ({ message: "Here's your prep brief.", links: [], doc: meetingBrief(next, prep, ctx.feeds?.find(feed => (next.subject ?? "").toLowerCase().includes(feed.name.toLowerCase().split(" ")[0])), [], notes) }) });
  }
  for (const away of collaboration.filter(item => item.kind === "oof" && item.relatedMeeting).slice(0, 3)) {
    out.push({ id: `away-${away.title}-${away.relatedMeeting}`, title: `${firstName(away.title) || away.title} is away for "${away.relatedMeeting}"`, detail: `Away ${dueLabel(away.date)}${away.endDate ? `–${dueLabel(away.endDate)}` : ""}. Draft a note to reschedule or plan around it.`, priority: PRIORITY.awayConflict, cost: "free", notesLabel: "Your message (saved as a draft, not sent)", runLabel: "Save the draft", copilot: ({ notes }) => awayRequest(away, notes),
      run: ({ notes }) => draftNewEmail(away.email, notes, `Rescheduling: ${away.relatedMeeting}`) });
  }
  return out;
}

export interface InboxMessage { id?: string; subject?: string; from?: string; isRead?: boolean }
export function emailActions(ctx: Ctx, replies: WorkItem[], unread: InboxMessage[], events: CalendarEvent[]): SuggestedAction[] {
  const out: SuggestedAction[] = [];
  for (const item of replies.slice(0, 10)) {
    out.push({ id: `reply-${item.title}-${item.url}`, title: `Reply: ${item.title}`, detail: item.summary, priority: PRIORITY.reply, cost: "free", notesLabel: "Your reply (saved as a draft in the thread, not sent)", runLabel: "Save the reply draft",
      open: item.url ? { label: "Open in Outlook", url: item.url } : undefined, copilot: ({ notes }) => replyRequest(item, notes), run: ({ notes }) => draftReply(item, notes) });
  }
  for (const message of unread.filter(message => message.id).slice(0, 5)) {
    out.push({ id: `triage-${message.id}`, title: `Triage: ${hideEmails(message.subject ?? "") || "No subject"}`, detail: "Flag it for follow-up or mark it as read.", priority: PRIORITY.niceToHave + 5, cost: "free",
      choices: [{ value: "flag", label: "Flag for follow-up" }, { value: "read", label: "Mark as read" }],
      run: ({ choice }) => choice === "read" ? markRead(message.id!, message.subject ?? "") : flagMessage(message.id!, message.subject ?? "") });
  }
  const away = events.find(event => { const start = startOf(event); return start && start > ctx.now && start.getTime() - ctx.now.getTime() < 14 * 864e5 && (event.showAs === "oof" || /\b(ooo|out of (the )?office|vacation|pto|annual leave|holiday)\b/i.test(event.subject ?? "")); });
  if (away) {
    const start = startOf(away)!, end = endOf(away) ?? new Date(start.getTime() + 864e5);
    out.push({ id: `oof-${away.id}`, title: `Turn on automatic replies for ${dayName(start)} – ${end.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" })}`, detail: "Your calendar shows you're away. Replies go to people inside your organization only.", priority: PRIORITY.followUp + 5, cost: "free",
      notesLabel: "Automatic reply message", notesDefault: `Thanks for your message. I'm out of the office from ${start.toLocaleDateString()} to ${end.toLocaleDateString()} with limited access to email. I'll reply when I'm back.`,
      run: ({ notes }) => setAutoReplies(start, end, notes) });
  }
  return out;
}

export function workActions(ctx: Ctx, work: WorkItem[]): SuggestedAction[] {
  const out: SuggestedAction[] = [];
  const dated = work.filter(item => (item.kind === "action" || item.kind === "urgent") && item.date)
    .map(item => ({ item, state: dueState(item.date, ctx.now) })).filter(entry => entry.state && entry.state !== "later").slice(0, 6);
  for (const { item, state } of dated) {
    out.push({ id: `todo-${item.title}-${item.date}`, title: `Add to To Do: ${item.title}`, detail: `${state === "overdue" ? "Overdue" : state === "today" ? "Due today" : "Due"} ${dueLabel(item.date)}${item.next ? ` · ${item.next}` : ""}`,
      priority: state === "overdue" ? PRIORITY.overdue - 5 : state === "today" ? PRIORITY.dueToday - 2 : PRIORITY.tasks, cost: "free", notesLabel: "Task notes (optional)", notesDefault: item.next,
      run: ({ notes }) => addToDo(item.title, item.date, notes, item.url) });
  }
  const tracked = work.filter(item => item.kind === "project" || item.kind === "action" || item.kind === "urgent");
  if (tracked.length) out.push({ id: "tracker", title: "Export a work tracker to Excel", detail: `${tracked.length} projects and actions, with status, due date, next step and source. Downloads as an Excel file.`, priority: PRIORITY.document, cost: "free",
    run: () => downloadDocument({ title: "Work tracker", sections: [] }, "xlsx", null, [["Type", "Item", "Status", "Urgency", "Due", "Next step", "Source"],
      ...tracked.map(item => [item.kind, item.title, item.status, item.severity, item.date ? dueLabel(item.date) : "", item.next, item.url])]) });
  const projects = work.filter(item => item.kind === "project");
  if (projects.length) out.push({ id: "status-report", title: "Build a weekly status report", detail: `One section per project (${projects.length}) from recent team posts and deadlines — download it as Word or PDF.`, priority: PRIORITY.document, cost: "free", notesLabel: "Your own notes to include (optional)", runLabel: "Build the report", copilotBrief: "status",
    run: async ({ notes }) => ({ message: "Here's your status report.", links: [], doc: statusReport(projects, ctx.feeds ?? [], work.filter(item => item.kind === "action"), midnight(ctx.now), notes) }) });
  return out;
}

export function teamsActions(waiting: WaitingChat[], now = new Date(), feeds: TeamFeed[] = []): SuggestedAction[] {
  // Only chats where someone else sent the latest message in the last 7 days (free Teams reads).
  const out: SuggestedAction[] = waiting.slice(0, 5).map((chat, index) => ({
    id: `teams-${chat.chatId}`, title: `Reply to ${chat.person} in Teams${chat.group ? ` (${chat.chatTitle})` : ""}`,
    detail: `"${chat.preview}" · ${agoLabel(chat.created, now)}.`, priority: PRIORITY.followUp + 2 - index, cost: "free" as const,
    open: { label: "Open in Teams", url: chat.messageUrl }, notesLabel: "What you want to say (optional)", copilot: ({ notes }) => teamsReplyRequest(chat, notes),
  }));
  out.push({ id: "teams-catchup", title: "Teams catch-up for the last 3 days", detail: "Who's waiting on you, and new posts in your project teams (with @mentions of you) — as a document you can download.", priority: PRIORITY.insight, cost: "free", runLabel: "Build the catch-up", copilotBrief: "catchup",
    run: async ({ notes }) => ({ message: "Here's your Teams catch-up.", links: [{ label: "Open Teams", url: "https://teams.microsoft.com/" }], doc: teamsCatchUp(waiting, feeds, now, notes) }) });
  return out;
}

export function peopleActions(collaboration: WorkItem[]): SuggestedAction[] {
  return collaboration.filter(item => item.kind === "person").slice(0, 5).map(person => ({
    id: `person-${person.url || person.title}`, title: `Follow up with ${firstName(person.title) || person.title}`, detail: person.summary, priority: PRIORITY.followUp, cost: "free" as const,
    open: person.url ? { label: "Open your email", url: person.url } : undefined, notesLabel: "Your follow-up (saved as a draft, not sent)", runLabel: "Save the draft", copilot: ({ notes }) => followUpRequest(person, notes),
    run: ({ notes }) => draftNewEmail(person.email, notes, `Following up: ${person.summary.match(/"([^"]+)"/)?.[1] ?? "my earlier email"}`),
  }));
}

export function filesActions(files: FileItem[]): SuggestedAction[] {
  return files.filter(file => file.score >= 20).slice(0, 6).map((file, index) => ({
    id: `file-${file.key}`, title: `Review: ${file.name}`, detail: file.reasons.slice(0, 2).join(" · "), priority: (file.score >= 38 ? PRIORITY.followUp : PRIORITY.document) - index, cost: "free" as const,
    sources: file.origins.includes("trending") && file.origins.length === 1 ? ["onedrive" as const] : ["onedrive" as const, ...(file.origins.some(origin => origin === "reply" || origin === "email" || origin === "meeting") ? ["outlook" as const] : []), ...(file.origins.some(origin => origin === "team" || origin === "chat" || origin === "mention") ? ["teams" as const] : [])],
    open: { label: "Open the file", url: file.url }, notesLabel: "Task notes (optional)", runLabel: "Add a review task to To Do",
    run: ({ notes }) => addToDo(`Review ${file.name}`.slice(0, 250), undefined, [notes.trim(), file.reasons[0]].filter(Boolean).join("\n"), file.url),
    copilot: ({ notes }) => fileRequest(file, notes),
  }));
}

export function oneNoteActions(ctx: Ctx, pages: OneNotePage[], plan: PlanStep[], events: CalendarEvent[]): SuggestedAction[] {
  const out: SuggestedAction[] = [];
  const date = ctx.now.toLocaleDateString([], { weekday: "long", month: "long", day: "numeric" });
  out.push({ id: "onenote-day", title: "Create today's plan page in OneNote", detail: "A page in Quick Notes with today's plan as a to-do checklist.", priority: PRIORITY.tasks - 5, cost: "free",
    run: () => oneNotePage(`Plan for ${date}`, `<h1>Plan for ${date}</h1>${todoHtml(plan.map(step => `${step.time ? `${step.time} — ` : ""}${step.text}`))}<h2>Notes</h2><p></p>`) });
  const next = events.filter(event => { const start = startOf(event); return start && start > ctx.now && hasAttendees(event) && !isPrivateEvent(event) && isAcceptedMeeting(event); })
    .sort((a, b) => startOf(a)!.getTime() - startOf(b)!.getTime())[0];
  if (next) {
    const start = startOf(next)!;
    out.push({ id: `onenote-meeting-${next.id}`, title: `Create meeting notes for ${eventTitle(next)}`, detail: `${dayName(start)}, ${clock(start)}. A page with agenda, notes, decisions and action items sections.`, priority: PRIORITY.tasks - 2, cost: "free",
      run: () => oneNotePage(`Meeting notes: ${eventTitle(next)}`, `<h1>${eventTitle(next)}</h1><p>${dayName(start)}, ${clock(start)}${next.organizer ? ` · Organizer: ${firstName(next.organizer)}` : ""}</p><h2>Agenda</h2><p></p><h2>Notes</h2><p></p><h2>Decisions</h2><p></p><h2>Action items</h2>${todoHtml(["", ""])}`) });
  }
  const todos = pages.flatMap(page => page.todos.map(todo => ({ todo, page }))).slice(0, 5);
  for (const { todo, page } of todos) out.push({ id: `onenote-todo-${page.id}-${todo}`, title: `Send to To Do: ${todo}`, detail: `From "${page.title}"`, priority: PRIORITY.tasks, cost: "free", notesLabel: "Task notes (optional)",
    run: ({ notes }) => addToDo(todo, undefined, notes, page.url) });
  return out;
}

export function insightActions(ctx: Ctx, events: CalendarEvent[], totals: { name: string; share: number }[], targets: Record<string, number>): SuggestedAction[] {
  const out: SuggestedAction[] = [];
  const days = new Map<string, { day: Date; count: number }>();
  for (const event of events.filter(isAcceptedMeeting)) {
    const start = startOf(event); if (!start || start < midnight(ctx.now)) continue;
    const key = start.toDateString(), entry = days.get(key) ?? { day: midnight(start), count: 0 };
    entry.count += 1; days.set(key, entry);
  }
  for (const { day, count } of [...days.values()].filter(entry => entry.count >= 7)) {
    out.push({ id: `heavy-${day.toDateString()}`, title: `${dayName(day)} is heavy: ${count} meetings`, detail: "Open that day in Outlook to decline, shorten or delegate meetings where you're optional.", priority: PRIORITY.insight + 5, cost: "free", runLabel: "Open the day in Outlook",
      run: async () => ({ message: "Opened links to that day. Review which meetings you can drop or shorten.", links: [{ label: `Open ${dayName(day)} in Outlook`, url: `https://outlook.office.com/calendar/view/day/${day.getFullYear()}/${day.getMonth() + 1}/${day.getDate()}` }] }) });
  }
  if (ctx.calendarId) for (const [name, target] of Object.entries(targets)) {
    const share = totals.find(total => total.name.toLowerCase() === name.toLowerCase())?.share ?? 0;
    if (target <= 0 || share >= target) continue;
    const day = midnight(ctx.now); day.setDate(day.getDate() + 1);
    const gap = freeGaps(events, day, ctx.now, 60)[0];
    if (!gap) continue;
    const end = new Date(gap.start.getTime() + 60 * 60000);
    out.push({ id: `target-${name}`, title: `Below your ${name} target (${Math.round(share)}% of ${target}%)`, detail: `Block an hour for ${name} tomorrow, ${clock(gap.start)}–${clock(end)}, tagged with your "${name}" category.`, priority: PRIORITY.insight, cost: "free",
      notesLabel: "Event title", notesDefault: name, run: ({ notes }) => blockTime(ctx.calendarId!, gap.start, end, notes.trim() || name, "", name) });
  }
  return out;
}

export function todayActions(ctx: Ctx, plan: PlanStep[]): SuggestedAction[] {
  const date = ctx.now.toLocaleDateString([], { weekday: "long", month: "long", day: "numeric" });
  return [
    { id: "plan-onenote", title: "Save today's plan to OneNote", detail: "Creates a checklist page in Quick Notes from the plan shown in the Today tab.", priority: PRIORITY.tasks - 5, cost: "free",
      run: () => oneNotePage(`Plan for ${date}`, `<h1>Plan for ${date}</h1>${todoHtml(plan.map(step => `${step.time ? `${step.time} — ` : ""}${step.text}`))}`) },
  ];
}
export { listHtml };
