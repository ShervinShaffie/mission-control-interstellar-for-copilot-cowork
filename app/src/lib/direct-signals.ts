/**
 * Work signals built only from the six direct Microsoft 365 connections (Outlook, Teams, To Do, OneNote,
 * OneDrive, Office 365 Users). No Copilot / Work IQ. Every rule here is plain, explainable logic over data
 * you can open, so results are stable and free.
 */
import { MicrosoftTeamsService } from "../../generated/services/MicrosoftTeamsService";
import { Office365OutlookService } from "../../generated/services/Office365OutlookService";
import { Office365UsersService } from "../../generated/services/Office365UsersService";
import type { WorkItem } from "@/lib/command-center-data";
import type { CalendarEvent } from "@/lib/command-center-evidence";
import { eventTitle, hasAttendees, isPrivateEvent } from "@/lib/command-center-evidence";
import { firstName, hideEmails, rememberFirstName } from "@/lib/privacy";
import { fileLinks, filesFromAttachments, filesFromHtml, type FileHit, type FileRef } from "@/lib/files";

const DAY = 864e5;
const item = (row: Partial<WorkItem> & Pick<WorkItem, "kind" | "title">): WorkItem => ({ summary: "", next: "", status: "", severity: "", date: "", endDate: "", email: "", url: "", evidence: "", relatedMeeting: "", ...row });
const rec = (value: unknown): Record<string, unknown> => (value && typeof value === "object" ? value as Record<string, unknown> : {});
const str = (value: unknown) => (typeof value === "string" ? value : "");
export function plain(html: string) {
  return (html || "").replace(/<style[\s\S]*?<\/style>/gi, " ").replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, "\"").replace(/&#39;/g, "'").replace(/\s+/g, " ").trim();
}
const clip = (text: string, max = 180) => (text.length > max ? `${text.slice(0, max - 1).trimEnd()}…` : text);
const shortDay = (date: Date) => date.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" });
const isoDay = (date: Date) => date.toLocaleDateString("en-CA");
async function limit<T, R>(items: T[], size: number, work: (item: T) => Promise<R>): Promise<R[]> {
  const out: R[] = new Array(items.length); let next = 0;
  await Promise.all(Array.from({ length: Math.min(size, items.length) }, async () => { while (next < items.length) { const i = next++; out[i] = await work(items[i]); } }));
  return out;
}
function ok<T>(result: { success: boolean; data: T; error?: unknown }, label: string): T {
  if (!result.success) { const error = new Error(`${label} unavailable`); (error as Error & { cause?: unknown }).cause = result.error; throw error; }
  return result.data;
}

// ---------- Plain-language patterns ----------
export const STATUS = /\b(RED|AMBER|GREEN)\b|\b(at risk|blocked|blocker|escalat\w*|overdue|missed|delayed|slipp\w*|behind schedule|on track)\b/i;
const REQUEST = /\?|\b(please|pls|can you|could you|would you|will you|need you|need your|your (input|review|approval|thoughts|feedback|sign-?off)|let me know|review|approve|approval|confirm|action required|asap|by (mon|tue|wed|thu|fri|eod|end of day|tomorrow|\d))/i;
const AUTOMATED = /no-?reply|do-?not-?reply|donotreply|notifications?@|mailer-daemon|postmaster|bounce|alerts?@|digest@/i;
const NOT_A_REQUEST = /^(accepted|declined|tentative|canceled|cancelled|updated invitation|invitation|automatic reply|out of office|undeliverable|read:|delivered:)\b/i;
const LEARNING = /\b(training|workshop|webinar|learning|course|bootcamp|certification|lunch (and|&) learn|onboarding|masterclass|skills?)\b/i;
const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
const WEEKDAYS = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];
const MONTH = "(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\\.?";
const DUE = new RegExp(`\\b(due|by|deadline|before|target|until|no later than|lands?)\\s+(?:on\\s+)?(?:(\\d{1,2})(?:st|nd|rd|th)?\\s+${MONTH}|${MONTH}\\s+(\\d{1,2})(?:st|nd|rd|th)?|(mon|tue|wed|thu|fri|sat|sun)[a-z]*(?:day)?|(today|tomorrow|eod|end of (?:the )?day|end of (?:the )?week))`, "gi");

/** Finds "due 23 Oct", "by Friday", "deadline Oct 30" in text, dated relative to when it was written. */
export function findDeadlines(text: string, written: Date): { date: Date; sentence: string }[] {
  const out: { date: Date; sentence: string }[] = [];
  const sentences = text.split(/(?<=[.!?])\s+|\n+/);
  for (const sentence of sentences) {
    for (const match of sentence.matchAll(DUE)) {
      let date: Date | undefined;
      const [, , d1, m1, m2, d2, weekday, relative] = match;
      if ((d1 && m1) || (m2 && d2)) {
        const month = MONTHS.indexOf((m1 ?? m2)!.slice(0, 3).toLowerCase()), day = Number(d1 ?? d2);
        if (month < 0 || day < 1 || day > 31) continue;
        // Pick the year that puts the date closest after (or just before) when it was written.
        const candidates = [written.getFullYear() - 1, written.getFullYear(), written.getFullYear() + 1].map(year => new Date(year, month, day, 17));
        date = candidates.find(candidate => candidate.getTime() >= written.getTime() - 60 * DAY) ?? candidates[1];
      } else if (weekday) {
        const target = WEEKDAYS.indexOf(weekday.slice(0, 3).toLowerCase());
        date = new Date(written); date.setHours(17, 0, 0, 0);
        const ahead = (target - date.getDay() + 7) % 7 || 7; date.setDate(date.getDate() + ahead);
      } else if (relative) {
        date = new Date(written); date.setHours(17, 0, 0, 0);
        if (/tomorrow/i.test(relative)) date.setDate(date.getDate() + 1);
        if (/week/i.test(relative)) date.setDate(date.getDate() + ((5 - date.getDay() + 7) % 7));
      }
      if (date) out.push({ date, sentence: clip(sentence.trim(), 200) });
    }
  }
  return out;
}

// ---------- Teams: projects from the teams you belong to ----------
export interface TeamPost { id: string; team: string; teamId: string; channel: string; channelId: string; author: string; authorId: string; created: Date; text: string; url: string; mentionsMe: boolean; important: boolean; files: FileRef[] }
export interface TeamFeed { id: string; name: string; description: string; channels: { id: string; name: string; webUrl: string; filesUrl: string }[]; posts: TeamPost[]; error?: boolean }
const PROJECTISH = /\b(launch|rollout|roll-out|expansion|project|program|programme|readiness|initiative|migration|implementation|upgrade|pilot|transformation|hardening|integration|go-live|deployment)\b/i;
const MAX_TEAMS = 8, MAX_CHANNELS = 3, POST_DAYS = 21;

/** Which teams to read: the ones you chose, else project-sounding teams first (max 8). */
export function chooseTeams<T extends { id: string; name: string }>(teams: T[], chosen?: string[]): T[] {
  if (chosen?.length) return teams.filter(team => chosen.includes(team.id)).slice(0, 12);
  return [...teams].sort((a, b) => Number(PROJECTISH.test(b.name)) - Number(PROJECTISH.test(a.name))).slice(0, MAX_TEAMS);
}
export async function loadJoinedTeamList(): Promise<{ id: string; name: string; description: string }[]> {
  const data = rec(ok(await MicrosoftTeamsService.GetAllTeams(), "Teams"));
  const value = Array.isArray(data.value) ? data.value : [];
  return value.map(rec).filter(team => str(team.id) && str(team.displayName)).map(team => ({ id: str(team.id), name: hideEmails(str(team.displayName)), description: hideEmails(str(team.description)) }));
}
export async function loadTeamFeeds(myId: string, chosen?: string[], now = new Date()): Promise<TeamFeed[]> {
  const teams = chooseTeams(await loadJoinedTeamList(), chosen);
  const cutoff = now.getTime() - POST_DAYS * DAY;
  return limit(teams, 3, async (team): Promise<TeamFeed> => {
    try {
      const channelData = ok(await MicrosoftTeamsService.GetChannelsForGroup(team.id), "Channels");
      const channels = (channelData.value ?? []).filter(channel => channel.id).slice(0, MAX_CHANNELS)
        .map(channel => ({ id: channel.id!, name: hideEmails(channel.displayName ?? "Channel"), webUrl: channel.webUrl ?? "", filesUrl: channel.filesFolderWebUrl ?? "" }));
      const lists = await limit(channels, 2, async channel => {
        const result = await MicrosoftTeamsService.GetMessagesFromChannel(team.id, channel.id);
        if (!result.success) return [];
        return (result.data?.value ?? []).flatMap((message): TeamPost[] => {
          const created = new Date(message.createdDateTime ?? "");
          const files = [...filesFromAttachments(rec(message).attachments), ...filesFromHtml(message.body?.content ?? "")].slice(0, 6);
          const text = hideEmails(plain(message.body?.content ?? "")) || hideEmails(message.subject ?? "") || (files.length ? `Shared ${files.map(file => file.name).join(", ")}` : "");
          if (!message.id || message.deleted || message.messageType !== "message" || Number.isNaN(created.getTime()) || created.getTime() < cutoff || !text) return [];
          const mentions = (message.mentions ?? []).map(rec);
          const mentionsMe = mentions.some(mention => str(rec(rec(mention.mentioned).user).id) === myId);
          const webUrl = str(rec(message).webUrl);
          const url = webUrl || `https://teams.microsoft.com/l/message/${encodeURIComponent(channel.id)}/${encodeURIComponent(message.id)}?groupId=${encodeURIComponent(team.id)}&parentMessageId=${encodeURIComponent(message.id)}&teamName=${encodeURIComponent(team.name)}&channelName=${encodeURIComponent(channel.name)}`;
          return [{ id: message.id, team: team.name, teamId: team.id, channel: channel.name, channelId: channel.id, author: firstName(message.from?.user?.displayName) || "Someone", authorId: message.from?.user?.id ?? "", created, text, url, mentionsMe, important: message.importance === "high" || message.importance === "urgent", files }];
        });
      });
      return { ...team, channels, posts: lists.flat().sort((a, b) => b.created.getTime() - a.created.getTime()) };
    } catch {
      return { ...team, channels: [], posts: [], error: true };
    }
  });
}

/** Projects, deadlines and urgent items from team activity. */
export function prioritiesFromTeams(feeds: TeamFeed[], myId: string, now = new Date()): WorkItem[] {
  const out: WorkItem[] = [];
  const seen = new Set<string>();
  for (const feed of feeds) {
    if (!feed.posts.length && !feed.channels.length) continue;
    const latest = feed.posts[0];
    const statusPost = feed.posts.find(post => STATUS.test(post.text));
    const statusWord = statusPost?.text.match(STATUS)?.[0] ?? "";
    const upcoming = feed.posts.flatMap(post => findDeadlines(post.text, post.created).map(found => ({ ...found, post })))
      .filter(found => found.date.getTime() > now.getTime() - 30 * DAY && found.date.getTime() < now.getTime() + 60 * DAY)
      .sort((a, b) => a.date.getTime() - b.date.getTime());
    const nextDue = upcoming.find(found => found.date >= now);
    out.push(item({
      kind: "project", title: feed.name,
      status: statusWord ? statusWord.toUpperCase().replace(/^(AT RISK|ON TRACK)$/, value => value.toLowerCase()) : "",
      summary: latest ? `${latest.author} in ${latest.channel}, ${shortDay(latest.created)}: ${clip(latest.text, 220)}` : "No posts in the last three weeks.",
      next: nextDue ? `${shortDay(nextDue.date)}: ${nextDue.sentence}` : "",
      url: (statusPost ?? latest)?.url || feed.channels[0]?.webUrl || "",
      evidence: `Teams › ${feed.name} · ${feed.posts.length} post${feed.posts.length === 1 ? "" : "s"} in the last 3 weeks`,
    }));
    for (const found of upcoming.slice(0, 6)) {
      const key = `${found.sentence.toLowerCase()}|${isoDay(found.date)}`;
      if (seen.has(key)) continue; seen.add(key);
      out.push(item({ kind: "action", title: clip(found.sentence, 140), date: isoDay(found.date), summary: `${found.post.author} in ${feed.name} › ${found.post.channel}, ${shortDay(found.post.created)}`, url: found.post.url, evidence: `Teams › ${feed.name} › ${found.post.channel}`, status: STATUS.test(found.sentence) ? (found.sentence.match(STATUS)?.[0] ?? "") : "" }));
    }
    for (const post of feed.posts.filter(post => post.created.getTime() > now.getTime() - 3 * DAY && post.authorId !== myId && (post.mentionsMe || post.important || /\b(RED|escalat\w*|blocked|urgent)\b/.test(post.text))).slice(0, 3)) {
      out.push(item({ kind: "urgent", title: post.mentionsMe ? `${post.author} mentioned you in ${feed.name} › ${post.channel}` : `${feed.name}: ${clip(post.text, 90)}`, summary: clip(post.text, 220), next: post.mentionsMe ? "Read and respond in Teams" : "", url: post.url, evidence: `Teams › ${feed.name} › ${post.channel}, ${shortDay(post.created)}`, severity: post.important ? "High importance" : "" }));
    }
  }
  return out;
}

// ---------- Outlook mail ----------
export interface MailRow { id: string; subject: string; from: string; to: string; cc: string; received: Date; preview: string; conversationId: string; importance: string; isRead: boolean; files: FileRef[]; hasAttachments: boolean }
function mailRows(data: unknown): MailRow[] {
  const value = rec(data).value;
  return (Array.isArray(value) ? value : []).map(rec).filter(row => str(row.id)).map(row => ({
    id: str(row.id), subject: str(row.subject), from: str(row.from), to: str(row.toRecipients), cc: str(row.ccRecipients),
    received: new Date(str(row.receivedDateTime)), preview: str(row.bodyPreview), conversationId: str(row.conversationId), importance: str(row.importance), isRead: row.isRead === true, files: filesFromHtml(str(row.body)), hasAttachments: row.hasAttachments === true,
  }));
}
const looked = new Set<string>();
/** Looks up first names for addresses in the directory (cached for the session, max 40 new per call, 6 at a time). */
export async function learnNames(values: string[], limit = 40): Promise<void> {
  const todo = [...new Set(values.flatMap(value => value.split(/[;,]/)).map(value => value.trim().toLowerCase()).filter(value => /^[^\s@]+@[^\s@]+$/.test(value) && !looked.has(value)))].slice(0, limit);
  for (let index = 0; index < todo.length; index += 6) {
    await Promise.all(todo.slice(index, index + 6).map(async address => {
      looked.add(address);
      try {
        const result = await Office365UsersService.UserProfile_V2(address, "givenName,displayName");
        if (result.success) rememberFirstName(address, firstName(str(rec(result.data).displayName), str(rec(result.data).givenName)));
      } catch { /* outside the directory: the alias is shown */ }
    }));
  }
}
export async function loadMail(): Promise<{ inbox: MailRow[]; sent: MailRow[] }> {
  const [inbox, sent] = await Promise.all([
    Office365OutlookService.GetEmailsV3("Inbox", undefined, undefined, undefined, undefined, undefined, false, undefined, false, false, undefined, false, undefined, 50),
    Office365OutlookService.GetEmailsV3("SentItems", undefined, undefined, undefined, undefined, undefined, false, undefined, false, false, undefined, false, undefined, 50),
  ]);
  const rows = { inbox: mailRows(ok(inbox, "Inbox")), sent: sent.success ? mailRows(sent.data) : [] };
  await learnNames([...rows.inbox.slice(0, 30).map(row => row.from), ...rows.sent.slice(0, 20).map(row => row.to.split(/[;,]/)[0] ?? "")]).catch(() => undefined);
  return rows;
}
export const mailLink = (id: string) => `https://outlook.office.com/mail/deeplink/read/${encodeURIComponent(id)}`;
const sameAddress = (list: string, me: string) => !!me && list.toLowerCase().split(/[;,]/).some(part => part.trim() === me.toLowerCase());

/** Emails that ask you for something and that you haven't answered in the same conversation. */
export function repliesFromMail(mail: { inbox: MailRow[]; sent: MailRow[] }, myEmail: string, now = new Date()): WorkItem[] {
  const answered = new Map<string, number>();
  for (const sent of mail.sent) answered.set(sent.conversationId, Math.max(answered.get(sent.conversationId) ?? 0, sent.received.getTime()));
  return mail.inbox.filter(row => row.received.getTime() > now.getTime() - 14 * DAY && !sameAddress(row.from, myEmail) && !AUTOMATED.test(row.from)
      && !NOT_A_REQUEST.test(row.subject.trim()) && !/unsubscribe|view (this|in) (browser|web)/i.test(row.preview)
      && (!myEmail || !row.to || sameAddress(row.to, myEmail)) && REQUEST.test(`${row.subject} ${row.preview}`)
      && (answered.get(row.conversationId) ?? 0) < row.received.getTime())
    .slice(0, 10).map(row => item({
      kind: "reply", title: hideEmails(row.subject) || "No subject", email: row.from,
      summary: `${firstName(row.from) ? `${firstName(row.from)}, ` : ""}${shortDay(row.received)}: ${clip(hideEmails(row.preview), 200)}${row.files?.length ? ` Files: ${fileLinks(row.files)}.` : row.hasAttachments ? " Has attachments." : ""}`,
      next: "Reply in Outlook", url: mailLink(row.id), evidence: `Email · Inbox, ${shortDay(row.received)}`, severity: row.importance === "high" ? "High importance" : "",
    }));
}
/** Emails you sent asking for something, with no answer yet (2–14 days old). */
export function waitingOnOthers(mail: { inbox: MailRow[]; sent: MailRow[] }, myEmail: string, now = new Date()): WorkItem[] {
  const replied = new Map<string, number>();
  for (const row of mail.inbox) if (!sameAddress(row.from, myEmail)) replied.set(row.conversationId, Math.max(replied.get(row.conversationId) ?? 0, row.received.getTime()));
  const seen = new Set<string>();
  return mail.sent.filter(row => { const age = now.getTime() - row.received.getTime(); return age > 2 * DAY && age < 14 * DAY && REQUEST.test(`${row.subject} ${row.preview}`) && !NOT_A_REQUEST.test(row.subject.trim()) && (replied.get(row.conversationId) ?? 0) < row.received.getTime(); })
    .filter(row => { if (seen.has(row.conversationId)) return false; seen.add(row.conversationId); return true; })
    .slice(0, 5).map(row => {
      const to = row.to.split(/[;,]/)[0]?.trim() ?? "";
      return item({ kind: "person", title: firstName(to) || "Someone", email: to, summary: `You asked about "${hideEmails(row.subject)}" on ${shortDay(row.received)} and haven't had a reply.`, next: "Send a short follow-up", url: mailLink(row.id), evidence: `Email · Sent Items, ${shortDay(row.received)}` });
    });
}

// ---------- Calendar ----------
export async function loadUpcomingEvents(calendarId: string, now = new Date()): Promise<CalendarEvent[]> {
  const start = new Date(now); start.setHours(0, 0, 0, 0);
  const end = new Date(start.getTime() + 15 * DAY);
  const data = ok(await Office365OutlookService.GetEventsCalendarViewV3(calendarId, start.toISOString(), end.toISOString(), undefined, "start/dateTime", 100, 0), "Calendar");
  const events = (data.value ?? []) as CalendarEvent[];
  await learnNames(events.slice(0, 40).map(event => event.organizer ?? "")).catch(() => undefined);
  return events;
}
const startOf = (event: CalendarEvent) => { const value = event.startWithTimeZone ?? event.start; const date = value ? new Date(value) : undefined; return date && !Number.isNaN(date.getTime()) ? date : undefined; };

/** Prep notes for the next meetings (free): agenda, who's coming, related team posts and emails. */
export async function meetingsFromCalendar(events: CalendarEvent[], feeds: TeamFeed[], mail: { inbox: MailRow[] } | undefined, now = new Date()): Promise<WorkItem[]> {
  const out: WorkItem[] = [];
  const soon = events.filter(event => { const start = startOf(event); return start && start > now && start.getTime() - now.getTime() < 2 * DAY && hasAttendees(event) && !isPrivateEvent(event) && event.responseType !== "declined" && !event.isAllDay; }).slice(0, 6);
  for (const event of soon) {
    const subject = event.subject ?? "";
    const words = subject.toLowerCase().split(/[^a-z0-9]+/).filter(word => word.length >= 4 && !/^(meeting|sync|weekly|review|call|with|update|check|catch|daily)$/.test(word));
    const team = feeds.find(feed => words.some(word => feed.name.toLowerCase().includes(word)));
    const related = (mail?.inbox ?? []).filter(row => words.length && words.filter(word => row.subject.toLowerCase().includes(word)).length >= Math.min(2, words.length)).slice(0, 2);
    const attendees = [event.requiredAttendees, event.optionalAttendees].join(";").split(";").filter(value => value.trim()).length;
    const agenda = clip(hideEmails(plain(event.body ?? "")), 260);
    const parts = [agenda ? `Agenda: ${agenda}` : "No agenda in the invite.", team?.posts[0] ? `Latest in ${team.name}: ${team.posts[0].author}, ${shortDay(team.posts[0].created)} — ${clip(team.posts[0].text, 160)} [${team.name} › ${team.posts[0].channel}](${team.posts[0].url})` : "",
      ...(filesFromHtml(event.body ?? "").length ? [`Files in the invite: ${fileLinks(filesFromHtml(event.body ?? ""))}.`] : []),
      ...related.map(row => `Related email: [${hideEmails(row.subject)}](${mailLink(row.id)}) from ${firstName(row.from) || "someone"}, ${shortDay(row.received)}`)].filter(Boolean);
    out.push(item({ kind: "prep", title: subject, summary: parts.join(" "), next: `${attendees} attendee${attendees === 1 ? "" : "s"}${event.organizer ? ` · organized by ${firstName(event.organizer) || "someone"}` : ""}${event.location ? ` · ${hideEmails(event.location)}` : ""}`, date: startOf(event)!.toISOString(), url: event.webLink ?? "", evidence: "Calendar invite" }));
  }
  for (const event of events.filter(event => { const start = startOf(event); return start && start > now && LEARNING.test(event.subject ?? "") && event.responseType !== "declined"; }).slice(0, 8)) {
    out.push(item({ kind: "training", title: eventTitle(event), date: startOf(event)!.toISOString(), url: event.webLink ?? "", summary: event.location ? hideEmails(event.location) : "", evidence: "Calendar" }));
  }
  return out;
}

/** Who's away: automatic replies of people in your meetings over the next week (Outlook mail tips). */
export async function awayFromMailTips(events: CalendarEvent[], myEmail: string, now = new Date()): Promise<WorkItem[]> {
  const week = events.filter(event => { const start = startOf(event); return start && start > now && start.getTime() - now.getTime() < 7 * DAY && hasAttendees(event) && event.responseType !== "declined"; });
  const addresses = [...new Set(week.flatMap(event => [event.organizer ?? "", ...[event.requiredAttendees, event.optionalAttendees].join(";").split(";")]).map(value => value.trim().toLowerCase()).filter(value => /@/.test(value) && value !== myEmail.toLowerCase()))].slice(0, 20);
  if (!addresses.length) return [];
  const result = await Office365OutlookService.GetMailTips_V2({ EmailAddresses: addresses, MailTipsOptions: "automaticReplies" });
  const tips = ok(result, "Mail tips").value ?? [];
  const out: WorkItem[] = [];
  const awayAddresses = tips.map((tip, index) => str(rec(rec(tip).emailAddress).address).toLowerCase() || addresses[index]).filter((_, index) => plain(str(rec(rec(tips[index]).automaticReplies).message)));
  await learnNames(awayAddresses).catch(() => undefined);
  tips.forEach((tip, index) => {
    const raw = rec(tip), auto = rec(raw.automaticReplies), message = plain(str(auto.message));
    if (!message) return;
    const address = str(rec(raw.emailAddress).address).toLowerCase() || addresses[index];
    const from = rec(auto.scheduledStartTime), to = rec(auto.scheduledEndTime);
    const start = str(from.dateTime) ? new Date(`${str(from.dateTime).replace(/Z?$/, "Z")}`) : now;
    const end = str(to.dateTime) ? new Date(`${str(to.dateTime).replace(/Z?$/, "Z")}`) : undefined;
    const organized = week.find(event => (event.organizer ?? "").toLowerCase() === address);
    out.push(item({ kind: "oof", title: firstName(address) || "Someone", email: address, date: isoDay(start), endDate: end ? isoDay(end) : "", relatedMeeting: organized ? eventTitle(organized) : "", summary: clip(hideEmails(message), 200), evidence: "Their Outlook automatic reply is on" }));
  });
  return out;
}

// ---------- Files ----------
/** Documents trending around you (Office 365 Users). Optional: any failure returns nothing. */
export async function loadTrendingFiles(): Promise<FileHit[]> {
  try {
    const trending = await Office365UsersService.MyTrendingDocuments();
    if (!trending.success) return [];
    return (trending.data?.value ?? []).slice(0, 10).flatMap((doc): FileHit[] => {
      const visual = doc.resourceVisualization ?? {}, reference = rec(rec(doc).resourceReference);
      const url = str(reference.webUrl);
      if (!visual.title || !/^https:\/\//i.test(url)) return [];
      return [{ file: { name: hideEmails(visual.title), url }, origin: "trending", context: visual.containerDisplayName ? hideEmails(visual.containerDisplayName) : "" }];
    });
  } catch { return []; }
}

/** Every file the dashboard has already seen, tagged with why it matters (no extra reads). */
export function collectFileHits(input: { mail?: { inbox: MailRow[] }; replyUrls: Set<string>; feeds: TeamFeed[]; waiting: { person: string; chatTitle: string; messageUrl: string; created: string; files?: FileRef[] }[]; events: CalendarEvent[]; trending: FileHit[]; now?: Date }): FileHit[] {
  const now = input.now ?? new Date();
  const hits: FileHit[] = [];
  for (const row of input.mail?.inbox ?? []) {
    if (!row.files?.length || row.received.getTime() < now.getTime() - 14 * DAY) continue;
    const reply = input.replyUrls.has(mailLink(row.id));
    for (const file of row.files) hits.push({ file, origin: reply ? "reply" : "email", who: firstName(row.from) || undefined, when: row.received, context: `Email: ${hideEmails(row.subject) || "No subject"}`, contextUrl: mailLink(row.id) });
  }
  for (const chat of input.waiting) for (const file of chat.files ?? []) hits.push({ file, origin: "chat", who: chat.person, when: new Date(chat.created), context: chat.chatTitle, contextUrl: chat.messageUrl });
  for (const feed of input.feeds) for (const post of feed.posts) for (const file of post.files ?? []) hits.push({ file, origin: post.mentionsMe ? "mention" : "team", who: post.author, when: post.created, context: `${feed.name} › ${post.channel}`, contextUrl: post.url });
  for (const event of input.events) {
    const start = startOf(event);
    if (!start || start < now || start.getTime() - now.getTime() > 2 * DAY || isPrivateEvent(event) || event.responseType === "declined") continue;
    for (const file of filesFromHtml(event.body ?? "")) hits.push({ file, origin: "meeting", when: start, meetingAt: start, context: eventTitle(event), contextUrl: event.webLink ?? undefined });
  }
  return [...hits, ...input.trending];
}
