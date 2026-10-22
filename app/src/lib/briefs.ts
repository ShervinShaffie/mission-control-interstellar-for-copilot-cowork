/**
 * Free documents assembled from what the dashboard already read (no Copilot). Sources are markdown links,
 * so they show as short clickable names in the app, in Word and in PDF.
 */
import type { CalendarEvent } from "@/lib/command-center-evidence";
import { eventTitle } from "@/lib/command-center-evidence";
import type { WorkItem } from "@/lib/command-center-data";
import type { DocContent } from "@/lib/documents";
import { plain, type TeamFeed } from "@/lib/direct-signals";
import { firstName, hideEmails } from "@/lib/privacy";
import type { OpenTask } from "@/lib/todo";
import type { WaitingChat } from "@/lib/teams-waiting";

const when = (date: Date) => date.toLocaleString([], { weekday: "long", month: "long", day: "numeric", hour: "numeric", minute: "2-digit" });
const day = (date: Date) => date.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" });
const link = (label: string, url?: string) => (url ? `[${label.replace(/[[\]]/g, "")}](${url})` : label);
const clip = (text: string, max = 260) => (text.length > max ? `${text.slice(0, max - 1).trimEnd()}…` : text);
const names = (list: string) => list.split(";").map(value => value.trim()).filter(Boolean).map(value => firstName(value) || value.split("@")[0]).join(", ");
const withNotes = (doc: DocContent, notes: string): DocContent => notes.trim() ? { ...doc, sections: [...doc.sections, { heading: "My notes", text: notes.trim() }] } : doc;

export function meetingBrief(event: CalendarEvent, prep: WorkItem | undefined, feed: TeamFeed | undefined, tasks: OpenTask[], notes = ""): DocContent {
  const start = new Date(event.startWithTimeZone ?? event.start ?? Date.now());
  const sections = [
    { heading: "Meeting", text: [`${when(start)}${event.location ? ` · ${hideEmails(event.location)}` : ""}`, event.organizer ? `Organizer: ${firstName(event.organizer) || "—"}` : "", `Attendees: ${names([event.requiredAttendees, event.optionalAttendees].join(";")) || "—"}`, link("Open the invite", event.webLink)].filter(Boolean).join("\n") },
    { heading: "Agenda from the invite", text: clip(hideEmails(plain(event.body ?? "")), 1500) || "The invite has no agenda. Consider sending one." },
  ];
  if (prep?.summary) sections.push({ heading: "Related emails and posts", text: prep.summary });
  if (feed) {
    sections.push({ heading: `Latest in ${feed.name}`, text: feed.posts.slice(0, 5).map(post => `- ${post.author}, ${day(post.created)}: ${clip(post.text, 220)} ${link(`${feed.name} › ${post.channel}`, post.url)}`).join("\n") || "No recent posts." });
  }
  if (tasks.length) sections.push({ heading: "Your open tasks that mention it", text: tasks.map(task => `- ${task.title}${task.due ? ` (due ${task.due})` : ""}`).join("\n") });
  sections.push({ heading: "Before the meeting", text: "- Decide the outcome you need from this meeting\n- Check the open items above and who owns each\n- Note the questions you want answered" });
  return withNotes({ title: `Prep brief: ${eventTitle(event)}`, sections }, notes);
}

export function projectBrief(project: WorkItem, feed: TeamFeed | undefined, deadlines: WorkItem[], meetings: { event: CalendarEvent; start: Date }[], tasks: OpenTask[], notes = ""): DocContent {
  const sections = [
    { heading: "Status", text: [project.status ? `Status mentioned in Teams: ${project.status}` : "No status word (red, amber, green, at risk…) in recent posts.", project.summary].filter(Boolean).join("\n") },
    { heading: "Recent activity", text: feed?.posts.slice(0, 8).map(post => `- ${post.author}, ${day(post.created)} in ${post.channel}: ${clip(post.text, 220)} ${link(`${feed.name} › ${post.channel}`, post.url)}`).join("\n") || "No posts in the last three weeks." },
    { heading: "Deadlines mentioned", text: deadlines.map(item => `- ${item.date}: ${item.title} ${link(item.evidence || "Source", item.url)}`).join("\n") || "None found in recent posts." },
    { heading: "Upcoming meetings", text: meetings.map(meeting => `- ${when(meeting.start)}: ${link(eventTitle(meeting.event), meeting.event.webLink)}`).join("\n") || "None in the next 14 days." },
  ];
  if (tasks.length) sections.push({ heading: "Your To Do tasks", text: tasks.map(task => `- ${task.title}${task.due ? ` (due ${task.due})` : ""}`).join("\n") });
  if (feed?.channels.some(channel => channel.filesUrl)) sections.push({ heading: "Files", text: feed.channels.filter(channel => channel.filesUrl).map(channel => `- ${link(`${feed.name} › ${channel.name} files`, channel.filesUrl)}`).join("\n") });
  return withNotes({ title: `Project briefing: ${project.title}`, sections }, notes);
}

export function statusReport(projects: WorkItem[], feeds: TeamFeed[], deadlines: WorkItem[], weekOf: Date, notes = ""): DocContent {
  return withNotes({ title: `Weekly status — week of ${weekOf.toLocaleDateString([], { month: "long", day: "numeric" })}`, sections: projects.map(project => {
    const feed = feeds.find(item => item.name === project.title);
    const due = deadlines.filter(item => item.evidence.includes(project.title)).slice(0, 4);
    return { heading: project.title, text: [project.status ? `Status: ${project.status}` : "", `Latest: ${project.summary}`, due.length ? `Coming up:\n${due.map(item => `- ${item.date}: ${item.title}`).join("\n")}` : "", feed?.posts[0] ? link("Open the latest post", feed.posts[0].url) : ""].filter(Boolean).join("\n") };
  }) }, notes);
}

export function teamsCatchUp(waiting: WaitingChat[], feeds: TeamFeed[], now = new Date(), notes = ""): DocContent {
  const since = now.getTime() - 3 * 864e5;
  const sections = [{ heading: "Waiting on you in chats", text: waiting.map(chat => `- ${chat.person}${chat.group ? ` (${chat.chatTitle})` : ""}: "${clip(chat.preview, 160)}" ${link("Open the chat", chat.messageUrl)}`).join("\n") || "No one is waiting on you." }];
  for (const feed of feeds) {
    const recent = feed.posts.filter(post => post.created.getTime() > since);
    if (recent.length) sections.push({ heading: `${feed.name} (${recent.length} new)`, text: recent.slice(0, 6).map(post => `- ${post.mentionsMe ? "@you · " : ""}${post.author} in ${post.channel}: ${clip(post.text, 200)} ${link("Open", post.url)}`).join("\n") });
  }
  return withNotes({ title: `Teams catch-up — last 3 days`, sections }, notes);
}
