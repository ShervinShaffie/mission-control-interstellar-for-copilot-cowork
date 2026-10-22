/**
 * Project board: joins each project team with the meetings, prep notes, deadlines and To Do tasks
 * that mention it, all from data already loaded (free). Actions only run when the user selects Run.
 */
import { MicrosoftTeamsService } from "../../generated/services/MicrosoftTeamsService";
import type { SuggestedAction } from "@/lib/actions";
import { addToDo } from "@/lib/action-runners";
import { meetingBrief, projectBrief } from "@/lib/briefs";
import type { TeamFeed } from "@/lib/direct-signals";
import { eventTitle, isPrivateEvent, type CalendarEvent } from "@/lib/command-center-evidence";
import type { WorkItem } from "@/lib/command-center-data";
import { hideEmails } from "@/lib/privacy";
import { firstLinkUrl } from "@/lib/source-links";
import { startOf } from "@/lib/suggestions";
import { taskState, type OpenTask } from "@/lib/todo";

const STOP = new Set("the and for with from into that this your our project program programme workstream team update review meeting sync weekly daily status plan phase rollout initiative".split(" "));
export const RISK = /\b(red|at risk|amber|blocked|blocker|behind|delayed|slipp\w*|escalat\w*|overdue|missed)\b/i;
export type Health = "risk" | "watch" | "ok" | "unknown";

export interface ProjectMeeting { event: CalendarEvent; start: Date; prep?: WorkItem }
export interface ProjectView {
  project: WorkItem; health: Health; attention: number; reasons: string[]; link?: string;
  meetings: ProjectMeeting[]; deadlines: WorkItem[]; tasks: OpenTask[]; actions: SuggestedAction[];
  /** Files shared in this team's recent posts (newest first) and the team's files folder. */
  files: { name: string; url: string; who: string; when: Date }[]; filesFolder?: string;
}

/** Words and codes that identify a project in other titles (for example "OPS-SECURE" or "Contoso"). */
export function projectTerms(title: string): { codes: string[]; words: string[] } {
  const codes = [...title.matchAll(/\b[A-Z][A-Z0-9]+(?:-[A-Z0-9]+)+\b|\b[A-Z]{3,}\b/g)].map(match => match[0].toLowerCase());
  const words = title.toLowerCase().replace(/\([^)]*\)/g, " ").split(/[^a-z0-9]+/).filter(word => word.length >= 4 && !STOP.has(word));
  return { codes: [...new Set(codes)], words: [...new Set(words)] };
}
export function mentions(text: string, terms: { codes: string[]; words: string[] }) {
  const value = text.toLowerCase();
  if (terms.codes.some(code => value.includes(code))) return true;
  const hits = terms.words.filter(word => value.includes(word)).length;
  return terms.words.length <= 1 ? hits === 1 && terms.words.length === 1 : hits >= 2;
}
export function healthOf(item: WorkItem): Health {
  const text = `${item.status} ${item.severity}`;
  if (/\bred\b|at risk|blocked|escalat|overdue|missed|behind/i.test(text)) return "risk";
  if (/amber|yellow|watch|delayed|slipp/i.test(text) || RISK.test(item.summary)) return "watch";
  if (/green|on track|complete|done/i.test(text)) return "ok";
  return "unknown";
}
const shortDay = (date: Date) => date.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" });
const clock = (date: Date) => date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });

export function buildProjects(input: { now: Date; projects: WorkItem[]; work: WorkItem[]; meetings: WorkItem[]; events: CalendarEvent[]; tasks: OpenTask[]; feeds?: TeamFeed[] }): ProjectView[] {
  const { now } = input, horizon = now.getTime() + 14 * 864e5;
  const upcoming = input.events.filter(event => { const start = startOf(event); return start && start > now && start.getTime() < horizon && !event.isAllDay && event.responseType !== "declined" && !isPrivateEvent(event); });
  const preps = input.meetings.filter(item => item.kind === "prep");
  return input.projects.map(project => {
    const terms = projectTerms(project.title);
    const meetings = upcoming.filter(event => mentions(event.subject ?? "", terms) || preps.some(prep => prep.title.trim().toLowerCase() === (event.subject ?? "").trim().toLowerCase() && mentions(`${prep.summary} ${prep.evidence}`, terms)))
      .map(event => ({ event, start: startOf(event)!, prep: preps.find(prep => prep.title.trim().toLowerCase() === (event.subject ?? "").trim().toLowerCase()) }))
      .sort((a, b) => a.start.getTime() - b.start.getTime()).slice(0, 4);
    const deadlines = input.work.filter(item => (item.kind === "action" || item.kind === "urgent") && (item.evidence.startsWith(`Teams › ${project.title} ›`) || mentions(`${item.title} ${item.summary} ${item.evidence}`, terms))).slice(0, 5);
    const tasks = input.tasks.filter(task => mentions(task.title, terms)).slice(0, 5);
    const health = healthOf(project);
    const overdue = deadlines.filter(item => item.date && new Date(/^\d{4}-\d{2}-\d{2}$/.test(item.date) ? `${item.date}T23:59:59` : item.date) < now).length + tasks.filter(task => taskState(task, now) === "overdue").length;
    const soon = meetings.filter(meeting => meeting.start.getTime() - now.getTime() < 2 * 864e5);
    const reasons = [
      health === "risk" ? `Status: ${project.status || project.severity || "at risk"}` : health === "watch" ? `Watch: ${project.status || "risk mentioned"}` : "",
      overdue ? `${overdue} overdue item${overdue > 1 ? "s" : ""}` : "",
      soon.length ? `Meeting ${sameDayLabel(soon[0].start, now)}` : "",
      meetings.some(meeting => meeting.prep) ? "Prep notes ready" : "",
    ].filter(Boolean);
    const attention = (health === "risk" ? 40 : health === "watch" ? 20 : 0) + overdue * 15 + soon.length * 12 + deadlines.length * 3 + tasks.length * 2;
    const link = project.url || firstLinkUrl(project.evidence) || undefined;
    const feed = input.feeds?.find(item => item.name === project.title);
    const seen = new Set<string>();
    const files = (feed?.posts ?? []).flatMap(post => (post.files ?? []).map(file => ({ ...file, who: post.author, when: post.created })))
      .filter(file => { const key = file.name.toLowerCase(); if (seen.has(key)) return false; seen.add(key); return true; }).slice(0, 5);
    const filesFolder = feed?.channels.find(channel => channel.filesUrl)?.filesUrl || undefined;
    return { project, health, attention, reasons, link, meetings, deadlines, tasks, files, filesFolder, actions: projectActions(project, meetings, link, deadlines, tasks, feed) };
  }).sort((a, b) => b.attention - a.attention);
}
function sameDayLabel(date: Date, now: Date) {
  if (date.toDateString() === now.toDateString()) return `today at ${clock(date)}`;
  const tomorrow = new Date(now); tomorrow.setDate(now.getDate() + 1);
  return date.toDateString() === tomorrow.toDateString() ? `tomorrow at ${clock(date)}` : `${shortDay(date)} at ${clock(date)}`;
}

function projectActions(project: WorkItem, meetings: ProjectMeeting[], link: string | undefined, deadlines: WorkItem[], tasks: OpenTask[], feed?: TeamFeed): SuggestedAction[] {
  const name = project.title, out: SuggestedAction[] = [];
  const next = meetings[0];
  if (next) out.push({ id: `proj-prep-${name}-${next.event.id}`, title: `Prep brief for ${eventTitle(next.event)}`, detail: `${shortDay(next.start)}, ${clock(next.start)}. A one-page brief from the invite, this team's recent posts and your tasks.`, priority: 80, cost: "free", sources: ["calendar", "teams"],
    runLabel: "Build the brief", notesLabel: "Your own notes to include (optional)", copilotBrief: "prep",
    run: async ({ notes }) => ({ message: "Here's your prep brief.", links: [], doc: meetingBrief(next.event, next.prep, feed, tasks, notes) }) });
  if (project.next) out.push({ id: `proj-todo-${name}`, title: `Add the next deadline to To Do`, detail: project.next, priority: 70, cost: "free", sources: ["todo"], runLabel: "Add to To Do", notesLabel: "Task notes (optional)", notesDefault: `Project: ${name}`,
    run: ({ notes }) => addToDo(`${name}: ${project.next}`.slice(0, 250), undefined, notes, link) });
  out.push({ id: `proj-brief-${name}`, title: `Project briefing`, detail: "Status, recent activity, deadlines mentioned, upcoming meetings, your tasks and the team's files — as a document you can download.", priority: 60, cost: "free", sources: ["teams"],
    runLabel: "Build the briefing", notesLabel: "Your own notes to include (optional)", copilotBrief: "project",
    run: async ({ notes }) => ({ message: "Here's your project briefing.", links: [], doc: projectBrief(project, feed, deadlines, meetings, tasks, notes) }) });
  if (link) out.push({ id: `proj-open-${name}`, title: `Open ${name} in Teams`, detail: "Read the latest posts and reply there.", priority: 40, cost: "free", sources: ["teams"], open: { label: "Open in Teams", url: link } });
  return out;
}

/** Joined Teams teams (free): shown when no team activity was found, so the board is never empty. */
export interface JoinedTeam { id: string; name: string; description: string }
export async function loadJoinedTeams(): Promise<JoinedTeam[]> {
  const result = await MicrosoftTeamsService.GetAllTeams();
  if (!result.success) throw result.error ?? new Error("Teams unavailable");
  const value = (result.data as { value?: unknown })?.value;
  if (!Array.isArray(value)) return [];
  return value.flatMap(row => {
    if (!row || typeof row !== "object") return [];
    const team = row as Record<string, unknown>;
    return typeof team.id === "string" && typeof team.displayName === "string" ? [{ id: team.id, name: hideEmails(team.displayName), description: typeof team.description === "string" ? hideEmails(team.description) : "" }] : [];
  }).slice(0, 20);
}
