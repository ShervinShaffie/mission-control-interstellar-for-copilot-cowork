import { useId, useState } from "react";
import type { UseQueryResult } from "@tanstack/react-query";
import { AlertTriangle, BriefcaseBusiness, CalendarClock, ChevronDown, ChevronsDownUp, ChevronsUpDown, ChevronUp, ExternalLink, FolderOpen, Paperclip, ListChecks, NotebookPen, Users } from "lucide-react";
import { Button } from "@/components/ui/button";
import { ActionRow, DismissedList } from "@/components/action-list";
import { useDismissed } from "@/lib/dismissed";
import { EmptyState, ErrorState, LoadingState } from "@/components/states";
import { RichText } from "@/components/rich-text";
import { eventTitle } from "@/lib/command-center-evidence";
import type { WorkItem } from "@/lib/command-center-data";
import type { JoinedTeam, ProjectView } from "@/lib/projects";
import { safeUrl } from "@/lib/safe-url";
import { cn } from "@/lib/utils";

const HEALTH_LABEL = { risk: "Needs attention", watch: "Watch", ok: "On track", unknown: "Status not stated" } as const;
const when = (date: Date) => `${date.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" })}, ${date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}`;
const dueText = (value: string) => value ? new Date(/^\d{4}-\d{2}-\d{2}$/.test(value) ? `${value}T12:00:00` : value).toLocaleDateString([], { month: "short", day: "numeric" }) : "";

/** Which project cards are collapsed, remembered per person in this browser (titles only, no business data beyond the team name). */
function collapsedKey(owner?: string) { return owner ? `mission-control-interstellar:collapsed-projects:v1:${owner}` : undefined; }
function readCollapsed(key?: string): string[] {
  if (!key) return [];
  try { const value: unknown = JSON.parse(localStorage.getItem(key) ?? "[]"); return Array.isArray(value) ? value.filter((item): item is string => typeof item === "string").slice(0, 100) : []; } catch { return []; }
}
function useCollapsedProjects(owner?: string) {
  const key = collapsedKey(owner);
  const [store, setStore] = useState(() => ({ key, list: readCollapsed(key) }));
  if (store.key !== key) setStore({ key, list: readCollapsed(key) });
  const save = (list: string[]) => {
    setStore({ key, list });
    if (key) try { localStorage.setItem(key, JSON.stringify(list.slice(0, 100))); } catch { /* keep it for this page only */ }
  };
  return { collapsed: store.list, save };
}

function ProjectCard({ view, mySite, collapsed, onToggle }: { view: ProjectView; mySite?: string; collapsed: boolean; onToggle: () => void }) {
  const bodyId = useId();
  const { isDismissed } = useDismissed();
  const actions = view.actions.filter(action => !isDismissed(action.id));
  const { project, health } = view;
  const link = safeUrl(view.link);
  return <li className={cn("project-card rounded-2xl border bg-card p-4", `project-${health}`)}>
    <div className="flex flex-wrap items-start justify-between gap-3">
      <div className="min-w-0 flex-1">
        <h4 className="flex flex-wrap items-center gap-2 font-heading text-lg font-bold">
          {health === "risk" && <AlertTriangle aria-hidden="true" className="size-4 text-red-700 dark:text-red-400" />}{project.title}
        </h4>
        <div className="mt-1.5 flex flex-wrap gap-1.5">
          <span className={cn("health-pill rounded-full px-2.5 py-0.5 text-xs font-semibold", `health-${health}`)}>{HEALTH_LABEL[health]}{project.status && health !== "unknown" ? ` · ${project.status}` : ""}</span>
          {view.reasons.filter(reason => !reason.startsWith("Status") && !reason.startsWith("Watch")).map(reason => <span key={reason} className="rounded-full border border-border px-2.5 py-0.5 text-xs">{reason}</span>)}
        </div>
      </div>
      <div className="flex shrink-0 flex-wrap items-center gap-2">
      {link && <a href={link} target="_blank" rel="noopener noreferrer" className="inline-flex shrink-0 items-center gap-1.5 rounded-full border border-border px-3 py-1.5 text-sm font-medium hover:bg-accent focus-visible:outline-2 focus-visible:outline-ring">Open the source<ExternalLink aria-hidden="true" className="size-3.5" /><span className="sr-only">(opens in a new tab)</span></a>}
      <Button type="button" variant="ghost" size="sm" className="rounded-full" aria-expanded={!collapsed} aria-controls={bodyId} onClick={onToggle}>
        {collapsed ? <ChevronDown aria-hidden="true" /> : <ChevronUp aria-hidden="true" />}{collapsed ? "Expand" : "Collapse"}<span className="sr-only"> {project.title}</span>
      </Button>
      </div>
    </div>
    <div id={bodyId} hidden={collapsed}>
    {project.summary && <p className="mt-3 break-words text-sm"><RichText text={project.summary} interactive={!link} /></p>}
    {project.next && <p className="mt-1 break-words text-sm"><strong>Next:</strong> <RichText text={project.next} interactive={!link} /></p>}

    <div className="mt-4 grid gap-4 lg:grid-cols-2">
      <section aria-label={`Upcoming meetings for ${project.title}`}>
        <h5 className="flex items-center gap-1.5 text-xs font-bold uppercase tracking-wide text-muted-foreground"><CalendarClock aria-hidden="true" className="size-3.5" />Upcoming meetings & prep</h5>
        {view.meetings.length ? <ul className="mt-2 space-y-2">{view.meetings.map(({ event, start, prep }) => <li key={event.id ?? start.toISOString()} className={cn("rounded-xl border border-border p-3 text-sm", prep && "prep-note")}>
          <p className="font-medium">{eventTitle(event)}</p>
          <p className="text-xs text-muted-foreground">{when(start)}</p>
          {prep ? <div className="mt-2 rounded-lg bg-accent/60 p-2">
            <p className="flex items-center gap-1.5 text-xs font-semibold"><NotebookPen aria-hidden="true" className="size-3.5" />Prep notes</p>
            <p className="mt-1 break-words text-sm"><RichText text={prep.summary} /></p>
            {prep.next && <p className="mt-1 break-words text-sm"><strong>Prepare:</strong> <RichText text={prep.next} /></p>}
          </div> : <p className="mt-1 text-xs text-muted-foreground">No prep notes found for this one.</p>}
        </li>)}</ul> : <p className="mt-2 text-sm text-muted-foreground">No meetings in the next 14 days mention this project.</p>}
      </section>
      <section aria-label={`Deadlines and tasks for ${project.title}`}>
        <h5 className="flex items-center gap-1.5 text-xs font-bold uppercase tracking-wide text-muted-foreground"><ListChecks aria-hidden="true" className="size-3.5" />Deadlines & tasks</h5>
        {view.deadlines.length || view.tasks.length ? <ul className="mt-2 space-y-1.5 text-sm">
          {view.deadlines.map((item: WorkItem) => <li key={`${item.title}-${item.date}`} className="flex gap-2"><span className="shrink-0 font-semibold">{item.date ? dueText(item.date) : "—"}</span><span className="min-w-0 break-words">{item.title}</span></li>)}
          {view.tasks.map(task => <li key={task.id} className="flex gap-2"><span className="shrink-0 font-semibold">{task.due ? dueText(task.due) : "To Do"}</span><span className="min-w-0 break-words">{task.title}</span></li>)}
        </ul> : <p className="mt-2 text-sm text-muted-foreground">No dated items or To Do tasks mention this project.</p>}
      </section>
    </div>

    <section aria-label={`Files for ${project.title}`} className="mt-4">
      <h5 className="flex items-center gap-1.5 text-xs font-bold uppercase tracking-wide text-muted-foreground"><Paperclip aria-hidden="true" className="size-3.5" />Files shared in this team</h5>
      {view.files.length ? <ul className="mt-2 flex flex-wrap gap-2">{view.files.map(file => <li key={file.url}>{safeUrl(file.url)
        ? <a href={safeUrl(file.url)} target="_blank" rel="noopener noreferrer" title={`${file.who}, ${dueText(file.when.toISOString())}`} className="inline-flex max-w-[22rem] items-center gap-1.5 rounded-full border border-border px-3 py-1 text-sm hover:bg-accent focus-visible:outline-2 focus-visible:outline-ring"><span className="truncate">{file.name}</span><span className="sr-only"> (shared by {file.who}, opens in a new tab)</span></a>
        : <span className="text-sm">{file.name}</span>}</li>)}</ul>
        : <p className="mt-2 text-sm text-muted-foreground">No files shared in recent posts.</p>}
      {safeUrl(view.filesFolder) && <a href={safeUrl(view.filesFolder)} target="_blank" rel="noopener noreferrer" className="mt-2 inline-flex items-center gap-1.5 text-sm font-medium text-primary underline underline-offset-2"><FolderOpen aria-hidden="true" className="size-3.5" />All of this team's files<span className="sr-only">(opens in a new tab)</span></a>}
    </section>

    <details className="project-actions mt-4 rounded-xl border border-border p-3" open={health === "risk"}>
      <summary className="cursor-pointer text-sm font-semibold focus-visible:outline-2 focus-visible:outline-ring">Actions for this project ({actions.length})</summary>
      <ol className="mt-3 space-y-3">{actions.map(action => <ActionRow key={action.id} action={action} mySite={mySite} section="projects" />)}</ol>
    </details>
    </div>
  </li>;
}

export function ProjectBoard({ views, query, teams, searchOn, mySite, owner }: { views: ProjectView[]; query: UseQueryResult<unknown>; teams: UseQueryResult<JoinedTeam[]>; searchOn: boolean; mySite?: string; owner?: string }) {
  const { collapsed, save } = useCollapsedProjects(owner);
  const titles = views.map(view => view.project.title);
  const allCollapsed = titles.length > 0 && titles.every(title => collapsed.includes(title));
  const toggle = (title: string) => save(collapsed.includes(title) ? collapsed.filter(item => item !== title) : [...collapsed, title]);
  const needs = views.filter(view => view.health === "risk" || view.health === "watch" || view.attention >= 15).length;
  return <div className="space-y-4">
    <p className="text-sm text-muted-foreground">Your project teams in Microsoft Teams, with their latest status, deadlines mentioned in posts, upcoming meetings and your To Do tasks. {views.length ? <>Most in need of attention first{needs ? <> — <strong className="text-foreground">{needs}</strong> need you now</> : null}.</> : null}</p>
    {!searchOn ? null : query.isPending ? <LoadingState rows={3} />
      : query.isError && !views.length ? <div role="alert"><ErrorState title="Your teams couldn't be read" action={<Button variant="outline" disabled={query.isFetching} onClick={() => void query.refetch()}>Retry</Button>} /></div>
      : views.length ? <><div className="flex justify-end"><Button type="button" variant="outline" size="sm" className="rounded-full" onClick={() => save(allCollapsed ? collapsed.filter(title => !titles.includes(title)) : [...new Set([...collapsed, ...titles])])}>{allCollapsed ? <ChevronsUpDown aria-hidden="true" /> : <ChevronsDownUp aria-hidden="true" />}{allCollapsed ? "Expand all projects" : "Collapse all projects"}</Button></div><ol className="space-y-4">{views.map(view => <ProjectCard key={view.project.title} view={view} mySite={mySite} collapsed={collapsed.includes(view.project.title)} onToggle={() => toggle(view.project.title)} />)}</ol></>
      : <EmptyState icon={BriefcaseBusiness} title="No project activity found" description="Use Choose project teams (top right) to pick which teams are your projects, then Refresh." />}
    <DismissedList section="projects" label="Dismissed project actions" />
    {(!searchOn || !views.length) && <section aria-labelledby="joined-teams-heading" className="rounded-2xl border border-border p-4">
      <h4 id="joined-teams-heading" className="flex items-center gap-2 font-semibold"><Users aria-hidden="true" className="size-4" />Your teams</h4>
      <p className="mt-1 text-xs text-muted-foreground">Read directly from Teams.</p>
      {teams.isPending ? <LoadingState rows={2} /> : teams.isError ? <p className="mt-2 text-sm text-muted-foreground">Your teams couldn't be loaded.</p>
        : teams.data?.length ? <ul className="mt-2 grid gap-2 sm:grid-cols-2">{teams.data.map(team => <li key={team.id} className="rounded-xl border border-border p-3 text-sm"><p className="font-medium">{team.name}</p>{team.description && <p className="mt-1 line-clamp-2 text-xs text-muted-foreground">{team.description}</p>}</li>)}</ul>
        : <p className="mt-2 text-sm text-muted-foreground">You're not a member of any teams.</p>}
    </section>}
  </div>;
}
