import { useMemo, useState, type ReactNode } from "react";
import { useQuery } from "@tanstack/react-query";
import { CheckSquare, ExternalLink, FileText, RefreshCw, Users } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { ConnectorErrorState, EmptyState } from "@/components/states";
import { useProfile } from "@/hooks/use-work-signals";
import { hideEmails } from "@/lib/privacy";
import { safeUrl } from "@/lib/safe-url";
import { cn } from "@/lib/utils";
import { InfoTip } from "@/components/info-tip";
import { PRIORITY, loadPlanner, plannerLinks, plannerTaskState, priorityOf, type PlannerPlan, type PlannerTask } from "@/lib/planner";

type View = "open" | "overdue" | "week" | "done";
const VIEWS: [View, string][] = [["open", "Open"], ["overdue", "Overdue"], ["week", "Due this week"], ["done", "Completed"]];
const WINDOW = 12;
const STATE_TEXT = { overdue: "Overdue", today: "Due today", week: "Due this week", later: "Due later", none: "No due date", done: "Completed" } as const;

function dueText(task: PlannerTask) {
  if (!task.due) return undefined;
  const [y, m, d] = task.due.split("-").map(Number);
  const date = new Date(y, m - 1, d);
  return date.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric", ...(y === new Date().getFullYear() ? {} : { year: "numeric" }) });
}
function progressText(task: PlannerTask) {
  return task.percentComplete >= 100 ? "Completed" : task.percentComplete > 0 ? "In progress" : "Not started";
}

/** One link out to Planner, opening in a new tab. */
function Out({ href, children, className }: { href: string; children: ReactNode; className?: string }) {
  const url = safeUrl(href);
  if (!url) return <span className={className}>{children}</span>;
  return <a href={url} target="_blank" rel="noopener noreferrer" className={cn("rounded-sm font-semibold text-primary underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-ring", className)}>
    {children}<span className="sr-only"> (opens in a new tab)</span>
  </a>;
}

function TaskRow({ task, plan, link }: { task: PlannerTask; plan?: PlannerPlan; link: string }) {
  const state = plannerTaskState(task);
  const priority = priorityOf(task.priority);
  const labels = task.categories.map(key => plan?.labels[key] ?? `Label ${key.replace("category", "")}`);
  return <li className="py-3">
    <div className="flex flex-wrap items-center gap-2">
      <Out href={link} className="break-words">{hideEmails(task.title)}</Out>
      <Badge variant="outline" className={cn("h-auto whitespace-normal", state === "overdue" && "border-red-700 bg-red-700 text-white", state === "today" && "border-amber-800 bg-amber-800 text-white", state === "done" && "border-green-700 bg-green-700 text-white")}>
        {STATE_TEXT[state]}{task.due && state !== "done" ? ` · ${dueText(task)}` : ""}
      </Badge>
      {priority && priority !== "medium" && <Badge variant="outline">Priority: {PRIORITY[priority]}</Badge>}
    </div>
    <p className="mt-1 flex flex-wrap gap-x-4 gap-y-1 text-xs text-muted-foreground">
      <span>{progressText(task)}</span>
      {task.checklistTotal ? <span className="inline-flex items-center gap-1"><CheckSquare aria-hidden="true" className="size-3.5" />Checklist {task.checklistTotal - (task.checklistOpen ?? 0)} of {task.checklistTotal} done</span> : null}
      {task.assigneeCount > 1 && <span className="inline-flex items-center gap-1"><Users aria-hidden="true" className="size-3.5" />Shared with {task.assigneeCount - 1} other{task.assigneeCount > 2 ? "s" : ""}</span>}
      {task.hasDescription && <span className="inline-flex items-center gap-1"><FileText aria-hidden="true" className="size-3.5" />Has notes</span>}
      {task.start && state !== "done" && <span>Starts {dueText({ ...task, due: task.start })}</span>}
      {task.createdBy && <span>Created by {hideEmails(task.createdBy)}</span>}
    </p>
    {labels.length > 0 && <div className="mt-1 flex flex-wrap gap-1" aria-label="Labels">{labels.map(name => <Badge key={name} variant="outline" className="h-auto whitespace-normal">{name}</Badge>)}</div>}
  </li>;
}

export function PlannerSection() {
  const profile = useProfile();
  const links = plannerLinks(profile.data?.userPrincipalName ?? profile.data?.mail ?? undefined);
  const planner = useQuery({ queryKey: ["planner"], queryFn: loadPlanner });
  const [view, setView] = useState<View>("open");
  const [shown, setShown] = useState(WINDOW);

  const counts = useMemo(() => {
    const tasks = planner.data?.tasks ?? [];
    const states = tasks.map(task => plannerTaskState(task));
    return {
      open: states.filter(s => s !== "done").length, overdue: states.filter(s => s === "overdue").length,
      week: states.filter(s => s === "today" || s === "week").length, done: states.filter(s => s === "done").length,
    };
  }, [planner.data]);

  if (planner.isPending) return <p role="status" className="py-6 text-sm text-muted-foreground">Loading your Planner tasks…</p>;
  if (planner.isError) return <ConnectorErrorState error={planner.error} connector="Planner" action={<Button variant="outline" size="sm" onClick={() => void planner.refetch()}><RefreshCw aria-hidden="true" />Try again</Button>} />;

  const { tasks, plans, more } = planner.data;
  const planOf = (id?: string) => plans.find(plan => plan.id === id);
  const order = { overdue: 0, today: 1, week: 2, later: 3, none: 4, done: 5 };
  const visible = tasks.filter(task => {
    const state = plannerTaskState(task);
    return view === "open" ? state !== "done" : view === "overdue" ? state === "overdue" : view === "week" ? state === "today" || state === "week" : state === "done";
  }).sort((a, b) => view === "done"
    ? (b.completed ?? "").localeCompare(a.completed ?? "")
    : order[plannerTaskState(a)] - order[plannerTaskState(b)] || (a.due ?? "9999").localeCompare(b.due ?? "9999"));

  // Group the visible window by plan, keeping the sorted order of each plan's first task.
  const windowed = visible.slice(0, shown);
  const groups: { plan?: PlannerPlan; planId?: string; tasks: PlannerTask[] }[] = [];
  for (const task of windowed) {
    let group = groups.find(g => g.planId === task.planId);
    if (!group) { group = { plan: planOf(task.planId), planId: task.planId, tasks: [] }; groups.push(group); }
    group.tasks.push(task);
  }

  return <div className="space-y-4">
    <dl className="grid grid-cols-2 gap-2 sm:grid-cols-4">
      {VIEWS.map(([key, label]) => <div key={key} className="rounded-xl border border-border bg-card px-3 py-2">
        <dt className="text-xs font-medium text-muted-foreground">{label}</dt>
        <dd className={cn("font-heading text-2xl font-bold tabular-nums", key === "overdue" && counts.overdue > 0 && "text-red-700 dark:text-red-400")}>{counts[key]}</dd>
      </div>)}
    </dl>
    <div className="flex flex-wrap items-center gap-2">
    <div role="group" aria-label="Planner view" className="flex flex-wrap gap-2">
      {VIEWS.map(([key, label]) => <Button key={key} size="sm" variant={view === key ? "default" : "outline"} aria-pressed={view === key} className="rounded-full"
        onClick={() => { setView(key); setShown(WINDOW); }}>{label} ({counts[key]})</Button>)}
    </div>
    <InfoTip label="Planner tasks" title="Which Planner tasks appear" points={[
      "Only tasks assigned to you, in plans you're a member of.",
      "Tasks in Planner's My Day or Private tasks are stored in To Do, so they appear in the To Do parts of the dashboard instead.",
      "Tasks in personal plans (made without a team or group) may not appear.",
      "If Planner returns only part of your tasks, a note says so and the counts cover what was returned.",
    ]} />
    </div>
    {more && <p className="rounded-xl border border-amber-700/40 bg-amber-50 p-3 text-sm text-amber-950 dark:bg-amber-950/30 dark:text-amber-100">
      Planner returned only part of your tasks. Counts here cover the tasks shown; <Out href={links.myTasks}>open My tasks in Planner</Out> for the full list.
    </p>}
    {tasks.length === 0 ? <EmptyState title="No Planner tasks assigned to you" className="py-8" />
      : visible.length === 0 ? <p className="rounded-xl border border-dashed border-border p-4 text-sm text-muted-foreground">Nothing in this view.</p>
      : <div className="space-y-4">
        {groups.map(group => <section key={group.planId ?? "none"} aria-label={group.plan?.title ?? "Plan"}>
          <h3 className="flex items-center justify-between gap-2 border-b border-border pb-1 text-sm font-bold">
            {group.planId ? <Out href={links.plan(group.planId)}>{group.plan?.title ?? "Plan"}</Out> : <span>Plan</span>}
            <span className="text-xs font-normal text-muted-foreground">{group.tasks.length} shown</span>
          </h3>
          <ul className="divide-y divide-border">{group.tasks.map(task => <TaskRow key={task.id} task={task} plan={group.plan} link={links.task(task.id)} />)}</ul>
        </section>)}
        {visible.length > shown && <Button variant="outline" size="sm" className="rounded-full" onClick={() => setShown(value => value + WINDOW)}>
          Show {Math.min(WINDOW, visible.length - shown)} more of {visible.length - shown}
        </Button>}
      </div>}
    <p className="flex flex-wrap gap-x-4 gap-y-1 text-sm">
      <Out href={links.myTasks} className="inline-flex items-center gap-1">My tasks in Planner<ExternalLink aria-hidden="true" className="size-3.5" /></Out>
      <Out href={links.home} className="inline-flex items-center gap-1">All my plans<ExternalLink aria-hidden="true" className="size-3.5" /></Out>
    </p>
  </div>;
}

/** Links for the Planner section's second tab: every plan you belong to. */
export function PlannerPlans() {
  const profile = useProfile();
  const links = plannerLinks(profile.data?.userPrincipalName ?? profile.data?.mail ?? undefined);
  const planner = useQuery({ queryKey: ["planner"], queryFn: loadPlanner });
  if (planner.isPending) return <p role="status" className="text-sm text-muted-foreground">Loading your plans…</p>;
  if (planner.isError) return <ConnectorErrorState error={planner.error} connector="Planner" />;
  const plans = [...planner.data.plans].sort((a, b) => a.title.localeCompare(b.title));
  if (!plans.length) return <p className="text-sm text-muted-foreground">You aren't a member of any plans.</p>;
  const openIn = (id: string) => planner.data.tasks.filter(task => task.planId === id && task.percentComplete < 100).length;
  return <ul className="grid gap-2 sm:grid-cols-2">{plans.map(plan => <li key={plan.id}>
    <a href={links.plan(plan.id)} target="_blank" rel="noopener noreferrer" className="flex items-center justify-between gap-2 rounded-xl border border-border p-3 text-sm hover:bg-muted focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-ring">
      <span className="min-w-0"><span className="block font-semibold text-primary break-words">{plan.title}</span><span className="text-xs text-muted-foreground">{openIn(plan.id)} open assigned to you</span></span>
      <ExternalLink aria-hidden="true" className="size-4 shrink-0" /><span className="sr-only"> (opens in a new tab)</span>
    </a>
  </li>)}</ul>;
}
