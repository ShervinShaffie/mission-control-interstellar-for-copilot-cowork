import { ClipboardList, ExternalLink, ListTodo } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { MY_DAY_URL, isoDay, useDatedTasks, type DayTask } from "@/lib/day-tasks";
import { cn } from "@/lib/utils";
import { InfoTip } from "@/components/info-tip";

function TaskLine({ task }: { task: DayTask }) {
  const Icon = task.source === "Planner" ? ClipboardList : ListTodo;
  return <li>
    <a href={task.url} target="_blank" rel="noopener noreferrer" className="flex items-start gap-3 px-4 py-2.5 hover:bg-accent focus-visible:outline-2 focus-visible:outline-ring">
      <Icon aria-hidden="true" className="mt-0.5 size-4 shrink-0 text-muted-foreground" />
      <span className="min-w-0 flex-1">
        <span className="block break-words text-sm font-medium">{task.title}</span>
        <span className="mt-0.5 flex flex-wrap items-center gap-x-2 text-xs text-muted-foreground">
          <span>{task.source} · {task.where}</span>
          {task.overdue && task.due && <span className="font-semibold text-red-700 dark:text-red-400">Overdue · was due {new Date(`${task.due}T12:00:00`).toLocaleDateString([], { month: "short", day: "numeric" })}</span>}
          {task.important && <Badge variant="outline" className="h-auto">{task.source === "Planner" ? "Important" : "High importance"}</Badge>}
        </span>
      </span>
      <ExternalLink aria-hidden="true" className="mt-0.5 size-4 shrink-0" /><span className="sr-only">(opens {task.source} in a new tab)</span>
    </a>
  </li>;
}

/** The to-dos for one agenda day: due that day, plus (on today) anything overdue and To Do tasks with no due date. */
export function AgendaTasks({ day, isToday }: { day: Date; isToday: boolean }) {
  const tasks = useDatedTasks();
  const due = tasks.dueOn(isoDay(day));
  const overdue = isToday ? tasks.overdue : [];
  const undated = isToday ? tasks.undatedTodo : [];
  return <section aria-labelledby="agenda-tasks-heading" className="border-t border-border">
    <div className="flex flex-wrap items-center justify-between gap-2 px-4 pt-3">
      <h3 id="agenda-tasks-heading" className="flex items-center gap-1 text-sm font-bold">To-dos{isToday ? " for today" : " due this day"}
        <InfoTip label="to-dos on the calendar" title="Which tasks appear here" points={[
          "To Do and Planner tasks appear on the day they're due.",
          "Planner tasks appear only when they're assigned to you.",
          "To Do doesn't tell other apps which tasks are in My Day, and tasks added in My Day have no due date — so today's list includes every To Do task without a due date.",
          "Give a task a due date in To Do or Planner to place it on a specific day.",
          "Reads up to 6 To Do lists, 50 tasks each.",
        ]} />
      </h3>
      <a href={MY_DAY_URL} target="_blank" rel="noopener noreferrer" className="text-xs font-semibold text-primary underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-ring">
        Open My Day in To Do<span className="sr-only"> (opens in a new tab)</span>
      </a>
    </div>
    {tasks.loading ? <p role="status" className="px-4 py-3 text-sm text-muted-foreground">Loading tasks…</p>
      : <>
        {overdue.length > 0 && <><p className="px-4 pt-2 text-xs font-semibold uppercase tracking-wide text-red-700 dark:text-red-400">Overdue ({overdue.length})</p>
          <ul className="divide-y divide-border">{overdue.map(task => <TaskLine key={task.key} task={task} />)}</ul></>}
        {due.length > 0 && <><p className="px-4 pt-2 text-xs font-semibold uppercase tracking-wide text-muted-foreground">Due {isToday ? "today" : "this day"} ({due.length})</p>
          <ul className="divide-y divide-border">{due.map(task => <TaskLine key={task.key} task={task} />)}</ul></>}
        {isToday && undated.length > 0 && <details className="group">
          <summary className={cn("cursor-pointer px-4 py-2 text-sm font-medium focus-visible:outline-2 focus-visible:outline-ring")}>
            To Do tasks with no due date ({undated.length})
          </summary>
          <p className="px-4 pb-1 text-xs text-muted-foreground">Tasks you add in My Day have no due date, and To Do doesn't tell other apps which tasks are in My Day — so they're listed here. Give a task a due date in To Do to place it on a specific day.</p>
          <ul className="divide-y divide-border">{undated.map(task => <TaskLine key={task.key} task={task} />)}</ul>
        </details>}
        {due.length === 0 && overdue.length === 0 && undated.length === 0 && <p className="px-4 py-3 text-sm text-muted-foreground">No To Do or Planner tasks due {isToday ? "today" : "this day"}.</p>}
        {tasks.failed.length > 0 && <p role="alert" className="px-4 py-2 text-xs text-red-700 dark:text-red-400">Couldn't load {tasks.failed.join(" or ")} tasks; others are shown.</p>}
      </>}
    <div className="h-2" />
  </section>;
}
