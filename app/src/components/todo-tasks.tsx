import { ExternalLink, ListTodo } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { EmptyState, ErrorState, LoadingState } from "@/components/states";
import { TODO_HOME, taskState, useOpenTasks, type OpenTask } from "@/lib/todo";

function dueLabel(task: OpenTask) {
  if (!task.due) return "No due date";
  const state = taskState(task);
  const date = new Date(`${task.due}T12:00:00`).toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" });
  return state === "overdue" ? `Overdue · was due ${date}` : state === "today" ? "Due today" : state === "tomorrow" ? "Due tomorrow" : `Due ${date}`;
}

/** Open To Do tasks. dueOnly shows just overdue and due-today tasks (used on Today). */
export function TodoTasks({ dueOnly = false, limit = 15, heading }: { dueOnly?: boolean; limit?: number; heading?: string }) {
  const tasks = useOpenTasks();
  const rows = (tasks.data ?? []).filter(task => !dueOnly || ["overdue", "today"].includes(taskState(task))).slice(0, limit);
  if (tasks.isPending) return <LoadingState rows={2} />;
  if (tasks.isError) return <div role="alert"><ErrorState title="Could not load your To Do tasks" description="Retry to read Microsoft To Do again." action={<Button variant="outline" disabled={tasks.isFetching} onClick={() => void tasks.refetch()}>Retry</Button>} /></div>;
  if (!rows.length) return dueOnly ? null : <EmptyState icon={ListTodo} title="No open To Do tasks" description="Tasks you add in Microsoft To Do or flag in Outlook show up here. Planner tasks have their own section." />;
  return <div>{heading && <h3 className="mb-2 text-sm font-semibold">{heading}</h3>}<ul className="divide-y divide-border rounded-xl border border-border">
    {rows.map(task => {
      const state = taskState(task);
      return <li key={task.id}>
        <a href={`${TODO_HOME}${encodeURIComponent(task.listId)}`} target="_blank" rel="noreferrer" className="flex items-start gap-3 p-3 hover:bg-accent focus-visible:outline-2 focus-visible:outline-ring">
          <ListTodo aria-hidden="true" className="mt-0.5 size-4 shrink-0 text-muted-foreground" />
          <div className="min-w-0 flex-1">
            <p className="break-words text-sm font-medium">{task.title}</p>
            <p className="mt-1 flex flex-wrap items-center gap-2 text-xs text-muted-foreground">
              <span className={state === "overdue" ? "font-semibold text-red-700 dark:text-red-400" : state === "today" ? "font-semibold text-amber-700 dark:text-amber-400" : undefined}>{dueLabel(task)}</span>
              <span>· {task.list}</span>
              {task.importance === "high" && <Badge variant="outline">High importance</Badge>}
            </p>
          </div>
          <ExternalLink aria-hidden="true" className="mt-0.5 size-4 shrink-0" />
          <span className="sr-only">(opens Microsoft To Do in a new tab)</span>
        </a>
      </li>;
    })}
  </ul></div>;
}
