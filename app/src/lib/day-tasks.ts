import { useQuery } from "@tanstack/react-query";
import { useProfile } from "@/hooks/use-work-signals";
import { hideEmails } from "@/lib/privacy";
import { loadPlanner, plannerLinks, type PlannerTask } from "@/lib/planner";
import { TODO_HOME, useOpenTasks, type OpenTask } from "@/lib/todo";

export const MY_DAY_URL = "https://to-do.office.com/tasks/myday";

/** Local calendar day as YYYY-MM-DD, the same form To Do and Planner due dates use here. */
export function isoDay(date: Date) { return date.toLocaleDateString("en-CA"); }

export interface DayTask { key: string; title: string; source: "To Do" | "Planner"; where: string; due?: string; url: string; overdue: boolean; important: boolean }

function fromTodo(task: OpenTask, today: string): DayTask {
  return {
    key: `todo-${task.id}`, title: task.title, source: "To Do", where: task.list, due: task.due, url: `${TODO_HOME}${encodeURIComponent(task.listId)}`,
    overdue: Boolean(task.due && task.due < today), important: task.importance === "high",
  };
}
function fromPlanner(task: PlannerTask, today: string, planTitle: string | undefined, url: string): DayTask {
  return {
    key: `planner-${task.id}`, title: hideEmails(task.title), source: "Planner", where: planTitle ?? "Planner", due: task.due, url,
    overdue: Boolean(task.due && task.due < today), important: task.priority !== undefined && task.priority <= 4,
  };
}

/** Open To Do and Planner tasks with a due date, keyed by day — shared by the week view and the agenda. */
export function useDatedTasks() {
  const todo = useOpenTasks();
  const profile = useProfile();
  const planner = useQuery({ queryKey: ["planner"], queryFn: loadPlanner });
  const links = plannerLinks(profile.data?.userPrincipalName ?? profile.data?.mail ?? undefined);
  const today = isoDay(new Date());
  const all: DayTask[] = [
    ...(todo.data ?? []).map(task => fromTodo(task, today)),
    ...(planner.data?.tasks ?? []).filter(task => task.percentComplete < 100)
      .map(task => fromPlanner(task, today, planner.data?.plans.find(plan => plan.id === task.planId)?.title, links.task(task.id))),
  ];
  const undatedTodo = (todo.data ?? []).filter(task => !task.due).map(task => fromTodo(task, today));
  return {
    dueOn: (day: string) => all.filter(task => task.due === day),
    overdue: all.filter(task => task.overdue),
    undatedTodo,
    loading: todo.isPending || planner.isPending,
    failed: [todo.isError && "To Do", planner.isError && "Planner"].filter(Boolean) as string[],
  };
}

