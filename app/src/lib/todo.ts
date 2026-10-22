import { useQuery } from "@tanstack/react-query";
import { MicrosoftTo_Do_Business_Service } from "../../generated/services/MicrosoftTo_Do_Business_Service";
import { hideEmails } from "@/lib/privacy";

export const TODO_HOME = "https://to-do.office.com/tasks/";
export interface OpenTask { id: string; title: string; list: string; listId: string; due?: string; importance?: string; status?: string }

function dayOf(value?: string) { return value && /^\d{4}-\d{2}-\d{2}/.test(value) ? value.slice(0, 10) : undefined; }
function dayKey(date: Date) { return date.toLocaleDateString("en-CA"); }

export function taskState(task: OpenTask, now = new Date()): "overdue" | "today" | "tomorrow" | "week" | "later" | "none" {
  if (!task.due) return "none";
  const today = dayKey(now), tomorrow = new Date(now), week = new Date(now);
  tomorrow.setDate(tomorrow.getDate() + 1); week.setDate(week.getDate() + 7);
  if (task.due < today) return "overdue";
  if (task.due === today) return "today";
  if (task.due === dayKey(tomorrow)) return "tomorrow";
  return task.due <= dayKey(week) ? "week" : "later";
}

/** Free read of the user's open Microsoft To Do tasks (no Copilot). Up to 6 lists, 50 tasks each. */
export async function loadOpenTasks(): Promise<OpenTask[]> {
  const lists = await MicrosoftTo_Do_Business_Service.GetAllTodoListsV2();
  if (!lists.success) throw lists.error ?? new Error("To Do lists unavailable");
  const chosen = (lists.data ?? []).filter(list => list.id)
    .sort((a, b) => (a.wellknownListName === "defaultList" ? -1 : 0) - (b.wellknownListName === "defaultList" ? -1 : 0)).slice(0, 6);
  const pages = await Promise.all(chosen.map(async list => {
    const result = await MicrosoftTo_Do_Business_Service.ListToDosByFolderV2(list.id!, 50);
    if (!result.success) return [];
    return (result.data ?? []).filter(task => task.id && task.title && task.status !== "completed").map((task): OpenTask => ({
      id: task.id!, title: hideEmails(task.title!), list: list.displayName ?? "Tasks", listId: list.id!,
      due: dayOf(task.dueDateTime?.dateTime), importance: task.importance, status: task.status,
    }));
  }));
  return pages.flat().sort((a, b) => (a.due ?? "9999").localeCompare(b.due ?? "9999") || (a.importance === "high" ? -1 : 0) - (b.importance === "high" ? -1 : 0));
}

export function useOpenTasks() {
  return useQuery({ retry: false, queryKey: ["todo-open-tasks"], queryFn: loadOpenTasks });
}
