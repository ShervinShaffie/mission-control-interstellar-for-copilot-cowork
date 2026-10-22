/**
 * Planner reads for the Planner section: your tasks across every plan, your plans, and each plan's label names.
 * The connector answers "my tasks" in one response; when Planner says more exist, the section says so rather
 * than presenting a partial list as complete.
 */
import { PlannerService } from "../../generated/services/PlannerService";
import { classify, reportConnection } from "@/lib/connection-health";

type Raw = Record<string, unknown>;
const str = (value: unknown) => typeof value === "string" && value.trim() ? value : undefined;
const num = (value: unknown) => typeof value === "number" && Number.isFinite(value) ? value : undefined;
const obj = (value: unknown): Raw | undefined => typeof value === "object" && value !== null && !Array.isArray(value) ? value as Raw : undefined;

export interface PlannerTask {
  id: string; title: string; planId?: string; bucketId?: string; percentComplete: number; due?: string; start?: string; created?: string;
  completed?: string; priority?: number; categories: string[]; checklistTotal?: number; checklistOpen?: number; hasDescription: boolean;
  assigneeCount: number; createdBy?: string;
}
export interface PlannerPlan { id: string; title: string; groupId?: string; labels: Record<string, string> }
export interface PlannerData { tasks: PlannerTask[]; plans: PlannerPlan[]; more: boolean }

function fail(label: string, error: unknown): never {
  console.error(`${label} request failed`, error);
  reportConnection("planner", classify(error));
  throw error ?? new Error(`${label} unavailable`);
}

function dayOf(value?: string) { return value && /^\d{4}-\d{2}-\d{2}/.test(value) ? value.slice(0, 10) : undefined; }

function toTask(row: unknown): PlannerTask | null {
  const task = obj(row);
  const id = str(task?.id), title = str(task?.title);
  if (!task || !id || !title) return null;
  const categories = Object.entries(obj(task.appliedCategories) ?? {}).filter(([, on]) => on === true).map(([key]) => key);
  const assignments = obj(task.assignments) ?? obj(task._assignments);
  const assigneeCount = Array.isArray(task._assignments) ? task._assignments.length : assignments ? Object.keys(assignments).length : 0;
  return {
    id, title, planId: str(task.planId), bucketId: str(task.bucketId), percentComplete: num(task.percentComplete) ?? 0,
    due: dayOf(str(task.dueDateTime)), start: dayOf(str(task.startDateTime)), created: str(task.createdDateTime), completed: str(task.completedDateTime),
    priority: num(task.priority), categories, checklistTotal: num(task.checklistItemCount), checklistOpen: num(task.activeChecklistItemCount),
    hasDescription: task.hasDescription === true, assigneeCount, createdBy: str(obj(obj(task.createdBy)?.user)?.displayName),
  };
}

/** Label names for one plan: Planner stores them as category1…category25 in the plan's details. */
async function loadPlanLabels(planId: string): Promise<Record<string, string>> {
  const result = await PlannerService.GetPlanDetails(planId);
  if (!result.success) return {};
  const descriptions = obj(obj(result.data as unknown)?.categoryDescriptions) ?? {};
  return Object.fromEntries(Object.entries(descriptions).flatMap(([key, value]) => { const name = str(value); return name ? [[key, name]] : []; }));
}

const LABEL_PLANS = 8;
export async function loadPlanner(): Promise<PlannerData> {
  const [tasksRes, plansRes] = await Promise.all([PlannerService.ListMyTasks_V2(), PlannerService.ListMyPlans_V2()]);
  if (!tasksRes.success) fail("Planner tasks", tasksRes.error);
  const taskData = obj(tasksRes.data as unknown);
  if (!taskData || !Array.isArray(taskData.value)) fail("Planner tasks", new Error("Unexpected Planner tasks shape"));
  reportConnection("planner", "ok");
  const tasks = taskData.value.map(toTask).filter((task): task is PlannerTask => task !== null);
  // Plans only add names; if they can't be read, tasks still show.
  const planRows = plansRes.success && Array.isArray(obj(plansRes.data as unknown)?.value) ? obj(plansRes.data as unknown)!.value as unknown[] : [];
  const plans: PlannerPlan[] = planRows.flatMap(row => {
    const plan = obj(row); const id = str(plan?.id);
    if (!plan || !id) return [];
    return [{ id, title: str(plan.title) ?? "Untitled plan", groupId: str(obj(plan.container)?.containerId) ?? str(plan.owner), labels: {} }];
  });
  // Label names for the plans your open tasks are in (bounded).
  const busy = [...new Set(tasks.filter(task => task.percentComplete < 100 && task.categories.length && task.planId).map(task => task.planId!))].slice(0, LABEL_PLANS);
  const labels = await Promise.all(busy.map(async id => [id, await loadPlanLabels(id)] as const));
  for (const [id, names] of labels) {
    const plan = plans.find(p => p.id === id);
    if (plan) plan.labels = names; else plans.push({ id, title: "Plan", labels: names });
  }
  return { tasks, plans, more: Boolean(str(taskData["@odata.nextLink"])) };
}

export function plannerTaskState(task: PlannerTask, now = new Date()): "done" | "overdue" | "today" | "week" | "later" | "none" {
  if (task.percentComplete >= 100) return "done";
  if (!task.due) return "none";
  const today = now.toLocaleDateString("en-CA"), week = new Date(now);
  week.setDate(week.getDate() + 7);
  if (task.due < today) return "overdue";
  if (task.due === today) return "today";
  return task.due <= week.toLocaleDateString("en-CA") ? "week" : "later";
}

export const PRIORITY: Record<string, string> = { urgent: "Urgent", important: "Important", medium: "Medium", low: "Low" };
/** Planner priority is 0–10: 0–1 urgent, 2–4 important, 5–7 medium, 8–10 low. */
export function priorityOf(value?: number) {
  if (value === undefined) return undefined;
  return value <= 1 ? "urgent" : value <= 4 ? "important" : value <= 7 ? "medium" : "low";
}

/** Planner web links. With your sign-in domain they open the exact task or plan; otherwise Planner's home. */
export function plannerLinks(upn?: string) {
  const domain = upn?.split("@")[1];
  const base = domain && /^[a-z0-9.-]+$/i.test(domain) ? `https://tasks.office.com/${domain}/Home` : undefined;
  return {
    home: base ? `${base}/Planner` : "https://planner.cloud.microsoft/",
    myTasks: base ? `${base}/Planner/#/mytasks` : "https://planner.cloud.microsoft/",
    task: (id: string) => base ? `${base}/Task/${encodeURIComponent(id)}` : "https://planner.cloud.microsoft/",
    plan: (id: string) => base ? `${base}/PlanViews/${encodeURIComponent(id)}` : "https://planner.cloud.microsoft/",
  };
}
