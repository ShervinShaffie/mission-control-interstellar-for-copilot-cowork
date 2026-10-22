import type { ComponentId, SearchGroup, SearchSwitches } from "@/lib/dashboard-preferences";

export type { SearchGroup };
export const SEARCH_LABELS: Record<SearchGroup, string> = {
  priorities: "Team activity (projects and deadlines)",
  meetings: "Meeting prep and learning",
  replies: "Emails needing a reply",
  collaboration: "Who's away, follow-ups and recent folders",
  pinned: "Changes in your pinned SharePoint sites (SharePoint)",
};
/** Plain-language list of what each search fills in (shown under its switch in Customize). */
export const SEARCH_FEEDS: Record<SearchGroup, string> = {
  priorities: "Fills in: Today's priorities, Projects, and Work actions",
  meetings: "Fills in: meeting prep notes, Projects' meetings, and Learning",
  replies: "Fills in: Email → Needs reply, and replies in Today",
  collaboration: "Fills in: Who's away, people to follow up with, and recently active folders",
  pinned: "Fills in: files changed this week in the sites you pin in Files, shown in Today. Runs only when you've pinned a site, uses a small amount of credits per site, and needs your SharePoint sign-in",
};
// Which components use which Copilot search.
const NEEDS: Record<SearchGroup, ComponentId[]> = {
  priorities: ["today", "projects", "files"],
  meetings: ["today", "projects", "calendar", "files"],
  replies: ["today", "email", "files"],
  collaboration: ["today", "calendar"],
  pinned: ["today"],
};
export type SearchPlan = Record<SearchGroup, boolean>;
/** A Copilot search runs only when it is switched on AND a component on the dashboard uses it. */
export function searchPlan(visible: readonly ComponentId[], switches: SearchSwitches): SearchPlan {
  const on = (group: SearchGroup) => switches[group] && NEEDS[group].some(id => visible.includes(id));
  return { priorities: on("priorities"), meetings: on("meetings"), replies: on("replies"), collaboration: on("collaboration"), pinned: on("pinned") };
}
export function plannedSearches(plan: SearchPlan): SearchGroup[] {
  return (Object.keys(plan) as SearchGroup[]).filter(group => plan[group]);
}
