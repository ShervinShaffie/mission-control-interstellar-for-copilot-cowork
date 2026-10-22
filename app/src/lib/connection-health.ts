/**
 * Last known state of each connection this session. Updated by real calls the app already makes, and by the
 * checks in the Connections panel. Nothing is stored after the page closes.
 */
import { useSyncExternalStore } from "react";
import { Office365OutlookService } from "../../generated/services/Office365OutlookService";
import { MicrosoftTeamsService } from "../../generated/services/MicrosoftTeamsService";
import { MicrosoftTo_Do_Business_Service } from "../../generated/services/MicrosoftTo_Do_Business_Service";
import { OneNote_Business_Service } from "../../generated/services/OneNote_Business_Service";
import { OneDriveforBusinessService } from "../../generated/services/OneDriveforBusinessService";
import { Office365UsersService } from "../../generated/services/Office365UsersService";
import { GitHubService } from "../../generated/services/GitHubService";
import { GoogleDriveService } from "../../generated/services/GoogleDriveService";
import { BoxService } from "../../generated/services/BoxService";
import { PlannerService } from "../../generated/services/PlannerService";
import { GoogleCalendarService } from "../../generated/services/GoogleCalendarService";

export type ConnectionKey = "outlook" | "teams" | "todo" | "onenote" | "onedrive" | "users" | "planner" | "github" | "googledrive" | "box" | "googlecalendar" | "copilot" | "sharepoint";
export type Health = "ok" | "signin" | "blocked" | "error" | "unknown" | "checking";
export interface ConnectionState { health: Health; at?: number; detail?: string }

let states: Partial<Record<ConnectionKey, ConnectionState>> = {};
const listeners = new Set<() => void>();
export function reportConnection(key: ConnectionKey, health: Health, detail?: string) {
  states = { ...states, [key]: { health, at: Date.now(), detail } };
  listeners.forEach(listener => listener());
}
export function useConnectionStates() {
  return useSyncExternalStore(listener => { listeners.add(listener); return () => listeners.delete(listener); }, () => states);
}

/** Classifies a connector failure into something a person can act on. */
export function classify(error: unknown): Health {
  const text = (() => { try { return typeof error === "string" ? error : error instanceof Error ? `${error.message} ${JSON.stringify(error)}` : JSON.stringify(error ?? ""); } catch { return String(error); } })();
  if (/\b401\b|unauthori[sz]ed|login|sign[- ]?in|consent|expired|token|ConsentRequired|\b1001\b/i.test(text)) return "signin";
  if (/\b442\b|\b403\b|forbidden|policy|blocked/i.test(text)) return "blocked";
  return "error";
}

type Check = () => Promise<{ success: boolean; error?: unknown }>;
/** Free checks: one small read per direct connection (no Copilot Credits). */
export const FREE_CHECKS: Partial<Record<ConnectionKey, Check>> = {
  outlook: () => Office365OutlookService.CalendarGetTables(),
  teams: () => MicrosoftTeamsService.GetAllTeams(),
  todo: () => MicrosoftTo_Do_Business_Service.GetAllTodoListsV2(),
  onenote: () => OneNote_Business_Service.GetNotebooks(),
  onedrive: () => OneDriveforBusinessService.ListRootFolder(),
  users: () => Office365UsersService.MyProfile_V2("id"),
  planner: () => PlannerService.ListMyPlans_V2(),
  github: () => GitHubService.GetUser(),
  googledrive: () => GoogleDriveService.ListRootFolder(),
  box: () => BoxService.ListRootFolder(),
  googlecalendar: () => GoogleCalendarService.ListCalendars(),
};
export async function runCheck(key: ConnectionKey, check: Check) {
  reportConnection(key, "checking");
  try {
    const result = await check();
    if (result.success) reportConnection(key, "ok");
    else reportConnection(key, classify(result.error));
  } catch (error) { reportConnection(key, classify(error)); }
}

/** The name each connection has in the Power Apps "Connections" page, where people sign in again. */
export const CONNECTOR_NAMES: Record<ConnectionKey, string> = {
  outlook: "Office 365 Outlook", teams: "Microsoft Teams", todo: "Microsoft To-Do (Business)", onenote: "OneNote (Business)",
  onedrive: "OneDrive for Business", users: "Office 365 Users", planner: "Planner", github: "GitHub", googledrive: "Google Drive", box: "Box", googlecalendar: "Google Calendar", copilot: "Work IQ Copilot MCP", sharepoint: "Work IQ SharePoint MCP",
};
/** The Power Apps page listing this person's own connections for the app's environment (sign in / fix / create). */
export function connectionsPageUrl(environmentId?: string) {
  return environmentId && /^[0-9a-z-]{8,64}$/i.test(environmentId) ? `https://make.powerapps.com/environments/${encodeURIComponent(environmentId)}/connections` : "https://make.powerapps.com/connections";
}
