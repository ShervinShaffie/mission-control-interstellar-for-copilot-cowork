/**
 * Reads for the Plugins section: GitHub, Google Drive, Box and Google Calendar, through the app's own connections.
 * Each adapter accepts only the shapes the connector documents and throws on anything else, so a
 * connector problem never shows up as an empty list.
 */
import { GitHubService } from "../../generated/services/GitHubService";
import { GoogleDriveService } from "../../generated/services/GoogleDriveService";
import { BoxService } from "../../generated/services/BoxService";
import { GoogleCalendarService } from "../../generated/services/GoogleCalendarService";
import type { BlobMetadata } from "../../generated/models/GoogleDriveModel";
import { classify, reportConnection } from "@/lib/connection-health";

export const REPO_PAGE_SIZE = 8;
export const WORK_PAGE_SIZE = 10;

type Raw = Record<string, unknown>;
const str = (value: unknown) => typeof value === "string" && value.trim() ? value : undefined;
const num = (value: unknown) => typeof value === "number" && Number.isFinite(value) ? value : undefined;
const obj = (value: unknown): Raw | undefined => typeof value === "object" && value !== null && !Array.isArray(value) ? value as Raw : undefined;

function fail(key: "github" | "googledrive" | "box" | "googlecalendar", label: string, error: unknown): never {
  console.error(`${label} request failed`, error);
  reportConnection(key, classify(error));
  throw error ?? new Error(`${label} unavailable`);
}

/* ---------------- GitHub ---------------- */

export interface GitHubProfile {
  login: string; name?: string; url?: string; bio?: string; company?: string; location?: string; website?: string; accountType?: string;
  memberSince?: string; publicRepos?: number; privateRepos?: number; publicGists?: number; followers?: number; following?: number; plan?: string;
}
export interface GitHubRepo {
  id: string; owner: string; name: string; fullName: string; url?: string; homepage?: string; description?: string;
  isPrivate: boolean; isFork: boolean; isArchived: boolean; isTemplate: boolean; language?: string; topics: string[]; license?: string;
  defaultBranch?: string; stars: number; forks: number; watchers: number; openIssues: number; sizeKb?: number;
  createdAt?: string; pushedAt?: string; hasIssues: boolean;
}
export interface GitHubLabel { name: string; color?: string }
export interface GitHubWorkItem {
  id: string; title: string; url?: string; number?: number; repo?: string; kind: "pull request" | "issue"; state?: string; isDraft: boolean;
  createdAt?: string; updatedAt?: string; comments?: number; author?: string; assignees: string[]; labels: GitHubLabel[]; milestone?: string;
}
export interface GitHubMilestone { id: string; title: string; url?: string; open: number; closed: number; dueOn?: string; description?: string }
export interface GitHubRepoDetail { pulls: GitHubWorkItem[]; issues: GitHubWorkItem[]; milestones: GitHubMilestone[]; labelCount?: number }

const strs = (value: unknown) => Array.isArray(value) ? value.filter((v): v is string => typeof v === "string" && v.trim() !== "") : [];
const logins = (value: unknown) => Array.isArray(value) ? value.flatMap(v => { const l = str(obj(v)?.login); return l ? [l] : []; }) : [];
const labelsOf = (value: unknown): GitHubLabel[] => Array.isArray(value) ? value.flatMap(v => {
  const label = obj(v); const name = str(label?.name) ?? (typeof v === "string" ? v : undefined);
  return name ? [{ name, color: str(label?.color) }] : [];
}) : [];

export async function loadGitHubProfile(): Promise<GitHubProfile> {
  const result = await GitHubService.GetUser();
  if (!result.success) fail("github", "GitHub profile", result.error);
  const user = obj(result.data);
  const login = str(user?.login);
  if (!user || !login) fail("github", "GitHub profile", new Error("Unexpected GitHub profile shape"));
  reportConnection("github", "ok");
  const blog = str(user.blog);
  return {
    login, name: str(user.name), url: str(user.html_url), bio: str(user.bio), company: str(user.company), location: str(user.location),
    website: blog && !/^https?:\/\//i.test(blog) ? `https://${blog}` : blog, accountType: str(user.type), memberSince: str(user.created_at),
    publicRepos: num(user.public_repos), privateRepos: num(user.total_private_repos), publicGists: num(user.public_gists),
    followers: num(user.followers), following: num(user.following), plan: str(obj(user.plan)?.name),
  };
}

function toRepo(row: unknown): GitHubRepo | null {
  const repo = obj(row);
  const name = str(repo?.name), id = repo?.id;
  if (!repo || !name || (typeof id !== "number" && typeof id !== "string")) return null;
  const fullName = str(repo.full_name) ?? name;
  return {
    id: String(id), owner: str(obj(repo.owner)?.login) ?? fullName.split("/")[0], name, fullName, url: str(repo.html_url), homepage: str(repo.homepage),
    description: str(repo.description), isPrivate: repo.private === true, isFork: repo.fork === true, isArchived: repo.archived === true,
    isTemplate: repo.is_template === true, language: str(repo.language), topics: strs(repo.topics),
    license: str(obj(repo.license)?.spdx_id) !== "NOASSERTION" ? str(obj(repo.license)?.spdx_id) ?? str(obj(repo.license)?.name) : str(obj(repo.license)?.name),
    defaultBranch: str(repo.default_branch), stars: num(repo.stargazers_count) ?? 0, forks: num(repo.forks_count) ?? 0,
    watchers: num(repo.subscribers_count) ?? num(repo.watchers_count) ?? 0, openIssues: num(repo.open_issues_count) ?? 0, sizeKb: num(repo.size),
    createdAt: str(repo.created_at), pushedAt: str(repo.pushed_at) ?? str(repo.updated_at), hasIssues: repo.has_issues !== false,
  };
}

/** One page of your repositories, most recently updated first. `page` is GitHub's own 1-based page number. */
export async function loadGitHubRepos(page: number): Promise<{ rows: GitHubRepo[]; hasMore: boolean }> {
  const result = await GitHubService.GetUserRepos(undefined, undefined, undefined, undefined, undefined, "updated", "desc", REPO_PAGE_SIZE, page);
  if (!result.success) fail("github", "GitHub repositories", result.error);
  if (!Array.isArray(result.data)) fail("github", "GitHub repositories", new Error("Unexpected GitHub repositories shape"));
  reportConnection("github", "ok");
  const rows = result.data.map(toRepo).filter((row): row is GitHubRepo => row !== null);
  // GitHub pages by number: a page shorter than the page size is the last one.
  return { rows, hasMore: result.data.length >= REPO_PAGE_SIZE };
}

function repoFromApiUrl(url?: string) {
  const match = url?.match(/\/repos\/([^/]+\/[^/]+)$/);
  return match?.[1];
}
function toWorkItem(row: unknown, fallbackRepo?: string, forceKind?: "pull request" | "issue"): GitHubWorkItem | null {
  const item = obj(row);
  const title = str(item?.title), id = item?.id;
  if (!item || !title || (typeof id !== "number" && typeof id !== "string")) return null;
  const assignees = logins(item.assignees);
  const single = str(obj(item.assignee)?.login);
  if (single && !assignees.includes(single)) assignees.push(single);
  return {
    id: String(id), title, url: str(item.html_url), number: num(item.number), repo: repoFromApiUrl(str(item.repository_url)) ?? fallbackRepo,
    kind: forceKind ?? (obj(item.pull_request) ? "pull request" : "issue"), state: str(item.state), isDraft: item.draft === true,
    createdAt: str(item.created_at), updatedAt: str(item.updated_at), comments: num(item.comments), author: str(obj(item.user)?.login),
    assignees, labels: labelsOf(item.labels), milestone: str(obj(item.milestone)?.title),
  };
}

/** Open issues and pull requests that involve you (author, assignee, mention, reviewer), newest activity first. */
export async function loadGitHubWork(login: string, page: number): Promise<{ rows: GitHubWorkItem[]; hasMore: boolean; total?: number }> {
  const result = await GitHubService.SearchIssues(`involves:${login} is:open`, "updated", "desc", WORK_PAGE_SIZE, page);
  if (!result.success) fail("github", "GitHub issues and pull requests", result.error);
  const data = obj(result.data);
  const body = data?.body;
  // GitHub's search answer is { total_count, items }; the connector may return it as-is or under `body`.
  const envelope = obj(body) ?? data;
  const items = Array.isArray(body) ? body : Array.isArray(envelope?.items) ? envelope.items : undefined;
  if (!items) fail("github", "GitHub issues and pull requests", new Error("Unexpected GitHub search shape"));
  reportConnection("github", "ok");
  const total = num(envelope?.total_count);
  const rows = items.map(row => toWorkItem(row)).filter((row): row is GitHubWorkItem => row !== null);
  return { rows, hasMore: total !== undefined ? page * WORK_PAGE_SIZE < total : items.length >= WORK_PAGE_SIZE, total };
}

const DETAIL_SIZE = 5;
function listOrFail(result: { success: boolean; data?: unknown; error?: unknown }, label: string): unknown[] {
  if (!result.success) fail("github", label, result.error);
  if (!Array.isArray(result.data)) fail("github", label, new Error(`Unexpected ${label} shape`));
  return result.data;
}
/** A repository's newest open pull requests and issues, open milestones and label count — read only when you expand it. */
export async function loadGitHubRepoDetail(owner: string, name: string): Promise<GitHubRepoDetail> {
  const full = `${owner}/${name}`;
  const [pullsRes, issuesRes, milestonesRes, labelsRes] = await Promise.all([
    GitHubService.GetPullRequests(owner, name, "open", undefined, undefined, "updated", "desc", DETAIL_SIZE, 1),
    GitHubService.GetIssues(owner, name, undefined, "open", undefined, undefined, undefined, undefined, "updated", "desc", undefined, DETAIL_SIZE * 2, 1),
    GitHubService.GetMilestones(owner, name, "open", "due_on", "asc", DETAIL_SIZE, 1),
    GitHubService.GetLabels(owner, name, 100, 1),
  ]);
  const pulls = listOrFail(pullsRes, "GitHub pull requests").map(row => toWorkItem(row, full, "pull request")).filter((r): r is GitHubWorkItem => r !== null);
  // GitHub's issues list also returns pull requests; keep only true issues.
  const issues = listOrFail(issuesRes, "GitHub issues").filter(row => !obj(obj(row)?.pull_request))
    .map(row => toWorkItem(row, full, "issue")).filter((r): r is GitHubWorkItem => r !== null).slice(0, DETAIL_SIZE);
  const milestones = listOrFail(milestonesRes, "GitHub milestones").flatMap((row): GitHubMilestone[] => {
    const m = obj(row); const title = str(m?.title), id = m?.id;
    if (!m || !title || (typeof id !== "number" && typeof id !== "string")) return [];
    return [{ id: String(id), title, url: str(m.html_url), open: num(m.open_issues) ?? 0, closed: num(m.closed_issues) ?? 0, dueOn: str(m.due_on), description: str(m.description) }];
  });
  const labelCount = labelsRes.success && Array.isArray(labelsRes.data) ? labelsRes.data.length : undefined;
  return { pulls, issues, milestones, labelCount };
}

/* ---------------- Google Drive ---------------- */

export interface DriveItem { id: string; name: string; path?: string; isFolder: boolean; kind: string; mediaType?: string; modified?: string; size?: number; url: string }

const GOOGLE_TYPES: Record<string, string> = {
  "application/vnd.google-apps.spreadsheet": "Google Sheet", "application/vnd.google-apps.document": "Google Doc",
  "application/vnd.google-apps.presentation": "Google Slides", "application/vnd.google-apps.form": "Google Form",
  "application/vnd.google-apps.drawing": "Google Drawing", "application/vnd.google-apps.map": "Google My Map",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "Word document",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "Excel workbook",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation": "PowerPoint deck",
  "application/pdf": "PDF",
};
function driveKind(meta: BlobMetadata) {
  if (meta.IsFolder) return "Folder";
  const type = meta.MediaType ?? "";
  if (GOOGLE_TYPES[type]) return GOOGLE_TYPES[type];
  if (type.startsWith("image/")) return "Image";
  if (type.startsWith("video/")) return "Video";
  if (type.startsWith("text/")) return "Text file";
  const extension = meta.Name?.match(/\.([a-z0-9]{1,5})$/i)?.[1];
  return extension ? `${extension.toUpperCase()} file` : "File";
}
function toDriveItem(meta: BlobMetadata): DriveItem | null {
  const id = str(meta.Id), name = str(meta.DisplayName) ?? str(meta.Name);
  if (!id || !name) return null;
  const encoded = encodeURIComponent(id);
  return {
    id, name, path: str(meta.Path), isFolder: meta.IsFolder === true, kind: driveKind(meta), mediaType: str(meta.MediaType), modified: str(meta.LastModified),
    size: meta.IsFolder ? undefined : num(meta.Size) || undefined,
    url: meta.IsFolder ? `https://drive.google.com/drive/folders/${encoded}` : `https://drive.google.com/open?id=${encoded}`,
  };
}
function sortDrive(rows: DriveItem[]) {
  return rows.sort((a, b) => (b.modified ?? "").localeCompare(a.modified ?? "") || a.name.localeCompare(b.name));
}

/** Everything at the top of your Drive (the connector returns it in one answer), newest first. */
export async function loadDriveRoot(): Promise<DriveItem[]> {
  const result = await GoogleDriveService.ListRootFolder();
  if (!result.success) fail("googledrive", "Google Drive", result.error);
  if (!Array.isArray(result.data)) fail("googledrive", "Google Drive", new Error("Unexpected Google Drive shape"));
  reportConnection("googledrive", "ok");
  return sortDrive(result.data.map(toDriveItem).filter((row): row is DriveItem => row !== null));
}
export async function loadDriveFolder(id: string): Promise<DriveItem[]> {
  const result = await GoogleDriveService.ListFolder(id);
  if (!result.success) fail("googledrive", "Google Drive folder", result.error);
  if (!Array.isArray(result.data)) fail("googledrive", "Google Drive folder", new Error("Unexpected Google Drive shape"));
  return sortDrive(result.data.map(toDriveItem).filter((row): row is DriveItem => row !== null));
}

/* ---------------- Box ---------------- */

export interface BoxItem { id: string; name: string; path?: string; isFolder: boolean; kind: string; mediaType?: string; modified?: string; modifiedBy?: string; size?: number; url: string }

/** Box web link. The connector's ids are Box's own numeric ids; anything else falls back to a Box search for the name. */
function boxUrl(id: string, name: string, isFolder: boolean) {
  if (/^\d+$/.test(id)) return `https://app.box.com/${isFolder ? "folder" : "file"}/${id}`;
  return `https://app.box.com/folder/0/search?query=${encodeURIComponent(name)}`;
}
function toBoxItem(meta: { Id?: string; Path?: string; Name?: string; DisplayName?: string; IsFolder?: boolean; MediaType?: string; LastModified?: string; LastModifiedBy?: string; Size?: number }): BoxItem | null {
  const id = str(meta.Id), name = str(meta.DisplayName) ?? str(meta.Name);
  if (!id || !name) return null;
  const isFolder = meta.IsFolder === true;
  return {
    id, name, path: str(meta.Path), isFolder, kind: driveKind(meta as BlobMetadata), mediaType: str(meta.MediaType), modified: str(meta.LastModified), modifiedBy: str(meta.LastModifiedBy),
    size: isFolder ? undefined : num(meta.Size) || undefined, url: boxUrl(id, name, isFolder),
  };
}
function sortBox(rows: BoxItem[]) {
  return rows.sort((a, b) => (b.modified ?? "").localeCompare(a.modified ?? "") || a.name.localeCompare(b.name));
}
/** Everything at the top of All Files in Box (the connector returns it in one answer), newest first. */
export async function loadBoxRoot(): Promise<BoxItem[]> {
  const result = await BoxService.ListRootFolder();
  if (!result.success) fail("box", "Box", result.error);
  if (!Array.isArray(result.data)) fail("box", "Box", new Error("Unexpected Box shape"));
  reportConnection("box", "ok");
  return sortBox(result.data.map(toBoxItem).filter((row): row is BoxItem => row !== null));
}
export async function loadBoxFolder(id: string): Promise<BoxItem[]> {
  const result = await BoxService.ListFolder(id);
  if (!result.success) fail("box", "Box folder", result.error);
  if (!Array.isArray(result.data)) fail("box", "Box folder", new Error("Unexpected Box shape"));
  return sortBox(result.data.map(toBoxItem).filter((row): row is BoxItem => row !== null));
}

/* ---------------- Google Calendar ---------------- */

export interface GoogleCalendarInfo { id: string; name: string; timeZone?: string; isPrimary: boolean }
export interface GoogleEvent {
  id: string; title: string; start?: Date; end?: Date; allDay: boolean; location?: string; status?: string; organizer?: string; url?: string;
  description?: string; attendees?: string; recurring: boolean;
}

/** Your Google calendars. The primary calendar's id is your Google account address; shared ones end in calendar.google.com. */
export async function loadGoogleCalendars(): Promise<GoogleCalendarInfo[]> {
  const result = await GoogleCalendarService.ListCalendars();
  if (!result.success) fail("googlecalendar", "Google Calendar", result.error);
  const items = result.data?.items;
  if (!Array.isArray(items)) fail("googlecalendar", "Google Calendar", new Error("Unexpected Google Calendar shape"));
  reportConnection("googlecalendar", "ok");
  const calendars = items.flatMap((item): GoogleCalendarInfo[] => {
    const id = str(item.id);
    if (!id) return [];
    return [{ id, name: str(item.summary) ?? id, timeZone: str(item.timeZone), isPrimary: id.includes("@") && !/calendar\.google\.com$/i.test(id) }];
  });
  return calendars.sort((a, b) => Number(b.isPrimary) - Number(a.isPrimary) || a.name.localeCompare(b.name));
}

function eventTime(value?: string): { date?: Date; allDay: boolean } {
  if (!value) return { allDay: false };
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const [y, m, d] = value.split("-").map(Number);
    return { date: new Date(y, m - 1, d), allDay: true };
  }
  const date = new Date(value);
  return { date: Number.isNaN(date.getTime()) ? undefined : date, allDay: false };
}

/**
 * Events on one calendar between `from` and `to`, earliest first. The connector returns a single page for the
 * window (it offers no next-page token), so the window is kept short.
 */
export async function loadGoogleEvents(calendarId: string, from: Date, to: Date, query?: string): Promise<GoogleEvent[]> {
  const result = await GoogleCalendarService.ListEvents(calendarId, from.toISOString(), to.toISOString(), query?.trim() || undefined);
  if (!result.success) fail("googlecalendar", "Google Calendar events", result.error);
  const items = result.data?.items;
  if (!Array.isArray(items)) fail("googlecalendar", "Google Calendar events", new Error("Unexpected Google Calendar events shape"));
  reportConnection("googlecalendar", "ok");
  return items.flatMap((item): GoogleEvent[] => {
    const id = str(item.id);
    if (!id || item.status === "cancelled") return [];
    const start = eventTime(str(item.start)), end = eventTime(str(item.end));
    return [{
      id, title: str(item.summary) ?? "(no title)", start: start.date, end: end.date, allDay: start.allDay, location: str(item.location),
      status: str(item.status), organizer: str(item.organizer), url: str(item.htmlLink), description: str(item.description),
      attendees: str(item.attendees),
      // A repeating event (or one occurrence of it): edits here could change the whole series, so the app sends these to Google Calendar.
      recurring: Array.isArray((item as Record<string, unknown>).recurrence) || Boolean(str((item as Record<string, unknown>).recurringEventId)),
    }];
  }).sort((a, b) => (a.start?.getTime() ?? 0) - (b.start?.getTime() ?? 0));
}

export function googleCalendarDayUrl(date: Date) {
  return `https://calendar.google.com/calendar/u/0/r/day/${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`;
}
