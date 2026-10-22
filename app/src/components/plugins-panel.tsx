import { useState, type ReactNode } from "react";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { toast } from "sonner";
import { getContext } from "@microsoft/managed-apps/app";
import type { LucideIcon } from "lucide-react";
import {
  Archive, ArrowDown, ArrowUp, Building2, CalendarPlus, ChevronLeft, ChevronRight, Copy, FilePlus2, FileArchive, House, MoreHorizontal, Pencil, Search, Trash2, Upload, CalendarDays, ChevronDown, CircleDot, ExternalLink, Eye, File, FileSpreadsheet, FileText, Flag, Folder, FolderGit2,
  GitBranch, GitFork, GitPullRequest, Globe, HardDrive, Link2, Loader2, Lock, MapPin, Presentation, RefreshCw, Scale, Star, UserRound,
} from "lucide-react";
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from "@/components/ui/collapsible";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { ConnectorErrorState } from "@/components/states";
import { hideEmails } from "@/lib/privacy";
import { NativeSelect, NativeSelectOption } from "@/components/ui/native-select";
import { safeUrl } from "@/lib/safe-url";
import { cn } from "@/lib/utils";
import { connectionsPageUrl } from "@/lib/connection-health";
import { PLUGIN_KEYS, type PluginKey } from "@/lib/dashboard-preferences";
import { PdfViewer } from "@/components/pdf-viewer";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Checkbox } from "@/components/ui/checkbox";
import { useConfirm } from "@/hooks/use-confirm";
import { InfoTip } from "@/components/info-tip";
import { recordHistory } from "@/lib/history";
import {
  createGoogleEvent, createTextFile, deleteGoogleEvent, deleteSourceFile, extractArchive, readUploadable, replaceTextFile, updateGoogleEvent, uploadTextFile, type EventDraft,
} from "@/lib/plugin-writes";
import { IMAGE_MAX_BYTES, THUMB_MAX_BYTES, editableText, imageDataUrl, loadFileContent, previewKindOf, textOf } from "@/lib/file-preview";
import {
  googleCalendarDayUrl, loadBoxFolder, loadBoxRoot, loadDriveFolder, loadDriveRoot, loadGitHubProfile, loadGitHubRepoDetail, loadGitHubRepos, loadGitHubWork,
  loadGoogleCalendars, loadGoogleEvents, type GitHubLabel, type GitHubProfile, type GitHubRepo, type GitHubWorkItem, type GoogleEvent,
} from "@/lib/plugin-sources";

const DRIVE_WINDOW = 10;

function when(value?: string) {
  if (!value) return undefined;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return undefined;
  const sameYear = date.getFullYear() === new Date().getFullYear();
  return date.toLocaleDateString([], { month: "short", day: "numeric", ...(sameYear ? {} : { year: "numeric" }) });
}
function size(bytes?: number) {
  if (!bytes) return undefined;
  const units = ["B", "KB", "MB", "GB"];
  let value = bytes, unit = 0;
  while (value >= 1024 && unit < units.length - 1) { value /= 1024; unit++; }
  return `${value >= 10 || unit === 0 ? Math.round(value) : value.toFixed(1)} ${units[unit]}`;
}

/** A link out to the source, opening in a new tab. Renders plain text when the URL isn't safe. */
function SourceLink({ href, children, className }: { href?: string; children: ReactNode; className?: string }) {
  const url = safeUrl(href);
  if (!url) return <span className={className}>{children}</span>;
  return <a href={url} target="_blank" rel="noopener noreferrer"
    className={cn("rounded-sm font-semibold text-primary underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-ring", className)}>
    {children}<span className="sr-only"> (opens in a new tab)</span>
  </a>;
}
function OpenButton({ href, label }: { href: string; label: string }) {
  return <Button asChild variant="outline" size="sm" className="rounded-full">
    <a href={href} target="_blank" rel="noopener noreferrer">{label}<ExternalLink aria-hidden="true" /><span className="sr-only"> (opens in a new tab)</span></a>
  </Button>;
}
function Loading({ label }: { label: string }) {
  return <p role="status" className="flex items-center gap-2 py-3 text-sm text-muted-foreground"><Loader2 aria-hidden="true" className="size-4 animate-spin" />{label}</p>;
}
function Fact({ label, value }: { label: string; value: ReactNode }) {
  return <div className="min-w-0 rounded-xl border border-border bg-card px-3 py-2">
    <dt className="text-xs font-medium text-muted-foreground">{label}</dt>
    <dd className="font-heading text-lg font-bold tabular-nums">{value}</dd>
  </div>;
}
function Pager({ page, hasMore, fetching, onPage, label }: { page: number; hasMore: boolean; fetching: boolean; onPage: (page: number) => void; label: string }) {
  if (page === 1 && !hasMore) return null;
  return <nav aria-label={`${label} pages`} className="mt-3 flex flex-wrap items-center gap-2">
    {page > 1 && <Button variant="outline" size="sm" className="rounded-full" disabled={fetching} onClick={() => onPage(page - 1)}>Previous page</Button>}
    <span className="text-xs text-muted-foreground">Page {page}</span>
    {hasMore && <Button variant="outline" size="sm" className="rounded-full" disabled={fetching} onClick={() => onPage(page + 1)}>{fetching ? "Loading…" : "Next page"}</Button>}
  </nav>;
}

interface PanelProps { open: boolean; onOpenChange: (open: boolean) => void; move?: ReactNode }

/** Which account a plugin is using, plus a link to check or switch it on the person's own Connections page. */
interface PanelAccount { name?: string; detail?: string; checking?: boolean }
function AccountLine({ title, account }: { title: string; account: PanelAccount }) {
  const context = useQuery({ queryKey: ["app-context"], queryFn: getContext, retry: false, staleTime: Infinity });
  const url = connectionsPageUrl(context.data?.app.environmentId);
  return <p className="flex flex-wrap items-center gap-x-2 gap-y-1 border-t border-border bg-muted/40 px-4 py-2 text-xs">
    <UserRound aria-hidden="true" className="size-3.5 shrink-0 text-muted-foreground" />
    {account.checking ? <span className="text-muted-foreground">Checking which {title} account is connected…</span>
      : account.name ? <span>Connected as <strong className="font-semibold break-all">{account.name}</strong>{account.detail ? <span className="text-muted-foreground"> ({account.detail})</span> : null}</span>
      : <span className="text-muted-foreground">{title} doesn't say which account is connected.</span>}
    <a href={url} target="_blank" rel="noopener noreferrer" className="rounded-sm font-semibold text-primary underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-ring">
      Check or change account<span className="sr-only"> for {title} (opens your Connections page in a new tab)</span>
    </a>
  </p>;
}

/** One collapsible plugin panel. */
function PluginPanel({ id, icon: Icon, title, summary, open, onOpenChange, sourceUrl, sourceLabel, account, move, children }: {
  id: string; icon: LucideIcon; title: string; summary?: string; open: boolean; onOpenChange: (open: boolean) => void;
  sourceUrl: string; sourceLabel: string; account: PanelAccount; move?: ReactNode; children: ReactNode;
}) {
  return <Collapsible open={open} onOpenChange={onOpenChange} className={cn("plugin-panel overflow-hidden rounded-2xl border border-border bg-card", `plugin-${id}`)}>
    <div className="flex flex-wrap items-center gap-3 px-4 py-3">
      <CollapsibleTrigger asChild>
        <button type="button" aria-controls={`plugin-${id}-content`}
          className="flex min-w-0 flex-1 items-center gap-3 rounded-xl py-1 text-left focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-ring">
          <span className="plugin-icon flex size-10 shrink-0 items-center justify-center rounded-xl" aria-hidden="true"><Icon className="size-5" /></span>
          <span className="min-w-0 flex-1">
            <span className="block font-heading text-lg font-bold">{title}</span>
            {summary && <span className="block truncate text-sm text-muted-foreground">{summary}</span>}
          </span>
          <ChevronDown aria-hidden="true" className={cn("size-5 shrink-0 transition-transform motion-reduce:transition-none", open && "rotate-180")} />
        </button>
      </CollapsibleTrigger>
      <OpenButton href={sourceUrl} label={sourceLabel} />
      {move}
    </div>
    <AccountLine title={title} account={account} />
    <CollapsibleContent id={`plugin-${id}-content`} className="border-t border-border px-4 py-4">{children}</CollapsibleContent>
  </Collapsible>;
}

/* ---------------- GitHub ---------------- */

function plural(n: number, one: string, many = `${one}s`) { return `${n.toLocaleString()} ${n === 1 ? one : many}`; }
function repoSize(kb?: number) { return kb ? size(kb * 1024) : undefined; }

/** GitHub label chip: the label's own colour as a small swatch, the name as text (colour is never the only cue). */
function LabelChip({ label }: { label: GitHubLabel }) {
  const colour = label.color && /^[0-9a-f]{6}$/i.test(label.color) ? `#${label.color}` : undefined;
  return <Badge variant="outline" className="h-auto gap-1.5 whitespace-normal">
    {colour && <span aria-hidden="true" className="size-2 shrink-0 rounded-full border border-border forced-colors:hidden" style={{ background: colour }} />}{label.name}
  </Badge>;
}

function GitHubProfileCard({ profile }: { profile: GitHubProfile }) {
  const meta = [
    profile.company && { icon: Building2, text: profile.company },
    profile.location && { icon: MapPin, text: profile.location },
  ].filter(Boolean) as { icon: LucideIcon; text: string }[];
  return <div className="space-y-3">
    <div className="space-y-1">
      <p className="font-heading text-lg font-bold">
        <SourceLink href={profile.url}>{profile.name ? `${profile.name} (${profile.login})` : profile.login}</SourceLink>
        {profile.accountType && profile.accountType !== "User" && <Badge variant="outline" className="ml-2 align-middle">{profile.accountType}</Badge>}
      </p>
      {profile.bio && <p className="text-sm break-words">{profile.bio}</p>}
      <p className="flex flex-wrap gap-x-4 gap-y-1 text-xs text-muted-foreground">
        {meta.map(({ icon: Icon, text }) => <span key={text} className="inline-flex items-center gap-1"><Icon aria-hidden="true" className="size-3.5" />{hideEmails(text)}</span>)}
        {profile.website && <span className="inline-flex items-center gap-1"><Link2 aria-hidden="true" className="size-3.5" /><SourceLink href={profile.website} className="font-medium">{profile.website.replace(/^https?:\/\//, "")}</SourceLink></span>}
        {when(profile.memberSince) && <span>Member since {when(profile.memberSince)}</span>}
        {profile.plan && <span>Plan: {profile.plan}</span>}
      </p>
    </div>
    <dl className="grid grid-cols-2 gap-2 sm:grid-cols-3 lg:grid-cols-6">
      {profile.publicRepos !== undefined && <Fact label="Public repositories" value={<SourceLink href={profile.url ? `${profile.url}?tab=repositories` : undefined}>{profile.publicRepos}</SourceLink>} />}
      {profile.privateRepos !== undefined && <Fact label="Private repositories" value={profile.privateRepos} />}
      {profile.followers !== undefined && <Fact label="Followers" value={<SourceLink href={profile.url ? `${profile.url}?tab=followers` : undefined}>{profile.followers}</SourceLink>} />}
      {profile.following !== undefined && <Fact label="Following" value={<SourceLink href={profile.url ? `${profile.url}?tab=following` : undefined}>{profile.following}</SourceLink>} />}
      {profile.publicGists !== undefined && <Fact label="Public gists" value={<SourceLink href={`https://gist.github.com/${encodeURIComponent(profile.login)}`}>{profile.publicGists}</SourceLink>} />}
      <Fact label="Stars given" value={<SourceLink href={profile.url ? `${profile.url}?tab=stars` : undefined}>View</SourceLink>} />
    </dl>
  </div>;
}

function WorkRow({ item, showRepo = true }: { item: GitHubWorkItem; showRepo?: boolean }) {
  const Icon = item.kind === "pull request" ? GitPullRequest : CircleDot;
  const details = [
    showRepo && item.repo && `${item.repo}${item.number ? ` #${item.number}` : ""}`, !showRepo && item.number && `#${item.number}`,
    item.author && `opened by ${item.author}`, when(item.createdAt) && `on ${when(item.createdAt)}`,
    when(item.updatedAt) && `updated ${when(item.updatedAt)}`, item.comments ? plural(item.comments, "comment") : undefined,
  ].filter(Boolean).join(" · ");
  return <li className="flex gap-3 py-3">
    <Icon aria-label={item.kind === "pull request" ? "Pull request" : "Issue"} className="mt-0.5 size-4 shrink-0 text-muted-foreground" />
    <div className="min-w-0 space-y-1">
      <div className="flex flex-wrap items-center gap-2">
        <SourceLink href={item.url} className="break-words">{item.title}</SourceLink>
        {item.isDraft && <Badge variant="outline">Draft</Badge>}
        {item.state && item.state !== "open" && <Badge variant="outline" className="capitalize">{item.state}</Badge>}
      </div>
      <p className="text-xs text-muted-foreground">{details}</p>
      {(item.assignees.length > 0 || item.milestone) && <p className="flex flex-wrap gap-x-4 text-xs text-muted-foreground">
        {item.assignees.length > 0 && <span>Assigned to {item.assignees.join(", ")}</span>}
        {item.milestone && <span className="inline-flex items-center gap-1"><Flag aria-hidden="true" className="size-3" />{item.milestone}</span>}
      </p>}
      {item.labels.length > 0 && <div className="flex flex-wrap gap-1" aria-label="Labels">{item.labels.map(label => <LabelChip key={label.name} label={label} />)}</div>}
    </div>
  </li>;
}

function RepoDetail({ repo }: { repo: GitHubRepo }) {
  const detail = useQuery({ queryKey: ["plugins", "github", "repo-detail", repo.fullName], queryFn: () => loadGitHubRepoDetail(repo.owner, repo.name) });
  if (detail.isPending) return <Loading label={`Loading ${repo.name}…`} />;
  if (detail.isError) return <ConnectorErrorState error={detail.error} connector="GitHub" className="py-4" action={<Button variant="outline" size="sm" onClick={() => void detail.refetch()}><RefreshCw aria-hidden="true" />Try again</Button>} />;
  const { pulls, issues, milestones, labelCount } = detail.data;
  const base = repo.url ?? `https://github.com/${repo.fullName}`;
  const block = (title: string, href: string, rows: GitHubWorkItem[], empty: string) => <section aria-label={`${repo.name} ${title}`}>
    <h5 className="flex items-center justify-between gap-2 text-sm font-semibold">{title}<SourceLink href={href} className="text-xs font-medium">View all</SourceLink></h5>
    {rows.length ? <ul className="divide-y divide-border">{rows.map(row => <WorkRow key={row.id} item={row} showRepo={false} />)}</ul>
      : <p className="py-2 text-xs text-muted-foreground">{empty}</p>}
  </section>;
  return <div className="space-y-4">
    <div className="grid gap-4 lg:grid-cols-2">
      {block("Open pull requests", `${base}/pulls`, pulls, "No open pull requests.")}
      {repo.hasIssues ? block("Open issues", `${base}/issues`, issues, "No open issues.") : <p className="text-xs text-muted-foreground">Issues are turned off for this repository.</p>}
    </div>
    <section aria-label={`${repo.name} milestones`}>
      <h5 className="flex items-center justify-between gap-2 text-sm font-semibold">Open milestones<SourceLink href={`${base}/milestones`} className="text-xs font-medium">View all</SourceLink></h5>
      {milestones.length ? <ul className="mt-1 space-y-2">{milestones.map(m => {
        const total = m.open + m.closed, pct = total ? Math.round((m.closed / total) * 100) : 0;
        return <li key={m.id} className="rounded-xl border border-border p-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <SourceLink href={m.url}>{m.title}</SourceLink>
            <span className="text-xs text-muted-foreground">{[when(m.dueOn) && `Due ${when(m.dueOn)}`, `${m.closed} closed, ${m.open} open`].filter(Boolean).join(" · ")}</span>
          </div>
          <div className="mt-2 h-1.5 overflow-hidden rounded-full bg-muted" role="progressbar" aria-label={`${m.title} progress`} aria-valuenow={pct} aria-valuemin={0} aria-valuemax={100}>
            <div className="h-full rounded-full bg-primary" style={{ width: `${pct}%` }} />
          </div>
          <p className="mt-1 text-xs text-muted-foreground">{pct}% complete</p>
        </li>;
      })}</ul> : <p className="py-2 text-xs text-muted-foreground">No open milestones.</p>}
    </section>
    <p className="flex flex-wrap gap-x-4 gap-y-1 text-xs">
      <SourceLink href={`${base}/commits`}>Commits</SourceLink>
      <SourceLink href={`${base}/branches`}>Branches</SourceLink>
      <SourceLink href={`${base}/releases`}>Releases</SourceLink>
      <SourceLink href={`${base}/actions`}>Actions</SourceLink>
      <SourceLink href={`${base}/labels`}>{labelCount !== undefined ? plural(labelCount, "label") : "Labels"}</SourceLink>
      <SourceLink href={`${base}/settings`}>Settings</SourceLink>
    </p>
  </div>;
}

function RepoRow({ repo }: { repo: GitHubRepo }) {
  const [expanded, setExpanded] = useState(false);
  const panelId = `github-repo-${repo.id}`;
  const base = repo.url ?? `https://github.com/${repo.fullName}`;
  return <li className="py-3">
    <div className="flex items-start gap-3">
      <div className="min-w-0 flex-1 space-y-1">
        <div className="flex flex-wrap items-center gap-2">
          <SourceLink href={repo.url} className="break-all">{repo.fullName}</SourceLink>
          <Badge variant="outline" className="gap-1">{repo.isPrivate ? <><Lock aria-hidden="true" />Private</> : <><Globe aria-hidden="true" />Public</>}</Badge>
          {repo.isFork && <Badge variant="outline">Fork</Badge>}
          {repo.isArchived && <Badge variant="outline">Archived</Badge>}
          {repo.isTemplate && <Badge variant="outline">Template</Badge>}
        </div>
        {repo.description && <p className="text-sm break-words">{repo.description}</p>}
        {repo.homepage && safeUrl(repo.homepage) && <p className="inline-flex items-center gap-1 text-xs"><Link2 aria-hidden="true" className="size-3.5 text-muted-foreground" /><SourceLink href={repo.homepage} className="break-all font-medium">{repo.homepage.replace(/^https?:\/\//, "")}</SourceLink></p>}
        {repo.topics.length > 0 && <div className="flex flex-wrap gap-1" aria-label="Topics">{repo.topics.map(topic => <SourceLink key={topic} href={`https://github.com/topics/${encodeURIComponent(topic)}`} className="rounded-full bg-accent px-2 py-0.5 text-xs font-medium text-accent-foreground">{topic}</SourceLink>)}</div>}
        <p className="flex flex-wrap gap-x-4 gap-y-1 text-xs text-muted-foreground">
          {repo.language && <span>{repo.language}</span>}
          <SourceLink href={`${base}/stargazers`} className="inline-flex items-center gap-1 font-normal text-muted-foreground"><Star aria-hidden="true" className="size-3.5" />{plural(repo.stars, "star")}</SourceLink>
          <SourceLink href={`${base}/forks`} className="inline-flex items-center gap-1 font-normal text-muted-foreground"><GitFork aria-hidden="true" className="size-3.5" />{plural(repo.forks, "fork")}</SourceLink>
          <SourceLink href={`${base}/watchers`} className="inline-flex items-center gap-1 font-normal text-muted-foreground"><Eye aria-hidden="true" className="size-3.5" />{plural(repo.watchers, "watcher")}</SourceLink>
          <span className="inline-flex items-center gap-1"><CircleDot aria-hidden="true" className="size-3.5" />{plural(repo.openIssues, "open issue or pull request", "open issues and pull requests")}</span>
          {repo.license && <span className="inline-flex items-center gap-1"><Scale aria-hidden="true" className="size-3.5" />{repo.license}</span>}
          {repo.defaultBranch && <span className="inline-flex items-center gap-1"><GitBranch aria-hidden="true" className="size-3.5" />{repo.defaultBranch}</span>}
          {repoSize(repo.sizeKb) && <span>{repoSize(repo.sizeKb)}</span>}
          {when(repo.createdAt) && <span>Created {when(repo.createdAt)}</span>}
          {when(repo.pushedAt) && <span>Last push {when(repo.pushedAt)}</span>}
        </p>
      </div>
      <Button variant="ghost" size="sm" className="shrink-0 rounded-full" aria-expanded={expanded} aria-controls={panelId} onClick={() => setExpanded(value => !value)}>
        {expanded ? "Hide details" : "Details"}<ChevronDown aria-hidden="true" className={cn("transition-transform motion-reduce:transition-none", expanded && "rotate-180")} />
      </Button>
    </div>
    {expanded && <div id={panelId} className="mt-3 rounded-xl border border-border bg-muted/30 p-3"><RepoDetail repo={repo} /></div>}
  </li>;
}

function GitHubWork({ login }: { login: string }) {
  const [page, setPage] = useState(1);
  const work = useQuery({ queryKey: ["plugins", "github", "work", login, page], queryFn: () => loadGitHubWork(login, page) });
  const pulls = work.data?.rows.filter(row => row.kind === "pull request").length ?? 0;
  const issues = (work.data?.rows.length ?? 0) - pulls;
  return <section aria-labelledby="github-work-heading">
    <h4 id="github-work-heading" className="flex flex-wrap items-center justify-between gap-2 font-heading text-base font-bold">
      <span>Open issues and pull requests involving you{work.data?.total !== undefined ? ` (${work.data.total})` : ""}</span>
      <span className="flex gap-3 text-xs font-medium">
        <SourceLink href={`https://github.com/pulls?q=${encodeURIComponent(`is:open is:pr involves:${login}`)}`}>All pull requests</SourceLink>
        <SourceLink href={`https://github.com/issues?q=${encodeURIComponent(`is:open is:issue involves:${login}`)}`}>All issues</SourceLink>
      </span>
    </h4>
    {work.data && work.data.rows.length > 0 && <p className="text-xs text-muted-foreground">On this page: {plural(pulls, "pull request")}, {plural(issues, "issue")}</p>}
    {work.isPending ? <Loading label="Loading issues and pull requests…" />
      : work.isError ? <ConnectorErrorState error={work.error} connector="GitHub" className="py-6" action={<Button variant="outline" size="sm" onClick={() => void work.refetch()}><RefreshCw aria-hidden="true" />Try again</Button>} />
      : work.data.rows.length === 0 && page === 1 ? <p className="mt-2 rounded-xl border border-dashed border-border p-3 text-sm text-muted-foreground">Nothing open involves you.</p>
      : <ul className="divide-y divide-border">{work.data.rows.map(item => <WorkRow key={item.id} item={item} />)}</ul>}
    <Pager page={page} hasMore={work.data?.hasMore ?? false} fetching={work.isFetching} onPage={setPage} label="Issues and pull requests" />
  </section>;
}

function GitHubRepos({ login }: { login: string }) {
  const [page, setPage] = useState(1);
  const repos = useQuery({ queryKey: ["plugins", "github", "repos", page], queryFn: () => loadGitHubRepos(page) });
  const rows = repos.data?.rows ?? [];
  const stars = rows.reduce((sum, repo) => sum + repo.stars, 0), forks = rows.reduce((sum, repo) => sum + repo.forks, 0);
  const languages = [...new Set(rows.map(repo => repo.language).filter(Boolean))];
  return <section aria-labelledby="github-repos-heading">
    <h4 id="github-repos-heading" className="flex flex-wrap items-center justify-between gap-2 font-heading text-base font-bold">
      <span className="flex items-center gap-1">Repositories, most recently updated
        <InfoTip label="GitHub data" title="What GitHub shows here" points={[
          "Repositories you own or collaborate on, 8 per page.",
          "Details load only when you open them: the 5 newest open pull requests and issues, and open milestones.",
          "Totals under the heading cover the current page only.",
          "Each page and each Details view counts toward your daily request allowance.",
        ]} />
      </span>
      <SourceLink href={`https://github.com/${encodeURIComponent(login)}?tab=repositories`} className="text-xs font-medium">All repositories</SourceLink>
    </h4>
    {rows.length > 0 && <p className="text-xs text-muted-foreground">
      On this page: {plural(stars, "star")}, {plural(forks, "fork")}{languages.length ? ` · ${languages.join(", ")}` : ""}
    </p>}
    {repos.isPending ? <Loading label="Loading repositories…" />
      : repos.isError ? <ConnectorErrorState error={repos.error} connector="GitHub" className="py-6" action={<Button variant="outline" size="sm" onClick={() => void repos.refetch()}><RefreshCw aria-hidden="true" />Try again</Button>} />
      : rows.length === 0 && page === 1 ? <p className="mt-2 rounded-xl border border-dashed border-border p-3 text-sm text-muted-foreground">No repositories yet.</p>
      : <ul className="divide-y divide-border">{rows.map(repo => <RepoRow key={repo.id} repo={repo} />)}</ul>}
    <Pager page={page} hasMore={repos.data?.hasMore ?? false} fetching={repos.isFetching} onPage={setPage} label="Repositories" />
  </section>;
}

function GitHubPanel({ open, onOpenChange, move }: PanelProps) {
  const profile = useQuery({ queryKey: ["plugins", "github", "profile"], queryFn: loadGitHubProfile });
  const login = profile.data?.login;
  const summary = profile.isPending ? "Loading…" : profile.isError ? "Couldn't load GitHub"
    : profile.data.publicRepos !== undefined ? `${plural(profile.data.publicRepos, "public repository", "public repositories")}` : "Your repositories, issues and pull requests";
  const account: PanelAccount = { checking: profile.isPending, name: login, detail: profile.data?.name };
  return <PluginPanel id="github" icon={FolderGit2} title="GitHub" summary={summary} open={open} onOpenChange={onOpenChange} account={account} move={move}
    sourceUrl={safeUrl(profile.data?.url) ?? "https://github.com"} sourceLabel="Open GitHub">
    {profile.isPending ? <Loading label="Loading your GitHub profile…" />
      : profile.isError ? <ConnectorErrorState error={profile.error} connector="GitHub" className="py-6" action={<Button variant="outline" size="sm" onClick={() => void profile.refetch()}><RefreshCw aria-hidden="true" />Try again</Button>} />
      : <div className="space-y-6">
        <GitHubProfileCard profile={profile.data} />
        <div className="grid gap-6 2xl:grid-cols-2">
          <GitHubRepos login={profile.data.login} />
          <GitHubWork login={profile.data.login} />
        </div>
        <nav aria-label="GitHub shortcuts" className="flex flex-wrap gap-2">
          <OpenButton href="https://github.com/notifications" label="Notifications" />
          <OpenButton href="https://github.com/pulls/review-requested" label="Review requests" />
          <OpenButton href={`https://github.com/${encodeURIComponent(profile.data.login)}?tab=stars`} label="Starred" />
          <OpenButton href="https://github.com/new" label="New repository" />
        </nav>
      </div>}
  </PluginPanel>;
}

/* ---------------- Files (Google Drive and Box) ---------------- */

interface FileEntry { id: string; name: string; path?: string; isFolder: boolean; kind: string; mediaType?: string; modified?: string; modifiedBy?: string; size?: number; url: string }
interface FileSource { key: "googledrive" | "box"; label: string; home: string; loadFolder: (id: string) => Promise<FileEntry[]> }

function fileIcon(item: FileEntry): LucideIcon {
  if (item.isFolder) return Folder;
  if (/Sheet|Excel/.test(item.kind)) return FileSpreadsheet;
  if (/Slides|PowerPoint/.test(item.kind)) return Presentation;
  if (/Doc|Word|PDF|Text/.test(item.kind)) return FileText;
  return File;
}

/** Picks a file and uploads it to a Google Drive or Box folder (text-based files only — see plugin-writes). */
function UploadButton({ source, folderPath, folderLabel }: { source: FileSource; folderPath: string; folderLabel: string }) {
  const queryClient = useQueryClient();
  const [busy, setBusy] = useState(false);
  const inputId = `upload-${source.key}-${folderPath.replace(/[^A-Za-z0-9_-]/g, "") || "root"}`;
  async function picked(file: File | undefined, input: HTMLInputElement) {
    input.value = "";
    if (!file) return;
    const check = await readUploadable(file);
    if (!check.ok) { toast.error(check.reason, { action: { label: `Open ${source.label}`, onClick: () => window.open(source.home, "_blank", "noopener") } }); return; }
    setBusy(true);
    try {
      await uploadTextFile(source.key, folderPath, file.name, check.text);
      toast.success(`Uploaded ${file.name} to ${folderLabel} in ${source.label}.`);
      recordHistory({ kind: "created", title: `Uploaded "${file.name}" to ${source.label}`, section: "plugins", status: "done", cost: "free" });
      await queryClient.invalidateQueries({ queryKey: ["plugins", source.key] });
    } catch {
      toast.error(`Couldn't upload ${file.name} to ${source.label}. A file with that name may already exist, or uploads may be turned off by your organization.`);
      recordHistory({ kind: "created", title: `Upload of "${file.name}" to ${source.label}`, section: "plugins", status: "failed", cost: "free" });
    } finally { setBusy(false); }
  }
  return <>
    <input id={inputId} type="file" className="sr-only" tabIndex={-1} aria-hidden="true"
      accept=".txt,.csv,.tsv,.md,.markdown,.json,.xml,.html,.htm,.css,.js,.ts,.yaml,.yml,.ini,.cfg,.log,.svg,text/*"
      onChange={event => void picked(event.target.files?.[0], event.target)} />
    <Button type="button" variant="outline" size="sm" className="rounded-full" disabled={busy} onClick={() => document.getElementById(inputId)?.click()}>
      {busy ? <Loader2 aria-hidden="true" className="animate-spin" /> : <Upload aria-hidden="true" />}{busy ? "Uploading…" : `Upload${folderPath === "/" ? "" : " here"}`}
      <span className="sr-only"> a text file to {folderLabel} in {source.label}</span>
    </Button>
    <InfoTip label={`uploading to ${source.label}`} title="Uploading from the dashboard" points={[
      "Text-based files only: TXT, CSV, Markdown, JSON, HTML, XML and similar.",
      `Images, PDFs and Word, Excel or PowerPoint files can't be uploaded here — they would arrive damaged. Upload them in ${source.label} instead.`,
      "Up to 4 MB per file.",
      source.key === "box" ? "Box won't accept a file whose name is already in the folder." : "Google Drive allows two files with the same name, so an upload never replaces an existing file — use Replace for that.",
      "Replace (in a file's … menu) works the same way and only for text files.",
      "Your organization can turn uploading off; you'll see an error if it has.",
      "Each upload counts toward your daily request allowance.",
    ]} />
  </>;
}

/** Write a new text file, or edit an existing one, and save it to Google Drive or Box. */
function TextFileDialog({ source, folderPath, folderLabel, file, open, onOpenChange }: {
  source: FileSource; folderPath: string; folderLabel: string; file?: FileEntry; open: boolean; onOpenChange: (open: boolean) => void;
}) {
  const queryClient = useQueryClient();
  const editing = Boolean(file);
  const existing = useQuery({ queryKey: ["plugins", source.key, "content", file?.id], queryFn: () => loadFileContent(source.key, file!.id), enabled: editing && open });
  const loadedText = editing && existing.data !== undefined ? editableText(existing.data) : undefined;
  const [name, setName] = useState(() => file?.name ?? `Note ${new Date().toLocaleDateString("en-CA")}.md`);
  const [text, setText] = useState<string>();
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string>();
  const body = text ?? loadedText ?? "";
  async function save() {
    const cleanName = name.trim();
    if (!editing && (!cleanName || /[\\/:*?"<>|]/.test(cleanName))) { setError("Give the file a name without \\ / : * ? \" < > or |."); return; }
    setSaving(true); setError(undefined);
    try {
      if (file) await replaceTextFile(source.key, file.id, body); else await createTextFile(source.key, folderPath, cleanName, body);
      toast.success(editing ? `Saved ${file!.name}.` : `Created ${cleanName} in ${folderLabel}.`);
      recordHistory({ kind: editing ? "action" : "created", title: `${editing ? "Edited" : "Created"} "${file?.name ?? cleanName}" in ${source.label}`, section: "plugins", status: "done", cost: "free" });
      await queryClient.invalidateQueries({ queryKey: ["plugins", source.key] });
      onOpenChange(false);
    } catch {
      setError(editing ? "Couldn't save the changes. Editing may be turned off by your organization." : `Couldn't create the file. ${source.key === "box" ? "A file with that name may already exist, or " : ""}creating files may be turned off by your organization.`);
    } finally { setSaving(false); }
  }
  return <Dialog open={open} onOpenChange={onOpenChange}>
    <DialogContent className="sm:max-w-2xl">
      <DialogHeader>
        <DialogTitle>{editing ? `Edit ${file!.name}` : "New file"}</DialogTitle>
        <DialogDescription>{editing ? `Changes are saved to ${source.label}.` : `Saved as a text file in ${folderLabel} in ${source.label}.`}</DialogDescription>
      </DialogHeader>
      <form className="space-y-3" onSubmit={e => { e.preventDefault(); void save(); }}>
        {!editing && <div className="space-y-1"><Label htmlFor={`${source.key}-new-name`}>File name</Label><Input id={`${source.key}-new-name`} value={name} onChange={e => setName(e.target.value)} required /></div>}
        {editing && existing.isPending ? <Loading label="Loading the file…" />
          : editing && (existing.isError || loadedText === null) ? <p role="alert" className="text-sm text-red-700 dark:text-red-400">This file can't be edited here — it couldn't be read as plain text, or it's over 1 MB. <SourceLink href={file!.url}>Open in {source.label}</SourceLink></p>
          : <div className="space-y-1"><Label htmlFor={`${source.key}-text`}>Contents</Label>
            <Textarea id={`${source.key}-text`} rows={14} className="font-mono text-xs" value={body} onChange={e => setText(e.target.value)} /></div>}
        {error && <p role="alert" className="text-sm font-medium text-red-700 dark:text-red-400">{error}</p>}
        <DialogFooter>
          <Button type="button" variant="outline" onClick={() => onOpenChange(false)}>Cancel</Button>
          <Button type="submit" disabled={saving || (editing && (existing.isPending || loadedText == null))}>{saving && <Loader2 aria-hidden="true" className="animate-spin" />}{editing ? "Save" : "Create file"}</Button>
        </DialogFooter>
      </form>
    </DialogContent>
  </Dialog>;
}

function NewFileButton({ source, folderPath, folderLabel }: { source: FileSource; folderPath: string; folderLabel: string }) {
  const [open, setOpen] = useState(false);
  return <>
    <Button type="button" variant="outline" size="sm" className="rounded-full" onClick={() => setOpen(true)}><FilePlus2 aria-hidden="true" />New file<span className="sr-only"> in {folderLabel}</span></Button>
    {open && <TextFileDialog source={source} folderPath={folderPath} folderLabel={folderLabel} open={open} onOpenChange={setOpen} />}
  </>;
}

/** Edit / Replace (text files), Extract here (.zip) and Delete for one Google Drive or Box file. */
function FileActions({ item, source, folderPath, folderLabel }: { item: FileEntry; source: FileSource; folderPath: string; folderLabel: string }) {
  const queryClient = useQueryClient();
  const confirm = useConfirm();
  const [editing, setEditing] = useState(false);
  const replaceId = `replace-${source.key}-${item.id.replace(/[^A-Za-z0-9_-]/g, "")}`;
  const isText = previewKindOf(item) === "text";
  const isZip = /\.zip$/i.test(item.name) && Boolean(item.path);
  async function remove() {
    const ok = await confirm({ title: `Delete ${item.name}?`, description: `This deletes the file from ${source.label}.`, confirmLabel: "Delete", destructive: true });
    if (!ok) return;
    try {
      await deleteSourceFile(source.key, item.id);
      toast.success(`Deleted ${item.name}.`);
      recordHistory({ kind: "action", title: `Deleted "${item.name}" from ${source.label}`, section: "plugins", status: "done", cost: "free" });
      await queryClient.invalidateQueries({ queryKey: ["plugins", source.key] });
    } catch { toast.error(`Couldn't delete ${item.name}. Deleting may be turned off by your organization, or you may not have permission.`); }
  }
  async function replace(file: File | undefined, input: HTMLInputElement) {
    input.value = "";
    if (!file) return;
    const check = await readUploadable(file);
    if (!check.ok) { toast.error(check.reason); return; }
    const ok = await confirm({ title: `Replace ${item.name}?`, description: `Its contents will be replaced with ${file.name}. The file keeps its name.`, confirmLabel: "Replace" });
    if (!ok) return;
    try {
      await replaceTextFile(source.key, item.id, check.text);
      toast.success(`Replaced the contents of ${item.name}.`);
      recordHistory({ kind: "action", title: `Replaced "${item.name}" in ${source.label}`, section: "plugins", status: "done", cost: "free" });
      await queryClient.invalidateQueries({ queryKey: ["plugins", source.key] });
    } catch { toast.error(`Couldn't replace ${item.name}.`); }
  }
  async function extract() {
    const ok = await confirm({ title: `Extract ${item.name}?`, description: `Its contents are unpacked into ${folderLabel}. Files with the same names aren't overwritten.`, confirmLabel: "Extract" });
    if (!ok) return;
    try {
      const count = await extractArchive(source.key, item.path!, folderPath);
      toast.success(`Extracted ${item.name}${count !== undefined ? ` (${count} item${count === 1 ? "" : "s"})` : ""}.`);
      recordHistory({ kind: "action", title: `Extracted "${item.name}" in ${source.label}`, section: "plugins", status: "done", cost: "free" });
      await queryClient.invalidateQueries({ queryKey: ["plugins", source.key] });
    } catch { toast.error(`Couldn't extract ${item.name}. Some of its files may already exist there.`); }
  }
  return <>
    {isText && <input id={replaceId} type="file" className="sr-only" tabIndex={-1} aria-hidden="true" onChange={event => void replace(event.target.files?.[0], event.target)} />}
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <Button variant="ghost" size="icon" className="size-8 shrink-0 rounded-full" aria-label={`More actions for ${item.name}`}><MoreHorizontal aria-hidden="true" /></Button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="end">
        {isText && <DropdownMenuItem onSelect={() => setEditing(true)}><Pencil aria-hidden="true" />Edit</DropdownMenuItem>}
        {isText && <DropdownMenuItem onSelect={() => document.getElementById(replaceId)?.click()}><Upload aria-hidden="true" />Replace with a file…</DropdownMenuItem>}
        {isZip && <DropdownMenuItem onSelect={() => void extract()}><FileArchive aria-hidden="true" />Extract here</DropdownMenuItem>}
        <DropdownMenuItem variant="destructive" onSelect={() => void remove()}><Trash2 aria-hidden="true" />Delete</DropdownMenuItem>
      </DropdownMenuContent>
    </DropdownMenu>
    {editing && <TextFileDialog source={source} folderPath={folderPath} folderLabel={folderLabel} file={item} open={editing} onOpenChange={setEditing} />}
  </>;
}

/** Small image thumbnail, loaded only for images under the thumbnail size limit; falls back to the file icon. */
function Thumbnail({ item, source, icon: Icon }: { item: FileEntry; source: FileSource; icon: LucideIcon }) {
  const eligible = previewKindOf(item) === "image" && (item.size === undefined || item.size <= THUMB_MAX_BYTES);
  const content = useQuery({ queryKey: ["plugins", source.key, "content", item.id], queryFn: () => loadFileContent(source.key, item.id), enabled: eligible });
  const src = eligible && content.data !== undefined ? imageDataUrl(content.data, THUMB_MAX_BYTES) : null;
  return <span className="flex size-10 shrink-0 items-center justify-center overflow-hidden rounded-lg border border-border bg-muted/50" aria-hidden="true">
    {src ? <img src={src} alt="" className="size-full object-cover" /> : <Icon className={cn("size-4 text-muted-foreground", eligible && content.isPending && "animate-pulse")} />}
  </span>;
}

function FilePreview({ item, source }: { item: FileEntry; source: FileSource }) {
  const kind = previewKindOf(item);
  const tooBig = kind === "image" && item.size !== undefined && item.size > IMAGE_MAX_BYTES;
  const content = useQuery({ queryKey: ["plugins", source.key, "content", item.id], queryFn: () => loadFileContent(source.key, item.id), enabled: kind !== "pdf" && !tooBig });
  const openLink = <SourceLink href={item.url} className="text-sm">Open in {source.label}</SourceLink>;
  if (kind === "pdf") {
    return <PdfViewer key={item.id} title={item.name} loadPdf={async () => {
      try { return { success: true, data: await loadFileContent(source.key, item.id) }; } catch { return { success: false, data: null }; }
    }} />;
  }
  if (tooBig) return <p className="text-sm text-muted-foreground">This image is too large to preview here. {openLink}</p>;
  if (content.isPending) return <Loading label={`Loading a preview of ${item.name}…`} />;
  if (content.isError) return <ConnectorErrorState error={content.error} connector={source.label} className="py-4" action={<Button variant="outline" size="sm" onClick={() => void content.refetch()}><RefreshCw aria-hidden="true" />Try again</Button>} />;
  if (kind === "image") {
    const src = imageDataUrl(content.data);
    return src ? <figure className="space-y-2">
      <img src={src} alt={`Preview of ${item.name}`} className="max-h-[28rem] w-auto max-w-full rounded-lg border border-border object-contain" />
      <figcaption className="text-xs text-muted-foreground">{item.name}</figcaption>
    </figure> : <p className="text-sm text-muted-foreground">This image can't be previewed here. {openLink}</p>;
  }
  const text = textOf(content.data);
  return text ? <div className="space-y-2">
    <pre className="max-h-80 overflow-auto whitespace-pre-wrap break-words rounded-lg border border-border bg-muted/40 p-3 text-xs" tabIndex={0} aria-label={`Contents of ${item.name}`}>{text.text}</pre>
    {text.truncated && <p className="text-xs text-muted-foreground">Showing the beginning of the file. {openLink}</p>}
  </div> : <p className="text-sm text-muted-foreground">This file can't be previewed here. {openLink}</p>;
}

interface Place { id?: string; name: string; path: string }

/**
 * One file or folder. Selecting it keeps you in the dashboard: a folder opens here, a file opens its preview
 * (or its details when it can't be shown). The small icon at the right is the secondary way out to the source.
 */
function FileRow({ item, source, place, onOpenFolder }: { item: FileEntry; source: FileSource; place: Place; onOpenFolder: (item: FileEntry) => void }) {
  const Icon = fileIcon(item);
  const details = [item.kind, when(item.modified) && `modified ${when(item.modified)}`, item.modifiedBy && `by ${hideEmails(item.modifiedBy)}`, size(item.size)].filter(Boolean).join(" · ");
  const previewId = `${source.key}-preview-${item.id.replace(/[^A-Za-z0-9_-]/g, "")}`;
  const kind = previewKindOf(item);
  const [open, setOpen] = useState(false);
  const external = safeUrl(item.url);
  return <li className="py-1.5">
    <div className={cn("flex items-center gap-2 rounded-xl pr-1", open && "bg-muted/50")}>
      <button type="button"
        className="flex min-w-0 flex-1 items-center gap-3 rounded-xl p-1.5 text-left hover:bg-muted focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-ring"
        onClick={() => item.isFolder ? onOpenFolder(item) : setOpen(value => !value)}
        aria-expanded={item.isFolder ? undefined : open} aria-controls={item.isFolder ? undefined : previewId}>
        <Thumbnail item={item} source={source} icon={Icon} />
        <span className="min-w-0 flex-1">
          <span className="block break-words text-sm font-semibold text-primary">{item.name}</span>
          <span className="block text-xs text-muted-foreground">{details}</span>
        </span>
        <span className="sr-only">{item.isFolder ? " — open this folder here" : open ? " — hide preview" : kind ? " — preview here" : " — show details"}</span>
        {item.isFolder ? <ChevronRight aria-hidden="true" className="size-4 shrink-0 text-muted-foreground" />
          : <ChevronDown aria-hidden="true" className={cn("size-4 shrink-0 text-muted-foreground transition-transform motion-reduce:transition-none", open && "rotate-180")} />}
      </button>
      {external && <Button asChild variant="ghost" size="icon" className="size-8 shrink-0 rounded-full">
        <a href={external} target="_blank" rel="noopener noreferrer" aria-label={`Open ${item.name} in ${source.label} (new tab)`} title={`Open in ${source.label}`}><ExternalLink aria-hidden="true" /></a>
      </Button>}
      {!item.isFolder && <FileActions item={item} source={source} folderPath={place.path} folderLabel={place.name} />}
    </div>
    {!item.isFolder && open && <div id={previewId} className="mt-2 rounded-xl border border-border bg-card p-3">
      {kind ? <FilePreview item={item} source={source} />
        : <div className="space-y-2 text-sm">
          <p className="font-semibold">{item.name}</p>
          <p className="text-muted-foreground">{details}</p>
          <p className="text-muted-foreground">{item.kind === "Folder" ? "" : `${item.kind}s can't be shown inside the dashboard.`}</p>
          {external && <Button asChild size="sm" className="rounded-full"><a href={external} target="_blank" rel="noopener noreferrer">Open in {source.label}<ExternalLink aria-hidden="true" /><span className="sr-only"> (new tab)</span></a></Button>}
        </div>}
    </div>}
  </li>;
}

function FileList({ items, source, place, onOpenFolder }: { items: FileEntry[]; source: FileSource; place: Place; onOpenFolder: (item: FileEntry) => void }) {
  const [shown, setShown] = useState(DRIVE_WINDOW);
  return <>
    <ul className="divide-y divide-border">{items.slice(0, shown).map(item => <FileRow key={item.id} item={item} source={source} place={place} onOpenFolder={onOpenFolder} />)}</ul>
    {items.length > shown && <Button variant="outline" size="sm" className="mt-2 rounded-full" onClick={() => setShown(value => value + DRIVE_WINDOW)}>
      Show {Math.min(DRIVE_WINDOW, items.length - shown)} more of {items.length - shown}
    </Button>}
  </>;
}

function FilesPanel({ id, icon, title, home, homeLabel, rootName, emptyText, loadRoot, source, open, onOpenChange, move }: {
  id: "googledrive" | "box"; icon: LucideIcon; title: string; home: string; homeLabel: string; rootName: string; emptyText: string;
  loadRoot: () => Promise<FileEntry[]>; source: FileSource; open: boolean; onOpenChange: (open: boolean) => void; move?: ReactNode;
}) {
  // Where you are: the top level, then each folder you opened (a breadcrumb trail).
  const [trail, setTrail] = useState<Place[]>([]);
  const here: Place = trail[trail.length - 1] ?? { name: rootName, path: "/" };
  const root = useQuery({ queryKey: ["plugins", id, "root"], queryFn: loadRoot });
  const folder = useQuery({ queryKey: ["plugins", id, "folder", here.id], queryFn: () => source.loadFolder(here.id!), enabled: Boolean(here.id) });
  const list = here.id ? folder : root;
  const folders = list.data?.filter(item => item.isFolder).length ?? 0;
  const files = (list.data?.length ?? 0) - folders;
  const latest = list.data?.[0];
  const summary = root.isPending ? "Loading…" : root.isError ? `Couldn't load ${title}`
    : `${(root.data.length - root.data.filter(i => i.isFolder).length)} files and ${root.data.filter(i => i.isFolder).length} folders at the top level`;
  const openFolder = (item: FileEntry) => setTrail(current => [...current, { id: item.id, name: item.name, path: item.path ?? `${here.path.replace(/\/$/, "")}/${item.name}` }]);
  return <PluginPanel id={id} icon={icon} title={title} summary={summary} open={open} onOpenChange={onOpenChange} sourceUrl={home} sourceLabel={homeLabel} account={{}} move={move}>
    <div className="space-y-4">
      <nav aria-label={`${title} folders`} className="flex flex-wrap items-center gap-1 text-sm">
        <Button variant="ghost" size="sm" className="h-7 rounded-full px-2" aria-current={trail.length === 0 ? "page" : undefined} disabled={trail.length === 0} onClick={() => setTrail([])}><House aria-hidden="true" />{rootName}</Button>
        {trail.map((place, index) => <span key={place.id} className="flex items-center gap-1">
          <ChevronRight aria-hidden="true" className="size-3.5 text-muted-foreground" />
          <Button variant="ghost" size="sm" className="h-7 rounded-full px-2" aria-current={index === trail.length - 1 ? "page" : undefined} disabled={index === trail.length - 1} onClick={() => setTrail(trail.slice(0, index + 1))}>{place.name}</Button>
        </span>)}
      </nav>
      <div className="flex flex-wrap items-center justify-between gap-2">
        <h4 id={`${id}-items-heading`} className="flex items-center gap-1 font-heading text-base font-bold">{here.name}, most recently changed
          <InfoTip label={`${title} previews`} title="Thumbnails, previews and editing" points={[
            "Select a folder to open it here, or a file to preview it here. The arrow icon on the right opens it in the service instead.",
            "Thumbnails show for images (PNG, JPG, GIF, WebP, BMP) up to 3 MB.",
            "Preview shows images up to 15 MB, PDFs up to 25 MB in the built-in viewer, and the beginning of text files.",
            "Edit (in a file's … menu) works for text files up to 1 MB.",
            "Extract here unpacks a .zip into the folder you're in; files that already exist there aren't overwritten.",
            `Word, Excel, PowerPoint and Google Docs, Sheets or Slides can't be previewed or edited here — open them in ${title}.`,
            ...(id === "box" ? ["If Box doesn't share a file's own address, its link opens a Box search for the file's name."] : []),
            "Each thumbnail, preview and change counts toward your daily request allowance.",
          ]} />
        </h4>
        <div className="flex flex-wrap items-center gap-1">
          <NewFileButton source={source} folderPath={here.path} folderLabel={here.name} />
          <UploadButton source={source} folderPath={here.path} folderLabel={here.name} />
        </div>
      </div>
      {list.isPending ? <Loading label={`Loading ${here.name}…`} />
        : list.isError ? <ConnectorErrorState error={list.error} connector={title} className="py-6" action={<Button variant="outline" size="sm" onClick={() => void list.refetch()}><RefreshCw aria-hidden="true" />Try again</Button>} />
        : list.data.length === 0 ? <p className="rounded-xl border border-dashed border-border p-3 text-sm text-muted-foreground">{here.id ? "This folder is empty." : emptyText}</p>
        : <>
          <dl className="grid grid-cols-2 gap-2 sm:grid-cols-3">
            <Fact label="Files" value={files} />
            <Fact label="Folders" value={folders} />
            {latest && <Fact label="Last change" value={when(latest.modified) ?? "—"} />}
          </dl>
          <section aria-labelledby={`${id}-items-heading`}>
            <FileList key={here.id ?? "root"} items={list.data} source={source} place={here} onOpenFolder={openFolder} />
          </section>
        </>}
    </div>
  </PluginPanel>;
}

const DRIVE_SOURCE: FileSource = { key: "googledrive", label: "Google Drive", home: "https://drive.google.com/drive/my-drive", loadFolder: loadDriveFolder };
const BOX_SOURCE: FileSource = { key: "box", label: "Box", home: "https://app.box.com/folder/0", loadFolder: loadBoxFolder };

function DrivePanel(props: PanelProps) {
  return <FilesPanel id="googledrive" icon={HardDrive} title="Google Drive" home="https://drive.google.com/drive/my-drive" homeLabel="Open Google Drive"
    rootName="My Drive" emptyText="Your Drive is empty." loadRoot={loadDriveRoot} source={DRIVE_SOURCE} {...props} />;
}
function BoxPanel(props: PanelProps) {
  return <FilesPanel id="box" icon={Archive} title="Box" home="https://app.box.com/folder/0" homeLabel="Open Box"
    rootName="All Files" emptyText="Your Box is empty." loadRoot={loadBoxRoot} source={BOX_SOURCE} {...props} />;
}

/* ---------------- Google Calendar ---------------- */

const CAL_DAYS = 7;
function startOfToday() { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }
function dayLabel(date: Date) {
  const today = startOfToday(), tomorrow = new Date(today); tomorrow.setDate(today.getDate() + 1);
  const key = date.toDateString();
  if (key === today.toDateString()) return "Today";
  if (key === tomorrow.toDateString()) return "Tomorrow";
  return date.toLocaleDateString([], { weekday: "long", month: "short", day: "numeric" });
}
function timeRange(event: GoogleEvent) {
  if (event.allDay) return "All day";
  if (!event.start) return "Time unavailable";
  const t = (d: Date) => d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  return event.end ? `${t(event.start)} – ${t(event.end)}` : t(event.start);
}

/* Event form helpers: local values for <input type="date"> / <input type="datetime-local">. */
const pad = (n: number) => String(n).padStart(2, "0");
const localDate = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const localDateTime = (d: Date) => `${localDate(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
function draftFor(event?: GoogleEvent): EventDraft {
  if (event?.start) {
    const end = event.end ?? new Date(event.start.getTime() + 3600000);
    if (event.allDay) { const last = new Date(end); last.setDate(last.getDate() - 1); return { title: event.title, allDay: true, start: localDate(event.start), end: localDate(last < event.start ? event.start : last), location: event.location ?? "", description: event.description ?? "", attendees: event.attendees ?? "" }; }
    return { title: event.title, allDay: false, start: localDateTime(event.start), end: localDateTime(end), location: event.location ?? "", description: event.description ?? "", attendees: event.attendees ?? "" };
  }
  const start = new Date(); start.setMinutes(0, 0, 0); start.setHours(start.getHours() + 1);
  return { title: "", allDay: false, start: localDateTime(start), end: localDateTime(new Date(start.getTime() + 3600000)), location: "", description: "", attendees: "" };
}

/** Add or edit a Google Calendar event. */
export function EventDialog({ calendarId, event, initial, open, onOpenChange, heading }: {
  calendarId: string; event?: GoogleEvent; initial?: EventDraft; open: boolean; onOpenChange: (open: boolean) => void; heading?: string;
}) {
  const queryClient = useQueryClient();
  const [draft, setDraft] = useState<EventDraft>(() => initial ?? draftFor(event));
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string>();
  const set = <K extends keyof EventDraft>(key: K, value: EventDraft[K]) => setDraft(current => ({ ...current, [key]: value }));
  const editing = Boolean(event);
  const invalid = !draft.title.trim() ? "Add a title." : !draft.start || !draft.end ? "Choose when it starts and ends."
    : (draft.allDay ? draft.end < draft.start : new Date(draft.end) <= new Date(draft.start)) ? "The end must be after the start." : undefined;
  function toggleAllDay(allDay: boolean) {
    setDraft(current => allDay
      ? { ...current, allDay, start: current.start.slice(0, 10), end: current.end.slice(0, 10) }
      : { ...current, allDay, start: `${current.start.slice(0, 10)}T09:00`, end: `${current.end.slice(0, 10)}T10:00` });
  }
  async function save() {
    if (invalid) { setError(invalid); return; }
    setSaving(true); setError(undefined);
    try {
      if (event) await updateGoogleEvent(calendarId, event.id, draft); else await createGoogleEvent(calendarId, draft);
      toast.success(editing ? `Updated "${draft.title.trim()}".` : `Added "${draft.title.trim()}" to Google Calendar.`);
      recordHistory({ kind: editing ? "action" : "created", title: `${editing ? "Edited" : "Added"} Google Calendar event "${draft.title.trim()}"`, section: "plugins", status: "done", cost: "free" });
      await queryClient.invalidateQueries({ queryKey: ["plugins", "googlecalendar", "events"] });
      onOpenChange(false);
    } catch {
      setError(`Couldn't ${editing ? "save the changes" : "add the event"}. Adding or editing events may be turned off by your organization — try again, or use Google Calendar.`);
    } finally { setSaving(false); }
  }
  const fieldType = draft.allDay ? "date" : "datetime-local";
  return <Dialog open={open} onOpenChange={onOpenChange}>
    <DialogContent className="sm:max-w-lg">
      <DialogHeader>
        <DialogTitle>{heading ?? (editing ? "Edit event" : "Add event")}</DialogTitle>
        <DialogDescription>{editing ? "Changes are saved to Google Calendar." : "The event is added to the calendar selected in the Google Calendar panel."}</DialogDescription>
      </DialogHeader>
      <form className="space-y-3" onSubmit={e => { e.preventDefault(); void save(); }}>
        <div className="space-y-1"><Label htmlFor="gcal-title">Title</Label><Input id="gcal-title" value={draft.title} onChange={e => set("title", e.target.value)} required /></div>
        <div className="flex items-center gap-2"><Checkbox id="gcal-allday" checked={draft.allDay} onCheckedChange={value => toggleAllDay(value === true)} /><Label htmlFor="gcal-allday">All day</Label></div>
        <div className="grid gap-3 sm:grid-cols-2">
          <div className="space-y-1"><Label htmlFor="gcal-start">Starts</Label><Input id="gcal-start" type={fieldType} value={draft.start} onChange={e => set("start", e.target.value)} required /></div>
          <div className="space-y-1"><Label htmlFor="gcal-end">Ends</Label><Input id="gcal-end" type={fieldType} value={draft.end} onChange={e => set("end", e.target.value)} required /></div>
        </div>
        <div className="space-y-1"><Label htmlFor="gcal-location">Location</Label><Input id="gcal-location" value={draft.location} onChange={e => set("location", e.target.value)} /></div>
        <div className="space-y-1"><Label htmlFor="gcal-description">Description</Label><Textarea id="gcal-description" rows={3} value={draft.description} onChange={e => set("description", e.target.value)} /></div>
        <div className="space-y-1">
          <Label htmlFor="gcal-attendees">Guests</Label>
          <Input id="gcal-attendees" value={draft.attendees} onChange={e => set("attendees", e.target.value)} placeholder="name@example.com, name2@example.com" aria-describedby="gcal-attendees-help" />
          <p id="gcal-attendees-help" className="text-xs text-muted-foreground">Email addresses separated by commas. Google may email them an invitation.</p>
        </div>
        {error && <p role="alert" className="text-sm font-medium text-red-700 dark:text-red-400">{error}</p>}
        <DialogFooter>
          <Button type="button" variant="outline" onClick={() => onOpenChange(false)}>Cancel</Button>
          <Button type="submit" disabled={saving}>{saving && <Loader2 aria-hidden="true" className="animate-spin" />}{editing ? "Save changes" : "Add event"}</Button>
        </DialogFooter>
      </form>
    </DialogContent>
  </Dialog>;
}

/** Plain text from Google's event description, which may contain simple HTML. */
function plainText(value?: string) {
  if (!value) return "";
  const doc = new DOMParser().parseFromString(value.replace(/<br\s*\/?>/gi, "\n").replace(/<\/p>/gi, "\n"), "text/html");
  return (doc.body.textContent ?? "").replace(/\n{3,}/g, "\n\n").trim();
}

/** Expandable details for one Google Calendar event: when, where, organizer, guests and description. */
export function EventDetails({ event }: { event: GoogleEvent }) {
  const description = plainText(event.description);
  const guests = (event.attendees ?? "").split(/[;,]/).map(guest => guest.trim()).filter(Boolean);
  const when = event.allDay ? "All day" : event.start ? `${event.start.toLocaleString([], { weekday: "long", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" })}${event.end ? ` – ${event.end.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}` : ""}` : "Time unavailable";
  return <details className="mt-1 text-xs">
    <summary className="cursor-pointer rounded-sm font-medium text-primary focus-visible:outline-2 focus-visible:outline-ring">Details</summary>
    <dl className="mt-2 grid grid-cols-[auto_1fr] gap-x-3 gap-y-1">
      <dt className="text-muted-foreground">When</dt><dd>{when}</dd>
      {event.location && <><dt className="text-muted-foreground">Where</dt><dd className="break-words">{event.location}</dd></>}
      {event.organizer && <><dt className="text-muted-foreground">Organizer</dt><dd className="break-all">{event.organizer}</dd></>}
      <dt className="text-muted-foreground">Guests</dt><dd className="break-all">{guests.length ? guests.join(", ") : "None"}</dd>
      {event.status && <><dt className="text-muted-foreground">Status</dt><dd className="capitalize">{event.status}</dd></>}
      {event.recurring && <><dt className="text-muted-foreground">Repeats</dt><dd>Yes — edit the series in Google Calendar</dd></>}
    </dl>
    {description ? <p className="mt-2 whitespace-pre-wrap break-words rounded-lg bg-muted/40 p-2">{description}</p> : <p className="mt-2 text-muted-foreground">No description.</p>}
    <p className="mt-2"><SourceLink href={event.url}>Open in Google Calendar</SourceLink></p>
  </details>;
}

/** Edit and Delete for one event; repeating events are edited in Google Calendar so the whole series isn't changed by mistake. */
export function EventActions({ event, calendarId }: { event: GoogleEvent; calendarId: string }) {
  const queryClient = useQueryClient();
  const confirm = useConfirm();
  const [editing, setEditing] = useState(false);
  const [duplicating, setDuplicating] = useState(false);
  const duplicateButton = <>
    <Button variant="ghost" size="icon" className="size-8 rounded-full" aria-label={`Duplicate ${event.title}`} title="Duplicate" onClick={() => setDuplicating(true)}><Copy aria-hidden="true" /></Button>
    {duplicating && <EventDialog calendarId={calendarId} initial={draftFor(event)} heading="Duplicate event" open={duplicating} onOpenChange={setDuplicating} />}
  </>;
  if (event.recurring) return <div className="flex shrink-0 items-center gap-1">{duplicateButton}<SourceLink href={event.url} className="text-xs">Edit in Google Calendar</SourceLink></div>;
  async function remove() {
    const ok = await confirm({ title: `Delete "${event.title}"?`, description: "This removes the event from Google Calendar. If it has guests, Google may tell them it was cancelled.", confirmLabel: "Delete", destructive: true });
    if (!ok) return;
    try {
      await deleteGoogleEvent(calendarId, event.id);
      toast.success(`Deleted "${event.title}".`);
      recordHistory({ kind: "action", title: `Deleted Google Calendar event "${event.title}"`, section: "plugins", status: "done", cost: "free" });
      await queryClient.invalidateQueries({ queryKey: ["plugins", "googlecalendar", "events"] });
    } catch { toast.error(`Couldn't delete "${event.title}".`); }
  }
  return <div className="flex shrink-0 gap-1">
    <Button variant="ghost" size="icon" className="size-8 rounded-full" aria-label={`Edit ${event.title}`} title="Edit" onClick={() => setEditing(true)}><Pencil aria-hidden="true" /></Button>
    {duplicateButton}
    <Button variant="ghost" size="icon" className="size-8 rounded-full" aria-label={`Delete ${event.title}`} title="Delete" onClick={() => void remove()}><Trash2 aria-hidden="true" /></Button>
    {editing && <EventDialog calendarId={calendarId} event={event} open={editing} onOpenChange={setEditing} />}
  </div>;
}

function AddEventButton({ calendarId }: { calendarId: string }) {
  const [open, setOpen] = useState(false);
  return <>
    <div className="flex items-center gap-1">
      <Button type="button" size="sm" className="rounded-full" onClick={() => setOpen(true)}><CalendarPlus aria-hidden="true" />Add event</Button>
      <InfoTip label="adding and editing events" title="Adding and editing events" points={[
        "Events are added to the calendar selected above.",
        "Guests you add may get an email invitation from Google; deleting an event may tell them it was cancelled.",
        "Repeating events are edited in Google Calendar, so a change can't accidentally alter the whole series.",
        "Shows 7, 14 or 30 days at a time; use the arrows to move. Search looks across the next 90 days of the selected calendar.",
        "Each event list is one read from Google Calendar; very busy ranges may not include every event.",
        "Duplicate copies an event's details into a new event you can adjust before saving.",
        "Your organization can turn adding or editing events off; you'll see an error if it has.",
        "Each add, edit or delete counts toward your daily request allowance.",
      ]} />
    </div>
    {open && <EventDialog calendarId={calendarId} open={open} onOpenChange={setOpen} />}
  </>;
}

const SEARCH_DAYS = 90;
function GoogleEvents({ calendarId }: { calendarId: string }) {
  const [from, setFrom] = useState(startOfToday);
  const [span, setSpan] = useState(CAL_DAYS);
  const [term, setTerm] = useState("");
  const [query, setQuery] = useState("");
  const isThisWeek = from.toDateString() === startOfToday().toDateString();
  const searching = Boolean(query);
  // Searching looks across the 90 days from today; browsing shows one 7-day window.
  const windowStart = searching ? startOfToday() : from;
  const to = new Date(windowStart); to.setDate(windowStart.getDate() + (searching ? SEARCH_DAYS : span));
  const events = useQuery({ queryKey: ["plugins", "googlecalendar", "events", calendarId, windowStart.toISOString(), query, searching ? SEARCH_DAYS : span], queryFn: () => loadGoogleEvents(calendarId, windowStart, to, query) });
  const shift = (days: number) => setFrom(current => { const next = new Date(current); next.setDate(next.getDate() + days); return next; });
  const last = new Date(to); last.setDate(last.getDate() - 1);
  const rangeLabel = `${windowStart.toLocaleDateString([], { month: "short", day: "numeric" })} – ${last.toLocaleDateString([], { month: "short", day: "numeric" })}`;
  const toolbar = <div className="flex flex-wrap items-center justify-between gap-2">
    <div className="flex flex-wrap items-center gap-1" role="group" aria-label="Choose the week">
      <Button variant="outline" size="icon" className="size-8 rounded-full" aria-label={`Previous ${span} days`} disabled={searching} onClick={() => shift(-span)}><ChevronLeft aria-hidden="true" /></Button>
      <Button variant="outline" size="sm" className="rounded-full" disabled={searching || isThisWeek} onClick={() => setFrom(startOfToday())}>Today</Button>
      <Button variant="outline" size="icon" className="size-8 rounded-full" aria-label={`Next ${span} days`} disabled={searching} onClick={() => shift(span)}><ChevronRight aria-hidden="true" /></Button>
      <label htmlFor={`gcal-span-${calendarId}`} className="sr-only">Days to show</label>
      <NativeSelect id={`gcal-span-${calendarId}`} value={String(span)} disabled={searching} onChange={e => setSpan(Number(e.target.value))} className="h-8">
        {[7, 14, 30].map(days => <NativeSelectOption key={days} value={String(days)}>{days} days</NativeSelectOption>)}
      </NativeSelect>
      <span className="ml-1 text-sm font-medium" aria-live="polite">{searching ? `Search: next ${SEARCH_DAYS} days` : rangeLabel}</span>
    </div>
    <form role="search" className="flex items-center gap-1" onSubmit={e => { e.preventDefault(); setQuery(term.trim()); }}>
      <label htmlFor={`gcal-search-${calendarId}`} className="sr-only">Search events</label>
      <Input id={`gcal-search-${calendarId}`} value={term} onChange={e => { setTerm(e.target.value); if (!e.target.value) setQuery(""); }} placeholder="Search events" className="h-8 w-44" />
      <Button type="submit" variant="outline" size="icon" className="size-8 rounded-full" aria-label="Search" disabled={!term.trim()}><Search aria-hidden="true" /></Button>
      {searching && <Button type="button" variant="ghost" size="sm" className="rounded-full" onClick={() => { setTerm(""); setQuery(""); }}>Clear</Button>}
    </form>
  </div>;
  if (events.isPending) return <div className="space-y-3">{toolbar}<Loading label="Loading events…" /></div>;
  if (events.isError) return <div className="space-y-3">{toolbar}<ConnectorErrorState error={events.error} connector="Google Calendar" className="py-6" action={<Button variant="outline" size="sm" onClick={() => void events.refetch()}><RefreshCw aria-hidden="true" />Try again</Button>} /></div>;
  const upcoming = events.data.filter(event => !isThisWeek || searching || !event.end || event.end.getTime() > Date.now() || event.allDay);
  if (upcoming.length === 0) return <div className="space-y-3">{toolbar}<p className="rounded-xl border border-dashed border-border p-3 text-sm text-muted-foreground">{searching ? `No events match "${query}" in the next ${SEARCH_DAYS} days.` : `Nothing on this calendar in these ${span} days.`}</p><AddEventButton calendarId={calendarId} /></div>;
  const days = new Map<string, { date: Date; items: GoogleEvent[] }>();
  for (const event of upcoming) {
    const date = event.start ? new Date(event.start) : from;
    date.setHours(0, 0, 0, 0);
    const key = date.toDateString();
    if (!days.has(key)) days.set(key, { date, items: [] });
    days.get(key)!.items.push(event);
  }
  const todayCount = days.get(startOfToday().toDateString())?.items.length ?? 0;
  return <div className="space-y-4">
    {toolbar}
    <div className="flex justify-end"><AddEventButton calendarId={calendarId} /></div>
    <dl className="grid grid-cols-2 gap-2 sm:grid-cols-3">
      {isThisWeek && !searching && <Fact label="Today" value={todayCount} />}
      <Fact label={searching ? "Matches" : isThisWeek ? `Next ${span} days` : `These ${span} days`} value={upcoming.length} />
      {upcoming[0]?.start && <Fact label="Next up" value={<span className="text-base">{upcoming[0].allDay ? dayLabel(upcoming[0].start) : upcoming[0].start.toLocaleString([], { weekday: "short", hour: "numeric", minute: "2-digit" })}</span>} />}
    </dl>
    <ol className="space-y-4">{[...days.values()].map(day => <li key={day.date.toDateString()}>
      <h4 className="flex items-center justify-between gap-2 font-heading text-base font-bold">
        {dayLabel(day.date)}
        <SourceLink href={googleCalendarDayUrl(day.date)} className="text-xs font-medium">View day</SourceLink>
      </h4>
      <ul className="mt-1 divide-y divide-border rounded-xl border border-border">{day.items.map(event => <li key={event.id} className="flex items-start gap-3 px-3 py-2.5">
        <span className="w-28 shrink-0 text-xs font-semibold tabular-nums text-muted-foreground">{timeRange(event)}</span>
        <div className="min-w-0 flex-1">
          <SourceLink href={event.url} className="break-words">{hideEmails(event.title)}</SourceLink>
          <p className="flex flex-wrap gap-x-3 text-xs text-muted-foreground">
            {event.status === "tentative" && <span>Tentative</span>}
            {event.location && <span className="inline-flex items-center gap-1"><MapPin aria-hidden="true" className="size-3" />{hideEmails(event.location)}</span>}
            {event.recurring && <span>Repeats</span>}
          </p>
          <EventDetails event={event} />
        </div>
        <EventActions event={event} calendarId={calendarId} />
      </li>)}</ul>
    </li>)}</ol>
  </div>;
}

function GoogleCalendarPanel({ open, onOpenChange, move }: PanelProps) {
  const calendars = useQuery({ queryKey: ["plugins", "googlecalendar", "calendars"], queryFn: loadGoogleCalendars });
  const [chosen, setChosen] = useState<string>();
  const calendarId = chosen ?? calendars.data?.[0]?.id;
  const count = calendars.data?.length ?? 0;
  const summary = calendars.isPending ? "Loading…" : calendars.isError ? "Couldn't load Google Calendar"
    : `${count} ${count === 1 ? "calendar" : "calendars"}`;
  // A Google account's main calendar is named after, and identified by, the account's email address.
  const main = calendars.data?.find(cal => cal.isPrimary);
  const account: PanelAccount = { checking: calendars.isPending, name: main?.id };
  return <PluginPanel id="googlecalendar" icon={CalendarDays} title="Google Calendar" summary={summary} open={open} onOpenChange={onOpenChange} account={account} move={move}
    sourceUrl="https://calendar.google.com/calendar/u/0/r/week" sourceLabel="Open Google Calendar">
    {calendars.isPending ? <Loading label="Loading your calendars…" />
      : calendars.isError ? <ConnectorErrorState error={calendars.error} connector="Google Calendar" className="py-6" action={<Button variant="outline" size="sm" onClick={() => void calendars.refetch()}><RefreshCw aria-hidden="true" />Try again</Button>} />
      : !calendarId ? <p className="rounded-xl border border-dashed border-border p-3 text-sm text-muted-foreground">No Google calendars found.</p>
      : <div className="space-y-4">
        {count > 1 && <div className="flex flex-wrap items-center gap-2">
          <label htmlFor="google-calendar-choice" className="text-sm font-medium">Calendar</label>
          <NativeSelect id="google-calendar-choice" value={calendarId} onChange={event => setChosen(event.target.value)}>
            {calendars.data.map(cal => <NativeSelectOption key={cal.id} value={cal.id}>{cal.isPrimary ? `${cal.name} (main)` : cal.name}</NativeSelectOption>)}
          </NativeSelect>
        </div>}
        <GoogleEvents key={calendarId} calendarId={calendarId} />
      </div>}
  </PluginPanel>;
}

type PanelKey = PluginKey;
const DEFAULT_PLUGIN_ORDER: PanelKey[] = [...PLUGIN_KEYS];
const PANEL_TITLES: Record<PanelKey, string> = { github: "GitHub", googledrive: "Google Drive", box: "Box", googlecalendar: "Google Calendar" };

/** GitHub, Google Drive, Box and Google Calendar — one section, a collapsible panel for each, in the order you choose. */
export function PluginsPanel({ order, onOrderChange }: { order: PanelKey[]; onOrderChange: (order: PanelKey[]) => void }) {
  const [open, setOpen] = useState<Record<PanelKey, boolean>>({ github: false, googledrive: true, box: false, googlecalendar: false });
  const [announcement, setAnnouncement] = useState("");
  const set = (key: PanelKey) => (value: boolean) => setOpen(current => ({ ...current, [key]: value }));
  const allOpen = Object.values(open).every(Boolean);
  const keys = [...order.filter(key => key in PANEL_TITLES), ...DEFAULT_PLUGIN_ORDER.filter(key => !order.includes(key))];
  function move(key: PanelKey, direction: -1 | 1) {
    const from = keys.indexOf(key), to = from + direction;
    if (to < 0 || to >= keys.length) return;
    const next = [...keys];
    [next[from], next[to]] = [next[to], next[from]];
    onOrderChange(next);
    setAnnouncement(`${PANEL_TITLES[key]} moved to position ${to + 1} of ${keys.length}.`);
    requestAnimationFrame(() => document.getElementById(`move-plugin-${key}-${direction < 0 ? "up" : "down"}`)?.focus());
  }
  const controls = (key: PanelKey) => {
    const index = keys.indexOf(key);
    return <div role="group" aria-label={`Reorder ${PANEL_TITLES[key]}`} className="flex gap-1">
      <Button id={`move-plugin-${key}-up`} type="button" variant="outline" size="icon" className="size-9 rounded-full" disabled={index === 0}
        aria-label={`Move ${PANEL_TITLES[key]} up`} title={`Move ${PANEL_TITLES[key]} up`} onClick={() => move(key, -1)}><ArrowUp aria-hidden="true" /></Button>
      <Button id={`move-plugin-${key}-down`} type="button" variant="outline" size="icon" className="size-9 rounded-full" disabled={index === keys.length - 1}
        aria-label={`Move ${PANEL_TITLES[key]} down`} title={`Move ${PANEL_TITLES[key]} down`} onClick={() => move(key, 1)}><ArrowDown aria-hidden="true" /></Button>
    </div>;
  };
  const panel = (key: PanelKey) => {
    const props = { open: open[key], onOpenChange: set(key), move: controls(key) };
    return key === "github" ? <GitHubPanel key={key} {...props} /> : key === "googledrive" ? <DrivePanel key={key} {...props} />
      : key === "box" ? <BoxPanel key={key} {...props} /> : <GoogleCalendarPanel key={key} {...props} />;
  };
  return <div className="space-y-3">
    <div className="flex justify-end">
      <Button variant="ghost" size="sm" className="rounded-full" onClick={() => setOpen({ github: !allOpen, googledrive: !allOpen, box: !allOpen, googlecalendar: !allOpen })}>
        {allOpen ? "Collapse all" : "Expand all"}
      </Button>
    </div>
    {keys.map(panel)}
    <p className="sr-only" aria-live="polite">{announcement}</p>
  </div>;
}
