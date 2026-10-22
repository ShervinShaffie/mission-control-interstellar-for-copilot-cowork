import { RichText } from "@/components/rich-text";
import { firstLinkUrl } from "@/lib/source-links";
import { PersonName } from "@/components/person-name";
import { hideEmails } from "@/lib/privacy";
import { useEffect, useState, type ReactNode } from "react";
import { useQuery } from "@tanstack/react-query";
import { AlertCircle, BookOpen, BriefcaseBusiness, CalendarDays, ChevronLeft, ChevronRight, ExternalLink, FolderOpen, Mail, MessageSquare, RefreshCw, Users } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Card } from "@/components/ui/card";
import { EmptyState, ErrorState, LoadingState } from "@/components/states";
import { safeUrl } from "@/lib/safe-url";
import { cn } from "@/lib/utils";
import { loadChats, loadChatMessages, loadRootFolders, loadFolder, WorkContextError, type WorkItem } from "@/lib/command-center-data";
import type { WorkSignals } from "@/hooks/use-work-signals";
import type { BlobMetadata } from "../../generated/models/OneDriveforBusinessModel";
import { messageReadState, type ChatRecord } from "@/lib/command-center-evidence";
import { teamsFailureDescription, replyFailureDescription } from "@/lib/command-center-failures";
type ContextQuery = WorkSignals["priorities"];
export function WorkSource({ url, children, className }: { url?: string; children: ReactNode; className?: string }) {
  const href = safeUrl(url);
  return href ? <a href={href} target="_blank" rel="noopener noreferrer" className={cn("block rounded-lg p-4 transition-colors hover:bg-muted focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-ring", className)}>{children}</a> : <div className={cn("p-4", className)}>{children}<p className="mt-2 text-xs text-muted-foreground">Source link unavailable.</p></div>;
}
function dateLabel(value: string) {
  if (!value) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return new Date(`${value}T12:00:00`).toLocaleDateString([], { month: "short", day: "numeric", year: "numeric" });
  return new Date(value).toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
}
function itemUrl(item: WorkItem) {
  if (safeUrl(item.url)) return item.url;
  const cited = firstLinkUrl(item.evidence) ?? firstLinkUrl(item.summary);
  if (safeUrl(cited)) return cited;
  if (["project", "training", "action", "urgent", "reply", "prep"].includes(item.kind)) return `https://outlook.office.com/mail/inbox?searchquery=${encodeURIComponent(item.title)}`;
  return undefined;
}
function ContextRow({ item }: { item: WorkItem }) {
  const linked = !!safeUrl(itemUrl(item));
  const health = item.kind === "project" ? item.status.toLowerCase() : "";
  const initials = item.title.split(/\s+/).slice(0, 2).map(part => part[0]).join("");
  return <WorkSource url={itemUrl(item)} className={item.kind === "urgent" ? "urgent-item" : undefined}>
    <div className="flex items-start gap-3">
      {item.kind === "person" && <span className="follow-avatar flex size-10 shrink-0 items-center justify-center rounded-full text-sm font-semibold" aria-hidden="true">{initials}</span>}
      <div className="min-w-0 flex-1 space-y-2">
        <div className="flex items-start justify-between gap-3">
          <p className="min-w-0 break-words font-semibold">{health && <span aria-hidden="true" className={cn("mr-2 inline-block size-2.5 rounded-full", health === "hot" ? "bg-red-700" : health === "watch" ? "bg-amber-700" : health === "active" ? "bg-green-700" : "bg-primary")} />}{item.title}</p>
          {safeUrl(itemUrl(item)) && <ExternalLink aria-hidden="true" className="mt-1 size-4 shrink-0" />}
        </div>
        {item.summary && <p className="break-words text-sm"><RichText text={item.summary} interactive={!linked} /></p>}
        {item.next && <p className="break-words text-sm"><strong>{item.kind === "prep" ? "Prepare:" : "Next:"}</strong> <RichText text={item.next} interactive={!linked} /></p>}
        {(item.status || item.severity) && <div className="flex flex-wrap gap-2">{item.status && <Badge variant="outline" className="h-auto whitespace-normal">Status: {item.status}</Badge>}{item.severity && <Badge variant="outline" className="h-auto whitespace-normal">Urgency: {item.severity}</Badge>}</div>}
        {item.date && <p className="text-sm">{item.kind === "file" ? "Modified: " : item.kind === "oof" ? "Away: " : item.kind === "action" || item.kind === "urgent" ? "Due: " : "When: "}{dateLabel(item.date)}{item.endDate && ` – ${dateLabel(item.endDate)}`}</p>}
        {item.relatedMeeting && <p className="break-words text-sm font-medium">Organizes your meeting: {item.relatedMeeting}</p>}
        <p className="break-words text-xs text-muted-foreground">Evidence: <RichText text={item.evidence} interactive={!linked} /></p>
        {!item.url && safeUrl(itemUrl(item)) && item.kind !== "person" && <p className="text-xs text-muted-foreground">Opens an inbox search for this item.</p>}
      </div>
    </div>
  </WorkSource>;
}
function dueByToday(date: string) {
  const day = /^\d{4}-\d{2}-\d{2}$/.test(date) ? new Date(`${date}T12:00:00`) : new Date(date);
  if (Number.isNaN(day.getTime())) return false;
  const end = new Date(); end.setHours(23, 59, 59, 999);
  return day <= end;
}
export function ContextSection({ title, kind, query, children, urgent = false, emptyTitle = "Nothing found", emptyDescription }: { title: string; kind: string; query: ContextQuery; children?: ReactNode; urgent?: boolean; emptyTitle?: string; emptyDescription?: string }) {
  // "Needs attention today" also shows deadlines that are overdue or due today.
  const rows = (query.data?.filter(item => item.kind === kind || urgent && item.kind === "action" && !!item.date && dueByToday(item.date)) ?? [])
    .sort((a, b) => urgent ? (a.kind === b.kind ? (a.date || "9").localeCompare(b.date || "9") : a.kind === "urgent" ? -1 : 1) : 0);
  const Icon = kind === "person" || kind === "oof" ? Users : kind === "file" ? FolderOpen : kind === "project" ? BriefcaseBusiness : kind === "training" ? BookOpen : kind === "reply" ? Mail : kind === "urgent" ? AlertCircle : CalendarDays;
  return <Card className={cn("min-w-0 gap-0 p-0", urgent && "urgent-banner")}>
    <section aria-labelledby={`${kind}-heading`} aria-busy={query.isFetching}>
      {urgent && <span className="urgent-watermark" aria-hidden="true">⚠</span>}
      <h2 id={`${kind}-heading`} className="relative px-5 pb-4 pt-5 font-heading text-xl font-bold">{title}</h2>
      {query.isPending ? <div className="px-5 pb-5"><LoadingState rows={3} /></div> : <>
        {query.isError && <div role="alert"><ErrorState title="This couldn't be read right now" description={query.data !== undefined ? "Showing what was loaded last time. Select Retry, or open Connections at the top." : "Select Retry, or open Connections at the top to check your sign-ins."} action={<Button variant="outline" disabled={query.isFetching} onClick={() => void query.refetch({ cancelRefetch: false })}>Retry {title.toLowerCase()}</Button>} className="px-5 py-5" /></div>}
        {rows.length ? <ul className="divide-y divide-border">{rows.map(item => <li key={`${item.kind}|${item.title}|${item.url}|${item.date}`}><ContextRow item={item} /></li>)}</ul> : !query.isError && <EmptyState icon={Icon} title={emptyTitle} description={emptyDescription} action={<Button variant="outline" disabled={query.isFetching} onClick={() => void query.refetch({ cancelRefetch: false })}>Check again</Button>} className="px-5 py-7" />}
      </>}
      {children}
    </section>
  </Card>;
}
export function ReplySection({ signals }: { signals: WorkSignals }) {
  const { replies: query, profile } = signals;
  const owner = profile.data?.id;
  const rows = query.data ?? [];
  const hasData = query.data !== undefined;
  const paused = query.fetchStatus === "paused";
  const validationFailed = query.error instanceof WorkContextError && query.error.stage === "validation";
  const search = <Button variant="outline" className="h-auto whitespace-normal" disabled={query.isFetching || paused} onClick={() => void query.refetch({ cancelRefetch: false })}>{query.isError ? "Retry email search" : hasData ? "Search again" : "Search emails"}</Button>;
  const announcement = !owner
    ? profile.isFetching ? "Waiting for your profile before searching email." : "Email search has not started."
    : query.isFetching
      ? hasData ? "Searching email again. Last validated results are still shown." : "Searching email for reply requests. This can take about 30 seconds."
      : paused ? "Email search is paused. Check your connection before searching again."
      : query.isError ? "Email search failed. No new result was validated."
      : query.isSuccess ? rows.length ? "Email search finished. Cited reply suggestions are shown." : "Email search finished with no cited reply suggestions."
      : "Email search has not started.";
  return <Card className="min-w-0 gap-0 p-0"><section aria-labelledby="reply-heading" aria-busy={query.isFetching || !owner && profile.isFetching}>
    <h2 id="reply-heading" className="px-5 pb-4 pt-5 font-heading text-xl font-bold">Emails needing a reply</h2>
    <p role="status" aria-live="polite" className="px-5 pb-3 text-sm">{announcement}</p>
    {!owner ? profile.isError
      ? <div role="alert"><ErrorState title="Email search has not started" description="Your profile could not be verified. Retry profile verification before searching your email." action={<Button variant="outline" className="h-auto whitespace-normal" disabled={profile.isFetching} onClick={() => void profile.refetch({ cancelRefetch: false })}>Retry profile verification</Button>} className="px-5 pb-5 pt-2" /></div>
      : <EmptyState icon={Mail} title="Waiting for profile verification" action={<Button variant="outline" className="h-auto whitespace-normal" disabled={profile.isFetching} onClick={() => void profile.refetch({ cancelRefetch: false })}>Verify profile</Button>} className="px-5 pb-5 pt-2" />
      : <>
        {query.isError && <div role="alert"><ErrorState title={validationFailed ? "Email results could not be validated" : "Email search or response failed"} description={`${validationFailed ? "Returned records failed request, evidence or record validation. This is not an empty result." : replyFailureDescription(query.error instanceof WorkContextError ? query.error.failureCause : "unknown")} ${hasData ? "Last successfully validated results are retained below." : ""}`} action={search} className="px-5 pb-5 pt-2" /></div>}
        {!hasData && query.isFetching && <div className="px-5 pb-5"><LoadingState rows={3} /></div>}
        {hasData && rows.length > 0 && <><ul className="divide-y divide-border">{rows.map(item => <li key={`${item.kind}|${item.title}|${item.url}|${item.date}`}><ContextRow item={item} /></li>)}</ul>{!query.isError && <div className="px-5 pb-5">{search}</div>}</>}
        {hasData && rows.length === 0 && <EmptyState icon={Mail} title={query.isFetching || query.isError || paused ? "The last validated search found no reply suggestions" : "No cited reply suggestions found"} description="Search covers received and sent email from the last 30 days; it is not an exhaustive unanswered-mail inventory." action={!query.isError ? search : undefined} className="px-5 pb-5 pt-2" />}
        {!hasData && !query.isFetching && !query.isError && <EmptyState icon={Mail} title={paused ? "Email search is paused" : "Email search has not started"} action={search} className="px-5 pb-5 pt-2" />}
      </>}
  </section></Card>;
}
export function WorkSearchStatus({ signals }: { signals: WorkSignals }) {
  const [slow, setSlow] = useState(false);
  const fetching = signals.priorities.isFetching || signals.meetings.isFetching || signals.replies.isFetching || signals.collaboration.isFetching;
  useEffect(() => {
    if (!fetching) return;
    const reset = window.setTimeout(() => setSlow(false), 0);
    const timer = window.setTimeout(() => setSlow(true), 10000);
    return () => { window.clearTimeout(reset); window.clearTimeout(timer); };
  }, [fetching]);
  return <p aria-live="polite" role="status" className="sr-only">{fetching ? slow ? "Still loading your Microsoft 365 data…" : "Loading your Microsoft 365 data." : ""}</p>;
}
export function SignalTiles({ signals, tomorrowCount, waitingCount, calendarPartial, calendarError, onOpen }: { signals: WorkSignals; tomorrowCount?: number; waitingCount?: number | "—"; calendarPartial: boolean; calendarError: boolean; onOpen: (id: "urgent" | "inbox" | "agenda" | "teams" | "project" | "oof") => void }) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const actions = signals.priorities.data?.filter(item => ["action", "urgent"].includes(item.kind) && item.date && (item.date.length === 10 ? new Date(`${item.date}T00:00:00`) < today : new Date(item.date) < new Date()));
  const overdue = actions ? new Set(actions.map(item => `${item.url || item.title}|${item.date}`)).size : undefined;
  const oof = signals.collaboration.data?.filter(item => item.kind === "oof" && item.date && new Date(`${(item.endDate || item.date).slice(0, 10)}T23:59:59`) >= today && new Date(`${item.date.slice(0, 10)}T00:00:00`) <= new Date(today.getFullYear(), today.getMonth(), today.getDate() + (7 - (today.getDay() || 7))));
  const values = [
    { label: "Overdue actions", value: overdue ?? (signals.priorities.isError ? "—" : undefined), scope: "Deadlines mentioned in your teams" },
    { label: "Unread emails", value: signals.unread.data ? `${signals.unread.data.count}${signals.unread.data.capped ? "+" : ""}` : signals.unread.isError ? "—" : undefined, scope: signals.unread.data?.capped ? "Inbox · first 100 unread checked" : "Unread in Inbox" },
    { label: "Meetings tomorrow", value: tomorrowCount ?? (calendarError ? "—" : undefined), scope: calendarPartial ? "Accepted meetings · current page" : "Accepted timed meetings with attendees" },
    { label: "Waiting in Teams", value: waitingCount, scope: "Chats where someone else wrote last (7 days)" },
    { label: "Active projects", value: signals.priorities.data?.filter(item => item.kind === "project").length ?? (signals.priorities.isError ? "—" : undefined), scope: "Your project teams" },
    { label: "OOF this week", value: oof ? new Set(oof.map(item => item.email || item.title)).size : signals.collaboration.isError ? "—" : undefined, scope: "People in your meetings with automatic replies on" },
  ];
  const targets = ["urgent", "inbox", "agenda", "teams", "project", "oof"] as const;
  return <section aria-labelledby="summary-heading"><h2 id="summary-heading" className="mb-4 font-heading text-2xl font-bold">Work summary</h2><div className="grid gap-px overflow-visible rounded-xl border border-border bg-border sm:grid-cols-3 xl:grid-cols-6">{values.map((value, index) => <button type="button" key={value.label} onClick={() => onOpen(targets[index])} className="min-w-0 bg-card px-5 py-5 text-left focus-visible:outline-2 focus-visible:outline-ring">
    <span className="block text-sm font-medium">{value.label}</span><span className="my-2 block text-3xl font-bold tabular-nums">{value.value ?? "…"}</span><span className="block text-xs text-muted-foreground">{value.scope}</span>
  </button>)}</div>{(calendarError || signals.priorities.isError || signals.unread.isError || signals.collaboration.isError) && <p role="alert" className="mt-3 text-sm">Some summary sources could not be refreshed. Last successfully loaded values are retained where available; unavailable values are shown as a dash.</p>}</section>;
}
export function PrioritySections({ signals }: { signals: WorkSignals }) {
  return <ContextSection title="Needs attention today" kind="urgent" query={signals.priorities} urgent />;
}
export function PrepSection({ signals }: { signals: WorkSignals }) {
  return <ContextSection title="Prep notes" kind="prep" query={signals.meetings} />;
}
export function RemainingWorkSections({ signals }: { signals: WorkSignals }) {
  return <div className="grid items-start gap-6 lg:grid-cols-2 xl:grid-cols-3">
    <ContextSection title="Top projects" kind="project" query={signals.priorities} />
    <ReplySection signals={signals} />
    <ContextSection title="Long-term action items" kind="action" query={signals.priorities} />
    <ContextSection title="Training & learning" kind="training" query={signals.meetings} />
    <ContextSection title="OOF radar" kind="oof" query={signals.collaboration} />
    <ContextSection title="People to follow up with" kind="person" query={signals.collaboration} />
    <TeamsSection currentUserId={signals.profile.data?.id} />
    <ContextSection title="Recent files in flight" kind="file" query={signals.collaboration} />
    <DriveSection />
  </div>;
}
export function DatedWorkSection({ signals }: { signals: WorkSignals }) {
  const queries = [signals.priorities, signals.meetings];
  const rows = [...(signals.priorities.data ?? []), ...(signals.meetings.data ?? [])].filter(item => ["urgent", "action", "training"].includes(item.kind) && item.date).sort((a, b) => Date.parse(a.date) - Date.parse(b.date));
  const fetching = queries.some(query => query.isFetching);
  const retry = <Button variant="outline" disabled={fetching} onClick={() => { for (const query of queries) void query.refetch({ cancelRefetch: false }); }}>Refresh dated work</Button>;
  return <Card className="min-w-0 gap-0 p-0"><section aria-labelledby="dated-heading" aria-busy={fetching}>
    <h2 id="dated-heading" className="px-5 pt-5 font-heading text-xl font-bold">Dated work</h2>
    <p className="px-5 py-3 text-sm text-muted-foreground">Actions and training ordered by their source dates, including past dates. Undated items remain in their individual components.</p>
    {queries.some(query => query.isPending) && <LoadingState className="px-5" />}
    {queries.some(query => query.isError) && <ErrorState title="Some dated work could not be refreshed" description="Last successfully loaded results are retained where available." action={retry} className="px-5 py-5" />}
    {rows.length ? <ul className="divide-y divide-border">{rows.map(item => <li key={`${item.kind}|${item.title}|${item.url}|${item.date}`}><p className="px-4 pt-3 text-xs font-medium">{item.kind === "training" ? "Training" : item.kind === "urgent" ? "Urgent action" : "Long-term action"}</p><ContextRow item={item} /></li>)}</ul> : queries.every(query => query.isSuccess) && <EmptyState icon={CalendarDays} title="No cited actions or training with dates" action={retry} className="px-5 py-7" />}
  </section></Card>;
}
function messagePreview(summary?: string, content?: string) {
  const text = summary || (content ?? "").replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, " ").replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, " ").replace(/<[^>]*>/g, " ").replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/&lt;/gi, "<").replace(/&gt;/gi, ">").replace(/&quot;/gi, '"').replace(/&#39;/g, "'").replace(/\s+/g, " ").trim();
  return text ? `${text.slice(0, 360)}${text.length > 360 ? "…" : ""}` : "Message preview unavailable.";
}
function ChatPreview({ chat, unreadOnly, currentUserId, showAll }: { chat: ChatRecord; unreadOnly: boolean; currentUserId?: string; showAll: () => void }) {
  const [expanded, setExpanded] = useState(false);
  const messages = useQuery({ queryKey: ["work-chat-messages", chat.id], queryFn: () => loadChatMessages(chat.id), retry: false });
  const annotated = (messages.data?.rows ?? []).map(message => ({ message, state: messageReadState(message, chat, currentUserId) }));
  const unread = annotated.filter(row => row.state === "unread");
  const unknown = annotated.filter(row => row.state === "unknown");
  const filtered = unreadOnly ? unread : annotated.filter(row => row.state !== "excluded");
  const visible = expanded ? filtered : filtered.slice(0, 3);
  const unreadSummary = unread.length > 0 || chat.readCutoff && currentUserId && unknown.length === 0
    ? `Returned messages: ${unread.length} confirmed unread${unknown.length ? ` · ${unknown.length} unknown` : ""}.`
    : `Unread count unknown${unknown.length ? ` · ${unknown.length} messages without sufficient evidence` : ""}.`;
  const labels = { unread: "Unread", read: "Read through cutoff", own: "Your message", excluded: "Excluded", unknown: "Unread unknown" };
  return <li className="min-w-0 py-4" aria-busy={messages.isFetching}>
    <WorkSource url={chat.url} className="px-0"><h3 className="break-words font-semibold">{hideEmails(chat.topic)}</h3>{chat.hidden === true && <p className="mt-1 text-xs">Hidden in Teams</p>}</WorkSource>
    <p role="status" aria-live="polite" className="mb-2 text-sm">{messages.isFetching ? messages.data !== undefined ? "Refreshing messages; last validated previews are retained." : "Loading chat messages…" : messages.fetchStatus === "paused" ? "Message loading is paused." : messages.isError ? "Message loading failed. No new result was validated." : ""}</p>
    {messages.isError && <div role="alert"><ErrorState title="Messages could not be refreshed" description={`${teamsFailureDescription(messages.error)} ${messages.data !== undefined ? "Showing the last successfully loaded previews." : ""}`} action={<Button variant="outline" className="h-auto whitespace-normal" disabled={messages.isFetching || messages.fetchStatus === "paused"} onClick={() => void messages.refetch({ cancelRefetch: false })}>Retry messages</Button>} className="py-5" /></div>}
    {messages.isPending && messages.isFetching ? <LoadingState rows={2} /> : messages.data !== undefined && <>
      <p className="mb-2 text-sm" role="status">{unreadSummary}</p>
      {(!chat.readCutoff || !currentUserId) && <p className="mb-3 text-xs text-muted-foreground">{!chat.readCutoff ? "Your read cutoff was not returned for this chat." : "Your identity could not be verified to exclude your own messages."} Unread status remains unknown where evidence is missing.</p>}
      {visible.length ? <ul className="divide-y divide-border">{visible.map(({ message, state }) => <li key={message.id}><WorkSource url={chat.url} className="px-0">
        <div className="flex flex-wrap items-start justify-between gap-2"><p className="break-words text-sm font-semibold"><PersonName value={message.sender} fallback="Sender unavailable" /></p><Badge variant={state === "unread" ? "default" : "outline"} className="h-auto whitespace-normal">{labels[state]}</Badge></div>
        <p className="mt-2 break-words text-sm">{hideEmails(messagePreview(message.summary, message.content))}</p>
        <p className="mt-2 text-xs text-muted-foreground">{message.created ? dateLabel(message.created) : "Message time unavailable"}</p>
      </WorkSource></li>)}</ul> : <EmptyState icon={MessageSquare} title={unreadOnly ? "No confirmed unread previews in these messages" : "No user messages returned"} action={unreadOnly ? <Button variant="outline" onClick={showAll}>Show all previews</Button> : <Button variant="outline" disabled={messages.isFetching} onClick={() => void messages.refetch({ cancelRefetch: false })}>Refresh messages</Button>} className="py-5" />}
      {filtered.length > 3 && <Button variant="outline" className="mt-2 h-auto whitespace-normal" aria-expanded={expanded} onClick={() => setExpanded(value => !value)}>{expanded ? "Show fewer previews" : "Show all returned previews"}</Button>}
      {messages.data?.hasMore && <p className="mt-3 text-xs text-muted-foreground">More history exists, but this action has no continuation parameter. Open the chat for the rest.</p>}
    </>}
  </li>;
}
export function TeamsSection({ currentUserId }: { currentUserId?: string }) {
  const [type, setType] = useState("oneOnOne");
  const [chatPage, setChatPage] = useState(0);
  const [unreadOnly, setUnreadOnly] = useState(false);
  const chats = useQuery({ queryKey: ["work-chats", type], queryFn: () => loadChats(type), retry: false });
  const paused = chats.fetchStatus === "paused";
  const rows = chats.data?.rows ?? [];
  const chatRows = rows.slice(chatPage * 5, (chatPage + 1) * 5);
  return <Card className="min-w-0 gap-0 p-5"><section aria-labelledby="teams-heading">
    <h2 id="teams-heading" className="font-heading text-xl font-bold">Teams chat previews</h2>
    <p className="my-3 text-sm text-muted-foreground">Unread means a message from someone else created after your verified read cutoff. Missing evidence stays unknown; unread-only shows confirmed unread previews. Five recent chats per page, up to 10 messages each; counts cover returned messages, not chat or Teams-wide totals. Source links open the full chat.</p>
    <label htmlFor="chat-type" className="text-sm font-medium">Chat type</label>
    <select id="chat-type" value={type} onChange={event => { setType(event.target.value); setChatPage(0); }} className="mt-1 w-full rounded-lg border border-input bg-background p-2 focus-visible:outline-2 focus-visible:outline-ring"><option value="oneOnOne">One-to-one</option><option value="group">Group</option><option value="meeting">Meeting</option></select>
    <div className="mt-3 flex flex-wrap gap-2" role="group" aria-label="Filter Teams previews"><Button variant={!unreadOnly ? "default" : "outline"} aria-pressed={!unreadOnly} className="h-auto whitespace-normal" onClick={() => setUnreadOnly(false)}>All previews</Button><Button variant={unreadOnly ? "default" : "outline"} aria-pressed={unreadOnly} className="h-auto whitespace-normal" onClick={() => setUnreadOnly(true)}>Unread only</Button></div>
    <div className="mt-4" aria-busy={chats.isFetching}>
      <p role="status" aria-live="polite" className="mb-3 text-sm">{chats.isFetching ? chats.data !== undefined ? "Refreshing chats; last validated results are retained." : "Loading recent chats…" : paused ? "Chat loading is paused." : chats.isError ? "Chat loading failed. No new result was validated." : chats.isSuccess ? "Chat-list request finished." : "Chat loading has not started."}</p>
      {chats.isError && <div role="alert"><ErrorState title="Chats could not be refreshed" description={`${teamsFailureDescription(chats.error)} ${chats.data !== undefined ? "Showing the last successfully loaded chat list." : ""}`} action={<Button className="h-auto whitespace-normal" disabled={chats.isFetching || paused} onClick={() => void chats.refetch({ cancelRefetch: false })}>Retry chats</Button>} className="py-6" /></div>}
      {chats.isPending && chats.isFetching ? <LoadingState /> : chats.data !== undefined && (!chatRows.length ? <EmptyState icon={MessageSquare} title={chats.isError || chats.isFetching || paused ? "The last validated chat list had no chats on this page" : "No chats on this page"} action={<Button variant="outline" className="h-auto whitespace-normal" disabled={chats.isFetching || paused} onClick={() => chatPage > 0 ? setChatPage(0) : void chats.refetch({ cancelRefetch: false })}>{chatPage > 0 ? "First preview page" : "Check again"}</Button>} className="py-6" /> : <ul className="divide-y divide-border">{chatRows.map(chat => <ChatPreview key={chat.id} chat={chat} currentUserId={currentUserId} unreadOnly={unreadOnly} showAll={() => setUnreadOnly(false)} />)}</ul>)}
      {chats.data === undefined && paused && <EmptyState icon={MessageSquare} title="Chat loading is paused" description="No chat-list result has been validated." action={<Button variant="outline" className="h-auto whitespace-normal" disabled>Refresh chat list</Button>} className="py-6" />}
    </div>
    {chats.data !== undefined && <p className="mt-3 text-xs text-muted-foreground">Recent chats · preview page {chatPage + 1} of {Math.max(1, Math.ceil(rows.length / 5))}.{chats.data.hasMore ? " Additional chats exist, but this action cannot page through them." : " This is not an exhaustive Teams inventory."}</p>}
    <div className="mt-4 flex flex-wrap gap-2">{chatPage > 0 && <Button variant="outline" onClick={() => setChatPage(value => value - 1)}>Previous page</Button>}{(chatPage + 1) * 5 < rows.length && <Button variant="outline" onClick={() => setChatPage(value => value + 1)}>Load more previews</Button>}<Button variant="outline" className="h-auto whitespace-normal" disabled={chats.isFetching || paused} onClick={() => void chats.refetch({ cancelRefetch: false })}><RefreshCw aria-hidden="true" />Refresh chat list</Button></div>
  </section></Card>;
}
export function DriveSection() {
  const [folder, setFolder] = useState<BlobMetadata | undefined>();
  const [token, setToken] = useState<string | undefined>();
  const [history, setHistory] = useState<(string | undefined)[]>([]);
  const root = useQuery({ queryKey: ["work-drive-root"], queryFn: loadRootFolders, retry: false });
  const contents = useQuery({ queryKey: ["work-drive-folder", folder?.Id, token], enabled: !!folder?.Id, queryFn: () => { if (!folder?.Id) throw new Error("Folder unavailable"); return loadFolder(folder.Id, token); }, retry: false });
  const query = folder ? contents : root;
  const rows = folder ? contents.data?.rows : root.data;
  const sorted = [...(rows ?? [])].sort((a, b) => (Date.parse(b.LastModified || "") || 0) - (Date.parse(a.LastModified || "") || 0));
  const [displayPage, setDisplayPage] = useState(0);
  const visible = sorted.slice(displayPage * 20, (displayPage + 1) * 20);
  return <Card className="min-w-0 gap-0 p-5"><section aria-labelledby="drive-heading">
    <h2 id="drive-heading" className="font-heading text-xl font-bold">Folder browser</h2>
    <p className="my-3 text-sm text-muted-foreground">Folder metadata has no web links. See Recently active for linked folders.</p>
    {folder && <Button variant="outline" className="h-auto whitespace-normal" onClick={() => { setFolder(undefined); setToken(undefined); setHistory([]); setDisplayPage(0); }}><ChevronLeft aria-hidden="true" />Back to root</Button>}
    <p className="mt-3 break-words text-sm font-medium">{folder?.DisplayName || folder?.Name || "Root folders and files"}</p>
    <div className="mt-4" aria-busy={query.isFetching}>
      {query.isError && <ErrorState title="Folder contents could not be refreshed" description={query.data !== undefined ? "Showing the last successfully loaded folder contents." : undefined} action={<Button disabled={query.isFetching} onClick={() => void query.refetch({ cancelRefetch: false })}>Retry folder</Button>} className="py-6" />}
      {query.isPending ? <LoadingState /> : query.data !== undefined && (!visible.length ? <EmptyState icon={FolderOpen} title="No items on this page" action={<Button variant="outline" disabled={query.isFetching} onClick={() => void query.refetch({ cancelRefetch: false })}>Check again</Button>} className="py-6" /> : <ul className="divide-y divide-border">{visible.map(item => <li key={item.Id || item.Path} className="py-3">
      {item.IsFolder && item.Id ? <Button variant="ghost" className="h-auto w-full justify-start whitespace-normal p-0 text-left" onClick={() => { setFolder(item); setToken(undefined); setHistory([]); setDisplayPage(0); }}><FolderOpen aria-hidden="true" /><span className="min-w-0 break-words">{item.DisplayName || item.Name || "Unnamed folder"}</span></Button> : <p className="break-words font-medium">{item.DisplayName || item.Name || "Unnamed file"}</p>}
      {item.LastModified && <p className="mt-1 text-xs text-muted-foreground">Modified {dateLabel(item.LastModified)}</p>}{item.LastModifiedBy && <p className="mt-1 break-words text-xs text-muted-foreground">By <PersonName value={item.LastModifiedBy} /></p>}
    </li>)}</ul>)}</div>
    <div className="mt-4 flex flex-wrap gap-2">
      {(displayPage > 0 || history.length > 0) && <Button variant="outline" disabled={query.isFetching} onClick={() => { if (displayPage > 0) setDisplayPage(value => value - 1); else { setToken(history[history.length - 1]); setHistory(value => value.slice(0, -1)); } }}>Previous page</Button>}
      {(sorted.length > (displayPage + 1) * 20 || contents.data?.next && folder) && <Button variant="outline" disabled={query.isFetching} onClick={() => { if (sorted.length > (displayPage + 1) * 20) setDisplayPage(value => value + 1); else if (contents.data?.next) { setHistory(value => [...value, token]); setToken(contents.data.next); setDisplayPage(0); } }}>Load more<ChevronRight aria-hidden="true" /></Button>}
      <Button variant="outline" disabled={query.isFetching} onClick={() => void query.refetch({ cancelRefetch: false })}>Refresh folder</Button>
    </div>
  </section></Card>;
}
