import { OneNoteSection } from "@/components/direct-sources";
import { ProjectTeamPicker } from "@/components/team-files";
import bannerImage from "@/assets/interstellar-banner.jpg";
import { PersonName } from "@/components/person-name";
import { hideEmails } from "@/lib/privacy";
import { useEffect, useRef, useState, type CSSProperties, type ReactNode } from "react";
import { refreshActiveData } from "@/lib/manual-refresh";
import { useQuery, useQueryClient, useIsFetching } from "@tanstack/react-query";
import { WorkSearchStatus, SignalTiles, ContextSection, TeamsSection } from "@/components/work-context";
import { FilesPanel } from "@/components/files-panel";
import { collectFileHits } from "@/lib/direct-signals";
import { rankFiles } from "@/lib/files";
import { TodoTasks } from "@/components/todo-tasks";
import { ProjectBoard } from "@/components/project-board";
import { QuickActions } from "@/components/quick-actions";
import { buildProjects, loadJoinedTeams } from "@/lib/projects";
import { TodayBriefing, type SourceCoverage, type SourceStatus } from "@/components/today-briefing";
import { buildAttention } from "@/lib/attention";
import { setDismissOwner } from "@/lib/dismissed";
import { useOpenTasks } from "@/lib/todo";
import { WaitingChats } from "@/components/waiting-chats";
import { loadWaitingChats } from "@/lib/teams-waiting";
import { SectionCard } from "@/components/section-card";
import { CreateForm } from "@/components/create-panel";
import { buildDayPlan, calendarActions, emailActions, filesActions, insightActions, oneNoteActions, teamsActions, todayActions, workActions } from "@/lib/suggestions";
import { loadChats } from "@/lib/command-center-data";
import { loadOneNotePages } from "@/lib/direct-sources";
import { Input } from "@/components/ui/input";
import { SEARCH_LABELS, type SearchGroup } from "@/lib/search-plan";
import { DashboardCustomizer } from "@/components/dashboard-customizer";
import { HowItWorks } from "@/components/how-it-works";
import { HistoryButton, HistoryPanel } from "@/components/history-panel";
import { ConnectionsButton } from "@/components/connections-panel";
import { initHistory } from "@/lib/history";
import { classify, reportConnection } from "@/lib/connection-health";
import { DashboardLayout } from "@/components/dashboard-layout";
import { PluginsPanel } from "@/components/plugins-panel";
import { AgendaTasks } from "@/components/agenda-tasks";
import { AgendaGoogle, type BusyBlock } from "@/components/agenda-google";
import { useCombinedGoogle } from "@/lib/combined-google";
import { Checkbox } from "@/components/ui/checkbox";
import { InfoTip } from "@/components/info-tip";
import { isoDay, useDatedTasks } from "@/lib/day-tasks";
import { PlannerPlans, PlannerSection } from "@/components/planner-section";
import { SectionCollapseContext, type SectionCollapse } from "@/lib/section-collapse";
import { useDashboardPreferences } from "@/hooks/use-dashboard-preferences";
import { type ComponentId, effectiveSwitches } from "@/lib/dashboard-preferences";
import { EmptyState } from "@/components/states";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { useProfile, useWorkSignals } from "@/hooks/use-work-signals";
import { searchPlan } from "@/lib/search-plan";
import { RefreshConfirm } from "@/components/refresh-confirm";
import { RichText } from "@/components/rich-text";
import { CalendarInsights } from "@/components/calendar-insights";
import { AlertCircle, BriefcaseBusiness, CalendarDays, ChartPie, ClipboardList, ExternalLink, FolderOpen, Mail, MessagesSquare, Moon, NotebookPen, Puzzle, RefreshCw, Sparkles, Sun } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Card } from "@/components/ui/card";
import { safeUrl } from "@/lib/safe-url";
import { cn } from "@/lib/utils";
import { Office365OutlookService } from "../../generated/services/Office365OutlookService";
import { eventTitle, isPrivateEvent, isAcceptedMeeting, type CalendarEvent } from "@/lib/command-center-evidence";
import { loadCalendarCategories, loadCalendarPage } from "@/lib/command-center-data";
import { suggestedTargets } from "@/lib/time-targets";

const PAGE_SIZE = 40;
const responseLabels = {
  none: "No response recorded", organizer: "Organizer", tentativelyAccepted: "Tentative",
  accepted: "Accepted", declined: "Declined", notResponded: "Not responded",
};
function dayKey(date: Date) {
  return `${date.getFullYear()}-${date.getMonth()}-${date.getDate()}`;
}
function instant(value?: string) {
  if (!value || !/(Z|[+-]\d{2}:\d{2})$/i.test(value)) return null;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}
function eventStart(event: CalendarEvent) { return instant(event.startWithTimeZone ?? event.start); }
function eventEnd(event: CalendarEvent) { return instant(event.endWithTimeZone ?? event.end); }
function eventKey(event: CalendarEvent) {
  return event.id ?? `${event.iCalUId}|${event.startWithTimeZone ?? event.start}|${event.subject}`;
}
function time(date: Date) { return date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }); }
function refreshed(timestamp: number) {
  return new Intl.DateTimeFormat(undefined, {
    month: "short", day: "numeric", hour: "numeric", minute: "2-digit",
    timeZoneName: "short",
  }).format(timestamp);
}
function SourceLink({ url, children, className }: { url?: string; children: ReactNode; className?: string }) {
  const href = safeUrl(url);
  return href ? <a href={href} target="_blank" rel="noopener noreferrer" className={cn("block rounded-lg focus-visible:outline-2 focus-visible:outline-offset-4 focus-visible:outline-ring hover:bg-muted", className)}>{children}</a>
    : <div className={className}>{children}<p className="mt-2 text-sm text-muted-foreground">Source link unavailable.</p></div>;
}
function Notice({ kind, children, action }: { kind: "loading" | "error" | "empty"; children: ReactNode; action?: ReactNode }) {
  const Icon = kind === "error" ? AlertCircle : kind === "loading" ? RefreshCw : CalendarDays;
  return <div className="flex flex-col items-center gap-3 px-5 py-12 text-center" role={kind === "error" ? "alert" : "status"}>
    <Icon className="size-6 text-muted-foreground" aria-hidden="true" />
    <p>{children}</p>{action}
  </div>;
}
function EventItem({ event, number, allEvents = [], prepTitles = [] }: { event: CalendarEvent; number: number; allEvents?: CalendarEvent[]; prepTitles?: string[] }) {
  const start = eventStart(event), end = eventEnd(event);
  const conflicts = !event.isAllDay && event.responseType !== "declined" && event.showAs !== "free" && start && end ? allEvents.filter(other => {
    const otherStart = eventStart(other), otherEnd = eventEnd(other);
    return eventKey(other) !== eventKey(event) && !other.isAllDay && other.responseType !== "declined" && other.showAs !== "free" && otherStart && otherEnd && start < otherEnd && end > otherStart;
  }) : [];
  const hasPrep = prepTitles.some(title => title.trim().toLowerCase() === hideEmails(event.subject)?.trim().toLowerCase());
  return <SourceLink url={event.webLink} className="px-4 py-4">
    <div className="flex items-start gap-3">
      <span className="flex min-h-7 min-w-7 shrink-0 items-center justify-center rounded-full border border-border px-1 text-sm font-semibold" aria-label={`Event ${number}`}>{number}</span>
      <div className="min-w-0 flex-1 space-y-2">
        <div className="flex flex-wrap items-start justify-between gap-2">
          <p className="font-semibold break-words">{eventTitle(event)}</p>
          {safeUrl(event.webLink) && <ExternalLink className="size-4 shrink-0 text-muted-foreground" aria-hidden="true" />}
        </div>
        <p className="text-sm">{event.isAllDay ? "All day" : start ? `${time(start)}${end ? ` – ${time(end)}` : ""}` : "Time unavailable"}</p>
        <div className="flex flex-wrap gap-2">
          <Badge variant="outline" className={cn("h-auto whitespace-normal", event.responseType === "accepted" && "bg-green-700 text-white border-green-700", event.responseType === "tentativelyAccepted" && "bg-amber-800 text-white border-amber-800", event.responseType === "notResponded" && "bg-red-700 text-white border-red-700")}>
            Response: {event.responseType ? responseLabels[event.responseType] : "Not available"}
          </Badge>
          {event.importance && event.importance !== "normal" && <Badge variant="outline" className="h-auto whitespace-normal">Importance: {event.importance}</Badge>}
          {hasPrep && !isPrivateEvent(event) && <Badge variant="outline" className="h-auto whitespace-normal">Prep note found</Badge>}
          {!isPrivateEvent(event) && <Badge variant="outline" className="h-auto whitespace-normal">{event.onlineStatus === true ? `Join link · ${event.onlineProvider ?? "online meeting"}` : "No join link in invite"}</Badge>}
        </div>
        {conflicts.length > 0 && <p className="break-words text-sm font-medium">Overlaps on this page: {conflicts.map(eventTitle).join("; ")}</p>}
        {!isPrivateEvent(event) && event.organizer && <p className="break-words text-sm text-muted-foreground">Organizer: <PersonName value={event.organizer} fallback="Outside your directory" /></p>}
        {!isPrivateEvent(event) && hideEmails(event.location) && <p className="break-words text-sm text-muted-foreground">{hideEmails(event.location)}</p>}
      </div>
    </div>
  </SourceLink>;
}

function CalendarPageTable({ rows, prepTitles }: { rows: CalendarEvent[]; prepTitles: string[] }) {
  return <div className="p-4" role="region" aria-label="Calendar-page event table"><Table className="w-full table-fixed">
    <TableHeader><TableRow><TableHead className="h-auto w-[28%] whitespace-normal align-top">Date</TableHead><TableHead className="h-auto whitespace-normal align-top">Appointment and response</TableHead></TableRow></TableHeader>
    <TableBody>{rows.map((event, index) => <TableRow key={eventKey(event)}><TableCell className="min-w-0 whitespace-normal align-top break-words">{eventStart(event)?.toLocaleDateString([], { weekday: "long", month: "short", day: "numeric" }) ?? "Date unavailable"}</TableCell><TableCell className="min-w-0 whitespace-normal align-top break-words p-0"><EventItem event={event} number={index + 1} allEvents={rows} prepTitles={prepTitles} /></TableCell></TableRow>)}</TableBody>
  </Table><p className="mt-3 text-xs text-muted-foreground">Overlap checks use this calendar page. Private and confidential titles are masked. Online meetings are identified by a join link in the invitation.</p></div>;
}

export function HomePage() {
  const queryClient = useQueryClient();
  const activeRequests = useIsFetching();
  const refreshLock = useRef(false);
  const [manualRefreshing, setManualRefreshing] = useState(false);
  const [refreshResult, setRefreshResult] = useState<"success" | "error" | null>(null);
  const [anchor, setAnchor] = useState(() => {
    const date = new Date(); date.setHours(0, 0, 0, 0); return date;
  });
  const days = Array.from({ length: 7 }, (_, index) => {
    const date = new Date(anchor); date.setDate(date.getDate() + index); return date;
  });
  const [selectedDay, setSelectedDay] = useState(1);
  const [agendaOffset, setAgendaOffset] = useState(0);
    const [offset, setOffset] = useState(0);
  const [mailFilter, setMailFilter] = useState<"unread" | "flagged" | "recent">("unread");
  const [systemDark, setSystemDark] = useState(() => window.matchMedia("(prefers-color-scheme: dark)").matches);
  useEffect(() => {
    const media = window.matchMedia("(prefers-color-scheme: dark)");
    const changed = () => setSystemDark(media.matches);
    media.addEventListener("change", changed);
    return () => media.removeEventListener("change", changed);
  }, []);
  const calendars = useQuery({
    queryKey: ["personal-calendars"], retry: false,
    queryFn: async () => {
      const result = await Office365OutlookService.CalendarGetTables();
      if (!result.success) { console.error("Calendar discovery failed", result.error); throw new Error("Calendar unavailable"); }
      return result.data.value ?? [];
    },
  });
  const calendarId = calendars.data?.find(row => row.DisplayName?.toLowerCase() === "calendar")?.Name || calendars.data?.[0]?.Name;
  const rangeEnd = new Date(anchor); rangeEnd.setDate(rangeEnd.getDate() + 7);
  const events = useQuery({
    queryKey: ["personal-calendar", calendarId, anchor.toISOString(), offset],
    enabled: Boolean(calendarId), retry: false,
    queryFn: async () => {
      if (!calendarId) throw new Error("Calendar unavailable");
      return loadCalendarPage(calendarId, anchor.toISOString(), rangeEnd.toISOString(), offset, PAGE_SIZE);
    },
  });
  const agendaEnd = new Date(days[selectedDay]); agendaEnd.setDate(agendaEnd.getDate() + 1);
  const agenda = useQuery({
    queryKey: ["personal-day-agenda", calendarId, days[selectedDay].toISOString(), agendaOffset],
    enabled: Boolean(calendarId), retry: false,
    queryFn: async () => {
      if (!calendarId) throw new Error("Calendar unavailable");
      return loadCalendarPage(calendarId, days[selectedDay].toISOString(), agendaEnd.toISOString(), agendaOffset, PAGE_SIZE, "start/dateTime asc");
    },
  });
  const mail = useQuery({
    queryKey: ["personal-inbox", mailFilter], retry: false,
    queryFn: async () => {
      const result = await Office365OutlookService.GetEmailsV3("Inbox", undefined, undefined, undefined, undefined, undefined, false, undefined, mailFilter === "unread", mailFilter === "flagged", undefined, false, undefined, 10);
      if (!result.success) { console.error("Inbox request failed", result.error); throw new Error("Inbox unavailable"); }
      return result.data.value ?? [];
    },
  });
  const profile = useProfile();
  const { preferences, update, storage, identityVerified } = useDashboardPreferences(profile.data?.id);
  const plan = searchPlan(preferences.visible, effectiveSwitches());
  const signals = useWorkSignals(calendarId, calendars.isSuccess || calendars.isError, plan, profile);
  const focusOwner = signals.profile.data?.id ?? "unverified";
  const dark = preferences.theme === "dark" || preferences.theme === "system" && systemDark;
  useEffect(() => {
    document.documentElement.classList.toggle("dark", dark);
    document.documentElement.classList.toggle("light", !dark);
  }, [dark]);
  function openComponent(id: ComponentId) {
    if (!preferences.visible.includes(id)) update({ ...preferences, visible: [...preferences.visible, id] });
    requestAnimationFrame(() => {
      const target = document.getElementById(`component-${id}`);
      target?.focus();
      target?.scrollIntoView({ behavior: "instant", block: "start" });
    });
  }
  const rows = events.data ?? [];
  const prepTitles = signals.meetings.data?.filter(item => item.kind === "prep").map(item => item.title) ?? [];
  const tomorrowCount = rows.filter(event => { const start = eventStart(event), end = eventEnd(event); return start && end && end > start && dayKey(start) === dayKey(days[1]) && isAcceptedMeeting(event); }).length;
  const calendarPartial = offset > 0 || rows.length === PAGE_SIZE || rows.some(event => { const start = eventStart(event), end = eventEnd(event); return !start || !end || end <= start; });
  const agendaRows = agenda.data ?? [];
  const selected = agendaRows.map((event, index) => ({ event, number: index + 1 }));
  const datedTasks = useDatedTasks();
  const google = useCombinedGoogle(days[0], 7, preferences.combineGoogle);
  const outlookBusy: BusyBlock[] = agendaRows.flatMap(event => {
    const start = eventStart(event), end = eventEnd(event);
    return !event.isAllDay && event.responseType !== "declined" && event.showAs !== "free" && start && end ? [{ title: eventTitle(event), start, end }] : [];
  });
  const agendaError = calendars.isError || agenda.isError;
  const agendaLoading = calendars.isPending || Boolean(calendarId) && agenda.isPending;
  const calendarError = calendars.isError || events.isError;
  const calendarLoading = calendars.isPending || (Boolean(calendarId) && events.isPending);
  const zone = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const anyFetching = manualRefreshing || activeRequests > 0;
  const refreshFailed = refreshResult === "error" || refreshResult !== null && queryClient.getQueryCache().findAll({ type: "active" }).some(query => query.state.status === "error");
  async function refreshAll() {
    if (refreshLock.current || queryClient.isFetching() > 0) return;
    refreshLock.current = true;
    setManualRefreshing(true);
    setRefreshResult(null);
    const date = new Date(); date.setHours(0, 0, 0, 0);
    if (dayKey(date) !== dayKey(anchor)) { setAnchor(date); setOffset(0); setAgendaOffset(0); }
    try {
      // Let an explicitly changed date commit; join its initial reads, rather than restart them.
      await new Promise<void>(resolve => requestAnimationFrame(() => resolve()));
      const result = await refreshActiveData(queryClient);
      setRefreshResult(result.failed > 0 ? "error" : "success");
    } catch { setRefreshResult("error"); }
    finally { refreshLock.current = false; setManualRefreshing(false); }
  }
  function resetCalendar() { if (events.isFetching || calendars.isFetching) return; if (offset > 0) setOffset(0); else if (calendarId) void events.refetch({ cancelRefetch: false }); else void calendars.refetch({ cancelRefetch: false }); }
  function resetAgenda() { if (agenda.isFetching || calendars.isFetching) return; if (calendarId) void agenda.refetch({ cancelRefetch: false }); else void calendars.refetch({ cancelRefetch: false }); }
  function retryCalendar() { if (events.isFetching || calendars.isFetching) return; if (calendarId) void events.refetch({ cancelRefetch: false }); else void calendars.refetch({ cancelRefetch: false }); }

  const calendarComponent = <Card className="gap-0 p-0" aria-labelledby="week-heading">
      <section>
        <div className="flex flex-wrap items-start justify-between gap-4 px-5 pt-6 sm:px-7">
          <div>
            <h2 id="week-heading" className="font-heading text-2xl font-bold">{preferences.calendarView === "graphic" && preferences.calendarDays === 1 ? "Day view" : preferences.calendarView === "graphic" && preferences.calendarDays === 3 ? "Three-day view" : "Next seven days"}</h2>
            <p className="mt-2 text-sm text-muted-foreground">{days[0].toLocaleDateString([], { month: "short", day: "numeric" })} – {days[6].toLocaleDateString([], { month: "short", day: "numeric" })} · {zone}</p>
          </div>
          <div className="flex flex-wrap items-center gap-3">
          <div className="flex items-center gap-2">
            <Checkbox id="combine-google" checked={preferences.combineGoogle} onCheckedChange={value => update({ ...preferences, combineGoogle: value === true })} />
            <label htmlFor="combine-google" className="text-sm font-medium">Include Google Calendar</label>
            <InfoTip label="the combined calendar" title="Outlook and Google together" points={[
              "Adds events from your main Google calendar to the 7-day view and the day agenda.",
              "In the 7-day view, Google events are the outlined bars at the right edge of each day; Outlook events are the numbered blocks.",
              "Outlook events are read-only here. Google events can be edited, duplicated or deleted from the agenda.",
              "A meeting that's on both calendars shows twice; the agenda flags Google events that overlap Outlook ones.",
              "Needs the Google Calendar connection, and Google Calendar allowed by your organization.",
              "Loading Google events adds one read each time the calendar loads.",
            ]} />
          </div>
          <div role="group" aria-label="Calendar view" className="flex flex-wrap gap-2">
            {(["graphic", "table"] as const).map(view => <Button key={view} variant={preferences.calendarView === view ? "default" : "outline"} aria-pressed={preferences.calendarView === view} onClick={() => update({ ...preferences, calendarView: view })} className="h-auto whitespace-normal">{view === "graphic" ? "Graphic" : "Table"}</Button>)}
          </div>
          </div>
        </div>
        {calendarError && <Notice kind="error" action={<div className="flex flex-wrap gap-2">{offset > 0 && <Button variant="outline" disabled={events.isFetching} onClick={() => setOffset(value => Math.max(0, value - PAGE_SIZE))}>Previous page</Button>}<Button variant="outline" disabled={events.isFetching || calendars.isFetching} onClick={retryCalendar}>{offset > 0 ? "Retry loading this page" : "Retry calendar"}</Button></div>}>Could not refresh your calendar. {events.data !== undefined ? "Showing the last successfully loaded events." : "Retry, or check your calendar access."}</Notice>}
        {calendarLoading ? <Notice kind="loading">Loading your calendar…</Notice>
          : calendarError && events.data === undefined ? null
          : !calendarId ? <Notice kind="empty" action={<Button variant="outline" disabled={calendars.isFetching} onClick={() => void calendars.refetch({ cancelRefetch: false })}>Check again</Button>}>No calendars were returned.</Notice>
          : rows.length === 0 ? <Notice kind="empty" action={<div className="flex flex-wrap gap-2">{offset > 0 && <Button variant="outline" onClick={() => setOffset(value => Math.max(0, value - PAGE_SIZE))}>Previous page</Button>}<Button variant="outline" disabled={events.isFetching || calendars.isFetching} onClick={resetCalendar}>Reload calendar</Button></div>}>{offset > 0 ? "No more events on this page." : "No events in the next seven days."}</Notice>
          : <>
            {preferences.calendarView === "graphic" && (() => {
              const span = preferences.calendarDays;
              const startIdx = span === 7 ? 0 : Math.min(selectedDay, 7 - span);
              const shown = days.slice(startIdx, startIdx + span);
              const colW = 945 / span, hourH = span === 7 ? 9.5 : span === 3 ? 16 : 30, top = 82, bottom = top + 24 * hourH, height = bottom + 28;
              const wide = span < 7;
              const hours = span === 7 ? [0, 6, 12, 18, 24] : span === 3 ? [0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24] : Array.from({ length: 25 }, (_, h) => h);
              // Larger type in the 1-day view, where there's room for it.
              const fs = span === 1 ? { label: 17, hour: 15, head: 19, small: 14 } : { label: 11, hour: 12, head: 13, small: 11 };
              const fit = (text: string, width: number) => { const max = Math.max(3, Math.floor(width / (fs.label * 0.56))); return text.length > max ? `${text.slice(0, max - 1)}…` : text; };
              const outlookBusy = rows.flatMap(event => { const s0 = eventStart(event), e0 = eventEnd(event); return !event.isAllDay && event.responseType !== "declined" && event.showAs !== "free" && s0 && e0 ? [{ title: eventTitle(event), start: s0, end: e0 }] : []; });
              const viewLabel = span === 7 ? "Seven-day" : span === 3 ? "Three-day" : "One-day";
              const step = (delta: number) => { const next = Math.max(0, Math.min(6, selectedDay + delta)); setSelectedDay(next); setAgendaOffset(0); };
              return <figure className="px-5 pt-5 sm:px-7">
              <div className="mb-3 flex flex-wrap items-center gap-2">
                <div role="group" aria-label="Days shown in the graphic" className="flex gap-1">
                  {([1, 3, 7] as const).map(n => <Button key={n} size="sm" variant={span === n ? "default" : "outline"} aria-pressed={span === n} className="rounded-full" onClick={() => update({ ...preferences, calendarDays: n })}>{n === 1 ? "1 day" : `${n} days`}</Button>)}
                </div>
                {wide && <div role="group" aria-label="Move the view" className="flex items-center gap-1">
                  <Button size="sm" variant="outline" className="rounded-full" disabled={startIdx === 0} onClick={() => step(-1)} aria-label="Previous day">‹</Button>
                  <span className="text-sm font-medium" aria-live="polite">{shown[0].toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" })}{span > 1 ? ` – ${shown[shown.length - 1].toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" })}` : ""}</span>
                  <Button size="sm" variant="outline" className="rounded-full" disabled={startIdx + span >= 7} onClick={() => step(1)} aria-label="Next day">›</Button>
                </div>}
                <InfoTip label="the calendar graphic" title="Reading the calendar graphic" points={[
                  "1 and 3 days give each event more room, so overlaps are easier to spot; 7 days shows the whole week.",
                  "Numbered filled blocks are Outlook events; their numbers match the full event list below.",
                  ...(google.enabled ? ["Outlined bars on the right of each day are Google Calendar events.", "In 1- and 3-day views, a Google event that overlaps an Outlook meeting is drawn in amber and marked “overlap”."] : []),
                  "The 1- and 3-day views follow the day selected in the agenda and stay within this 7-day page.",
                  "Declined meetings and time marked free don't count as overlaps.",
                ]} />
              </div>
              <div role="region" aria-label={`${viewLabel} calendar map`} tabIndex={0} className="overflow-x-auto rounded-xl border border-border bg-card focus-visible:outline-2 focus-visible:outline-ring">
                <svg role="img" aria-labelledby="calendar-map-title calendar-map-description" viewBox={`0 0 1024 ${height}`} className="calendar-map block min-w-[700px] w-full">
                  <title id="calendar-map-title">{`${viewLabel} calendar, page ${offset / PAGE_SIZE + 1}`}</title>
                  <desc id="calendar-map-description">Each column is a day. Vertical position represents local time from midnight to midnight. Numbered blocks represent the actual duration of Outlook events on this page. All-day events appear above the time grid, and each day shows how many To Do and Planner tasks are due.{google.enabled ? ` Outlined bars at the right of a day are Google Calendar events${wide ? "; amber ones overlap an Outlook meeting" : ""}.` : ""} Use the agenda and full event list below for titles, responses and source links.</desc>
                  <rect x="0" y="0" width="1024" height={height} fill="var(--card)" />
                  {hours.map(hour => <g key={hour}>
                    <text x="12" y={top + 4 + hour * hourH} className="map-label" fontSize={fs.hour}>{hour === 24 ? "24:00" : `${String(hour).padStart(2, "0")}:00`}</text>
                    <line x1="70" x2="1016" y1={top + hour * hourH} y2={top + hour * hourH} className="map-grid" />
                  </g>)}
                  <text x="12" y="58" fontSize={fs.small} className="map-label">All day</text>
                  {shown.map((day, column) => {
                    const x0 = 70 + column * colW;
                    const tomorrow = new Date(day); tomorrow.setDate(tomorrow.getDate() + 1);
                    const outlookW = wide ? colW * 0.6 : 120, googleX = wide ? x0 + colW * 0.64 : x0 + 123, googleW = wide ? colW * 0.33 : 9;
                    const dayRows = rows.map((event, index) => ({ event, index })).filter(({ event }) => { const s0 = eventStart(event), e0 = eventEnd(event); return s0 && e0 && e0 > day && s0 < tomorrow && e0 > s0; });
                    const lanes = wide ? Math.max(1, Math.min(4, dayRows.length)) : 4;
                    const laneW = outlookW / lanes;
                    return <g key={dayKey(day)}>
                    <rect x={x0} y="8" width={colW} height="28" fill={startIdx + column === selectedDay ? "var(--accent)" : "var(--card)"} />
                    <text x={x0 + colW / 2} y="28" textAnchor="middle" fontSize={fs.head} fontWeight="600" className="map-label">{day.toLocaleDateString([], span === 1 ? { weekday: "long", month: "short", day: "numeric" } : span === 3 ? { weekday: "short", month: "short", day: "numeric" } : { weekday: "short", day: "numeric" })}</text>
                    <line x1={x0} x2={x0} y1="39" y2={bottom} className="map-grid" />
                    {(() => { const n = datedTasks.dueOn(isoDay(day)).length; return n ? <text x={x0 + colW / 2} y="77" textAnchor="middle" fontSize={fs.small} fontWeight="600" className="map-label">{`✓ ${n} to-do${n === 1 ? "" : "s"} due`}</text> : null; })()}
                    {dayRows.map(({ event, index }, order) => {
                      const start = eventStart(event)!, end = eventEnd(event)!;
                      const startMinute = start < day ? 0 : start.getHours() * 60 + start.getMinutes();
                      const endMinute = end >= tomorrow ? 1440 : end.getHours() * 60 + end.getMinutes();
                      const lane = wide ? order % lanes : index % 4;
                      const x = x0 + 5 + lane * (wide ? laneW : 30);
                      const w = wide ? Math.max(18, laneW - 4) : 27;
                      const y = event.isAllDay ? 42 : top + startMinute / 60 * hourH;
                      const h = event.isAllDay ? 23 : Math.max(2, (endMinute - startMinute) / 60 * hourH);
                      const label = wide && (event.isAllDay || h >= fs.label + 3) ? fit(`${index + 1}. ${eventTitle(event)}`, w - 8) : undefined;
                      return <g key={eventKey(event)}><title>{`${index + 1}. ${eventTitle(event)}`}</title>
                        <rect x={x} y={y} width={w} height={h} rx="3" className="map-event" />
                        {label ? <text x={x + 5} y={event.isAllDay ? 59 : y + fs.label} fontSize={fs.label} fontWeight={span === 1 ? 600 : undefined} className="map-event-text">{label}</text>
                          : <text x={x + w / 2} y={event.isAllDay ? 58 : Math.min(y + h + 11, bottom + 16)} textAnchor="middle" fontSize="10" className={event.isAllDay ? "map-event-text" : "map-label"}>{index + 1}</text>}
                      </g>;
                    })}
                    {google.events.map(event => {
                      const start = event.start, end = event.end ?? event.start;
                      if (!start || !end || start >= tomorrow || (end <= day && !(event.allDay && start.toDateString() === day.toDateString()))) return null;
                      const startMinute = event.allDay || start < day ? 0 : start.getHours() * 60 + start.getMinutes();
                      const endMinute = event.allDay || end >= tomorrow ? 1440 : end.getHours() * 60 + end.getMinutes();
                      const y = event.allDay ? 42 : top + startMinute / 60 * hourH;
                      const h = event.allDay ? 23 : Math.max(3, (endMinute - startMinute) / 60 * hourH);
                      const clash = wide && !event.allDay && event.end ? outlookBusy.filter(block => start < block.end && event.end! > block.start).map(block => block.title) : [];
                      const label = wide && (event.allDay || h >= fs.label + 3) ? fit(event.title, googleW - 8) : undefined;
                      return <g key={`g-${event.id}`}><title>{`Google: ${event.title}${clash.length ? ` — overlaps Outlook: ${clash.join("; ")}` : ""}`}</title>
                        <rect x={googleX} y={y} width={googleW} height={h} rx="3" className={clash.length ? "map-google map-google-overlap" : "map-google"} />
                        {label && <text x={googleX + 5} y={event.allDay ? 59 : y + fs.label} fontSize={fs.label} fontWeight={span === 1 ? 600 : undefined} className="map-label">{label}</text>}
                        {clash.length > 0 && h >= fs.label * 2 + 6 && <text x={googleX + 5} y={y + fs.label * 2 + 2} fontSize={fs.small} fontWeight="700" className="map-overlap-text">overlap</text>}
                      </g>;
                    })}
                  </g>;
                  })}
                </svg>
              </div>
              <figcaption className="mt-3 text-sm text-muted-foreground">Block numbers match the full calendar-page event list below. Agenda numbers are separate.{google.enabled ? ` Outlined bars at the right of each day are Google Calendar events${wide ? "; amber ones overlap an Outlook meeting" : ""}; the day agenda lists them.` : ""} Short events use a minimum-height mark; events without an available time appear in the full list.</figcaption>
            </figure>;
            })()}
            {preferences.calendarView === "table" ? <CalendarPageTable rows={rows} prepTitles={prepTitles} /> : <details className="mx-5 mt-4 rounded-lg border border-border sm:mx-7"><summary className="cursor-pointer rounded-lg px-4 py-3 font-medium focus-visible:outline-2 focus-visible:outline-ring">All events on this calendar page</summary><CalendarPageTable rows={rows} prepTitles={prepTitles} /></details>}
            <div className="flex flex-wrap items-center justify-between gap-3 px-5 py-5 sm:px-7">
              <p className="text-sm">Calendar page {offset / PAGE_SIZE + 1}{rows.length === PAGE_SIZE ? " · More events may be available" : ""}</p>
              <div className="flex flex-wrap gap-2">
                {offset > 0 && <Button variant="outline" disabled={events.isFetching} onClick={() => setOffset(value => Math.max(0, value - PAGE_SIZE))} className="h-auto whitespace-normal py-2">Previous page</Button>}
                {rows.length === PAGE_SIZE && <Button variant="outline" disabled={events.isFetching} onClick={() => setOffset(value => value + PAGE_SIZE)} className="h-auto whitespace-normal py-2">Load more</Button>}
              </div>
            </div>
            {events.dataUpdatedAt > 0 && <p className="refresh-stamp px-5 pb-5 text-xs text-muted-foreground sm:px-7">Calendar refreshed {refreshed(events.dataUpdatedAt)}</p>}
          </>}
      </section>
    </Card>;

  const agendaComponent = <section aria-labelledby="agenda-heading" className="min-w-0">
        <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
          <h2 id="agenda-heading" className="font-heading text-2xl font-bold">{selectedDay === 1 ? "Tomorrow’s agenda" : selectedDay === 0 ? "Today’s agenda" : "Day agenda"}</h2>
        </div>
        <div className="mb-4 flex flex-wrap gap-2" role="group" aria-label="Select agenda day">
          {days.map((day, index) => <Button key={dayKey(day)} variant={selectedDay === index ? "default" : "outline"} aria-pressed={selectedDay === index} onClick={() => { setSelectedDay(index); setAgendaOffset(0); }} className="h-auto whitespace-normal px-3 py-2 text-xs">
            {index === 0 ? "Today" : index === 1 ? "Tomorrow" : day.toLocaleDateString([], { weekday: "short", day: "numeric" })}
          </Button>)}
        </div>
        <Card className="gap-0 p-0">
          {agendaError && <Notice kind="error" action={<Button variant="outline" disabled={agenda.isFetching || calendars.isFetching} onClick={resetAgenda}>Retry agenda</Button>}>The agenda could not be refreshed. {agenda.data !== undefined ? "Showing the last successfully loaded events." : "Retry to load events."}</Notice>}
          {agendaLoading ? <Notice kind="loading">Loading events…</Notice>
            : agendaError && agenda.data === undefined ? null
            : selected.length === 0 ? <Notice kind="empty" action={<Button variant="outline" disabled={agenda.isFetching || calendars.isFetching} onClick={resetAgenda}>Reload agenda</Button>}>{agendaOffset > 0 ? "No more events for this day." : "No events for this day."}</Notice>
            : <ul className="divide-y divide-border">{selected.map(({ event, number }) => <li key={eventKey(event)}><EventItem event={event} number={number} allEvents={agendaRows} prepTitles={prepTitles} /></li>)}</ul>}
          <AgendaGoogle google={google} day={days[selectedDay]} outlook={outlookBusy} />
          <AgendaTasks day={days[selectedDay]} isToday={dayKey(days[selectedDay]) === dayKey(new Date())} />
          <div className="flex flex-wrap items-center justify-between gap-2 border-t border-border px-4 py-3">
            <p className="text-xs text-muted-foreground">Agenda page {agendaOffset / PAGE_SIZE + 1} · {days[selectedDay].toLocaleDateString([], { month: "short", day: "numeric" })}</p>
            <div className="flex flex-wrap gap-2">{agendaOffset > 0 && <Button variant="outline" disabled={agenda.isFetching} onClick={() => setAgendaOffset(value => Math.max(0, value - PAGE_SIZE))}>Previous page</Button>}{agendaRows.length === PAGE_SIZE && <Button variant="outline" disabled={agenda.isFetching} onClick={() => setAgendaOffset(value => value + PAGE_SIZE)}>Load more</Button>}</div>
          </div>
          {agenda.data !== undefined && <div className="border-t border-border px-4 py-3 text-xs text-muted-foreground"><p>Overlap checks use this agenda page; overlaps are not conflict decisions. Prep flags match the meeting title. Private and confidential titles are masked; source links retain the original appointment.</p><p className="mt-2">Online meetings are identified by a join link in the invitation.</p></div>}
        </Card>
      </section>;
  const inboxComponent = <section aria-labelledby="inbox-heading" className="min-w-0">
        <h2 id="inbox-heading" className="mb-4 font-heading text-2xl font-bold">Inbox</h2>
        <div className="mb-4 flex flex-wrap gap-2" role="group" aria-label="Filter inbox">
          {(["unread", "flagged", "recent"] as const).map(filter => <Button key={filter} variant={mailFilter === filter ? "default" : "outline"} aria-pressed={mailFilter === filter} onClick={() => setMailFilter(filter)} className="h-auto whitespace-normal px-3 py-2 text-xs">{filter === "unread" ? "Unread" : filter === "flagged" ? "Flagged" : "Recent"}</Button>)}
        </div>
        <Card className="gap-0 p-0">
          {mail.isError && <Notice kind="error" action={<Button variant="outline" disabled={mail.isFetching} onClick={() => void mail.refetch({ cancelRefetch: false })}>Retry inbox</Button>}>Could not refresh your inbox. {mail.data !== undefined ? "Showing the last successfully loaded messages." : "Retry, or check your mail access."}</Notice>}
          {mail.isPending ? <Notice kind="loading">Loading your inbox…</Notice>
            : mail.data === undefined ? null
            : !mail.data.length ? <Notice kind="empty" action={<Button variant="outline" onClick={() => setMailFilter(mailFilter === "recent" ? "unread" : "recent")}>Show {mailFilter === "recent" ? "unread" : "recent"} mail</Button>}>No {mailFilter} messages were returned.</Notice>
            : <ul className="divide-y divide-border">{mail.data.map(message => <li key={message.id ?? message.internetMessageId}>
              <SourceLink url={message.id ? `https://outlook.office.com/mail/deeplink/read/${encodeURIComponent(message.id)}` : undefined} className="space-y-2 px-5 py-4">
                <div className="flex items-start justify-between gap-3"><p className="min-w-0 break-words text-sm font-medium"><PersonName value={message.from} fallback="External sender" /></p><ExternalLink className="size-4 shrink-0 text-muted-foreground" aria-hidden="true" /></div>
                <p className="break-words font-semibold">{hideEmails(message.subject) || "No subject"}</p>
                {message.bodyPreview && <p className="break-words text-sm text-muted-foreground">{hideEmails(message.bodyPreview)}</p>}
                <div className="flex flex-wrap items-center gap-2 text-xs text-muted-foreground">
                  {message.receivedDateTime && <span>{instant(message.receivedDateTime)?.toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }) ?? "Received time unavailable"}</span>}
                  {message.importance && <Badge variant="outline" className="h-auto whitespace-normal">Importance: {message.importance}</Badge>}
                </div>
              </SourceLink>
            </li>)}</ul>}
          {mail.data !== undefined && <div className="space-y-1 border-t border-border px-5 py-4 text-xs text-muted-foreground"><p>Up to 10 {mailFilter} inbox messages.</p><p>Inbox refreshed {refreshed(mail.dataUpdatedAt)}</p></div>}
        </Card>
      </section>;

  // ---------- data used only for suggestions (shared caches, no extra reads) ----------
  const owner = profile.data?.id, mySite = profile.data?.mySite, me = profile.data?.displayName;
  useEffect(() => { setDismissOwner(owner); void initHistory(owner, mySite); }, [owner, mySite]);
  const unread = useQuery({ queryKey: ["personal-inbox", "unread"], retry: false, queryFn: async () => {
    const result = await Office365OutlookService.GetEmailsV3("Inbox", undefined, undefined, undefined, undefined, undefined, false, undefined, true, false, undefined, false, undefined, 10);
    if (!result.success) { console.error("Inbox request failed", result.error); throw new Error("Inbox unavailable"); }
    return result.data.value ?? [];
  } });
  const oneToOne = useQuery({ queryKey: ["work-chats", "oneOnOne"], queryFn: () => loadChats("oneOnOne"), retry: false });
  const groupChats = useQuery({ queryKey: ["work-chats", "group"], queryFn: () => loadChats("group"), retry: false });
  const waiting = useQuery({ queryKey: ["teams-waiting", owner, oneToOne.dataUpdatedAt, groupChats.dataUpdatedAt], enabled: !!owner && oneToOne.isFetched && groupChats.isFetched, retry: false,
    queryFn: () => loadWaitingChats(oneToOne.data?.rows ?? [], groupChats.data?.rows ?? [], owner!) });
  const flagged = useQuery({ queryKey: ["personal-inbox", "flagged-today"], retry: false, queryFn: async () => {
    const result = await Office365OutlookService.GetEmailsV3("Inbox", undefined, undefined, undefined, undefined, undefined, false, undefined, false, true, undefined, false, undefined, 10);
    if (!result.success) throw new Error("Flagged email unavailable");
    return result.data.value ?? [];
  } });
  const openTasks = useOpenTasks();
  const joinedTeams = useQuery({ queryKey: ["teams-joined"], queryFn: loadJoinedTeams, retry: false });
  const notebooks = useQuery({ queryKey: ["direct-source", "onenote"], queryFn: loadOneNotePages, retry: false });
  useEffect(() => { if (events.isSuccess) reportConnection("outlook", "ok"); else if (events.isError) reportConnection("outlook", classify(events.error)); }, [events.isSuccess, events.isError, events.error]);
  useEffect(() => { if (oneToOne.isSuccess || groupChats.isSuccess) reportConnection("teams", "ok"); else if (oneToOne.isError && groupChats.isError) reportConnection("teams", classify(oneToOne.error)); }, [oneToOne.isSuccess, groupChats.isSuccess, oneToOne.isError, groupChats.isError, oneToOne.error]);
  useEffect(() => { if (openTasks.isSuccess) reportConnection("todo", "ok"); else if (openTasks.isError) reportConnection("todo", classify(openTasks.error)); }, [openTasks.isSuccess, openTasks.isError, openTasks.error]);
  useEffect(() => { if (notebooks.isSuccess) reportConnection("onenote", "ok"); else if (notebooks.isError) reportConnection("onenote", classify(notebooks.error)); }, [notebooks.isSuccess, notebooks.isError, notebooks.error]);
  useEffect(() => { if (profile.isSuccess) reportConnection("users", "ok"); else if (profile.isError) reportConnection("users", classify(profile.error)); }, [profile.isSuccess, profile.isError, profile.error]);
  const now = new Date();
  const ctx = { now, calendarId, mySite, me, feeds: signals.teams.data ?? [], tasks: openTasks.data ?? [], waiting: waiting.data ?? [] };
  const work = [...(signals.priorities.data ?? []), ...(signals.meetings.data ?? [])];
  const replies = signals.replies.data ?? [];
  const fileItems = rankFiles(collectFileHits({ mail: signals.mail.data, replyUrls: new Set(replies.map(item => item.url)), feeds: signals.teams.data ?? [], waiting: waiting.data ?? [], events: signals.upcoming.data ?? [], trending: signals.trending.data ?? [], now }), now);
  const filesLoading = [signals.mail, signals.teams, signals.upcoming, signals.trending].some(query => query.isPending && query.fetchStatus !== "idle");
  const oneDriveUrl = mySite && /^https:\/\/[^/]+-my\.sharepoint\.com\//i.test(mySite) ? `${mySite.replace(/\/+$/, "")}/_layouts/15/onedrive.aspx` : "https://www.office.com/launch/onedrive";
  const collaboration = signals.collaboration.data ?? [];
  const dayPlan = buildDayPlan(rows, work, replies, now);
  const shares = categoryShares(rows);
  const categoryList = useQuery({ queryKey: ["calendar-categories"], queryFn: loadCalendarCategories, retry: false });
  const suggested = suggestedTargets((categoryList.data?.categories ?? []).map(category => category.name), shares);
  const targets = preferences.targetsSet ? preferences.targets : suggested;
  const actions = {
    calendar: calendarActions(ctx, rows, signals.meetings.data ?? [], collaboration),
    email: emailActions(ctx, replies, unread.data ?? [], rows),
    teams: teamsActions(waiting.data ?? [], now, signals.teams.data ?? []),
    onenote: oneNoteActions(ctx, notebooks.data ?? [], dayPlan, rows),
    files: filesActions(fileItems),
    insights: insightActions(ctx, rows, shares, targets),
  };
  const attention = buildAttention({ now, calendarId, events: rows, unread: unread.data ?? [], flagged: flagged.data ?? [], waiting: waiting.data ?? [], tasks: openTasks.data ?? [], pages: notebooks.data ?? [],
    replies, priorities: signals.priorities.data ?? [], meetings: signals.meetings.data ?? [], collaboration, feeds: signals.teams.data ?? [], files: fileItems });
  const helpers = todayActions(ctx, dayPlan);
  const projectViews = buildProjects({ now, projects: (signals.priorities.data ?? []).filter(item => item.kind === "project"), work: signals.priorities.data ?? [], meetings: signals.meetings.data ?? [], events: rows, tasks: openTasks.data ?? [], feeds: signals.teams.data ?? [] });
  const projectReports = workActions(ctx, signals.priorities.data ?? []).filter(action => action.id === "tracker" || action.id === "status-report");
  const projectActionList = projectViews.flatMap(view => view.actions.map(action => ({ ...action, id: `pa-${action.id}`, title: `${view.project.title}: ${action.title}`, priority: action.priority / 4 + view.attention })));
  const status = (query: { isPending: boolean; isError: boolean; fetchStatus?: string }, on = true): SourceStatus => !on ? "off" : query.isError ? "error" : query.isPending && query.fetchStatus !== "idle" ? "loading" : query.isPending ? "off" : "ready";
  const coverage: SourceCoverage[] = [
    { key: "calendar", status: calendarError ? "error" : calendarLoading ? "loading" : "ready", read: rows.length, note: "events" },
    { key: "outlook", status: status(unread), read: (unread.data?.length ?? 0) + (flagged.data?.length ?? 0) + replies.length, note: "emails" },
    { key: "teams", label: "Teams chats", status: oneToOne.isError && groupChats.isError ? "error" : status(waiting), read: Math.min(10, (oneToOne.data?.rows.length ?? 0) + (groupChats.data?.rows.length ?? 0)), note: "chats" },
    { key: "copilot", label: "Team activity", status: status(signals.teams, plan.priorities || plan.meetings), read: (signals.teams.data ?? []).reduce((sum, feed) => sum + feed.posts.length, 0), note: `posts in ${(signals.teams.data ?? []).length} teams`, countWith: ["teams"] },
    { key: "onedrive", label: "Files", status: filesLoading ? "loading" : signals.mail.isError && signals.teams.isError ? "error" : "ready", read: fileItems.length, note: "files found" },
    { key: "todo", status: status(openTasks), read: openTasks.data?.length ?? 0, note: "open tasks" },
    { key: "onenote", status: status(notebooks), read: notebooks.data?.length ?? 0, note: "pages" },
  ];
  const tileTarget = (id: "urgent" | "inbox" | "agenda" | "teams" | "project" | "oof"): ComponentId => ({ urgent: "today", inbox: "email", agenda: "calendar", teams: "teams", project: "projects", oof: "calendar" } as const)[id];
  const [emailView, setEmailView] = useState<"reply" | "unread" | "flagged" | "recent">("reply");
  const [askView, setAskView] = useState<"chat" | "create">("chat");
  const [calendarChip, setCalendarChip] = useState<"week" | "away">("week");
  const searchOff = (group: SearchGroup) => <p className="rounded-2xl border border-dashed border-border p-5 text-sm text-muted-foreground">{SEARCH_LABELS[group]} is turned off in Customize.</p>;
  const chips = <T extends string>(label: string, value: T, set: (value: T) => void, options: [T, string][]) => <div role="group" aria-label={label} className="mb-4 flex flex-wrap gap-2">
    {options.map(([option, text]) => <Button key={option} size="sm" variant={value === option ? "default" : "outline"} aria-pressed={value === option} onClick={() => set(option)} className="rounded-full">{text}</Button>)}
  </div>;

  const sectionCollapse: SectionCollapse = {
    isCollapsed: id => preferences.collapsed.includes(id as ComponentId),
    toggle: id => {
      const key = id as ComponentId;
      update({ ...preferences, collapsed: preferences.collapsed.includes(key) ? preferences.collapsed.filter(item => item !== key) : [...preferences.collapsed, key] });
    },
  };
  const components: Record<ComponentId, ReactNode> = {
    today: <SectionCard id="today" title="Today" icon={Sun} tone="tone-today" dataLabel="Today" mySite={mySite} actions={attention} actionsLabel="Priorities"
      actionsContent={<TodayBriefing items={attention} coverage={coverage} helpers={helpers} mySite={mySite} />}
      actionsIntro={<p>Everything that needs you, from your calendar, Outlook, Teams chats and team posts, To Do, OneNote and OneDrive, ranked in one list. Click any item to open the original.</p>}>
      <SignalTiles signals={signals} waitingCount={waiting.data ? waiting.data.length : waiting.isError ? "—" : undefined} tomorrowCount={calendarId && events.data !== undefined ? tomorrowCount : undefined} calendarPartial={calendarPartial} calendarError={calendarError || calendars.isSuccess && !calendarId} onOpen={id => openComponent(tileTarget(id))} />
      <div className="mt-6 grid gap-6 xl:grid-cols-[1.2fr_1fr]">
        <section aria-labelledby="plan-heading">
          <h3 id="plan-heading" className="font-heading text-lg font-bold">Your day</h3>
          {calendarLoading ? <p className="mt-3 text-sm text-muted-foreground">Building your plan…</p> : dayPlan.length ? <ol className="day-plan mt-3 space-y-2">{dayPlan.map((step, index) => <li key={index} className={`plan-step plan-${step.tone} flex gap-3 rounded-xl border border-border bg-card p-3`}>
            <span className="plan-dot mt-1.5 size-2.5 shrink-0 rounded-full" aria-hidden="true" />
            <div className="min-w-0">{step.time && <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">{step.time}</p>}<p className="break-words text-sm"><RichText text={step.text} /></p></div>
          </li>)}</ol> : <p className="mt-3 rounded-xl border border-dashed border-border p-4 text-sm text-muted-foreground">Nothing scheduled or due for the rest of today.</p>}
          <p className="mt-2 text-xs text-muted-foreground">Built from your calendar, tasks and the last refresh.</p>
        </section>
        {plan.priorities ? <ContextSection title="Needs attention today" kind="urgent" query={signals.priorities} urgent emptyTitle="Nothing needs attention today" emptyDescription="No urgent team posts, @mentions of you, or deadlines due today were found in your project teams.">
          <div className="px-5 pb-5"><TodoTasks dueOnly heading="Overdue or due today in Microsoft To Do" /></div>
        </ContextSection> : searchOff("priorities")}
      </div>
    </SectionCard>,
    ask: <SectionCard id="ask" title="Quick actions" icon={Sparkles} tone="tone-ask" dataLabel="Do" mySite={mySite} actions={[]} actionsLabel="History"
      actionsContent={<HistoryPanel mySite={mySite} />}>
      {chips("Quick actions view", askView, setAskView, [["chat", "Email, meeting, task or note"], ["create", "Create a document"]])}
      <div hidden={askView !== "chat"}><QuickActions calendarId={calendarId} /></div>
      <div hidden={askView !== "create"} className="space-y-3">
        <p className="text-sm text-muted-foreground">Make a Word, Excel or PDF file from an outline you write, then download it.</p>
        <CreateForm mySite={mySite} />
      </div>
    </SectionCard>,
    projects: <SectionCard id="projects" title="Projects" icon={BriefcaseBusiness} tone="tone-projects" dataLabel="Projects" mySite={mySite} actions={[...projectActionList, ...projectReports]}
      actionsIntro={<p>Actions for your projects, most in need of attention first. Nothing runs until you select it.</p>}>
      <div className="mb-3 flex justify-end"><ProjectTeamPicker owner={owner} /></div>
      <ProjectBoard owner={owner} views={projectViews} query={signals.priorities} teams={joinedTeams} searchOn={plan.priorities} mySite={mySite} />
      {!plan.priorities && searchOff("priorities")}
    </SectionCard>,
    calendar: <SectionCard id="calendar" title="Calendar" icon={CalendarDays} tone="tone-calendar" dataLabel="Schedule" mySite={mySite} actions={actions.calendar}>
      {chips("Calendar view", calendarChip, setCalendarChip, [["week", "Schedule"], ["away", "Who's away"]])}
      {calendarChip === "week" ? <div className="grid gap-6 xl:grid-cols-[1.6fr_1fr]">{calendarComponent}{agendaComponent}</div>
        : plan.collaboration ? <ContextSection title="Who's away" kind="oof" query={signals.collaboration} emptyTitle="No one is away this week or next" /> : searchOff("collaboration")}
    </SectionCard>,
    email: <SectionCard id="email" title="Email" icon={Mail} tone="tone-email" dataLabel="Mail" mySite={mySite} actions={actions.email}>
      {chips("Email view", emailView, value => { setEmailView(value); if (value !== "reply") setMailFilter(value); }, [["reply", "Needs reply"], ["unread", "Unread"], ["flagged", "Flagged"], ["recent", "Recent"]])}
      {emailView === "reply" ? (plan.replies ? <ContextSection title="Emails needing a reply" kind="reply" query={signals.replies} emptyTitle="No emails waiting on a reply" emptyDescription="Shows emails from the last 14 days that ask you something and that you haven't answered in the same conversation." /> : searchOff("replies")) : inboxComponent}
    </SectionCard>,
    teams: <SectionCard id="teams" title="Teams" icon={MessagesSquare} tone="tone-teams" dataLabel="Chats" mySite={mySite} actions={actions.teams}>
      <WaitingChats query={waiting} />
      <div className="mt-5"><TeamsSection currentUserId={signals.profile.data?.id} /></div>
    </SectionCard>,
    files: <SectionCard id="files" title="Files" icon={FolderOpen} tone="tone-files" dataLabel="Files" mySite={mySite} actions={actions.files}
      actionsIntro={<p>Files that need you, most important first. Open one, add a review task, or ask your Copilot to summarize it.</p>}>
      <FilesPanel files={fileItems} loading={filesLoading} oneDriveUrl={oneDriveUrl} />
    </SectionCard>,
    onenote: <SectionCard id="onenote" title="OneNote" icon={NotebookPen} tone="tone-onenote" dataLabel="Notes" mySite={mySite} actions={actions.onenote}>
      <OneNoteSection />
    </SectionCard>,
    planner: <SectionCard id="planner" title="Planner" icon={ClipboardList} tone="tone-planner" dataLabel="Tasks" mySite={mySite} actions={[]} actionsLabel="Plans"
      actionsContent={<PlannerPlans />}>
      <PlannerSection />
    </SectionCard>,
    plugins: <SectionCard id="plugins" title="Plugins" icon={Puzzle} tone="tone-plugins" dataLabel="Sources" mySite={mySite} actions={[]} actionsLabel="Links"
      actionsContent={<ul className="grid gap-2 sm:grid-cols-2">{[
        ["GitHub notifications", "https://github.com/notifications"], ["GitHub pull requests", "https://github.com/pulls"],
        ["GitHub issues", "https://github.com/issues"], ["Google Drive: Recent", "https://drive.google.com/drive/recent"],
        ["Google Drive: Shared with me", "https://drive.google.com/drive/shared-with-me"], ["Box: Recents", "https://app.box.com/recents"], ["Google Calendar: This week", "https://calendar.google.com/calendar/u/0/r/week"],
      ].map(([label, url]) => <li key={url}><a href={url} target="_blank" rel="noopener noreferrer" className="flex items-center justify-between gap-2 rounded-xl border border-border p-3 text-sm font-semibold text-primary hover:bg-muted focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-ring">{label}<ExternalLink aria-hidden="true" className="size-4 shrink-0" /><span className="sr-only"> (opens in a new tab)</span></a></li>)}</ul>}>
      <PluginsPanel order={preferences.pluginOrder} onOrderChange={order => update({ ...preferences, pluginOrder: order })} />
    </SectionCard>,
    insights: <SectionCard id="insights" title="Time insights" icon={ChartPie} tone="tone-insights" dataLabel="Insights" mySite={mySite} actions={actions.insights}>
      <div className="space-y-5">
        <div className="grid gap-5 xl:grid-cols-2">
        <CalendarInsights section="density" hasData={events.data !== undefined} fetching={events.isFetching || calendars.isFetching} rows={rows} days={days} page={offset / PAGE_SIZE + 1} partial={calendarPartial} pending={calendarLoading} error={calendarError || calendars.isSuccess && !calendarId} retry={resetCalendar} view={preferences.densityView} onViewChange={view => { if (view === "bar" || view === "line" || view === "table") update({ ...preferences, densityView: view }); }} />
        <CalendarInsights section="allocation" hasData={events.data !== undefined} fetching={events.isFetching || calendars.isFetching} rows={rows} days={days} page={offset / PAGE_SIZE + 1} partial={calendarPartial} pending={calendarLoading} error={calendarError || calendars.isSuccess && !calendarId} retry={resetCalendar} view={preferences.allocationView} onViewChange={view => { if (view === "bar" || view === "pie" || view === "table") update({ ...preferences, allocationView: view }); }} />
        </div>
        <TimeTargets shares={shares} targets={targets} suggested={!preferences.targetsSet} onChange={next => update({ ...preferences, targets: next, targetsSet: true })} onReset={() => update({ ...preferences, targets: {}, targetsSet: false })} />
      </div>
    </SectionCard>,
  };

  return <div className="command-center mx-auto max-w-[1440px] space-y-7 px-4 py-8 sm:px-8">
    <header className="command-header command-banner flex min-h-56 flex-col items-end justify-between gap-5 rounded-3xl p-6 text-right sm:min-h-64 sm:p-8" style={{ "--banner-image": `url(${bannerImage})` } as CSSProperties}>
      <div className="flex items-center justify-end gap-4"><div className="banner-title"><h1 className="font-heading text-3xl font-bold tracking-tight sm:text-4xl">Mission Control Interstellar</h1>{signals.profile.data?.givenName && <p className="mt-1 text-lg opacity-90">Hello, {signals.profile.data.givenName}.</p>}{signals.profile.isError && <p className="mt-2 text-sm">Your profile could not be refreshed.</p>}</div></div>
      <div className="flex flex-wrap justify-end gap-2">
        <HistoryButton mySite={mySite} />
        <ConnectionsButton />
        <DashboardCustomizer triggerId="customize-components" preferences={preferences} update={update} storage={storage} identityVerified={identityVerified} />
        <Button variant="secondary" onClick={() => update({ ...preferences, theme: dark ? "light" : "dark" })} className="h-auto whitespace-normal rounded-full py-2">{dark ? <Sun aria-hidden="true" /> : <Moon aria-hidden="true" />} {dark ? "Light mode" : "Dark mode"}</Button>
        <RefreshConfirm searches={[]} owner={profile.data?.id} onConfirm={() => void refreshAll()} disabled={anyFetching} label={manualRefreshing ? "Refreshing…" : anyFetching ? "Loading…" : "Refresh"} />
      </div>
    </header>
    {refreshResult && !anyFetching && refreshFailed && <p role="alert" className="sr-only">Some sections could not be refreshed. Last successfully loaded data is shown where available; review section errors and retry with Refresh.</p>}
    <WorkSearchStatus signals={signals} />
    <SectionCollapseContext.Provider value={sectionCollapse}>
    <DashboardLayout key={focusOwner} preferences={preferences} update={update} components={components} focusOnly={false}
      emptyState={<EmptyState icon={CalendarDays} title="No dashboard components selected" action={<DashboardCustomizer preferences={preferences} update={update} storage={storage} identityVerified={identityVerified} />} />} />
    </SectionCollapseContext.Provider>
    <footer className="flex justify-center border-t border-border pt-5" data-pdf-exclude><HowItWorks trigger="link" /></footer>
  </div>;
}

/** Share of scheduled (timed, not declined) hours per calendar category in the loaded week. */
function categoryShares(rows: CalendarEvent[]) {
  const totals = new Map<string, number>(); let all = 0;
  for (const event of rows) {
    const start = eventStart(event), end = eventEnd(event);
    if (!start || !end || end <= start || event.isAllDay || event.responseType === "declined") continue;
    const minutes = (end.getTime() - start.getTime()) / 60000, labels = (event.categories ?? []).filter(Boolean);
    all += minutes;
    for (const label of labels.length ? labels : ["Uncategorized"]) totals.set(label, (totals.get(label) ?? 0) + minutes / Math.max(1, labels.length));
  }
  return [...totals.entries()].map(([name, minutes]) => ({ name, share: all ? minutes / all * 100 : 0 })).sort((a, b) => b.share - a.share);
}

function TimeTargets({ shares, targets, suggested, onChange, onReset }: { shares: { name: string; share: number }[]; targets: Record<string, number>; suggested: boolean; onChange: (targets: Record<string, number>) => void; onReset: () => void }) {
  const names = [...new Set([...Object.keys(targets), ...shares.map(share => share.name).filter(name => name !== "Uncategorized")])];
  if (!names.length) return null;
  const total = names.reduce((sum, name) => sum + (targets[name] ?? 0), 0);
  return <section aria-labelledby="targets-heading" className="rounded-2xl border border-border p-4">
    <div className="flex flex-wrap items-start justify-between gap-3">
      <div>
        <h3 id="targets-heading" className="font-semibold">Targets</h3>
        <p className="mt-1 text-sm text-muted-foreground">{suggested
          ? "Suggested starting targets, based on your category names. Change any number to make them yours."
          : "Your targets for how your scheduled time should be split across your categories."} The Actions tab suggests time blocks when you're below a target.</p>
      </div>
      {!suggested && <Button variant="outline" size="sm" className="rounded-full" onClick={onReset}>Use suggested targets</Button>}
    </div>
    <ul className="mt-3 grid gap-x-6 gap-y-2 sm:grid-cols-2">{names.map(name => {
      const share = shares.find(item => item.name === name)?.share ?? 0, target = targets[name];
      const below = target !== undefined && share + 0.5 < target;
      return <li key={name} className="flex flex-wrap items-center gap-3">
        <label htmlFor={`target-${name}`} className="min-w-32 flex-1 break-words text-sm font-medium">{name}<span className={cn("block text-xs font-normal", below ? "font-semibold text-amber-700 dark:text-amber-300" : "text-muted-foreground")}>Now {Math.round(share)}%{below ? " · below target" : ""}</span></label>
        <Input id={`target-${name}`} type="number" min={0} max={100} inputMode="numeric" className="w-24" value={target ?? ""} placeholder="—"
          onChange={event => { const value = event.target.value === "" ? undefined : Math.max(0, Math.min(100, Number(event.target.value))); const next = { ...targets }; if (value === undefined || Number.isNaN(value)) delete next[name]; else next[name] = value; onChange(next); }} />
        <span className="text-sm text-muted-foreground" aria-hidden="true">%</span>
      </li>;
    })}</ul>
    <p className={cn("mt-3 text-xs", total === 100 ? "text-muted-foreground" : "font-semibold text-amber-700 dark:text-amber-300")}>Targets add up to {total}%{total === 100 ? "." : " — they don't need to total 100%, but it helps."}</p>
  </section>;
}
