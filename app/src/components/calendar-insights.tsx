import { useQuery } from "@tanstack/react-query";
import { loadCalendarCategories } from "@/lib/command-center-data";
import { Card } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { EmptyState, ErrorState, LoadingState } from "@/components/states";
import { CalendarDays } from "lucide-react";
import { hasAttendees, isAcceptedMeeting, type CalendarEvent } from "@/lib/command-center-evidence";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import type { DensityView, AllocationView } from "@/lib/dashboard-preferences";

const UNCATEGORIZED = "Uncategorized";
const UNCATEGORIZED_COLOR = "#8a8886";
// Used only when the user's own category colors cannot be read.
const FALLBACK_COLORS = ["#0078d4", "#8764b8", "#13a10e", "#ca5010", "#c239b3", "#038387", "#a4262c", "#5c2e91", "#c19c00", "#004e8c"];
function instant(value?: string) {
  if (!value || !/(Z|[+-]\d{2}:\d{2})$/i.test(value)) return undefined;
  const date = new Date(value); return Number.isNaN(date.getTime()) ? undefined : date;
}
export function CalendarInsights({ rows, days, page, partial, pending, error, retry, section, view, onViewChange, hasData, fetching }: { rows: CalendarEvent[]; days: Date[]; page: number; partial: boolean; pending: boolean; error: boolean; retry: () => void; section: "density" | "allocation"; view: DensityView | AllocationView; onViewChange: (view: DensityView | AllocationView) => void; hasData: boolean; fetching: boolean }) {
  const categoryList = useQuery({ queryKey: ["calendar-categories"], queryFn: loadCalendarCategories, retry: false, enabled: section === "allocation" });
  const slots = rows.map(event => ({ event, start: instant(event.startWithTimeZone ?? event.start), end: instant(event.endWithTimeZone ?? event.end) })).filter(slot => !slot.event.isAllDay && slot.event.responseType !== "declined" && slot.start && slot.end);
  const meetings = slots.filter(slot => isAcceptedMeeting(slot.event) && slot.start && slot.end && slot.end > slot.start);
  const loads = days.map(day => {
    const next = new Date(day); next.setDate(day.getDate() + 1);
    const onDay = slots.filter(slot => slot.start && slot.start >= day && slot.start < next);
    return { day, count: meetings.filter(slot => slot.start && slot.start >= day && slot.start < next).length, solo: onDay.filter(slot => !hasAttendees(slot.event) && slot.end && slot.start && slot.end > slot.start).length };
  });
  const invalidMeetings = rows.filter(event => isAcceptedMeeting(event) && (!instant(event.startWithTimeZone ?? event.start) || !instant(event.endWithTimeZone ?? event.end) || Date.parse(event.endWithTimeZone ?? event.end ?? "") <= Date.parse(event.startWithTimeZone ?? event.start ?? ""))).length;
  // Categories and their names come from the user's own Outlook calendar.
  const known = new Map((categoryList.data?.categories ?? []).map(category => [category.name.toLowerCase(), category]));
  const totals = new Map<string, { group: string; minutes: number; color?: string }>();
  let excluded = rows.filter(event => !instant(event.startWithTimeZone ?? event.start) || !instant(event.endWithTimeZone ?? event.end)).length;
  let multiCategory = 0;
  const endRange = new Date(days[6]); endRange.setDate(endRange.getDate() + 1);
  for (const slot of slots) {
    if (!slot.start || !slot.end || slot.end <= slot.start) { excluded += 1; continue; }
    const duration = Math.max(0, Math.min(slot.end.getTime(), endRange.getTime()) - Math.max(slot.start.getTime(), days[0].getTime())) / 60000;
    const labels = [...new Set((slot.event.categories ?? []).map(label => label.trim()).filter(Boolean))];
    if (labels.length > 1) multiCategory += 1;
    // An event with several categories shares its time equally between them.
    for (const label of labels.length ? labels : [UNCATEGORIZED]) {
      const official = known.get(label.toLowerCase());
      const name = official?.name ?? label;
      const entry = totals.get(name) ?? { group: name, minutes: 0, color: name === UNCATEGORIZED ? UNCATEGORIZED_COLOR : official?.color };
      entry.minutes += duration / Math.max(1, labels.length);
      totals.set(name, entry);
    }
  }
  const sortedGroups = [...totals.values()].filter(row => row.minutes > 0).sort((a, b) => (a.group === UNCATEGORIZED ? 1 : 0) - (b.group === UNCATEGORIZED ? 1 : 0) || b.minutes - a.minutes || a.group.localeCompare(b.group));
  let fallbackIndex = 0;
  const minutes = sortedGroups.map(row => ({ ...row, color: row.color ?? FALLBACK_COLORS[fallbackIndex++ % FALLBACK_COLORS.length] }));
  const total = minutes.reduce((sum, row) => sum + row.minutes, 0);
  const maximum = Math.max(1, ...loads.map(row => row.count));
  const maxMinutes = Math.max(1, ...minutes.map(row => row.minutes));
  const scope = `${days[0].toLocaleDateString([], { month: "short", day: "numeric" })} – ${days[6].toLocaleDateString([], { month: "short", day: "numeric" })}${partial ? ` · Calendar page ${page} only` : ""}`;
  const state = pending ? <LoadingState className="mt-5" /> : error ? <div role="alert"><ErrorState title="Calendar insights could not be refreshed" description={hasData ? "Showing the last successfully loaded calendar data." : undefined} action={<Button variant="outline" disabled={fetching} onClick={retry}>Retry calendar</Button>} /></div> : undefined;
  const options = section === "density" ? ["bar", "line", "table"] as const : ["bar", "pie", "table"] as const;
  const viewPicker = <div className="mt-4 flex flex-wrap gap-2" role="group" aria-label={`${section === "density" ? "Meeting density" : "Time allocation"} view`}>
    {options.map(option => <Button key={option} size="sm" variant={view === option ? "default" : "outline"} aria-pressed={view === option} className="h-auto whitespace-normal" onClick={() => onViewChange(option)}>{option === "table" ? "Table" : option === "bar" ? "Bar chart" : option === "line" ? "Line chart" : "Pie chart"}</Button>)}
  </div>;
  const tableClass = view === "table" ? "mt-4" : "sr-only";
  if (section === "density") return <Card className="min-w-0 p-5"><section aria-labelledby="load-heading" aria-busy={fetching || pending}>
    <h2 id="load-heading" className="font-heading text-xl font-bold">Accepted meeting density</h2>
    <p className="mt-2 text-sm text-muted-foreground">{scope}. Accepted timed meetings with attendees only; tentative, unanswered, declined and organizer responses are excluded. Solo blocks do not contribute to density.</p>
    {viewPicker}
    {state}
    {!pending && (!error || hasData) && <>
      <div className={tableClass}><Table className="w-full table-fixed">
        <TableHeader><TableRow><TableHead className="h-auto whitespace-normal align-top">Day</TableHead><TableHead className="h-auto whitespace-normal align-top">Accepted meetings</TableHead><TableHead className="h-auto whitespace-normal align-top">Solo blocks</TableHead></TableRow></TableHeader>
        <TableBody>{loads.map(row => <TableRow key={row.day.toISOString()}><TableCell className="whitespace-normal align-top">{row.day.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" })}</TableCell><TableCell className="whitespace-normal align-top">{row.count}{row.count >= 7 ? " (Heavy)" : row.count === 2 ? " (Light)" : ""}</TableCell><TableCell className="whitespace-normal align-top">{row.solo}</TableCell></TableRow>)}</TableBody>
      </Table></div>
      {view !== "table" && <figure className="mt-5"><div role="region" aria-label="Meeting density graphic" tabIndex={0} className="overflow-x-auto rounded-lg focus-visible:outline-2 focus-visible:outline-ring">
        <svg className="insight-chart block w-full min-w-[400px]" aria-hidden="true" focusable="false" viewBox="0 0 560 270" role="img" aria-labelledby="density-chart-title density-chart-description">
          <title id="density-chart-title">Accepted meetings by day</title><desc id="density-chart-description">Counts of accepted timed meetings with attendees, for the displayed calendar scope. Equivalent counts and solo blocks are available in the adjacent table.</desc>
          {[0, maximum].map(tick => <g key={tick}><line className="insight-axis" x1="45" x2="540" y1={220 - tick / maximum * 170} y2={220 - tick / maximum * 170} /><text x="36" y={224 - tick / maximum * 170} textAnchor="end" fontSize="12">{tick}</text></g>)}
          <text x="45" y="20" fontSize="12">Accepted meetings</text>
          {view === "line" && <polyline className="insight-line" points={loads.map((row, index) => `${80 + index * 72},${220 - row.count / maximum * 170}`).join(" ")} fill="none" strokeWidth="3" />}
          {loads.map((row, index) => <g key={row.day.toISOString()}>
            {view === "bar" ? <rect className={`insight-mark ${row.count >= 7 ? "heavy" : row.count === 2 ? "light" : ""}`} x={57 + index * 72} y={220 - row.count / maximum * 170} width="46" height={row.count / maximum * 170} rx="3" /> : <circle className={`insight-mark ${row.count >= 7 ? "heavy" : row.count === 2 ? "light" : ""}`} cx={80 + index * 72} cy={220 - row.count / maximum * 170} r="5" />}
            <text x={80 + index * 72} y={210 - row.count / maximum * 170} textAnchor="middle" fontSize="13" fontWeight="600">{row.count}{row.count >= 7 ? " (H)" : row.count === 2 ? " (L)" : ""}</text>
            <text x={80 + index * 72} y="245" textAnchor="middle" fontSize="12">{row.day.toLocaleDateString([], { weekday: "short", day: "numeric" })}</text>
          </g>)}
        </svg>
      </div><figcaption className="mt-2 text-xs text-muted-foreground">H: Heavy, 7+ accepted meetings · L: Light, 2 accepted meetings. Overlaps count separately. Select Table to compare solo blocks.</figcaption></figure>}
      {invalidMeetings > 0 && <p className="mt-2 text-xs text-muted-foreground">{invalidMeetings} accepted meeting{invalidMeetings === 1 ? "" : "s"} with missing or invalid times excluded.</p>}
    </>}
  </section></Card>;

  let angle = -Math.PI / 2;
  const wedges = minutes.map((row, index) => {
    const begin = angle; angle += total ? row.minutes / total * Math.PI * 2 : 0;
    return { ...row, index, begin, end: angle };
  });
  return <Card className="min-w-0 p-5"><section aria-labelledby="allocation-heading" aria-busy={fetching || pending}>
    <h2 id="allocation-heading" className="font-heading text-xl font-bold">Time allocation</h2>
    <p className="mt-2 text-sm text-muted-foreground">{scope}. Scheduled timed-event duration, including tentative and unanswered items and solo blocks, excluding declined items. Overlaps each contribute their duration; this is not your entire workweek.</p>
    {viewPicker}
    {state}
    {!pending && (!error || hasData) && (total > 0 ? <>
      <p className="mt-3 text-sm">Categories and names come from your Outlook calendar.{categoryList.data && !categoryList.data.colorsAvailable ? " Category colors are assigned by the app, so they may differ from Outlook's." : ""}{multiCategory > 0 ? ` ${multiCategory} event${multiCategory === 1 ? " has" : "s have"} several categories; their time is shared equally.` : ""}</p>
      {excluded > 0 && <p className="mt-2 text-xs text-muted-foreground">{excluded} event{excluded === 1 ? "" : "s"} with missing or invalid times excluded.</p>}
      <div className={tableClass}><Table className="w-full table-fixed">
        <TableHeader><TableRow><TableHead className="h-auto w-1/2 whitespace-normal align-top">Category</TableHead><TableHead className="h-auto whitespace-normal align-top">Hours</TableHead><TableHead className="h-auto whitespace-normal align-top">Share</TableHead></TableRow></TableHeader>
        <TableBody>{minutes.map(row => <TableRow key={row.group}><TableCell className="whitespace-normal align-top"><span className="flex items-center gap-2"><span aria-hidden="true" className="inline-block size-3 shrink-0 rounded-sm" style={{ backgroundColor: row.color }} />{row.group}</span></TableCell><TableCell className="whitespace-normal align-top">{(row.minutes / 60).toFixed(1)}</TableCell><TableCell className="whitespace-normal align-top">{(row.minutes / total * 100).toFixed(1)}%</TableCell></TableRow>)}</TableBody>
      </Table></div>
      {view !== "table" && <figure className="mt-5"><div role="region" aria-label="Scheduled time allocation graphic" tabIndex={0} className="overflow-x-auto rounded-lg focus-visible:outline-2 focus-visible:outline-ring">
        <svg className="insight-chart block min-w-[400px] w-full" viewBox={view === "pie" ? `0 0 560 ${Math.max(340, 40 + minutes.length * 50)}` : `0 0 560 ${Math.max(300, 50 + minutes.length * 39)}`} aria-hidden="true" focusable="false" role="img" aria-labelledby="allocation-chart-title allocation-chart-description">
          <title id="allocation-chart-title">Scheduled duration by calendar category</title><desc id="allocation-chart-description">Hours and percentage shares for the displayed calendar scope, available as an equivalent table.</desc>
          {view === "pie" ? <>
            {wedges.filter(row => row.minutes > 0).map(row => {
              const x1 = 130 + Math.cos(row.begin) * 110, y1 = 155 + Math.sin(row.begin) * 110;
              const x2 = 130 + Math.cos(row.end) * 110, y2 = 155 + Math.sin(row.end) * 110;
              const style = { fill: row.color };
              return row.minutes === total ? <circle key={row.group} className="insight-segment" cx="130" cy="155" r="110" style={style} /> : <path key={row.group} className="insight-segment" d={`M 130 155 L ${x1} ${y1} A 110 110 0 ${row.end - row.begin > Math.PI ? 1 : 0} 1 ${x2} ${y2} Z`} style={style} />;
            })}
            {minutes.map((row, index) => <g key={row.group}><rect className="insight-segment" x="260" y={25 + index * 50} width="12" height="12" style={{ fill: row.color }} /><text x="282" y={36 + index * 50} fontSize="13">{row.group}</text><text x="282" y={53 + index * 50} fontSize="12">{(row.minutes / 60).toFixed(1)} h · {(row.minutes / total * 100).toFixed(1)}%</text></g>)}
          </> : <>
            <text x="170" y="18" fontSize="12">Scheduled hours</text>
            {minutes.map((row, index) => <g key={row.group}><text x="160" y={47 + index * 39} textAnchor="end" fontSize="13">{row.group}</text><rect className="insight-segment" x="170" y={31 + index * 39} width={row.minutes / maxMinutes * 300} height="24" rx="3" style={{ fill: row.color }} /><text x={180 + row.minutes / maxMinutes * 300} y={47 + index * 39} fontSize="12">{(row.minutes / 60).toFixed(1)} h</text></g>)}
          </>}
        </svg>
      </div><figcaption className="mt-2 text-xs text-muted-foreground">Hours of scheduled events in the displayed scope, not free or available time.</figcaption></figure>}
    </> : <EmptyState icon={CalendarDays} title="No timed events available for allocation" action={<Button variant="outline" disabled={fetching} onClick={retry}>Reload calendar</Button>} className="py-8" />)}
  </section></Card>;
}
