import { CalendarDays } from "lucide-react";
import { InfoTip } from "@/components/info-tip";
import { EventActions, EventDetails } from "@/components/plugins-panel";
import type { CombinedGoogle } from "@/lib/combined-google";
import { hideEmails } from "@/lib/privacy";
import { safeUrl } from "@/lib/safe-url";
import type { GoogleEvent } from "@/lib/plugin-sources";

export interface BusyBlock { title: string; start: Date; end: Date }

function sameDay(a: Date, b: Date) { return a.toDateString() === b.toDateString(); }
function onDay(event: GoogleEvent, day: Date) {
  if (!event.start) return false;
  const next = new Date(day); next.setDate(day.getDate() + 1);
  const end = event.end ?? event.start;
  return event.start < next && (end > day || sameDay(event.start, day));
}
function timeText(event: GoogleEvent) {
  if (event.allDay) return "All day";
  if (!event.start) return "Time unavailable";
  const t = (d: Date) => d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  return event.end ? `${t(event.start)} – ${t(event.end)}` : t(event.start);
}

/** Google Calendar events for one agenda day, with overlaps against the Outlook events shown above them. */
export function AgendaGoogle({ google, day, outlook }: { google: CombinedGoogle; day: Date; outlook: BusyBlock[] }) {
  if (!google.enabled) return null;
  const items = google.events.filter(event => onDay(event, day));
  return <section aria-labelledby="agenda-google-heading" className="border-t border-border">
    <div className="flex items-center gap-1 px-4 pt-3">
      <CalendarDays aria-hidden="true" className="size-4 text-muted-foreground" />
      <h3 id="agenda-google-heading" className="text-sm font-bold">Google Calendar{google.calendarName ? ` · ${google.calendarName}` : ""}</h3>
      <InfoTip label="Google Calendar in the agenda" title="Google events on this calendar" points={[
        "Shows your main Google calendar. Pick another calendar in Plugins → Google Calendar.",
        "Google events can be edited, duplicated or deleted here; Outlook events stay read-only.",
        "Overlaps compare Google events with the Outlook events on this agenda page, ignoring declined and free time.",
        "If the same meeting is on both calendars, it appears twice.",
        "Repeating Google events are edited in Google Calendar.",
        "Turn this off with Include Google Calendar at the top of the Calendar section.",
      ]} />
    </div>
    {google.loading ? <p role="status" className="px-4 py-3 text-sm text-muted-foreground">Loading Google events…</p>
      : google.failed ? <p role="alert" className="px-4 py-3 text-sm text-red-700 dark:text-red-400">Couldn't load Google Calendar. Check its connection in Connections.</p>
      : google.noCalendar ? <p className="px-4 py-3 text-sm text-muted-foreground">No Google calendar was found.</p>
      : items.length === 0 ? <p className="px-4 py-3 text-sm text-muted-foreground">No Google events this day.</p>
      : <ul className="divide-y divide-border">{items.map(event => {
        const overlaps = !event.allDay && event.start && event.end
          ? outlook.filter(block => event.start! < block.end && event.end! > block.start).map(block => block.title) : [];
        const href = safeUrl(event.url);
        return <li key={event.id} className="flex items-start gap-3 px-4 py-2.5">
          <span className="w-24 shrink-0 text-xs font-semibold tabular-nums text-muted-foreground">{timeText(event)}</span>
          <div className="min-w-0 flex-1">
            {href ? <a href={href} target="_blank" rel="noopener noreferrer" className="break-words text-sm font-semibold text-primary underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-ring">{hideEmails(event.title)}<span className="sr-only"> (opens Google Calendar in a new tab)</span></a>
              : <span className="break-words text-sm font-semibold">{hideEmails(event.title)}</span>}
            {overlaps.length > 0 && <p className="text-xs font-medium text-amber-800 dark:text-amber-300">Overlaps with Outlook: {overlaps.join("; ")}</p>}
            <EventDetails event={event} />
          </div>
          {google.calendarId && <EventActions event={event} calendarId={google.calendarId} />}
        </li>;
      })}</ul>}
  </section>;
}
