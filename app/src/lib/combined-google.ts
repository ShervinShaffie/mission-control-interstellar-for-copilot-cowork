import { useQuery } from "@tanstack/react-query";
import { loadGoogleCalendars, loadGoogleEvents } from "@/lib/plugin-sources";

/**
 * Google Calendar events for the Calendar section's window, from your main Google calendar. Uses the same cache
 * keys as the Plugins panel, so edits made in either place refresh both.
 */
export function useCombinedGoogle(start: Date, days: number, enabled: boolean) {
  const calendars = useQuery({ queryKey: ["plugins", "googlecalendar", "calendars"], queryFn: loadGoogleCalendars, enabled });
  const calendar = calendars.data?.find(cal => cal.isPrimary) ?? calendars.data?.[0];
  const end = new Date(start); end.setDate(start.getDate() + days);
  const events = useQuery({
    queryKey: ["plugins", "googlecalendar", "events", calendar?.id, start.toISOString(), "", days],
    queryFn: () => loadGoogleEvents(calendar!.id, start, end),
    enabled: enabled && Boolean(calendar?.id),
  });
  return {
    enabled, calendarId: calendar?.id, calendarName: calendar?.name,
    events: enabled ? events.data ?? [] : [],
    loading: enabled && (calendars.isPending || (Boolean(calendar?.id) && events.isPending)),
    failed: enabled && (calendars.isError || events.isError),
    noCalendar: enabled && calendars.isSuccess && !calendar,
  };
}
export type CombinedGoogle = ReturnType<typeof useCombinedGoogle>;
