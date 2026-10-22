import { useQuery } from "@tanstack/react-query";
import { loadInboxUnreadCount, unwrap } from "@/lib/command-center-data";
import type { SearchPlan } from "@/lib/search-plan";
import { Office365UsersService } from "../../generated/services/Office365UsersService";
import { awayFromMailTips, loadMail, loadTrendingFiles, loadTeamFeeds, loadUpcomingEvents, meetingsFromCalendar, prioritiesFromTeams, repliesFromMail, waitingOnOthers } from "@/lib/direct-signals";
import { readTeamChoice } from "@/lib/team-choice";

type Source = { isFetched: boolean; error: unknown; dataUpdatedAt: number };
/** A group computed from shared reads; it re-runs whenever one of its sources is refreshed. */
function useDerived<T>(settled: boolean, me: string, spec: { name: string; on: boolean; sources: Source[]; build: () => Promise<T> | T }) {
  return useQuery({
    retry: false, enabled: settled && spec.on && spec.sources.every(source => source.isFetched),
    queryKey: ["direct", spec.name, me, ...spec.sources.map(source => source.dataUpdatedAt)],
    queryFn: async () => { const failed = spec.sources.find(source => source.error); if (failed && spec.sources.every(source => source.error)) throw failed.error; return spec.build(); },
  });
}
export function useProfile() {
  return useQuery({ retry: false, queryKey: ["work-profile"], queryFn: async () => unwrap(await Office365UsersService.MyProfile_V2("id,givenName,displayName,mySite,mail,userPrincipalName"), "Profile") });
}
/**
 * Work signals from the direct connections only (Outlook, Teams, OneDrive, Office 365 Users). Free and stable:
 * three shared reads (team activity, mail, upcoming calendar) feed four derived groups.
 */
export function useWorkSignals(calendarId: string | undefined, ready: boolean, plan: SearchPlan, profile: ReturnType<typeof useProfile>) {
  const me = profile.data?.id ?? "", myEmail = (profile.data?.mail || profile.data?.userPrincipalName || "").toLowerCase();
  const settled = ready && !!me;
  const choice = readTeamChoice(me);
  const teams = useQuery({ retry: false, enabled: settled && (plan.priorities || plan.meetings), queryKey: ["direct", "teams", me, choice.join("|")], queryFn: () => loadTeamFeeds(me, choice) });
  const mail = useQuery({ retry: false, enabled: settled && (plan.replies || plan.collaboration || plan.meetings), queryKey: ["direct", "mail", me], queryFn: loadMail });
  const upcoming = useQuery({ retry: false, enabled: settled && !!calendarId && (plan.meetings || plan.collaboration), queryKey: ["direct", "upcoming", calendarId], queryFn: () => loadUpcomingEvents(calendarId!) });
  const derived = <T,>(name: string, on: boolean, sources: Source[], build: () => Promise<T> | T) => ({ name, on, sources, build });
  const priorities = useDerived(settled, me, derived("priorities", plan.priorities, [teams], () => { if (teams.error) throw teams.error; return prioritiesFromTeams(teams.data ?? [], me); }));
  const meetings = useDerived(settled, me, derived("meetings", plan.meetings, [upcoming, teams, mail], () => { if (upcoming.error) throw upcoming.error; return meetingsFromCalendar(upcoming.data ?? [], teams.data ?? [], mail.data); }));
  const replies = useDerived(settled, me, derived("replies", plan.replies, [mail], () => { if (mail.error) throw mail.error; return repliesFromMail(mail.data!, myEmail); }));
  const collaboration = useDerived(settled, me, derived("collaboration", plan.collaboration, [upcoming, mail], async () => {
    const away = upcoming.data ? await awayFromMailTips(upcoming.data, myEmail).catch(() => []) : [];
    return [...away, ...(mail.data ? waitingOnOthers(mail.data, myEmail) : [])];
  }));
  const unread = useQuery({ retry: false, queryKey: ["work-inbox-unread-count"], queryFn: loadInboxUnreadCount });
  // Optional: documents trending around you. A failure just leaves this source out.
  const trending = useQuery({ retry: false, enabled: settled, queryKey: ["direct", "trending", me], queryFn: loadTrendingFiles });
  return { priorities, meetings, replies, collaboration, unread, profile, teams, mail, upcoming, trending };
}
export type WorkSignals = ReturnType<typeof useWorkSignals>;
