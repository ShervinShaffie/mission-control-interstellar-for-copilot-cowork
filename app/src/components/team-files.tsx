import { useState } from "react";
import type { UseQueryResult } from "@tanstack/react-query";
import { useQuery } from "@tanstack/react-query";
import { ExternalLink, FolderOpen, Settings2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { EmptyState, ErrorState, LoadingState } from "@/components/states";
import { loadJoinedTeamList, type TeamFeed } from "@/lib/direct-signals";
import { readTeamChoice, writeTeamChoice } from "@/lib/team-choice";
import { safeUrl } from "@/lib/safe-url";

/** Choose which teams count as projects (stored in this browser). Applies on the next Refresh. */
export function ProjectTeamPicker({ owner }: { owner?: string }) {
  const [open, setOpen] = useState(false);
  const teams = useQuery({ queryKey: ["teams-list"], queryFn: loadJoinedTeamList, enabled: open, retry: false });
  const [chosen, setChosen] = useState<string[]>(() => readTeamChoice(owner));
  const [saved, setSaved] = useState(false);
  if (!open) return <Button variant="ghost" size="sm" className="rounded-full" onClick={() => setOpen(true)}><Settings2 aria-hidden="true" />Choose project teams</Button>;
  return <div className="rounded-2xl border p-3 text-sm">
    <p className="font-semibold">Which teams are your projects?</p>
    <p className="text-xs text-muted-foreground">Pick up to 12. With none picked, the app uses up to 8 teams, project-sounding names first (launch, rollout, readiness…).</p>
    {teams.isPending ? <LoadingState rows={2} /> : teams.isError ? <p className="mt-2 text-destructive">Your teams couldn't be loaded.</p>
      : <ul className="mt-2 grid gap-1 sm:grid-cols-2">{(teams.data ?? []).map(team => <li key={team.id}><label className="flex items-center gap-2"><input type="checkbox" checked={chosen.includes(team.id)} disabled={!chosen.includes(team.id) && chosen.length >= 12}
        onChange={event => { setSaved(false); setChosen(list => event.target.checked ? [...list, team.id] : list.filter(id => id !== team.id)); }} />{team.name}</label></li>)}</ul>}
    <div className="mt-2 flex flex-wrap items-center gap-2">
      <Button size="sm" className="rounded-full" onClick={() => { writeTeamChoice(owner, chosen); setSaved(true); }}>Save</Button>
      <Button size="sm" variant="outline" className="rounded-full" onClick={() => { setChosen([]); writeTeamChoice(owner, []); setSaved(true); }}>Use automatic</Button>
      <Button size="sm" variant="ghost" className="rounded-full" onClick={() => setOpen(false)}>Close</Button>
      {saved && <span role="status" className="text-xs">Saved — select Refresh to update Projects.</span>}
    </div>
  </div>;
}

/** Files folders of each project team's channels (from Teams, free). */
export function TeamFilesSection({ query, owner }: { query: UseQueryResult<TeamFeed[]>; owner?: string }) {
  const feeds = (query.data ?? []).filter(feed => feed.channels.some(channel => channel.filesUrl));
  return <section aria-label="Team files" className="space-y-3">
    <p className="text-sm text-muted-foreground">Each project team's files, by channel. These open in SharePoint.</p>
    {query.isPending ? <LoadingState rows={3} /> : query.isError ? <div role="alert"><ErrorState title="Your teams couldn't be read" action={<Button variant="outline" onClick={() => void query.refetch()}>Retry</Button>} /></div>
      : !feeds.length ? <EmptyState icon={FolderOpen} title="No team files found" />
      : <ul className="grid gap-3 sm:grid-cols-2">{feeds.map(feed => <li key={feed.id} className="rounded-2xl border p-3">
        <p className="font-semibold">{feed.name}</p>
        <ul className="mt-1 space-y-1">{feed.channels.filter(channel => safeUrl(channel.filesUrl)).map(channel => <li key={channel.id}>
          <a href={safeUrl(channel.filesUrl)} target="_blank" rel="noopener noreferrer" className="inline-flex items-center gap-1.5 text-sm text-primary underline-offset-2 hover:underline"><FolderOpen aria-hidden="true" className="size-3.5" />{channel.name}<ExternalLink aria-hidden="true" className="size-3" /><span className="sr-only"> (opens in a new tab)</span></a>
        </li>)}</ul>
      </li>)}</ul>}
    <ProjectTeamPicker owner={owner} />
  </section>;
}
