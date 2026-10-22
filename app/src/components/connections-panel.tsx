import { useEffect, useRef, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { getContext } from "@microsoft/managed-apps/app";
import { AlertTriangle, CheckCircle2, CircleHelp, ExternalLink, KeyRound, Loader2, Plug, RefreshCw, ShieldAlert } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogDescription, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { CONNECTOR_NAMES, FREE_CHECKS, connectionsPageUrl, runCheck, useConnectionStates, type ConnectionKey, type Health } from "@/lib/connection-health";
import { cn } from "@/lib/utils";

const ROWS: { key: ConnectionKey; name: string; powers: string; paid?: boolean }[] = [
  { key: "outlook", name: "Office 365 Outlook", powers: "Calendar, email, drafts, invites, focus blocks, automatic replies" },
  { key: "teams", name: "Microsoft Teams", powers: "Chats waiting on you, your teams" },
  { key: "todo", name: "Microsoft To Do", powers: "Your tasks, adding tasks" },
  { key: "onenote", name: "OneNote", powers: "Recent pages and to-dos, new Quick Notes pages" },
  { key: "onedrive", name: "OneDrive for Business", powers: "My OneDrive in Files, your History copy" },
  { key: "users", name: "Office 365 Users", powers: "Your profile, first names, people search, trending documents" },
  { key: "planner", name: "Planner", powers: "Your Planner tasks and plans" },
  { key: "github", name: "GitHub", powers: "Plugins: your profile, repositories, open issues and pull requests" },
  { key: "googledrive", name: "Google Drive", powers: "Plugins: files and folders in your Drive" },
  { key: "box", name: "Box", powers: "Plugins: files and folders in your Box" },
  { key: "googlecalendar", name: "Google Calendar", powers: "Plugins: your Google calendars and upcoming events" },
];
export interface CopilotSearchStatus { group: string; label: string; on: boolean; state: "ok" | "failed" | "checking" | "waiting" | "off"; reason?: string }
export interface CopilotHealth { health: Health | "partial"; detail?: string; searches?: CopilotSearchStatus[]; retry?: () => void }
const LABEL: Record<Health | "partial", string> = { partial: "Partly working", ok: "Working", signin: "Needs sign-in", blocked: "Blocked by your organization", error: "Not working", unknown: "Not checked yet", checking: "Checking…" };

function Status({ health }: { health: Health | "partial" }) {
  const Icon = health === "ok" ? CheckCircle2 : health === "partial" ? AlertTriangle : health === "checking" ? Loader2 : health === "unknown" ? CircleHelp : health === "blocked" ? ShieldAlert : AlertTriangle;
  return <span className={cn("conn-status inline-flex items-center gap-1.5 rounded-full px-2.5 py-0.5 text-xs font-semibold", `conn-${health}`)}>
    <Icon aria-hidden="true" className={cn("size-3.5", health === "checking" && "animate-spin")} />{LABEL[health]}
  </span>;
}
function help(key: ConnectionKey, health: Health | "partial") {
  if (health === "signin") return <>This connection's sign-in has expired. Select <strong>Sign in again</strong> and follow the steps — it takes about a minute.</>;
  if (health === "blocked") return <>Your organization's data policy blocks this. Signing in again won't change it; the rest of the dashboard keeps working. Ask your IT team if you need it.</>;
  if (health === "error") return <>It didn't respond as expected. Select Try again in a minute. If it keeps failing, select <strong>Sign in again</strong>.</>;
  return null;
}

/** Step-by-step sign-in in the Power Apps Connections page (opens in a new tab), then an automatic re-check on return. */
function SignInSteps({ name, url, onDone }: { name: string; url: string; onDone: () => void }) {
  return <ol className="list-decimal space-y-1 rounded-xl border border-primary/30 bg-accent/40 p-3 pl-7 text-xs">
    <li>Select <a href={url} target="_blank" rel="noopener noreferrer" className="font-semibold text-primary underline-offset-2 hover:underline">Open your connections<ExternalLink aria-hidden="true" className="ml-1 inline size-3" /><span className="sr-only"> (opens in a new tab)</span></a> (Microsoft Power Apps, signed in as you).</li>
    <li>Find <strong>{name}</strong> (use Search at the top right). If it shows <em>Fix connection</em>, select it; otherwise select <strong>⋯ → Edit</strong>.</li>
    <li>Select <strong>Sign in</strong> (or <strong>Save</strong>) and sign in with your work account.</li>
    <li>If <strong>{name}</strong> isn't listed, select <strong>+ New connection</strong>, search for it, select <strong>Create</strong> and sign in.</li>
    <li>Come back here and select <button type="button" onClick={onDone} className="font-semibold text-primary underline underline-offset-2">Check again</button>. Free connections are re-checked automatically when you return.</li>
    <li>Still showing a problem? <button type="button" onClick={() => window.location.reload()} className="font-semibold text-primary underline underline-offset-2">Reload the app</button> once — the app picks up a renewed sign-in when it starts.</li>
  </ol>;
}

/** Connection status, with free checks for direct connections and a tiny paid check for SharePoint. */
export function ConnectionsPanel() {
  const states = useConnectionStates();
  const [busy, setBusy] = useState(false);
  const [signingIn, setSigningIn] = useState<ConnectionKey>();
  const context = useQuery({ queryKey: ["app-context"], queryFn: getContext, retry: false, staleTime: Infinity });
  const pageUrl = connectionsPageUrl(context.data?.app.environmentId);
  const awaiting = useRef<ConnectionKey | undefined>(undefined);
  useEffect(() => {
    // When the person returns from signing in, re-check a free connection automatically.
    const onFocus = () => { const key = awaiting.current; const check = key ? FREE_CHECKS[key] : undefined; if (key && check) { awaiting.current = undefined; void runCheck(key, check); } };
    window.addEventListener("focus", onFocus);
    return () => window.removeEventListener("focus", onFocus);
  }, []);
  useEffect(() => {
    // Opening the panel shows the current state: run the free checks once.
    void Promise.all(Object.entries(FREE_CHECKS).map(([key, check]) => runCheck(key as ConnectionKey, check!)));
  }, []);
  const healthOf = (key: ConnectionKey): Health | "partial" => states[key]?.health ?? "unknown";
  async function checkFree() {
    setBusy(true);
    await Promise.all(Object.entries(FREE_CHECKS).map(([key, check]) => runCheck(key as ConnectionKey, check!)));
    setBusy(false);
  }
  const problems = ROWS.filter(row => ["signin", "error", "blocked", "partial"].includes(healthOf(row.key))).length;
  return <div className="space-y-4">
    <div className="flex flex-wrap items-center justify-between gap-3">
      <p className="text-sm text-muted-foreground">{problems ? <><strong className="text-foreground">{problems}</strong> connection{problems > 1 ? "s need" : " needs"} attention.</> : "Statuses come from what the app has done so far this session, or from a check you run here."}</p>
      <Button className="rounded-full" disabled={busy} onClick={() => void checkFree()}>{busy ? <Loader2 aria-hidden="true" className="animate-spin" /> : <RefreshCw aria-hidden="true" />}Check all — free</Button>
    </div>
    <ul className="divide-y rounded-2xl border">{ROWS.map(row => {
      const health = healthOf(row.key), at = states[row.key]?.at;
      const check = FREE_CHECKS[row.key];
      return <li key={row.key} className="space-y-2 p-3">
        <div className="flex flex-wrap items-start justify-between gap-2">
          <div className="min-w-0 flex-1">
            <p className="flex flex-wrap items-center gap-2 text-sm font-semibold">{row.name}</p>
            <p className="text-xs text-muted-foreground">{row.powers}</p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <Status health={health} />
            {at && <span className="text-xs text-muted-foreground">{new Date(at).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}</span>}
            {check && <Button size="sm" variant="outline" className="rounded-full" disabled={health === "checking"} onClick={() => void runCheck(row.key, check)}>
              {health === "signin" || health === "error" ? "Try again" : "Check"}</Button>}
            <Button size="sm" variant={health === "signin" ? "default" : "ghost"} className="rounded-full" aria-expanded={signingIn === row.key} onClick={() => { setSigningIn(value => value === row.key ? undefined : row.key); awaiting.current = row.key; }}>
              <KeyRound aria-hidden="true" />Sign in again</Button>
          </div>
        </div>
        {help(row.key, health) && <p className="rounded-xl bg-accent/60 p-2.5 text-xs">{help(row.key, health)}</p>}
        {signingIn === row.key && <SignInSteps name={CONNECTOR_NAMES[row.key]} url={pageUrl} 
          onDone={() => { setSigningIn(undefined); if (check) void runCheck(row.key, check); }} />}
      </li>;
    })}</ul>
    <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border p-3 text-xs">
      <p className="min-w-0 flex-1 text-muted-foreground">Sign-in is handled by Microsoft 365, on your own Power Apps <em>Connections</em> page. You can renew any connection there at any time — no page reload needed here.</p>
      <a href={pageUrl} target="_blank" rel="noopener noreferrer" className="inline-flex items-center gap-1.5 rounded-full border border-primary/40 px-3 py-1.5 font-semibold text-primary hover:bg-accent"><KeyRound aria-hidden="true" className="size-3.5" />Open all my connections<ExternalLink aria-hidden="true" className="size-3" /><span className="sr-only"> (opens in a new tab)</span></a>
    </div>
  </div>;
}

export function ConnectionsButton() {
  const states = useConnectionStates();
  const bad = Object.values(states).some(state => state && ["signin", "error", "blocked"].includes(state.health));
  return <Dialog>
    <DialogTrigger asChild><Button variant="secondary" className="h-auto whitespace-normal rounded-full py-2"><Plug aria-hidden="true" />Connections<span className={cn("ml-1 inline-block size-2.5 rounded-full", bad ? "bg-amber-500" : "bg-emerald-500")} aria-label={bad ? "Some connections need attention" : "Connections look fine"} /></Button></DialogTrigger>
    <DialogContent className="max-h-[88vh] overflow-y-auto text-left sm:max-w-3xl">
      <DialogHeader>
        <DialogTitle>Connections</DialogTitle>
        <DialogDescription>Check that each Microsoft 365 connection is working and retry the ones that aren't — without reloading the whole dashboard.</DialogDescription>
      </DialogHeader>
      <ConnectionsPanel />
    </DialogContent>
  </Dialog>;
}
