import { useState } from "react";
import { AlertTriangle, CheckCircle2, Cloud, CloudOff, Copy, History, Loader2, Search, Trash2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Dialog, DialogContent, DialogDescription, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { DocumentResult, ResultLinks } from "@/components/action-list";
import { ReplyView } from "@/components/copilot-text";
import { clearHistory, deleteHistory, retryHistorySync, searchHistory, useHistory, HISTORY_FILE, type HistoryEntry, type HistoryKind } from "@/lib/history";
import { safeUrl } from "@/lib/safe-url";
import { cn } from "@/lib/utils";

const FILTERS: [("all" | HistoryKind | "failed"), string][] = [["all", "All"], ["document", "Documents"], ["action", "Actions"], ["created", "Quick actions & created"], ["failed", "Didn't work"]];
const KIND_LABEL: Record<HistoryKind, string> = { action: "Action", ask: "Quick action", document: "Document", created: "Created" };
const SECTION_LABEL: Record<string, string> = { today: "Today", "today-helpers": "Today", ask: "Quick actions", projects: "Projects", calendar: "Calendar", email: "Email", work: "Work", teams: "Teams", files: "Files", onenote: "OneNote", insights: "Time insights", create: "Create a document" };

const dayLabel = (at: number) => {
  const date = new Date(at), today = new Date(), yesterday = new Date(); yesterday.setDate(today.getDate() - 1);
  return date.toDateString() === today.toDateString() ? "Today" : date.toDateString() === yesterday.toDateString() ? "Yesterday" : date.toLocaleDateString([], { weekday: "long", month: "long", day: "numeric", year: date.getFullYear() === today.getFullYear() ? undefined : "numeric" });
};
const clock = (at: number) => new Date(at).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });

function Entry({ entry, mySite, allowPdf }: { entry: HistoryEntry; mySite?: string; allowPdf: boolean }) {
  const [copied, setCopied] = useState(false);
  const failed = entry.status === "failed";
  return <li className={cn("history-entry rounded-2xl border bg-card", failed && "history-failed")}>
    <details>
      <summary className="flex cursor-pointer list-none flex-wrap items-start gap-x-3 gap-y-1 p-3 focus-visible:outline-2 focus-visible:outline-ring">
        <span className="text-xs tabular-nums text-muted-foreground">{clock(entry.at)}</span>
        <span className="min-w-0 flex-1 break-words text-sm font-semibold">{failed ? <AlertTriangle aria-label="Didn't work" className="mr-1 inline size-4 text-amber-700" /> : <CheckCircle2 aria-hidden="true" className="mr-1 inline size-4 text-sky-700" />}{entry.title}</span>
        <span className="flex flex-wrap gap-1.5 text-xs">
          <span className="rounded-full border px-2">{KIND_LABEL[entry.kind]}</span>
          {entry.section && <span className="rounded-full border px-2">{SECTION_LABEL[entry.section] ?? entry.section}</span>}
          {entry.cost === "credits" && <span className="rounded-full border border-amber-600/50 bg-amber-50 px-2 text-amber-900 dark:bg-amber-950 dark:text-amber-100">Used credits</span>}
        </span>
      </summary>
      <div className="space-y-3 border-t px-3 pb-3 pt-2 text-sm">
        {entry.request && <p><span className="font-semibold">You asked:</span> {entry.request}</p>}
        {entry.notes && <p><span className="font-semibold">Your notes:</span> {entry.notes}</p>}
        {entry.choice && <p><span className="font-semibold">Choice:</span> {entry.choice}</p>}
        {entry.message && <p>{entry.message}</p>}
        {entry.error && <p className="text-destructive">{entry.error}</p>}
        {entry.links?.some(link => safeUrl(link.url)) && <ResultLinks links={entry.links} />}
        {entry.text && (entry.kind === "ask" ? <ReplyView text={entry.text} /> : <div className="rounded-xl border bg-muted/40 p-3"><p className="whitespace-pre-wrap break-words">{entry.text}</p>
          <Button variant="outline" size="sm" className="mt-2 rounded-full" onClick={() => { void navigator.clipboard?.writeText(entry.text ?? "").then(() => setCopied(true)); }}><Copy aria-hidden="true" />{copied ? "Copied" : "Copy"}</Button></div>)}
        {entry.doc && <DocumentResult doc={entry.doc} mySite={mySite} allowPdf={allowPdf} />}
        <div className="flex justify-end"><Button variant="ghost" size="sm" className="rounded-full text-muted-foreground" onClick={() => deleteHistory(entry.id)}><Trash2 aria-hidden="true" />Delete from history<span className="sr-only">: {entry.title}</span></Button></div>
      </div>
    </details>
  </li>;
}

/** Searchable history of everything done from the dashboard. */
export function HistoryPanel({ mySite, allowPdf = true }: { mySite?: string; allowPdf?: boolean }) {
  const { entries, sync, fileUrl } = useHistory();
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<(typeof FILTERS)[number][0]>("all");
  const [confirmClear, setConfirmClear] = useState(false);
  const [limit, setLimit] = useState(30);
  const found = searchHistory(entries, query, filter);
  const shown = found.slice(0, limit);
  const days: { label: string; items: HistoryEntry[] }[] = [];
  for (const entry of shown) { const label = dayLabel(entry.at); const last = days.at(-1); if (last?.label === label) last.items.push(entry); else days.push({ label, items: [entry] }); }
  const file = safeUrl(fileUrl);
  return <div className="space-y-4">
    <div className="flex flex-wrap items-center justify-between gap-2 text-xs text-muted-foreground">
      <p className="flex items-center gap-1.5">
        {sync === "saved" ? <Cloud aria-hidden="true" className="size-4 text-sky-700" /> : sync === "loading" || sync === "saving" ? <Loader2 aria-hidden="true" className="size-4 animate-spin" /> : <CloudOff aria-hidden="true" className="size-4" />}
        {sync === "saved" ? <>Saved in this browser and in your OneDrive{file ? <> — <a href={file} target="_blank" rel="noopener noreferrer" className="font-semibold text-primary underline-offset-2 hover:underline">Mission Control Interstellar › {HISTORY_FILE}<span className="sr-only"> (opens in a new tab)</span></a></> : null}.</>
          : sync === "loading" ? "Loading your history from OneDrive…" : sync === "saving" ? "Saving to your OneDrive…"
          : sync === "local" ? <>Saved in this browser only — the OneDrive copy couldn't be reached. <Button variant="link" className="h-auto p-0 text-xs" onClick={retryHistorySync}>Try again</Button></> : "Saved in this browser."}
      </p>
    </div>
    <div className="relative">
      <Search aria-hidden="true" className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
      <label htmlFor="history-search" className="sr-only">Search history</label>
      <Input id="history-search" type="search" value={query} onChange={event => { setQuery(event.target.value); setLimit(30); }} placeholder="Search titles, notes and text…" className="rounded-full pl-9" />
    </div>
    <div role="group" aria-label="Filter history" className="flex flex-wrap gap-2">{FILTERS.map(([value, label]) =>
      <Button key={value} size="sm" variant={filter === value ? "default" : "outline"} aria-pressed={filter === value} className="rounded-full" onClick={() => { setFilter(value); setLimit(30); }}>{label}</Button>)}</div>
    <p className="sr-only" aria-live="polite">{found.length} matching items.</p>
    {!entries.length ? <div className="rounded-2xl border border-dashed p-6 text-center text-sm text-muted-foreground"><History aria-hidden="true" className="mx-auto mb-2 size-6" />Nothing here yet. Everything you run or create from the dashboard appears here — and stays after Refresh.</div>
      : !found.length ? <p className="rounded-2xl border border-dashed p-5 text-sm text-muted-foreground">Nothing matches. Try fewer words or another filter.</p>
      : days.map(day => <section key={day.label} aria-label={day.label}>
        <h4 className="mb-2 text-xs font-bold uppercase tracking-wide text-muted-foreground">{day.label}</h4>
        <ol className="space-y-2">{day.items.map(entry => <Entry key={entry.id} entry={entry} mySite={mySite} allowPdf={allowPdf} />)}</ol>
      </section>)}
    {found.length > shown.length && <Button variant="outline" className="rounded-full" onClick={() => setLimit(value => value + 30)}>Show more ({found.length - shown.length})</Button>}
    {entries.length > 0 && <div className="flex flex-wrap items-center justify-end gap-2 border-t pt-3 text-xs">
      {confirmClear ? <><span>Delete all {entries.length} items, here and in OneDrive?</span><Button size="sm" variant="destructive" className="rounded-full" onClick={() => { clearHistory(); setConfirmClear(false); }}>Delete all</Button><Button size="sm" variant="outline" className="rounded-full" onClick={() => setConfirmClear(false)}>Cancel</Button></>
        : <Button size="sm" variant="ghost" className="rounded-full text-muted-foreground" onClick={() => setConfirmClear(true)}><Trash2 aria-hidden="true" />Clear history</Button>}
    </div>}
  </div>;
}

export function HistoryButton({ mySite }: { mySite?: string }) {
  const { entries } = useHistory();
  return <Dialog>
    <DialogTrigger asChild><Button variant="secondary" className="h-auto whitespace-normal rounded-full py-2"><History aria-hidden="true" />History{entries.length ? <span className="ml-1 rounded-full bg-primary px-1.5 text-xs text-primary-foreground">{entries.length}</span> : null}</Button></DialogTrigger>
    <DialogContent className="max-h-[88vh] overflow-y-auto text-left sm:max-w-4xl">
      <DialogHeader>
        <DialogTitle>History</DialogTitle>
        <DialogDescription>Everything you've done or created from this dashboard — drafts, invites, tasks, pages and documents — so you can find it again later.</DialogDescription>
      </DialogHeader>
      <HistoryPanel mySite={mySite} allowPdf={false} />
      <p className="text-xs text-muted-foreground">To download a document as PDF, open it from Quick actions → History on the dashboard.</p>
    </DialogContent>
  </Dialog>;
}
