import { useRef, useState } from "react";
import { Check, Coins, Copy, Download, EyeOff, ExternalLink, Loader2, NotebookPen, Play, RotateCcw, Sparkles, X } from "lucide-react";
import { dismiss, restore, setRetention, useDismissed, RETENTION_CHOICES } from "@/lib/dismissed";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { RichText } from "@/components/rich-text";
import { safeUrl } from "@/lib/safe-url";
import { cn } from "@/lib/utils";
import { CREDIT_NOTE, SOURCE_LABELS, rankActions, type ActionResult, type SourceKey, type SuggestedAction } from "@/lib/actions";
import { FORMAT_LABELS, type DocContent, type DocFormat } from "@/lib/documents";
import { downloadDocument, saveDocToOneNote } from "@/lib/action-runners";
import { recordHistory } from "@/lib/history";
import { DocView } from "@/components/copilot-text";
import { CopilotHandoff } from "@/components/copilot-handoff";
import { briefRequest } from "@/lib/copilot-prompts";

export function CostBadge({ cost }: { cost: "free" | "credits" }) {
  return cost === "free"
    ? <Badge variant="outline" className="h-auto gap-1 whitespace-normal border-sky-700/40 bg-sky-50 text-sky-900 dark:bg-sky-950 dark:text-sky-100"><Check aria-hidden="true" className="size-3" />Free</Badge>
    : <Badge variant="outline" className="h-auto gap-1 whitespace-normal border-amber-700/50 bg-amber-50 text-amber-950 dark:bg-amber-950 dark:text-amber-100"><Coins aria-hidden="true" className="size-3" />Uses Copilot Credits</Badge>;
}
function errorMessage(error: unknown) {
  return error instanceof Error && error.name === "ActionError" ? error.message : "That didn't work. Nothing else was changed. Try again, or check your connections with Refresh.";
}
export function ResultLinks({ links }: { links: { label: string; url: string }[] }) {
  const safe = links.filter(link => safeUrl(link.url));
  if (!safe.length) return null;
  return <ul className="mt-2 flex flex-wrap gap-2">{safe.map(link => <li key={link.url + link.label}>
    <a href={safeUrl(link.url)} target="_blank" rel="noopener noreferrer" className="inline-flex items-center gap-1.5 rounded-full border border-primary/40 bg-background px-3 py-1.5 text-sm font-medium text-primary underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-ring">
      {link.label}<ExternalLink aria-hidden="true" className="size-3.5" /><span className="sr-only">(opens in a new tab)</span>
    </a>
  </li>)}</ul>;
}
function CopyText({ text }: { text: string }) {
  const [copied, setCopied] = useState(false);
  return <div className="mt-3 rounded-xl border border-border bg-background p-3">
    <p className="whitespace-pre-wrap break-words text-sm">{text}</p>
    <Button variant="outline" size="sm" className="mt-2" onClick={() => { void navigator.clipboard?.writeText(text).then(() => { setCopied(true); window.setTimeout(() => setCopied(false), 2500); }); }}>
      {copied ? <Check aria-hidden="true" /> : <Copy aria-hidden="true" />}{copied ? "Copied" : "Copy text"}
    </Button>
  </div>;
}

/** Preview of generated content, with buttons to save it as Word, Excel or PDF in OneDrive. */
export function DocumentResult({ doc, formats: wanted = ["docx", "pdf", "xlsx"], rows, allowPdf = true, mySite }: { doc: DocContent; mySite?: string; formats?: DocFormat[]; rows?: (string | number)[][]; allowPdf?: boolean }) {
  const formats = allowPdf ? wanted : wanted.filter(format => format !== "pdf");
  const target = useRef<HTMLDivElement>(null);
  const [saving, setSaving] = useState<string>();
  const [saved, setSaved] = useState<ActionResult[]>([]);
  const [error, setError] = useState("");
  async function download(format: DocFormat) {
    setSaving(`dl-${format}`); setError("");
    try { const result = await downloadDocument(doc, format, target.current, rows); setSaved(value => [...value, { ...result, message: `${result.message} To keep it in OneDrive, open your OneDrive and drag the file in.`, links: [...result.links, { label: "Open my OneDrive", url: oneDriveUrl }] }]); recordHistory({ kind: "created", title: `Downloaded "${doc.title}" as ${FORMAT_LABELS[format]}`, status: "done", cost: "free", message: result.message }); }
    catch (cause) { setError(errorMessage(cause)); }
    finally { setSaving(undefined); }
  }
  const oneDriveUrl = mySite && /^https:\/\/[^/]+-my\.sharepoint\.com\//i.test(mySite) ? `${mySite.replace(/\/+$/, "")}/_layouts/15/onedrive.aspx` : "https://www.office.com/launch/onedrive";
  async function toOneNote() {
    setSaving("onenote"); setError("");
    try { const result = await saveDocToOneNote(doc); setSaved(value => [...value, result]); recordHistory({ kind: "created", title: `Saved "${doc.title}" to OneNote`, status: "done", cost: "free", message: result.message, links: result.links, doc }); }
    catch (cause) { setError(errorMessage(cause)); }
    finally { setSaving(undefined); }
  }
  return <div className="mt-3 space-y-3">
    <div ref={target} className="doc-preview max-h-[28rem] overflow-y-auto rounded-xl border border-border bg-card p-5" data-pdf-expand>
      <DocView doc={doc} />
    </div>
    <div data-pdf-exclude>
            <div className="mt-2 flex flex-wrap gap-2">{formats.map(format => <Button key={`dl-${format}`} variant="default" className="rounded-full" disabled={!!saving} onClick={() => void download(format)}>
        {saving === `dl-${format}` ? <Loader2 aria-hidden="true" className="animate-spin" /> : <Download aria-hidden="true" />}Download {FORMAT_LABELS[format]}
      </Button>)}
        <Button variant="outline" className="rounded-full" disabled={!!saving} onClick={() => void toOneNote()} title="Creates a OneNote page in Quick Notes, with formatting and clickable sources. OneNote is stored in your OneDrive.">
          {saving === "onenote" ? <Loader2 aria-hidden="true" className="animate-spin" /> : <NotebookPen aria-hidden="true" />}Save to OneNote
        </Button></div>

      {error && <p role="alert" className="mt-2 text-sm text-destructive">{error}</p>}
      {saved.map((result, index) => <div key={index} role="status" className="mt-2 text-sm"><p>{result.message}</p><ResultLinks links={result.links} /></div>)}
    </div>
  </div>;
}

export function SourceBadges({ sources }: { sources?: SourceKey[] }) {
  if (!sources?.length) return null;
  return <ul className="flex flex-wrap gap-1.5" aria-label="Found in">{sources.map(source => <li key={source}>
    <span className={cn("source-badge inline-flex items-center rounded-full px-2 py-0.5 text-xs font-semibold", `source-${source}`)}>{SOURCE_LABELS[source]}</span>
  </li>)}</ul>;
}
export function ActionRow({ action, mySite, rank, tone, section }: { action: SuggestedAction; mySite?: string; rank?: number; tone?: string; section?: string }) {
  const [open, setOpen] = useState(false);
  const [notes, setNotes] = useState(action.notesDefault ?? "");
  const [choice, setChoice] = useState(action.choices?.[0]?.value);
  const [state, setState] = useState<"idle" | "running" | "done" | "error">("idle");
  const [result, setResult] = useState<ActionResult>();
  const [error, setError] = useState("");
  const credits = action.cost === "credits";
  const openUrl = action.open ? safeUrl(action.open.url) : undefined;
  const hasCopilot = !!action.copilot || !!(action.copilotBrief && action.run);
  const showNotes = credits || action.notesLabel !== undefined || hasCopilot;
  async function copilotRequest() {
    if (action.copilot) return action.copilot({ notes, choice });
    const outcome = await action.run!({ notes: "", choice });
    return briefRequest(action.copilotBrief!, outcome.doc ?? { title: action.title, sections: [] }, notes);
  }
  async function run() {
    setState("running"); setError("");
    if (!action.run) return;
    const chosen = action.choices?.find(option => option.value === choice)?.label;
    const used = notes.trim() && notes.trim() !== (action.notesDefault ?? "").trim() ? notes.trim() : undefined;
    try {
      const outcome = await action.run({ notes, choice });
      setResult(outcome); setState("done");
      recordHistory({ kind: outcome.doc ? "document" : "action", title: action.title, section, cost: action.cost, status: "done", notes: used, choice: chosen, message: outcome.message, links: outcome.links, text: outcome.text, doc: outcome.doc });
    } catch (cause) {
      setError(errorMessage(cause)); setState("error");
      recordHistory({ kind: "action", title: action.title, section, cost: action.cost, status: "failed", notes: used, choice: chosen, error: errorMessage(cause) });
    }
  }
  const id = `action-${action.id.replace(/[^a-z0-9-]/gi, "-")}`;
  return <li className={cn("rounded-2xl border border-border bg-card p-4 transition-colors", tone, open && "border-primary/50 shadow-sm", state === "done" && "border-emerald-700/40")}>
    <div className="flex flex-wrap items-start justify-between gap-3">
      <div className="min-w-0 flex-1">
        <p className="break-words font-semibold">{rank !== undefined && <span className="mr-2 inline-flex size-6 items-center justify-center rounded-full bg-primary text-xs font-bold text-primary-foreground" aria-label={`Priority ${rank}`}>{rank}</span>}{state === "done" && <Check aria-label="Done" className="mr-1 inline size-4 text-sky-700" />}{openUrl
          ? <a href={openUrl} target="_blank" rel="noopener noreferrer" className="source-title underline-offset-4 hover:underline focus-visible:outline-2 focus-visible:outline-ring"><RichText text={action.title} interactive={false} /><span className="sr-only"> ({action.open!.label}, opens in a new tab, free)</span></a>
          : <RichText text={action.title} interactive={false} />}</p>
        {action.detail && <p className="mt-1 break-words text-sm text-muted-foreground"><RichText text={action.detail} /></p>}
        <div className="mt-2 flex flex-wrap items-center gap-2">{action.sources && <SourceBadges sources={action.sources} />}{action.run && <CostBadge cost={action.cost} />}</div>
      </div>
      {openUrl && <a href={openUrl} target="_blank" rel="noopener noreferrer" title="Opens the original item" className="inline-flex shrink-0 items-center gap-1.5 rounded-full border border-primary/40 bg-background px-3 py-1.5 text-sm font-medium text-primary hover:bg-accent focus-visible:outline-2 focus-visible:outline-ring">
        {action.open!.label}<ExternalLink aria-hidden="true" className="size-3.5" /><span className="sr-only">(opens in a new tab)</span>
      </a>}
      {(action.run || hasCopilot) && state !== "done" && !open && <Button variant={credits ? "outline" : "default"} className="shrink-0 rounded-full" aria-expanded={false} aria-controls={id} onClick={() => setOpen(true)}>
        {credits || !action.run ? <Sparkles aria-hidden="true" /> : <Play aria-hidden="true" />}{credits ? "Review" : !action.run ? "Draft a reply" : action.runLabel ?? "Start"}
      </Button>}
      {section && state !== "running" && !open && <Button variant="ghost" size="sm" className="dismiss-pill shrink-0 rounded-full" title="Hide this suggestion. You can restore it from Dismissed for a few days." onClick={() => dismiss(action.id, action.title, section)}>
        <EyeOff aria-hidden="true" />Dismiss<span className="sr-only">: {action.title}</span>
      </Button>}
    </div>
    {open && state !== "done" && <div id={id} className="mt-4 space-y-3 border-t border-border pt-4">
      {action.choices && <fieldset><legend className="text-sm font-medium">Choose</legend><div className="mt-2 flex flex-wrap gap-2">{action.choices.map(option =>
        <Button key={option.value} type="button" size="sm" variant={choice === option.value ? "default" : "outline"} aria-pressed={choice === option.value} onClick={() => setChoice(option.value)}>{option.label}</Button>)}</div></fieldset>}
      {showNotes && <div>
        <Label htmlFor={`${id}-notes`}>{action.notesLabel ?? "Notes (optional)"}</Label>
        <Textarea id={`${id}-notes`} value={notes} maxLength={1500} onChange={event => setNotes(event.target.value)} className="mt-1.5 min-h-20" placeholder={credits ? "For example: keep it short, mention the Q3 numbers, decline politely…" : ""} />
      </div>}
      {credits && <p className="flex items-start gap-2 rounded-xl bg-amber-50 p-3 text-sm text-amber-950 dark:bg-amber-950 dark:text-amber-100"><Coins aria-hidden="true" className="mt-0.5 size-4 shrink-0" />{CREDIT_NOTE}</p>}
      {hasCopilot && <p className="text-xs text-muted-foreground">{action.run ? "Write it yourself and select the button, or use Draft with Copilot: your notes above go into the request." : "Your notes go into the request for Copilot. The app never posts to Teams; you paste the reply there yourself."}</p>}
      <div className="flex flex-wrap gap-2">
        {action.run && <Button onClick={() => void run()} disabled={state === "running"} className="rounded-full">
          {state === "running" ? <Loader2 aria-hidden="true" className="animate-spin" /> : credits ? <Coins aria-hidden="true" /> : <Play aria-hidden="true" />}
          {state === "running" ? "Working…" : action.runLabel ?? (credits ? "Run — uses Copilot Credits" : "Run")}
        </Button>}
        <Button variant="ghost" disabled={state === "running"} onClick={() => { setOpen(false); setState("idle"); setError(""); }}><X aria-hidden="true" />Cancel</Button>
      </div>
      {state === "error" && <p role="alert" className="text-sm text-destructive">{error}</p>}
      {hasCopilot && <div className="border-t border-dashed border-border pt-3"><CopilotHandoff build={copilotRequest} title={action.title} section={section} disabled={state === "running"} /></div>}
    </div>}
    {state === "done" && result && <div role="status" className="mt-3 border-t border-border pt-3">
      <p className="text-sm">{result.message}</p>
      <ResultLinks links={result.links} />
      {result.text && <CopyText text={result.text} />}
      {result.doc && <DocumentResult doc={result.doc} mySite={mySite} />}
    </div>}
  </li>;
}

/** Suggested actions, most important first. Nothing runs until the user selects Run. */
/** Items dismissed from this section, restorable until they expire. */
export function DismissedList({ section, label = "Dismissed" }: { section: string; label?: string }) {
  const { items, retention, daysLeft } = useDismissed();
  const mine = items.filter(item => item.section === section);
  if (!mine.length) return null;
  return <details className="dismissed-list rounded-2xl border border-dashed border-border p-3">
    <summary className="cursor-pointer text-sm font-semibold text-muted-foreground focus-visible:outline-2 focus-visible:outline-ring">{label} ({mine.length})</summary>
    <div className="mt-3 space-y-3">
      <div className="flex flex-wrap items-center gap-2 text-xs text-muted-foreground">
        <span>Keep dismissed items here for</span>
        {RETENTION_CHOICES.map(days => <Button key={days} size="sm" variant={retention === days ? "default" : "outline"} aria-pressed={retention === days} className="h-7 rounded-full px-3 text-xs" onClick={() => setRetention(days)}>{days} day{days > 1 ? "s" : ""}</Button>)}
        <span>— then they're removed for good.</span>
      </div>
      <ul className="divide-y divide-border">{mine.map(item => {
        const left = daysLeft(item), ago = Math.floor((Date.now() - item.at) / 864e5);
        return <li key={item.key} className="flex flex-wrap items-center justify-between gap-2 py-2">
          <div className="min-w-0 flex-1"><p className="break-words text-sm text-muted-foreground line-through decoration-muted-foreground/40">{item.title}</p>
            <p className="text-xs text-muted-foreground">Dismissed {ago === 0 ? "today" : ago === 1 ? "yesterday" : `${ago} days ago`} · removed for good {left <= 1 ? "within a day" : `in ${left} days`}</p></div>
          <Button size="sm" variant="outline" className="rounded-full" onClick={() => restore(item.key)}><RotateCcw aria-hidden="true" />Restore<span className="sr-only">: {item.title}</span></Button>
        </li>;
      })}</ul>
    </div>
  </details>;
}
export function ActionList({ actions, mySite, section, empty = "Nothing to suggest right now. Suggestions update when you select Refresh." }: { actions: SuggestedAction[]; mySite?: string; section?: string; empty?: string }) {
  const { isDismissed } = useDismissed();
  const ranked = rankActions(section ? actions.filter(action => !isDismissed(action.id)) : actions);
  return <div className="space-y-3">
    {ranked.length ? <ol className="space-y-3" aria-label="Suggested actions, most important first">{ranked.map(action => <ActionRow key={action.id} action={action} mySite={mySite} section={section} />)}</ol>
      : <p className="rounded-2xl border border-dashed border-border p-5 text-sm text-muted-foreground">{actions.length && section ? "You've dismissed everything here. Restore items below, or wait for new suggestions after Refresh." : empty}</p>}
    {section && <DismissedList section={section} />}
  </div>;
}
