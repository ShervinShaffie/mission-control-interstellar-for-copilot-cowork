import { useState } from "react";
import { CopilotHandoff } from "@/components/copilot-handoff";
import { newEmailRequest } from "@/lib/copilot-prompts";
import { CalendarPlus, ListTodo, Loader2, Mail, NotebookPen, Search, UserPlus, X } from "lucide-react";
import type { LucideIcon } from "lucide-react";
import { Office365UsersService } from "../../generated/services/Office365UsersService";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { ResultLinks } from "@/components/action-list";
import { addToDo, oneNotePage, outlookComposeLink, saveEmailDraft, sendMeetingInvite, todoHtml } from "@/lib/action-runners";
import type { ActionResult } from "@/lib/actions";
import { recordHistory } from "@/lib/history";
import { firstName, hideEmails } from "@/lib/privacy";
import { cn } from "@/lib/utils";

type Kind = "email" | "meeting" | "todo" | "onenote";
const KINDS: { kind: Kind; label: string; icon: LucideIcon; hint: string }[] = [
  { kind: "email", label: "Email draft", icon: Mail, hint: "Saved in Outlook Drafts — you send it from Outlook." },
  { kind: "meeting", label: "Meeting invite", icon: CalendarPlus, hint: "Sent from your calendar when you select Send." },
  { kind: "todo", label: "To Do task", icon: ListTodo, hint: "Added to your default To Do list." },
  { kind: "onenote", label: "OneNote page", icon: NotebookPen, hint: "Created in your Quick Notes. Start a line with “- [ ]” for a to-do." },
];
interface Person { name: string; email: string }

/** Find people in your organization by name (Office 365 Users; free). Shows first names only. */
function PeoplePicker({ people, onChange, label }: { people: Person[]; onChange: (people: Person[]) => void; label: string }) {
  const [term, setTerm] = useState("");
  const [results, setResults] = useState<Person[]>();
  const [busy, setBusy] = useState(false);
  async function search() {
    const query = term.trim();
    if (!query) return;
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(query)) { onChange([...people, { name: query.split("@")[0], email: query }]); setTerm(""); return; }
    setBusy(true);
    const result = await Office365UsersService.SearchUserV2(query, 6, true);
    setBusy(false);
    setResults(result.success ? (result.data?.value ?? []).filter(user => user.Mail).map(user => ({ name: firstName(user.DisplayName, user.GivenName) || user.DisplayName || "Someone", email: user.Mail! })) : []);
  }
  return <div>
    <Label htmlFor="qa-people">{label}</Label>
    {people.length > 0 && <ul className="mt-1 flex flex-wrap gap-1.5">{people.map(person => <li key={person.email} className="inline-flex items-center gap-1 rounded-full border px-2 py-0.5 text-sm">{person.name}
      <button type="button" aria-label={`Remove ${person.name}`} onClick={() => onChange(people.filter(item => item.email !== person.email))} className="rounded-full p-0.5 hover:bg-accent"><X aria-hidden="true" className="size-3" /></button></li>)}</ul>}
    <form className="mt-1 flex gap-2" onSubmit={event => { event.preventDefault(); void search(); }}>
      <Input id="qa-people" value={term} onChange={event => setTerm(event.target.value)} placeholder="Type a name and press Find" />
      <Button type="submit" variant="outline" className="rounded-full" disabled={busy || !term.trim()}>{busy ? <Loader2 aria-hidden="true" className="animate-spin" /> : <Search aria-hidden="true" />}Find</Button>
    </form>
    {results && <ul className="mt-1 rounded-xl border">{results.length ? results.map(person => <li key={person.email}>
      <button type="button" className="flex w-full items-center gap-2 px-3 py-1.5 text-left text-sm hover:bg-accent" onClick={() => { if (!people.some(item => item.email === person.email)) onChange([...people, person]); setResults(undefined); setTerm(""); }}>
        <UserPlus aria-hidden="true" className="size-3.5" />{person.name}<span className="text-xs text-muted-foreground">({person.email.split("@")[1]})</span></button></li>)
      : <li className="px-3 py-1.5 text-sm text-muted-foreground">No one found.</li>}</ul>}
  </div>;
}

const toLocalInput = (date: Date) => new Date(date.getTime() - date.getTimezoneOffset() * 60000).toISOString().slice(0, 16);
function defaultStart() { const date = new Date(); date.setDate(date.getDate() + 1); date.setHours(10, 0, 0, 0); return date; }

/** Do something in Microsoft 365 straight from the dashboard. Nothing happens until you select a button; Copilot help is a copy-and-paste hand-off. */
export function QuickActions({ calendarId }: { calendarId?: string }) {
  const [kind, setKind] = useState<Kind>("email");
  const [people, setPeople] = useState<Person[]>([]);
  const [title, setTitle] = useState("");
  const [body, setBody] = useState("");
  const [start, setStart] = useState(toLocalInput(defaultStart()));
  const [minutes, setMinutes] = useState(30);
  const [location, setLocation] = useState("");
  const [due, setDue] = useState("");
  const [state, setState] = useState<"idle" | "running">("idle");
  const [result, setResult] = useState<ActionResult>();
  const [error, setError] = useState("");
  const current = KINDS.find(item => item.kind === kind)!;
  const startDate = new Date(start), endDate = new Date(startDate.getTime() + minutes * 60000);
  function reset() { setTitle(""); setBody(""); setPeople([]); setDue(""); setLocation(""); }
  async function run(task: () => Promise<ActionResult>) {
    setState("running"); setError(""); setResult(undefined);
    const what = `${current.label}: ${title || "(untitled)"}`;
    try { const outcome = await task(); setResult(outcome); recordHistory({ kind: "created", title: what, section: "ask", cost: "free", status: "done", message: outcome.message, links: outcome.links, text: body }); reset(); }
    catch (cause) { const message = cause instanceof Error && cause.name === "ActionError" ? cause.message : "That didn't work. Nothing else was changed."; setError(message); recordHistory({ kind: "created", title: what, section: "ask", status: "failed", error: message }); }
    finally { setState("idle"); }
  }
  const addresses = people.map(person => person.email).join(";");
  return <div className="space-y-4">
    <div role="group" aria-label="What do you want to do?" className="flex flex-wrap gap-2">{KINDS.map(item => <Button key={item.kind} size="sm" variant={kind === item.kind ? "default" : "outline"} aria-pressed={kind === item.kind} className="rounded-full" onClick={() => { setKind(item.kind); setResult(undefined); setError(""); }}><item.icon aria-hidden="true" />{item.label}</Button>)}</div>
    <div className="quick-form space-y-3 rounded-2xl border p-4">
      <p className="text-xs text-muted-foreground">{current.hint}</p>
      {(kind === "email" || kind === "meeting") && <PeoplePicker people={people} onChange={setPeople} label={kind === "email" ? "To" : "Attendees"} />}
      <div><Label htmlFor="qa-title">{kind === "email" || kind === "meeting" ? "Subject" : "Title"}</Label><Input id="qa-title" value={title} maxLength={200} onChange={event => setTitle(event.target.value)} className="mt-1" /></div>
      {kind === "meeting" && <div className="grid gap-3 sm:grid-cols-3">
        <div><Label htmlFor="qa-start">Starts</Label><Input id="qa-start" type="datetime-local" value={start} onChange={event => setStart(event.target.value)} className="mt-1" /></div>
        <div><Label htmlFor="qa-length">Length</Label><select id="qa-length" value={minutes} onChange={event => setMinutes(Number(event.target.value))} className="mt-1 h-9 w-full rounded-md border bg-background px-2 text-sm">{[15, 30, 45, 60, 90].map(value => <option key={value} value={value}>{value} minutes</option>)}</select></div>
        <div><Label htmlFor="qa-loc">Location</Label><Input id="qa-loc" value={location} onChange={event => setLocation(event.target.value)} className="mt-1" placeholder="Optional" /></div>
      </div>}
      {kind === "todo" && <div className="max-w-xs"><Label htmlFor="qa-due">Due date (optional)</Label><Input id="qa-due" type="date" value={due} onChange={event => setDue(event.target.value)} className="mt-1" /></div>}
      <div><Label htmlFor="qa-body">{kind === "meeting" ? "Agenda" : kind === "todo" ? "Notes" : "Text"}</Label><Textarea id="qa-body" value={body} maxLength={8000} onChange={event => setBody(event.target.value)} className="mt-1 min-h-28" /></div>
      <div className="flex flex-wrap items-center gap-2">
        {kind === "email" && <Button className="rounded-full" disabled={state === "running" || !title.trim()} onClick={() => void run(() => saveEmailDraft(addresses, title, body))}>{state === "running" ? <Loader2 aria-hidden="true" className="animate-spin" /> : <Mail aria-hidden="true" />}Save as Outlook draft</Button>}
        {kind === "meeting" && <>
          <Button className="rounded-full" disabled={state === "running" || !calendarId || !title.trim() || !people.length || Number.isNaN(startDate.getTime())} onClick={() => void run(() => sendMeetingInvite(calendarId!, { subject: title, start: startDate, end: endDate, attendees: addresses, body, location }))}>{state === "running" ? <Loader2 aria-hidden="true" className="animate-spin" /> : <CalendarPlus aria-hidden="true" />}Send the invite</Button>
          {!Number.isNaN(startDate.getTime()) && <a className="inline-flex items-center rounded-full border px-3 py-1.5 text-sm font-medium hover:bg-accent" href={outlookComposeLink(title, startDate, endDate, body)} target="_blank" rel="noopener noreferrer">Open in Outlook instead<span className="sr-only"> (opens in a new tab)</span></a>}
        </>}
        {kind === "todo" && <Button className="rounded-full" disabled={state === "running" || !title.trim()} onClick={() => void run(() => addToDo(title, due || undefined, body, undefined))}><ListTodo aria-hidden="true" />Add to To Do</Button>}
        {kind === "onenote" && <Button className="rounded-full" disabled={state === "running" || !title.trim()} onClick={() => void run(() => {
          const lines = body.split(/\r?\n/), todo = /^\s*-\s*\[\s?\]\s*/;
          const html = lines.filter(line => !todo.test(line)).map(line => `<p>${line.replace(/[<>&]/g, char => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;" })[char]!)}</p>`).join("");
          const todos = lines.filter(line => todo.test(line)).map(line => line.replace(todo, ""));
          return oneNotePage(title, `<h1>${title.replace(/[<>&]/g, "")}</h1>${html}${todos.length ? todoHtml(todos) : ""}`);
        })}><NotebookPen aria-hidden="true" />Create the OneNote page</Button>}
        {kind === "meeting" && <span className="text-xs text-muted-foreground">Attendees get the invite as soon as you select Send.</span>}
      </div>
      {kind === "email" && <div className="border-t border-dashed border-border pt-3"><CopilotHandoff title={`Email: ${title || "(no subject)"}`} section="ask" build={() => newEmailRequest(people.map(person => person.name).join(", "), title, body)} />
        <p className="mt-1 text-xs text-muted-foreground">Write a few words in Text about what you want to say; they go into the request. Paste Copilot's answer back into Text, then save the draft.</p></div>}
      {error && <p role="alert" className="text-sm text-destructive">{error}</p>}
      {result && <div role="status" className={cn("rounded-xl bg-accent/50 p-3 text-sm")}><p>{hideEmails(result.message)}</p><ResultLinks links={result.links} /></div>}
    </div>
  </div>;
}
