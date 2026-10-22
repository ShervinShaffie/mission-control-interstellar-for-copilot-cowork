import { useState, type ReactNode } from "react";
import { ArrowDown, ArrowRight, BookOpen, CalendarDays, Download, EyeOff, FolderOpen, HardDrive, Hand, Info, ListTodo, Mail, MessagesSquare, NotebookPen, ShieldCheck, Sparkles, UserRound } from "lucide-react";
import type { LucideIcon } from "lucide-react";
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from "@/components/ui/accordion";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogDescription, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { cn } from "@/lib/utils";

function Node({ icon: Icon, title, note, tone = "read" }: { icon: LucideIcon; title: string; note?: string; tone?: "read" | "copilot" | "write" | "local" }) {
  return <li className={cn("hiw-node flex items-start gap-2 rounded-xl border px-3 py-2", `hiw-${tone}`)}>
    <Icon aria-hidden="true" className="mt-0.5 size-4 shrink-0" />
    <span className="min-w-0"><span className="block text-sm font-semibold">{title}</span>{note && <span className="block text-xs opacity-80">{note}</span>}</span>
  </li>;
}
function Arrow({ label }: { label: string }) {
  return <div className="hiw-arrow flex shrink-0 flex-col items-center justify-center gap-1 px-1 py-2 text-center text-xs font-semibold" aria-hidden="true">
    <ArrowRight className="hidden size-6 lg:block" /><ArrowDown className="size-6 lg:hidden" /><span className="max-w-24">{label}</span>
  </div>;
}
function Column({ title, subtitle, children, className }: { title: string; subtitle: string; children: ReactNode; className?: string }) {
  return <section className={cn("hiw-column min-w-0 flex-1 rounded-2xl border p-3", className)} aria-label={title}>
    <h4 className="text-sm font-bold">{title}</h4>
    <p className="mb-2 text-xs text-muted-foreground">{subtitle}</p>
    <ul className="space-y-1.5">{children}</ul>
  </section>;
}

/** The read → think → you decide → write flow, drawn with real elements (selectable, accessible text). */
export function HowItWorksDiagram() {
  return <figure className="hiw-diagram rounded-2xl border p-3 sm:p-4">
    <div className="flex flex-col gap-1 lg:flex-row lg:items-stretch">
      <Column title="1 · Read" subtitle="When the app opens and when you select Refresh">
        <Node icon={CalendarDays} title="Outlook calendar & mail" note="Events, invites, emails waiting on you, who's away" />
        <Node icon={MessagesSquare} title="Teams" note="Chats waiting on you, posts in your project teams" />
        <Node icon={ListTodo} title="To Do" note="Open tasks" />
        <Node icon={NotebookPen} title="OneNote" note="Recent pages, open to-dos" />
        <Node icon={HardDrive} title="OneDrive" note="Your files and your History file" />
        <Node icon={UserRound} title="Your profile" note="First names, never email addresses" />
      </Column>
      <Arrow label="Combined in your browser" />
      <Column title="2 · Understand" subtitle="Worked out in the app with simple, visible rules">
        <Node icon={Sparkles} title="Today → Priorities" note="One ranked list across every source" tone="local" />
        <Node icon={CalendarDays} title="Your day plan" note="Built from your calendar and tasks" tone="local" />
        <Node icon={FolderOpen} title="Projects board" note="Meetings, prep notes, deadlines per project" tone="local" />
        <Node icon={Info} title="Suggested actions" note="Each component's Actions tab" tone="local" />
      </Column>
      <Arrow label="You review" />
      <Column title="3 · You decide" subtitle="Nothing happens until you choose">
        <Node icon={Hand} title="Open the original" note="Free — read it first" tone="local" />
        <Node icon={UserRound} title="Review, add your text, Run" note="Nothing happens before that" tone="local" />
        <Node icon={EyeOff} title="Or Dismiss it" note="Restorable for 1–7 days" tone="local" />
        <Node icon={Sparkles} title="Or Draft with Copilot" note="Copies a ready-made request; you paste it into your own Copilot" tone="local" />
        <Node icon={Sparkles} title="Or use Quick actions" note="Email, invite, task, note or document" tone="local" />
      </Column>
      <Arrow label="Only when you select Run" />
      <Column title="4 · Write" subtitle="Back into your own Microsoft 365">
        <Node icon={Mail} title="Outlook" note="Drafts (never sent), invite replies, focus blocks, meeting invites, flags, auto-replies" tone="write" />
        <Node icon={ListTodo} title="To Do" note="New tasks" tone="write" />
        <Node icon={NotebookPen} title="OneNote" note="Pages in Quick Notes" tone="write" />
                <Node icon={Download} title="Your computer" note="Downloads" tone="write" />
        <Node icon={BookOpen} title="History" note="Everything you ran or created — searchable" tone="write" />
      </Column>
    </div>
    <figcaption className="mt-3 flex flex-wrap gap-x-4 gap-y-1 text-xs text-muted-foreground">
      <span className="flex items-center gap-1.5"><span className="hiw-key hiw-read" aria-hidden="true" />Read from Microsoft 365</span>
      <span className="flex items-center gap-1.5"><span className="hiw-key hiw-local" aria-hidden="true" />Worked out in the app</span>
      <span className="flex items-center gap-1.5"><span className="hiw-key hiw-write" aria-hidden="true" />Writes, only when you choose</span>
    </figcaption>
  </figure>;
}

const READS: [string, string, string][] = [
  ["Office 365 Outlook", "Your calendar (next 2 weeks), inbox and sent mail (last 50 of each), and the automatic-reply status of people in your meetings", "Who's away, emails needing a reply, emails waiting on others, meeting prep, file links in emails and invites, Calendar, Email"],
  ["Microsoft Teams", "Your chats, and the latest posts in up to 8 of your teams (you can choose which teams are projects)", "Projects, deadlines and @mentions, Teams, files shared in posts and chats"],
  ["Microsoft To Do", "Your open tasks (up to 6 lists)", "Today, Projects"],
  ["OneNote", "Recently changed pages and their unchecked to-dos", "OneNote, Today"],
  ["OneDrive for Business", "Your folders and files (My OneDrive), and your History file", "Files, History"],
  ["Office 365 Users", "Your profile, first names, people search and documents trending around you", "Greeting, names, Quick actions, Files"],
];
const WRITES: [string, string, string][] = [
  ["Answer an invite, or decline one of two overlapping meetings", "Outlook calendar (the organizer is notified)", "Calendar, Today"],
  ["Block focus time or time for a category", "Outlook calendar (private, busy)", "Calendar, Insights, Today"],
  ["Send a meeting invite", "Outlook calendar — attendees get it right away", "Quick actions"],
  ["Flag or mark an email read; turn on automatic replies", "Outlook", "Email, Today"],
  ["Save a reply or follow-up you write", "Outlook Drafts — you send it yourself", "Email, Today, Projects"],
  ["Add a task", "Microsoft To Do", "Today, OneNote, Projects, Quick actions"],
  ["Create a day plan, meeting notes or a page", "OneNote Quick Notes", "OneNote, Today, Quick actions"],
  ["Build a prep brief, project briefing, status report or Teams catch-up", "A preview in the app, to download as Word, PDF or Excel", "Calendar, Projects, Teams"],
  ["Save a brief, report or document to OneNote", "OneNote Quick Notes (stored in your OneDrive), with formatting and clickable sources", "Any document preview"],
  ["Draft with Copilot", "Your clipboard, and Microsoft 365 Copilot opens in a new tab — you paste and send", "Replies, follow-ups, Teams replies, briefs, reports, Quick actions"],
];
function Table({ head, rows }: { head: [string, string, string]; rows: [string, string, string][] }) {
  return <div className="overflow-x-auto rounded-xl border"><table className="w-full min-w-[36rem] text-left text-sm">
    <thead className="bg-accent/60"><tr>{head.map(cell => <th key={cell} scope="col" className="px-3 py-2 font-semibold">{cell}</th>)}</tr></thead>
    <tbody className="divide-y">{rows.map(row => <tr key={row[0]}><th scope="row" className="px-3 py-2 align-top font-medium">{row[0]}</th><td className="px-3 py-2 align-top">{row[1]}</td>
      <td className={cn("px-3 py-2 align-top text-xs font-semibold", "text-muted-foreground")}>{row[2]}</td></tr>)}</tbody>
  </table></div>;
}

export function HowItWorks({ trigger = "button" }: { trigger?: "button" | "link" }) {
  const [open, setOpen] = useState(false);
  return <Dialog open={open} onOpenChange={setOpen}>
    <DialogTrigger asChild>
      {trigger === "button"
        ? <Button variant="secondary" className="h-auto whitespace-normal rounded-full py-2"><BookOpen aria-hidden="true" />How it works</Button>
        : <Button variant="link" className="h-auto p-0 text-sm"><BookOpen aria-hidden="true" />How Mission Control Interstellar works</Button>}
    </DialogTrigger>
    <DialogContent className="max-h-[88vh] overflow-y-auto text-left sm:max-w-6xl">
      <DialogHeader>
        <DialogTitle>How Mission Control Interstellar works</DialogTitle>
        <DialogDescription>It reads your own Microsoft 365 through six standard connections, works out what needs you, and suggests what it can do. It only changes anything when you select a button.</DialogDescription>
      </DialogHeader>
      <HowItWorksDiagram />
      <Accordion type="multiple" defaultValue={["read", "write"]} className="mt-2">
        <AccordionItem value="read">
          <AccordionTrigger>The read side — where your data comes from</AccordionTrigger>
          <AccordionContent className="space-y-3">
            <p className="text-sm">Everything is read with <strong>your</strong> sign-in, so you only see what you already have access to. Data loads once when you open the app and again only when you select <strong>Refresh</strong>. Nothing runs on a schedule.</p>
            <Table head={["Connection", "What it reads", "Used for"]} rows={READS} />
            <p className="text-xs text-muted-foreground">Only these six standard Microsoft 365 connections are used, so the app works the same for everyone you share it with.</p>
          </AccordionContent>
        </AccordionItem>
        <AccordionItem value="write">
          <AccordionTrigger>The write side — what the actions can do for you</AccordionTrigger>
          <AccordionContent className="space-y-3">
            <p className="text-sm">Suggestions are worked out from what was read. An action only happens after you open it and select its button. When something is created you get links to it, and it\'s kept in <strong>History</strong> (top of the page, or Quick actions → History) with a search bar.</p>
            <Table head={["Action", "Where it lands", "Found in"]} rows={WRITES} />
            <p className="text-sm"><strong>Never done by the app:</strong> sending email, posting to Teams, deleting anything, or acting without your click.</p>
          </AccordionContent>
        </AccordionItem>
        <AccordionItem value="components">
          <AccordionTrigger>What each part of the dashboard is for</AccordionTrigger>
          <AccordionContent>
            <ul className="grid gap-2 text-sm sm:grid-cols-2">
              <li><strong>Today</strong> — your day plan, what needs attention, and <em>Priorities</em>: one ranked list across every source, grouped Overdue → Next 2 hours → Today → Tomorrow → This week.</li>
              <li><strong>Quick actions</strong> — draft an email, send a meeting invite, add a task, create a OneNote page or a document. Plus your History.</li>
              <li><strong>Files</strong> — files that need you now, everything shared with you recently, and your OneDrive.</li>
              <li><strong>Projects</strong> — your project teams, most in need of attention first, with status, deadlines mentioned in posts, meetings, briefings, a weekly status report and a tracker you can download. Use <strong>Choose project teams</strong> to pick which teams count, and Collapse any project (or all of them) to keep the board short; the app remembers it in this browser.</li>
              <li><strong>Calendar</strong> — the next 7 days, a day agenda, and who's away.</li>
              <li><strong>Email</strong> — emails needing a reply, unread, flagged and recent.</li>
              <li><strong>Teams</strong> — who's waiting on you, and your recent chats.</li>
              <li><strong>OneNote</strong> — recent pages and open to-dos.</li>
              <li><strong>Time insights</strong> — meeting load, time by your Outlook categories, and targets.</li>
            </ul>
            <p className="mt-2 text-sm">Each component has a data tab and an <strong>Actions</strong> tab, ranked most important first. Click any item's title or <strong>Open</strong> to see the original email, chat, file or event for free before deciding.</p>
          </AccordionContent>
        </AccordionItem>
        <AccordionItem value="rules">
          <AccordionTrigger>How the app decides what needs you</AccordionTrigger>
          <AccordionContent>
            <ul className="list-disc space-y-1.5 pl-5 text-sm">
              <li><strong>Projects</strong> are your Teams teams — the ones you choose (Projects → Choose project teams), or up to 8 with project-sounding names. Status comes from words like RED, AMBER, at risk or blocked in recent posts.</li>
              <li><strong>Deadlines</strong> are phrases like "due 23 Oct", "by Friday" or "deadline Oct 30" in team posts from the last 3 weeks.</li>
              <li><strong>Urgent</strong> means a post in the last 3 days that @mentions you, is marked important, or says RED, escalated, blocked or urgent.</li>
              <li><strong>Emails needing a reply</strong> are emails from the last 14 days, addressed to you, that ask something (a question, "please", "can you", "review"…) and that you haven't answered in the same conversation.</li>
              <li><strong>Waiting on others</strong> are emails you sent 2–14 days ago that asked something and have no reply yet.</li>
              <li><strong>Who's away</strong> is anyone in your meetings this week whose Outlook automatic reply is on.</li>
              <li><strong>Files for you now</strong> are files linked or attached in an email you owe a reply (highest), a Teams chat waiting on you, an invite for a meeting in the next 2 days, a post that mentions you, your project teams' posts or other recent email, plus documents trending around you. The same file found in several places becomes one row with every reason.</li>
              <li><strong>Meeting prep</strong> gathers the invite's agenda, the related team's latest post and related emails for meetings in the next 2 days.</li>
            </ul>
            <p className="mt-2 text-sm">These are plain rules, so they can miss things or include something you don't need — use <strong>Dismiss</strong> to hide an item.</p>
          </AccordionContent>
        </AccordionItem>
        <AccordionItem value="copilot">
          <AccordionTrigger>Getting help from Copilot</AccordionTrigger>
          <AccordionContent>
            <p className="text-sm">The dashboard itself uses no AI: the suggestions come from the plain rules above, and drafts contain only what you type. When you want AI help, select <strong>Draft with Copilot</strong> on a reply, follow-up, Teams reply, brief, report, Quick actions email or document.</p>
            <ol className="mt-2 list-decimal space-y-1 pl-5 text-sm">
              <li>The app writes a ready-made request: the task, the context it already read (subject, who, dates, the message preview, related posts and links), your notes, and clear instructions (use your Microsoft 365 data, don't invent facts, mark gaps in [brackets]).</li>
              <li>It copies the request and opens Microsoft 365 Copilot in a new tab.</li>
              <li>You paste it (Ctrl+V) and press Enter, then copy Copilot's answer into your Outlook draft, Teams chat or document.</li>
            </ol>
            <p className="mt-2 text-sm">It uses your own Copilot, under your own sign-in and licence. The app never calls Copilot, never uses Copilot Credits and never sends anything for you. If your browser blocks copying, select <strong>Show the request</strong> and copy it yourself. Each request is saved in History so you can copy it again.</p>
          </AccordionContent>
        </AccordionItem>
        <AccordionItem value="privacy">
          <AccordionTrigger>Privacy, settings and what's saved</AccordionTrigger>
          <AccordionContent className="space-y-2 text-sm">
            <p className="flex items-start gap-2"><ShieldCheck aria-hidden="true" className="mt-0.5 size-4 shrink-0" />Everything is read with your own sign-in, so you only see what you already have access to. No one\'s email address is shown — people appear by first name.</p>
            <p>Saved in this browser only, for your account: your layout, theme, time targets, chosen project teams and dismissed suggestions. Your emails, chats and files are read fresh each time and not stored.</p>
            <p><strong>History</strong> keeps what you did from the dashboard — drafts, invites, tasks, pages and documents — in this browser and in a file in your own OneDrive (<em>Mission Control Interstellar › Activity history.json</em>). You can delete single items or clear it all.</p>
            <p>Dismissed suggestions wait in a <em>Dismissed</em> list at the bottom of each Actions tab for 1, 3 or 7 days (your choice) so you can restore them, then they're removed for good. <strong>Customize</strong> can bring everything back.</p>
          </AccordionContent>
        </AccordionItem>
        <AccordionItem value="help">
          <AccordionTrigger>If something looks empty or wrong</AccordionTrigger>
          <AccordionContent>
            <ul className="list-disc space-y-1.5 pl-5 text-sm">
              <li><strong>A section says it couldn't be read</strong> — open <strong>Connections</strong> at the top. Select Try again, or <strong>Sign in again</strong> and follow the steps.</li>
              <li><strong>Projects shows the wrong teams</strong> — Projects → <strong>Choose project teams</strong>, then Refresh.</li>
              <li><strong>Something is listed that you don't need</strong> — select <strong>Dismiss</strong>. It can be restored for a few days.</li>
              <li><strong>Can't find something you created</strong> — open <strong>History</strong> and search for it.</li>
              <li><strong>Times look off</strong> — times use your browser's time zone.</li>
              <li><strong>Want to start over</strong> — Customize → Restore default layout (keeps your theme).</li>
            </ul>
          </AccordionContent>
        </AccordionItem>
      </Accordion>
    </DialogContent>
  </Dialog>;
}
