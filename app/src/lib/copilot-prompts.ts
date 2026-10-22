/**
 * "Draft with Copilot": ready-made requests for the person's OWN Microsoft 365 Copilot.
 * The app never calls Copilot. It builds a request from data it already read, copies it, and opens
 * Copilot in a new tab; the person pastes it and sends. Nothing here makes a network call.
 */
import type { WorkItem } from "@/lib/command-center-data";
import type { DocContent } from "@/lib/documents";
import { hideEmails } from "@/lib/privacy";

export const COPILOT_URL = "https://m365.cloud.microsoft/chat";
/** Long requests are trimmed so they paste cleanly into Copilot. */
export const MAX_REQUEST = 12000;
export type BriefKind = "prep" | "project" | "status" | "catchup" | "document";

interface Request { task: string; context?: (string | false | undefined)[]; material?: string; notes?: string; output: string }

const RULES = [
  "Use my Microsoft 365 data (emails, Teams chats and channel posts, meetings and files) to fill in the latest details. Search by the subjects and names above.",
  "Don't invent facts, numbers, dates, owners or commitments. Put anything I need to fill in as [square brackets].",
  "Keep source links as short clickable names.",
];

/** Turns a request into plain text for Copilot. Email addresses are never included. */
export function buildRequest({ task, context = [], material, notes, output }: Request): string {
  const lines = [task, ""];
  const facts = context.filter((line): line is string => !!line && !!line.trim());
  if (facts.length) lines.push("Context:", ...facts.map(line => `- ${line}`), "");
  const head = [...lines];
  const tail = [...(notes?.trim() ? ["My notes (follow these):", notes.trim(), ""] : []), "Instructions:", ...RULES.map(rule => `- ${rule}`), `- ${output}`];
  if (material?.trim()) {
    const room = MAX_REQUEST - head.join("\n").length - tail.join("\n").length - 200;
    const text = material.trim().length > room ? `${material.trim().slice(0, Math.max(0, room)).trimEnd()}\n[…trimmed]` : material.trim();
    head.push("Material from my Mission Control Interstellar dashboard:", text, "");
  }
  return hideEmails([...head, ...tail].join("\n")).slice(0, MAX_REQUEST);
}

/** A document's sections as plain text with markdown links kept. */
export function docText(doc: DocContent): string {
  return doc.sections.map(section => [section.heading && `## ${section.heading}`, section.text].filter(Boolean).join("\n")).filter(Boolean).join("\n\n");
}

const quoted = (text: string) => `"${text.replace(/\s+/g, " ").trim()}"`;

export function replyRequest(item: Pick<WorkItem, "title" | "summary" | "url">, notes: string): string {
  return buildRequest({
    task: `Draft a reply to the email ${quoted(item.title)}.`,
    context: [`Email: ${item.summary}`, item.url && `Link: [Open the email](${item.url})`, "Find the full thread in my Outlook by its subject and answer every question it asks."],
    notes,
    output: "Give me only the reply body, friendly and professional, under 150 words unless the questions need more. I'll paste it into the draft I saved in Outlook.",
  });
}

export function followUpRequest(person: Pick<WorkItem, "title" | "summary" | "url">, notes: string): string {
  return buildRequest({
    task: `Draft a short, polite follow-up email to ${person.title}.`,
    context: [person.summary, person.url && `My original email: [Open it](${person.url})`, "Check my Outlook for the original email and any reply before drafting."],
    notes,
    output: "Give me a subject line and the email body (under 100 words), with one clear ask and a suggested date.",
  });
}

export function awayRequest(away: Pick<WorkItem, "title" | "relatedMeeting" | "date" | "endDate" | "summary">, notes: string): string {
  return buildRequest({
    task: `Draft a short email to ${away.title}, who is away during my meeting ${quoted(away.relatedMeeting)} (they organize it).`,
    context: [`Away ${away.date}${away.endDate ? ` to ${away.endDate}` : ""}.`, away.summary && `Their automatic reply: ${quoted(away.summary)}`, "Look up the meeting in my calendar for its time and attendees."],
    notes,
    output: "Give me a subject line and body (under 100 words) asking whether to reschedule, go ahead without them, or name a stand-in; suggest two new times if useful.",
  });
}

export function teamsReplyRequest(chat: { person: string; chatTitle: string; group: boolean; preview: string }, notes: string): string {
  return buildRequest({
    task: `Draft a reply to ${chat.person}'s latest Teams message${chat.group ? ` in ${quoted(chat.chatTitle)}` : ""}.`,
    context: [`Their message: ${quoted(chat.preview)}`, "Read the rest of that Teams chat for context before drafting."],
    notes,
    output: "Give me only the reply text, short and conversational, ready to paste into Teams.",
  });
}

const BRIEF_TASK: Record<BriefKind, (title: string) => { task: string; output: string }> = {
  prep: title => ({ task: `Write a one-page prep brief for my meeting ${quoted(title.replace(/^Prep brief[:\s-–]*/i, ""))}.`,
    output: "Use these headings: Purpose and what I need from it; Where things stand; Risks, decisions and open questions; Questions to ask; What to bring or do beforehand. Short bullets." }),
  project: title => ({ task: `Write a short project briefing for me on ${quoted(title.replace(/^Project briefing[:\s-–]*/i, ""))}.`,
    output: "Use these headings: Status in one line (only as stated in the material); What changed recently; Risks and blockers; Upcoming deadlines and meetings; My next actions. Short bullets." }),
  status: () => ({ task: "Write my weekly status report from the material below.",
    output: "Start with a 3-bullet summary, then one section per project: Status (Green/Amber/Red only if stated), Progress this week, Risks and blockers, Next week. Short bullets." }),
  catchup: () => ({ task: "Give me a Teams catch-up from the material below.",
    output: "List, most urgent first: who is waiting on me and what they need; where I was @mentioned; decisions, risks and deadlines in my project teams. Short bullets with links." }),
  document: title => ({ task: `Write a polished document titled ${quoted(title)} from my outline below.`,
    output: "Keep my structure and headings, expand each point into clear prose or bullets, and give me the finished text ready to paste into Word." }),
};

/** A Copilot request built from a document the app already assembled (brief, report, catch-up or outline). */
export function briefRequest(kind: BriefKind, doc: DocContent, notes: string): string {
  const { task, output } = BRIEF_TASK[kind](doc.title);
  return buildRequest({ task, material: docText(doc), notes, output });
}

/** Quick actions → Email draft. */
export function newEmailRequest(to: string, subject: string, notes: string): string {
  return buildRequest({
    task: `Draft an email${to ? ` to ${to}` : ""}${subject.trim() ? ` with the subject ${quoted(subject)}` : ""}.`,
    context: ["Check my recent emails, chats and meetings with them for anything relevant."],
    notes,
    output: "Give me only the email body, clear and friendly, under 150 words unless my notes need more.",
  });
}

/** Files → "Draft with Copilot": summarize a file and pull out what it needs from me. */
export function fileRequest(file: { name: string; url: string; reasons: string[]; contexts: { label: string; url?: string }[] }, notes: string): string {
  return buildRequest({
    task: `Summarize the file ${quoted(file.name)} for me and tell me what it needs from me.`,
    context: [`File: [${file.name.replace(/[[\]]/g, "")}](${file.url})`, `Why it matters: ${file.reasons.join("; ")}`, ...file.contexts.slice(0, 3).map(context => `Found in: ${context.url ? `[${context.label.replace(/[[\]]/g, "")}](${context.url})` : context.label}`), "Open the file through its link (I have access to it)."],
    notes,
    output: "Give me: a 3-line summary; decisions and deadlines in it; anything that asks me to act (with who and when); and 2-3 questions I should raise. Short bullets.",
  });
}
