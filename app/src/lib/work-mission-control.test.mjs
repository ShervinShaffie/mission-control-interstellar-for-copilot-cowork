// Local checks only: connectors are intercepted; nothing here is live proof.
import { test, expect, mock } from "bun:test";
import { Window } from "happy-dom";
const dom = new Window({ url: "https://example.com/" });
for (const name of ["window", "document", "navigator", "DOMParser"]) Object.defineProperty(globalThis, name, { value: dom[name], configurable: true });
let calls = [], respond;
mock.module("@microsoft/managed-apps/data", () => ({ getClient: () => ({ executeAsync: async request => { calls.push(request.connectorOperation); return respond(request.connectorOperation); } }) }));
const docs = await import("./documents.ts");
const prefs = await import("./dashboard-preferences.ts");
const plans = await import("./search-plan.ts");
const actions = await import("./actions.ts");
const sugg = await import("./suggestions.ts");
const runners = await import("./action-runners.ts");
const ok = data => ({ success: true, data });

test("Word files are valid ZIP packages with the document parts", () => {
  const bytes = docs.buildDocx({ title: "Plan & <notes>", sections: [{ heading: "Goals", text: "One\n\nTwo" }] });
  expect([...bytes.slice(0, 4)]).toEqual([0x50, 0x4b, 0x03, 0x04]);
  const text = new TextDecoder().decode(bytes);
  for (const part of ["[Content_Types].xml", "word/document.xml", "word/styles.xml", "_rels/.rels"]) expect(text).toContain(part);
  expect(text).toContain("Plan &amp; &lt;notes&gt;");
  expect(docs.crc32(new TextEncoder().encode("123456789"))).toBe(0xcbf43926);
  expect(docs.fileName("Q4: plan/report", "docx")).toMatch(/^Q4 plan report \d{4}-\d{2}-\d{2}\.docx$/);
  expect(docs.buildXlsxBase64([{ name: "Tracker", rows: [["A"], ["b"]] }]).slice(0, 4)).toBe("UEsD");
});

test("All 9 components start on the dashboard, in order, with every search on", () => {
  const defaults = prefs.defaultPreferences();
  expect(defaults.visible).toEqual(["today", "ask", "projects", "calendar", "email", "teams", "files", "onenote", "insights"]);
  // Full-width or same-size pairs: 4 wide, then email+teams, work+onenote, then files and insights wide.
  expect(prefs.componentCatalog.filter(item => item.wide).map(item => item.id)).toEqual(["today", "ask", "projects", "calendar", "insights"]);
  // People and Create were folded into other components; old saved layouts drop them.
  expect(prefs.parsePreferences({ version: 2, visible: ["today", "people", "create", "projects", "ask"] }).visible).toEqual(["today", "ask", "projects", "files"]);
  // An older saved layout gets Projects and Ask once and the new balanced order once; later choices stick.
  const older = prefs.parsePreferences({ version: 2, visible: ["calendar", "today", "email"] });
  expect(older.visible).toEqual(["today", "ask", "projects", "calendar", "email", "files"]);
  expect(prefs.parsePreferences({ ...older, visible: ["calendar", "today"] }).visible).toEqual(["calendar", "today"]);
  expect(older.targetsSet).toBe(false);
  expect(Object.values(defaults.searches).every(Boolean)).toBe(true);
  const parsed = prefs.parsePreferences({ ...defaults, searches: { replies: false }, targets: { Focus: 30, Bad: 400 } });
  expect(parsed.searches).toEqual({ priorities: true, meetings: true, replies: false, collaboration: true, pinned: true });
  expect(parsed.targets).toEqual({ Focus: 30 });
});

test("Copilot searches run only when switched on and used by a visible component", () => {
  const on = { priorities: true, meetings: true, replies: true, collaboration: true };
  expect(plans.plannedSearches(plans.searchPlan(["calendar", "teams"], on))).toEqual(["meetings", "collaboration"]);
  expect(plans.plannedSearches(plans.searchPlan(["email"], { ...on, replies: false }))).toEqual([]);
  expect(plans.plannedSearches(plans.searchPlan(["today"], on))).toHaveLength(4);
});

test("Actions rank most important first, free before credits on ties", () => {
  const run = async () => ({ message: "", links: [] });
  const ranked = actions.rankActions([
    { id: "a", title: "B", priority: 50, cost: "credits", run }, { id: "b", title: "A", priority: 50, cost: "free", run }, { id: "c", title: "C", priority: 100, cost: "credits", run },
  ]);
  expect(ranked.map(item => item.id)).toEqual(["c", "b", "a"]);
});

const at = (h, m = 0, dayOffset = 0) => { const d = new Date(); d.setDate(d.getDate() + dayOffset); d.setHours(h, m, 0, 0); return d; };
const ev = (id, h1, h2, extra = {}) => ({ id, subject: `Meeting ${id}`, startWithTimeZone: at(h1, 0, 1).toISOString(), endWithTimeZone: at(h2, 0, 1).toISOString(), responseType: "accepted", requiredAttendees: "a@example.com;", showAs: "busy", webLink: "https://outlook.office.com/x", ...extra });

test("Calendar suggestions: invites, overlaps, focus gaps; nothing runs until asked", () => {
  calls = [];
  const events = [ev("1", 9, 10), ev("2", 9, 11, { responseType: "notResponded" }), ev("3", 13, 14)];
  const list = sugg.calendarActions({ now: at(10, 0, 0), calendarId: "cal" }, events, [], []);
  expect(list.some(item => item.id === "invite-2" && item.cost === "free" && item.choices.length === 3)).toBe(true);
  expect(list.some(item => item.id.startsWith("overlap-"))).toBe(true);
  expect(list.some(item => item.id.startsWith("focus-") && /11:00/.test(item.title))).toBe(true);
  expect(list.find(item => item.id.startsWith("brief-")).cost).toBe("free");
  expect(calls).toHaveLength(0);
  expect(sugg.freeGaps(events, at(0, 0, 1), at(0, 0, 1), 60).map(gap => gap.start.getHours())).toEqual([11, 14]);
});

test("Day plan orders overdue first and uses only loaded data", () => {
  const steps = sugg.buildDayPlan([], [{ kind: "action", title: "Expenses", date: "2020-01-01", next: "", summary: "", status: "", severity: "", endDate: "", email: "", url: "", evidence: "x", relatedMeeting: "" }], [], at(8));
  expect(steps[0].tone).toBe("urgent"); expect(steps[0].text).toContain("Expenses");
});

test("Respond to invite calls the right Outlook action and records links", async () => {
  calls = []; respond = () => ok(undefined);
  const result = await runners.respondToInvite(ev("9", 9, 10, { responseType: "notResponded" }), "decline", "Conflict");
  expect(calls[0].operationName).toBe("RespondToEvent_V2");
  expect(calls[0].parameters).toMatchObject({ event_id: "9", response: "decline", body: { Comment: "Conflict", SendResponse: true } });
  expect(result.links[0].url).toBe("https://outlook.office.com/x");
});

test("A failed action reports a clear error and changes nothing else", async () => {
  calls = []; respond = () => ({ success: false, error: { status: 442 } });
  const error = await runners.flagMessage("m1", "Hello").catch(value => value);
  expect(error.name).toBe("ActionError"); expect(calls).toHaveLength(1);
});

test("Teams follow-ups keep only chats where someone else wrote last, recently", async () => {
  const waitingLib = await import("./teams-waiting.ts");
  const recent = new Date(Date.now() - 3600e3).toISOString(), old = new Date(Date.now() - 20 * 864e5).toISOString();
  const msg = (id, sender, senderId, created, text) => ({ id, createdDateTime: created, messageType: "message", from: { user: { id: senderId, displayName: sender } }, body: { content: `<p>${text}</p>` } });
  const pages = {
    a: [msg("1", "Jordan Miwa", "jordan", recent, "Can you review the outline by Friday?")],
    b: [msg("2", "Alex", "me", recent, "Done, thanks")],
    c: [msg("3", "Lydia Bauer", "lydia", old, "Old note")],
  };
  calls = []; respond = op => ok({ value: pages[op.parameters["chat-id"] ?? op.parameters.chatId ?? Object.values(op.parameters)[0]] ?? [] });
  const rows = await waitingLib.loadWaitingChats([{ id: "a", topic: "Chat without a title" }, { id: "b", topic: "Chat without a title" }], [{ id: "c", topic: "Launch" }], "me");
  expect(rows.map(row => [row.person, row.chatTitle])).toEqual([["Jordan", "Chat with Jordan"]]);
  expect(rows[0].preview).toBe("Can you review the outline by Friday?");
  const actions = sugg.teamsActions(rows);
  expect(actions[0].title).toBe("Reply to Jordan in Teams");
  expect(actions.some(item => item.title.includes("Chat without a title"))).toBe(false);
});

test("Today ranks across every source, merges duplicates and never calls a service", async () => {
  const attention = await import("./attention.ts");
  calls = [];
  const now = at(10, 0, 0), y = new Date(now); y.setDate(y.getDate() - 2);
  const ymd = d => d.toLocaleDateString("en-CA");
  const items = attention.buildAttention({
    now, calendarId: "cal", events: [], unread: [{ id: "m1", subject: "Approve budget", importance: "high", from: "Jane Doe" }], flagged: [],
    waiting: [{ chatId: "c1", chatTitle: "Chat with Jordan", person: "Jordan", preview: "Can you review?", created: new Date(now.getTime() - 3600e3).toISOString(), group: false }],
    tasks: [{ id: "t1", title: "Annterra quality agreement outline", list: "Tasks", listId: "L", due: ymd(y) }],
    pages: [], replies: [], meetings: [], collaboration: [],
    priorities: [
      { kind: "action", title: "Annterra quality agreement outline", date: ymd(y), summary: "Draft due", next: "", status: "", severity: "", endDate: "", email: "", url: "https://contoso.sharepoint.com/sites/x/doc.docx", evidence: "Teams post", relatedMeeting: "" },
      { kind: "project", title: "Fabrikam Upgrade (OPS-SECURE)", status: "RED", summary: "Gated on oversharing", next: "Close X-04", date: "", severity: "", endDate: "", email: "", url: "", evidence: "email", relatedMeeting: "" },
    ],
  });
  expect(calls).toHaveLength(0);
  const merged = items.find(item => item.title === "Annterra quality agreement outline");
  expect(merged.horizon).toBe("overdue");
  expect(merged.sources.sort()).toEqual(["sharepoint", "todo"]);
  expect(items[0].horizon).toBe("overdue");
  expect(items.some(item => item.title.startsWith("Project needs attention: Fabrikam Upgrade") && item.cost === "free")).toBe(true);
  expect(items.some(item => item.title === "Jordan is waiting on you in Teams")).toBe(true);
  expect(items.some(item => item.sources.includes("outlook") && item.title.includes("Approve budget"))).toBe(true);
});

test("Projects pick up their meetings, prep notes, deadlines and To Do tasks", async () => {
  const projects = await import("./projects.ts");
  const now = at(10, 0, 0);
  const meeting = ev("p1", 14, 15, { subject: "OPS-SECURE checkpoint" });
  const views = projects.buildProjects({ now, events: [meeting, ev("p2", 9, 10, { subject: "Lunch" })], tasks: [{ id: "t", title: "OPS-SECURE label coverage", list: "Tasks", listId: "L" }],
    projects: [{ kind: "project", title: "Fabrikam Upgrade & Security Migration (OPS-SECURE)", status: "AMBER", summary: "", next: "Confirm owners", date: "", severity: "", endDate: "", email: "", url: "", evidence: "x", relatedMeeting: "" }],
    work: [], meetings: [{ kind: "prep", title: "OPS-SECURE checkpoint", summary: "Bring the oversharing numbers", next: "Review X-04", date: "", severity: "", endDate: "", email: "", url: "", evidence: "x", relatedMeeting: "" }] });
  expect(views[0].health).toBe("watch");
  expect(views[0].meetings.map(m => m.event.id)).toEqual(["p1"]);
  expect(views[0].meetings[0].prep.summary).toContain("oversharing");
  expect(views[0].tasks).toHaveLength(1);
  expect(views[0].actions.map(a => a.title)).toEqual(["Prep brief for OPS-SECURE checkpoint", "Add the next deadline to To Do", "Project briefing"]);
  expect(views[0].actions.every(a => a.cost === "free")).toBe(true);
});

test("Suggested time targets come from category names and add up to 100%", async () => {
  const { suggestedTargets } = await import("./time-targets.ts");
  const targets = suggestedTargets(["Focus time", "Customer meetings", "Admin", "Blue category"], [{ name: "Team sync", share: 40 }, { name: "Uncategorized", share: 30 }]);
  expect(Object.keys(targets).sort()).toEqual(["Admin", "Customer meetings", "Focus time", "Team sync"]);
  expect(Object.values(targets).reduce((a, b) => a + b, 0)).toBe(100);
  expect(Object.values(targets).every(value => value >= 5 && value % 5 === 0)).toBe(true);
  expect(targets["Team sync"]).toBeGreaterThan(targets.Admin);
  expect(suggestedTargets([], [])).toEqual({});
});

test("Every priority links to its original item", async () => {
  const attention = await import("./attention.ts");
  const waitingLib = await import("./teams-waiting.ts");
  const link = waitingLib.teamsMessageLink("19:abc@thread.v2", "1791387000000");
  expect(link.startsWith("https://teams.microsoft.com/l/message/19%3Aabc%40thread.v2/1791387000000?context=")).toBe(true);
  const now = at(10, 0, 0);
  const items = attention.buildAttention({
    now, calendarId: "cal", events: [], unread: [], flagged: [], tasks: [], pages: [], meetings: [], collaboration: [],
    waiting: [{ chatId: "19:abc@thread.v2", chatTitle: "Chat with Sam", person: "Sam", preview: "Will do", created: new Date(now.getTime() - 7 * 864e5 + 3600e3).toISOString(), group: false, messageUrl: link }],
    replies: [{ kind: "reply", title: "Budget question", summary: "Asks for numbers", next: "Reply", status: "", severity: "", date: "", endDate: "", email: "", url: "", evidence: "email", relatedMeeting: "" }],
    priorities: [],
  });
  expect(items.every(item => item.open && item.open.url.startsWith("https://"))).toBe(true);
  expect(items.find(item => item.sources.includes("teams")).open.url).toBe(link);
  expect(items.find(item => item.title.startsWith("Reply needed")).open.label).toBe("Find it in Outlook");
});

test("Dismissed items wait for the chosen days, then stay hidden for good", async () => {
  const d = await import("./dismissed.ts");
  d.resetDismissedForTests();
  expect(d.dismissKey("att-teams-19:abc")).toBe(d.dismissKey("teams-19:abc"));
  expect(d.dismissKey("pa-proj-brief-X")).toBe("proj-brief-X");
  const now = Date.now(), day = 864e5;
  const store = { retention: 3, items: [{ key: "a", title: "A", section: "today", at: now - 1 * day }, { key: "b", title: "B", section: "email", at: now - 4 * day }], hidden: [] };
  const next = d.expireDismissed(store, now);
  expect(next.items.map(item => item.key)).toEqual(["a"]);
  expect(next.hidden).toEqual(["b"]);
  d.dismiss("att-reply-Budget", "Reply needed: Budget", "today");
  d.dismiss("invite-9", "Answer the invite", "calendar");
  d.restore("invite-9");
});

test("Copilot text shows short named sources instead of long links, in the app and in Word", async () => {
  const c = await import("./citations.ts");
  const doc = { title: "Prep brief - Northwind launch risk register review", sections: [
    { heading: "Prep brief - Northwind launch risk register review", text: "Meeting: Wednesday 1:00 PM. [1](https://teams.microsoft.com/l/meeting/details?eventId=AAMk\\_Q6%3d)" },
    { heading: "##### Current readiness", text: "Two workstreams are at risk. [3](https://teams.microsoft.com/l/message/19:abc@thread.v2/1791?context=%7B%22contextType%22:%22chat%22%7D)[5](https://teams.microsoft.com/l/message/19:x@thread.tacv2/1790?teamName=Northwind%20Launch%20Readiness&channelName=General)\n- **Lead times** confirmed? [5](https://teams.microsoft.com/l/message/19:x@thread.tacv2/1790?teamName=Northwind%20Launch%20Readiness&channelName=General)\n- Plan [Project Plan](https://contoso.sharepoint.com/sites/x/_layouts/15/Doc.aspx?file=Project%20Plan.xlsx)" },
  ] };
  const clean = c.cleanDoc(doc);
  expect(clean.sections[0].heading).toBe("");
  expect(clean.sections[1].heading).toBe("Current readiness");
  expect(clean.sources.map(s => s.name)).toEqual(["Meeting", "Teams chat", "Northwind Launch Readiness › General", "Project Plan"]);
  const para = clean.sections[1].blocks[0];
  expect(para.inlines[1].kind).toBe("cites");
  expect(para.inlines[1].cites).toHaveLength(2);
  expect(clean.sections[1].blocks[1].kind).toBe("bullet");
  expect(c.plainText(clean.sections[1].blocks)).not.toContain("https://");
  expect(c.plainText(clean.sections[1].blocks)).not.toContain("**");
  const text = new TextDecoder().decode(docs.buildDocx(doc));
  expect(text).toContain('relationships/hyperlink" Target=');
  expect(text).toContain("Northwind Launch Readiness › General");
  expect(text).not.toContain("[1](");
});

test("Escaped citations (\\[1\\](url)) also become short source names", async () => {
  const c = await import("./citations.ts");
  const links = await import("./source-links.ts");
  const raw = "coming week. \\[1\\](https://teams.microsoft.com/l/meeting/details?eventId=AAMk\\_Q6%3d)\\[2\\](https://teams.microsoft.com/l/message/19:1c@thread.tacv2/178?teamName=Northwind%20Launch%20Readiness&channelName=General)";
  const doc = c.cleanDoc({ title: "T", sections: [{ heading: "H", text: raw }] });
  expect(doc.sources.map(s => s.name)).toEqual(["Meeting", "Northwind Launch Readiness › General"]);
  expect(c.plainText(doc.sections[0].blocks)).toBe("coming week. (Meeting; Northwind Launch Readiness › General)");
  expect(links.parseRichText(raw).filter(s => s.kind === "link")).toHaveLength(2);
});

test("History keeps every result, merges this browser with OneDrive, and is searchable", async () => {
  const h = await import("./history.ts");
  h.resetHistoryForTests();
  const remote = [{ id: "r1", at: Date.now() - 864e5, kind: "document", title: "Prep brief - Northwind launch risk register review", status: "done", cost: "credits", doc: { title: "Prep brief", sections: [{ heading: "Depot contingency", text: "Kim owns the January depot contingency." }] } }];
  calls = [];
  respond = op => op.operationName === "GetFileMetadataByPath" ? ok({ Id: "f1", Path: "/Work Mission Control/Activity history.json" })
    : op.operationName === "GetFileContentByPath" ? ok(JSON.stringify({ entries: remote })) : ok({ Id: "f1" });
  await h.initHistory("11111111-2222-3333-4444-555555555555", "https://contoso-my.sharepoint.com/personal/alex_contoso_com/");
  h.recordHistory({ kind: "action", title: "Answer the invite: Budget review", section: "calendar", cost: "free", status: "done", choice: "Accept", message: "Accepted." });
  h.recordHistory({ kind: "action", title: "Draft a reply: Q4 numbers", section: "email", cost: "credits", status: "failed", error: "Saving the draft failed." });
  const all = h.searchHistory(h.mergeHistory(), "", "all");
  expect(all).toEqual([]); // mergeHistory() with no lists is empty
  const found = h.searchHistory(h.parseHistory(JSON.stringify(remote)), "kim depot");
  expect(found.map(e => e.id)).toEqual(["r1"]);
  expect(calls.map(c => c.operationName)).toEqual(["GetFileMetadataByPath", "GetFileContentByPath"]);
  expect(h.mergeHistory([{ ...remote[0], title: "old", at: 1 }], remote)[0].title).toContain("Prep brief");
  expect(h.parseHistory("not json")).toEqual([]);
});

test("Connection problems are explained in plain terms", async () => {
  const health = await import("./connection-health.ts");
  expect(health.classify({ status: 401, message: "Unauthorized" })).toBe("signin");
  expect(health.classify(new Error("Login unsuccessful"))).toBe("signin");
  expect(health.classify({ status: 442 })).toBe("blocked");
  expect(health.classify(new Error("boom"))).toBe("error");
  calls = []; respond = () => ({ success: false, error: { status: 401 } });
  await health.runCheck("todo", health.FREE_CHECKS.todo);
  respond = () => ok([]);
  await health.runCheck("onenote", health.FREE_CHECKS.onenote);
  expect(calls.map(c => c.operationName)).toEqual(["GetAllTodoListsV2", "GetNotebooks"]);
});

test("Every connection has a sign-in route to the person's own Connections page", async () => {
  const health = await import("./connection-health.ts");
  expect(Object.keys(health.CONNECTOR_NAMES)).toHaveLength(8);
  expect(health.connectionsPageUrl("00000000-1111-2222-3333-444444444444")).toBe("https://make.powerapps.com/environments/00000000-1111-2222-3333-444444444444/connections");
  expect(health.connectionsPageUrl(undefined)).toBe("https://make.powerapps.com/connections");
  expect(health.connectionsPageUrl("bad/../id")).toBe("https://make.powerapps.com/connections");
});

test("Direct rules: deadlines in team posts, projects with status, urgent @mentions", async () => {
  const d = await import("./direct-signals.ts");
  const written = new Date(2026, 9, 9, 10);
  const found = d.findDeadlines("Annterra quality agreement outline is due 23 Oct. I'll circulate a draft. Denver confirmation is due Friday.", written);
  expect(found.map(f => f.date.toDateString())).toEqual([new Date(2026, 9, 23).toDateString(), new Date(2026, 9, 16).toDateString()]);
  const now = new Date(2026, 9, 14, 9);
  const post = (text, extra = {}) => ({ id: Math.random().toString(), team: "Northwind Launch Readiness", teamId: "t", channel: "General", channelId: "c", author: "Kim", authorId: "k", created: new Date(2026, 9, 13, 9), text, url: "https://teams.microsoft.com/l/message/x", mentionsMe: false, important: false, ...extra });
  const items = d.prioritiesFromTeams([{ id: "t", name: "Northwind Launch Readiness", description: "", channels: [{ id: "c", name: "General", webUrl: "https://t", filesUrl: "https://f" }],
    posts: [post("Status, week ending 2 Oct: RED. Contingency plan due 16 Oct."), post("Can you review the deck?", { mentionsMe: true })] }], "me", now);
  const project = items.find(i => i.kind === "project");
  expect(project.status).toBe("RED");
  expect(project.next).toContain("Contingency plan due 16 Oct");
  expect(items.some(i => i.kind === "action" && i.date === "2026-10-16" && i.evidence === "Teams › Northwind Launch Readiness › General")).toBe(true);
  expect(items.some(i => i.kind === "urgent" && i.title === "Kim mentioned you in Northwind Launch Readiness › General")).toBe(true);
});

test("Direct rules: emails needing a reply and emails waiting on others", async () => {
  const d = await import("./direct-signals.ts");
  const now = new Date(2026, 9, 14, 9), day = 864e5;
  const row = (id, extra) => ({ id, subject: "Budget", from: "jordan@contoso.com", to: "alex@contoso.com", cc: "", received: new Date(now.getTime() - day), preview: "Can you approve the budget by Friday?", conversationId: id, importance: "normal", isRead: false, ...extra });
  const mail = {
    inbox: [row("a"), row("b", { subject: "Newsletter", preview: "Read more. Unsubscribe here", conversationId: "b" }), row("c", { subject: "Answered?", conversationId: "c" }), row("d", { from: "noreply@contoso.com", conversationId: "d" }), row("e", { subject: "Accepted: Sync", conversationId: "e" })],
    sent: [row("s1", { from: "alex@contoso.com", to: "jordan@contoso.com", conversationId: "c", received: new Date(now.getTime() - 3600e3) }), row("s2", { from: "alex@contoso.com", to: "lydia@contoso.com", subject: "Deck review", preview: "Could you review the deck?", conversationId: "z", received: new Date(now.getTime() - 4 * day) })],
  };
  const replies = d.repliesFromMail(mail, "alex@contoso.com", now);
  expect(replies.map(r => r.title)).toEqual(["Budget"]);
  expect(replies[0].url).toBe("https://outlook.office.com/mail/deeplink/read/a");
  const waiting = d.waitingOnOthers(mail, "alex@contoso.com", now);
  expect(waiting.map(w => w.email)).toEqual(["lydia@contoso.com"]);
});

test("Free briefs use only what was read, with short source links", async () => {
  const b = await import("./briefs.ts");
  const event = { id: "e1", subject: "Northwind risk review", startWithTimeZone: new Date(2026, 9, 14, 13).toISOString(), organizer: "jordan@contoso.com", requiredAttendees: "jordan@contoso.com;lydia@contoso.com", body: "<p>Review R-04 and depot contingency</p>", webLink: "https://outlook.office.com/x" };
  const doc = b.meetingBrief(event, undefined, { id: "t", name: "Northwind Launch Readiness", description: "", channels: [], posts: [{ id: "p", team: "Northwind Launch Readiness", teamId: "t", channel: "General", channelId: "c", author: "Kim", authorId: "k", created: new Date(2026, 9, 13), text: "Depot readiness is 61% and at risk.", url: "https://teams.microsoft.com/l/message/p", mentionsMe: false, important: false }] }, [], "Ask about R-04");
  const text = doc.sections.map(s => `${s.heading}\n${s.text}`).join("\n");
  expect(text).toContain("Review R-04 and depot contingency");
  expect(text).toContain("[Northwind Launch Readiness › General](https://teams.microsoft.com/l/message/p)");
  expect(doc.sections.at(-1)).toEqual({ heading: "My notes", text: "Ask about R-04" });
});

test("People are never shown by email address: first name when known, else the alias", async () => {
  const p = await import("./privacy.ts");
  expect(p.firstName("pat.lee@contoso.com")).toBe("pat.lee");
  p.rememberFirstName("pat.lee@contoso.com", "Pat");
  expect(p.firstName("pat.lee@contoso.com")).toBe("Pat");
  expect(p.firstName("Doe, Jane")).toBe("Jane");
});

test("Draft with Copilot builds a clear request from loaded data, with no addresses and no calls", async () => {
  const cp = await import("./copilot-prompts.ts");
  const reply = cp.replyRequest({ title: "Budget sign-off", summary: "Pat, Tue Oct 13: Can you approve by Friday? Contact pat@contoso.com", url: "https://outlook.office.com/mail/deeplink/read/x" }, "Yes, approved");
  expect(reply).toContain('Draft a reply to the email "Budget sign-off"');
  expect(reply).toContain("Yes, approved");
  expect(reply).toContain("[square brackets]");
  expect(reply).not.toContain("pat@contoso.com");
  const big = cp.briefRequest("prep", { title: "Prep brief: Launch review", sections: [{ heading: "Agenda", text: "x".repeat(20000) }] }, "");
  expect(big.length).toBeLessThanOrEqual(cp.MAX_REQUEST);
  expect(big).toContain('prep brief for my meeting "Launch review"');
  expect(big).toContain("[…trimmed]");
});

test("Files: links and attachments become one ranked list, most important first", async () => {
  const f = await import("./files.ts");
  const html = '<p>See <a href="https://contoso.sharepoint.com/:w:/s/launch/EabC">Launch plan</a> and https://contoso.sharepoint.com/sites/launch/Shared%20Documents/Budget.xlsx and <a href="https://contoso.sharepoint.com/sites/launch">the site</a></p>';
  const refs = f.filesFromHtml(html);
  expect(refs.map(ref => ref.name)).toEqual(["Launch plan", "Budget.xlsx"]);
  expect(f.fileKind("Launch plan", refs[0].url)).toBe("word");
  expect(f.filesFromAttachments([{ contentType: "reference", name: "Deck.pptx", contentUrl: "https://contoso.sharepoint.com/sites/x/Deck.pptx" }, { contentType: "image/png" }]).map(ref => ref.name)).toEqual(["Deck.pptx"]);
  const now = new Date(2026, 9, 14, 12);
  const ranked = f.rankFiles([
    { file: { name: "Budget.xlsx", url: "https://a.sharepoint.com/b/Budget.xlsx" }, origin: "trending", context: "Finance" },
    { file: { name: "Launch plan.docx", url: "https://a.sharepoint.com/b/Launch plan.docx" }, origin: "team", who: "Kim", when: new Date(2026, 9, 13), context: "Launch › General" },
    { file: { name: "Launch plan.docx", url: "https://a.sharepoint.com/b/Launch plan.docx" }, origin: "reply", who: "Pat", when: new Date(2026, 9, 14), context: "Email: Sign-off" },
  ], now);
  expect(ranked[0].name).toBe("Launch plan.docx");
  expect(ranked[0].reasons.length).toBe(2);
  expect(ranked[0].reasons[0]).toContain("waiting on your reply");
  expect(ranked.at(-1).name).toBe("Budget.xlsx");
});
