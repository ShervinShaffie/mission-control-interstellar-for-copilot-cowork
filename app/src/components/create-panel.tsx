import { useState } from "react";
import { CopilotHandoff } from "@/components/copilot-handoff";
import { briefRequest } from "@/lib/copilot-prompts";
import { FilePlus2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { DocumentResult, ResultLinks } from "@/components/action-list";
import { useCreatedItems } from "@/lib/actions";
import { recordHistory } from "@/lib/history";
import { FORMAT_LABELS, type DocContent, type DocFormat } from "@/lib/documents";

/** Data tab: everything created or changed from this dashboard during this session. */
export function CreatedList() {
  const items = useCreatedItems();
  if (!items.length) return <div className="rounded-2xl border border-dashed border-border p-6 text-center text-sm text-muted-foreground">
    <FilePlus2 aria-hidden="true" className="mx-auto mb-2 size-6" />Nothing created yet this session. Drafts, events, tasks, pages and files you create from any Actions tab appear here with links.
  </div>;
  return <ul className="space-y-3">{items.map(item => <li key={item.id} className="rounded-2xl border border-border bg-card p-4">
    <p className="break-words font-semibold">{item.title}</p>
    <p className="text-xs text-muted-foreground">{item.kind} · {new Date(item.at).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })}</p>
    <ResultLinks links={item.links} />
  </li>)}</ul>;
}

/** "## Heading" lines start a new section; everything else is that section's text. */
function textToDoc(title: string, body: string): DocContent {
  const sections: DocContent["sections"] = [];
  let current = { heading: "", text: "" };
  for (const line of body.split(/\r?\n/)) {
    const heading = line.match(/^#{1,3}\s+(.+)/);
    if (heading) { if (current.heading || current.text.trim()) sections.push({ ...current, text: current.text.trim() }); current = { heading: heading[1].trim(), text: "" }; }
    else current.text += `${line}\n`;
  }
  if (current.heading || current.text.trim()) sections.push({ ...current, text: current.text.trim() });
  return { title, sections: sections.length ? sections : [{ heading: "", text: "" }] };
}

/** Make a Word, Excel or PDF document from text you write, then download it. */
export function CreateForm({ mySite }: { mySite?: string }) {
  const [format, setFormat] = useState<DocFormat>("docx");
  const [title, setTitle] = useState("");
  const [body, setBody] = useState("");
  const [doc, setDoc] = useState<DocContent>();
  function make() {
    const name = title.trim() || "Untitled document";
    const made = textToDoc(name, body);
    setDoc(made);
    recordHistory({ kind: "document", title: name, section: "create", cost: "free", status: "done", doc: made });
  }
  return <div className="rounded-2xl border border-border bg-card p-4">
    <h3 className="font-semibold">New document</h3>
    <fieldset className="mt-3"><legend className="text-sm font-medium">Type</legend><div className="mt-2 flex flex-wrap gap-2">{(["docx", "xlsx", "pdf"] as DocFormat[]).map(value =>
      <Button key={value} size="sm" variant={format === value ? "default" : "outline"} aria-pressed={format === value} onClick={() => setFormat(value)} className="rounded-full">{FORMAT_LABELS[value]}</Button>)}</div></fieldset>
    <div className="mt-3"><Label htmlFor="create-title">Title</Label><Input id="create-title" value={title} maxLength={120} onChange={event => setTitle(event.target.value)} placeholder="For example: Q4 launch plan" className="mt-1.5" /></div>
    <div className="mt-3"><Label htmlFor="create-body">Content</Label>
      <Textarea id="create-body" value={body} maxLength={20000} onChange={event => setBody(event.target.value)} className="mt-1.5 min-h-40" placeholder={"## Background\nWhy this matters…\n\n## Goals\n- First goal\n- Second goal"} />
      <p className="mt-1 text-xs text-muted-foreground">Start a line with “## ” for a heading and “- ” for a bullet.</p></div>
    <div className="mt-3 flex flex-wrap items-start gap-3"><Button className="rounded-full" onClick={make}><FilePlus2 aria-hidden="true" />Preview the document</Button>
      <CopilotHandoff title={`Document: ${title.trim() || "Untitled document"}`} section="ask" disabled={!body.trim()} build={() => briefRequest("document", textToDoc(title.trim() || "Untitled document", body), "")} /></div>
    <p className="mt-1 text-xs text-muted-foreground">Want Copilot to write it up? Draft with Copilot sends your outline; paste its answer back into Content and preview again.</p>
    {doc && <DocumentResult doc={doc} mySite={mySite} formats={[format]} />}
  </div>;
}
