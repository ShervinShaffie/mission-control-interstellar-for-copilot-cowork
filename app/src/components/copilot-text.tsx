import { Fragment } from "react";
import { ExternalLink } from "lucide-react";
import { chipName, cleanDoc, parseBlocks, SourceIndex, type Block, type Cite, type CleanDoc } from "@/lib/citations";
import type { DocContent } from "@/lib/documents";
import { safeUrl } from "@/lib/safe-url";

function CiteChips({ cites }: { cites: Cite[] }) {
  return <span className="cite-group">{cites.map(cite => {
    const href = safeUrl(cite.url);
    return href
      ? <a key={cite.url} href={href} target="_blank" rel="noopener noreferrer" title={`${cite.name} — opens in a new tab`} className="cite-chip">{chipName(cite.name)}<span className="sr-only"> (source {cite.n}, opens in a new tab)</span></a>
      : <span key={cite.url} className="cite-chip">{chipName(cite.name)}</span>;
  })}</span>;
}
function Inlines({ block }: { block: Block }) {
  return <>{block.inlines.map((item, index) => item.kind === "text" ? <Fragment key={index}>{item.text}</Fragment> : <CiteChips key={index} cites={item.cites} />)}</>;
}
/** Paragraphs and bullet lists with short, clickable source names instead of long links. */
export function Blocks({ blocks }: { blocks: Block[] }) {
  const groups: { kind: Block["kind"]; items: Block[] }[] = [];
  for (const block of blocks) {
    const last = groups.at(-1);
    if (last && last.kind === "bullet" && block.kind === "bullet") last.items.push(block); else groups.push({ kind: block.kind, items: [block] });
  }
  return <>{groups.map((group, index) => group.kind === "bullet"
    ? <ul key={index} className="my-2 list-disc space-y-1 pl-5">{group.items.map((block, i) => <li key={i}><Inlines block={block} /></li>)}</ul>
    : <p key={index} className="my-2 first:mt-0 last:mb-0"><Inlines block={group.items[0]} /></p>)}</>;
}
export function SourcesList({ sources, heading = "Sources" }: { sources: Cite[]; heading?: string }) {
  if (!sources.length) return null;
  return <section className="sources-list mt-5 border-t border-border pt-3">
    <h5 className="text-xs font-bold uppercase tracking-wide text-muted-foreground">{heading}</h5>
    <ol className="mt-1.5 grid gap-x-6 gap-y-1 text-sm sm:grid-cols-2">{sources.map(cite => {
      const href = safeUrl(cite.url);
      return <li key={cite.url} className="flex min-w-0 gap-1.5"><span className="shrink-0 text-muted-foreground">{cite.n}.</span>
        {href ? <a href={href} target="_blank" rel="noopener noreferrer" className="min-w-0 truncate font-medium text-primary underline-offset-2 hover:underline" title={cite.name}>{cite.name}<ExternalLink aria-hidden="true" className="ml-1 inline size-3" /><span className="sr-only"> (opens in a new tab)</span></a>
          : <span className="min-w-0 truncate">{cite.name}</span>}</li>;
    })}</ol>
  </section>;
}
/** A Copilot-written document, cleaned up for reading. */
export function DocView({ doc }: { doc: DocContent | CleanDoc }) {
  const clean = "sources" in doc ? doc : cleanDoc(doc);
  return <article className="copilot-doc text-sm leading-relaxed">
    <h4 className="font-heading text-xl font-bold text-primary">{clean.title}</h4>
    {clean.sections.map((section, index) => <section key={index} className="mt-4">
      {section.heading && <h5 className="mb-1 font-semibold">{section.heading}</h5>}
      <Blocks blocks={section.blocks} />
    </section>)}
    <SourcesList sources={clean.sources} />
  </article>;
}
/** A Copilot chat reply, cleaned up for reading. */
export function ReplyView({ text }: { text: string }) {
  const index = new SourceIndex();
  const blocks = parseBlocks(text, index);
  return <div className="copilot-doc text-sm leading-relaxed"><Blocks blocks={blocks} /><SourcesList sources={index.list} /></div>;
}
