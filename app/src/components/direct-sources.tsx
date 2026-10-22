import { type ReactNode } from "react";
import { useQuery } from "@tanstack/react-query";
import { NotebookPen } from "lucide-react";
import { Card } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { EmptyState, ErrorState, LoadingState } from "@/components/states";
import { safeUrl } from "@/lib/safe-url";
import { DirectSourceError, loadOneNotePages } from "@/lib/direct-sources";

function when(value?: string) {
  const date = value ? new Date(value) : undefined;
  return date && !Number.isNaN(date.getTime()) ? date.toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }) : "";
}
function errorText(error: unknown) {
  if (error instanceof DirectSourceError) {
    if (error.failureCause === "policy") return "Your organization's data policy blocks this app from reading it directly.";
    if (error.failureCause === "access") return "Sign-in or permission was refused. Renew the connection, then select Refresh.";
  }
  return "It could not be read right now. Try again.";
}
function Shell({ id, title, children }: { id: string; title: string; children: ReactNode }) {
  return <Card className="min-w-0 gap-0 p-5"><section aria-labelledby={`${id}-heading`}>
    <h2 id={`${id}-heading`} className="mb-3 font-heading text-xl font-bold">{title}</h2>{children}
  </section></Card>;
}
function Retry({ title, query }: { title: string; query: { error: unknown; data?: unknown; isFetching: boolean; refetch: (options: { cancelRefetch: boolean }) => unknown } }) {
  return <div role="alert"><ErrorState title={title} description={`${errorText(query.error)}${query.data ? " Showing the last loaded results." : ""}`} action={<Button variant="outline" disabled={query.isFetching} onClick={() => void query.refetch({ cancelRefetch: false })}>Retry</Button>} /></div>;
}
export function OneNoteSection() {
  const query = useQuery({ queryKey: ["direct-source", "onenote"], queryFn: loadOneNotePages, retry: false });
  return <Shell id="onenote" title="OneNote">
    {query.isError && <Retry title="OneNote could not be refreshed" query={query} />}
    {query.isPending ? <LoadingState /> : query.data && (query.data.length ? <ul className="divide-y divide-border">{query.data.map(page => {
      const href = safeUrl(page.url);
      const body = <>
        <div className="flex flex-wrap items-start justify-between gap-2"><p className="flex items-start gap-2 break-words font-semibold"><NotebookPen aria-hidden="true" className="mt-0.5 size-4 shrink-0" />{page.title}</p>{page.todos.length > 0 && <Badge className="h-auto whitespace-normal">{page.todos.length} open to-do{page.todos.length === 1 ? "" : "s"}</Badge>}</div>
        <p className="mt-1 text-xs text-muted-foreground">{[page.notebook, page.section, when(page.modified) && `Changed ${when(page.modified)}`].filter(Boolean).join(" · ")}</p>
        {page.todos.length > 0 && <ul className="mt-2 list-disc space-y-1 pl-5 text-sm">{page.todos.map((todo, index) => <li key={index} className="break-words">{todo}</li>)}</ul>}
      </>;
      return <li key={page.id}>{href ? <a href={href} target="_blank" rel="noopener noreferrer" className="block rounded-lg px-2 py-3 hover:bg-muted focus-visible:outline-2 focus-visible:outline-ring">{body}</a> : <div className="px-2 py-3">{body}</div>}</li>;
    })}</ul> : <EmptyState icon={NotebookPen} title="No recent OneNote pages found" className="py-6" />)}
  </Shell>;
}
