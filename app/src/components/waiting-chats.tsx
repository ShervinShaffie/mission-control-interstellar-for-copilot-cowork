import type { UseQueryResult } from "@tanstack/react-query";
import { ExternalLink, MessageSquare } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { EmptyState, ErrorState, LoadingState } from "@/components/states";
import { safeUrl } from "@/lib/safe-url";
import { agoLabel, type WaitingChat } from "@/lib/teams-waiting";

/** People whose latest Teams chat message is to you (free Teams reads, no Copilot). */
export function WaitingChats({ query }: { query: UseQueryResult<WaitingChat[]> }) {
  const rows = query.data ?? [];
  return <Card className="mt-5 min-w-0 gap-0 p-0">
    <section aria-labelledby="waiting-chats-heading" aria-busy={query.isFetching} className="p-5">
      <h3 id="waiting-chats-heading" className="font-heading text-xl font-bold">Waiting on you in Teams</h3>
      <p className="mt-1 text-xs text-muted-foreground">Recent chats where someone else sent the last message (last 7 days).</p>
      <div className="mt-3">
        {query.isPending ? <LoadingState rows={2} />
          : query.isError ? <div role="alert"><ErrorState title="Could not check your Teams chats" action={<Button variant="outline" disabled={query.isFetching} onClick={() => void query.refetch()}>Retry</Button>} /></div>
          : !rows.length ? <EmptyState icon={MessageSquare} title="No one is waiting on you in Teams" />
          : <ul className="divide-y divide-border rounded-xl border border-border">{rows.map(chat => {
            const url = safeUrl(chat.messageUrl) ?? safeUrl(chat.url);
            const body = <div className="flex items-start gap-3 p-3">
              <span className="follow-avatar flex size-9 shrink-0 items-center justify-center rounded-full text-sm font-semibold" aria-hidden="true">{chat.person.slice(0, 1).toUpperCase()}</span>
              <div className="min-w-0 flex-1">
                <p className="break-words text-sm font-semibold">{chat.person}{chat.group && <span className="font-normal text-muted-foreground"> · {chat.chatTitle}</span>}</p>
                <p className="mt-1 break-words text-sm">“{chat.preview}”</p>
                <p className="mt-1 text-xs text-muted-foreground">{agoLabel(chat.created)}</p>
              </div>
              {url && <ExternalLink aria-hidden="true" className="mt-1 size-4 shrink-0" />}
            </div>;
            return <li key={chat.chatId}>{url ? <a href={url} target="_blank" rel="noreferrer" className="block hover:bg-accent focus-visible:outline-2 focus-visible:outline-ring">{body}<span className="sr-only">(opens this message in Teams in a new tab)</span></a> : body}</li>;
          })}</ul>}
      </div>
    </section>
  </Card>;
}
