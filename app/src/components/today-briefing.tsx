import { useState } from "react";
import { AlertTriangle, CheckCircle2, CircleDashed, Loader2, PowerOff } from "lucide-react";
import { Button } from "@/components/ui/button";
import { ActionList, ActionRow, DismissedList } from "@/components/action-list";
import { useDismissed } from "@/lib/dismissed";
import { SOURCE_LABELS, type SourceKey, type SuggestedAction } from "@/lib/actions";
import { HORIZONS, type AttentionItem } from "@/lib/attention";
import { cn } from "@/lib/utils";

export type SourceStatus = "ready" | "loading" | "off" | "error";
export interface SourceCoverage { key: SourceKey; status: SourceStatus; read: number; note?: string; label?: string; countWith?: SourceKey[]; detail?: string }

const STATUS_TEXT: Record<SourceStatus, string> = { ready: "read", loading: "loading", off: "search off", error: "couldn't load" };
const PAGE = 12;

/** Today → Actions: one ranked list across every source, grouped by when it needs attention. */
export function TodayBriefing({ items: all, coverage, helpers, mySite }: { items: AttentionItem[]; coverage: SourceCoverage[]; helpers: SuggestedAction[]; mySite?: string }) {
  const [showAll, setShowAll] = useState(false);
  const { isDismissed } = useDismissed();
  const items = all.filter(item => !isDismissed(item.id));
  const shown = showAll ? items : items.slice(0, PAGE);
  const rankOf = new Map(items.map((item, index) => [item.id, index + 1]));
  const counts = new Map<SourceKey, number>();
  for (const item of items) for (const source of item.sources) counts.set(source, (counts.get(source) ?? 0) + 1);
  const pressing = items.filter(item => item.horizon === "overdue" || item.horizon === "now" || item.horizon === "today").length;
  return <div className="space-y-6">
    <section aria-labelledby="coverage-heading" className="rounded-2xl border border-border bg-accent/40 p-4">
      <h3 id="coverage-heading" className="text-sm font-semibold">What I checked for you</h3>
      <p className="mt-1 text-sm text-muted-foreground">
        {items.length ? <><strong className="text-foreground">{items.length}</strong> things need your attention{pressing ? <>, <strong className="text-foreground">{pressing}</strong> of them today or overdue</> : null}. Ranked by urgency, due dates, who is waiting on you and importance — the same task found in several places is combined.</> : "Nothing needs your attention right now."}
      </p>
      <ul className="mt-3 flex flex-wrap gap-2" aria-label="Sources checked">
        {coverage.map(source => {
          const Icon = source.status === "ready" ? CheckCircle2 : source.status === "loading" ? Loader2 : source.status === "off" ? PowerOff : source.status === "error" ? AlertTriangle : CircleDashed;
          const found = source.countWith ? items.filter(item => item.sources.some(key => key === source.key || source.countWith!.includes(key))).length : counts.get(source.key) ?? 0;
          return <li key={source.key} className={cn("source-chip flex items-center gap-2 rounded-full border border-border bg-card py-1 pl-1 pr-3 text-xs", source.status !== "ready" && "opacity-80")}>
            <span className={cn("source-badge inline-flex items-center gap-1 rounded-full px-2 py-0.5 font-semibold", `source-${source.key}`)}>
              <Icon aria-hidden="true" className={cn("size-3.5", source.status === "loading" && "animate-spin")} />{source.label ?? SOURCE_LABELS[source.key]}
            </span>
            <span>{source.status === "ready" ? `${source.read} ${source.note ?? "read"} · ${found} to act on` : STATUS_TEXT[source.status]}{source.detail ? ` · ${source.detail}` : ""}</span>
          </li>;
        })}
      </ul>
      <p className="mt-3 text-xs text-muted-foreground">Built from your last refresh. Nothing runs until you select it.</p>
    </section>

    {HORIZONS.map(horizon => {
      const group = shown.filter(item => item.horizon === horizon.id);
      if (!group.length) return null;
      return <section key={horizon.id} aria-labelledby={`horizon-${horizon.id}`}>
        <h3 id={`horizon-${horizon.id}`} className={cn("horizon-heading mb-2 flex items-center gap-2 text-sm font-bold uppercase tracking-wide", `horizon-${horizon.id}`)}>
          <span className="horizon-dot size-2.5 rounded-full" aria-hidden="true" />{horizon.label}<span className="font-normal normal-case text-muted-foreground">({items.filter(item => item.horizon === horizon.id).length})</span>
        </h3>
        <ol className="space-y-3">{group.map(item => <ActionRow key={item.id} action={item} mySite={mySite} rank={rankOf.get(item.id)} tone={`attn-${item.horizon}`} section="today" />)}</ol>
      </section>;
    })}
    {items.length > PAGE && <Button variant="outline" className="rounded-full" aria-expanded={showAll} onClick={() => setShowAll(value => !value)}>{showAll ? "Show the top 12 only" : `Show all ${items.length}`}</Button>}

    <DismissedList section="today" label="Dismissed priorities" />

    <section aria-labelledby="plan-helpers-heading" className="border-t border-border pt-5">
      <h3 id="plan-helpers-heading" className="mb-3 text-sm font-semibold">Plan helpers</h3>
      <ActionList actions={helpers} mySite={mySite} section="today-helpers" />
    </section>
  </div>;
}
