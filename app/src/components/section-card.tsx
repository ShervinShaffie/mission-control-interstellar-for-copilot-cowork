import { useState, type ReactNode } from "react";
import { ChevronDown, type LucideIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { useSectionCollapse } from "@/lib/section-collapse";
import { Card } from "@/components/ui/card";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { ActionList } from "@/components/action-list";
import { useDismissed } from "@/lib/dismissed";
import type { SuggestedAction } from "@/lib/actions";
import { cn } from "@/lib/utils";

/** One dashboard component: a data tab and an Actions tab with ranked suggestions. */
export function SectionCard({ id, title, icon: Icon, tone, dataLabel, children, actions, actionsIntro, mySite, extraActions, actionsContent, actionsLabel = "Actions" }: {
  id: string; title: string; icon: LucideIcon; tone: string; dataLabel: string; children: ReactNode;
  actions: SuggestedAction[]; actionsIntro?: ReactNode; mySite?: string; extraActions?: ReactNode;
  /** Replaces the default ranked list (used by Today). */ actionsContent?: ReactNode; actionsLabel?: string;
}) {
  const [tab, setTab] = useState("data");
  const { isDismissed } = useDismissed();
  const shown = actions.filter(action => !isDismissed(action.id)).length;
  const collapse = useSectionCollapse();
  const collapsed = collapse?.isCollapsed(id) ?? false;
  return <Card className="section-card min-w-0 gap-0 overflow-hidden p-0" data-collapsed={collapsed || undefined}>
    <section aria-labelledby={`${id}-title`}>
      <Tabs value={tab} onValueChange={setTab} className="gap-0">
        <div className={cn("flex flex-wrap items-center justify-between gap-3 border-b border-border px-5 py-4", tone)}>
          <h2 id={`${id}-title`} className="flex items-center gap-3 font-heading text-xl font-bold">
            <span className="section-icon flex size-10 items-center justify-center rounded-2xl" aria-hidden="true"><Icon className="size-5" /></span>{title}
          </h2>
          <div className="ml-auto flex items-center gap-2">
            {!collapsed && <TabsList aria-label={`${title} views`} className="rounded-full">
              <TabsTrigger value="data" className="rounded-full px-4">{dataLabel}</TabsTrigger>
              <TabsTrigger value="actions" className="rounded-full px-4">{actionsLabel}{shown ? <span className="ml-1.5 rounded-full bg-primary px-1.5 text-xs font-semibold text-primary-foreground" aria-label={`${shown} suggested`}>{shown}</span> : null}</TabsTrigger>
            </TabsList>}
            {collapse && <Button type="button" variant="ghost" size="icon" className="size-9 shrink-0 rounded-full" data-pdf-exclude
              aria-expanded={!collapsed} aria-controls={`${id}-body`} aria-label={collapsed ? `Expand ${title}` : `Collapse ${title}`}
              title={collapsed ? `Expand ${title}` : `Collapse ${title}`} onClick={() => collapse.toggle(id)}>
              <ChevronDown aria-hidden="true" className={cn("transition-transform motion-reduce:transition-none", !collapsed && "rotate-180")} />
            </Button>}
          </div>
        </div>
        <div id={`${id}-body`} hidden={collapsed} className={collapsed ? undefined : "flex flex-1 flex-col"}>
        <TabsContent value="data" className="tab-panel p-5" forceMount hidden={tab !== "data"}>{children}</TabsContent>
        <TabsContent value="actions" className="tab-panel space-y-4 p-5" forceMount hidden={tab !== "actions"}>
          {actionsIntro && <div className="text-sm text-muted-foreground">{actionsIntro}</div>}
          {actionsContent ?? <ActionList actions={actions} mySite={mySite} section={id} />}
          {extraActions}
        </TabsContent>
        </div>
      </Tabs>
    </section>
  </Card>;
}
