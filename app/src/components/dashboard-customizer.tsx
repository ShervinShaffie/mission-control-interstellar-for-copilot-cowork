import { useRef, useState } from "react";
import { ArrowDown, ArrowUp, Plus, RotateCcw, Settings2, X } from "lucide-react";
import { RETENTION_CHOICES, restoreAll, setRetention, useDismissed } from "@/lib/dismissed";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogDescription, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { componentCatalog, defaultPreferences, moveComponent, type ComponentId, type DashboardPreferences } from "@/lib/dashboard-preferences";

export function DashboardCustomizer({ preferences, update, storage, identityVerified, triggerId }: { preferences: DashboardPreferences; update: (next: DashboardPreferences) => void; storage: "browser" | "session"; identityVerified: boolean; triggerId?: string }) {
  const [announcement, setAnnouncement] = useState("");
  const dismissed = useDismissed();
  const actions = useRef<Partial<Record<ComponentId, HTMLButtonElement | null>>>({});
  const visible = preferences.visible.map(id => componentCatalog.find(item => item.id === id)).filter(item => item !== undefined);
  const optional = componentCatalog.filter(item => !preferences.visible.includes(item.id));
  function toggle(id: ComponentId, title: string, add: boolean) {
    update({ ...preferences, visible: add ? [...preferences.visible, id] : preferences.visible.filter(value => value !== id) });
    setAnnouncement(`${title} ${add ? "added at the end of your layout" : "removed; available under Removed components in Customize"}.`);
    requestAnimationFrame(() => actions.current[id]?.focus());
  }
  function move(id: ComponentId, title: string, direction: -1 | 1) {
    const next = moveComponent(preferences.visible, id, direction);
    update({ ...preferences, visible: next });
    setAnnouncement(`${title} moved to position ${next.indexOf(id) + 1} of ${next.length}.`);
  }
  return <Dialog>
    <DialogTrigger asChild><Button id={triggerId} variant="outline" className="h-auto whitespace-normal rounded-full py-2"><Settings2 aria-hidden="true" />Customize</Button></DialogTrigger>
    <DialogContent className="max-h-[85dvh] overflow-y-auto sm:max-w-2xl">
      <DialogHeader><DialogTitle>Customize</DialogTitle><DialogDescription>{storage === "browser" ? "Layout preferences use your verified identity in this browser only, not across devices. Business data is not stored with them." : identityVerified ? "Browser storage is unavailable. Preferences last for this open page only." : "Until your identity is verified, preferences last for this open page only."}</DialogDescription></DialogHeader>
      <p role="status" aria-live="polite" className="sr-only">{announcement}</p>
      <section aria-labelledby="visible-components-heading">
        <h3 id="visible-components-heading" className="font-semibold">On your dashboard</h3>
        {visible.length === 0 ? <div className="mt-3 flex items-start gap-2 rounded-lg border border-border p-3"><Settings2 className="size-5 shrink-0" aria-hidden="true" /><p>No components selected. Use an Add button below or restore the default layout.</p></div> : <ol className="mt-3 divide-y divide-border">{visible.map((item, index) => <li key={item.id} className="py-3">
          <div className="min-w-0"><p className="break-words font-medium">{index + 1}. {item.title}</p><p className="mt-1 break-words text-sm text-muted-foreground">{item.detail}</p></div>
          <div className="mt-2 flex flex-wrap gap-2">
            <Button variant="outline" size="sm" className="h-auto whitespace-normal" disabled={index === 0} aria-label={`Move up ${item.title}`} onClick={() => move(item.id, item.title, -1)}><ArrowUp aria-hidden="true" />Move up</Button>
            <Button variant="outline" size="sm" className="h-auto whitespace-normal" disabled={index === visible.length - 1} aria-label={`Move down ${item.title}`} onClick={() => move(item.id, item.title, 1)}><ArrowDown aria-hidden="true" />Move down</Button>
            <Button ref={node => { actions.current[item.id] = node; }} variant="ghost" size="sm" className="h-auto whitespace-normal" aria-label={`Remove ${item.title}`} onClick={() => toggle(item.id, item.title, false)}><X aria-hidden="true" />Remove</Button>
          </div>
        </li>)}</ol>}
      </section>
      <section aria-labelledby="dismissed-heading" className="border-t border-border pt-4">
        <h3 id="dismissed-heading" className="font-semibold">Dismissed suggestions</h3>
        <p className="mt-1 text-sm text-muted-foreground">Dismissed suggestions wait in a "Dismissed" list at the bottom of each Actions tab, where you can restore them. After the time you choose they're removed for good and won't come back.</p>
        <div className="mt-3 flex flex-wrap items-center gap-2 text-sm">
          <span>Keep them for</span>
          {RETENTION_CHOICES.map(days => <Button key={days} size="sm" variant={dismissed.retention === days ? "default" : "outline"} aria-pressed={dismissed.retention === days} className="rounded-full" onClick={() => { setRetention(days); setAnnouncement(`Dismissed suggestions are kept for ${days} day${days > 1 ? "s" : ""}.`); }}>{days} day{days > 1 ? "s" : ""}</Button>)}
        </div>
        <p className="mt-2 text-xs text-muted-foreground">{dismissed.items.length} waiting to be restored · {dismissed.hiddenCount} removed for good.</p>
        {(dismissed.items.length > 0 || dismissed.hiddenCount > 0) && <Button variant="outline" size="sm" className="mt-2 rounded-full" onClick={() => { restoreAll(); setAnnouncement("All dismissed suggestions were restored."); }}><RotateCcw aria-hidden="true" />Bring back everything I dismissed</Button>}
      </section>
      <section aria-labelledby="optional-components-heading" className="border-t border-border pt-4">
        <h3 id="optional-components-heading" className="font-semibold">Removed components</h3>
        {optional.length ? <ul className="mt-2 divide-y divide-border">{optional.map(item => <li key={item.id} className="flex flex-wrap items-start justify-between gap-3 py-3"><div className="min-w-0 flex-1"><p className="break-words font-medium">{item.title}</p><p className="mt-1 break-words text-sm text-muted-foreground">{item.detail}</p></div><Button ref={node => { actions.current[item.id] = node; }} variant="outline" size="sm" className="h-auto whitespace-normal" aria-label={`Add ${item.title}`} onClick={() => toggle(item.id, item.title, true)}><Plus aria-hidden="true" />Add</Button></li>)}</ul> : <p className="mt-3 text-sm">All components are on your dashboard. Removed components appear here so you can add them back.</p>}
      </section>
      <Button variant="outline" className="h-auto whitespace-normal" onClick={() => { update({ ...defaultPreferences(), theme: preferences.theme }); setAnnouncement("Default component order and views restored."); }}><RotateCcw aria-hidden="true" />Restore default layout</Button>
    </DialogContent>
  </Dialog>;
}
