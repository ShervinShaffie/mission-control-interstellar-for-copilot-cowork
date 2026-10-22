import { useState } from "react";
import { RefreshCw } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { Label } from "@/components/ui/label";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { SEARCH_LABELS, type SearchGroup } from "@/lib/search-plan";

function skipKey(owner?: string) {
  return owner && /^[0-9a-f-]{36}$/i.test(owner) ? `mission-control:skip-refresh-confirm:v1:${owner.toLowerCase()}` : undefined;
}
function readSkip(owner?: string) {
  try { const key = skipKey(owner); return key ? localStorage.getItem(key) === "1" : false; } catch { return false; }
}

/**
 * The top Refresh button. When the refresh will run Copilot searches, it first
 * explains that they use Copilot Credits billed to the organization.
 */
export function RefreshConfirm({ searches, owner, onConfirm, disabled, label }: { searches: SearchGroup[]; owner?: string; onConfirm: () => void; disabled: boolean; label: string }) {
  const [open, setOpen] = useState(false);
  const [dontAsk, setDontAsk] = useState(false);
  function start() {
    if (searches.length === 0 || readSkip(owner)) { onConfirm(); return; }
    setDontAsk(false);
    setOpen(true);
  }
  function confirm() {
    if (dontAsk) { try { const key = skipKey(owner); if (key) localStorage.setItem(key, "1"); } catch { /* ask again next time */ } }
    setOpen(false);
    onConfirm();
  }
  const count = searches.length;
  return <>
    <Button onClick={start} disabled={disabled} aria-busy={disabled} className="h-auto whitespace-normal rounded-full py-2"><RefreshCw aria-hidden="true" />{label}</Button>
    <Dialog open={open} onOpenChange={setOpen}>
      <DialogContent className="text-left sm:max-w-lg">
        <DialogHeader>
          <DialogTitle>Refresh uses Copilot Credits</DialogTitle>
          <DialogDescription>This refresh runs {count} Copilot search{count === 1 ? "" : "es"}, and each one consumes Copilot Credits that are billed to your organization.</DialogDescription>
        </DialogHeader>
        <div className="space-y-3 text-sm">
          <ul className="list-disc space-y-1 pl-5">{searches.map(group => <li key={group}>{SEARCH_LABELS[group]}</li>)}</ul>
          <p>The cost of each search varies with how much work it finds. Your Outlook, Teams, OneDrive, SharePoint and OneNote boxes also reload; those reads count toward your organization's app usage.</p>
          <p>To run fewer searches, turn off the ones you don't need in Customize. Suggested actions in the Actions tabs never run on their own.</p>
          <div className="flex items-center gap-2 pt-1">
            <Checkbox id="refresh-dont-ask" checked={dontAsk} onCheckedChange={value => setDontAsk(value === true)} />
            <Label htmlFor="refresh-dont-ask" className="font-normal">Don't ask again in this browser</Label>
          </div>
        </div>
        <DialogFooter className="gap-2">
          <Button variant="outline" onClick={() => setOpen(false)}>Cancel</Button>
          <Button onClick={confirm}>Refresh and use credits</Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  </>;
}
