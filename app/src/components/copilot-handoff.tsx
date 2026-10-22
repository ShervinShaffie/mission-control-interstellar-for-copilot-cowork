import { useId, useState } from "react";
import { Check, ExternalLink, Loader2, Sparkles } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { COPILOT_URL } from "@/lib/copilot-prompts";
import { recordHistory } from "@/lib/history";

/** Copies text. Tries the Clipboard API, then the classic copy command (works in more embedded frames). */
async function copyText(text: string): Promise<boolean> {
  try { await navigator.clipboard.writeText(text); return true; } catch { /* try the fallback */ }
  try {
    const area = document.createElement("textarea");
    area.value = text; area.setAttribute("readonly", ""); area.style.position = "fixed"; area.style.opacity = "0";
    document.body.appendChild(area); area.select();
    const ok = document.execCommand("copy");
    area.remove();
    return ok;
  } catch { return false; }
}

/**
 * "Draft with Copilot": builds a ready-made request from what the dashboard already read, copies it and
 * opens Microsoft 365 Copilot in a new tab. The person pastes it (Ctrl+V) and sends. The app never calls Copilot.
 */
export function CopilotHandoff({ build, title, section, disabled }: { build: () => string | Promise<string>; title: string; section?: string; disabled?: boolean }) {
  const id = useId();
  const [state, setState] = useState<"idle" | "working" | "copied" | "manual">("idle");
  const [request, setRequest] = useState("");
  const [error, setError] = useState("");
  async function go() {
    setState("working"); setError("");
    try {
      const text = await build();
      setRequest(text);
      const copied = await copyText(text);
      if (copied) window.open(COPILOT_URL, "_blank", "noopener,noreferrer");
      setState(copied ? "copied" : "manual");
      recordHistory({ kind: "action", title: `Copilot request: ${title}`, section, cost: "free", status: "done", message: "Copied a ready-made request for Microsoft 365 Copilot.", links: [{ label: "Microsoft 365 Copilot", url: COPILOT_URL }], text });
    } catch {
      setError("The request couldn't be prepared. Nothing was changed."); setState("idle");
    }
  }
  return <div className="copilot-handoff space-y-2">
    <div className="flex flex-wrap items-center gap-2">
      <Button type="button" variant="outline" className="rounded-full" disabled={disabled || state === "working"} onClick={() => void go()}>
        {state === "working" ? <Loader2 aria-hidden="true" className="animate-spin" /> : state === "copied" ? <Check aria-hidden="true" /> : <Sparkles aria-hidden="true" />}
        {state === "copied" ? "Copied — copy again" : "Draft with Copilot"}
      </Button>
      {state === "idle" && <span className="text-xs text-muted-foreground">Copies a ready-made request and opens Microsoft 365 Copilot. Paste it (Ctrl+V) and send.</span>}
    </div>
    {(state === "copied" || state === "manual") && <div role="status" className="rounded-xl border border-dashed border-primary/50 bg-accent/40 p-3 text-sm">
      <p>{state === "copied" ? <>Request copied. In Copilot, paste it (<kbd>Ctrl</kbd>+<kbd>V</kbd>, or <kbd>⌘</kbd>+<kbd>V</kbd> on a Mac) and press Enter. Then copy Copilot's answer back into your draft.</>
        : <>Your browser didn't allow copying. Select all the text below (<kbd>Ctrl</kbd>+<kbd>A</kbd>), copy it, then paste it into Copilot.</>}</p>
      <a href={COPILOT_URL} target="_blank" rel="noopener noreferrer" className="mt-2 inline-flex items-center gap-1.5 font-medium text-primary underline underline-offset-2">Open Microsoft 365 Copilot<ExternalLink aria-hidden="true" className="size-3.5" /><span className="sr-only">(opens in a new tab)</span></a>
      <details className="mt-2" open={state === "manual"}>
        <summary className="cursor-pointer text-xs font-semibold text-muted-foreground">Show the request</summary>
        <label htmlFor={id} className="sr-only">Request for Copilot</label>
        <Textarea id={id} readOnly value={request} className="mt-2 min-h-40 font-mono text-xs" onFocus={event => event.currentTarget.select()} />
      </details>
      <p className="mt-1 text-xs text-muted-foreground">Saved in History, so you can copy it again later.</p>
    </div>}
    {error && <p role="alert" className="text-sm text-destructive">{error}</p>}
  </div>;
}
