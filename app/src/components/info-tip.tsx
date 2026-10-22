import type { ReactNode } from "react";
import { Info } from "lucide-react";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { cn } from "@/lib/utils";

/**
 * A small "i" button that opens a short note about a feature's limits. A popover rather than a hover tooltip,
 * so it works with a keyboard, a screen reader and touch.
 */
export function InfoTip({ label, title, points, className }: { label: string; title?: string; points: ReactNode[]; className?: string }) {
  return <Popover>
    <PopoverTrigger asChild>
      <button type="button" aria-label={`About ${label}`} title={`About ${label}`}
        className={cn("inline-flex size-6 shrink-0 items-center justify-center rounded-full text-muted-foreground hover:bg-muted hover:text-foreground focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-ring", className)}>
        <Info aria-hidden="true" className="size-4" />
      </button>
    </PopoverTrigger>
    <PopoverContent align="end" className="w-80 text-sm">
      <p className="font-semibold">{title ?? label}</p>
      <ul className="mt-2 list-disc space-y-1.5 pl-5 text-muted-foreground">{points.map((point, index) => <li key={index}>{point}</li>)}</ul>
    </PopoverContent>
  </Popover>;
}
