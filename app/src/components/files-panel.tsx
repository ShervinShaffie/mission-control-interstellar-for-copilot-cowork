import { useState } from "react";
import { ExternalLink, File, FileSpreadsheet, FileText, FileType, NotebookPen, Presentation, Shapes, type LucideIcon } from "lucide-react";
import { Button } from "@/components/ui/button";
import { EmptyState, LoadingState } from "@/components/states";
import { DriveSection } from "@/components/work-context";
import { kindLabel, type FileItem, type FileKind } from "@/lib/files";
import { safeUrl } from "@/lib/safe-url";
import { cn } from "@/lib/utils";

const ICON: Record<FileKind, { icon: LucideIcon; color: string }> = {
  word: { icon: FileText, color: "text-blue-700 dark:text-blue-300" },
  excel: { icon: FileSpreadsheet, color: "text-emerald-700 dark:text-emerald-300" },
  powerpoint: { icon: Presentation, color: "text-orange-700 dark:text-orange-300" },
  pdf: { icon: FileType, color: "text-red-700 dark:text-red-300" },
  onenote: { icon: NotebookPen, color: "text-purple-700 dark:text-purple-300" },
  loop: { icon: Shapes, color: "text-pink-700 dark:text-pink-300" },
  other: { icon: File, color: "text-muted-foreground" },
};
const day = (date?: Date) => date ? date.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" }) : "";

export function FileRow({ file, compact }: { file: FileItem; compact?: boolean }) {
  const { icon: Icon, color } = ICON[file.kind];
  const href = safeUrl(file.url);
  return <div className="flex items-start gap-3 py-3">
    <Icon aria-hidden="true" className={cn("mt-0.5 size-5 shrink-0", color)} />
    <div className="min-w-0 flex-1">
      <p className="break-words font-medium">{href ? <a href={href} target="_blank" rel="noopener noreferrer" className="underline-offset-4 hover:underline focus-visible:outline-2 focus-visible:outline-ring">{file.name}<span className="sr-only"> ({kindLabel(file.kind)}, opens in a new tab)</span></a> : file.name}</p>
      {!compact && <div className="mt-1 flex flex-wrap gap-1.5">{file.reasons.slice(0, 2).map(reason => <span key={reason} className="rounded-full border border-border px-2 py-0.5 text-xs">{reason}</span>)}</div>}
      <p className="mt-1 text-xs text-muted-foreground">{[kindLabel(file.kind), file.who && `from ${file.who}`, day(file.when)].filter(Boolean).join(" · ")}
        {file.contexts[0] && <> · {safeUrl(file.contexts[0].url) ? <a href={safeUrl(file.contexts[0].url)} target="_blank" rel="noopener noreferrer" className="underline underline-offset-2">{file.contexts[0].label}</a> : file.contexts[0].label}</>}</p>
    </div>
    {href && <a href={href} target="_blank" rel="noopener noreferrer" className="inline-flex shrink-0 items-center gap-1 rounded-full border border-primary/40 px-3 py-1 text-sm font-medium text-primary hover:bg-accent focus-visible:outline-2 focus-visible:outline-ring">Open<ExternalLink aria-hidden="true" className="size-3.5" /><span className="sr-only">{file.name} (opens in a new tab)</span></a>}
  </div>;
}

/**
 * Files: what needs you (ranked), everything shared with you recently, and your OneDrive.
 * Built from data the dashboard already reads; only "trending" is an extra (optional) read.
 */
export function FilesPanel({ files, loading, oneDriveUrl }: { files: FileItem[]; loading: boolean; oneDriveUrl: string }) {
  const [view, setView] = useState<"now" | "shared" | "drive">("now");
  const [all, setAll] = useState(false);
  const ranked = files;
  const shared = files.filter(file => !(file.origins.length === 1 && file.origins[0] === "trending")).sort((a, b) => (b.when?.getTime() ?? 0) - (a.when?.getTime() ?? 0));
  const list = view === "now" ? ranked : shared;
  const shown = all ? list : list.slice(0, 8);
  return <div className="space-y-3">
    <div role="group" aria-label="Files view" className="flex flex-wrap gap-2">
      {([["now", "For you now"], ["shared", "Shared with you"], ["drive", "My OneDrive"]] as const).map(([value, label]) =>
        <Button key={value} size="sm" variant={view === value ? "default" : "outline"} aria-pressed={view === value} className="rounded-full" onClick={() => { setView(value); setAll(false); }}>{label}</Button>)}
      <a href={oneDriveUrl} target="_blank" rel="noopener noreferrer" className="ml-auto inline-flex items-center gap-1 self-center text-sm font-medium text-primary underline underline-offset-2">Open OneDrive<ExternalLink aria-hidden="true" className="size-3.5" /><span className="sr-only">(opens in a new tab)</span></a>
    </div>
    {view === "drive" ? <DriveSection />
      : <>
        <p className="text-sm text-muted-foreground">{view === "now" ? "Files linked in emails you owe a reply, chats waiting on you, meetings in the next two days and your project teams — most important first. Plus documents trending around you." : "Every file linked or attached in your recent email (14 days), Teams posts and chats waiting on you — newest first."}</p>
        {loading && !list.length ? <LoadingState rows={3} />
          : !list.length ? <EmptyState icon={File} title={view === "now" ? "No files need you right now" : "No files shared with you recently"} description="Files appear here when an email, Teams post, chat or meeting invite links to one." />
          : <><ul className="divide-y divide-border">{shown.map(file => <li key={file.key}><FileRow file={file} /></li>)}</ul>
            {list.length > shown.length && <Button variant="ghost" size="sm" className="rounded-full" onClick={() => setAll(true)}>Show all {list.length}</Button>}</>}
      </>}
  </div>;
}
