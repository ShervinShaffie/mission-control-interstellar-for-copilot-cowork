import { Fragment } from "react";
import { parseRichText } from "@/lib/source-links";
import { safeUrl } from "@/lib/safe-url";

/**
 * Renders citation text with short source names instead of long URLs. Inside
 * an element that is already a link, pass interactive={false} so names render
 * as plain text (links cannot be nested).
 */
export function RichText({ text, interactive = true }: { text: string; interactive?: boolean }) {
  return <>{parseRichText(text).map((segment, index) => {
    if (segment.kind === "text") return <Fragment key={index}>{segment.text}</Fragment>;
    const href = safeUrl(segment.url);
    if (!interactive || !href) return <span key={index} className="font-medium">{segment.text}</span>;
    return <a key={index} href={href} target="_blank" rel="noopener noreferrer" className="font-medium underline underline-offset-2 focus-visible:outline-2 focus-visible:outline-ring">{segment.text}</a>;
  })}</>;
}
