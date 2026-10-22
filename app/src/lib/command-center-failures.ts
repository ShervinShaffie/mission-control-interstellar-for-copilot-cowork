import { isRecord } from "@/lib/command-center-evidence";

export type ReadFailureCause = "not-found" | "access" | "policy" | "configuration-timeout" | "response" | "unknown";
// Classification only selects fixed UI text. Never expose upstream messages or IDs.
export function readFailureCause(error: unknown): ReadFailureCause {
  const record = isRecord(error) ? error : undefined;
  const message = error instanceof Error ? error.message : typeof record?.message === "string" ? record.message : typeof error === "string" ? error : "";
  const status = record?.statusCode ?? record?.status;
  const code = typeof status === "number" ? status : typeof status === "string" && /^\d{3}$/.test(status) ? Number(status) : Number(message.match(/\bstatus(?:\s+code)?\s*[:=]?\s*(\d{3})\b/i)?.[1]);
  if (/\bDLP\b|data loss prevention|advanced connector policy/i.test(message) && /block|denied|not allowed|restrict/i.test(message)) return "policy";
  if (code === 442) return "policy";
  if (code === 401 || code === 403) return "access";
  if (code === 404) return "not-found";
  if (/connection configuration/i.test(message) && /timed?\s*out|timeout/i.test(message)) return "configuration-timeout";
  // The canonical Work IQ parser gives these bounded, app-owned diagnostics.
  if (/^(Invalid JSON in |Work IQ Copilot reply item \d+ failed schema validation\.|No Work IQ Copilot reply item passed schema validation\.|Work IQ returned no JSON-RPC response |Work IQ JSON-RPC result was missing\.|Work IQ response nesting is too deep\.)/.test(message)) return "response";
  return "unknown";
}

export class TeamsReadError extends Error {
  constructor(readonly stage: "chat-list" | "messages" | "validation", readonly failureCause: ReadFailureCause) {
    super("Teams read failed");
    this.name = "TeamsReadError";
  }
}
export function teamsFailureDescription(error: unknown): string {
  if (!(error instanceof TeamsReadError)) return "The Teams request did not produce a validated result. This is not an empty result.";
  const scope = error.stage === "messages" ? "Message request" : "Chat-list request";
  if (error.stage === "validation") return "The returned Teams records could not be validated. This is not an empty result.";
  const reasons: Record<ReadFailureCause, string> = {
    "not-found": "returned resource not found (404)",
    access: "was denied access",
    policy: "was blocked by an organization data policy",
    "configuration-timeout": "timed out while loading its connection configuration",
    response: "returned an unsupported response",
    unknown: "failed",
  };
  return `${scope} ${reasons[error.failureCause]}. No new result was validated.`;
}
export function replyFailureDescription(cause: ReadFailureCause): string {
  if (cause === "access") return "Email search was denied access. No new result was validated.";
  if (cause === "policy") return "Email search was blocked by an organization data policy. No new result was validated.";
  if (cause === "configuration-timeout") return "Email search timed out while loading its connection configuration. No new result was validated.";
  if (cause === "response") return "The email-search response could not be parsed or did not match the requested fields. This is not an empty result.";
  return "An access, transport or response problem prevented a validated result. This is not an empty result.";
}
