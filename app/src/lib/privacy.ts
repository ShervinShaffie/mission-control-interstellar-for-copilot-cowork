// Email addresses are never displayed. Names are shown instead when the
// directory can resolve them; otherwise a neutral label is used.
const EMAIL = /[A-Z0-9._%+'-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
const SINGLE_EMAIL = /^\s*(?:mailto:)?[A-Z0-9._%+'-]+@[A-Z0-9.-]+\.[A-Z]{2,}\s*$/i;

/** True when the value is just one email address. */
export function isEmailAddress(value?: string | null): boolean {
  return !!value && SINGLE_EMAIL.test(value);
}

/** Returns the bare address when the value is one email address. */
export function emailAddress(value?: string | null): string | undefined {
  return isEmailAddress(value) ? value!.trim().replace(/^mailto:/i, "").toLowerCase() : undefined;
}

/**
 * Removes every email address from display text. "Name <address>" keeps the
 * name; any other address is replaced with a neutral marker.
 */
export function hideEmails(value: string): string;
export function hideEmails(value: string | undefined): string | undefined;
export function hideEmails(value?: string): string | undefined {
  if (!value) return value;
  return value
    .replace(/\s*<\s*(?:mailto:)?[A-Z0-9._%+'-]+@[A-Z0-9.-]+\.[A-Z]{2,}\s*>/gi, "")
    .replace(/\(\s*(?:mailto:)?[A-Z0-9._%+'-]+@[A-Z0-9.-]+\.[A-Z]{2,}\s*\)/gi, "")
    .replace(/mailto:/gi, "")
    .replace(EMAIL, "(email hidden)")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

/** Applies hideEmails to every string field of a record. */
export function hideEmailsInRecord<T extends Record<string, unknown>>(record: T, keep: (keyof T)[] = []): T {
  const result: Record<string, unknown> = { ...record };
  for (const [key, value] of Object.entries(result)) {
    if (typeof value === "string" && !keep.includes(key as keyof T)) result[key] = hideEmails(value);
  }
  return result as T;
}

/**
 * First name only. "Doe, Jane" -> "Jane"; "Jane Doe (Contoso)" -> "Jane";
 * prefers the directory's given name when supplied.
 */
const KNOWN_NAMES = new Map<string, string>();
/** Remembers a directory first name for an address (session only), so later displays never show the address. */
export function rememberFirstName(address: string, name: string) {
  const key = address.trim().toLowerCase();
  if (key && name && !isEmailAddress(name)) KNOWN_NAMES.set(key, name);
}
/** The part of an address before the @ ("jane.doe@x.com" -> "jane.doe"), used when no name is known. */
export function aliasOf(address: string): string {
  return address.trim().replace(/^mailto:/i, "").split("@")[0] ?? "";
}
export function firstName(displayName?: string | null, givenName?: string | null): string {
  const address = emailAddress(displayName);
  if (address && !givenName?.trim()) return KNOWN_NAMES.get(address) ?? aliasOf(address);
  const given = hideEmails(givenName ?? "")?.trim();
  if (given) return given.split(/\s+/)[0];
  let name = hideEmails(displayName ?? "")?.replace(/\([^)]*\)|\[[^\]]*\]/g, " ").trim() ?? "";
  if (!name || name === "(email hidden)") return "";
  if (name.includes(",")) name = name.split(",")[1]?.trim() || name.split(",")[0].trim();
  return name.split(/\s+/)[0] ?? "";
}
