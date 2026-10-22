/** Which teams to treat as projects (per user, this browser). Empty = automatic (project-sounding teams first). */
const key = (owner?: string) => owner && /^[0-9a-f-]{36}$/i.test(owner) ? `mission-control-interstellar:project-teams:v1:${owner.toLowerCase()}` : undefined;
export function readTeamChoice(owner?: string): string[] {
  try { const k = key(owner); const value: unknown = k ? JSON.parse(localStorage.getItem(k) ?? "[]") : []; return Array.isArray(value) ? value.filter((id): id is string => typeof id === "string").slice(0, 12) : []; } catch { return []; }
}
export function writeTeamChoice(owner: string | undefined, ids: string[]) {
  try { const k = key(owner); if (k) localStorage.setItem(k, JSON.stringify(ids.slice(0, 12))); } catch { /* memory only */ }
}
