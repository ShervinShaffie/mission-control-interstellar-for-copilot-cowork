import { useState } from "react";
import { defaultPreferences, parsePreferences, preferenceKey, type DashboardPreferences } from "@/lib/dashboard-preferences";

type PreferenceState = { owner?: string; preferences: DashboardPreferences; storage: "browser" | "session" };
function readPreferences(owner?: string): PreferenceState {
  if (!owner) return { owner, preferences: defaultPreferences(), storage: "session" };
  try {
    let raw = localStorage.getItem(owner);
    if (!raw) {
      // Carry over a layout saved by an earlier build under an older key (same app and person, longer key).
      const [prefix, , version, person] = owner.split(":");
      for (let i = 0; i < localStorage.length && !raw; i++) {
        const key = localStorage.key(i);
        if (key && key.startsWith(`${prefix}:`) && key.endsWith(`:${version}:${person}`) && key.split(":").length === 5) raw = localStorage.getItem(key);
      }
    }
    return { owner, preferences: raw ? parsePreferences(JSON.parse(raw)) : defaultPreferences(), storage: "browser" };
  } catch { return { owner, preferences: defaultPreferences(), storage: "session" }; }
}
export function useDashboardPreferences(playerId?: string) {
  const owner = preferenceKey(playerId);
  const [state, setState] = useState(() => readPreferences(owner));
  // Reset before rendering another player; do not write the previous player's layout.
  if (state.owner !== owner) setState(readPreferences(owner));
  const current = state.owner === owner ? state : readPreferences(owner);
  function update(preferences: DashboardPreferences) {
    const clean = parsePreferences(preferences);
    let storage: PreferenceState["storage"] = "session";
    if (owner) {
      try { localStorage.setItem(owner, JSON.stringify(clean)); storage = "browser"; } catch { /* Session preferences still work. */ }
    }
    setState({ owner, preferences: clean, storage });
  }
  return { preferences: current.preferences, storage: current.storage, identityVerified: Boolean(owner), update };
}
