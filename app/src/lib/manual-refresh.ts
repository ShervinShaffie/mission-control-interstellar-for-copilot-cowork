import type { QueryClient } from "@tanstack/react-query";

// First loads and explicit selections can read; cached data never refreshes itself.
export const manualQueryOptions = {
  staleTime: Infinity,
  retry: false,
  retryOnMount: false,
  refetchInterval: false,
  refetchIntervalInBackground: false,
  refetchOnMount: false,
  refetchOnWindowFocus: false,
  refetchOnReconnect: false,
} as const;

export interface RefreshSummary { failed: number }
const flights = new WeakMap<QueryClient, Promise<RefreshSummary>>();
export function refreshActiveData(client: QueryClient): Promise<RefreshSummary> {
  const existing = flights.get(client);
  if (existing) return existing;
  const flight = (async () => {
    // Join any read already started by the explicit action; never cancel/restart it.
    await client.refetchQueries({ type: "active" }, { cancelRefetch: false });
    const failed = client.getQueryCache().findAll({ type: "active" }).filter(query => query.state.status === "error").length;
    return { failed };
  })().finally(() => flights.delete(client));
  flights.set(client, flight);
  return flight;
}
