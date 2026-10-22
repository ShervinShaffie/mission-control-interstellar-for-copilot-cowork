import { QueryClient } from "@tanstack/react-query";
import { manualQueryOptions } from "@/lib/manual-refresh";

// Applies to every data query, including optional components restored later.
export const queryClient = new QueryClient({
  defaultOptions: { queries: manualQueryOptions },
});
