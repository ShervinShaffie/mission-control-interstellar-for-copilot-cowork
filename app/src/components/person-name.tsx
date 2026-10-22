import { useQuery } from "@tanstack/react-query";
import { Office365UsersService } from "../../generated/services/Office365UsersService";
import { aliasOf, emailAddress, firstName } from "@/lib/privacy";

/**
 * Shows a person by first name only, never by email address. An address is resolved to a
 * display name through the directory (cached for the session); people outside
 * the directory get the fallback label.
 */
export function PersonName({ value, fallback = "Name unavailable" }: { value?: string | null; fallback?: string }) {
  const address = emailAddress(value);
  const lookup = useQuery({
    queryKey: ["directory-first-name", address],
    enabled: !!address,
    retry: false,
    queryFn: async () => {
      if (!address) return "";
      const result = await Office365UsersService.UserProfile_V2(address, "givenName,displayName");
      if (!result.success) return "";
      const given = typeof result.data?.givenName === "string" ? result.data.givenName : undefined;
      const display = typeof result.data?.displayName === "string" ? result.data.displayName : undefined;
      return firstName(display, given);
    },
  });
  if (!value?.trim()) return <>{fallback}</>;
  if (!address) return <>{firstName(value) || fallback}</>;
  if (lookup.isPending) return <span aria-busy="true">…</span>;
  return <>{lookup.data || aliasOf(address) || fallback}</>;
}
