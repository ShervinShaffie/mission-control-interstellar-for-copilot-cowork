import { createContext, useContext } from "react";

/** Which dashboard components are collapsed to their title bar. Saved with the person's dashboard layout. */
export interface SectionCollapse { isCollapsed: (id: string) => boolean; toggle: (id: string) => void }
export const SectionCollapseContext = createContext<SectionCollapse | null>(null);
export function useSectionCollapse(): SectionCollapse | null { return useContext(SectionCollapseContext); }
