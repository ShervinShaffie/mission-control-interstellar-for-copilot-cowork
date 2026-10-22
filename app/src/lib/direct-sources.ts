import type { IOperationResult } from "@microsoft/managed-apps/data";
import { OneNote_Business_Service } from "../../generated/services/OneNote_Business_Service";
import { readFailureCause, type ReadFailureCause } from "@/lib/command-center-failures";
import { hideEmails } from "@/lib/privacy";

export class DirectSourceError extends Error {
  constructor(readonly source: "SharePoint" | "OneNote", readonly failureCause: ReadFailureCause) {
    super(`${source} could not be read`);
    this.name = "DirectSourceError";
  }
}
async function read<T>(source: DirectSourceError["source"], call: () => Promise<IOperationResult<T>>): Promise<T> {
  let result: IOperationResult<T>;
  try {
    result = await call();
    if (!result.success) throw result.error;
  } catch (error) {
    const failureCause = readFailureCause(error);
    console.error(`${source} request failed`, { cause: failureCause });
    throw new DirectSourceError(source, failureCause);
  }
  return result.data;
}

// ---------- OneNote: recent pages and open to-dos ----------

/** One page's HTML for the to-do check. Uses the section's own id (the listing read takes its pages address). */
async function pageContent(notebookKey: string, sectionId: string, pageId: string): Promise<string | undefined> {
  try {
    const result = await OneNote_Business_Service.GetPageContent(notebookKey, sectionId, pageId, false);
    if (result.success && typeof result.data === "string") return result.data;
    console.warn("OneNote page content unavailable; its to-dos are shown as not checked", { cause: readFailureCause(result.error) });
  } catch (error) {
    console.warn("OneNote page content unavailable; its to-dos are shown as not checked", { cause: readFailureCause(error) });
  }
  return undefined;
}

export interface OneNotePage { id: string; title: string; notebook: string; section: string; modified?: string; url?: string; todos: string[]; todosChecked: boolean }
const MAX_NOTEBOOKS = 3, MAX_SECTIONS = 4, MAX_PAGES = 8, TODO_PAGES = 5, TODOS_PER_PAGE = 3;

/** Open (unchecked) to-do items in a OneNote page's HTML, as plain text. */
export function openTodos(html: string): string[] {
  const doc = new DOMParser().parseFromString(html, "text/html");
  return [...doc.querySelectorAll("[data-tag]")]
    .filter(node => (node.getAttribute("data-tag") ?? "").split(",").some(tag => tag.trim() === "to-do"))
    .map(node => hideEmails((node.textContent ?? "").replace(/\s+/g, " ").trim()))
    .filter(Boolean);
}

export async function loadOneNotePages(): Promise<OneNotePage[]> {
  const notebooks = await read("OneNote", () => OneNote_Business_Service.GetNotebooks());
  if (!Array.isArray(notebooks)) throw new DirectSourceError("OneNote", "response");
  const pages: (OneNotePage & { notebookKey: string; sectionKey: string; sectionId?: string })[] = [];
  for (const notebook of notebooks.slice(0, MAX_NOTEBOOKS)) {
    if (!notebook.Key) continue;
    const sections = await read("OneNote", () => OneNote_Business_Service.GetSectionsInNotebook(notebook.Key!));
    for (const section of (sections?.value ?? []).slice(0, MAX_SECTIONS)) {
      const sectionKey = section.pagesUrl ?? section.id;
      if (!sectionKey) continue;
      const list = await read("OneNote", () => OneNote_Business_Service.GetPagesInSection(notebook.Key!, sectionKey));
      for (const page of list?.value ?? []) {
        if (!page.id) continue;
        pages.push({ id: page.id, title: hideEmails(page.title || "Untitled page"), notebook: hideEmails((notebook.FileName ?? "Notebook").replace(/\.one(toc2)?$/i, "")), section: hideEmails(section.name ?? ""), modified: page.lastModifiedTime, url: page.links?.oneNoteWebUrl?.href, todos: [], todosChecked: false, notebookKey: notebook.Key, sectionKey, sectionId: section.id });
      }
    }
  }
  pages.sort((a, b) => (Date.parse(b.modified ?? "") || 0) - (Date.parse(a.modified ?? "") || 0));
  const recent = pages.slice(0, MAX_PAGES);
  // Optional enrichment: open to-dos from the most recent pages. A page whose content can't be read still shows
  // (its to-dos stay "not checked"), so a miss here is noted as a warning rather than an error.
  for (const page of recent.slice(0, TODO_PAGES)) {
    const html = await pageContent(page.notebookKey, page.sectionId ?? page.sectionKey, page.id);
    if (typeof html === "string") { page.todos = openTodos(html).slice(0, TODOS_PER_PAGE); page.todosChecked = true; }
  }
  return recent.map(page => ({ id: page.id, title: page.title, notebook: page.notebook, section: page.section, modified: page.modified, url: page.url, todos: page.todos, todosChecked: page.todosChecked }));
}
