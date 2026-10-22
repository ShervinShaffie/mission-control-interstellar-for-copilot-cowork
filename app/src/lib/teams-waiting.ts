import { loadChatMessages } from "@/lib/command-center-data";
import type { ChatRecord } from "@/lib/command-center-evidence";
import { firstName, hideEmails } from "@/lib/privacy";

/** A chat whose latest real message came from someone else recently: they are likely waiting on you. */
export interface WaitingChat { chatId: string; chatTitle: string; url?: string | undefined; messageUrl: string; person: string; preview: string; created: string; group: boolean; files?: { name: string; url: string }[] }

/** Deep link to one message in a chat (opens Teams at that message). */
export function teamsMessageLink(chatId: string, messageId?: string) {
  if (!messageId) return `https://teams.microsoft.com/l/chat/${encodeURIComponent(chatId)}/0`;
  return `https://teams.microsoft.com/l/message/${encodeURIComponent(chatId)}/${encodeURIComponent(messageId)}?context=${encodeURIComponent(JSON.stringify({ contextType: "chat" }))}`;
}

const MAX_CHATS = 10, MAX_AGE_DAYS = 7;

export function plainPreview(html?: string) {
  const text = (html ?? "").replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, "\"").replace(/&#39;/g, "'").replace(/\s+/g, " ").trim();
  return hideEmails(text);
}

/**
 * Free Teams reads (no Copilot): for up to 10 chats, read the latest messages and keep chats whose newest
 * real message was sent by someone else in the last 7 days. Newest first.
 */
export async function loadWaitingChats(oneToOne: ChatRecord[], group: ChatRecord[], myId: string, now = new Date()): Promise<WaitingChat[]> {
  const chats = [...oneToOne.map(chat => ({ chat, group: false })), ...group.map(chat => ({ chat, group: true }))].filter(item => !item.chat.hidden).slice(0, MAX_CHATS);
  const cutoff = now.getTime() - MAX_AGE_DAYS * 864e5;
  const found = await Promise.all(chats.map(async ({ chat, group: isGroup }): Promise<WaitingChat | undefined> => {
    try {
      const page = await loadChatMessages(chat.id);
      const latest = page.rows.filter(row => !row.deleted && row.messageType === "message" && row.created).sort((a, b) => Date.parse(b.created!) - Date.parse(a.created!))[0];
      if (!latest || !latest.senderId || latest.senderId === myId || Date.parse(latest.created!) < cutoff) return undefined;
      const person = firstName(latest.sender) || "Someone";
      const preview = plainPreview(latest.summary || latest.content).slice(0, 160);
      if (!preview && !latest.files?.length) return undefined;
      const title = chat.topic && chat.topic !== "Chat without a title" ? hideEmails(chat.topic) : isGroup ? `Group chat with ${person}` : `Chat with ${person}`;
      return { chatId: chat.id, chatTitle: title, url: chat.url || teamsMessageLink(chat.id), messageUrl: teamsMessageLink(chat.id, latest.id), person, preview: preview || "Shared a file", created: latest.created!, group: isGroup, files: latest.files ?? [] };
    } catch {
      return undefined; // one unreadable chat never hides the others
    }
  }));
  return found.filter((item): item is WaitingChat => !!item).sort((a, b) => Date.parse(b.created) - Date.parse(a.created));
}

export function agoLabel(iso: string, now = new Date()) {
  const minutes = Math.max(0, Math.round((now.getTime() - Date.parse(iso)) / 60000));
  if (minutes < 60) return `${minutes} min ago`;
  const hours = Math.round(minutes / 60);
  if (hours < 24) return `${hours} h ago`;
  const days = Math.round(hours / 24);
  return days === 1 ? "yesterday" : `${days} days ago`;
}
