# Mission Control Interstellar

Your whole workday on one screen: a Microsoft 365 dashboard you install with Copilot Cowork, which also brings in GitHub, Google Drive, Google Calendar and Box.

![Mission Control Interstellar header: an astronaut walking across an icy planet](app/src/assets/interstellar-banner.jpg)

Version 1.0.0. See [CHANGELOG.md](CHANGELOG.md) for the full history.

Read the story behind it in **[Build a Mission Control App for About $5 To Run Your Workday](https://cosi.pro/work-mission-control-app-by-copilot-cowork)** on Collaboration Simplified.

Mission Control Interstellar is a dashboard app built with Copilot App Builder. It runs on Microsoft Copilot Managed Runtime, inside your own Microsoft 365 tenant. It reads your calendar, mail, Teams chats, files, notes and tasks, lines up what needs you today, and lets you act on it without jumping between apps. It goes beyond Microsoft 365 too. Your Google calendar sits next to Outlook with overlaps flagged, and you can browse, preview, edit and upload files in Google Drive and Box from the same screen.

Not an official Microsoft product. This is a personal project, shared as is for personal and demo use. It is not built, endorsed, or supported by Microsoft, and it is not meant for production. Make sure you review the Licensing section before you run it. See the Disclaimer and license section below.

## What it does

The dashboard is a set of sections. You can show, hide, reorder and collapse any of them.

| Section | What you get |
|---|---|
| **Today** | Your plan for the day, what needs attention now, and the most important actions across everything |
| **Quick actions** | Draft an email, send a meeting invite, add a To Do task, create a OneNote page or a document, plus your activity history |
| **Projects** | Your project teams, most in need of attention first, with status, deadlines, meetings and tasks |
| **Calendar** | Outlook and Google Calendar together in one view (1, 3 or 7 days), with overlaps highlighted, To Do and Planner tasks on their due days, and a day agenda |
| **Email** | Emails waiting on a reply, unread, flagged and recent mail |
| **Teams** | Who's waiting on you, and your recent chats and team activity |
| **Files** | Files that need you, from your email, Teams and meetings, plus your OneDrive |
| **OneNote** | Recent pages and open to-dos |
| **Planner** | Your Planner tasks across every plan: overdue, due this week and in progress, with progress, labels and checklists |
| **Plugins** | Google Drive, Box, Google Calendar and GitHub, each in its own collapsible panel, in the order you choose |
| **Time insights** | Your meeting load and how your time splits across calendar categories, against targets you set |

The plugins are two-way, not just read-only:

- **Google Drive and Box:** browse folders inside the dashboard, preview images, PDFs and text files, edit text files in place, create new files, upload text files, unzip archives, and delete files.
- **Google Calendar:** add, edit, duplicate and delete events, search the next 90 days, and look 7, 14 or 30 days ahead.
- **GitHub:** your profile, your repositories with their details, milestones, and the open issues and pull requests that involve you.

Wherever a feature has a limit, an **i** button next to it explains what it is.

Every **Draft with Copilot** button prepares a ready-made request from what's on the dashboard, copies it, and opens Microsoft 365 Copilot. The dashboard itself never calls an AI service.

## What you need

- **A Microsoft 365 work account.** The dashboard uses your own Outlook, Teams, OneDrive, OneNote, To Do and Planner.
- **A license to run it:** Power Apps Premium, or Copilot Credits. See Licensing below.
- **To install it with Copilot:** Copilot Cowork with App Builder (a Microsoft Copilot license, Copilot Credits, and a tenant enrolled in the Frontier program).
- **Your admin's go-ahead** for the GitHub, Google Drive, Google Calendar and Box connections. See Admin setup below. The Microsoft 365 parts work without this.
- **Optional:** your own GitHub, Google and Box accounts for the Plugins section. Each person signs in to their own accounts.
- **Optional:** a Microsoft 365 Copilot license, to get the most out of the Draft with Copilot buttons.

You don't need to add any plugins in Copilot's Customize page. The dashboard uses its own connections, which Copilot sets up during the install.

## Licensing

Mission Control Interstellar runs on Microsoft Copilot Managed Runtime, which is in preview. Licensing for preview features can change, so check Microsoft's current guidance before you roll it out.

**To run the app**, each person needs one of these:

- **Power Apps Premium.** Running the app doesn't use Copilot Credits. Each data read or write counts toward the daily request allowance that comes with the license.
- **Copilot Credits (pay as you go).** Opening the app and each data read or write is charged in Copilot Credits. Your admin turns on usage-based billing and adds you to a spending policy for managed apps.

The Power Apps rights included with Microsoft 365 don't cover this app, even though most of its connections are standard Microsoft 365 ones. Sharing the app with someone doesn't give them a license.

**To install it with Copilot**, you need Copilot Cowork: a Microsoft Copilot license, Copilot Credits, and a tenant in the Frontier program. Building and installing are billed separately from running the app.

**Microsoft 365 Copilot is optional.** It makes the Draft with Copilot buttons more useful, but the dashboard works fully without it.

**How it keeps running costs down:**

- The dashboard never uses AI that's charged per use. Every rule and summary runs in the app itself, and deeper questions hand off to your own Copilot.
- It never refreshes on its own. It reads data when you open it, select **Refresh**, or open something like a folder, a preview or a repository's details.

## Admin setup

Apps built with Copilot App Builder follow a connector policy that your admin controls. Microsoft's default policy already allows the Microsoft 365 connectors the dashboard uses (Office 365 Outlook, Office 365 Users, Microsoft Teams, OneDrive for Business, OneNote (Business), Microsoft To-Do (Business) and Planner), so those work straight away.

The Plugins section needs four more. An admin (Global Administrator or Power Platform Administrator) adds them once:

1. Go to the **Microsoft 365 admin center → Apps → Settings → Environment groups**.
2. Open the default environment group for Copilot Managed Runtime, then the **Connectors and MCP servers** rule.
3. Select **Edit this policy**. This takes the connector list out of Microsoft's automatic updates and puts it under your organization's control.
4. Select **Add connectors** and add **GitHub**, **Google Drive**, **Google Calendar** and **Box**.
5. Check that the actions the dashboard uses are switched on, including uploading or creating files and creating, updating and deleting events.

Changes can take a while to reach your environment. Until they do, the dashboard shows a "Blocked by your organization's policy" note, and the rest of the dashboard keeps working.

## Install it

Let Copilot do it. Copilot Cowork can't download files from GitHub, so you hand it the release zip:

1. Download the latest `mission-control-interstellar-for-copilot-cowork-v<version>.zip` from the [Releases](../../releases) page. You don't need to unzip it.
2. Open **Copilot Cowork**, start a new task, and attach the zip.
3. Paste this:

> Install Mission Control Interstellar from the attached zip. Read INSTALL-WITH-COPILOT.md inside it and follow it exactly, including the stop conditions.

From there, Copilot does everything in that same task:

1. It creates the app in your environment and loads the code.
2. It connects the eleven data sources the dashboard uses.
3. It asks you to sign in to GitHub, Google and Box with **your own** accounts.
4. It checks the app and publishes it.

When it's done, you get a link to your own copy of Mission Control Interstellar.

Always attach the zip. Pasting a link to this repository doesn't work, because Cowork can't download the code from GitHub.

## Using it

- **Customize** shows, hides and reorders sections. Each section has a collapse button at its top right.
- **Connections** shows whether each data source is working, and lets you sign in again or switch accounts.
- **Refresh** reloads everything on the dashboard.
- **History** keeps a record of what you did from the dashboard, in your browser and in a file in your own OneDrive.

## Platform support

Mission Control Interstellar runs in the browser, on Windows, macOS or anything else that runs a modern browser, wherever you can open Microsoft 365 apps. Installing it needs Copilot Cowork with App Builder. There's nothing to install on your computer.

## Privacy

- **Your data stays where it is.** The dashboard reads your data through your own signed-in connections, inside your organization's environment. It doesn't copy your data anywhere else.
- **Each person sees only their own accounts.** Connections can't be shared, so everyone you share the app with signs in to their own GitHub, Google and Box accounts. Each plugin panel shows which account it's using, with a link to check or change it.
- **Where things are stored:** your layout and settings stay in your browser. Your activity history is saved to a file in your own OneDrive.
- **Nothing goes out on its own.** The dashboard never sends emails, posts messages or changes files on its own. Everything happens when you select it, and deleting always asks you first.
- **The code is checked for personal data and tenant details by verify-clean.** `scripts/verify-clean.mjs` fails on any email address, tenant, environment or app ID, connection ID, or tenant-specific link.

## Disclaimer and license

This is a personal, community project, and it is not an official Microsoft product. It is not built, endorsed, supported, or maintained by Microsoft. Names like Microsoft 365, Copilot, Power Apps, GitHub, Google Drive, Google Calendar and Box are used only to describe the services this tool works with. Any views or work here are the author's own and do not represent Microsoft.

It is provided as is, for personal and demo use, and it is not meant for production. Mission Control Interstellar is shared free of charge with no warranty of any kind. It is provided without support or any obligation to fix issues or provide updates. It can read and act on your own Microsoft 365, GitHub, Google and Box data. With your approval in the dashboard, it can create, change and delete files and calendar events, so you use it at your own risk. To the maximum extent allowed by law, the author is not liable for any outcome, data loss, missed or mistaken change, or other damage that comes from using it.

The header image is an original image created for this project and is not from any film.

Licensed under the [MIT License](LICENSE).

Built by Shervin Shaffie.
