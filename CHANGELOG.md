# Changelog

This page lists what changed in each release of Mission Control Interstellar, newest first.

## What Mission Control Interstellar does today

Mission Control Interstellar is a Microsoft 365 dashboard you install with Copilot Cowork. As of the latest release:

- It puts your day on one screen: what needs attention, your calendar, mail, Teams, files, OneNote, To Do and Planner.
- It shows Outlook and Google Calendar side by side, in 1-, 3- and 7-day views, and highlights where they overlap.
- It lets you work in Google Drive, Box and Google Calendar without leaving the dashboard: browse, preview, edit, upload and add events. GitHub shows your repositories, issues and pull requests.
- It explains its own limits with an **i** next to anything that has one.
- It runs without AI that's charged per use. Deeper questions hand off to your own Copilot.

## Releases

### 1.0.0

First public release.

- **Started from Work Mission Control.** This release began as a copy of the author's earlier Work Mission Control dashboard, rebuilt on the current App Builder template, with its own settings and history so the two never mix.
- **Microsoft 365 at the core:**
  - Today, Quick actions, Projects, Calendar, Email, Teams, Files, OneNote and Time insights, all reading your data directly.
  - Activity history saved in a Mission Control Interstellar folder in your own OneDrive.
- **Planner section:**
  - Your Planner tasks across every plan, with Open, Overdue, Due this week and Completed views.
  - Progress, priority, labels, checklists and links to each task and plan.
- **Tasks on the calendar:**
  - To Do and Planner tasks appear on the day they're due, with a count on each day of the calendar graphic.
  - Today's agenda also lists overdue tasks and To Do tasks with no due date, which is where My Day tasks appear.
- **Combined calendar:**
  - Your main Google calendar alongside Outlook, with overlaps flagged.
  - 1-, 3- and 7-day views, with larger, easier-to-read text in the 1-day view.
  - Outlook stays read-only. Google events can be edited from the agenda.
- **Plugins section** with collapsible, reorderable panels:
  - **Google Drive and Box:**
    - Browse folders inside the dashboard, with thumbnails for images.
    - Preview images, PDFs and text files.
    - Edit text files in place, create new files, upload text files, unzip archives, and delete files (with a confirmation).
    - Selecting a file or folder keeps you in the dashboard; a separate button opens it at the source.
  - **Google Calendar:** add, edit, duplicate and delete events, search 90 days ahead, choose 7, 14 or 30 days, and expand any event's details.
  - **GitHub:**
    - Your profile and repositories with license, topics, branch, size and dates.
    - Per-repository details: open pull requests, issues and milestone progress.
    - The open issues and pull requests that involve you, with labels and assignees.
- **Know which account you're using:** each plugin panel shows the connected account where the service reports it, with a link to check or change it.
- **Collapse any section** from its top-right corner. Collapsed sections stay collapsed next time.
- **Limits explained in place:** an **i** button next to uploads, previews, the calendar graphic, to-dos, Planner and GitHub.
- **New header:** an original image of an astronaut on an icy planet, created for this project.
- **Privacy check:** `scripts/verify-clean.mjs` scans the repository before each release and fails on any personal or organization-specific data.
- **Fixes before release:**
  - Removed a request that used Outlook's general "Send an HTTP request" action. Microsoft's default policy blocks that action, and it showed a "blocked by your organization's policy" note even though everything worked.
  - OneNote to-do checks now read each page with its section's own ID, and a page that can't be read no longer counts as an error.
