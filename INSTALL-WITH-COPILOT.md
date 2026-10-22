# Install Mission Control Interstellar (guide for Copilot)

You are Microsoft Copilot, working in a Copilot Cowork task, and you are installing Mission Control Interstellar for the user as an App Builder app in their own environment. Follow this guide in order. Use your judgment to recover when a step fails, but stay inside the guardrails at the bottom. They are not optional.

## What "done" looks like

The whole job is complete when all of these are true:

1. A new App Builder app named **Mission Control Interstellar** exists in the user's environment, and its `src/` includes every file from this repository's `app/src/` folder.
2. All eleven data sources in Step 4 are added to the app.
3. Every connection the user agreed to is connected with **the user's own accounts**. Anything that couldn't be connected, or that the organization's policy blocks, is reported plainly.
4. The app passes its checks and is **published**, and the user has the app's link.

When all four are true, stop.

## Before you start

- **Confirm App Builder is available.** You need it in this task. If it isn't available (for example, the tenant isn't in the Frontier program), stop and tell the user they need Copilot Cowork with App Builder.
- **Confirm it's a new install.** If the user already has an app named Mission Control Interstellar, ask whether to install a separate copy (named "Mission Control Interstellar 2") or stop. Never overwrite or change an existing app.
- **Tell the user what to expect, in one short message:**
  - Installing uses their Copilot Credits.
  - Running the app needs Power Apps Premium or Copilot Credits.
  - They'll be asked to sign in to GitHub, Google and Box, which is optional.

## Step 1: Get the package

Work out which of these applies, in order:

1. **An attached package:** if the user attached `mission-control-interstellar-for-copilot-cowork-v<version>.zip`, extract it fully. The extracted folder is the package folder.
2. **The latest release:** otherwise, download the latest release's zip from `https://github.com/ShervinShaffie/mission-control-interstellar-for-copilot-cowork/releases` and extract it.
3. **The repository:** if you can't download the release, read the files from the public repository `https://github.com/ShervinShaffie/mission-control-interstellar-for-copilot-cowork`.

Check that the package folder has `app/src/pages/home.tsx`, `app/src/components/plugins-panel.tsx` and `app/src/assets/interstellar-banner.jpg`. If it doesn't, you have the wrong folder. Fix that first, and if you can't, stop and ask the user to attach the release zip.

## Step 2: Create the app

Create a new App Builder app named **Mission Control Interstellar**. Wait until the app's environment is set up (its configuration file appears) before going on.

## Step 3: Load the code

The package's `app/` folder holds only the app's own code. The new app's template provides everything else.

- **Copy** every file under the package's `app/src/` into the new app's `src/`, keeping the same paths.
  - This adds the app's own files.
  - It replaces three template files the app changes: `src/index.css`, `src/lib/query-client.ts` and `src/pages/home.tsx`.
- **Don't copy** `app/README.md`, and don't edit the template's protected build files (`package.json`, the lock files, `vite.config.ts`, `tsconfig*.json`, `eslint.config.js`, `index.html` and similar).
- **Never create or edit the app's environment configuration file by hand**, except its `description`. You may set that to: "Mission Control Interstellar: a Microsoft 365 dashboard with GitHub, Google Drive, Google Calendar and Box."
- **If the code doesn't compile after Step 4** because the template has changed (for example, a template component it uses was renamed), stop and report exactly what's missing. Don't rewrite the app to make it fit.

## Step 4: Add the data sources

Add these eleven connectors to the app as data sources, one at a time, using App Builder's data source tools:

| Connector | ID | Used for |
|---|---|---|
| Office 365 Outlook | `shared_office365` | Calendar, mail, drafts, invites |
| Office 365 Users | `shared_office365users` | Profile, people search |
| Microsoft Teams | `shared_teams` | Chats and teams |
| OneDrive for Business | `shared_onedriveforbusiness` | Files and the activity history file |
| OneNote (Business) | `shared_onenote` | Pages and to-dos |
| Microsoft To-Do (Business) | `shared_todo` | Tasks |
| Planner | `shared_planner` | Planner tasks and plans |
| GitHub | `shared_github` | Plugins: GitHub |
| Google Drive | `shared_googledrive` | Plugins: Google Drive |
| Google Calendar | `shared_googlecalendar` | Plugins and the combined calendar |
| Box | `shared_box` | Plugins: Box |

After adding them, check that each connector's generated service file exists and that the app's code compiles against them. The app expects services named `Office365OutlookService`, `Office365UsersService`, `MicrosoftTeamsService`, `OneDriveforBusinessService`, `OneNote_Business_Service`, `MicrosoftTo_Do_Business_Service`, `PlannerService`, `GitHubService`, `GoogleDriveService`, `GoogleCalendarService` and `BoxService`.

Don't add any other connectors. In particular, don't add Work IQ, Copilot or other AI connectors. The app is designed to run with no AI charged per use.

## Step 5: Connect the user's accounts

- **Microsoft 365 connectors** usually connect automatically with the user's work account. Confirm each one is usable.
- **GitHub, Google Drive, Google Calendar and Box** need the user to sign in:
  1. Ask once whether they want them now. All four are optional, and the rest of the dashboard works without them.
  2. For each one they want, create the connection only with their approval, and give them the sign-in link.
  3. Tell them to sign in with **their own** account, and to check which Google account they pick if they have more than one.
  4. After they confirm, check that the connection shows as connected.
- If a connector is **blocked by the organization's policy**, don't try to work around it. Tell the user plainly that their admin needs to add it. The steps are in the README's Admin setup section. Then carry on.

## Step 6: Check and publish

1. Run the app's checks (type check, lint and build). If they fail, fix only what this install changed. If the failure is in the app's own code or a protected file, stop and report it.
2. Open the preview long enough for the dashboard to load, then check the preview's error log.
   - Connector-policy blocks and connections the user chose not to set up are expected. Report them, but don't treat them as failures.
   - Anything else is a failure. Stop and report it.
3. Publish the app.

## Step 7: Hand over

Give the user a short, plain summary:

- The app's link.
- Which data sources are connected.
- Which need their admin (from Step 5).
- How to share the app. Each person they share it with needs their own license and signs in to their own accounts.

## Guardrails (these are not optional)

- Don't attempt the whole install more than twice. If it fails twice, stop and show the user exactly what you saw. Don't keep trying.
- Retry any single step at most once, and never re-run a step that already succeeded.
- Never change, overwrite or delete an existing app, and never edit the app's protected build files.
- Never create a connection without the user's approval. Never connect with anyone's account but the user's own.
- Never work around a connector policy block. That's the admin's decision.
- Don't add connectors, features or AI services that aren't in this guide.
- The definition of done is at the top of this guide. The moment it's met, stop.
- Only do what this guide says. If a situation isn't covered here, stop and ask the user rather than inventing a step.
