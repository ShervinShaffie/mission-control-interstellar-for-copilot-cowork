# app/

The source code of Mission Control Interstellar: a Copilot App Builder app (React, TypeScript and Vite) that runs on Microsoft Copilot Managed Runtime.

You don't need to work with these files by hand. To install the app, follow the main [README](../README.md). Copilot sets everything up from this folder using [INSTALL-WITH-COPILOT.md](../INSTALL-WITH-COPILOT.md).

**This folder holds only the app's own code.** That means the files written for Mission Control Interstellar, plus three of the template's files that it changes: `src/index.css`, `src/lib/query-client.ts` and `src/pages/home.tsx`. Everything else comes from the App Builder template when Copilot creates the app in your environment:

- the build setup and its supporting files
- the shared UI components
- the template's own helper components

These are left out on purpose:

- `generated/`: the typed clients for the app's data connections. These are created in your environment when Copilot connects the data sources.
- The file that ties an app to one organization's environment. Each installation creates its own.
- Template files, preview logs and downloaded packages.

The header image, `src/assets/interstellar-banner.jpg`, is an original image created for this project.
