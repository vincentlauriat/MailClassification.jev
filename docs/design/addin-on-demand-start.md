# Design — Outlook add-in: start jevOutlook on demand (option 2)

Status: **designed, not built** (2026-09-28). Option 1, the macOS LaunchAgent
(`jevoutlook service install`), is what ships today; see ARCHITECTURE_EN.md §10.

## Problem

The add-in's task pane (`taskpane.html`) is served by `jevoutlook ui` on
`https://localhost:5178`. When that process is not running, Outlook has nothing to load,
and an Office web add-in runs in a sandboxed webview: it cannot spawn a local process.
Option 1 avoids the problem by keeping the server up at all times. Option 2 lets the pane
itself start the server when the user needs it, so nothing has to run in the background.

## Overview

```
Outlook ──loads──▶ static pane (GitHub Pages, https)
                     │ probe GET http://127.0.0.1:5177/health
                     ├─ up   → navigate to https://localhost:5178/taskpane.html (today's pane)
                     └─ down → "Start jevOutlook" button → jevoutlook://start
                                                            │
                                   jevOutlook Launcher.app ◀┘ (CFBundleURLTypes: jevoutlook)
                                     ├─ agent installed → launchctl kickstart gui/$UID/com.vincentlauriat.jevoutlook
                                     └─ no agent       → jevoutlook ui --no-open (detached)
                     pane keeps probing /health, then switches to the local pane
```

## 1. Static pane

- Hosted on GitHub Pages next to the landing page: `docs/addin/pane/index.html` (+ icons). `docs/addin/` itself is the add-in documentation page.
- On load it probes `http://127.0.0.1:5177/health` with `fetch` and a short timeout. It
  accepts only a body whose `app` is `jevoutlook` (same rule as `UiServer.ParseHealth`).
- Server up: `location.replace('https://localhost:5178/taskpane.html')`. The Office.js
  context carries over because the local origin is listed in the manifest `<AppDomains>`.
- Server down: show "Start jevOutlook" linking to `jevoutlook://start`, then poll `/health`
  every second for about 20 s and switch as soon as it answers. After the timeout, show the
  manual fallback: `jevoutlook service install`.

## 2. Server changes it will need

- **CORS on `GET /health` only**, for the Pages origin (`https://vincentlauriat.github.io`):
  - `Access-Control-Allow-Origin: <that origin>`, no credentials.
  - Answer the preflight with `Access-Control-Allow-Private-Network: true`. Chrome's
    Private Network Access asks for it before a public https page may reach 127.0.0.1.
  - It goes through `UiServer.AllowedOrigins(port, httpsPort)`: add a separate
    "health-only" origin set there. Never widen `/api/*`, which stays same-origin, JSON-only
    and POST-only.
- **Manifest**: `SourceLocation` and `Taskpane.Url` point to the Pages URL.
  `<AppDomains>` keeps `https://localhost:5178` so the pane may navigate there.
  `BuildManifest` gains a base-URL parameter, and `addin manifest` gets
  `--pane static|local`, with `local` as the default until option 2 is validated.

## 3. Launcher `.app`

- A minimal bundle, `jevOutlook Launcher.app`. Its `Info.plist` declares `CFBundleURLTypes`
  with the scheme `jevoutlook`, and `LSUIElement` true so it has no Dock icon.
- Handler, for `jevoutlook://start` only; any other path is ignored:
  1. If the LaunchAgent is loaded (`launchctl print gui/$UID/com.vincentlauriat.jevoutlook`),
     run `launchctl kickstart gui/$UID/com.vincentlauriat.jevoutlook`.
  2. Otherwise, start `~/.jevoutlook/bin/jevoutlook ui --no-open` detached, logging to
     `~/.jevoutlook/logs/ui.log`.
  3. Quit. It never takes arguments from the URL, so a web page cannot pass options.
- Language: a small Swift AppKit target, or an AppleScript applet. Processes are started
  with argument arrays, never a shell string.
- Signed with the Developer ID and notarized through the standard sign + notarize pipeline
  (hardened runtime, `notarytool`, stapled DMG in `release/`).

## 4. Known risks

- **Custom schemes in Outlook's webview.** New Outlook, Outlook on the web and Outlook for
  Mac may block or silently drop navigation to `jevoutlook://`. Test all three first.
  Fallback: `Office.context.ui.openBrowserWindow('jevoutlook://start')`, or a "copy this
  command" hint.
- **Mixed content.** The https pane probing `http://127.0.0.1` is allowed for loopback in
  Chromium and WebKit, but embedded webviews may be stricter. The alternative is to probe
  `https://localhost:5178/health`, which needs the dev certificate trusted, as it is today.
- **Private Network Access.** Chrome's rules keep evolving (preflight today, a permission
  prompt later). Probing only `/health`, which returns no data, limits the exposure.
- **Tenant policy.** Tenants that disable custom add-ins still refuse the sideload,
  whatever the pane's origin.
- **Security.** A public page can now learn whether jevOutlook runs on the machine (the
  version string). That is accepted; `/api/*` stays unreachable cross-origin.

## 5. Steps (see TODOS.md, "Option 2 — on-demand start")

1. Spike: can each Outlook client open `jevoutlook://` from a pane? This decides go or no-go.
2. `/health` CORS + PNA preflight for the Pages origin, with tests on `AllowedOrigins`.
3. Static pane `docs/addin/pane/`, with probe, switch, start button and fallback.
4. `BuildManifest(baseUrl)` and `addin manifest --pane static|local`.
5. Launcher `.app` with URL handler; sign, notarize, DMG.
6. End-to-end test on Outlook for Mac and new Outlook; update README §6.6.
