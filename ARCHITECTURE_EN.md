# MailClassification — Architecture (source of truth)

> French mirror: `ARCHITECTURE.md`. Keep both in sync.

## 1. Purpose

Port of the jevMail Gmail/Apps Script classifier to a **local multi-mailbox client**:
classify messages of several mailboxes — Microsoft 365 / Outlook.com (Graph), Gmail
(IMAP + app password) and any IMAP server — with the Jev decision model (TypeSafe,
reached through OpenRouter), apply one label per message (Outlook category, Gmail
label or IMAP keyword), optionally archive high-confidence disposable mail, under a
strict run budget and with resumable, checkpointed processing. No hosted backend, no
Azure work beyond one shared app registration for the Microsoft mailboxes.

One local process, three ways in, all sharing the same state (`~/.mailclassification`):

| Surface | Served by | Use |
| --- | --- | --- |
| **Dashboard** | `ui` (ASP.NET Core on 127.0.0.1:5177) | setup, preview runs, batches on every mailbox |
| **CLI** | the same executable (`mailclassification run`, `continue`…) | scripting, `--all-accounts` |
| **Outlook add-in** | the same `ui` process over HTTPS (localhost:5178), kept up by the LaunchAgent (§10) | classify / apply on the open message, a preview-first Inbox run, inside Outlook (Microsoft 365 / Outlook.com) |

The add-in pane, its API and its manifest are built and tested; sideloading into a real Outlook
client is not yet validated end to end (see the Outlook add-in row in §3 and [docs/ADDIN.md](docs/ADDIN.md)).

The product was called **jevOutlook** until 2026-09-28; the rename and its migration are in §7.

## 2. Stack

| Concern | Choice | Why |
| --- | --- | --- |
| Language / runtime | C# 13, .NET 10 (`net10.0`), one executable: local web dashboard + CLI | Microsoft-platform language; the dashboard mirrors the Apps Script web app, the CLI suits scripting |
| Microsoft mailboxes | Microsoft Graph v1.0 REST via `HttpClient` | Full control of `$batch`, `Prefer` headers and OData filters; no SDK version drift. Exchange Online has no password IMAP, so Graph + Microsoft sign-in is the only path |
| Microsoft sign-in | `Azure.Identity` (`DeviceCodeCredential`, `InteractiveBrowserCredential`) | Delegated permissions, persisted MSAL token cache, one `AuthenticationRecord` per account; one Entra app registration shared by every Microsoft mailbox |
| Gmail / IMAP mailboxes | MailKit (`ImapClient`) | Mature IMAP client with Gmail extensions (`X-GM-LABELS`), keywords, special-use folders |
| Passwords | macOS Keychain through `/usr/bin/security` (service `mailclassification`, account = account id); `secrets.json` 0600 elsewhere | Nothing secret next to the configuration; no extra dependency |
| Model | Jev `~typesafe/jev-latest` via OpenRouter Decisions (`/api/alpha/decisions`); optional direct TypeSafe (`/v1/systemone`) | Same request/response contract; OpenRouter adds `usage.cost` |
| State | JSON files in `~/.mailclassification/` (0600), per-account subdirectories | Counterpart of Apps Script User Properties |
| Tests | xUnit | Pure logic: rules, payload, contract parsing, body compaction, filters, cursor paging, IMAP helpers |

## 3. Component map

```
Cli/Program.cs ─┐                                   ┌─► Graph/GraphMailClient ──► Microsoft Graph
Web/UiServer  ──┼─► Triage/TriageEngine ─► Mail/IMailbox ┤
                │          │                        └─► Imap/ImapMailbox (MailKit) ──► imap.gmail.com / any IMAP
                │          ├─► Jev/JevClient ──► OpenRouter / TypeSafe
                │          └─► Storage/JobStore (per account)
                └─► Mail/MailboxFactory ─► Storage/AccountStore, SecretStore, Graph/GraphAuth
```

| Module | Responsibility |
| --- | --- |
| `AppConstants` | All tunables (concurrency windows, retry delays, thresholds, defaults) |
| `Rules/*` | `LabelRule` record, 7 playbooks (identical to jevMail), validation (≤12 rules, no `,`/`;`, no two names mapping to the same IMAP keyword, id regeneration) |
| `Mail/IMailbox` | The exact mailbox surface the engine needs: identity, capabilities, estimate, cursor listing, batched reads (`ReadPart`), `StoredLabel` (rule name → stored label), `EnsureLabels`, `ApplyLabels`, `Archive`, label cleanup (`FindMessagesWithLabel`, `RemoveLabel`, `DeleteLabel`). Ids and cursors are opaque strings owned by the provider; `MailboxException(transient)` drives pause vs. error. `IAsyncDisposable`: an IMAP mailbox holds a live connection, Graph disposal is a no-op |
| `Mail/MailboxFactory` | Opens the right provider for a `MailAccount` (Graph: saved sign-in record; IMAP: password from the secret store); interactive Microsoft sign-in; readiness check |
| `Mail/MessageMetadata`, `MessageContent` | The flat `email` object sent to Jev (built from Graph JSON or IMAP envelope/headers); HTML→text; head/tail compaction |
| `Storage/*` | Paths, atomic JSON writes with 0600, `AppConfig` (global), `RuleStore` (global rules), `AccountStore` (`accounts.json`), `SecretStore`, `JobStore` (one per account: job + frozen rules + the `job.lock` cross-process lock) |
| `Graph/GraphAuth` | Credential per account (tenant override), device-code / browser sign-in, record persistence under `accounts/<id>/`, legacy single-mailbox migration |
| `Graph/GraphMailClient` | `IMailbox` over Graph: `receivedDateTime` cursor listing, `$batch` reads, master categories (colour presets), category PATCH, archive move, throttling handling |
| `Imap/ImapMailbox` | `IMailbox` over MailKit: UID cursor listing, envelope/header fetch, Gmail labels or IMAP keywords (add-only writes; `ImapSupport.ToKeyword` maps names, reserved ones to `jev-…`), archive (All Mail / Archive folder), one serialised connection with reconnect, closed on dispose |
| `Jev/JevPayloadBuilder` | Builds the Choice question (options `L0..Ln`), size guard (29 000 bytes), cost estimate |
| `Jev/JevClient` | Sends one Decisions request; strict contract validation; cost accounting source |
| `Triage/TriageEngine` | Session lifecycle, batches, two waves, adaptive concurrency, budget, circuit breaker, grouped writes — provider-agnostic |
| `Cli/*` | `account` commands, `--account` / `--all-accounts`, run/continue/status, terminal rendering, hidden password prompt; `ServiceCommand` (`service install/status/restart/uninstall`, macOS LaunchAgent, see §10) |
| `Web/UiServer` | Local dashboard (ASP.NET Core minimal API on 127.0.0.1): serves the embedded page and exposes jevMail's server functions as `POST /api/{function}` (args as JSON array) plus `addAccount` / `removeAccount` / `testAccount`; one engine and one device-code sign-in state per account. Engines live in a `ConcurrentDictionary`; a replaced or evicted engine's mailbox is disposed, and one-shot mailboxes (add-account probe, `testAccount`, `authStatus`, `classifyItem`, `applyItem`) are `await using`, so no IMAP connection leaks (Gmail caps simultaneous IMAP connections at 15). Job calls (`startTriageJob`, `processNextBatch`, `resumeTriageJob`, `cancelTriageJob`, `clearFinishedJob`) run under the in-process `_jobLock` and try-acquire the account's `job.lock`; `signOut` runs under `_jobLock` because it may dispose an IMAP mailbox. `GET /health` → `{app, version, https}` identifies a running MailClassification (single instance, `service status`); `AllowedOrigins(port, httpsPort)` is the one place that builds the Host/Origin allowlists |
| `Web/wwwroot/*` | `index.html` + `style.css` adapted from jevMail (MIT): a bridge emulates `google.script.run` over `fetch`; Mailboxes card (chips, add/test/sign-in/remove, "Run all mailboxes" queue); `taskpane.html` (Office.js) and icons for the Outlook add-in |
| Outlook add-in | Classic XML manifest served at `/manifest.xml`; the pane is served by the local server over HTTPS, which the LaunchAgent (§10) keeps running. One work tenant refused the sideload (generic "installation failed" although the manifest passes Microsoft's validator): tenants that disable custom add-ins still block it. An empty account id in the API resolves to the first ready Microsoft mailbox |

## 4. Provider mapping

| Concept | Outlook (Graph) | Gmail (IMAP) | Generic IMAP (e.g. Dovecot) |
| --- | --- | --- | --- |
| Label | Master category (`POST /me/outlook/masterCategories` with a colour) + `PATCH /me/messages/{id}` `{categories}` (full replacement → existing categories merged, never dropped) | Gmail label (a folder under the personal namespace, created if missing, named after the rule) added with `STORE +X-GM-LABELS`; never removed by a run (only `cleanup remove-category`, `-X-GM-LABELS`). Gmail's spam is the system label `\Spam`: a rule's label is a user folder under the personal namespace, and the writer never adds a `\…` label (`ImapSupport.LabelsToAdd`) | IMAP keyword (custom flag) added with `STORE +FLAGS` when the folder advertises `PERMANENTFLAGS \*`; otherwise labelling fails closed with a clear message. Name → keyword (`ImapSupport.ToKeyword`): atoms verbatim, other characters `_`, reserved names (`junk`, `nonjunk`, `notjunk`, `forwarded`, `phishing`, `mdnsent`, any `$…`, case-insensitive) → `jev-<name>` so a run never sets a client's junk or state marker. Never removed by a run (only `cleanup`, `-FLAGS`) |
| Archive | `POST /me/messages/{id}/move` `{destinationId:"archive"}` | Move Inbox → `[Gmail]/All Mail` (= remove the Inbox label; the message stays in All Mail) | Move to the `Archive` folder (special-use or created) |
| Marker | none: "already processed" = carries one of the configured labels | same | same, compared with each rule's stored keyword (`StoredLabel`), case-insensitively |
| Scope `inbox` / `all` | `/me/mailFolders/inbox/messages` / `/me/messages` with client-side exclusion of Junk, Deleted Items, Drafts | `INBOX` / `[Gmail]/All Mail` | `INBOX` only (`SupportsAllScope=false`, refused at start) |
| Unread | `isRead eq false` | `NOT SEEN` | `NOT SEEN` |
| Message id | Graph id (changes on move → 404 on replay handled) | `scope:uidvalidity:uid` | same |
| Sort key / cursor | `receivedDateTime` ISO-8601 (7 fractional digits, `Z`) | `uidvalidity:uid` zero-padded to 12 digits (lexicographic = numeric) | same |
| Metadata read | `$batch` of `GET …?$select=…,internetMessageHeaders` (20 per batch) | one `FETCH` per window: `ENVELOPE`, `FLAGS`, `X-GM-LABELS`, selected header fields, body structure, 2 KB text snippet | same without labels |
| Full read | `$select=id,body` with `Prefer: outlook.body-content-type="text"` | `FETCH BODY.PEEK[]` → `TextBody` or `HtmlBody`→text | same |
| Estimate | Inbox counts; `$count=true` for the whole mailbox | folder `Count` / `Unread` after `SELECT` | same |
| Rate limits | 429 `ApplicationThrottled` (MailboxConcurrency = 4 per mailbox; Graph fans a `$batch` out 4 at a time) → exactly **one `$batch` in flight**, window 5→20; reads retried 3× honouring Retry-After (≤30 s); writes sent batch by batch and retried 5× (≤30 s) | single connection, serialised; connection loss → transient pause | same |
| Sign-in | device code / browser, MSAL cache + per-account record | app password (2-step verification) in the Keychain | mailbox password in the Keychain |

## 5. Listing cursor

The engine lists newest-first from a provider-owned **sort key** and only advances the
cursor over messages it actually **examined**, so a page can hold more candidates than one
batch needs without losing any. Ids sharing the cursor key are kept in `CursorBoundaryIds`
and excluded client-side (ties are neither repeated nor lost); a full page of boundary ids
switches the next query to exclusive (`lt`). Because the cursor is a key and not a page
token, archiving messages out of the Inbox mid-run never shifts pages.

- Graph: `$filter=receivedDateTime le {cursor}[ and isRead eq false]&$orderby=receivedDateTime desc&$top=N`
  (Graph requires `$orderby` properties to appear first in `$filter`). Initial cursor = start + 5 min.
- IMAP: `UID SEARCH 1:{cursorUid}` (`cursorUid-1` when exclusive) `[NOT SEEN]`, sorted by UID
  descending, then `FETCH` of the page. Initial cursor = `UIDNEXT`. A `UIDVALIDITY` change fails
  the session with a clear message (UIDs are meaningless across a rebuild).
- Messages received after the session started are not part of it.

## 6. Processing algorithm (one batch)

1. **Fill pending** (≤ 50) from the cursor; complete when exhausted or the limit is reached.
2. **Metadata read** (adaptive window). Items already carrying a configured label without a
   final decision are dropped (processed elsewhere). Not found → skipped safely.
3. **Wave 1 — metadata**: one independent Choice request per pending item without a result.
   Successful answers are checkpointed before failures are handled.
4. **Decide**: confidence ≥ metadata threshold (and, in archive mode for an archive-eligible
   label, ≥ archive threshold) → final; otherwise the item goes to wave 2.
5. **Full read** + text extraction. Empty body → safe `review` fallback (non-archive) when
   such a rule exists, else skipped.
6. **Wave 2 — full body**: independent requests; answers checkpointed.
7. **Ready prefix**: leading pending items with a final decision are written in two grouped
   idempotent passes — the current labels of each message are **re-read in the same moment**
   (a replacement write and a minutes-old snapshot would clobber user changes), then
   `ApplyLabels` (current ∪ {label}), then `Archive` for archive decisions. Graph PATCHes that
   merged set (a full replacement, so the re-read matters); IMAP/Gmail only add the missing
   label (`ImapSupport.LabelsToAdd`) and never remove one, so a label set by a filter or another
   client between the re-read and the write survives. A transient
   mailbox failure pauses the session with every decision saved. On replay, a pending item
   that already has a final decision but is no longer found (its id changed because the
   archive move succeeded before the checkpoint) is counted as processed, not skipped.
8. **Counters** (processed, metadata-only, full, archived, per-label), results table,
   pending trimmed, job saved. Budget block → status `budget`.

### Wave dispatch (`DispatchWaveAsync`)

- Budget reservation before sending (`SpentUsd + reserve + estimate ≤ MaxSpendUsd`), replaced
  by the actual cost afterwards (`reported` > `input_tokens` > `estimated`).
- Requests run concurrently in windows of `JevConcurrency` (25 → 50, halved on failure).
- Session-scoped retryable failures (429/5xx/transport): if the whole window failed with the
  same code, **probe one request** before retrying the rest; otherwise retry the failed ones once.
- Message-scoped contract failures (invalid JSON, unknown option, bad probabilities) are
  retried once, bounded by the circuit breaker (3 consecutive → pause).

### Failure policy (fail closed)

| Situation | Effect |
| --- | --- |
| Provider 401/402/403/4xx | Session error / pause, no mailbox change |
| Provider 429/5xx after retry | Pause `provider-temporary`, resumable. A 429 is not billed (reservation released); timeouts/5xx keep the conservative estimate |
| Invalid model response ×1 | Retry once; then skip message, `ModelResponseSkips++` |
| 3 consecutive invalid responses | Pause `jev-response-circuit-breaker` |
| Mailbox 429/5xx / IMAP connection loss after retries (reads 3×, writes 5×, Retry-After honoured) | Pause `mailbox-temporary` (reads, pre-write re-read or grouped writes); `continue` resumes |
| Mailbox 401/403 / IMAP authentication refused | Error, ask to sign in again (`account login`) or store a new password (`account password`) |
| Message not found | Skipped safely (reads); treated as done (writes); counted as processed when the item already carries a final decision (replay after an archive move) |
| 40 listing pages without a new candidate | Pause `scan-guard` (everything already carries a configured label); `continue` keeps scanning |
| Ctrl+C during a wave | Reservations of unanswered requests are released; they are re-dispatched on resume |
| Budget would be exceeded | Status `budget`; `continue --max-spend` to raise |
| Ctrl+C | Pause `user-stop`; `continue` resumes (and re-arms the circuit breaker) |
| Another process holds the account's `job.lock` (dashboard batch vs. CLI `run`/`continue`) | Refused before any change: "Another MailClassification process is processing this mailbox right now (dashboard or CLI). Wait for it or stop it, then retry." `run --all-accounts` skips that mailbox (exit 1) |

## 7. Persistence

```
~/.mailclassification/
  config.json            client id, tenant, provider, endpoint/model overrides, optional API key, device-code flag
  rules.json             saved rules (shared by every mailbox)
  accounts.json          [{id, email, kind: graph|imap, host, port, gmail, username, tenantId}] — no secrets
  secrets.json           IMAP passwords, non-macOS only (macOS: Keychain, service "mailclassification")
  accounts/<id>/
    auth-record.json     MSAL AuthenticationRecord (Microsoft mailboxes; tokens live in the OS cache)
    job.json             current session: options, cursor (sort key), counters, spend, pending items with decisions
    job-rules.json       rules frozen for the current session
    job.lock             exclusive lock (FileStream + FileShare.None = advisory flock on Unix); never deleted
  logs/ui.log            stdout/stderr of the LaunchAgent
~/Library/LaunchAgents/com.vincentlauriat.mailclassification.plist   LaunchAgent (macOS), no secret
```

`<id>` is derived from the address (`alice@contoso.com` → `alice-contoso.com`). All
files are written atomically (`.tmp` + move) with mode 0600; directories are 0700. The
pre-multi-account layout (`auth-record.json`, `job.json` at the root) is migrated once into an
account entry on first start; the old session is dropped because its cursor type changed.

**One processor per session.** `job.lock` is held by the CLI for the whole `run` / `continue` loop (and
`clear-job`), and try-acquired (non-blocking) by the dashboard for each job call. Without it, the always-on
dashboard and a terminal `continue` could process the same `Running` session at once: both would pay for the same
messages and overwrite each other's `SpentUsd` (up to about twice the cap) and status (a cancel lost). The lock is
per open file description, so a second open fails even inside the same process; the kernel releases it when the
holder exits or crashes, and `using` releases it on exceptions and cancellation.

**Rename migration (jevOutlook → MailClassification, 2026-09-28).**

| Legacy | Current | How it moves |
| --- | --- | --- |
| `~/.jevoutlook/` | `~/.mailclassification/` | `AppPaths.MigrateLegacyRoot()`, first thing in `Main`: when no custom home is set and the new root is missing, `Directory.Move`; when the new root holds only `bin/` (a publish done before the first run), the legacy entries are moved one by one and names already present stay behind. Never runs once the new root has state. One line on stderr |
| `JEVOUTLOOK_HOME`, `JEVOUTLOOK_DEBUG` | `MAILCLASSIFICATION_HOME`, `MAILCLASSIFICATION_DEBUG` | the legacy names are read as fallbacks |
| Keychain service `jevoutlook` | `mailclassification` | lazy: on a read miss the legacy item is copied to the new service, then deleted (`SecretStore.GetWithLegacyFallback`); `Delete` removes both |
| LaunchAgent `com.vincentlauriat.jevoutlook` | `com.vincentlauriat.mailclassification` | `service install` boots out and deletes the legacy agent; `status` / `uninstall` report a leftover one |
| `/health` `app: "jevoutlook"` | `"mailclassification"` | `ParseHealth` accepts both, so an old server still running during the upgrade is not "foreign" |
| MSAL token cache name `jevoutlook` | **unchanged** | an internal identifier; renaming it would drop every cached Microsoft sign-in and force a new device-code login. Auth records (`accounts/<id>/auth-record.json`) move with the state directory |
| Add-in manifest `<Id>` | **unchanged** | stable identity; `DisplayName`/`ProviderName` changed, manifest version 1.0.0.0 → 1.1.0.0 |

## 8. Security notes

- Microsoft: delegated permissions only (`Mail.ReadWrite`, `MailboxSettings.ReadWrite`, `User.Read`);
  the app never sends mail. IMAP: the password is passed to MailKit only; it is never logged.
- Passwords are stored in the macOS Keychain via `/usr/bin/security` with `ArgumentList` (never
  through a shell) and never written to `accounts.json` or `config.json`. Gmail requires an **app
  password** (2-step verification), never the account password.
- Provider error bodies are never echoed (they may contain request data).
- Email content is passed to the model as untrusted data; the instructions tell the model never
  to follow instructions found in the email.
- The API key is read from the environment first; storing it on disk is opt-in.
- The dashboard listens on 127.0.0.1 only and rejects DNS-rebinding / cross-site calls: `Host` allowlist
  (421), same-origin `Origin` / `Sec-Fetch-Site` and `application/json` required on `/api/*` (403).
  Passwords entered in the page travel only to that loopback server. `GET /health` sits behind the same Host check.
- The LaunchAgent plist is plaintext: `BuildLaunchAgentPlist` writes only `DOTNET_ROOT`, `MAILCLASSIFICATION_HOME` and
  the legacy `JEVOUTLOOK_HOME` (allowlist), never `OPENROUTER_API_KEY` / `JEV_API_KEY`; the agent reads the key stored with `key set`.
  `launchctl` and `id -u` are started with `ArgumentList`, never through a shell.

## 9. Verification status (2026-09-28)

- 2026-09-24: `dotnet build`: 0 warnings, 0 errors. `dotnet test`: 66 tests pass (18 IMAP helper tests included; MailKit 4.18.0).
- Independent code-review pass (separate agent) + documentation verification of Graph/Azure.Identity/TypeSafe
  facts on 2026-09-22; Major findings fixed.
- Live, on a Microsoft 365 work mailbox: device-code sign-in, listing, `$batch` reads, both Jev stages,
  category PATCH (20 messages) and category cleanup; after the multi-account refactor the legacy
  state migrated in place and a dashboard preview run of 10 Inbox messages completed ($0.00067).
- 2026-09-24, live run of 100 Inbox messages on that Microsoft 365 mailbox (categories only): the first attempt paused twice on Graph
  429 `ApplicationThrottled` (two `$batch` in flight = 8 concurrent requests > MailboxConcurrency 4; the pre-write
  re-read had no retry). After the fix (one batch in flight, adaptive re-read, Retry-After-aware retries) the run
  completed: 99 categorized, 1 skipped (no readable text), 120 Jev requests, $0.009, zero 429.
- Not yet exercised live: the archive move; Gmail and generic IMAP (no credentials available yet:
  Gmail app passwords, the IMAP password, the second Microsoft tenant's admin consent).
- 2026-09-28: `dotnet build -c Release` 0 warnings, 0 errors; `dotnet test` 79 tests pass (+13: LaunchAgent plist
  contents, secret exclusion, XML escaping, `plutil -lint`, origin allowlist, `/health` body recognition).
- 2026-09-28: 88 tests pass (+9: the add-in's single-message rule — body review for an archive-eligible label below
  0.93, and "Apply + archive" offered only at 0.93, the batch run's archive threshold).
- 2026-09-28: 114 tests pass (+26 net: 28 new, 2 retired `Diff` tests), `dotnet build -c Release` 0 warnings, 0 errors: reserved
  IMAP keyword mapping, add-only label writes, keyword-based "already labelled" recognition and collision validation,
  `IMailbox` disposal (Graph no-op), per-account `job.lock` (same-process refusal, release on exception). A second
  process was checked by hand: refused while held (also by `flock(LOCK_NB)`), acquired once the holder exited.
- 2026-09-28, rename to MailClassification: 128 tests pass (+14: state-directory migration on temp directories — move,
  bin-only merge, no-op when the new root has state, idempotence — Keychain legacy fallback through injected delegates,
  legacy `/health` name, legacy LaunchAgent plist detection, manifest name and stable id). The published binary was run
  against a scratch `HOME` holding a seeded `.jevoutlook`: moved on the first command, no-op on the second.

## 10. Keeping the server running (macOS LaunchAgent)

The Outlook add-in pane is served by `mailclassification ui` itself, and an Office web add-in cannot start a local process.
Decision of 2026-09-28 (option 1): a per-user LaunchAgent keeps the server up from login onward.

| Piece | Behaviour |
| --- | --- |
| `service install [--port] [--https-port] [--exe]` | Writes `~/Library/LaunchAgents/com.vincentlauriat.mailclassification.plist` (`ProgramArguments` = exe `ui --no-open --port --https-port`, `RunAtLoad`, `KeepAlive`, `ThrottleInterval` 30, stdout/stderr → `~/.mailclassification/logs/ui.log`), then `launchctl bootout` (if loaded) + `bootstrap gui/<uid>`. Exe defaults to the running one; warns for `bin/Debug`/`bin/Release` and recommends a `dotnet publish` copy in `~/.mailclassification/bin` |
| `service status` | Plist present, agent loaded (`launchctl print`), `/health` probe, log path; exit 1 when the server does not answer |
| `service restart` | `launchctl kickstart -k gui/<uid>/<label>` (after a new publish) |
| `service uninstall` | `launchctl bootout`, then deletes the plist; mentions a leftover jevOutlook agent |
| Legacy agent | `service install` first boots out `com.vincentlauriat.jevoutlook` (if loaded) and deletes its plist (`ServiceCommand.LegacyPlistIn`), before bootstrapping and before the `/health` probe, since the probe also accepts the legacy server's answer |
| Single instance | `ui` first probes `http://127.0.0.1:{port}/health` (2 s): a MailClassification (or legacy jevOutlook) answer → "already running", open the browser, exit 0; any other answer → clear error instead of Kestrel's bind exception |

Known behaviour: while a `ui` started by hand holds the port, the agent's copy exits 0 ("already running") and launchd,
with `KeepAlive` true, starts it again every 30 s (one log line each time). `KeepAlive = {SuccessfulExit: false}` would
stop that, at the cost of not restarting after a clean exit.

Option 2 (on-demand start from a statically hosted pane through a `mailclassification://start` URL handler) is designed, not
built: see [docs/design/addin-on-demand-start.md](docs/design/addin-on-demand-start.md).

User-facing guide to the add-in (install, sideload per client, security, troubleshooting): [docs/ADDIN.md](docs/ADDIN.md).
