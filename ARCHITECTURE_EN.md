# jevOutlook — Architecture (source of truth)

> French mirror: `ARCHITECTURE.md`. Keep both in sync.

## 1. Purpose

Port of the jevMail Gmail/Apps Script classifier to the Microsoft ecosystem: classify
Outlook messages with the Jev decision model (TypeSafe, reached through OpenRouter),
apply one Outlook category per message, optionally archive high-confidence disposable
mail, under a strict run budget and with resumable, checkpointed processing.

## 2. Stack

| Concern | Choice | Why |
| --- | --- | --- |
| Language / runtime | C# 13, .NET 10 (`net10.0`), console app | Microsoft-platform language; a CLI is the natural equivalent of a paste-and-run Apps Script |
| Mailbox API | Microsoft Graph v1.0 REST via `HttpClient` | Full control of `$batch`, `Prefer` headers and OData filters; no SDK version drift |
| Sign-in | `Azure.Identity` (`InteractiveBrowserCredential`, `DeviceCodeCredential`) | Delegated permissions, persisted MSAL token cache, silent renewal via `AuthenticationRecord` |
| Model | Jev `~typesafe/jev-latest` via OpenRouter Decisions (`/api/alpha/decisions`); optional direct TypeSafe (`/v1/systemone`) | Same request/response contract; OpenRouter adds `usage.cost` |
| State | JSON files in `~/.jevoutlook/` (0600) | Counterpart of Apps Script User Properties |
| Tests | xUnit | Pure logic: rules, payload, contract parsing, body compaction, filters, options |

## 3. Component map

```
Cli/Program.cs ──► Triage/TriageEngine ──► Graph/GraphMailClient ──► Microsoft Graph
      │                    │                       ▲
      │                    ├──► Jev/JevClient ─────┼──► OpenRouter / TypeSafe
      │                    │                       │
      │                    └──► Storage/JobStore, RuleStore      Graph/GraphAuth (Azure.Identity)
      └──► Storage/AppConfig, Rules/Playbooks, Rules/RuleValidator
```

| Module | Responsibility |
| --- | --- |
| `AppConstants` | All tunables (concurrency windows, retry delays, thresholds, marker name, defaults) |
| `Rules/*` | `LabelRule` record, 7 playbooks (identical to jevMail), validation (≤12 rules, no `,`/`;`, no reserved marker, id regeneration) |
| `Storage/*` | Paths, atomic JSON writes with 0600, `AppConfig`, `RuleStore` (saved rules + frozen job rules), `JobStore` |
| `Graph/GraphAuth` | Credential construction, interactive sign-in, record persistence, token acquisition |
| `Graph/GraphMailClient` | Cursor listing, `$batch` reads, master categories, category PATCH, archive move, throttling handling |
| `Graph/MessageMetadata` | Normalises a Graph message into the flat `email` object sent to Jev; HTML→text; head/tail compaction |
| `Jev/JevPayloadBuilder` | Builds the Choice question (options `L0..Ln`), size guard (29 000 bytes), cost estimate |
| `Jev/JevClient` | Sends one Decisions request; strict contract validation; cost accounting source |
| `Triage/TriageEngine` | Session lifecycle, batches, two waves, adaptive concurrency, budget, circuit breaker, grouped writes |
| `Cli/*` | Commands, argument parsing, terminal rendering |

## 4. Gmail → Outlook mapping

| Concept | Gmail (jevMail) | Outlook (jevOutlook) |
| --- | --- | --- |
| Label | Gmail label (created via `labels.create`) | Category: `POST /me/outlook/masterCategories` `{displayName, color: presetN}` then `PATCH /me/messages/{id}` `{categories:[…]}` (full replacement → existing categories are merged, never dropped) |
| Archive | `messages.batchModify` remove `INBOX` | `POST /me/messages/{id}/move` `{destinationId:"archive"}` |
| Marker | label `jev-triaged` | category `jev-triaged` (colour `none`) |
| Scope | `in:inbox` / `-in:trash -in:spam` | `/me/mailFolders/inbox/messages` / `/me/messages` with client-side exclusion of Junk, Deleted Items, Drafts (folder ids resolved once) |
| Unread | `is:unread` | `isRead eq false` |
| Skip processed | `-label:jev-triaged` in the query | client-side: skip refs whose `categories` contain the marker (Graph has no reliable `not categories/any`) |
| Pagination | `pageToken` (dry-run) / query shrinks (live) | `receivedDateTime` cursor, see §5 |
| Metadata read | `messages.get format=metadata` + selected headers, HTTP multipart batch | `$batch` of `GET /me/messages/{id}?$select=…,internetMessageHeaders` (20 per batch) |
| Full read | `messages.get format=full` + MIME walk | `$batch` of `GET /me/messages/{id}?$select=id,body` with `Prefer: outlook.body-content-type="text"` |
| Estimate | `resultSizeEstimate` | Inbox: `unreadItemCount` / `totalItemCount`; all: `$count=true` (may be unavailable → unknown) |
| Rate limits | 429 + Retry-After, window 25→50 | 429/503 + Retry-After; ≤2 `$batch` calls in flight (`SemaphoreSlim`, permit held until the body is read — every inner request counts against the ~4 concurrent per app × mailbox), window 20→40 messages |

## 5. Listing cursor

Graph requires `$orderby` properties to appear first in `$filter`. The engine therefore lists
newest-first with:

```
$filter=receivedDateTime le {cursor}[ and isRead eq false]&$orderby=receivedDateTime desc&$top=N
```

- The cursor only advances over messages that were **examined** (collected or skipped); a
  page can therefore contain more candidates than one batch needs without losing any.
- Ids sharing the cursor timestamp are kept in `CursorBoundaryIds` and excluded client-side,
  so ties are neither repeated nor lost. If a full page consists only of boundary ids (more
  identical timestamps than one page holds), the next query switches to `lt` (exclusive):
  nothing in between is lost or duplicated; only further messages with that exact timestamp
  would be left out. Collected ids are de-duplicated defensively.
- Because the cursor is time-based, archiving messages out of the Inbox mid-run never shifts
  pages, and live and preview runs use the same mechanism.
- Messages received after the session started (cursor = start + 5 min) are not part of it.

## 6. Processing algorithm (one batch)

1. **Fill pending** (≤ 50) from the cursor; complete when exhausted or the limit is reached.
2. **Metadata read** (adaptive `$batch` window). Items already carrying the marker without a
   final decision are dropped (processed elsewhere). 404 → skipped safely.
3. **Wave 1 — metadata**: one independent Choice request per pending item without a result.
   Successful answers are checkpointed before failures are handled.
4. **Decide**: confidence ≥ metadata threshold (and, in archive mode for an archive-eligible
   category, ≥ archive threshold) → final; otherwise the item goes to wave 2.
5. **Full read** + text extraction. Empty body → safe `review` fallback (non-archive) when
   such a rule exists, else skipped.
6. **Wave 2 — full body**: independent requests; answers checkpointed.
7. **Ready prefix**: leading pending items with a final decision are written in two grouped
   idempotent passes — the current `categories` of each message are **re-read in the same
   moment** (a PATCH replaces the whole collection, and the batch snapshot may be minutes old),
   then `PATCH categories` (current ∪ {category, marker}), then `move` for archive decisions.
   A retryable Graph failure pauses the session with every decision saved. On replay, a pending
   item that already has a final decision but answers 404 (its id changed because the archive
   move succeeded before the checkpoint) is counted as processed, not skipped.
8. **Counters** (processed, metadata-only, full, archived, per-category), results table,
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
| Provider 401/402/403/4xx | Session error / pause, no Outlook change |
| Provider 429/5xx after retry | Pause `provider-temporary`, resumable. A 429 is not billed (reservation released); timeouts/5xx keep the conservative estimate |
| Invalid model response ×1 | Retry once; then skip message, `ModelResponseSkips++` |
| 3 consecutive invalid responses | Pause `jev-response-circuit-breaker` |
| Graph 429/5xx after retry | Pause `graph-temporary` (reads) or after grouped write (`graph-temporary`) |
| Graph 401/403 | Error, ask to sign in again |
| Message 404 | Skipped safely (reads); treated as done (writes); counted as processed when the item already carries a final decision (replay after an archive move) |
| 40 listing pages without a new candidate | Pause `scan-guard` (everything already carries the marker); `continue` keeps scanning |
| Ctrl+C during a wave | Reservations of unanswered requests are released; they are re-dispatched on resume |
| Budget would be exceeded | Status `budget`; `continue --max-spend` to raise |
| Ctrl+C | Pause `user-stop`; `continue` resumes (and re-arms the circuit breaker) |

## 7. Persistence

| File | Content |
| --- | --- |
| `config.json` | client id, tenant, provider, endpoint/model overrides, optional API key, device-code flag |
| `auth-record.json` | MSAL `AuthenticationRecord` (no secrets; tokens live in the OS cache) |
| `rules.json` | saved rules |
| `job.json` | current session: options, cursor, counters, spend, pending items with checkpointed decisions |
| `job-rules.json` | rules frozen for the current session |

All files are written atomically (`.tmp` + move) with mode 0600; the directory is 0700.

## 8. Security notes

- Delegated permissions only (`Mail.ReadWrite`, `MailboxSettings.ReadWrite`, `User.Read`);
  the app never sends mail.
- Provider error bodies are never echoed (they may contain request data).
- Email content is passed to the model as untrusted data; the instructions tell the model never
  to follow instructions found in the email.
- The API key is read from the environment first; storing it on disk is opt-in.

## 9. Verification status (2026-09-22)

- `dotnet build`: 0 warnings, 0 errors. `dotnet test`: 48 tests pass.
- Independent code-review pass (separate agent) + documentation verification of Graph/Azure.Identity/TypeSafe
  facts; Major findings fixed (replay 404 counting, scan guard, pre-write category re-read, budget
  reservation leak on cancel, transient Graph failures pausing instead of erroring).
- Live check: OpenRouter error path exercised with an invalid key (HTTP 401 → clean message).
- Not yet exercised against a real mailbox (requires an Entra app registration): listing filter,
  `$batch` shapes and category/move writes were validated against the Graph documentation only.
