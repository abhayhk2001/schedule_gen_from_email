# schedule_gen_from_email

A Vercel serverless function that takes the contents of an email and uses an
LLM to extract every calendar event mentioned and returns it as JSON suitable
for creating Outlook events (start/end dates, start/end times, whole-day
flag, IANA timezone, event name, description). The caller chooses the LLM
per request via a `model` field.

Supported models:

| `model`        | Provider | Key env var       |
|----------------|----------|-------------------|
| `gpt-4o-mini`  | OpenAI   | `OPENAI_API_KEY`  |
| `MiniMax-M3`   | MiniMax  | `MINIMAX_API_KEY` |

Both providers are reached through the official `openai` SDK with strict
`json_schema` response format. MiniMax is reached via its OpenAI-compatible
endpoint.

## Project layout

```
.
├── api/
│   └── extract-event.ts     # POST /api/extract-event handler
├── lib/
│   ├── router.ts            # picks provider based on requested model
│   ├── openai.ts            # OpenAI extraction (gpt-4o-mini)
│   ├── minimax.ts           # MiniMax extraction (MiniMax-M3)
│   ├── extract.ts           # shared chat-completion + JSON-parse helper
│   ├── errors.ts            # ConfigError class
│   ├── manifest-template.ts # Outlook manifest XML template (__AZURE_CLIENT_ID__ placeholder)
│   ├── prompt.ts            # shared system prompt + JSON schema + model list
│   └── types.ts             # Event / request / response types
├── tests/
│   ├── handler.test.ts      # 12 input-validation cases + error-wrapping tests
│   ├── router.test.ts       # dispatch tests
│   ├── providers.test.ts    # missing-key ConfigError tests
│   └── extract.test.ts      # runExtraction helper tests
├── public/
│   ├── index.html           # landing page at the deployment root
│   ├── site.css             # styles for the landing page
│   ├── site.js              # calls /api/extract-event from the landing page
│   └── outlook-addin/       # Outlook Add-in (static files served by Vercel)
│       ├── index.html
│       ├── app.js
│       ├── app.css
│       ├── theme.js
│       └── assets/          # 16/32/64/80/128 px icons referenced by the manifest
├── assets/
│   └── app-icon.png         # master source for the manifest icons (697x697). Regenerate sizes with `sips -z`.
├── run-tests.sh             # live regression battery (8 multi-day scenarios)
├── vitest.config.ts
├── package.json
├── tsconfig.json
├── vercel.json
├── .env.example
└── README.md
```

## Setup

1. Install dependencies:

   ```bash
   npm install
   ```

2. Create `.env.local` for local dev:

   ```bash
   cp .env.example .env.local
   # then set whichever API key(s) you need
   ```

3. Run locally:

   ```bash
   npx vercel dev
   ```

   - The website is at `http://localhost:3000/`.
   - The function is at `http://localhost:3000/api/extract-event`.
   - The add-in source is at `http://localhost:3000/outlook-addin/`.

## Deploy

Pushing to `main` deploys: the Vercel GitHub integration builds and
promotes to production automatically. The steps below cover deploying by
hand, and the one-time project setup.

The Vercel CLI is a devDependency, so `npm run deploy` and `npm run dev`
resolve it from `node_modules/.bin` — no global install needed. Sign in
once before the first manual deploy:

```bash
npx vercel login
```

1. Set the secret(s) in your Vercel project:

   ```bash
   vercel env add OPENAI_API_KEY production     # for gpt-4o-mini requests
   vercel env add MINIMAX_API_KEY production    # for MiniMax-M3 requests
   ```

2. (Required for this codebase) Set the bundling env var as a **plain config**
   value, not a Secret — Secrets are hidden from the build step:

   ```bash
   echo "1" | vercel env add VERCEL_API_FUNCTION_BUNDLING production --type config
   ```

3. (Optional) Override MiniMax base URL:

   ```bash
   echo "https://api.minimax.io/v1" | vercel env add MINIMAX_BASE_URL production
   ```

4. Deploy by hand (typechecks and runs the tests first, then promotes to
   production). Note this ships your **working tree**, which may differ
   from what is on `main`:

   ```bash
   npm run deploy
   ```

> **Important — `package.json` has no `"type": "module"`.** Vercel's @vercel/node
> bundler generates a single `___vc_bundled_api_handler.js` that mixes the
> bundled-handler (which uses CommonJS `require`) with your code. Forcing the
> package to ESM would crash that bundle with `require is not defined in ES
> module scope`. The `VERCEL_API_FUNCTION_BUNDLING=1` config var above is also
> required — without it, Vercel tries to ship your TS files separately and
> Node ESM can't resolve the cross-file `.js` imports.

## Endpoint

### `POST /api/extract-event`

**Request**

```http
POST /api/extract-event
Content-Type: application/json

{
  "email": "<the full email body as a string>",
  "model": "gpt-4o-mini" | "MiniMax-M3"
}
```

`model` is **required**. `email` must be a non-empty string.

**Response — `200 OK`**

```json
{
  "events": [
    {
      "date": "2025-07-10",
      "time": "13:00",
      "timezone": "America/Chicago",
      "event_name": "Hackathon Day 1",
      "description": "Day 1 of the campus hackathon, 1pm to 5pm.",
      "whole_day": false,
      "end_date": null,
      "end_time": "17:00"
    },
    {
      "date": "2025-07-11",
      "time": "13:00",
      "timezone": "America/Chicago",
      "event_name": "Hackathon Day 2",
      "description": "Day 2 of the campus hackathon, 1pm to 3pm.",
      "whole_day": false,
      "end_date": null,
      "end_time": "15:00"
    }
  ]
}
```

Any field the model cannot infer is returned as `null`. If no event is
mentioned, the response is `{ "events": [] }`.

### Event fields

| Field        | Type             | Notes                                                       |
|--------------|------------------|-------------------------------------------------------------|
| `date`       | string \| null   | ISO YYYY-MM-DD. Start date.                                 |
| `time`       | string \| null   | 24h HH:MM. Start time. `null` when `whole_day=true`.        |
| `timezone`   | string \| null   | IANA name (e.g. `America/Chicago`). `null` if not inferable. |
| `event_name` | string \| null   | Short title.                                                |
| `description`| string \| null   | One-sentence summary.                                       |
| `whole_day`  | boolean          | `true` ⇒ no specific time-of-day on that date.              |
| `end_date`   | string \| null   | ISO YYYY-MM-DD end date. `null` when entry spans one day.   |
| `end_time`   | string \| null   | 24h HH:MM end time. `null` when `whole_day=true` or absent. |

### Multi-day handling

| Input pattern | Output |
|---------------|--------|
| Same-day with explicit end time (`"2pm to 3pm"`, `"9am-5pm"`) | 1 entry: `whole_day=false`, `time=<start>`, `end_time=<end>`, `end_date=null` |
| Date range, no times (`"Oct 5 to Oct 8"`, `"Oct 10-20"`) | One `whole_day=true` entry per calendar day in the range; `time=null`, `end_time=null` |
| Date range with `"each day"` / `"daily"` times (`"Oct 10-12, 9am-5pm each day"`, `"Mon Oct 5 to Wed Oct 7, daily 10am-4pm"`) | One entry per calendar day, each with `whole_day=false`, `time=<start-of-day>`, `end_time=<end-of-day>` |
| Independent events (`"Day 1: ... Day 2: ..."`) | One entry per event, each with its own `time` / `end_time` / `whole_day` |

Registration deadlines and RSVP dates are **not** emitted as events.

### Error responses

| Status | When                                                                        |
|-------:|-----------------------------------------------------------------------------|
| 400    | `email` missing/empty, or `model` is not one of the supported values         |
| 405    | Non-`POST` method                                                           |
| 500    | Missing API key for the chosen provider, or upstream LLM failure             |

All errors are returned as `{ "error": "<message>" }`.

## Sample calls

### Single event, OpenAI

```bash
curl -X POST https://<your-deployment>.vercel.app/api/extract-event \
  -H "Content-Type: application/json" \
  -d '{
    "model": "gpt-4o-mini",
    "email": "Reminder: Career Fair on Friday Oct 3, 2025 from 1:00 PM to 5:00 PM at the Illini Union. Bring resumes."
  }'
```

### Multiple events, MiniMax

```bash
curl -X POST https://<your-deployment>.vercel.app/api/extract-event \
  -H "Content-Type: application/json" \
  -d '{
    "model": "MiniMax-M3",
    "email": "Hackathon kicks off July 10 1:00 PM to 5:00 PM and continues July 11 1:00 PM to 3:00 PM in Siebel Center."
  }'
```

Expected response (identical shape regardless of provider):

```json
{
  "events": [
    {
      "date": "2025-07-10",
      "time": "13:00",
      "timezone": "America/Chicago",
      "event_name": "Hackathon Day 1",
      "description": "Day 1 of the hackathon in Siebel Center, 1pm to 5pm.",
      "whole_day": false,
      "end_date": null,
      "end_time": "17:00"
    },
    {
      "date": "2025-07-11",
      "time": "13:00",
      "timezone": "America/Chicago",
      "event_name": "Hackathon Day 2",
      "description": "Day 2 of the hackathon in Siebel Center, 1pm to 3pm.",
      "whole_day": false,
      "end_date": null,
      "end_time": "15:00"
    }
  ]
}
```

### Bad request — missing `model`

```bash
curl -i -X POST https://<your-deployment>.vercel.app/api/extract-event \
  -H "Content-Type: application/json" \
  -d '{"email":"some email"}'
# HTTP/1.1 400
# {"error":"Request body must include a \"model\" field. Supported values: gpt-4o-mini, MiniMax-M3."}
```

## Implementation notes

- `lib/router.ts` dispatches on `model`. Adding a new provider means adding a
  case there and a corresponding `lib/<provider>.ts` module.
- The shared chat-completion + JSON-parse logic lives in `lib/extract.ts`,
  which both providers call. Provider-specific differences (the MiniMax
  `thinking: disabled` flag) are passed as an `extraBody` argument.
- `lib/errors.ts` exports a `ConfigError` class. The handler distinguishes
  `ConfigError` (passes the message through — these are deployment
  configuration issues the caller should see) from any other error (logged
  server-side via `console.error` and replaced with a generic
  `"Extraction failed"` response so upstream details never leak).
- Both providers share the system prompt and JSON schema in `lib/prompt.ts`,
  so output shape is identical regardless of which model is called.
- Both use `response_format: { type: "json_schema", strict: true, ... }` so the
  SDK guarantees the response matches `Event[]` — no coercion needed.
- MiniMax additionally sends `thinking: { type: "disabled" }` via `extra_body`
  so M3 does not inject `<thinking>...</thinking>` content that would corrupt
  the JSON output.
- `temperature` is `0` for deterministic extraction.
- API keys are read from `process.env.OPENAI_API_KEY` and
  `process.env.MINIMAX_API_KEY` and never echoed back to the caller.
- The handler returns `400` if `model` is not one of the supported values
  rather than falling back silently — the caller should know which provider
  was used.

## Type-check

```bash
npm run typecheck
```

## Tests

```bash
npm test
```

28 unit tests covering the 12 validation cases from the API analysis,
provider routing, the shared extraction helper (null content, malformed JSON,
non-array events, `extra_body` propagation, `whole_day=true` handling), and
the error-wrapping behaviour (Fix 6) — `ConfigError` messages pass through;
any other error is logged server-side and replaced with a generic
`"Extraction failed"` response so upstream details never leak to the caller.

## Live regression battery

`run-tests.sh` exercises 8 multi-day scenarios against the deployed API. Use
it after any prompt change to confirm the model still follows the rules in
the [Multi-day handling](#multi-day-handling) table.

## Notes

- `vercel.json` is intentionally empty. Vercel auto-detects `api/*.ts` as a
  Node.js serverless function based on file path; specifying `runtime`
  explicitly can fail with "Function Runtimes must have a valid version" on
  some CLI versions. If you need to pin a runtime, add it back to
  `vercel.json`, e.g. `{ "functions": { "api/*.ts": { "runtime": "nodejs20.x" } } }`.
- This service is server-to-server only (Outlook Actions / Power Automate /
  arbitrary HTTP clients). It does not set CORS headers. If you ever need
  browser callers, add an `Access-Control-Allow-Origin` header in the
  handler.

## Outlook Add-in

A small Outlook Add-in lives in `public/outlook-addin/` and is served as
static files by the same Vercel project. It appears in the read pane of any
message, reads the subject/sender/body via Office.js, posts to
`/api/extract-event` (same-origin, so no CORS changes are needed), and
renders the returned events.

### Install (Outlook on the web)

1. Open Outlook on the web (https://outlook.office.com or
   https://outlook.live.com).
2. **Settings** (gear icon) → **View all Outlook settings** → **Mail** →
   **Customize actions** → **Manage add-ins** (or in newer Outlook:
   **Settings** → **Integrations** → **Add-ins**).
3. At the bottom of the add-ins list click **+ Add custom add-in** → **Add
   from URL**.
4. Paste:

   ```
https://schedule-gen-from-email.vercel.app/outlook-addin/manifest.xml
   ```

5. Click **Install**, accept the permission prompt (`ReadItem`).
6. Open any email — the **Event Extractor** button appears in the ribbon.
   Click it; the add-in pane opens, fetches events from the API, and
   displays them.

### Install (Outlook desktop, Windows)

1. **File** → **Manage Add-ins** (or **Get Add-ins** → **My add-ins**).
2. Click **+ Add a custom add-in** → **Add from File…** (the on-the-web
   flow only allows URL-based manifests in newer builds; for desktop you may
   need to download the manifest to disk first via
   `https://schedule-gen-from-email.vercel.app/outlook-addin/manifest.xml`).

### Files

| Path | Purpose |
|------|---------|
| `lib/manifest-template.ts` | Manifest XML template (single source of truth). Contains the `__AZURE_CLIENT_ID__` placeholder. |
| `scripts/build-manifest.mjs` | Pre-build step that substitutes `$AZURE_CLIENT_ID` into the template and writes `public/outlook-addin/manifest.xml`. |
| `public/outlook-addin/manifest.xml` | Generated at build time, served as a static asset at `/outlook-addin/manifest.xml`. Gitignored. |
| `public/outlook-addin/index.html` | Task-pane UI. |
| `public/outlook-addin/app.js` | Office.js + fetch logic. |
| `public/outlook-addin/app.css` | Card styling. |
| `public/outlook-addin/assets/` | PNG icons (16/32/64/80/128). Replace with your real logo. |

### How the manifest reaches users

`scripts/build-manifest.mjs` runs on every Vercel deploy via the
`buildCommand` in `vercel.json`. It reads `AZURE_CLIENT_ID` from the
project environment and writes a fully-rendered XML file into
`public/outlook-addin/manifest.xml`. The browser and Outlook both
fetch that file directly — no server function, no runtime rewrite.
Because the env var is read **at build time**, it must be set as
**Config** (not Secret) in the Vercel dashboard:

```bash
vercel env add AZURE_CLIENT_ID production --type config
# paste your Azure client GUID when prompted
```

Client IDs are public identifiers (Microsoft's docs say so
explicitly), so plaintext storage here is safe.

### Limitations / next steps

- **No auth on the API.** The add-in calls `/api/extract-event` directly;
  anyone with the URL can use the backend. Fine for single-user use; add a
  shared-secret header before exposing to others.
- **No copy-to-clipboard.** Per your direction, events are rendered but not
  copyable in this iteration.

## Creating calendar events

After the events render, each card has a checkbox + **× remove** button, and
the footer shows a green **"Create N events in calendar"** button.

### Transport: Microsoft Graph

The add-in creates calendar events with Microsoft Graph, calling
`https://graph.microsoft.com/v1.0/me/events` directly from the pane.

It gets the Bearer token from whichever of two paths works on the host:
the Office SSO fast lane (`Office.auth.getAccessToken`) where that is
implemented, otherwise a PKCE sign-in shown in an Office dialog. See
[Auth](#auth-azure-app-registration-required) and `docs/AUTH.md`.

For each non-removed event, the add-in POSTs:

```json
POST https://graph.microsoft.com/v1.0/me/events
Authorization: Bearer <Graph access token>
Content-Type: application/json

{
  "subject": "{event_name}",
  "body": { "contentType": "Text", "content": "{description}" },
  "start": { "dateTime": "{date}T{time}:00", "timeZone": "{tz}" },
  "end":   { "dateTime": "{end_date}T{end_time}:00", "timeZone": "{tz}" },
  "isAllDay": false
}
```

| API field        | `whole_day` | Graph payload                                                                                  |
|------------------|:-----------:|------------------------------------------------------------------------------------------------|
| All day          | true        | `start.dateTime: "{date}"`, `end.dateTime: "{end_date ?? date+1d}"`, `isAllDay: true`           |
| Timed            | false       | `start.dateTime: "{date}T{time}:00"`, `end.dateTime: "{end_date ?? date}T{end_time ?? time+1h}:00"`, `isAllDay: false` |
| Both (always)    | —           | `subject: "{event_name}"`, `body.contentType: "Text"`, `body.content: "{description}"`            |

### Fallback rules

- Missing `end_time` on a timed event → end = start + 1 hour.
- Missing `end_date` on a whole-day event → end = start + 1 day.
- Missing `timezone` → browser's IANA timezone via `Intl.DateTimeFormat`.

### Multi-day handling

For an email like *"Oct 10-20, 2026"* the API emits one entry per calendar
day. The add-in shows all 11 rows; remove the ones you don't want before
clicking **Create**. Each row becomes its own Outlook event.

### Auth (Azure App Registration required)

The add-in creates calendar events with Microsoft Graph. It uses
two cooperating auth paths:

| Path | When | How |
|------|------|-----|
| **Office SSO fast lane** | Hosts where `Office.auth.getAccessToken` actually returns — Outlook on the web, Outlook desktop on Windows, future Mac builds where Microsoft fixes the silent-hang bug. | `<WebApplicationInfo>` in the manifest; Office calls Entra internally, returns a Graph token to the add-in. Fast (<1 s), no user interaction after first consent. |
| **Office dialog sign-in** | Every other host — including **all current Outlook for Mac builds**, where `Office.auth.getAccessToken` silently never invokes its callback. | The add-in calls `Office.context.ui.displayDialogAsync` on a same-origin shim that redirects to `login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize`. Consent renders in a window Outlook owns — it cannot be popup-blocked and never opens the OS default browser. The callback page hands the auth code back with `Office.context.ui.messageParent`, and the pane trades it (+ S256 PKCE verifier) for a Graph access token directly at Entra. |

> An earlier design used `window.open` instead of the Office dialog. It
> could not work: the popup opened in the user's *default browser*, which
> shares no `localStorage` and no `window.opener` with Outlook's WebView,
> so the auth code could never get back. `docs/AUTH.md` §4 has the full
> post-mortem.

Both flows use a **public client** — there is no client secret. The
PKCE pattern (random 64-byte `code_verifier`, `code_challenge=S256(verifier)`)
binds the auth code to the requesting browser, so a leaked code is
useless without the verifier.

#### Azure setup

1. Register an app in **Microsoft Entra admin center** →
   *Applications → App registrations → New registration*.
   - **Name**: `Email Event Extractor` (or similar).
   - **Supported account types**: *Accounts in any organizational directory* (multi-tenant) OR *Single tenant* if you prefer. Default in this codebase is multi-tenant (`/common`).
   - **Redirect URI**: leave blank for now (Office handles that itself).
2. Note the **Application (client) ID** from the *Overview* blade —
   you'll paste it as `AZURE_CLIENT_ID`.
3. **Authentication → Add a platform → Single-page application** →
   add this redirect URI (exact):
   ```
   https://schedule-gen-from-email.vercel.app/outlook-addin/auth-callback.html
   ```
   If you'll develop locally, also add:
   ```
   http://localhost:3000/outlook-addin/auth-callback.html
   ```
   Save. (Off-by-one — case, slash, `http` vs `https` — is the single
   most common cause of `AADSTS50011`.)

   > Keep this on the **Single-page application** platform. That platform
   > makes Entra require an `Origin` header when the code is redeemed,
   > which is why the pane redeems it itself rather than posting to
   > `/api/exchange-token`. Moving the exchange back to the server without
   > first switching this to *Mobile and desktop applications* fails with
   > `AADSTS9002327`.
4. **API permissions** → *Microsoft Graph* → *Delegated permissions* →
   **Add permissions**:
   - `User.Read`
   - `Calendars.ReadWrite`
   - `openid`, `profile`, `offline_access`
   Click **Grant admin consent for &lt;tenant&gt;**.
5. **Expose an API** (only required if you keep the Office SSO fast lane and
   want V1_0 Office desktop clients to use `client_credentials` flows
   later; not needed for the Office dialog path):
   - *Set* the Application ID URI to
     `api://schedule-gen-from-email.vercel.app/<client-id>` (use the same
     `<client-id>` that you'll put in the manifest).
   - Add a scope `access_as_user` with *Admins and users* consent.

#### Env vars

```bash
# Production:
vercel env add AZURE_CLIENT_ID production   --type config
# paste your GUID when prompted

# Optional. Default authority is /common (multi-tenant).
# Override to lock a tenant or restrict to work/school:
vercel env add OAUTH_AUTHORITY production   --type config
# value e.g. https://login.microsoftonline.com/<your-tenant-id>
```

`.env.local` for local dev:
```
AZURE_CLIENT_ID=<your-guid>
# OAUTH_AUTHORITY=...
```

### Manifest wiring

`lib/manifest-template.ts` is the single source of truth. It is
rendered to `public/outlook-addin/manifest.xml` at build time by
`scripts/build-manifest.mjs` (invoked via Vercel's `buildCommand` in
`vercel.json`). The file is served as a static asset — no server
function, no runtime rewrite.

The manifest contains:

- A `<VersionOverrides xsi:type="VersionOverridesV1_1">` **nested inside**
  the V1_0 block — required by the schema; the validator at
  `npx office-addin-validator` reports `Validation: Passed`.
- `<Permissions>ReadWriteItem</Permissions>` — the Office API surface we
  use is just reading the current message; `ReadWriteMailbox` is no
  longer needed.
- `<WebApplicationInfo>` with the Azure client ID, Application ID URI,
  and the five Graph scopes listed above.

The real client ID is **never committed to the repo**. Set it as a
**plain config** env var so the build step can read it:

```bash
vercel env add AZURE_CLIENT_ID production --type config
# paste your Azure client GUID when prompted
```

If the env var is missing the build fails with a clear message — the
absence is loud, not silent. Bump the `Version` in
`lib/manifest-template.ts` when you change the manifest.

### Required permission

On first click of **Create**:

- **Hosts where Office SSO works** (Outlook on the web, Windows desktop):
  Office shows a one-time consent dialog asking the user to allow the
  add-in to *read and write calendar items through Microsoft Graph*
  using their sign-in. Click **Accept**. The token is cached; on `401`
  the add-in clears its cache and re-acquires once.
- **Outlook for Mac** (all current builds): the add-in opens a Microsoft
  sign-in dialog *inside Outlook*. Sign in once and click **Accept**.
  The dialog closes itself; the access token and a refresh token are
  cached in `localStorage`, so they survive the pane being recreated
  when you switch emails. Subsequent **Create** clicks are silent until
  the token expires, and expiry is then handled by a silent refresh.

If your tenant admin has blocked user consent for Graph, the install or
the first create will fail — ask the admin to grant admin consent for
the Graph permissions listed above (Entra → API permissions → **Grant
admin consent for &lt;tenant&gt;**).

### Error policy

Failures are reported per event. A failed row stays in the list with its
error message and a **Retry** button; successful rows show **"Created —
open your Outlook calendar to view"**. One failure does not block the
others.

### Debug log (Outlook on Mac or any environment without DevTools)

The add-in pane ships with a visible **Debug** section at the bottom.
It is hidden until a log entry is appended. Each entry is a coloured
breadcrumb: timestamped, prefixed by level, and reachable from the
pane without DevTools.

| Button | Action |
|--------|--------|
| **Test Dialog auth** | Skips the Office SSO fast lane and runs the Office dialog sign-in directly — no Graph POST. This is the path the add-in actually uses on Mac, so it is the fastest way to verify auth without extracting an email first. |
| **Test SSO** | Calls `Office.auth.getAccessToken` only — no Graph POST. Surfaces consent errors, sign-in prompts, host compatibility issues, and the 13001/13004/13005/13006/13012 codes that mean a dialog was rejected, suppressed, or never offered. |
| **Copy log** | Copies the entire log to the clipboard so you can paste it back here. |
| **Clear** | Empties the panel. |

Log contents cover: Office.js / `Office.auth.getAccessToken`
availability, scope requested, token-acquired token prefix, Graph
POST/response, every error code + message + trace. The first four
lines are emitted at startup so the panel tells you — without you
clicking anything — whether the add-in loaded in a real Office host.

Typical failures:

| Code | Meaning | Fix |
|------|---------|-----|
| 13001 | Office.js not authorized to call `getAccessToken` | Manifest's `<WebApplicationInfo>` is missing or the Application ID URI in Entra doesn't match `<Resource>` exactly |
| 13004 | User declined the consent dialog | They need to click **Accept** once; on managed devices admins may have hidden this |
| 13005 | Consent dialog suppressed (tenant policy) | Grant admin consent in Entra for the five delegated permissions |
| 13006 | No Office identity (Outlook signed out) | Sign in to Outlook first |
| 13012 | MFA / Conditional Access triggered | Dialog flashed and closed; user must complete MFA in Office or auth gets forwarded silently |

## Theming

### Add-in pane (`public/outlook-addin/`)

The pane paints **dark** by default and replaces its tokens with the
host's Office palette as soon as `Office.context.officeTheme` resolves.
On `Office.EventType.ThemeChanged` the same swap repeats, so the pane
re-themes instantly when the user switches Office themes (Office,
Office on the web, dark gray, black, high-contrast).

Source of truth: `public/outlook-addin/theme.js`. It writes the
following CSS variables to `:root`:

| Variable         | Comes from                       |
|------------------|----------------------------------|
| `--bg`           | `bodyBackgroundColor`            |
| `--fg`           | `bodyForegroundColor`            |
| `--surface`      | `controlBackgroundColor`         |
| `--surface-fg`   | `controlForegroundColor`         |
| `--border`       | `controlBorderColor`             |
| `--accent`       | `accent1`                        |
| `--accent-hover` | accent1 × 0.85 each channel      |
| `--accent-on`    | white or black by RGB luminance  |
| `--accent-soft`  | accent1 @ 18% alpha              |

The emerald `--success-fg` / `--success-fg-strong` / `--success-bg`
ramp is **not** derived from the Office accent — it is fixed at
`#047857` / `#059669` so the *Create in calendar* button and the
success pill stay visually consistent regardless of theme.

### Manual override

A sun/moon toggle in the pane header overrides Office's theme. The
choice is stored under `localStorage["addCalEvent.themeOverride"]` as
`"light"` or `"dark"`. While set:

- The pane ignores `Office.EventType.ThemeChanged`.
- The value is reconciled across multiple open panes via the
  `storage` event.
- Clearing it from DevTools (`localStorage.removeItem(...)`)
  hands control back to Office.

### Landing page (`public/`)

`public/site.css` uses `@media (prefers-color-scheme: dark)` so the
page flips when the user's OS appearance changes — no JS. The token
list mirrors the add-in's; gradient logo, code block, cards, inputs,
buttons, status pills all flow from the same variables.

| Token override lives in | Trigger |
|-------------------------|---------|
| `:root` (light)         | default                                |
| `@media (prefers-color-scheme: dark)` in `site.css` | when the user's OS reports dark |

`<meta name="color-scheme" content="light dark">` in `public/index.html`
makes the browser's address bar and native form controls follow the
active theme.
