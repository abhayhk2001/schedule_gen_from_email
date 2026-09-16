# schedule_gen_from_email

A Vercel serverless function that takes the contents of an email and uses an
LLM to extract every calendar event mentioned — date, time, IANA timezone,
event name, and a short description — and returns them as JSON. The caller
chooses the LLM per request via a `model` field.

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
│   ├── prompt.ts            # shared system prompt + JSON schema + model list
│   └── types.ts             # Event / request / response types
├── tests/
│   ├── handler.test.ts      # 12 input-validation cases + error-wrapping tests
│   ├── router.test.ts       # dispatch tests
│   ├── providers.test.ts    # missing-key ConfigError tests
│   └── extract.test.ts      # runExtraction helper tests
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

   The function is available at `http://localhost:3000/api/extract-event`.

## Deploy

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

4. Deploy:

   ```bash
   vercel --prod
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
      "description": "Day 1 of the campus hackathon, 1pm to 5pm."
    },
    {
      "date": "2025-07-11",
      "time": "13:00",
      "timezone": "America/Chicago",
      "event_name": "Hackathon Day 2",
      "description": "Day 2 of the campus hackathon, 1pm to 3pm."
    }
  ]
}
```

Any field the model cannot infer is returned as `null`. If no event is
mentioned, the response is `{ "events": [] }`.

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
      "description": "Day 1 of the hackathon in Siebel Center, 1pm to 5pm."
    },
    {
      "date": "2025-07-11",
      "time": "13:00",
      "timezone": "America/Chicago",
      "event_name": "Hackathon Day 2",
      "description": "Day 2 of the hackathon in Siebel Center, 1pm to 3pm."
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

27 unit tests covering the 12 validation cases from the API analysis,
provider routing, the shared extraction helper (null content, malformed JSON,
non-array events, `extra_body` propagation), and the error-wrapping behaviour
(Fix 6) — `ConfigError` messages pass through; any other error is logged
server-side and replaced with a generic `"Extraction failed"` response so
upstream details never leak to the caller.

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
