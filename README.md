# schedule_gen_from_email

A Vercel serverless function that takes the contents of an email and uses OpenAI's
`gpt-4o-mini` (via the official `openai` SDK with strict `json_schema` response
format) to extract every calendar event mentioned — date, time, IANA timezone,
event name, and a short description — and returns them as JSON.

## Project layout

```
.
├── api/
│   └── extract-event.ts     # POST /api/extract-event handler
├── lib/
│   ├── openai.ts            # OpenAI call + JSON schema + prompt
│   └── types.ts             # Event / request / response types
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

2. Create `.env.local` for local dev (or set the env var however you prefer):

   ```bash
   cp .env.example .env.local
   # then edit .env.local and set OPENAI_API_KEY
   ```

3. Run locally:

   ```bash
   npx vercel dev
   ```

   The function is available at `http://localhost:3000/api/extract-event`.

## Deploy

1. Set the secret in your Vercel project:

   ```bash
   vercel env add OPENAI_API_KEY production
   ```

2. Deploy:

   ```bash
   vercel --prod
   ```

## Endpoint

### `POST /api/extract-event`

**Request**

```http
POST /api/extract-event
Content-Type: application/json

{ "email": "<the full email body as a string>" }
```

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

| Status | When                                                |
|-------:|-----------------------------------------------------|
| 400    | Body missing or `email` is not a non-empty string  |
| 405    | Non-`POST` method                                   |
| 500    | Missing `OPENAI_API_KEY` or upstream OpenAI failure |

All errors are returned as `{ "error": "<message>" }`.

## Sample calls

### Single event

```bash
curl -X POST https://<your-deployment>.vercel.app/api/extract-event \
  -H "Content-Type: application/json" \
  -d '{
    "email": "Reminder: Career Fair on Friday Oct 3, 2025 from 1:00 PM to 5:00 PM at the Illini Union. Bring resumes."
  }'
```

Expected response:

```json
{
  "events": [
    {
      "date": "2025-10-03",
      "time": "13:00",
      "timezone": "America/Chicago",
      "event_name": "Career Fair",
      "description": "Career fair at the Illini Union, 1pm to 5pm. Bring resumes."
    }
  ]
}
```

### Multiple events

```bash
curl -X POST https://<your-deployment>.vercel.app/api/extract-event \
  -H "Content-Type: application/json" \
  -d '{
    "email": "Hackathon kicks off July 10 1:00 PM to 5:00 PM and continues July 11 1:00 PM to 3:00 PM in Siebel Center."
  }'
```

Expected response:

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

## Implementation notes

- The model is forced into a strict JSON schema via OpenAI's
  `response_format: { type: "json_schema", ... }`, so the function does not need
  to validate or coerce the shape — the SDK guarantees it matches `Event[]`.
- The system prompt injects today's date so relative phrases like "tomorrow" or
  "next Tuesday" resolve correctly.
- `temperature` is set to `0` for deterministic extraction.
- The OpenAI key is read from `process.env.OPENAI_API_KEY` and never logged or
  echoed back to the caller.

## Type-check

```bash
npm run typecheck
```

## Notes

- `vercel.json` is intentionally empty. Vercel auto-detects `api/*.ts` as a
  Node.js serverless function based on file path; specifying `runtime`
  explicitly can fail with "Function Runtimes must have a valid version" on
  some CLI versions. If you need to pin a runtime, add it back to
  `vercel.json`, e.g. `{ "functions": { "api/*.ts": { "runtime": "nodejs20.x" } } }`.
