# Authentication architecture (Microsoft Graph via the Office Dialog API)

> Technical record of the calendar-creation auth flow, and of the two rounds of
> work that produced it: an initial `window.open` + PKCE fallback that could not
> work on Outlook for Mac, and the Office-Dialog rewrite that replaced it.

| Git ref | Title |
|---|---|
| `82feb97` | **feat(auth)** — initial MSAL.js popup fallback |
| `29658c2` | **fix(msal)** — drop broken import of `./oauth-config.js` |
| `8f6fa5d` | **fix(auth)** — persist tokens in `localStorage` + add popup-blocked fallback |
| `4106072` | **fix(auth)** — surface popup-blocked fallback on every event row |
| *(current)* | **fix(auth)** — replace `window.open` with `Office.context.ui.displayDialogAsync` |

## 1. Background — why this was needed

The Outlook add-in creates calendar events via `POST https://graph.microsoft.com/v1.0/me/events`, which requires a Bearer token with `Calendars.ReadWrite`. It also reads the source email's `webLink` via `GET /v1.0/me/messages/{restId}?$select=webLink`, which requires `Mail.ReadBasic` — metadata only, no bodies or attachments. Microsoft's recommended path for an Office Add-in is:

```js
Office.auth.getAccessToken({ forMSGraphAccess: true })
```

This relies on **Office.js** brokering the SSO handshake internally. It works on Outlook on the web and on Outlook desktop (Windows), but on **Outlook for Mac** the call **silently never invokes its callback**. The function-shape check passes (`typeof Office.auth.getAccessToken === "function"`), but the implementation is a no-op: no popup, no error, no token — just a hang.

> This is not a "legacy Outlook" problem. It reproduces on current Microsoft 365 builds (confirmed on **16.113.1**). An earlier version of this document blamed legacy Mac Outlook and told the user to upgrade; that diagnosis was wrong, and the user-agent sniff in `app.js` that produced the warning has been removed.

`mailbox.getCallbackTokenAsync({ isRest: true })` and the older `Office.context.auth.getAccessToken` are also dead ends on Outlook for Mac.

**Goal:** make event creation work on every Outlook host without depending on the broken Office SSO path, while keeping the fast path intact on hosts where it works.

## 2. Architecture — control flow

```
user clicks "Create N events in calendar"
                │
                ▼
createGraphEvent(ev)
   │
   └── getGraphToken(scopes)
          │
          ├─[FAST LANE]────────────────────────────────────────────────┐
          │ Office.auth.getAccessToken({forMSGraphAccess:true})        │
          │   • 4 s timeout (FAST_LANE_TIMEOUT_MS)                     │
          │   • success? use the token                                 │
          │   • timeout/fail? mark addCalEvent.fastLaneBroken and      │
          │     skip this lane entirely on every later call            │
          └────────────────────────────────────────────────────────────┘
          │
          └─[OFFICE DIALOG]───────────────────────────────────────────┐
                                                                      │
                          ┌───────────────────────────────────────────┘
                          ▼
                  msalLogin(scopes) {
                    cached access_token still valid?  → return it
                    refresh_token present?            → redeem at Entra
                    else {
                      GET /api/auth-config → client_id, authority,
                                             authorization_url, token_url,
                                             redirect_uri, scopes
                      PKCE: 64-byte verifier, S256 challenge, random state
                        ↳ both held in module variables, never in storage
                      build the /authorize URL
                      Office.context.ui.displayDialogAsync(
                        /outlook-addin/auth-start.html?url=<authorize url>)
                        ↳ auth-start.html redirects to login.microsoftonline.com
                        ↳ consent renders inside Outlook's own dialog window
                        ↳ Entra redirects to auth-callback.html (same origin)
                        ↳ auth-callback.html calls Office.context.ui.messageParent
                      DialogMessageReceived → { ok, code, state }
                      verify state, close the dialog
                      POST directly to Entra's /token from the pane
                      → { access_token, refresh_token, expires_in }
                      cache tokens in localStorage
                      return access_token
                    }
                  }
```

Two properties make this host-agnostic where the old flow was not:

1. **The dialog is owned by Office, not by the browser.** It cannot be popup-blocked, and it never opens in the user's default browser. The previous `window.open` design failed precisely here: on a Mac whose default browser is Firefox, consent completed in Firefox while the add-in ran in Outlook's WebView, and the two share no `localStorage`, no `window.opener`, and no way to talk.
2. **The code returns in-process via `messageParent`.** No storage hand-off between windows is required at all, which is why the PKCE verifier and state are now ordinary variables.

## 3. File-by-file map

### 3.1 Auth files

| Path | Purpose |
|------|---------|
| `lib/oauth-config.ts` | Shared server-side auth constants. Reads `OAUTH_AUTHORITY` (defaults to `https://login.microsoftonline.com/common`). Exports `GRAPH_DEFAULT_SCOPES`, `REDIRECT_URI`, `AUTHORITY`, `AUTHZ_URL`, `TOKEN_URL`. |
| `api/auth-config.ts` | Public-config endpoint (GET). Returns `{client_id, authority, authorization_url, token_url, redirect_uri, scopes}`. The pane fetches this instead of having `client_id` compiled into `msal.js`. |
| `public/outlook-addin/auth-start.html` | Same-origin shim. `displayDialogAsync` requires its initial URL to be on the add-in's own domain, so the pane passes the Entra `/authorize` URL as `?url=` and this page redirects to it. Refuses any target that is not `https://login.microsoftonline.com/`. |
| `public/outlook-addin/auth-callback.html` | Entra's redirect target. Loads `office.js`, reads `code`/`error`/`state` from the query string, and hands `{ok, code, state}` to the pane with `Office.context.ui.messageParent`. Does **not** validate `state` itself — the pane holds the expected value and compares. |
| `public/outlook-addin/msal.js` | The whole PKCE flow: `loadConfig` → cached token → refresh → Office dialog → direct token exchange. Pure browser code, no imports, no dependencies. Exports `msalLogin(scopes, opts)` and `invalidateAccessToken()`. |

### 3.2 Consumers

| Path | Role |
|------|------|
| `public/outlook-addin/app.js` | `getGraphToken(scopes, opts)` runs the fast lane then `msalLogin`. `createGraphEvent` calls `invalidateAccessToken()` before its single 401 retry. Every stage is logged to the console. |
| `public/outlook-addin/index.html` | Hosts the task pane. |

## 4. How the popup design failed, and what replaced it

### 4.1 The original design (`82feb97` … `4106072`)

`window.open()` on the Entra `/authorize` URL, with the auth code returned by `postMessage` from `auth-callback.html` to `window.opener`. Four commits of patching followed: an inline `REDIRECT_URI` after a broken import, a move from `sessionStorage` to `localStorage` for tokens, a "popup blocked" fallback button, and finally a `localStorage` polling loop meant to cover the case where `window.opener` is `null`.

### 4.2 Why none of it could work

Three defects, each sufficient on its own:

1. **Cross-browser gap (fatal, unfixable client-side).** The add-in runs in Outlook's WebView; `window.open` escaping that WebView opens the user's default browser. When that browser is Firefox, consent completes in a process that shares no `localStorage` and no `window.opener` with the pane. The auth code is simply unreachable. §12 of the previous revision of this document acknowledged the constraint and then depended on it anyway.
2. **The polling bridge was never wired up.** `4106072` claimed `auth-callback.html` writes the result to `localStorage`; it wrote to **`sessionStorage`**, while `msal.js` polled **`localStorage`**. The same split silently disabled the CSRF `state` check (pane wrote local, callback read session, so `expectedState` was always `""`) and the silent-refresh path (`tryRefresh` read the refresh token from session, `storeTokens` wrote it to local — `tryRefresh` had never once succeeded).
3. **The token exchange would have been rejected anyway.** The redirect URI is registered as a Single-page application, and the old `api/exchange-token.ts` redeemed the code from the server with no `Origin` header. Entra answers that with `AADSTS9002327`. This stayed invisible only because no real code ever reached the endpoint. That endpoint has since been deleted — the exchange belongs in the browser as long as the registration stays SPA (§9).

### 4.3 The current design

`Office.context.ui.displayDialogAsync` opens Microsoft's consent screen in a window Outlook owns. It cannot be popup-blocked and never reaches the default browser, so the flow stays in one process from start to finish and `messageParent` delivers the code directly to the pane.

That removes the entire class of bug above: no `postMessage` across windows, no storage polling, no cross-browser hand-off. The PKCE verifier and `state` are module-local variables in `msal.js`, because nothing has to survive a page load any more.

The token exchange moved into the pane (`redeemAtEntra`), where the browser supplies the `Origin` header the SPA registration requires. Both the authorization-code and refresh-token grants go through it.

## 5. Storage layout

Everything the client persists is in `localStorage` (per-origin, survives pane recreation, Outlook restart, and reboot). Nothing uses `sessionStorage` any more — the split between the two backends was the source of three separate bugs.

| Key | Lifetime | Purpose |
|------|----------|---------|
| `addCalEvent.accessToken` | ~1 h | cached Graph token |
| `addCalEvent.refreshToken` | until revoked or rejected | silent refresh; cleared automatically when Entra rejects it |
| `addCalEvent.accessTokenExpiresAt` | ~1 h | epoch ms, checked before the access token is used (60 s safety margin) |
| `addCalEvent.oauth.configCache` | 5 min | `/api/auth-config` response |
| `addCalEvent.fastLaneBroken` | until cleared | set once `Office.auth.getAccessToken` has failed on this host, so later calls skip the 4 s timeout |
| `addCalEvent.tokenScopes` | until the scope set changes | fingerprint of the scopes the cached tokens were granted for. When it no longer matches the requested set, the cached access **and** refresh tokens are dropped so a newly added scope takes effect on the next sign-in rather than whenever the old token happens to expire. |

The PKCE verifier and `state` are deliberately **not** in this table. They live in `msal.js` module variables for the duration of a single `msalLogin` call.

## 6. Error model

`msalLogin` throws `Error` objects carrying a `code` property, so callers can distinguish a cancellation from a real failure:

| `code` | Meaning |
|--------|---------|
| `dialog_unavailable` | `Office.context.ui.displayDialogAsync` is missing in this host — sign-in cannot be shown at all |
| `dialog_open_failed` | Office refused to open the dialog (the Office error code is in the message; `12007` means a dialog is already open) |
| `dialog_cancelled` | the user closed the dialog (`DialogEventReceived` code `12006`) |
| `dialog_event` | the dialog closed for some other Office-reported reason |
| `dialog_timeout` | no message arrived within 120 s |
| `dialog_bad_message` | the callback sent something that was not JSON |
| `state_mismatch` | the returned `state` did not match the one this pane generated — the code is refused |
| `missing_code` / `no_access_token` | the callback or the token endpoint returned nothing usable |
| `auth_config_unavailable` / `auth_config_invalid` | `/api/auth-config` is missing or misconfigured |
| `no_webcrypto` / `insecure_context` | the WebView cannot do PKCE |
| *(Entra's own code)* | e.g. `invalid_grant`, `AADSTS…` — passed through from the token endpoint |

The old `popup_blocked` code and everything built on it (the global "Open sign-in here" button, the per-row fallback button, the re-throw out of `createEvents`) are gone: an Office dialog cannot be popup-blocked.

## 7. Timeouts and waits

| Where | Wait | Why |
|------|------|-----|
| Fast lane (`Office.auth.getAccessToken`) | 4 s (`FAST_LANE_TIMEOUT_MS`) | Long enough for a real SSO prompt on Windows desktop or Outlook on the web; short enough that Mac's hang is recognised quickly. Paid **once** per host — the failure is then remembered in `addCalEvent.fastLaneBroken` and the lane is skipped. |
| Office dialog → `messageParent` | 120 s (`DIALOG_TIMEOUT_MS`) | Consent can take a minute if the user walks away or has to do MFA |
| Access-token expiry margin | 60 s | Refresh slightly early rather than hand Graph a token about to expire |
| `/api/auth-config` cache | 5 min | Avoids a round-trip on every create |

## 8. Endpoints summary

| Endpoint | Method | Purpose | Auth |
|---------|--------|---------|------|
| `/api/extract-event` | POST | LLM extraction of events from email body | none (open) |
| `/api/manifest.xml` | GET (rewrite) | serves `/outlook-addin/manifest.xml` | none |
| `/api/auth-config` | GET | returns `{client_id, authority, authorization_url, token_url, redirect_uri, scopes}` | none (public) |

The token exchange itself goes **straight from the pane to `https://login.microsoftonline.com/{authority}/oauth2/v2.0/token`**, not through this deployment.

`AZURE_CLIENT_ID` is read from `process.env` by `api/auth-config.ts` (and by the build script for the manifest). It is a public identifier, not a secret, and the pane receives it as a JSON response rather than having it compiled in.

## 9. Azure setup (user-side, ~3 minutes)

1. **App registration → Authentication → Add a platform → Single-page application** → add the redirect URI **exactly**:
   ```
   https://schedule-gen-from-email.vercel.app/outlook-addin/auth-callback.html
   ```
   Off-by-one (case, trailing slash, `http` vs `https`) is the single most common cause of `AADSTS50011`.

   > **The platform type is load-bearing.** "Single-page application" makes Entra enforce cross-origin token redemption: the `/token` call must carry an `Origin` header, which only a browser can supply. That is why `msal.js` redeems the code itself, and why the server-side `api/exchange-token.ts` was deleted rather than kept around. Re-introducing a server-side exchange without first moving this redirect URI to the **Mobile and desktop applications** platform (and enabling *Allow public client flows*) will fail with:
   > ```
   > AADSTS9002327: Tokens issued for the 'Single-Page Application' client-type
   > may only be redeemed via cross-origin requests.
   > ```

2. **API permissions → Microsoft Graph → Delegated**: `User.Read`, `Calendars.ReadWrite`, `Mail.ReadBasic`, `openid`, `offline_access`.

3. **Grant admin consent for the tenant** so end-users don't see a permission rationale screen.

4. **Supported account types** must be **multi-tenant** — `/common` and `/organizations` refuse single-tenant apps with `AADSTS50194`.

5. (Optional) **Expose an API** for the Office-SSO fast lane, only if you want `WebApplicationInfo` SSO to keep working on Windows desktop and Outlook on the web. Application ID URI `api://schedule-gen-from-email.vercel.app/<client-id>` with an `access_as_user` scope. **Not** required for the dialog flow.

## 10. Vercel env vars

| Name | Required | Default | Used by |
|------|----------|---------|---------|
| `AZURE_CLIENT_ID` | yes | — | `scripts/build-manifest.mjs`, `api/auth-config.ts` |
| `OPENAI_API_KEY` | yes (production) | — | `api/extract-event.ts` (OpenAI provider) |
| `MINIMAX_API_KEY` | yes (production, alt provider) | — | `api/extract-event.ts` (MiniMax provider) |
| `VERCEL_API_FUNCTION_BUNDLING` | yes | — | build bundler |
| `OAUTH_AUTHORITY` | no | `https://login.microsoftonline.com/common` | `lib/oauth-config.ts` |
| `MINIMAX_BASE_URL` | no | `https://api.minimax.io/v1` | `lib/minimax.ts` |

`AZURE_CLIENT_ID` must be **Config** (plain) — Vercel hides Secrets from the build step, but the build script generates the manifest by substituting it. Currently set to the user's real GUID (`aa709c01-def6-4036-9852-77bc35ae7002`) which is a public identifier per Microsoft.

---

## 11. What the dialog rewrite did **not** change

- **Manifest XML** — untouched *by the auth work*, so the dialog rewrite alone needed no reinstall. (A later change added `<SupportsPinning>` and bumped the version to 1.4.0.0, which does require reinstalling.) `WebApplicationInfo` still references the same Application ID URI, and the Office-SSO fast lane keeps working on hosts where it works.
- **`/api/extract-event`** and the whole LLM extraction path.
- **The Graph call** — still `POST /v1.0/me/events` with the same body; only the token acquisition changed.
- **`mapToGraphFields`** event-schema mapping.
- **UI layout** was left alone by the dialog rewrite. It was reworked separately afterwards: the debug panel was removed, the header rebuilt, and the palette moved onto Fluent tokens.

## 12. Known edge cases and how they're handled

| Case | Behaviour |
|------|-----------|
| Access token expired, refresh token still good | `msalLogin` checks `accessTokenExpiresAt` first, then redeems the refresh token against Entra. Silent — no dialog. |
| Refresh token rejected (revoked consent, expired) | `tryRefresh` catches the failure and **deletes** the stored refresh token, so the next attempt goes straight to the dialog instead of retrying a dead grant. |
| No refresh token at all | Straight to the dialog. |
| User closes the dialog mid-flow | `DialogEventReceived` fires with `12006`; `msalLogin` rejects with `dialog_cancelled` and a plain "sign-in was cancelled" message. No 120 s hang. |
| A dialog is already open | `displayDialogAsync` fails with Office code `12007`, surfaced as `dialog_open_failed`. |
| Graph returns 401 on an apparently valid token | `createGraphEvent` calls `invalidateAccessToken()` — which really does clear the cached token now, unlike the old code that only logged that it had — and retries once with the fast lane skipped. The refresh token is kept, so the retry usually renews silently rather than re-prompting. |
| Office SSO hangs (every Outlook for Mac build) | 4 s timeout on the first attempt only; `addCalEvent.fastLaneBroken` is then set and every later create skips straight to the dialog. |
| User clears `localStorage` | Tokens are lost; the next create re-prompts. Normal. |
| `auth-start.html` opened directly, or with a non-Microsoft `?url=` | Refuses to redirect and says so. It will only bounce to `https://login.microsoftonline.com/`. |

## 13. Diagnostics

The in-pane debug panel is gone. It existed to diagnose the Office SSO hang
from a host with no DevTools; that is solved, and the panel was taking up most
of the task pane. Everything it logged now goes to the console with an
`[addCalEvent:<level>]` prefix.

A healthy Mac run reads:

```
[addCalEvent:info] Office.onReady host=Outlook platform=...
[addCalEvent:info] auth.getAccessToken available: true (known broken on this host — skipping it)
[addCalEvent:info] fast lane skipped (Office SSO already known to fail on this host)
[addCalEvent:info] auth: opening the Office dialog sign-in flow
[addCalEvent:info] msal: auth-dialog-opening
[addCalEvent:info] msal: auth-dialog-opened
[addCalEvent:info] msal: auth-exchanging-code
[addCalEvent:info] msal: auth-token-cached
[addCalEvent:info] POST https://graph.microsoft.com/v1.0/me/events
[addCalEvent:info] Graph response status=201
```

On the very first run the fast lane is still attempted, adding
`fast lane timed out after 4000ms`. Later creates show `msal: auth-cache-hit`
instead of the dialog stages.

## 14. Verifying a deployment

```bash
npm run typecheck && npm test
curl -s https://schedule-gen-from-email.vercel.app/api/auth-config | jq
#   → must include token_url
curl -sI https://schedule-gen-from-email.vercel.app/outlook-addin/auth-start.html
#   → 200
```

Outlook for Mac caches add-in web resources aggressively. After deploying, clear it or the pane will keep running the old `msal.js`:

```bash
rm -rf ~/Library/Containers/com.microsoft.Outlook/Data/Library/Caches/com.microsoft.Outlook/
```

then restart Outlook.

Then click **Create** on an extracted event: an Outlook-owned dialog should open (never the default browser), show the Microsoft account picker, close itself, and log `msal: auth-token-cached`.

## 15. Future work

- **iOS / Android Outlook** — the Dialog API is supported there too; untested.
- **Silent SSO on Windows and web** — the fast lane still covers those; the `addCalEvent.fastLaneBroken` memo is per-origin, so a user who moves between hosts on the same browser profile may skip a lane that would have worked. Clearing the key re-enables it.
- **BFF pattern** — moving token handling server-side would keep the access token out of the client entirely, but requires re-registering the redirect URI away from the SPA platform (§9).
- **CSP** — if the host page ever gets a strict CSP, the minimum is `connect-src https://login.microsoftonline.com https://graph.microsoft.com`.
