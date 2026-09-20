const DEFAULT_REDIRECT_URI =
  "https://schedule-gen-from-email.vercel.app/outlook-addin/auth-callback.html";

const AUTH_START_PATH = "/outlook-addin/auth-start.html";

const CONFIG_CACHE_KEY = "addCalEvent.oauth.configCache";
const CONFIG_TTL_MS = 5 * 60 * 1000;

const STORAGE_ACCESS_KEY = "addCalEvent.accessToken";
const STORAGE_REFRESH_KEY = "addCalEvent.refreshToken";
const STORAGE_EXPIRES_KEY = "addCalEvent.accessTokenExpiresAt";

const DIALOG_TIMEOUT_MS = 120_000;

function safeLocalGet(key) {
  try {
    return window.localStorage?.getItem(key) ?? null;
  } catch {
    return null;
  }
}

function safeLocalSet(key, value) {
  try {
    if (value === null || value === undefined) window.localStorage.removeItem(key);
    else window.localStorage.setItem(key, value);
  } catch {}
}

function authError(code, message, extra = {}) {
  const err = new Error(message);
  err.code = code;
  Object.assign(err, extra);
  return err;
}

async function loadConfig() {
  const cachedRaw = safeLocalGet(CONFIG_CACHE_KEY);
  if (cachedRaw) {
    try {
      const parsed = JSON.parse(cachedRaw);
      if (parsed?.client_id && parsed?.authorization_url && parsed?.scopes?.length) {
        if (Date.now() - (parsed.__fetchedAt ?? 0) < CONFIG_TTL_MS) {
          return parsed;
        }
      }
    } catch {}
  }
  const resp = await fetch("/api/auth-config", { method: "GET" });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw authError(
      "auth_config_unavailable",
      `auth_config_unavailable (${data?.error ?? resp.status}): ${data?.message ?? ""}`.trim(),
    );
  }
  if (!data?.client_id || !data?.authorization_url || !Array.isArray(data.scopes)) {
    throw authError(
      "auth_config_invalid",
      "auth_config response is missing client_id/authorization_url/scopes",
    );
  }
  safeLocalSet(CONFIG_CACHE_KEY, JSON.stringify({ ...data, __fetchedAt: Date.now() }));
  return data;
}

function tokenUrlFor(config) {
  if (config?.token_url) return config.token_url;
  const authority = (config?.authority ?? "https://login.microsoftonline.com/common").replace(
    /\/+$/,
    "",
  );
  return `${authority}/oauth2/v2.0/token`;
}

function redirectUriFor(config) {
  return config?.redirect_uri ?? DEFAULT_REDIRECT_URI;
}

function b64UrlEncode(bytes) {
  let str = "";
  for (let i = 0; i < bytes.length; i++) str += String.fromCharCode(bytes[i]);
  return btoa(str).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function makeCodeVerifier() {
  const bytes = new Uint8Array(64);
  crypto.getRandomValues(bytes);
  return b64UrlEncode(bytes);
}

async function makeCodeChallenge(verifier) {
  const data = new TextEncoder().encode(verifier);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return b64UrlEncode(new Uint8Array(digest));
}

function makeState() {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  return b64UrlEncode(bytes);
}

/**
 * Redeem a grant directly against Entra from the browser.
 *
 * The redirect URI is registered as a Single-page application, and Entra
 * refuses SPA-issued codes that are redeemed without an Origin header
 * (AADSTS9002327). That rules out a server-side exchange, so this call must
 * stay in the pane. No client secret is involved — this is a public client
 * using PKCE.
 */
async function redeemAtEntra(config, fields) {
  const params = new URLSearchParams({ client_id: config.client_id, ...fields });
  const resp = await fetch(tokenUrlFor(config), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params.toString(),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok || data?.error) {
    throw authError(
      data?.error ?? "token_exchange_failed",
      `${data?.error ?? `HTTP ${resp.status}`}: ${data?.error_description ?? resp.statusText}`,
    );
  }
  return data;
}

function storeTokens(accessToken, refreshToken, expiresInSec) {
  safeLocalSet(STORAGE_ACCESS_KEY, accessToken);
  if (refreshToken) safeLocalSet(STORAGE_REFRESH_KEY, refreshToken);
  if (expiresInSec) {
    safeLocalSet(
      STORAGE_EXPIRES_KEY,
      String(Date.now() + (Number(expiresInSec) - 60) * 1000),
    );
  }
}

function readCachedAccessToken() {
  const access = safeLocalGet(STORAGE_ACCESS_KEY);
  const expiresAt = Number(safeLocalGet(STORAGE_EXPIRES_KEY));
  if (!access || !expiresAt) return null;
  if (Date.now() >= expiresAt) return null;
  return access;
}

async function tryRefresh(config, scopes) {
  const refreshToken = safeLocalGet(STORAGE_REFRESH_KEY);
  if (!refreshToken) return null;
  try {
    const data = await redeemAtEntra(config, {
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      scope: scopes.join(" "),
    });
    if (data?.access_token) {
      storeTokens(data.access_token, data.refresh_token ?? refreshToken, data.expires_in);
      return data.access_token;
    }
  } catch (err) {
    // A rejected refresh token is not recoverable — drop it so the next
    // attempt goes straight to the dialog instead of retrying forever.
    console.warn("[msal] refresh failed", err);
    safeLocalSet(STORAGE_REFRESH_KEY, null);
  }
  return null;
}

function buildAuthorizeUrl(config, challenge, state) {
  const url = new URL(config.authorization_url);
  url.searchParams.set("client_id", config.client_id);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", redirectUriFor(config));
  url.searchParams.set("response_mode", "query");
  url.searchParams.set("scope", config.scopes.join(" "));
  url.searchParams.set("code_challenge", challenge);
  url.searchParams.set("code_challenge_method", "S256");
  url.searchParams.set("state", state);
  url.searchParams.set("prompt", "select_account");
  return url.toString();
}

function dialogStartUrl(authorizeUrl) {
  return `${window.location.origin}${AUTH_START_PATH}?url=${encodeURIComponent(authorizeUrl)}`;
}

/**
 * Open Microsoft's consent screen in an Office-owned dialog.
 *
 * This replaces window.open: the dialog is hosted by Outlook itself, so it
 * cannot be popup-blocked and never escapes to the OS default browser. The
 * auth code comes back in-process via messageParent, which means no storage
 * hand-off between windows is needed at all.
 */
function openAuthDialog(startUrl, onProgress) {
  return new Promise((resolve, reject) => {
    const ui = window.Office?.context?.ui;
    if (typeof ui?.displayDialogAsync !== "function") {
      reject(
        authError(
          "dialog_unavailable",
          "Office.context.ui.displayDialogAsync is not available in this host, so sign-in cannot be shown. Update Outlook or use Outlook on the web.",
        ),
      );
      return;
    }

    let settled = false;
    let dialog = null;
    let timer = null;

    function finish(fn, arg) {
      if (settled) return;
      settled = true;
      if (timer) clearTimeout(timer);
      try {
        dialog?.close();
      } catch {}
      fn(arg);
    }

    ui.displayDialogAsync(
      startUrl,
      { height: 60, width: 30, promptBeforeOpen: false },
      (result) => {
        if (result?.status !== Office.AsyncResultStatus.Succeeded) {
          const code = result?.error?.code ?? "?";
          finish(
            reject,
            authError(
              "dialog_open_failed",
              `Could not open the sign-in dialog (code ${code}): ${result?.error?.message ?? "unknown"}`,
            ),
          );
          return;
        }

        dialog = result.value;
        if (typeof onProgress === "function") onProgress("auth-dialog-opened");

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          let payload;
          try {
            payload = JSON.parse(arg.message);
          } catch {
            finish(
              reject,
              authError(
                "dialog_bad_message",
                "The sign-in dialog sent a message that could not be parsed.",
              ),
            );
            return;
          }
          finish(resolve, payload);
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          const code = arg?.error;
          if (code === 12006) {
            finish(
              reject,
              authError("dialog_cancelled", "Sign-in was cancelled — the dialog was closed."),
            );
            return;
          }
          finish(
            reject,
            authError(
              "dialog_event",
              `The sign-in dialog closed unexpectedly (code ${code ?? "?"}).`,
            ),
          );
        });

        timer = setTimeout(() => {
          finish(
            reject,
            authError(
              "dialog_timeout",
              `Sign-in did not complete within ${Math.round(DIALOG_TIMEOUT_MS / 1000)}s.`,
            ),
          );
        }, DIALOG_TIMEOUT_MS);
      },
    );
  });
}

export async function msalLogin(scopes, opts = {}) {
  const progress = typeof opts.onProgress === "function" ? opts.onProgress : () => {};

  const cached = readCachedAccessToken();
  if (cached) {
    progress("auth-cache-hit");
    return cached;
  }

  if (typeof window.crypto?.subtle?.digest !== "function") {
    throw authError(
      "no_webcrypto",
      "Web Crypto SubtleCrypto is not available in this WebView, so PKCE sign-in cannot run.",
    );
  }
  if (!window.isSecureContext) {
    throw authError(
      "insecure_context",
      "PKCE requires a secure context (https). The add-in pane does not appear to be secure.",
    );
  }

  const config = await loadConfig();

  const refreshed = await tryRefresh(config, scopes);
  if (refreshed) {
    progress("auth-refresh");
    return refreshed;
  }

  const verifier = makeCodeVerifier();
  const challenge = await makeCodeChallenge(verifier);
  const state = makeState();

  progress("auth-dialog-opening");
  const payload = await openAuthDialog(
    dialogStartUrl(buildAuthorizeUrl(config, challenge, state)),
    progress,
  );

  if (!payload?.ok) {
    throw authError(
      payload?.error ?? "auth_callback_failed",
      `Sign-in failed: ${payload?.error ?? "unknown"} ${payload?.error_description ?? ""}`.trim(),
    );
  }
  if (payload.state !== state) {
    throw authError(
      "state_mismatch",
      "Sign-in state did not match the value this pane generated — refusing the code.",
    );
  }
  if (!payload.code) {
    throw authError("missing_code", "The sign-in dialog returned no authorization code.");
  }

  progress("auth-exchanging-code");
  const data = await redeemAtEntra(config, {
    grant_type: "authorization_code",
    code: payload.code,
    code_verifier: verifier,
    redirect_uri: redirectUriFor(config),
    scope: config.scopes.join(" "),
  });

  if (!data?.access_token) {
    throw authError("no_access_token", "The token endpoint returned no access_token.");
  }

  storeTokens(data.access_token, data.refresh_token, data.expires_in);
  progress("auth-token-cached");
  return data.access_token;
}

/**
 * Drop only the access token, keeping the refresh token.
 *
 * Used when Graph rejects a token with 401: the grant itself is usually still
 * good, so the next call can refresh silently instead of re-prompting.
 */
export function invalidateAccessToken() {
  safeLocalSet(STORAGE_ACCESS_KEY, null);
  safeLocalSet(STORAGE_EXPIRES_KEY, null);
}

