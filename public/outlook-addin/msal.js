const DEFAULT_REDIRECT_URI =
  "https://schedule-gen-from-email.vercel.app/outlook-addin/auth-callback.html";

const CONFIG_CACHE_KEY = "addCalEvent.oauth.configCache";
const CONFIG_TTL_MS = 5 * 60 * 1000;

function safeStorageGet(key) {
  try {
    return window.sessionStorage?.getItem(key) ?? null;
  } catch {
    return null;
  }
}

function safeStorageSet(key, value) {
  try {
    if (value === null || value === undefined) window.sessionStorage.removeItem(key);
    else window.sessionStorage.setItem(key, value);
  } catch {}
}

async function loadConfig() {
  const cachedRaw = safeStorageGet(CONFIG_CACHE_KEY);
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
    throw new Error(
      `auth_config_unavailable (${data?.error ?? resp.status}): ${data?.message ?? ""}`.trim(),
    );
  }
  if (!data?.client_id || !data?.authorization_url || !Array.isArray(data.scopes)) {
    throw new Error("auth_config response is missing client_id/authorization_url/scopes");
  }
  safeStorageSet(CONFIG_CACHE_KEY, JSON.stringify({ ...data, __fetchedAt: Date.now() }));
  return data;
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

const STORAGE_STATE_KEY = "addCalEvent.oauthState";
const STORAGE_VERIFIER_KEY = "addCalEvent.pkceVerifier";
const STORAGE_ACCESS_KEY = "addCalEvent.accessToken";
const STORAGE_REFRESH_KEY = "addCalEvent.refreshToken";
const STORAGE_EXPIRES_KEY = "addCalEvent.accessTokenExpiresAt";

async function exchangeViaApi(payload) {
  const resp = await fetch("/api/exchange-token", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(
      `token_exchange_failed (${data?.error ?? resp.status}): ${data?.message ?? ""}`.trim(),
    );
  }
  return data;
}

function storeTokens(accessToken, refreshToken, expiresInSec) {
  safeStorageSet(STORAGE_ACCESS_KEY, accessToken);
  if (refreshToken) safeStorageSet(STORAGE_REFRESH_KEY, refreshToken);
  if (expiresInSec) {
    safeStorageSet(
      STORAGE_EXPIRES_KEY,
      String(Date.now() + (Number(expiresInSec) - 60) * 1000),
    );
  }
}

function clearTokens() {
  safeStorageSet(STORAGE_ACCESS_KEY, null);
  safeStorageSet(STORAGE_REFRESH_KEY, null);
  safeStorageSet(STORAGE_EXPIRES_KEY, null);
}

function readCachedAccessToken() {
  const access = safeStorageGet(STORAGE_ACCESS_KEY);
  const expiresAt = Number(safeStorageGet(STORAGE_EXPIRES_KEY));
  if (!access || !expiresAt) return null;
  if (Date.now() >= expiresAt) return null;
  return access;
}

async function tryRefresh(scopes) {
  const refreshToken = safeStorageGet(STORAGE_REFRESH_KEY);
  if (!refreshToken) return null;
  try {
    const data = await exchangeViaApi({
      refresh_token: refreshToken,
      scope: scopes.join(" "),
    });
    if (data?.access_token) {
      storeTokens(
        data.access_token,
        data.refresh_token ?? refreshToken,
        data.expires_in,
      );
      return data.access_token;
    }
  } catch (err) {
    console.warn("[msal] refresh failed", err);
  }
  return null;
}

async function popupLoginOnce(config) {
  if (typeof window.crypto?.subtle?.digest !== "function") {
    throw new Error(
      "Web Crypto SubtleCrypto is not available in this WebView. MSAL fallback requires a modern browser.",
    );
  }
  if (!window.isSecureContext) {
    throw new Error(
      "PKCE requires a secure context (https or localhost). The add-in iframe does not appear to be secure.",
    );
  }

  const verifier = makeCodeVerifier();
  const challenge = await makeCodeChallenge(verifier);
  const state = makeState();
  safeStorageSet(STORAGE_VERIFIER_KEY, verifier);
  safeStorageSet(STORAGE_STATE_KEY, state);

  const url = new URL(config.authorization_url);
  url.searchParams.set("client_id", config.client_id);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("redirect_uri", config.redirect_uri ?? DEFAULT_REDIRECT_URI);
  url.searchParams.set("response_mode", "query");
  url.searchParams.set("scope", config.scopes.join(" "));
  url.searchParams.set("code_challenge", challenge);
  url.searchParams.set("code_challenge_method", "S256");
  url.searchParams.set("state", state);
  url.searchParams.set("prompt", "select_account");

  return { url: url.toString(), state };
}

function awaitPopupMessage(expectedState, timeoutMs = 90_000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      window.removeEventListener("message", onMessage);
      reject(
        new Error(
          `Popup did not deliver an auth code within ${Math.round(timeoutMs / 1000)}s. The popup may have been blocked, the user cancelled, or the consent dialog stalled.`,
        ),
      );
    }, timeoutMs);

    function onMessage(event) {
      if (event.origin !== window.location.origin) return;
      const data = event.data;
      if (!data || data.source !== "addCalEvent.auth") return;
      clearTimeout(timer);
      window.removeEventListener("message", onMessage);
      resolve(data);
    }
    window.addEventListener("message", onMessage);
  });
}

export async function msalLogin(scopes, opts = {}) {
  const cached = readCachedAccessToken();
  if (cached) {
    if (typeof opts.onProgress === "function") opts.onProgress("auth-cache-hit");
    return cached;
  }

  const refreshed = await tryRefresh(scopes);
  if (refreshed) {
    if (typeof opts.onProgress === "function") opts.onProgress("auth-refresh");
    return refreshed;
  }

  if (typeof opts.onProgress === "function") opts.onProgress("auth-popup-opening");

  const config = await loadConfig();
  const { url, state } = await popupLoginOnce(config);
  const popup = window.open(
    url,
    "msal-consent",
    "width=520,height=640,menubar=no,toolbar=no,location=no,status=no",
  );
  if (!popup) {
    safeStorageSet(STORAGE_VERIFIER_KEY, null);
    safeStorageSet(STORAGE_STATE_KEY, null);
    throw new Error(
      "Popup was blocked. Allow popups for this origin and try again.",
    );
  }

  let payload;
  try {
    payload = await awaitPopupMessage(state);
  } catch (err) {
    try { popup.close(); } catch {}
    throw err;
  }
  try { popup.close(); } catch {}

  if (!payload?.ok) {
    safeStorageSet(STORAGE_VERIFIER_KEY, null);
    safeStorageSet(STORAGE_STATE_KEY, null);
    throw new Error(
      `PKCE callback failed: ${payload?.error ?? "unknown"} (${payload?.error_description ?? ""})`.trim(),
    );
  }

  const verifier = safeStorageGet(STORAGE_VERIFIER_KEY);
  safeStorageSet(STORAGE_VERIFIER_KEY, null);
  safeStorageSet(STORAGE_STATE_KEY, null);

  if (!payload.code || !verifier) {
    throw new Error("PKCE callback delivered no code or verifier was cleared.");
  }

  const data = await exchangeViaApi({
    code: payload.code,
    code_verifier: verifier,
    redirect_uri: config.redirect_uri ?? DEFAULT_REDIRECT_URI,
  });

  if (!data?.access_token) {
    throw new Error("Token endpoint returned no access_token.");
  }

  storeTokens(data.access_token, data.refresh_token, data.expires_in);
  if (typeof opts.onProgress === "function") opts.onProgress("auth-token-cached");
  return data.access_token;
}

export function resetMsalCache() {
  safeStorageSet(STORAGE_STATE_KEY, null);
  safeStorageSet(STORAGE_VERIFIER_KEY, null);
  clearTokens();
}
