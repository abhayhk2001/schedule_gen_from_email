import type { VercelRequest, VercelResponse } from "@vercel/node";
import { AUTHORITY, TOKEN_URL } from "../lib/oauth-config.js";

interface ExchangeBody {
  code?: string;
  code_verifier?: string;
  redirect_uri?: string;
  refresh_token?: string;
}

function missingField(field: string): string {
  return `${field} is required for this grant`;
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  const clientId = process.env.AZURE_CLIENT_ID?.trim();
  if (!clientId) {
    return res.status(500).json({
      error: "AZURE_CLIENT_ID is not configured on this deployment",
    });
  }

  const body =
    typeof req.body === "object" && req.body !== null
      ? (req.body as ExchangeBody)
      : {};

  const params = new URLSearchParams({ client_id: clientId });

  const refreshToken =
    typeof body.refresh_token === "string" ? body.refresh_token.trim() : "";
  if (refreshToken) {
    params.set("grant_type", "refresh_token");
    params.set("refresh_token", refreshToken);
    params.set(
      "scope",
      "openid offline_access User.Read Calendars.ReadWrite",
    );
  } else {
    const code = typeof body.code === "string" ? body.code.trim() : "";
    const verifier =
      typeof body.code_verifier === "string" ? body.code_verifier.trim() : "";
    const redirectUri =
      typeof body.redirect_uri === "string" ? body.redirect_uri.trim() : "";

    if (!code) return res.status(400).json({ error: missingField("code") });
    if (!verifier) {
      return res.status(400).json({ error: missingField("code_verifier") });
    }
    if (!redirectUri) {
      return res.status(400).json({ error: missingField("redirect_uri") });
    }

    params.set("grant_type", "authorization_code");
    params.set("code", code);
    params.set("code_verifier", verifier);
    params.set("redirect_uri", redirectUri);
  }

  let upstream: Response;
  try {
    upstream = await fetch(TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString(),
    });
  } catch (err) {
    return res.status(502).json({
      error: "token_endpoint_unreachable",
      message: err instanceof Error ? err.message : String(err),
      authority: AUTHORITY,
    });
  }

  const data = (await upstream.json().catch(() => ({}))) as {
    access_token?: string;
    refresh_token?: string;
    expires_in?: number;
    token_type?: string;
    scope?: string;
    error?: string;
    error_description?: string;
  };

  if (!upstream.ok || data.error) {
    return res.status(upstream.status || 400).json({
      error: data.error ?? "token_exchange_failed",
      message: data.error_description ?? "",
      authority: AUTHORITY,
    });
  }

  return res.status(200).json({
    access_token: data.access_token,
    refresh_token: data.refresh_token,
    expires_in: data.expires_in,
    scope: data.scope,
    token_type: data.token_type,
  });
}
