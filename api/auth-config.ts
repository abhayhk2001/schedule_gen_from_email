import type { VercelRequest, VercelResponse } from "@vercel/node";
import {
  AUTHORITY,
  AUTHZ_URL,
  TOKEN_URL,
  REDIRECT_URI,
  GRAPH_DEFAULT_SCOPES,
} from "../lib/oauth-config.js";

export default function handler(_req: VercelRequest, res: VercelResponse) {
  res.setHeader("Cache-Control", "no-store");
  const clientId = process.env.AZURE_CLIENT_ID?.trim();
  if (!clientId) {
    return res.status(503).json({
      error: "AZURE_CLIENT_ID is not configured on this deployment",
    });
  }
  return res.status(200).json({
    client_id: clientId,
    authority: AUTHORITY,
    authorization_url: AUTHZ_URL,
    token_url: TOKEN_URL,
    redirect_uri: REDIRECT_URI,
    scopes: GRAPH_DEFAULT_SCOPES,
  });
}
