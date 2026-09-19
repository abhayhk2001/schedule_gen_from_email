export const GRAPH_DEFAULT_SCOPES = [
  "openid",
  "offline_access",
  "User.Read",
  "Calendars.ReadWrite",
];

export const REDIRECT_URI =
  process.env.OAUTH_REDIRECT_URI ??
  "https://schedule-gen-from-email.vercel.app/outlook-addin/auth-callback.html";

export const AUTHORITY =
  process.env.OAUTH_AUTHORITY ??
  "https://login.microsoftonline.com/common";

export const AUTHZ_URL = `${AUTHORITY}/oauth2/v2.0/authorize`;
export const TOKEN_URL = `${AUTHORITY}/oauth2/v2.0/token`;

