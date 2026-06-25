// src/authConfig.js
// Microsoft Authentication Library (MSAL) configuration

export const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_TENANT_ID}`,
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    // sessionStorage clears on tab close; use localStorage for persistent login
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

// Permissions the app needs — admin must grant consent once in Azure portal
export const graphScopes = {
  scopes: ["Sites.ReadWrite.All", "User.Read"],
};

// Emails allowed to use Edit Mode (read from .env / Azure SWA env vars)
export const EDITOR_EMAILS = (import.meta.env.VITE_EDITOR_EMAILS || "")
  .split(",")
  .map((e) => e.trim().toLowerCase())
  .filter(Boolean);
