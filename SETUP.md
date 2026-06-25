# Portfolio Strategy Overview — Setup & Deployment Guide

## What this is
A real-time React dashboard connected directly to your SharePoint list via Microsoft Graph API.
Users sign in with their Adobe Microsoft account. Edit Mode is restricted to emails you configure.

---

## Step 1 — Register an Azure AD App (15 min, needs IT or you if you have Azure access)

1. Go to **portal.azure.com** → search "App registrations" → **New registration**
2. Fill in:
   - **Name:** Portfolio Strategy Overview
   - **Supported account types:** Accounts in this organizational directory only (Adobe)
   - **Redirect URI:** Single-page application (SPA) → `http://localhost:5173`
3. Click **Register**
4. On the Overview page, copy and save:
   - **Application (client) ID** → goes into `VITE_CLIENT_ID`
   - **Directory (tenant) ID** → goes into `VITE_TENANT_ID`

---

## Step 2 — Grant API Permissions (needs SharePoint/Global Admin consent)

1. In your app registration → **API permissions** → **Add a permission**
2. Choose **Microsoft Graph** → **Delegated permissions**
3. Search and add both:
   - `Sites.ReadWrite.All`
   - `User.Read`
4. Click **Grant admin consent for Adobe** (the blue button)
   - This requires a Global Admin or SharePoint Admin to click it once
   - After this, any Adobe user can sign in without individual consent prompts

---

## Step 3 — Configure environment variables

Copy `.env.example` to `.env` and fill in your values:

```
VITE_CLIENT_ID=paste-your-client-id-here
VITE_TENANT_ID=paste-your-tenant-id-here
VITE_EDITOR_EMAILS=you@adobe.com,colleague@adobe.com
```

`VITE_EDITOR_EMAILS` controls who sees the Edit Mode button.
Leave it blank to allow all signed-in users to edit (not recommended for production).

---

## Step 4 — Run locally to test

```bash
npm install
npm run dev
```

Open http://localhost:5173 → click "Sign in with Microsoft" → your SharePoint data should load.

If you see a permissions error, admin consent in Step 2 hasn't been granted yet.

---

## Step 5 — Deploy to Azure Static Web Apps

### Option A: GitHub (recommended, gets you automatic deployments on every push)

1. Push this folder to a GitHub repo
2. In **portal.azure.com** → Create a resource → **Static Web App**
3. Connect to your GitHub repo
4. Build settings:
   - **App location:** `/`
   - **Output location:** `dist`
   - **API location:** (leave blank)
5. Azure creates a GitHub Actions workflow automatically
6. After first deployment, copy the generated URL (e.g. `https://gentle-wave-xxx.azurestaticapps.net`)

### Option B: Manual deploy with SWA CLI

```bash
npm run build

# Install the Azure SWA CLI (one-time)
npm install -g @azure/static-web-apps-cli

# Deploy
swa deploy dist --deployment-token YOUR_DEPLOYMENT_TOKEN
```

Get your deployment token from: Azure portal → your Static Web App → Overview → Manage deployment token

---

## Step 6 — Add production URL to Azure AD App

1. Back in **portal.azure.com** → App registrations → your app → **Authentication**
2. Under "Single-page application" redirect URIs, click **Add URI**
3. Add your Azure Static Web Apps URL: `https://your-app-name.azurestaticapps.net`
4. Click **Save**

---

## Step 7 — Set environment variables in Azure (for production)

In your Static Web App in Azure portal → **Configuration** → Add application settings:
- `VITE_CLIENT_ID` = your client ID
- `VITE_TENANT_ID` = your tenant ID  
- `VITE_EDITOR_EMAILS` = comma-separated editor emails

> Note: These Vite env vars get baked in at build time. If you use GitHub Actions,
> add them as GitHub repo secrets and reference them in the workflow file instead.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| "AADSTS50011: Reply URL mismatch" | Redirect URI not registered | Add your URL to Azure AD app → Authentication |
| "Insufficient privileges" on load | Admin consent not granted | Ask IT to grant consent in Step 2 |
| All sites load but columns empty | SharePoint field names differ | Open browser console, the field map is logged |
| Edit Mode button not visible | Email not in VITE_EDITOR_EMAILS | Add your email to the env var |
| "interaction_in_progress" error | Multiple login popups triggered | Refresh the page and try again |

---

## Architecture

```
Browser (Azure Static Web App)
  └── React + MSAL.js
        ↓ Sign in with Microsoft (popup)
        ↓ Get access token (silent refresh)
        ↓ Microsoft Graph API calls
              ↓ sites/Monitor_SP/lists/portfolio strategy data temp/items
              ↓ Read (GET) / Write (PATCH) / Create (POST) / Delete (DELETE)
```

No backend needed. No API keys in client code. Auth is handled entirely by Microsoft.
