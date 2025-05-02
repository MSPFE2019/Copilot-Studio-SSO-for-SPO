# Copilot-Studio-SSO-for-SPO

# Deploy Microsoft Copilot Studio Copilot as a SharePoint Component with SSO

This guide explains how to deploy a **Microsoft Copilot Studio Copilot** as a **SharePoint Framework (SPFx) component** with **Single Sign-On (SSO)** enabled using a prebuilt `.sppkg` file.

---

## Overview

To complete the deployment, follow these high-level steps:

1. Configure Microsoft Entra ID authentication for your Copilot.
2. Register your SharePoint site as a Canvas app.
3. (Optional) Build the SPFx component yourself — or use the prebuilt one.
4. Upload the SPFx component to SharePoint and configure it.

---

## Step 1: Configure Microsoft Entra ID Authentication for Your Copilot

Follow the [official guide](https://learn.microsoft.com/en-us/power-virtual-agents/configure-user-authentication) and additionally:

* ✅ **Required**: Populate the **Token Exchange URL** in your Copilot authentication settings. This must be the full URI of the **custom scope**.
* ✅ **Optional**: If using **Generative Answers** over SharePoint/OneDrive, grant additional API permissions.

### Example Token Exchange URL

```
api://46982118-ac3a-424f-b9a2-fee8ced5708f/SPO.Read
```

Once configured, verify the Copilot canvas signs you in correctly. If "Require users to sign in" is enabled, the canvas should auto-trigger sign-in.

> ⚠️ Signing in during authoring uses a validation code, but post-deployment, users will sign in silently via SSO.

---

## Step 2: Register Your SharePoint Site as a Canvas App

Create an app registration in Azure AD with the following settings:

* Platform: **Single-page application (SPA)**
* Redirect URIs:

  * `https://yourtenant.sharepoint.com/sites/yoursite`
  * `https://yourtenant.sharepoint.com/sites/yoursite/`

> ⚠️ **Redirect URIs are case-sensitive**. Ensure you add both with and without the trailing slash.

Grant this app permission to the **Copilot custom API scope** you created in Step 1.

---

## Step 3: Download and Configure the SPFx Component

You have two options:

### Option A: Use Prebuilt Package

Download the `.sppkg` file from the official GitHub repo:

* [pva-extension-sso.sppkg](https://github.com/microsoft/CopilotStudioSamples/blob/main/SSOSamples/SharePointSSOComponent/sharepoint/solution/pva-extension-sso.sppkg)

Skip to **Step 4** if you choose this route.

### Option B: Build It Yourself

1. Clone the repo:

```
git clone https://github.com/microsoft/CopilotStudioSamples.git
cd CopilotStudioSamples/SSOSamples/SharePointSSOComponent
```

2. Populate `elements.xml` using one of:

   * `python .\populate_elements_xml.py`
   * Manually update JSON string
   * Leave untouched and configure via PowerShell
3. Build:

```
npm install
gulp bundle --ship
gulp package-solution --ship
```

This generates the `.sppkg` under `sharepoint/solution/`.

---

## Step 4: Upload the Component to SharePoint

1. Open the [SharePoint App Catalog](https://admin.microsoft.com/)
2. Upload `pva-extension-sso.sppkg`
3. Choose **Enable App**, *not* "Enable this app and add to all sites"
4. Add the app to the same site used as your redirect URI.

---

## Step 5: Configure Site-Level Properties

Run the helper script to configure the SPFx component settings:

```powershell
.\Configure-McsForSite.ps1 \ 
  -siteUrl "https://yourtenant.sharepoint.com/sites/yoursite" \ 
  -botUrl "https://yourcopilot.botframework.com" \ 
  -botName "Copilot Assistant" \ 
  -greet \ 
  -customScope "api://46982118-ac3a-424f-b9a2-fee8ced5708f/SPO.Read" \ 
  -clientId "<copilot-client-id>" \ 
  -authority "https://login.microsoftonline.com/<tenant-id>" \ 
  -buttonLabel "Chat Now"
```

> 🔍 The script uses PnP PowerShell. The `clientId` in the script refers to **your Copilot Studio app**, not the PowerShell app registration.

After running this, a **Chat** button appears on every page in your SharePoint site. Clicking it launches the Copilot canvas with seamless SSO.

---

## A. Authentication Flow Diagram

```text
[User] ─▶ [SharePoint Page w/ SPFx] ─▶ [MSAL.js] ─▶ [Azure AD Token] ─▶ [Copilot Token Exchange URL] ─▶ [BotFramework WebChat]
                                                │                                 ▲
                                                └───── SSO with SPA + Copilot Custom Scope ─────┘
```

---

## B. PowerShell Script Template (Pre-Filled Example)

```powershell
.\Configure-McsForSite.ps1 \ 
  -siteUrl "https://contoso.sharepoint.com/sites/hrportal" \ 
  -botUrl "https://hrbot.botframework.com" \ 
  -botName "HR Assistant" \ 
  -greet \ 
  -customScope "api://46982118-ac3a-424f-b9a2-fee8ced5708f/SPO.Read" \ 
  -clientId "12345678-aaaa-bbbb-cccc-ddddeeeeffff" \ 
  -authority "https://login.microsoftonline.com/contoso.onmicrosoft.com" \ 
  -buttonLabel "Chat with HR"
```

Use this as a reference or automation input when managing multiple sites.

---

## ✅ Final Notes

* Only **one** app registration is required — the Copilot’s.
* The `.sppkg` is prebuilt, minimizing setup effort.
* Redirect URIs **must match casing and trailing slash** of your SharePoint URL.
* Use the PowerShell script to set/override deployment properties site-by-site.
