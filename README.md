````markdown
# Deploy Microsoft Copilot Studio Copilot as a SharePoint Component with SSO

This guide explains how to deploy a **Microsoft Copilot Studio Copilot** as a **SharePoint Framework (SPFx) component** with **Single Sign-On (SSO)** enabled using a prebuilt `.sppkg` file.

---

## Overview

To complete the deployment, follow these high-level steps:

1. Configure Microsoft Entra ID authentication for your Copilot.  
2. Modify the auto-created Copilot Studio App Registration in Azure AD.  
3. Download the prebuilt SPFx component.  
4. Upload the SPFx component to SharePoint and configure it.  
5. Configure site-level properties via PowerShell.

---

## Step 1: Configure Microsoft Entra ID Authentication for Your Copilot

Follow the [official guide](https://learn.microsoft.com/en-us/power-virtual-agents/configure-user-authentication) and then:

1. **Create Token Exchange URL** in your Copilot app registration:  
   - Navigate to **Expose an API**.  
   - Create a new scope (e.g. `SPO.Read`).  
   - Copy the full URI of that scope (you will use it as the Token Exchange URL).

2. **Grant Copilot the following Microsoft Graph delegated permissions**:  
   - `Files.Read.All`  
   - `openid`  
   - `profile`  
   - `Sites.Read.All`  
   - `User.Read`

3. **Populate the Token Exchange URL** in your Copilot authentication settings with your custom scope URI:  
   ```text
   api://<your-client-id>/SPO.Read
````

4. *(Optional)* If you plan to use **Generative Answers** over SharePoint/OneDrive, grant any additional Graph permissions required.

> **Note:**
>
> * In authoring, sign-in is validated via a code prompt.
> * In production, users will sign in silently via SSO if **Require users to sign in** is enabled.

---

## Step 2: Modify the Auto-Created Copilot Studio App Registration

1. In the **Azure AD** portal, locate the App Registration automatically created by Copilot Studio (named after your Copilot).
2. Select **Authentication**, then **Add a platform** → **Single-page application (SPA)**.

   * Under **Redirect URIs**, add:

     ```
     https://<your-tenant>.sharepoint.com/sites/<your-site>
     ```
   * Enable both:

     * **Access tokens** (used for implicit flows)
     * **ID tokens** (used for implicit and hybrid flows)
3. Click **Manifest**, locate the `spa` section, append a wildcard (`*`) to your URI(s), for example:

   ```json
   "spa": {
     "redirectUris": [
       "https://<your-tenant>.sharepoint.com/sites/<your-site>*"
     ]
   }
   ```

   Then save the manifest.
4. Under **API permissions**, click **Add a permission → My APIs**, select your Copilot Studio app, and grant the custom scope you created in Step 1 (e.g. `SPO.Read`).

---

## Step 3: Download the Prebuilt SPFx Component

Download the `.sppkg` file:

* [pva-extension-sso.sppkg](https://github.com/microsoft/CopilotStudioSamples/blob/main/SSOSamples/SharePointSSOComponent/sharepoint/solution/pva-extension-sso.sppkg)

> **Original source and version updates:**
> [SharePoint SSO Component samples](https://github.com/microsoft/CopilotStudioSamples/tree/main/SSOSamples/SharePointSSOComponent)
> **Support:** Please direct any questions about the SharePoint package to that repository.

---

## Step 4: Upload the Component to SharePoint

1. Go to the SharePoint **App Catalog** in the Microsoft 365 admin center.
2. Upload **pva-extension-sso.sppkg**.
3. Select **Enable App** (do **not** choose “Enable this app and add to all sites”).
4. On the target site (the same one used as your redirect URI), add the newly uploaded app.

---

## Setup in Copilot Studio

In Copilot Studio, navigate to **Settings → Security → Authenticate manually** and configure:

* **Require users to sign in**: ☑️
* **Redirect URL**:

  ```
  https://token.botframework.com/.auth/web/redirect
  ```
* **Service provider**:

  ```
  Azure Active Directory v2
  ```
* **Client ID**: Your Copilot app’s Client ID
* **Client Secret**: Your Copilot app’s Client Secret
* **Token Exchange URL (SSO)**:

  ```
  api://<your-client-id>/SPO.Read
  ```
* **Tenant ID**: Your Azure AD tenant ID
* **Scopes**:

  ```
  profile openid
  ```

---

## Step 5: Configure Site-Level Properties

To get your **bot URL**, go to the Power Virtual Agents portal for your bot, select **Channels**, choose **Mobile App**, then copy the **Token Endpoint**. Use this URL as the `-botUrl` parameter in the script. It should resemble:

```text
https://f3479ae949c40ecb110d7d82a1729e2.3.environment.api.gov.powerplatform.microsoft.us/powervirtualagents/botsbyschema/cr48c_powerPlatformLicensingBot/directline/token?api-version=2022-03-01-preview
```

1. Clone or download the repository containing **Configure-McsForSite.ps1**:

   ```
   https://github.com/MSPFE2019/Copilot-Studio-SSO-for-SPO
   ```
2. If you do not yet have a PnP PowerShell app registration, run:

   ```powershell
   .\PNP App Registration Creation.ps1
   ```
3. Run the configuration script:

   ```powershell
   .\Configure-McsForSite.ps1 `
     -siteUrl      "https://<your-tenant>.sharepoint.com/sites/<your-site>" `
     -botUrl       "<your-token-endpoint-url>" `
     -botName      "Copilot Assistant" `
     -greet        `
     -customScope  "api://<your-client-id>/SPO.Read" `
     -clientId     "<copilot-client-id>" `
     -authority    "https://login.microsoftonline.com/<tenant-id>" `
     -buttonLabel  "Chat Now"
   ```

> **Note:**
> The `-clientId` parameter refers to **your Copilot Studio app**, not the PnP registration.

After successful execution, a **Chat** button will appear on every page of your specified site. Clicking it launches the Copilot canvas with seamless SSO.

---

## A. Authentication Flow Diagram

```text
[ User ] ──▶ [ SharePoint Page (SPFx) ] ──▶ [ MSAL.js ] ──▶ [ Azure AD Token ]
                                   │                         │
                                   └────── SSO via SPA + Copilot Custom Scope ──────┘
                                            ▼
                                     [ Bot Framework WebChat ]
```

---

## B. PowerShell Script Template (Example)

```powershell
.\Configure-McsForSite.ps1 `
  -siteUrl      "https://contoso.sharepoint.com/sites/hrportal" `
  -botUrl       "https://hrbot.botframework.com" `
  -botName      "HR Assistant" `
  -greet        `
  -customScope  "api://46982118-ac3a-424f-b9a2-fee8ced5708f/SPO.Read" `
  -clientId     "12345678-aaaa-bbbb-cccc-ddddeeeeffff" `
  -authority    "https://login.microsoftonline.com/contoso.onmicrosoft.com" `
  -buttonLabel  "Chat with HR"
```

---

## ✅ Final Notes

* Only **one** app registration is required—Copilot Studio’s auto-generated registration.
* The provided `.sppkg` is prebuilt for ease of deployment.
* Use the PowerShell script to customize the chat button and scope site by site.

```
```
