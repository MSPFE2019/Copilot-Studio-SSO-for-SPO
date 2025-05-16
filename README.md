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

1. **Create a Token Exchange URL** in your Copilot app registration:
    - Go to **Expose an API**.
    - Create a new scope, e.g., `SPO.Read`.
    - Copy the full URI of the scope (you will use it later as the Token Exchange URL).

2. **Grant the following Microsoft Graph delegated permissions** to the Copilot app:
    - `Files.Read.All`
    - `openid`
    - `profile`
    - `Sites.Read.All`
    - `User.Read`

3. **Set the Token Exchange URL** in Copilot Studio settings:
    ```text
    api://<your-client-id>/SPO.Read
    ```

4. *(Optional)* If you plan to use Generative Answers over SharePoint/OneDrive, grant additional Graph permissions.

> **Note:**  
> - In authoring, users sign in with a prompt.  
> - In production, users sign in silently via SSO if "Require users to sign in" is enabled.

---

## Step 2: Modify the Auto-Created Copilot Studio App Registration

1. Go to [Azure Portal](https://portal.azure.com) > **App registrations**, and locate the app auto-created by Copilot Studio.

2. Under **Authentication**:
    - Click **Add a platform** > **Single-page application (SPA)**.
    - Add the following redirect URI:
      ```
      https://<your-tenant>.sharepoint.com/sites/<your-site>
      ```
    - Enable both:
        - ✔️ Access tokens
        - ✔️ ID tokens

3. Under **Manifest**, find the `spa` section and append a wildcard `*`:
    ```json
    "spa": {
      "redirectUris": [
        "https://<your-tenant>.sharepoint.com/sites/<your-site>*"
      ]
    }
    ```

4. Go to **API permissions** > **Add a permission** > **My APIs**:
    - Select your Copilot Studio app.
    - Add the custom scope `SPO.Read`.

---

## Step 3: Download the Prebuilt SPFx Component

Download the prebuilt `.sppkg` file:

- [Download pva-extension-sso.sppkg](https://github.com/microsoft/CopilotStudioSamples/blob/main/SSOSamples/SharePointSSOComponent/sharepoint/solution/pva-extension-sso.sppkg)

> **Source and Support:**  
> Refer to [SharePoint SSO Component samples](https://github.com/microsoft/CopilotStudioSamples/tree/main/SSOSamples/SharePointSSOComponent) for updates or issues.

---

## Step 4: Upload the Component to SharePoint

1. Go to the **App Catalog** in SharePoint admin center.

2. Upload the `pva-extension-sso.sppkg` file.

3. Enable the app (do **not** choose “Enable this app and add to all sites”).

4. Navigate to the SharePoint site you plan to use, and add the app from "Site Contents".

---

## Step 5: Setup in Copilot Studio

1. In Copilot Studio, go to **Settings > Security > Authenticate manually**.

2. Fill in the following fields:
    - **Require users to sign in**: ✅
    - **Redirect URL**:
      ```
      https://token.botframework.com/.auth/web/redirect
      ```
    - **Service provider**:  
      ```
      Azure Active Directory v2
      ```
    - **Client ID**: Copilot app Client ID  
    - **Client Secret**: Copilot app Client Secret  
    - **Token Exchange URL (SSO)**:
      ```
      api://<your-client-id>/SPO.Read
      ```
    - **Tenant ID**: Azure AD Tenant ID  
    - **Scopes**:
      ```
      profile openid
      ```

---

## Step 6: Configure Site-Level Properties with PowerShell

1. Get your **bot URL**:
    - Open your bot in Power Virtual Agents.
    - Go to **Channels > Mobile App**.
    - Copy the **Token Endpoint**.
      Example:
      ```text
      https://f3479ae949c40ecb110d7d82a1729e2.3.environment.api.gov.powerplatform.microsoft.us/powervirtualagents/botsbyschema/cr48c_powerPlatformLicensingBot/directline/token?api-version=2022-03-01-preview
      ```

2. Clone or download the configuration script:
    ```
    https://github.com/MSPFE2019/Copilot-Studio-SSO-for-SPO
    ```

3. If needed, create a PnP PowerShell app registration:
    ```powershell
    .\PNP App Registration Creation.ps1
    ```

4. Run the configuration script:
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

> **Important:** `-clientId` is your Copilot Studio app ID, not the PnP app registration ID.

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


✅ Final Notes
----
Only one app registration (Copilot Studio’s) is required.

The provided .sppkg file is ready for deployment.

Use the PowerShell script to scope deployment and customize the chat button.
