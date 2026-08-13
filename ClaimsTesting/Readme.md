
# Entra Optional Claims Demo Tools

Three single-file PowerShell + Windows Forms tools for **demonstrating Microsoft Entra ID optional claims**. Each tool creates (or reuses) an app registration, lets you pick which optional claims to request from the [official optional claims reference](https://learn.microsoft.com/en-us/entra/identity-platform/optional-claims-reference), signs a user in, and shows the **decoded, human-readable tokens** so you can confirm exactly which claims were emitted.

| Script | Protocol / Endpoint | Sign-in method | Tokens shown |
| --- | --- | --- | --- |
| `OptionalClaimsDemo.ps1` | OAuth 2.0 / OpenID Connect (**v2.0** endpoint) | MSAL interactive | v2 `id_token` + access token |
| `OptionalClaimsDemoV1.ps1` | **v1.0** endpoint (`/oauth2/authorize` + `/oauth2/token`) | Manual auth-code flow (loopback listener) | v1 `id_token` + access token |
| `SamlOptionalClaimsDemo.ps1` | **SAML 2.0** SSO | SP-initiated SAML (loopback ACS listener) | Signed SAML **assertion** (attribute claims) |

---

## What each script does

### `OptionalClaimsDemo.ps1` (v2.0 / OIDC)
- Connects to Microsoft Graph as an admin (`Application.ReadWrite.All`).
- Creates/reuses a public-client app that:
  - Signs in interactively via `http://localhost` (MSAL).
  - Exposes its own API scope `api://<appId>/access_as_user` so **access-token** optional claims can be demonstrated.
  - Has Microsoft Graph `User.Read` for comparison.
- Lets you check optional claims (each tagged `[ID]`, `[Access]`, or `[ID+Access]`) plus a free-text box for any claim not in the list, then PATCHes them onto the app's `optionalClaims`.
- Signs in and shows the decoded header, payload, readable timestamps, and full claim list for both the **ID token** and the **access token**.

### `OptionalClaimsDemoV1.ps1` (v1.0)
- Same experience, but uses the **Azure AD v1.0 endpoint for everything** (MSAL is v2-only, so this performs a manual v1 authorization-code flow with a local loopback `HttpListener`).
- Uses the v1 **`resource`** parameter instead of scopes.
- App is configured with `requestedAccessTokenVersion = 1`, so you get **v1-format tokens** (note the `ver` claim, `upn`/`unique_name`, resource-style `aud`).
- Resource selector: **This app's API** (`resource=<appId GUID>`) or **Microsoft Graph** (`resource=https://graph.microsoft.com`).

> **Note on access-token optional claims:** they only appear when **your app is the token audience** (the "This app's API" option). A Microsoft Graph access token is owned by Graph, so your custom optional claims will **not** appear in it. ID-token optional claims always reflect your selection.

### `SamlOptionalClaimsDemo.ps1` (SAML 2.0)
- Same experience, but for the **SAML** protocol so you can inspect the raw **SAML Response / assertion** instead of JWTs.
- Connects to Microsoft Graph as an admin (`Application.ReadWrite.All`) and creates/reuses an enterprise app configured for SAML SSO:
  - Service principal `preferredSingleSignOnMode = 'saml'`.
  - A **token-signing certificate** so Entra can sign the SAML Response.
  - Entity ID `api://<appId>` and an ACS (Assertion Consumer Service) reply URL pointing at a local listener the tool hosts on `http://localhost:<port>/acs/`.
  - `appRoleAssignmentRequired = false` so no explicit assignment is needed to test.
- Applies your selected optional claims to `optionalClaims.saml2Token`.
- Runs **SP-initiated SSO**: builds a SAML `AuthnRequest`, opens the browser to Entra's SAML endpoint, then captures the POSTed `SAMLResponse` on the local ACS listener, base64-decodes it, and shows:
  - **Claims (readable)** — NameID, Issuer, Audience, conditions, and every `<Attribute>` claim.
  - **SAML Response (XML)** — pretty-printed assertion.
  - **Raw (base64)** — the exact `SAMLResponse` form value.

> **SAML claims differ from OIDC:** there is no ID/access token — claims are XML `<Attribute>` elements inside the signed assertion, optional claims live under `optionalClaims.saml2Token`, and the user identifier is the `<NameID>` in the Subject.

---

## Prerequisites

- **Windows** with **Windows PowerShell 5.1** (STA by default) or **PowerShell 7+** (`pwsh -STA`).
- An Entra tenant account with permission to **create app registrations** and grant admin consent (`Application.ReadWrite.All`).
- Internet access to install modules (auto-installed to `CurrentUser` on first run):
  - `Microsoft.Graph.Authentication`
  - `Microsoft.Graph.Applications`
  - `MSAL.PS` *(v2/OIDC script only)*

---

## How to run

### 1. Launch the tool
Windows Forms requires an STA thread. In Windows PowerShell 5.1 this is the default; in PowerShell 7 pass `-STA`.

```powershell
# v2.0 / OIDC tool
powershell -STA -File ".\OptionalClaimsDemo.ps1"

# v1.0 tool
powershell -STA -File ".\OptionalClaimsDemoV1.ps1"

# SAML 2.0 tool
powershell -STA -File ".\SamlOptionalClaimsDemo.ps1"
```

Optional parameters:

```powershell
# Pin a tenant and/or a custom app display name
powershell -STA -File ".\OptionalClaimsDemo.ps1" -TenantId "contoso.onmicrosoft.com" -AppDisplayName "My Claims Demo"

# v1 tool: change the loopback port used by the redirect listener (default 8400)
powershell -STA -File ".\OptionalClaimsDemoV1.ps1" -Port 8450

# SAML tool: change the local ACS listener port (default 8080)
powershell -STA -File ".\SamlOptionalClaimsDemo.ps1" -AcsPort 8090
```

### 2. Connect (admin)
Click **1) Connect (admin)**. Optionally enter a tenant id/domain first (blank = interactive). Approve the sign-in / consent prompt. The status bar shows the connected tenant and account.

### 3. Choose optional claims and create/update the app
1. Check the optional claims you want in the list (hover any item for a description).
2. Optionally add extra claim names in the **Extra claim names** box (comma-separated).
3. Click **2) Create / Update App**. This creates (or reuses) the app registration and applies your selected claims.

> Manifest changes can take a few minutes to propagate. If a claim doesn't appear, wait and sign in again.

### 4. Sign in and decode
1. Pick the token resource:
   - **This app's API** — shows access-token optional claims.
   - **Microsoft Graph** — for comparison (custom claims not shown).
2. Leave **Force fresh sign-in** checked so a new token (reflecting your claim changes) is issued.
3. Click **3) Sign In & Decode Tokens**.
   - *v2 tool:* an MSAL browser/WAM window opens.
   - *v1 tool:* your browser opens to the v1 authorize endpoint; the code is captured on the loopback listener.
4. Review the **ID Token** and **Access Token** tabs — decoded header, payload, readable timestamps, and the full list of claim names. Raw JWTs are on the **Raw** tabs; the **Log** tab mirrors the terminal.

### 5. Close
Use the red **Close / Kill** button to exit — it remains responsive even while a sign-in is pending (auth runs on a background runspace).

### SAML flow (`SamlOptionalClaimsDemo.ps1`)
Steps 2–4 differ slightly for the SAML tool:
- **2) Create / Update SAML App** configures the enterprise app, adds a token-signing certificate, and applies your claims to `saml2Token`. Allow a few minutes for the SAML config + signing cert to propagate before the first sign-in.
- **3) Start SAML SSO & Capture Response** hosts the local ACS listener, opens the browser to Entra's SAML endpoint, and captures the signed `SAMLResponse`. Review the **Claims (readable)**, **SAML Response (XML)**, and **Raw (base64)** tabs.
- The ACS reply URL uses `http://localhost:<port>/acs/`. If your tenant rejects a non-HTTPS SAML reply URL, the error appears on step 2.
- To test **IdP-initiated** SSO from the My Apps portal, assign your user to the enterprise app in the Entra portal first, and keep the tool's ACS listener running (click step 3) so the My Apps tile has somewhere to POST.

---

## Tips & troubleshooting

- **A claim you requested didn't appear:** allow a few minutes for propagation, confirm you're viewing the right token type, and remember some claims only exist for certain account types (e.g., `upn` is not present for personal Microsoft accounts).
- **`login_hint` looks encoded:** that's expected — it's an opaque value meant to be passed back on a later sign-in, not decoded.
- **The sign-in window seems stuck:** it may be behind the tool window; check the taskbar. The Log tab/terminal prints a heartbeat while waiting.
- **v1 self-token error (AADSTS90009):** the v1 tool requests the app's own API using the **App ID GUID** as the resource (required for self-token requests).

---

## Examples

> Replace the placeholders below with your own screenshots.

### `OptionalClaimsDemo.ps1` (v2.0 / OIDC)

<!-- Upload your screenshot and update the path/URL below -->
<img width="1080" height="749" alt="image" src="https://github.com/user-attachments/assets/f8e1953f-4b04-41de-8fd0-a7d8fbca473e" />


*Example: the tool after signing in, showing the decoded v2 `id_token` with the selected optional claims.*

### `OptionalClaimsDemoV1.ps1` (v1.0)

<!-- Upload your screenshot and update the path/URL below -->
<img width="982" height="753" alt="image" src="https://github.com/user-attachments/assets/e678c980-541e-49ed-b0ae-85566ea88d05" />


*Example: the tool after a v1 sign-in, showing the `ver: 1.0` token with v1-style claims (`upn`, `unique_name`).*

### `SamlOptionalClaimsDemo.ps1` (SAML 2.0)

<!-- Upload your screenshot and update the path/URL below -->
![SamlOptionalClaimsDemo.ps1 - decoded SAML assertion](images/samloptionalclaimsdemo.png)

*Example: the tool after a SAML sign-in, showing the decoded assertion with NameID, Issuer, Audience, and the `<Attribute>` claims.*

---

## Notes

- These tools create real app registrations in your tenant. Delete the demo apps (`Optional Claims Demo` / `Optional Claims Demo (v1)` / `SAML Optional Claims Demo`) when you're done if you don't need them.
- No secrets are used; both tools rely on public-client / interactive flows.
- Reference: [Microsoft Entra optional claims](https://learn.microsoft.com/en-us/entra/identity-platform/optional-claims-reference).
