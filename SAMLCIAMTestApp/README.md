# SAML SP Tester for Entra External ID

This tutorial will show you how to test Entra External ID as a SAML Identity Provider.  The web app has already been deployed to https://aka.ms/samlciamtestapp .  To utilize in your own Entra External ID tenant, use the below powershell script to register the test app in your tenant and then populate the client id and tenant info with your tenant to sign in and inspect the SAML Response.

## Step 1. Register Client ID In Entra External ID

Download script: [Configure-EntraExternalIdSamlTestApp.ps1](scripts/Configure-EntraExternalIdSamlTestApp.ps1)

Run with generic parameters:

```powershell
.\scripts\Configure-EntraExternalIdSamlTestApp.ps1 `
	-TenantId "<your-tenant-id-guid>" `
	-TenantDomain "<your-tenant>.onmicrosoft.com" `
	-ExternalIdDomain "<your-tenant>.ciamlogin.com" `
	-AppDisplayName "SAML SP Tester" `
	-BaseUrl ""https://samlciam-ca-53525.azurewebsites.net""
```

Small test app that behaves like a simple SAML Service Provider (SP):

- Accepts `domain` + `client_id` (and optional `tenantDomain`) from the UI.
- Builds federation metadata URL for Entra External ID.
- Sends an SP-initiated SAML AuthnRequest to the IdP.
- Receives SAMLResponse at `/acs` and displays decoded XML plus common assertion fields.
- 
## Step 2. Test The SAML Test App
1. Visit  https://aka.ms/samlciamtestapp
2. Populate the tenant info + client ID you created in previous step

## EXTRA: SAML Test App Source Code

If you are interested in deploying this test web app yourself, you can download the source from this repository using git and run locally using below steps

1. Install dependencies:

```bash
npm install
```

2. Run:

```bash
npm start
```

3. Open:

- `http://localhost:3000`

## Runtime settings

- `PORT` (optional): defaults to `3000`
- `BASE_URL` (optional): force ACS base URL, e.g. `https://myapp.example.com`
- `SESSION_SECRET` (recommended in non-dev)

## Notes

- This utility is intended for **inspection/testing**.
- It currently does **not** validate XML signatures or decrypt encrypted assertions.
- The app attempts tenant domain auto-derivation from domain (`contoso.ciamlogin.com` -> `contoso.onmicrosoft.com`). Override `tenantDomain` if needed.
