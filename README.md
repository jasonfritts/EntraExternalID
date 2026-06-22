# SAML SP Tester for Entra External ID

## Register Client ID In Entra External ID

Download script: [Configure-EntraExternalIdSamlTestApp.ps1](scripts/Configure-EntraExternalIdSamlTestApp.ps1)

Run with generic parameters:

```powershell
.\scripts\Configure-EntraExternalIdSamlTestApp.ps1 `
	-TenantId "<your-tenant-id-guid>" `
	-TenantDomain "<your-tenant>.onmicrosoft.com" `
	-ExternalIdDomain "<your-tenant>.ciamlogin.com" `
	-AppDisplayName "SAML SP Tester" `
	-BaseUrl "http://localhost:3000"
```

Small test app that behaves like a simple SAML Service Provider (SP):

- Accepts `domain` + `client_id` (and optional `tenantDomain`) from the UI.
- Builds federation metadata URL for Entra External ID.
- Sends an SP-initiated SAML AuthnRequest to the IdP.
- Receives SAMLResponse at `/acs` and displays decoded XML plus common assertion fields.

## Quick start

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
