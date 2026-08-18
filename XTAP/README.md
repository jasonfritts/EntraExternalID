# Update-XtapPartnerPolicy

A single-file **Windows Forms GUI** (PowerShell) for adding or removing an application from the **inbound** or **outbound** B2B collaboration application list of a Microsoft Entra **cross-tenant access policy (XTAP)** partner configuration.

It talks to Microsoft Graph, lists your configured partners, shows the live policy JSON, previews the exact change before you commit, and only writes after explicit confirmation.

---

## What it does

- Connects to Microsoft Graph with admin credentials.
- Lists all **partner-specific** cross-tenant access policies in your tenant.
- Lets you pick a partner and choose **Inbound** or **Outbound**.
- Shows the **current** policy and a live **proposed** policy as you add/remove an app.
- Preserves the rest of the policy (`usersAndGroups`, `accessType`) — it only edits the `applications.targets` list.
- Requires a Yes/No confirmation, applies the change via a Graph `PATCH`, then re-reads and displays the saved result.

> The application list is matched against **both** the client app and the resource app in a cross-tenant token request — a match on either satisfies the rule.

---

## Prerequisites

- **PowerShell** 5.1 (Windows PowerShell) or 7+ (`pwsh`) on Windows (WinForms is Windows-only).
- The **Microsoft.Graph.Authentication** module:
  ```powershell
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  ```
- Graph permission **`Policy.ReadWrite.CrossTenantAccess`** (e.g. **Security Administrator** or equivalent). You'll consent on first sign-in.

> **Direction ownership:** *Inbound* policy is owned by the **resource** tenant (which of your apps external users may access). *Outbound* policy is owned by **your (home)** tenant (which partner apps your users may access). Sign in to the tenant that owns the direction you want to edit.

---

## Running it

```powershell
# Fully interactive
.\Update-XtapPartnerPolicy.ps1

# Optional: pre-seed values (you still select the partner and confirm)
.\Update-XtapPartnerPolicy.ps1 -PartnerTenantId <guid> -Direction Outbound -AppId <guid> -Action Add
```

If WinForms throws a threading error on PowerShell 7, launch in single-threaded apartment mode:
```powershell
pwsh -STA -File .\Update-XtapPartnerPolicy.ps1
```
(`powershell.exe` is STA by default.)

### Parameters (all optional)

| Parameter | Values | Purpose |
|---|---|---|
| `-PartnerTenantId` | GUID | Pre-selects this partner in the dropdown after connecting. |
| `-Direction` | `Inbound` \| `Outbound` | Pre-selects the policy direction. |
| `-AppId` | GUID | Pre-fills the application ID field. |
| `-Action` | `Add` \| `Remove` | Pre-selects the action. |

---

## Step-by-step in the UI

1. **Connect (admin)** — sign in. The tool loads your partner policies.
2. **Partner policy** — pick a partner tenant from the dropdown (use **Refresh** to reload).
3. **Policy direction** — choose **Inbound** or **Outbound**; the current policy repaints instantly.

   <img width="783" height="751" alt="image" src="https://github.com/user-attachments/assets/e7e66948-1dd2-4ece-a528-eca8091da9ac" />

   
4. **Application (App) ID** — paste the app's GUID.
5. **Action** — choose **Add** or **Remove**; the **PROPOSED** JSON updates live.
6. **Apply** — confirm Yes/No. The change is saved and the **CURRENT** view refreshes to the result.

---

## Safety features

- **Preview before write** — current vs. proposed JSON shown side by side; nothing is saved until you click **Apply** and confirm.
- **Backup before write** — an optional (default-on) checkbox saves the partner's **full current policy** (both directions and all settings) to a timestamped `.json` file before applying, so you can restore. Cancelling the save dialog aborts the change.
- **No-op guards** — Apply is disabled when adding an app that's already present or removing one that isn't.
- **Block-list warning** — if the list's `accessType` is `blocked` and you choose **Add**, the tool warns that a partner policy has a single applications list with one `accessType` (you can't keep both an allow-list and a block-list for the same partner) and requires an acknowledgment checkbox before enabling Apply.
- **`AllApplications` note** — flags a redundant specific-app add on an allow-list.
- **Non-destructive edits** — only `applications.targets` is changed; `usersAndGroups` is preserved.

---

## Notes & limitations

- The portal's application picker often can't find first-party **client** apps (e.g. by name or appId); this tool writes the raw appId directly via Graph, which works regardless. The portal may then display the app as a bare GUID — that's expected.
- Only **explicitly configured** partners appear in the list (the `default` configuration is not listed).
- **Public cloud** Graph endpoint (`graph.microsoft.com`). For sovereign clouds, point at the appropriate national cloud endpoint.
- Uses Microsoft Graph **v1.0**.

---

## Disclaimer

Provided as-is, with no warranty. Cross-tenant access policy changes affect who can authenticate across tenants — review the proposed JSON carefully and test in a non-production tenant first.
