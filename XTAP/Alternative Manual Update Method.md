# Manage XTAP Partner Policy App List with Graph Explorer

Alternate, no-script steps to review an **inbound** or **outbound** cross-tenant access policy (XTAP) partner configuration and **add or remove an application ID** from its `applications` allow/block list — using **Graph Explorer** (<https://aka.ms/ge>).

Every step below includes a simple example. The examples use:

| Placeholder | Example value |
|---|---|
| Partner tenant ID | `0db7a84e-fa5e-4c85-bbc6-262831a58131` |
| Application (App) ID to add/remove | `75f31797-37c9-498e-8dc9-53c16a36afca` (Microsoft Planner Client) |

> **Inbound vs outbound ownership:** *Inbound* (`b2bCollaborationInbound`) is owned by the **resource** tenant (which of your apps external users can access). *Outbound* (`b2bCollaborationOutbound`) is owned by **your (home)** tenant (which partner apps your users can access). Sign in to the tenant that owns the direction you want to edit.

---

## Step 1 — Open Graph Explorer and sign in

1. Go to <https://aka.ms/ge>.
2. Click **Sign in to Graph Explorer** and authenticate as an admin.
3. Consent to the delegated permission **`Policy.ReadWrite.CrossTenantAccess`**:
   - Click the **gear / Settings → Select permissions**, search for `Policy.ReadWrite.CrossTenantAccess`, add it, and **Consent**.

> A **Security Administrator** (or equivalent) role is required to write cross-tenant access policy.

---

## Step 2 — List all partner policies

Find the partners that have an explicit configuration.

**Request**
```http
GET https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners
```

**Example response (trimmed)**
```json
{
  "value": [
    { "tenantId": "0db7a84e-fa5e-4c85-bbc6-262831a58131" },
    { "tenantId": "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" }
  ]
}
```

Pick the `tenantId` you want to edit.

   <img width="520" height="483" alt="chrome-capture-2026-08-20" src="https://github.com/user-attachments/assets/4d04cada-762c-4c55-986d-0f577250adc1" />


---

## Step 3 — Review the current policy for one partner

**Request**
```http
GET https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners/0db7a84e-fa5e-4c85-bbc6-262831a58131
```

**Example response (trimmed)**
```json
{
  "tenantId": "0db7a84e-fa5e-4c85-bbc6-262831a58131",
  "b2bCollaborationInbound": {
    "usersAndGroups": {
      "accessType": "allowed",
      "targets": [ { "target": "AllUsers", "targetType": "user" } ]
    },
    "applications": {
      "accessType": "allowed",
      "targets": [
        { "target": "00000003-0000-0ff1-ce00-000000000000", "targetType": "application" }
      ]
    }
  },
  "b2bCollaborationOutbound": null
}
```

   <img width="520" height="483" alt="getpartner" src="https://github.com/user-attachments/assets/f746e50d-ed86-489f-bbf2-a26ecff0df08" />


Note two things about the direction you plan to edit (`b2bCollaborationInbound` or `b2bCollaborationOutbound`):

- **`applications.accessType`** — `allowed` = allow-list (only listed apps permitted); `blocked` = block-list (listed apps denied). A partner has **one** applications list with **one** `accessType` — you cannot keep both an allow-list and a block-list for the same partner.
- **`applications.targets`** — the current app entries. **Copy this array** — you'll resend it (plus or minus one app) in the next step.

> If the direction is `null`, there is no explicit configuration yet; the partner inherits the `default` policy. Editing it (Step 4/5) creates an explicit override — include a `usersAndGroups` block too.

   <img width="530" height="482" alt="copypolicy" src="https://github.com/user-attachments/assets/f6b71a9d-dff2-45c2-a216-d0d56c86786d" />


---

## Step 4 — Add an app to the list

`PATCH` **replaces the entire** `b2bCollaborationInbound` / `b2bCollaborationOutbound` object, so send the **full** object: `usersAndGroups` **and** the complete `applications.targets` list = your existing targets **plus** the new app.

**Request (inbound example — adding Planner)**
```http
PATCH https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners/0db7a84e-fa5e-4c85-bbc6-262831a58131
Content-Type: application/json
```
```json
{
  "b2bCollaborationInbound": {
    "usersAndGroups": {
      "accessType": "allowed",
      "targets": [ { "target": "AllUsers", "targetType": "user" } ]
    },
    "applications": {
      "accessType": "allowed",
      "targets": [
        { "target": "00000003-0000-0ff1-ce00-000000000000", "targetType": "application" },
        { "target": "75f31797-37c9-498e-8dc9-53c16a36afca", "targetType": "application" }
      ]
    }
  }
}
```

A successful `PATCH` returns **`204 No Content`**.

> For **outbound**, use `"b2bCollaborationOutbound"` as the property name instead — the body shape is identical.

---

## Step 5 — Remove an app from the list

Same as Step 4, but resend the `applications.targets` list **without** the app you want to remove.

**Request (inbound example — removing Planner)**
```http
PATCH https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners/0db7a84e-fa5e-4c85-bbc6-262831a58131
Content-Type: application/json
```
```json
{
  "b2bCollaborationInbound": {
    "usersAndGroups": {
      "accessType": "allowed",
      "targets": [ { "target": "AllUsers", "targetType": "user" } ]
    },
    "applications": {
      "accessType": "allowed",
      "targets": [
        { "target": "00000003-0000-0ff1-ce00-000000000000", "targetType": "application" }
      ]
    }
  }
}
```

Returns **`204 No Content`** on success.

---

## Step 6 — Verify the result

Re-run the Step 3 `GET` and confirm the `applications.targets` list now reflects your change.

**Request**
```http
GET https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners/0db7a84e-fa5e-4c85-bbc6-262831a58131

```

**Example response (after add — trimmed)**
```json
{
  "b2bCollaborationInbound": {
    "applications": {
      "accessType": "allowed",
      "targets": [
        { "target": "00000003-0000-0ff1-ce00-000000000000", "targetType": "application" },
        { "target": "75f31797-37c9-498e-8dc9-53c16a36afca", "targetType": "application" }
      ]
    }
  }
}
```

---

## Optional — Back up before you change

Before editing, run the Step 3 `GET`, click **Response preview**, and copy the JSON to a `.json` file. To restore, `PATCH` the saved `b2bCollaborationInbound` / `b2bCollaborationOutbound` object back.

---

## Reference

**Target types** used in `targets[]`:

| `targetType` | Example `target` | Meaning |
|---|---|---|
| `user` | `AllUsers` | All external users (or a specific user objectId) |
| `group` | a group objectId | A specific group |
| `application` | an app **appId** (GUID) | A specific app, or `AllApplications` for all |

**Tips & gotchas**
- `PATCH` replaces the whole direction object — always include `usersAndGroups` and the **complete** `applications.targets` list, not just the delta.
- Other properties (`b2bDirectConnectInbound`/`Outbound`, `automaticUserConsentSettings`, `inboundTrust`) are separate and are left untouched as long as you don't include them.
- The applications list is matched against **both** the client app and the resource app in a cross-tenant token request — a match on either satisfies the rule.
- Portal can't always find first-party **client** apps in its picker; Graph accepts the raw appId regardless. It may then display as a bare GUID — expected.
- Endpoint shown is **public cloud** (`graph.microsoft.com`) and API **v1.0**. For national clouds, use the corresponding Graph endpoint.
