# Remove-TeamMemberships.ps1

To execute the script, create an empty **reports** folder in the same directory. The output files will be stored in this folder.

Bulk-removes a list of users from every Microsoft Teams team they belong to (as Member or Owner) using Microsoft Graph, with a full pre-change audit trail and post-change results log. Protects any Team from being left ownerless by promoting a designated service account before removing a sole owner.

---

## What it does

1. Authenticates to Microsoft Graph via certificate (app-only, no client secret).
2. Reads a CSV of target users (`UserEmail` column).
3. Resolves each user against Azure AD; skips and logs any UPN not found in the tenant.
4. Enumerates **every** Team in the tenant.
5. Scans each Team for each target user's membership (Member or Owner) and current owner count.
6. Writes a **pre-change audit CSV** before any modification — this is your source of truth for review and rollback.
7. **Live run only:** for each membership found —
   - If the user is the **sole owner** of a Team, adds a designated service account as owner first, verifies the promotion has replicated, then proceeds.
   - Removes the user from the Team.
   - Retries automatically on the transient `"last owner"` Graph error that can occur immediately after a sole-owner promotion.
8. Writes a **post-change results CSV** with per-row outcome (`Removed` / `Failed` / `AlreadyGone`).
9. Supports `-WhatIf` for a complete dry-run: discovery and the pre-change audit CSV are produced, but no Graph writes occur.

---

## Prerequisites

- PowerShell 7+ (Windows PowerShell 5.1 also works)
- Modules:
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser -Force
  ```
- An Azure AD App Registration with a certificate credential and the following Graph **Application** permissions, consented by an admin:

  | Permission | Why it's needed |
  |---|---|
  | `Team.ReadBasic.All` | Enumerate all Teams in the tenant |
  | `TeamMember.Read.All` | Read member/owner lists per Team |
  | `TeamMember.ReadWrite.All` | Add/remove members and owners |
  | `Group.Read.All` | Read group metadata backing each Team |
  | `Group.ReadWrite.All` | Required for owner manipulation on the underlying group |
  | `User.Read.All` | Resolve UPNs to Azure AD object IDs |

- The certificate's thumbprint must be installed in the **current user's certificate store** (`Cert:\CurrentUser\My`) on the machine running the script.

---

## Parameters

| Parameter | Required | Description |
|---|---|---|
| `-CsvPath` | Yes | Path to input CSV with a `UserEmail` column |
| `-ServiceAccount` | Yes | UPN of the break-glass/service account promoted to owner when a target user is a Team's sole owner |
| `-AuditDir` | No | Output folder for logs and CSVs (defaults to script folder) |
| `-TenantId` | Yes | Azure AD tenant ID |
| `-ClientId` | Yes | App registration's Application (client) ID |
| `-Thumbprint` | Yes | Certificate thumbprint used for authentication |
| `-WhatIf` | No | Dry-run — discovers and writes the pre-change audit CSV, makes no changes |

---

## Usage

### Dry-run (always do this first)

```powershell
.\Remove-TeamMemberships.ps1 `
    -CsvPath        ".\users.csv" `
    -ServiceAccount "admin@yourtenant.onmicrosoft.com" `
    -TenantId       "<tenant-id>" `
    -ClientId       "<client-id>" `
    -Thumbprint     "<cert-thumbprint>" `
    -WhatIf
```

Review the resulting `Audit_PRE_<timestamp>.csv` before proceeding.

### Live run

Drop `-WhatIf`:

```powershell
.\Remove-TeamMemberships.ps1 `
    -CsvPath        ".\users.csv" `
    -ServiceAccount "admin@yourtenant.onmicrosoft.com" `
    -TenantId       "<tenant-id>" `
    -ClientId       "<client-id>" `
    -Thumbprint     "<cert-thumbprint>"
```

---

## Input CSV format

```csv
UserEmail
AdeleV@yourtenant.onmicrosoft.com
AllanD@yourtenant.onmicrosoft.com
```

Single column, header must be exactly `UserEmail`. Users not found in the target tenant are logged and skipped — they don't block processing of the rest of the batch.

---

## Output files

All written to `-AuditDir` (defaults to the script's own folder):

| File | When | Contents |
|---|---|---|
| `Audit_PRE_<timestamp>.csv` | WhatIf and live | Every membership found before any change: user, team, role, owner count, sole-owner flag, membership ID |
| `Audit_POST_<timestamp>.csv` | Live only | Same rows plus `Result` (`Removed` / `Failed` / `AlreadyGone`), `ServiceAcctAdded`, `ErrorDetail`, `ProcessedAt` |
| `RunLog_<timestamp>_<MODE>.txt` | Always | Full console log written to disk, regardless of `-WhatIf` |

`Audit_PRE_<timestamp>.csv` is the file to keep for audit and rollback purposes — it has everything needed to reconstruct original memberships.

---

## How sole-owner protection works

For each Team where a target user is found:

1. The script counts current owners of that Team.
2. If the user is an **Owner** and is the **only** owner, the configured `-ServiceAccount` is added as an owner before the user is removed.
3. The script then **polls Microsoft Graph** (up to 10 attempts, 1 second apart) to confirm the new owner role has actually replicated and is readable, since Graph can lag briefly after a write.
4. If verification fails after all attempts, that row is marked `Failed` and is **not** removed — the Team is left untouched rather than risking an ownerless state.
5. Once verified, the removal proceeds. If Graph still returns a transient `"last owner"` error on delete (a known race condition even after verification), the script retries up to 5 times with a 1.5 second backoff before giving up.
6. If the service account is already an owner of the Team, the promotion step is skipped entirely.

This guarantees no Team is ever left without an owner as a result of running this script.

---

## Notes & known behaviors

- **Tenant matching matters.** Target users must exist in the same tenant the app is registered in. A UPN from a different tenant will log as `NOT FOUND` and be skipped.
- **Idempotency.** Re-running against users who have already been removed simply finds no remaining memberships for them — safe to re-run.
- **Rate limiting.** Short delays are built in between Graph calls to avoid throttling on large batches; expect roughly 2–4 seconds per membership processed.
- **CSV over JSON.** The audit format is CSV (not JSON) so it stays reviewable in Excel and scales cleanly to tenants with thousands of Teams.

---

## Disclaimer

This script makes real changes to Microsoft Teams membership and group ownership. Always run with `-WhatIf` first, review the generated `Audit_PRE_*.csv`, and validate in a non-production tenant before running against production.
