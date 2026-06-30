#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Teams, Microsoft.Graph.Users, Microsoft.Graph.Groups

<#
.SYNOPSIS
    Removes specified users from all Microsoft Teams (member or owner).
    Produces audit CSV/JSON before changes, and a results CSV after.

.PARAMETER CsvPath        Path to input CSV with a 'UserEmail' column.
.PARAMETER ServiceAccount UPN of break-glass account to add as owner when target
                          user is the sole owner of a Team.
.PARAMETER AuditDir       Folder for output files (defaults to script folder).
.PARAMETER TenantId       From Register-AppAndPermissions output.
.PARAMETER ClientId       From Register-AppAndPermissions output.
.PARAMETER Thumbprint     From Register-AppAndPermissions output.
.PARAMETER WhatIf         Discover memberships and write audit files — NO changes made.

.EXAMPLE
    # Dry-run — always do this first
    .\Remove-TeamMemberships.ps1 -CsvPath .\users.csv -ServiceAccount "admin@M365x6433245.onmicrosoft.com" `
        -TenantId "fbbda3de-75be-40ae-8319-8738189a37e5" `
        -ClientId "3756b518-287a-408a-8e22-befcafdbed63" `
        -Thumbprint "8E6DF2A647740" -WhatIf

.EXAMPLE
    # Live run
    .\Remove-TeamMemberships.ps1 -CsvPath .\users.csv -ServiceAccount "admin@M365x6433245.onmicrosoft.com" `
        -TenantId "fbbda3de-75be-40ae-8319-8738189a37e5" `
        -ClientId "3756b518-287a-408a-8e22-befcafdbed63" `
        -Thumbprint "8E6DF2A64774"
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)][string] $CsvPath,
    [Parameter(Mandatory)][string] $ServiceAccount,
    [string] $AuditDir = $PSScriptRoot,
    [Parameter(Mandatory)][string] $TenantId,
    [Parameter(Mandatory)][string] $ClientId,
    [Parameter(Mandatory)][string] $Thumbprint
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── HELPERS ──────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Message"
    switch ($Level) {
        "INFO"    { Write-Host $line -ForegroundColor Cyan   }
        "WARN"    { Write-Host $line -ForegroundColor Yellow }
        "ERROR"   { Write-Host $line -ForegroundColor Red    }
        "SUCCESS" { Write-Host $line -ForegroundColor Green  }
        default   { Write-Host $line }
    }
    # Always write log to disk regardless of WhatIf
    Add-Content -Path $RunLogPath -Value $line -WhatIf:$false
}

function Get-GraphUser {
    param([string]$Upn)
    try   { Get-MgUser -UserId $Upn -ErrorAction Stop }
    catch { $null }
}

function Get-TeamMembershipEntry {
    # Returns the full membership object (or $null) for a userId in a teamId
    param([string]$TeamId, [string]$UserId)
    try {
        $members = @(Get-MgTeamMember -TeamId $TeamId -All -ErrorAction Stop)
        return $members | Where-Object {
            $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.aadUserConversationMember' -and
            $_.AdditionalProperties.userId -eq $UserId
        } | Select-Object -First 1
    } catch { return $null }
}

function Get-TeamOwnerCount {
    param([string]$TeamId)
    try {
        $members = @(Get-MgTeamMember -TeamId $TeamId -All -ErrorAction Stop)
        return @($members | Where-Object { $_.Roles -contains "owner" }).Count
    } catch { return 0 }
}

function Add-TeamOwnerByUserId {
    param([string]$TeamId, [string]$UserId)
    $body = @{
        "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
        roles             = @("owner")
        "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$UserId')"
    }
    New-MgTeamMember -TeamId $TeamId -BodyParameter $body | Out-Null
}

# ── INIT ─────────────────────────────────────────────────────────────────────

$Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$RunMode    = if ($WhatIfPreference) { "WHATIF" } else { "LIVE" }

# Always write these to disk regardless of WhatIf
New-Item -ItemType Directory -Path $AuditDir -Force -WhatIf:$false | Out-Null

$RunLogPath    = Join-Path $AuditDir "\reports\RunLog_${Timestamp}_${RunMode}.txt"
$PreAuditCsv   = Join-Path $AuditDir "\reports\Audit_PRE_${Timestamp}.csv"
$PostResultCsv = Join-Path $AuditDir "\reports\REMOVE_POST_${Timestamp}.csv"

Write-Log "═══════════════════════════════════════════════════════"
Write-Log " Teams Membership Removal  |  Run: $Timestamp  |  Mode: $RunMode"
Write-Log "═══════════════════════════════════════════════════════"

# ── AUTHENTICATE ─────────────────────────────────────────────────────────────
Write-Log "Authenticating to Microsoft Graph (certificate)..."
Connect-MgGraph -TenantId $TenantId -ClientId $ClientId `
                -CertificateThumbprint $Thumbprint -NoWelcome
Write-Log "Connected." "SUCCESS"

# ── READ CSV ─────────────────────────────────────────────────────────────────
Write-Log "Reading CSV: $CsvPath"
$TargetUsers = Import-Csv -Path $CsvPath
if (-not ($TargetUsers | Get-Member -Name 'UserEmail' -ErrorAction SilentlyContinue)) {
    Write-Log "CSV must contain a 'UserEmail' column." "ERROR"; exit 1
}
Write-Log "Target users in CSV: $($TargetUsers.Count)"

# ── RESOLVE USERS ─────────────────────────────────────────────────────────────
Write-Log "Resolving user objects in Azure AD..."
$ResolvedUsers = [System.Collections.Generic.List[hashtable]]::new()
foreach ($row in $TargetUsers) {
    $upn  = $row.UserEmail.Trim()
    $user = Get-GraphUser -Upn $upn
    if ($user) {
        Write-Log "  Resolved: $upn  →  $($user.Id)"
        $ResolvedUsers.Add(@{ Upn = $upn; Id = $user.Id; DisplayName = $user.DisplayName })
    } else {
        Write-Log "  NOT FOUND in Azure AD: $upn  — skipping." "WARN"
    }
}

if ($ResolvedUsers.Count -eq 0) {
    Write-Log "No users could be resolved. Check that the UPNs match this tenant. Exiting." "ERROR"
    Disconnect-MgGraph | Out-Null; exit 1
}

# ── RESOLVE SERVICE ACCOUNT ──────────────────────────────────────────────────
Write-Log "Resolving service account: $ServiceAccount"
$SvcUser = Get-GraphUser -Upn $ServiceAccount
if (-not $SvcUser) {
    Write-Log "Service account '$ServiceAccount' not found in Azure AD." "ERROR"
    Disconnect-MgGraph | Out-Null; exit 1
}
Write-Log "  Service account ID: $($SvcUser.Id)" "SUCCESS"

# ── ENUMERATE ALL TEAMS ──────────────────────────────────────────────────────
Write-Log "Enumerating all Teams in tenant..."
$AllTeams = @(Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" `
                           -Property "Id,DisplayName" -All)
Write-Log "  Total Teams found: $($AllTeams.Count)"

# ── DISCOVER MEMBERSHIPS ──────────────────────────────────────────────────────
Write-Log "Scanning each Team for target users (this may take a few minutes)..."
$AuditRecords = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($userEntry in $ResolvedUsers) {
    Write-Log "─── Checking Teams for $($userEntry.Upn) ───"

    foreach ($team in $AllTeams) {
        $entry = Get-TeamMembershipEntry -TeamId $team.Id -UserId $userEntry.Id
        if (-not $entry) { continue }

        $role        = if ($entry.Roles -contains "owner") { "Owner" } else { "Member" }
        $ownerCount  = Get-TeamOwnerCount -TeamId $team.Id
        $isSoleOwner = ($role -eq "Owner") -and ($ownerCount -eq 1)

        $record = [PSCustomObject]@{
            UserUpn          = $userEntry.Upn
            UserDisplayName  = $userEntry.DisplayName
            UserId           = $userEntry.Id
            TeamId           = $team.Id
            TeamName         = $team.DisplayName
            RoleInTeam       = $role
            TotalOwners      = $ownerCount
            IsSoleOwner      = $isSoleOwner
            MembershipId     = $entry.Id
            # Populated during live run:
            ServiceAcctAdded = $false
            Removed          = $false
            Result           = "Pending"
            ErrorDetail      = ""
            ProcessedAt      = ""
        }
        $AuditRecords.Add($record)
        Write-Log ("  [{0}] {1}  (owners:{2}  soleOwner:{3})" -f $role, $team.DisplayName, $ownerCount, $isSoleOwner)
    }
}

Write-Log "Discovery complete. Total membership records: $($AuditRecords.Count)"

# ── WRITE PRE-CHANGE AUDIT (always written, even in WhatIf) ──────────────────
Write-Log "Writing pre-change audit file..."
$AuditRecords | Select-Object UserUpn,UserDisplayName,UserId,TeamName,TeamId,RoleInTeam,TotalOwners,IsSoleOwner,MembershipId |
    Export-Csv -Path $PreAuditCsv -NoTypeInformation -Encoding utf8 -WhatIf:$false
Write-Log "  PRE audit CSV : $PreAuditCsv" "SUCCESS"

# ── WHATIF EXIT ───────────────────────────────────────────────────────────────
if ($WhatIfPreference) {
    Write-Log ""
    Write-Log "══ WhatIf Summary (no changes made) ══════════════════" "WARN"
    Write-Log "  Users resolved        : $($ResolvedUsers.Count)" "WARN"
    Write-Log "  Memberships found     : $($AuditRecords.Count)" "WARN"
    #Write-Log "  Sole-owner situations : $(@($AuditRecords | Where-Object { $_.IsSoleOwner -eq $true }).Count)" "WARN"
    Write-Log "  Sole-owner situations : $(@($AuditRecords | Where-Object IsSoleOwner).Count)" "WARN"
    Write-Log "  PRE audit CSV written : $PreAuditCsv" "WARN"
    Write-Log "  Review the CSV then re-run WITHOUT -WhatIf to apply changes." "WARN"
    Write-Log "══════════════════════════════════════════════════════" "WARN"
    Disconnect-MgGraph | Out-Null
    exit 0
}

# ── LIVE RUN — APPLY CHANGES ──────────────────────────────────────────────────
Write-Log "═══ Starting live membership removal ═══"

foreach ($rec in $AuditRecords) {
    $rec.ProcessedAt = (Get-Date -Format "o")
    Write-Log "Processing: $($rec.UserUpn)  →  $($rec.TeamName)  [$($rec.RoleInTeam)]"
    
    try {
        # Step 1 — sole owner guard
        if ($rec.IsSoleOwner) {
            Write-Log "  Sole owner — checking if service account already present..."
            $svcEntry = Get-TeamMembershipEntry -TeamId $rec.TeamId -UserId $SvcUser.Id
            if ($svcEntry -and ($svcEntry.Roles -contains "owner")) {
                Write-Log "  Service account already an owner — skipping add." "WARN"
            } else {
                Write-Log "  Adding service account as owner..."
                Add-TeamOwnerByUserId -TeamId $rec.TeamId -UserId $SvcUser.Id
                $rec.ServiceAcctAdded = $true
                Write-Log "  Service account added as owner. Verifying replication..." "SUCCESS"

                # Poll until the new owner role is actually readable, since Graph can
                # lag a beat before the write propagates tenant-wide. Without this,
                # the delete below can fail with:
                # "Remove member not allowed, member is the last owner of the team."
                $verified = $false
                for ($i = 1; $i -le 10; $i++) {
                    Start-Sleep -Milliseconds 1000
                    $check = Get-TeamMembershipEntry -TeamId $rec.TeamId -UserId $SvcUser.Id
                    if ($check -and ($check.Roles -contains "owner")) {
                        $verified = $true
                        Write-Log "  Service account owner role verified (attempt $i)." "SUCCESS"
                        break
                    }
                    Write-Log "  Owner role not yet visible — retrying ($i/10)..." "WARN"
                }
                if (-not $verified) {
                    Write-Log "  Could not verify service account owner role after 10 attempts — skipping removal for this Team." "ERROR"
                    $rec.Result      = "Failed"
                    $rec.ErrorDetail = "Service account owner promotion did not replicate in time."
                    continue
                }
            }
        }

        # Step 2 — refresh membership ID in case it changed
        $freshEntry = Get-TeamMembershipEntry -TeamId $rec.TeamId -UserId $rec.UserId
        if (-not $freshEntry) {
            Write-Log "  User no longer a member (already removed?)." "WARN"
            $rec.Removed = $true; $rec.Result = "AlreadyGone"
            continue
        }

        # Step 3 — remove, with retry on the transient "last owner" race condition
        $removeAttempts = 0
        $removed = $false
        do {
            $removeAttempts++
            try {
                Remove-MgTeamMember -TeamId $rec.TeamId -ConversationMemberId $freshEntry.Id
                $removed = $true
            } catch {
                if ($_.Exception.Message -match "last owner" -and $removeAttempts -lt 5) {
                    Write-Log "  Transient last-owner error — waiting and retrying ($removeAttempts/5)..." "WARN"
                    Start-Sleep -Milliseconds 1500
                } else {
                    throw
                }
            }
        } while (-not $removed -and $removeAttempts -lt 5)

        $rec.Removed = $true
        $rec.Result  = "Removed"
        Write-Log "  Removed successfully." "SUCCESS"
    }
    catch {
        $rec.Result      = "Failed"
        $rec.ErrorDetail = $_.Exception.Message
        Write-Log "  ERROR: $($_.Exception.Message)" "ERROR"
    }

    Start-Sleep -Milliseconds 300
}

# ── WRITE POST-CHANGE RESULTS CSV ────────────────────────────────────────────
Write-Log "Writing post-change results..."
$AuditRecords |
    Select-Object UserUpn, UserDisplayName, TeamName, RoleInTeam, IsSoleOwner,
                  ServiceAcctAdded, Removed, Result, ErrorDetail, ProcessedAt, TeamId, UserId |
    Export-Csv -Path $PostResultCsv -NoTypeInformation -Encoding utf8 -WhatIf:$false
Write-Log "  POST results CSV : $PostResultCsv" "SUCCESS"

$total    = @($AuditRecords).Count
$removed  = @($AuditRecords | Where-Object { $_.Result -eq "Removed" }).Count
$svcAdded = @($AuditRecords | Where-Object { $_.ServiceAcctAdded }).Count
$failed   = @($AuditRecords | Where-Object { $_.Result -eq "Failed" }).Count
$gone     = @($AuditRecords | Where-Object { $_.Result -eq "AlreadyGone" }).Count

Write-Log "═══════════════════════════════════════════════════════"
Write-Log "FINAL SUMMARY"
Write-Log "  Total memberships processed  : $($total)"
Write-Log "  Removed successfully         : $($removed)"
Write-Log "  Already gone (skipped)       : $($gone)"
Write-Log "  Service account promoted     : $($svcAdded)"
Write-Log "  Failed                       : $($failed)"
Write-Log "  PRE  audit  : $PreAuditCsv"
Write-Log "  POST results: $PostResultCsv"
Write-Log "  Run log     : $RunLogPath"
Write-Log "═══════════════════════════════════════════════════════"

Disconnect-MgGraph | Out-Null
Write-Log "Done." "SUCCESS"
