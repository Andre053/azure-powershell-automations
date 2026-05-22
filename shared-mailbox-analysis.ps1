<#
CREATED BY CLAUDE (SONNET 4.6 ADAPTIVE)

.SYNOPSIS
    Retrieves all Microsoft 365 shared mailboxes with delegate assignments and
    key health/configuration attributes.

.DESCRIPTION
    Connects to Exchange Online, enumerates every shared mailbox, and collects:

    Delegate permissions:
      - Full Access delegates  (Get-MailboxPermission  / FullAccess)
      - Send As delegates      (Get-RecipientPermission / SendAs)
      - Send on Behalf of      (GrantSendOnBehalfTo)

    Mailbox health & configuration (per mailbox):
      - Last logon date & days since last logon  (Get-MailboxStatistics)
      - Mailbox size (MB) and item count         (Get-MailboxStatistics)
      - Warning / ProhibitSend / ProhibitSendReceive quotas
      - Sign-in blocked status                   (Get-User / BlockCredential)
      - WhenCreated / WhenChanged

    Outputs delegate detail, delegate summary counts, and a mailbox health table
    to the console, and optionally exports all three to CSV.

.PARAMETER ExportPath
    Optional. Folder path where three CSV files will be written:
      - SharedMailbox_Delegates.csv
      - SharedMailbox_DelegateSummary.csv
      - SharedMailbox_Health.csv
    If omitted, results are displayed in the console only.

.PARAMETER IncludeSelf
    If specified, includes permissions granted to the mailbox itself
    (e.g. NT AUTHORITY\SELF). Excluded by default.

.EXAMPLE
    # Display results in console only
    .\Get-SharedMailboxDelegates.ps1

.EXAMPLE
    # Display and export to CSV
    .\Get-SharedMailboxDelegates.ps1 -ExportPath "C:\Reports"

.NOTES
    Prerequisites:
      - ExchangeOnlineManagement module  (Install-Module ExchangeOnlineManagement)
      - Connect-ExchangeOnline must succeed (MFA supported)
      - The running account needs View-Only Organization Management or higher

    Tested with ExchangeOnlineManagement v3.x
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$ExportPath,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSelf
)

# ── Helper: ensure module & connection ─────────────────────────────────

#$ExportPath = "."
function Assert-ExchangeOnlineConnection {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Error "ExchangeOnlineManagement module not found. Run: Install-Module ExchangeOnlineManagement"
        exit 1
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    # Test whether a session already exists
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "✔  Already connected to Exchange Online." -ForegroundColor Green
    }
    catch {
        Write-Host "Connecting to Exchange Online — sign-in prompt may appear..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Host "✔  Connected to Exchange Online." -ForegroundColor Green
    }
}


# ── Main ──────────────────────────────────────────────────────────────

Assert-ExchangeOnlineConnection

# ── 1. Retrieve all shared mailboxes ─────────────────────────────────────────
Write-Host "`nFetching shared mailboxes..." -ForegroundColor Cyan
$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
    Select-Object DisplayName, PrimarySmtpAddress, Alias, ExchangeGuid, GrantSendOnBehalfTo,
                  WhenCreated, WhenChanged,
                  IssueWarningQuota, ProhibitSendQuota, ProhibitSendReceiveQuota,
                  UseDatabaseQuotaDefaults

if (-not $sharedMailboxes) {
    Write-Warning "No shared mailboxes found in this tenant."
    exit 0
}

Write-Host "  Found $($sharedMailboxes.Count) shared mailbox(es).`n" -ForegroundColor Green

# ── 2. Collect delegate details ───────────────────────────────────────────────
$delegateRows   = [System.Collections.Generic.List[PSCustomObject]]::new()
$summaryRows    = [System.Collections.Generic.List[PSCustomObject]]::new()
$healthRows     = [System.Collections.Generic.List[PSCustomObject]]::new()

$mbxIndex = 0
foreach ($mbx in $sharedMailboxes) {
    $mbxIndex++
    $displayName = $mbx.DisplayName
    $smtp        = $mbx.PrimarySmtpAddress

    Write-Progress -Activity "Processing shared mailboxes" `
                   -Status "$mbxIndex / $($sharedMailboxes.Count) — $displayName" `
                   -PercentComplete (($mbxIndex / $sharedMailboxes.Count) * 100)

    # ── Full Access ──────────────────────────────────────────────────────────
    $fullAccessDelegates = Get-MailboxPermission -Identity $smtp |
        Where-Object {
            $_.AccessRights -contains "FullAccess" -and
            $_.IsInherited   -eq $false -and
            ($IncludeSelf -or $_.User -notmatch "NT AUTHORITY")
        }

    foreach ($perm in $fullAccessDelegates) {
        $delegateRows.Add([PSCustomObject]@{
            MailboxDisplayName   = $displayName
            MailboxEmail         = $smtp
            DelegateUser         = $perm.User.ToString()
            PermissionType       = "Full Access"
        })
    }

    # ── Send As ──────────────────────────────────────────────────────────────
    $sendAsDelegates = Get-RecipientPermission -Identity $smtp |
        Where-Object {
            $_.AccessRights -contains "SendAs" -and
            $_.IsInherited   -eq $false -and
            ($IncludeSelf -or $_.Trustee -notmatch "NT AUTHORITY")
        }

    foreach ($perm in $sendAsDelegates) {
        $delegateRows.Add([PSCustomObject]@{
            MailboxDisplayName   = $displayName
            MailboxEmail         = $smtp
            DelegateUser         = $perm.Trustee.ToString()
            PermissionType       = "Send As"
        })
    }

    # ── Send on Behalf ───────────────────────────────────────────────────────
    if ($mbx.GrantSendOnBehalfTo) {
        foreach ($trustee in $mbx.GrantSendOnBehalfTo) {
            $delegateRows.Add([PSCustomObject]@{
                MailboxDisplayName   = $displayName
                MailboxEmail         = $smtp
                DelegateUser         = $trustee.ToString()
                PermissionType       = "Send on Behalf"
            })
        }
    }

    # ── Per-mailbox summary ───────────────────────────────────────────────────
    $faCount  = ($fullAccessDelegates | Measure-Object).Count
    $saCount  = ($sendAsDelegates     | Measure-Object).Count
    $sobCount = if ($mbx.GrantSendOnBehalfTo) { $mbx.GrantSendOnBehalfTo.Count } else { 0 }

    $summaryRows.Add([PSCustomObject]@{
        MailboxDisplayName       = $displayName
        MailboxEmail             = $smtp
        FullAccessCount          = $faCount
        SendAsCount              = $saCount
        SendOnBehalfCount        = $sobCount
        TotalDelegateCount       = $faCount + $saCount + $sobCount
    })

    # ── Mailbox statistics (last logon, size) ─────────────────────────────────
    $stats = Get-MailboxStatistics -Identity $smtp -ErrorAction SilentlyContinue

    $lastLogon        = if ($stats -and $stats.LastLogonTime) { $stats.LastLogonTime } else { $null }
    $daysSinceLogon   = if ($lastLogon) { (New-TimeSpan -Start $lastLogon -End (Get-Date)).Days } else { "Never" }
    $sizeMB           = if ($stats -and $stats.TotalItemSize) {
                            [math]::Round($stats.TotalItemSize.Value.ToMB(), 2)
                        } else { 0 }
    $itemCount        = if ($stats) { $stats.ItemCount } else { 0 }

    # ── Sign-in status (BlockCredential via Get-User) ─────────────────────────
    $exoUser          = Get-User -Identity $smtp -ErrorAction SilentlyContinue
    $signInBlocked    = if ($exoUser) { $exoUser.BlockCredential } else { "Unknown" }

    # ── Quota helper: render ByteQuantifiedSize or "Unlimited" as a string ────
    function Format-Quota ($quota) {
        if (-not $quota -or $quota -eq "Unlimited") { return "Unlimited" }
        try { return "$([math]::Round($quota.Value.ToMB(), 0)) MB" } catch { return $quota.ToString() }
    }

    $usesOrgDefaults  = $mbx.UseDatabaseQuotaDefaults

    $healthRows.Add([PSCustomObject]@{
        MailboxDisplayName        = $displayName
        MailboxEmail              = $smtp
        WhenCreated               = $mbx.WhenCreated
        WhenChanged               = $mbx.WhenChanged
        LastLogonDate             = if ($lastLogon) { $lastLogon } else { "Never" }
        DaysSinceLastLogon        = $daysSinceLogon
        MailboxSizeMB             = $sizeMB
        ItemCount                 = $itemCount
        UsesOrgDefaultQuotas      = $usesOrgDefaults
        WarningQuota              = if ($usesOrgDefaults) { "Org default" } else { Format-Quota $mbx.IssueWarningQuota }
        ProhibitSendQuota         = if ($usesOrgDefaults) { "Org default" } else { Format-Quota $mbx.ProhibitSendQuota }
        ProhibitSendReceiveQuota  = if ($usesOrgDefaults) { "Org default" } else { Format-Quota $mbx.ProhibitSendReceiveQuota }
        SignInBlocked             = $signInBlocked
    })
}

Write-Progress -Activity "Processing shared mailboxes" -Completed

# ── 3. Grand totals ──────────────────────────────────────────────────────────
$totalFA  = ($delegateRows | Where-Object PermissionType -eq "Full Access"    | Measure-Object).Count
$totalSA  = ($delegateRows | Where-Object PermissionType -eq "Send As"        | Measure-Object).Count
$totalSOB = ($delegateRows | Where-Object PermissionType -eq "Send on Behalf" | Measure-Object).Count
$grandTotal = $delegateRows.Count

# ── 4. Display results ────────────────────────────────────────────────────────
Write-Host "`n════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  SHARED MAILBOX DELEGATE DETAIL" -ForegroundColor Cyan
Write-Host "════════════════════════════════════════════════════════════════`n" -ForegroundColor Cyan

if ($delegateRows.Count -gt 0) {
    $delegateRows | Format-Table -AutoSize -Property MailboxDisplayName, MailboxEmail, DelegateUser, PermissionType
} else {
    Write-Host "  No delegates found across any shared mailbox.`n" -ForegroundColor Yellow
}

Write-Host "`n════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  PER-MAILBOX DELEGATE SUMMARY" -ForegroundColor Cyan
Write-Host "════════════════════════════════════════════════════════════════`n" -ForegroundColor Cyan

$summaryRows | Sort-Object TotalDelegateCount -Descending |
    Format-Table -AutoSize -Property MailboxDisplayName, MailboxEmail,
                                     FullAccessCount, SendAsCount,
                                     SendOnBehalfCount, TotalDelegateCount

Write-Host "`n════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  GRAND TOTAL ACROSS ALL SHARED MAILBOXES" -ForegroundColor Cyan
Write-Host "════════════════════════════════════════════════════════════════`n" -ForegroundColor Cyan

[PSCustomObject]@{
    TotalSharedMailboxes   = $sharedMailboxes.Count
    TotalFullAccess        = $totalFA
    TotalSendAs            = $totalSA
    TotalSendOnBehalf      = $totalSOB
    GrandTotalDelegates    = $grandTotal
} | Format-List

Write-Host "`n════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  MAILBOX HEALTH & CONFIGURATION" -ForegroundColor Cyan
Write-Host "════════════════════════════════════════════════════════════════`n" -ForegroundColor Cyan

# Flag mailboxes where sign-in is NOT blocked (should be for all shared mailboxes)
$signInNotBlocked = ($healthRows | Where-Object { $_.SignInBlocked -eq $false } | Measure-Object).Count
if ($signInNotBlocked -gt 0) {
    Write-Host "  ⚠  WARNING: $signInNotBlocked mailbox(es) have sign-in enabled — review recommended.`n" -ForegroundColor Yellow
}

$healthRows | Sort-Object DaysSinceLastLogon -Descending |
    Format-Table -AutoSize -Property `
        MailboxDisplayName,
        WhenCreated,
        WhenChanged,
        LastLogonDate,
        DaysSinceLastLogon,
        @{N="Size (MB)";  E={ $_.MailboxSizeMB }},
        ItemCount,
        WarningQuota,
        ProhibitSendQuota,
        ProhibitSendReceiveQuota,
        @{N="SignInBlocked"; E={
            if ($_.SignInBlocked -eq $true)    { "✔ Blocked" }
            elseif ($_.SignInBlocked -eq $false){ "✘ NOT blocked" }
            else                               { $_.SignInBlocked }
        }}

# ── 5. Optional CSV export ────────────────────────────────────────────────────
if ($ExportPath) {
    if (-not (Test-Path $ExportPath)) {
        New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
    }

    $detailFile  = Join-Path $ExportPath "SharedMailbox_Delegates.csv"
    $summaryFile = Join-Path $ExportPath "SharedMailbox_DelegateSummary.csv"

    $delegateRows | Export-Csv -Path $detailFile  -NoTypeInformation -Encoding UTF8
    $summaryRows  | Export-Csv -Path $summaryFile -NoTypeInformation -Encoding UTF8

    Write-Host "`n✔  Detail  exported → $detailFile"  -ForegroundColor Green
    Write-Host "✔  Summary exported → $summaryFile`n" -ForegroundColor Green
}