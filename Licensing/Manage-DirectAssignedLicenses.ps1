#Requires -Module Microsoft.Graph.Authentication
#Requires -Module Microsoft.Graph.Users
<#
.SYNOPSIS
    Finds and optionally removes directly assigned licenses outside group-based licensing.
.DESCRIPTION
    Connects to Microsoft Graph and retrieves all users with license assignments.
    Identifies users who have at least one license assigned directly (not through a
    group) by checking LicenseAssignmentStates.AssignedByGroup. Reports the direct
    SKUs by name, along with user context (department, account status) for governance
    review.

    With -RemoveDirect, removes ALL direct license assignments unconditionally.
    With -RemoveSafe, removes direct assignments ONLY when the same SKU is also
    assigned through a group with an Active state.

    Both removal modes support -WhatIf and log each action to console and optionally
    to a log file.

    Use this to detect and clean up license drift -- users who got a direct assignment
    during troubleshooting, onboarding, or migration and are now outside your group-based
    licensing model.
.AUTHOR
    Daniel Petri
.EXAMPLE
    .\Manage-DirectAssignedLicenses.ps1
    Finds all users with direct license assignments and shows results in Out-GridView.
.EXAMPLE
    .\Manage-DirectAssignedLicenses.ps1 -ExportCsv "C:\Reports\DirectLicenses.csv"
    Exports the results to CSV.
.EXAMPLE
    .\Manage-DirectAssignedLicenses.ps1 -RemoveDirect
    Removes ALL direct license assignments regardless of group coverage.
.EXAMPLE
    .\Manage-DirectAssignedLicenses.ps1 -RemoveSafe
    Removes direct assignments only when the same SKU is also assigned via a group.
.EXAMPLE
    .\Manage-DirectAssignedLicenses.ps1 -RemoveSafe -WhatIf
    Shows what would be removed without making any changes.
.EXAMPLE
    .\Manage-DirectAssignedLicenses.ps1 -RemoveSafe -LogFile "C:\Logs\LicenseCleanup.log"
    Removes safe direct assignments and writes a log file.
.EXAMPLE
    .\Manage-DirectAssignedLicenses.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -NoGridView
    Targets a specific tenant without opening the grid view.
.NOTES
    Requires: Microsoft.Graph.Authentication, Microsoft.Graph.Users
    Permissions required: User.Read.All, Organization.Read.All (delegated)
        With -RemoveDirect or -RemoveSafe: User.ReadWrite.All, Organization.Read.All (delegated)
    Version: 2.1.0
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$TenantId,

    [string]$ExportCsv,

    [switch]$NoGridView,

    [switch]$RemoveDirect,

    [switch]$RemoveSafe,

    [string]$LogFile
)

if ($RemoveDirect -and $RemoveSafe) {
    Write-Warning 'Use either -RemoveDirect or -RemoveSafe, not both.'
    return
}

#region Logging helper
function Write-Log {
    param([string]$Message, [string]$Color = 'White')
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] $Message"
    Write-Host $line -ForegroundColor $Color
    if ($LogFile) {
        $line | Out-File -FilePath $LogFile -Append -Encoding UTF8
    }
}
#endregion

#region Connect to Microsoft Graph
$scopes = @('Organization.Read.All')
if ($RemoveDirect -or $RemoveSafe) {
    $scopes += 'User.ReadWrite.All'
} else {
    $scopes += 'User.Read.All'
}

$connectParams = @{ Scopes = $scopes }
if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
    $connectParams['TenantId'] = $TenantId
}
Connect-MgGraph @connectParams -ErrorAction Stop | Out-Null
$ctx = Get-MgContext
Write-Host "Connected to Microsoft Graph | $($ctx.Account) | TenantId: $($ctx.TenantId)" -ForegroundColor Green
#endregion

#region Build SKU name lookup
Write-Host 'Loading SKU definitions...' -ForegroundColor Cyan
$skuResponse = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus?$select=skuId,skuPartNumber'
$skuMap = @{}
foreach ($sku in $skuResponse.value) {
    $skuMap[$sku.skuId] = $sku.skuPartNumber
}
Write-Host "  SKUs in tenant: $($skuMap.Count)"
#endregion

#region Retrieve users with license assignments
Write-Host 'Retrieving licensed users...' -ForegroundColor Cyan
$users = Get-MgUser -All `
    -Property Id, UserPrincipalName, DisplayName, Department, AccountEnabled, LicenseAssignmentStates, AssignedLicenses

$licensedUsers = @($users | Where-Object { $_.AssignedLicenses -and @($_.AssignedLicenses).Count -gt 0 })
Write-Host "  Licensed users: $($licensedUsers.Count)"

if ($licensedUsers.Count -gt 0 -and $null -eq $licensedUsers[0].LicenseAssignmentStates) {
    Write-Warning 'LicenseAssignmentStates is null. The property may not have been returned by Graph. Verify the -Property parameter includes LicenseAssignmentStates.'
    return
}
#endregion

#region Find users with direct assignments
$report = [System.Collections.Generic.List[object]]::new()
$usersWithDirectSkus = [System.Collections.Generic.List[object]]::new()

foreach ($user in $licensedUsers) {
    if (-not $user.LicenseAssignmentStates) { continue }

    $directAssignments = @($user.LicenseAssignmentStates | Where-Object { $null -eq $_.AssignedByGroup })
    if ($directAssignments.Count -eq 0) { continue }

    $groupAssignments = @($user.LicenseAssignmentStates | Where-Object { $null -ne $_.AssignedByGroup })
    $activeGroupSkuIds = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($groupAssignments | Where-Object { $_.State -eq 'Active' } | ForEach-Object { $_.SkuId }),
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($da in $directAssignments) {
        $skuName = if ($skuMap[$da.SkuId]) { $skuMap[$da.SkuId] } else { $da.SkuId }

        [void]$report.Add([PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Department        = $user.Department
            AccountEnabled    = $user.AccountEnabled
            License           = $skuName
            DirectAssigned    = $true
            GroupAssigned     = $activeGroupSkuIds.Contains($da.SkuId)
        })
    }

    [void]$usersWithDirectSkus.Add(@{
        User               = $user
        DirectAssignments  = $directAssignments
        ActiveGroupSkuIds  = $activeGroupSkuIds
    })
}

Write-Host ''
Write-Host "Users with direct license assignments: $($report.Count)" -ForegroundColor Yellow
#endregion

#region Output
if ($report.Count -eq 0) {
    Write-Host 'No direct license assignments found. All licenses are group-assigned.' -ForegroundColor Green
    return
}

$disabledWithLicenses = @($report | Where-Object { $_.AccountEnabled -eq $false }).Count
if ($disabledWithLicenses -gt 0) {
    Write-Host "  Disabled accounts with direct licenses: $disabledWithLicenses" -ForegroundColor Red
}

if (-not $NoGridView) {
    try {
        $report | Out-GridView -Title "Users with Direct License Assignments"
    }
    catch {
        Write-Warning 'Out-GridView not available. Use -NoGridView to suppress.'
    }
}

if ($ExportCsv) {
    $report | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to: $ExportCsv" -ForegroundColor Green
}
#endregion

#region Remove direct assignments
if ($RemoveDirect -or $RemoveSafe) {
    $mode = if ($RemoveDirect) { 'RemoveDirect (all)' } else { 'RemoveSafe (group-covered only)' }
    Write-Host ''
    Write-Log "Starting removal -- mode: $mode" 'Cyan'

    $removed = 0
    $skipped = 0
    $failed = 0

    foreach ($entry in $usersWithDirectSkus) {
        $user = $entry.User
        $activeGroupSkuIds = $entry.ActiveGroupSkuIds

        foreach ($assignment in $entry.DirectAssignments) {
            $skuId = $assignment.SkuId
            $skuName = if ($skuMap[$skuId]) { $skuMap[$skuId] } else { $skuId }

            if ($RemoveSafe -and -not $activeGroupSkuIds.Contains($skuId)) {
                Write-Log "SKIP  $($user.UserPrincipalName) | $skuName (no active group coverage)" 'DarkYellow'
                $skipped++
                continue
            }

            if (-not $PSCmdlet.ShouldProcess("$($user.UserPrincipalName) | $skuName", 'Remove direct license')) {
                $skipped++
                continue
            }

            try {
                Set-MgUserLicense -UserId $user.Id `
                    -AddLicenses @() `
                    -RemoveLicenses @($skuId) `
                    -ErrorAction Stop | Out-Null

                Write-Log "OK    $($user.UserPrincipalName) | $skuName removed" 'Green'
                $removed++
            }
            catch {
                Write-Log "FAIL  $($user.UserPrincipalName) | $skuName | $($_.Exception.Message)" 'Red'
                $failed++
            }
        }
    }

    Write-Host ''
    Write-Log "Done. Removed: $removed | Skipped: $skipped | Failed: $failed" 'Cyan'
}
#endregion
