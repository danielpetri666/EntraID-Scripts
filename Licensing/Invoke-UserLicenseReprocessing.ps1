#Requires -Module Microsoft.Graph.Authentication
#Requires -Module Microsoft.Graph.Users
#Requires -Module Microsoft.Graph.Users.Actions
<#
.SYNOPSIS
    Triggers license reprocessing for users in the tenant.
.DESCRIPTION
    Connects to Microsoft Graph and calls Invoke-MgLicenseUser for each target user to force
    the license service to re-evaluate group-based license assignments. Useful after bulk group
    changes or when users are stuck in a license error state.

    -UserPrincipalName: targets a single user.
    -CsvPath: targets users listed in a CSV file (must contain a 'UPN' column).
    Neither: targets all licensed member users in the tenant.
.AUTHOR
    Daniel Petri
.EXAMPLE
    .\Invoke-UserLicenseReprocessing.ps1
    Connects interactively and reprocesses licenses for all licensed member users.
.EXAMPLE
    .\Invoke-UserLicenseReprocessing.ps1 -UserPrincipalName "user@contoso.com"
    Reprocesses licenses for a single user.
.EXAMPLE
    .\Invoke-UserLicenseReprocessing.ps1 -CsvPath "C:\Reports\Users.csv"
    Reprocesses licenses only for users listed in the CSV.
.EXAMPLE
    .\Invoke-UserLicenseReprocessing.ps1 -CsvPath "C:\Reports\Users.csv" -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    Connects to the specified tenant and reprocesses licenses for each user in the CSV.
.NOTES
    Requires: Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Users.Actions
    Permissions required: User.ReadWrite.All
    CSV format: Must contain a column named 'UPN'.
    Version: 3.0.0
#>

[CmdletBinding(DefaultParameterSetName = 'AllUsers')]
param(
    [Parameter(ParameterSetName = 'SingleUser', Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$UserPrincipalName,

    [Parameter(ParameterSetName = 'CsvFile', Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [string]$TenantId
)

# Connect to Microsoft Graph
$connectParams = @{
    Scopes = @('User.ReadWrite.All')
}
if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
    $connectParams['TenantId'] = $TenantId
}
Connect-MgGraph @connectParams -ErrorAction Stop | Out-Null

#region Build user list
$users = [System.Collections.Generic.List[object]]::new()

switch ($PSCmdlet.ParameterSetName) {
    'SingleUser' {
        Write-Host "Looking up user: '$UserPrincipalName'..." -ForegroundColor Cyan
        $user = Get-MgUser -Filter "userPrincipalName eq '$UserPrincipalName'" -ErrorAction SilentlyContinue
        if ($user) {
            $users.Add($user)
        } else {
            Write-Warning "User not found: '$UserPrincipalName'"
            return
        }
    }
    'CsvFile' {
        $csvUsers = Import-Csv -Path $CsvPath
        if (-not $csvUsers -or $csvUsers.Count -eq 0) {
            Write-Warning "The CSV file '$CsvPath' contains no records."
            return
        }
        if (-not ($csvUsers | Get-Member -Name 'UPN')) {
            throw "The CSV file '$CsvPath' must contain a column named 'UPN'."
        }
        Write-Host "Users loaded from CSV: $($csvUsers.Count)" -ForegroundColor Cyan
        foreach ($csvUser in $csvUsers) {
            $upn = $csvUser.UPN
            $user = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue
            if ($user) {
                $users.Add($user)
            } else {
                Write-Warning "User not found: '$upn'"
            }
        }
    }
    default {
        # All licensed members
        Write-Host 'Retrieving licensed member users...' -ForegroundColor Cyan
        [Array]$allUsers = Get-MgUser `
            -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
            -ConsistencyLevel eventual `
            -CountVariable Records `
            -All `
            -Property Id, AssignedLicenses, UserPrincipalName |
            Sort-Object UserPrincipalName
        if ($allUsers) {
            foreach ($u in $allUsers) { $users.Add($u) }
        }
    }
}

if ($users.Count -eq 0) {
    Write-Warning 'No users to process.'
    return
}
Write-Host "Users to reprocess: $($users.Count)" -ForegroundColor Cyan
#endregion

#region Reprocess license assignments
$successCount = 0
$errorCount = 0
foreach ($user in $users) {
    try {
        Invoke-MgLicenseUser -UserId $user.Id -ErrorAction Stop | Out-Null
        Write-Host "Reprocessed: $($user.UserPrincipalName)" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Host "Error: $($user.UserPrincipalName): $_" -ForegroundColor Red
        $errorCount++
    }
}
Write-Host "Done. $successCount succeeded, $errorCount failed." -ForegroundColor Cyan
#endregion
