#Requires -Version 7.0
#Requires -Module Az.Accounts
#Requires -Module Az.Resources
<#
.SYNOPSIS
    Removes Microsoft Graph application permissions from a Managed Identity service principal.
.DESCRIPTION
    Connects to Azure and removes the specified Microsoft Graph application permissions from a
    Managed Identity (or any service principal) identified by its object ID. Permissions that
    are not currently assigned are silently skipped. Supports -WhatIf for safe preview.
.AUTHOR
    Daniel Petri
.EXAMPLE
    .\Remove-ManagedIdentityGraphPermissions.ps1 -PrincipalId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -GraphPermissions @("User.Read.All")
    Removes User.Read.All from the specified service principal.
.EXAMPLE
    .\Remove-ManagedIdentityGraphPermissions.ps1 -PrincipalId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -GraphPermissions @("User.Read.All","Group.Read.All") -WhatIf
    Previews which permissions would be removed without making changes.
.NOTES
    Requires: Az.Accounts, Az.Resources
    Version: 1.0.0
    The account running this script must have the Privileged Role Administrator or
    Global Administrator role to remove application permissions.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$PrincipalId,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string[]]$GraphPermissions
)

# Suppress WhatIf for all read operations (Az cmdlets inherit $WhatIfPreference)
$savedWhatIf = $WhatIfPreference
$WhatIfPreference = $false

Connect-AzAccount -ErrorAction Stop

# Look up the Microsoft Graph service principal by its well-known AppId
$graphServicePrincipal = Get-AzADServicePrincipal -SearchString 'Microsoft Graph' | Select-Object -First 1
if (-not $graphServicePrincipal) {
    throw 'Could not locate the Microsoft Graph service principal in this tenant.'
}

# Build role name-to-ID map
$roleMap = @{}
$graphServicePrincipal.AppRole | ForEach-Object { $roleMap[$_.Value] = $_.Id }

# Get current assignments
$currentAssignments = Get-AzADServicePrincipalAppRoleAssignment -ServicePrincipalId $PrincipalId

# Restore WhatIf for the actual write operations
$WhatIfPreference = $savedWhatIf

foreach ($permName in $GraphPermissions) {
    if (-not $roleMap.ContainsKey($permName)) {
        Write-Warning "Permission '$permName' was not found on Microsoft Graph. Check the spelling."
        continue
    }

    $roleId = $roleMap[$permName]
    $assignment = $currentAssignments | Where-Object { $_.AppRoleId -eq $roleId }

    if (-not $assignment) {
        Write-Output "Not assigned: '$permName' on service principal $PrincipalId (skipping)"
        continue
    }

    if ($PSCmdlet.ShouldProcess($permName, "Remove from service principal $PrincipalId")) {
        try {
            Remove-AzADServicePrincipalAppRoleAssignment `
                -AppRoleAssignmentId $assignment.Id `
                -ServicePrincipalId $PrincipalId -ErrorAction Stop
            Write-Output "Removed: '$permName' from service principal $PrincipalId"
        }
        catch {
            Write-Error "Failed to remove '$permName' from service principal $PrincipalId : $($_.Exception.Message)"
        }
    }
}

# Show remaining permissions
Write-Host "`nRemaining permissions for $PrincipalId`:" -ForegroundColor Cyan
$WhatIfPreference = $false
$remaining = Get-AzADServicePrincipalAppRoleAssignment -ServicePrincipalId $PrincipalId
if (-not $remaining -or $remaining.Count -eq 0) {
    Write-Host '  (none)'
}
else {
    $idToName = @{}
    $graphServicePrincipal.AppRole | ForEach-Object { $idToName[$_.Id] = $_.Value }
    $remaining | ForEach-Object {
        $name = if ($idToName.ContainsKey($_.AppRoleId)) { $idToName[$_.AppRoleId] } else { $_.AppRoleId }
        Write-Host "  $name"
    }
}
