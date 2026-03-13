#Requires -Version 7.0
#Requires -Module Az.Accounts
#Requires -Module Az.Resources
<#
.SYNOPSIS
    Grants Microsoft Graph application permissions to a Managed Identity service principal.
.DESCRIPTION
    Connects to Azure and assigns the specified Microsoft Graph application permissions to a
    Managed Identity (or any service principal) identified by its object ID. Existing
    assignments are checked to prevent duplicates. Unrecognized permission names are reported
    as warnings. Supports -WhatIf for safe preview of changes.
.AUTHOR
    Daniel Petri
.EXAMPLE
    .\Add-ManagedIdentityGraphPermissions.ps1 -PrincipalId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    Grants the default permission (User.Read.All) to the specified service principal.
.EXAMPLE
    .\Add-ManagedIdentityGraphPermissions.ps1 -PrincipalId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -GraphPermissions @("User.Read.All","Group.Read.All","AuditLog.Read.All") -WhatIf
    Previews which permissions would be assigned without making changes.
.NOTES
    Requires: Az.Accounts, Az.Resources
    Version: 2.0.0
    The account running this script must have the Privileged Role Administrator or
    Global Administrator role to grant application permissions.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$PrincipalId,

    [string[]]$GraphPermissions = @('User.Read.All')
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

# Find matching app roles
$graphAppRoles = $graphServicePrincipal.AppRole | Where-Object {
    $GraphPermissions -contains $_.Value -and $_.AllowedMemberType -contains 'Application'
}

# Warn about unrecognized permissions
$foundNames = @($graphAppRoles | ForEach-Object { $_.Value })
foreach ($requested in $GraphPermissions) {
    if ($requested -notin $foundNames) {
        Write-Warning "Permission '$requested' was not found as an application permission on Microsoft Graph. Check the spelling."
    }
}

if (-not $graphAppRoles -or $graphAppRoles.Count -eq 0) {
    Write-Warning 'No matching Graph application permissions found. Nothing to assign.'
    return
}

# Get current assignments
$currentPermissions = Get-AzADServicePrincipalAppRoleAssignment -ServicePrincipalId $PrincipalId

# Restore WhatIf for the actual write operations
$WhatIfPreference = $savedWhatIf

# Assign missing permissions
foreach ($role in $graphAppRoles) {
    if ($currentPermissions.AppRoleId -notcontains $role.Id) {
        if ($PSCmdlet.ShouldProcess($role.Value, "Assign to service principal $PrincipalId")) {
            try {
                New-AzADServicePrincipalAppRoleAssignment `
                    -ServicePrincipalId $PrincipalId `
                    -ResourceId $graphServicePrincipal.Id `
                    -AppRoleId $role.Id | Out-Null
                Write-Output "Assigned: '$($role.Value)' to service principal $PrincipalId"
            }
            catch {
                Write-Error "Failed to assign '$($role.Value)' to service principal $PrincipalId : $($_.Exception.Message)"
            }
        }
    }
    else {
        Write-Output "Already assigned: '$($role.Value)' on service principal $PrincipalId"
    }
}

# Verify current state
Write-Host "`nCurrent permissions for $PrincipalId`:" -ForegroundColor Cyan
$WhatIfPreference = $false
$updated = Get-AzADServicePrincipalAppRoleAssignment -ServicePrincipalId $PrincipalId
$roleMap = @{}
$graphServicePrincipal.AppRole | ForEach-Object { $roleMap[$_.Id] = $_.Value }
$updated | ForEach-Object {
    $name = if ($roleMap.ContainsKey($_.AppRoleId)) { $roleMap[$_.AppRoleId] } else { $_.AppRoleId }
    Write-Host "  $name"
}
