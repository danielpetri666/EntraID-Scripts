#Requires -Version 7.0
#Requires -Module Microsoft.Graph.Authentication
<#
.SYNOPSIS
    Enforces a lifecycle policy for inactive guest accounts: disable and delete based on thresholds.
.DESCRIPTION
    Retrieves all guest accounts from Microsoft Graph (beta endpoint, required for
    signInActivity) and evaluates each against configurable inactivity thresholds.
    Guests inactive beyond the disable threshold are disabled; guests inactive beyond
    the delete threshold (and already disabled) are permanently deleted.
    Inactivity is determined from the later of interactive and non-interactive last sign-in.
    For guests that have never signed in, the account creation date is used as the baseline.
    Supports -WhatIf for safe preview of all actions before committing changes.
.AUTHOR
    Daniel Petri
.EXAMPLE
    .\Invoke-InactiveGuestLifecycle.ps1 -DisableThresholdMonths 6 -DeleteThresholdMonths 12
    Disables guests inactive for more than 6 months and deletes guests inactive for more than
    12 months (if already disabled).
.EXAMPLE
    .\Invoke-InactiveGuestLifecycle.ps1 -WhatIf
    Previews which accounts would be disabled or deleted without making any changes.
.NOTES
    Requires: Microsoft.Graph.Authentication
    Version: 2.0.0
    Permissions required: User.ReadWrite.All, AuditLog.Read.All
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [int]$DisableThresholdMonths = 6,

    [int]$DeleteThresholdMonths = 12,

    [string]$TenantId
)

function Invoke-GraphGetSafe {
    param([Parameter(Mandatory)][string]$Uri)

    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
        }
        catch {
            $msg = $_.Exception.Message
            if (($msg -match '429') -or ($msg -match 'Too Many Requests') -or ($msg -match '503') -or ($msg -match '504')) {
                $wait = [Math]::Min(5 * [Math]::Pow(2, $attempt - 1), 60)
                Write-Host "  Throttled (attempt $attempt/5), waiting ${wait}s..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }

    throw "Failed after 5 retries: $Uri"
}

# Connect to Microsoft Graph
$connectParams = @{
    Scopes = @('User.ReadWrite.All', 'AuditLog.Read.All')
}
if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
    $connectParams['TenantId'] = $TenantId
}
Connect-MgGraph @connectParams -ErrorAction Stop | Out-Null

$mgContext = Get-MgContext
Write-Host "Connected: $($mgContext.Account) | Tenant: $($mgContext.TenantId)" -ForegroundColor DarkCyan

#region Retrieve guest accounts via beta (required for signInActivity)
Write-Host 'Retrieving guest accounts...' -ForegroundColor Cyan
$allGuests = [System.Collections.Generic.List[object]]::new()
$uri = "https://graph.microsoft.com/beta/users?`$filter=userType eq 'Guest'&`$select=id,userPrincipalName,mail,creationType,createdDateTime,signInActivity,accountEnabled,userType,externalUserState"
do {
    $response = Invoke-GraphGetSafe -Uri $uri
    if ($response.value) {
        foreach ($item in $response.value) { $allGuests.Add($item) }
    }
    $uri = $response.'@odata.nextLink'
} while ($uri)

if ($allGuests.Count -eq 0) {
    throw 'No guest accounts found in this tenant.'
}
Write-Host "Guests retrieved: $($allGuests.Count)"
#endregion

#region Build guest lifecycle report
$disableDate = (Get-Date).AddMonths(-$DisableThresholdMonths)
$deleteDate  = (Get-Date).AddMonths(-$DeleteThresholdMonths)

$guestReport = $allGuests | ForEach-Object {
    $lastInteractive    = $_.signInActivity.lastSignInDateTime
    $lastNonInteractive = $_.signInActivity.lastNonInteractiveSignInDateTime
    $lastSignIn = if ($lastNonInteractive -gt $lastInteractive) { $lastNonInteractive } else { $lastInteractive }

    # Determine lifecycle action based on last activity or creation date
    $baselineDate = if ($null -ne $lastSignIn) { [datetime]$lastSignIn } else { [datetime]$_.createdDateTime }
    $action = 'None'
    if ($null -ne $baselineDate) {
        if ($baselineDate -lt $disableDate) { $action = 'Disable' }
        if ($baselineDate -lt $deleteDate)  { $action = 'Delete' }
    }

    [PSCustomObject][ordered]@{
        Id                = $_.id
        UserPrincipalName = $_.userPrincipalName
        Mail              = $_.mail
        CreationType      = $_.creationType
        CreatedDateTime   = $_.createdDateTime
        ExternalUserState = $_.externalUserState
        LastSignIn        = if ($lastSignIn) { ([datetime]$lastSignIn).ToString('yyyy-MM-dd') } else { 'N/A' }
        AccountEnabled    = $_.accountEnabled
        Action            = $action
    }
}
#endregion

#region Disable inactive guest accounts
$errors = [System.Collections.Generic.List[string]]::new()

$guestsToDisable = @($guestReport | Where-Object {
    $_.Action -eq 'Disable' -and
    $_.AccountEnabled -eq $true -and
    $_.ExternalUserState -ne 'PendingAcceptance'
})
Write-Host "Guests to disable: $($guestsToDisable.Count)"

foreach ($guest in $guestsToDisable) {
    if ($PSCmdlet.ShouldProcess($guest.UserPrincipalName, 'Disable guest account')) {
        try {
            Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$($guest.Id)" `
                -Body (@{ accountEnabled = $false } | ConvertTo-Json) -ContentType 'application/json' -ErrorAction Stop
            Write-Output "Disabled: $($guest.UserPrincipalName)"
        }
        catch {
            $shortError = "ERROR: Failed to disable '$($guest.UserPrincipalName)': $($_.Exception.Message)"
            Write-Output $shortError
            $errors.Add($shortError)
        }
    }
}
#endregion

#region Delete inactive guest accounts
$guestsToDelete = @($guestReport | Where-Object { $_.Action -eq 'Delete' -and $_.AccountEnabled -eq $false })
Write-Host "Guests to delete: $($guestsToDelete.Count)"

foreach ($guest in $guestsToDelete) {
    if ($PSCmdlet.ShouldProcess($guest.UserPrincipalName, 'Delete guest account')) {
        try {
            Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/users/$($guest.Id)" -ErrorAction Stop
            Write-Output "Deleted: $($guest.UserPrincipalName)"
        }
        catch {
            $shortError = "ERROR: Failed to delete '$($guest.UserPrincipalName)': $($_.Exception.Message)"
            Write-Output $shortError
            $errors.Add($shortError)
        }
    }
}
#endregion

#region Summary
Write-Host "`nLifecycle summary:" -ForegroundColor Cyan
Write-Host "  Guests evaluated: $($allGuests.Count)"
Write-Host "  Disabled: $($guestsToDisable.Count)"
Write-Host "  Deleted: $($guestsToDelete.Count)"

if ($errors.Count -gt 0) {
    Write-Output "Completed with $($errors.Count) error(s). Review messages above."
    throw "Lifecycle enforcement completed with $($errors.Count) error(s)."
}
else {
    Write-Output 'Lifecycle enforcement completed successfully.'
}
#endregion
