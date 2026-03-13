#Requires -Version 7.0
#Requires -Module Microsoft.Graph.Authentication
#Requires -Module ImportExcel
<#
.SYNOPSIS
    Reports on all guest users in the tenant including last sign-in and manager details.
.DESCRIPTION
    Connects to Microsoft Graph and retrieves all guest accounts via the beta endpoint
    (required for SignInActivity). Builds a report with identity, creation date, last
    sign-in (interactive and non-interactive), account state, and the guest's manager UPN.
    The report is displayed in a grid view and exported to Excel.
.AUTHOR
    Daniel Petri
.EXAMPLE
    .\Get-GuestsInfo.ps1
    Connects interactively, retrieves all guests, and exports the report to the system
    temp folder.
.EXAMPLE
    .\Get-GuestsInfo.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -OutputPath "C:\Reports\GuestStatus.xlsx"
    Connects to the specified tenant and exports to the given path.
.NOTES
    Requires: Microsoft.Graph.Authentication, ImportExcel
    Version: 2.0.0
    Permissions required: User.Read.All, AuditLog.Read.All
#>

param(
    [string]$TenantId,
    [string]$OutputPath = (Join-Path $env:TEMP "$(Get-Date -Format 'yyyy-MM-dd')-GuestStatus.xlsx")
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
            return $null
        }
    }

    Write-Warning "Failed after 5 retries: $Uri"
    return $null
}

function Invoke-GraphGetPaged {
    param([Parameter(Mandatory)][string]$Uri)

    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri

    while (-not [string]::IsNullOrWhiteSpace($next)) {
        $response = Invoke-GraphGetSafe -Uri $next
        if ($null -eq $response) {
            Write-Warning "Failed to fetch page: $next"
            break
        }
        if ($response -is [System.Collections.IDictionary] -and $response.ContainsKey('value')) {
            foreach ($item in @($response.value)) { [void]$items.Add($item) }
            $next = $response.'@odata.nextLink'
        }
        else {
            if ($null -ne $response) { [void]$items.Add($response) }
            $next = $null
        }
    }

    return @($items)
}

# Connect to Microsoft Graph
$connectParams = @{
    Scopes = @('User.Read.All', 'AuditLog.Read.All')
}
if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
    $connectParams['TenantId'] = $TenantId
}
Connect-MgGraph @connectParams -ErrorAction Stop | Out-Null

$mgContext = Get-MgContext
Write-Host "Connected: $($mgContext.Account) | Tenant: $($mgContext.TenantId)" -ForegroundColor DarkCyan

# Retrieve all guest accounts via beta (required for signInActivity)
Write-Host 'Retrieving guest accounts...' -ForegroundColor Cyan
$allGuests = Invoke-GraphGetPaged -Uri "https://graph.microsoft.com/beta/users?`$filter=userType eq 'Guest'&`$select=id,userPrincipalName,mail,creationType,createdDateTime,signInActivity,accountEnabled,userType"

if (-not $allGuests -or $allGuests.Count -eq 0) {
    Write-Warning 'No guest accounts found in this tenant.'
    return
}
Write-Host "Guests found: $($allGuests.Count)"

# Build report with manager lookup
Write-Host 'Building report (resolving managers)...' -ForegroundColor Cyan

$report = $allGuests | ForEach-Object {
    $lastInteractive = $_.signInActivity.lastSignInDateTime
    $lastNonInteractive = $_.signInActivity.lastNonInteractiveSignInDateTime
    $lastSignIn = if ($lastNonInteractive -gt $lastInteractive) { $lastNonInteractive } else { $lastInteractive }

    # Resolve manager
    $mgr = Invoke-GraphGetSafe -Uri "https://graph.microsoft.com/v1.0/users/$($_.id)/manager?`$select=userPrincipalName"
    $managerUpn = if ($mgr -and $mgr.userPrincipalName) { $mgr.userPrincipalName } else { '' }

    [PSCustomObject][ordered]@{
        Id                = $_.id
        UserPrincipalName = $_.userPrincipalName
        Mail              = $_.mail
        CreationType      = $_.creationType
        CreatedDateTime   = $_.createdDateTime
        LastSignIn        = if ($lastSignIn) { ([datetime]$lastSignIn).ToString('yyyy-MM-dd') } else { 'N/A' }
        AccountEnabled    = $_.accountEnabled
        Manager           = $managerUpn
    }
}

# Output
try {
    $report | Out-GridView -Title "Guest Account Status ($($mgContext.TenantId))"
}
catch {
    Write-Warning 'Out-GridView not available (requires desktop environment).'
}

$report | Export-Excel -Path $OutputPath -TableStyle Medium2 -AutoSize
Write-Host "Exported to: $OutputPath" -ForegroundColor Green
