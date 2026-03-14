#Requires -Module Microsoft.Graph.Authentication
#Requires -Module Az.OperationalInsights
#Requires -Module ImportExcel
<#
.SYNOPSIS
    Reports which MFA methods are actually used by a list of users during sign-in.
.DESCRIPTION
    For a supplied list of user UPNs (from a CSV), this script:
      1. Retrieves each user's registered MFA methods from the Microsoft Graph
         authentication methods registration report.
      2. Queries a Log Analytics workspace for sign-in logs over the specified
         number of days and identifies which MFA methods were used in practice.
      3. Optionally enriches the report with user properties from Entra ID
         (e.g., companyName, department, extensionAttribute1-15).
      4. Produces a report showing the most-used method, all used methods, and
         which apps triggered MFA challenges, for each user.
    Requires an active Azure (Az module) and Microsoft Graph connection.
.AUTHOR
    Daniel Petri
.PARAMETER CsvPath
    Path to a semicolon-delimited, UTF-8 CSV file containing a 'UPN' column.
.PARAMETER WorkspaceId
    The ID of the Log Analytics workspace to query for sign-in logs.
.PARAMETER Days
    Number of days to look back in the sign-in logs. Default: 60.
.PARAMETER TenantId
    Optional tenant ID for multi-tenant environments.
.PARAMETER OutputPath
    Path for the Excel output file. Defaults to a timestamped file in $env:TEMP.
.PARAMETER UserProperties
    Optional list of Graph user property names to include in the report.
    Accepts standard properties (e.g., companyName, department, jobTitle,
    displayName) and extension attributes (extensionAttribute1-15).
    When specified, the script calls the /v1.0/users endpoint and adds the
    requested properties as extra columns after UserPrincipalName.
    Requires the User.Read.All Graph scope.
.EXAMPLE
    .\Get-MfaMethodUsedFromUpn.ps1 -CsvPath "C:\Reports\Users.csv" -WorkspaceId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    Runs the report for the users in the CSV against the specified Log Analytics workspace
    using the default 60-day window. Only core MFA columns are included.
.EXAMPLE
    .\Get-MfaMethodUsedFromUpn.ps1 -CsvPath "C:\Reports\Users.csv" -WorkspaceId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -Days 90 -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy"
    Uses a 90-day window and connects to the specified tenant.
.EXAMPLE
    .\Get-MfaMethodUsedFromUpn.ps1 -CsvPath "C:\Reports\Users.csv" -WorkspaceId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -UserProperties companyName, extensionAttribute13, department
    Includes CompanyName, ExtensionAttribute13, and Department columns from Entra ID
    user objects alongside the core MFA columns.
.NOTES
    Requires: Microsoft.Graph.Authentication, Az.OperationalInsights, ImportExcel
    Version: 1.1.0
    Permissions required (Graph): UserAuthenticationMethod.Read.All, AuditLog.Read.All
                                  User.Read.All is only needed when -UserProperties is used.
    Permissions required (Azure): Reader on the Log Analytics workspace
    CSV format: Must contain a column named 'UPN'.
#>

param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkspaceId,

    [int]$Days = 60,

    [string]$TenantId,

    [string]$OutputPath = (Join-Path $env:TEMP "$(Get-Date -Format 'yyyy-MM-dd')-MfaMethodsUsed.xlsx"),

    [string[]]$UserProperties
)

# Connect to Azure (required for Log Analytics query)
$azParams = @{}
if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $azParams['TenantId'] = $TenantId }
Connect-AzAccount @azParams -ErrorAction Stop | Out-Null

# Connect to Microsoft Graph
$graphScopes = @('UserAuthenticationMethod.Read.All', 'AuditLog.Read.All')
if ($UserProperties) { $graphScopes += 'User.Read.All' }

$mgParams = @{ Scopes = $graphScopes }
if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $mgParams['TenantId'] = $TenantId }
Connect-MgGraph @mgParams -ErrorAction Stop | Out-Null

#region Load and validate CSV
$csvUsers = Import-Csv -Path $CsvPath -Delimiter ';' -Encoding UTF8
if (-not $csvUsers -or $csvUsers.Count -eq 0) {
    Write-Error "No records found in CSV: '$CsvPath'" -ErrorAction Stop
}
if (-not ($csvUsers | Get-Member -Name 'UPN')) {
    throw "The CSV file '$CsvPath' must contain a column named 'UPN'."
}
Write-Output "Users loaded from CSV: $($csvUsers.Count)"
#endregion

#region Retrieve Entra ID user details (only when -UserProperties is specified)
$entraUsers = $null
if ($UserProperties) {
    # Build $select: always include id and userPrincipalName for matching
    $selectFields = [System.Collections.Generic.List[string]]::new()
    $selectFields.Add('id')
    $selectFields.Add('userPrincipalName')

    # Determine which properties need onPremisesExtensionAttributes vs direct select
    $hasExtensionAttributes = $false
    foreach ($prop in $UserProperties) {
        if ($prop -match '^extensionAttribute\d+$') {
            $hasExtensionAttributes = $true
        } else {
            if ($prop -notin $selectFields) { $selectFields.Add($prop) }
        }
    }
    if ($hasExtensionAttributes) { $selectFields.Add('onPremisesExtensionAttributes') }

    $entraUsers = [System.Collections.Generic.List[object]]::new()
    $uri = "https://graph.microsoft.com/v1.0/users?`$select=$($selectFields -join ',')"
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        if ($response.value) { foreach ($item in $response.value) { $entraUsers.Add($item) } }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    if ($entraUsers.Count -eq 0) {
        Write-Error 'Could not retrieve Entra users from Microsoft Graph.' -ErrorAction Stop
    }
    Write-Output "Entra users retrieved: $($entraUsers.Count)"
}
#endregion

#region Retrieve MFA registration details
$registrationDetails = [System.Collections.Generic.List[object]]::new()
$uri = 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails'
do {
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    if ($response.value) { foreach ($item in $response.value) { $registrationDetails.Add($item) } }
    $uri = $response.'@odata.nextLink'
} while ($uri)

$filteredRegistrations = @($registrationDetails | Where-Object { $_.userPrincipalName -in $csvUsers.UPN })
if ($filteredRegistrations.Count -eq 0) {
    Write-Error 'No matching MFA registration details found for the supplied UPNs.' -ErrorAction Stop
}
Write-Output "MFA registration records matched: $($filteredRegistrations.Count)"
#endregion

#region Query sign-in logs from Log Analytics
$query = @"
SigninLogs
| where TimeGenerated > ago($($Days)d)
| extend AuthenticationMethod1 = tostring(parse_json(AuthenticationDetails)[0].authenticationMethod)
| extend AuthenticationMethodDetail1 = tostring(parse_json(AuthenticationDetails)[0].authenticationMethodDetail)
| extend AuthenticationMethodSucceeded1 = tostring(parse_json(AuthenticationDetails)[0].succeeded)
| extend AuthenticationMethod2 = tostring(parse_json(AuthenticationDetails)[1].authenticationMethod)
| extend AuthenticationMethodDetail2 = tostring(parse_json(AuthenticationDetails)[1].authenticationMethodDetail)
| extend AuthenticationMethodSucceeded2 = tostring(parse_json(AuthenticationDetails)[1].succeeded)
| where AuthenticationMethod2 != ''
| where not(AuthenticationMethod1 == 'Previously satisfied' and AuthenticationMethod2 == 'Previously satisfied')
| where (AuthenticationMethodSucceeded1 == 'true' and AuthenticationMethodSucceeded2 == 'true')
| project TimeGenerated, UserPrincipalName, AppDisplayName, ResultType, AuthenticationRequirement, Status_additionalDetails = Status.additionalDetails, AuthenticationMethod1, AuthenticationMethodDetail1, AuthenticationMethodSucceeded1, AuthenticationMethod2, AuthenticationMethodDetail2, AuthenticationMethodSucceeded2
"@

$queryResult = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceId -Query $query -ErrorAction Stop
$signinLogs = @($queryResult.Results | Where-Object { $_.UserPrincipalName -in $csvUsers.UPN })

if ($signinLogs.Count -eq 0) {
    Write-Error "No sign-in log entries found for the supplied UPNs in the last $Days days." -ErrorAction Stop
}
Write-Output "Sign-in log entries matched: $($signinLogs.Count)"
#endregion

#region Generate per-user report
$report = [System.Collections.Generic.List[object]]::new()
foreach ($csvUser in $csvUsers) {
    $userSigninLogs = @($signinLogs | Where-Object { $_.UserPrincipalName -eq $csvUser.UPN })
    $registeredMethods = ($filteredRegistrations | Where-Object { $_.userPrincipalName -eq $csvUser.UPN }).methodsRegistered -join ', '

    $methodCount = @{}
    $appCount = @{}

    foreach ($log in $userSigninLogs) {
        foreach ($method in @($log.AuthenticationMethod1, $log.AuthenticationMethod2)) {
            if ($method -and $method -notin @('Previously satisfied', 'Password')) {
                $methodCount[$method] = ($methodCount[$method] ?? 0) + 1
            }
        }
        if ($log.AppDisplayName) {
            $appCount[$log.AppDisplayName] = ($appCount[$log.AppDisplayName] ?? 0) + 1
        }
    }

    $methodsFormatted = $methodCount.GetEnumerator() | ForEach-Object { "$($_.Key) ($($_.Value))" } | Sort-Object
    $appsFormatted    = $appCount.GetEnumerator()    | ForEach-Object { "$($_.Key) ($($_.Value))" } | Sort-Object
    $mostUsedMethod   = if ($methodCount.Count -gt 0) {
        $top = $methodCount.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1
        "$($top.Key) ($($top.Value))"
    } else { '' }

    # Build the output object: start with UPN, then optional user properties, then MFA columns
    $row = [ordered]@{
        UserPrincipalName = $csvUser.UPN
    }

    if ($UserProperties -and $entraUsers) {
        $entraUser = $entraUsers | Where-Object { $_.userPrincipalName -eq $csvUser.UPN } | Select-Object -First 1
        foreach ($prop in $UserProperties) {
            # Convert property name to PascalCase for the column header
            $columnName = $prop.Substring(0, 1).ToUpper() + $prop.Substring(1)
            if ($prop -match '^extensionAttribute(\d+)$') {
                $row[$columnName] = $entraUser.onPremisesExtensionAttributes.$prop
            } else {
                $row[$columnName] = $entraUser.$prop
            }
        }
    }

    $row['RegisteredMethods'] = $registeredMethods
    $row['MostUsedMethod']    = $mostUsedMethod
    $row['UsedMethods']       = $methodsFormatted -join ', '
    $row['Apps']              = $appsFormatted -join ', '
    $row['SignInCount']       = $userSigninLogs.Count

    $report.Add([PSCustomObject]$row)
}
#endregion

#region Output
$report | Out-GridView -Title 'MFA Methods Used'
$report | Export-Excel -Path $OutputPath -TableStyle Medium2 -AutoSize -Show
Write-Host "Exported to: $OutputPath" -ForegroundColor Green
#endregion
