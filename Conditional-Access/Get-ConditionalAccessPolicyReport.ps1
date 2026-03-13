#Requires -Version 7.0
#Requires -Module Microsoft.Graph.Authentication
<#
.SYNOPSIS
    Generates a detailed Conditional Access policy report from Microsoft Graph.
.DESCRIPTION
    Fetches all Conditional Access policies from the Microsoft Graph API and produces a
    flat, human-readable report with one row per policy. Complex nested properties
    (conditions, grant controls, session controls) are expanded into individual columns
    and optional summary columns. Supports export to CSV, Excel (requires ImportExcel),
    and JSON, as well as interactive display via Out-GridView.
.AUTHOR
    Daniel Petri
.EXAMPLE
    .\Get-ConditionalAccessPolicyReport.ps1 -ExportExcel
    Connects to Microsoft Graph interactively, fetches all CA policies, and exports
    the report to an Excel file in the system temp folder.
.EXAMPLE
    .\Get-ConditionalAccessPolicyReport.ps1 -ExportCsvPath C:\Reports\CA-Policies.csv -NoGridView
    Exports the report to the specified CSV path without opening a grid view.
.EXAMPLE
    .\Get-ConditionalAccessPolicyReport.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ExportJson
    Connects to the specified tenant and exports raw policy JSON.
.NOTES
    Requires: Microsoft.Graph.Authentication, ImportExcel (only when -ExportExcel is used)
    Version: 1.0.0
    Permissions required: Policy.Read.All, Directory.Read.All, Agreement.Read.All (for Terms of Use)
#>

param(
    [string]$TenantId,

    [switch]$ExportCsv,
    [switch]$ExportExcel,
    [switch]$ExportJson,

    [string]$ExportCsvPath,
    [string]$ExportExcelPath,
    [string]$ExportJsonPath,

    [switch]$NoGridView
)

function Initialize-ParentDirectory {
    param([Parameter(Mandatory)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $parent = [System.IO.Path]::GetDirectoryName($fullPath)
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path $parent)) {
        [void](New-Item -Path $parent -ItemType Directory -Force)
    }

    return $fullPath
}

function Resolve-ExportPath {
    param(
        [string]$RequestedPath,
        [Parameter(Mandatory)][string]$DefaultBaseName,
        [Parameter(Mandatory)][string]$Extension
    )

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $defaultFileName = "{0}-{1}.{2}" -f $DefaultBaseName, $timestamp, $Extension

    if ([string]::IsNullOrWhiteSpace($RequestedPath)) {
        return Initialize-ParentDirectory -Path (Join-Path $env:TEMP $defaultFileName)
    }

    $candidate = $RequestedPath.Trim()
    $fullCandidate = [System.IO.Path]::GetFullPath($candidate)

    if ((Test-Path $fullCandidate) -and (Get-Item $fullCandidate).PSIsContainer) {
        return Initialize-ParentDirectory -Path (Join-Path $fullCandidate $defaultFileName)
    }

    $ext = [System.IO.Path]::GetExtension($fullCandidate)
    if ([string]::IsNullOrWhiteSpace($ext)) {
        return Initialize-ParentDirectory -Path (Join-Path $fullCandidate $defaultFileName)
    }

    return Initialize-ParentDirectory -Path $fullCandidate
}

function Convert-RowsForExcelLineBreaks {
    param(
        [Parameter(Mandatory)][array]$Rows,
        [Parameter(Mandatory)][string[]]$Columns
    )

    $outRows = New-Object System.Collections.Generic.List[object]

    function Get-JsonLines {
        param(
            [Parameter(Mandatory)]$Object,
            [string]$Prefix = ''
        )

        $lines = New-Object System.Collections.Generic.List[string]

        if ($null -eq $Object) {
            if (-not [string]::IsNullOrWhiteSpace($Prefix)) {
                $lines.Add("$Prefix=") | Out-Null
            }
            return $lines
        }

        if ($Object -is [System.Collections.IDictionary]) {
            foreach ($k in $Object.Keys) {
                $key = [string]$k
                $child = $Object[$k]
                $childPrefix = if ([string]::IsNullOrWhiteSpace($Prefix)) { $key } else { "$Prefix.$key" }
                $childLines = Get-JsonLines -Object $child -Prefix $childPrefix
                foreach ($line in $childLines) { $lines.Add($line) | Out-Null }
            }
            return $lines
        }

        if ($Object -is [System.Collections.IEnumerable] -and -not ($Object -is [string])) {
            $arr = @($Object)
            if ($arr.Count -eq 0) {
                $lines.Add("$Prefix=") | Out-Null
                return $lines
            }

            $simple = $true
            foreach ($item in $arr) {
                if (($item -is [System.Collections.IDictionary]) -or (($item -is [System.Collections.IEnumerable]) -and -not ($item -is [string]))) {
                    $simple = $false
                    break
                }
            }

            if ($simple) {
                $lines.Add(("{0}={1}" -f $Prefix, (($arr | ForEach-Object { [string]$_ }) -join ', '))) | Out-Null
                return $lines
            }

            for ($i = 0; $i -lt $arr.Count; $i++) {
                $childPrefix = "{0}[{1}]" -f $Prefix, $i
                $childLines = Get-JsonLines -Object $arr[$i] -Prefix $childPrefix
                foreach ($line in $childLines) { $lines.Add($line) | Out-Null }
            }

            return $lines
        }

        $lines.Add(("{0}={1}" -f $Prefix, [string]$Object)) | Out-Null
        return $lines
    }

    $jsonPrettyColumns = @('inclGuestsExternal', 'exclGuestsExternal', 'deviceFilter', 'authenticationFlows')

    foreach ($row in $Rows) {
        $newRow = [ordered]@{}
        foreach ($prop in $row.PSObject.Properties) {
            $propName = [string]$prop.Name
            $val = $prop.Value

            if ($Columns -contains $propName -and $val -is [string] -and -not [string]::IsNullOrWhiteSpace($val)) {
                if ($jsonPrettyColumns -contains $propName -and $val.TrimStart().StartsWith('{')) {
                    try {
                        $obj = $val | ConvertFrom-Json -AsHashtable -Depth 30
                        $jsonLines = Get-JsonLines -Object $obj
                        $val = ($jsonLines -join "`n")
                    }
                    catch {
                        $val = $val -replace ';\s*', "`n"
                    }
                }
                else {
                    $val = $val -replace ';\s*', "`n"
                    if ($propName -in @('conditions', 'grantControls', 'sessionControls', 'conditions_flat', 'grantControls_flat', 'sessionControls_flat')) {
                        $val = $val -replace '\s*\|\s*', "`n"
                    }
                }
            }

            $newRow[$propName] = $val
        }

        $outRows.Add([pscustomobject]$newRow) | Out-Null
    }

    return [object[]]$outRows.ToArray()
}

function Set-WorksheetWrapByColumnName {
    param(
        [Parameter(Mandatory)]$Worksheet,
        [Parameter(Mandatory)][string[]]$WrapColumns,
        [string[]]$NoWrapColumns
    )

    if (-not $Worksheet -or -not $Worksheet.Dimension) { return }

    $maxCol = $Worksheet.Dimension.End.Column
    $maxRow = $Worksheet.Dimension.End.Row
    $headerMap = @{}

    for ($col = 1; $col -le $maxCol; $col++) {
        $header = [string]$Worksheet.Cells[1, $col].Text
        if (-not [string]::IsNullOrWhiteSpace($header)) {
            $headerMap[$header] = $col
        }
    }

    foreach ($colName in $WrapColumns) {
        if ($headerMap.ContainsKey($colName)) {
            $colIndex = $headerMap[$colName]
            $Worksheet.Cells[2, $colIndex, $maxRow, $colIndex].Style.WrapText = $true
            $Worksheet.Cells[2, $colIndex, $maxRow, $colIndex].Style.VerticalAlignment = 'Top'
        }
    }

    foreach ($colName in $NoWrapColumns) {
        if ($headerMap.ContainsKey($colName)) {
            $colIndex = $headerMap[$colName]
            $Worksheet.Cells[2, $colIndex, $maxRow, $colIndex].Style.WrapText = $false
        }
    }
}

function Test-IsGuid {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $guidValue = [guid]::Empty
    return [guid]::TryParse($Value, [ref]$guidValue)
}

function Invoke-GraphGetSafe {
    param([Parameter(Mandatory)][string]$Uri)

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
        }
        catch {
            $msg = $_.Exception.Message
            if (($msg -match '429') -or ($msg -match 'Too Many Requests') -or ($msg -match '503') -or ($msg -match '504')) {
                Start-Sleep -Milliseconds (200 * $attempt)
                continue
            }
            return $null
        }
    }

    return $null
}

function Invoke-GraphGetPaged {
    # Returns a PSCustomObject with a 'value' array containing all pages of results.
    param([Parameter(Mandatory)][string]$Uri)

    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri

    while (-not [string]::IsNullOrWhiteSpace($next)) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
        if ($null -ne $response -and $response.ContainsKey('value')) {
            foreach ($item in @($response.value)) { [void]$items.Add($item) }
            $next = $response.'@odata.nextLink'
        }
        else {
            if ($null -ne $response) { [void]$items.Add($response) }
            $next = $null
        }
    }

    return [PSCustomObject]@{ value = @($items) }
}

function Resolve-UserUpn {
    param([string]$Id)

    if (-not (Test-IsGuid $Id)) { return $Id }
    if ($script:userById.ContainsKey($Id)) { return $script:userById[$Id] }

    $u = Invoke-GraphGetSafe -Uri "https://graph.microsoft.com/v1.0/users/$Id?`$select=id,userPrincipalName,displayName"
    if (-not $u) {
        $u = Invoke-GraphGetSafe -Uri "https://graph.microsoft.com/v1.0/directoryObjects/$Id"
    }

    $name = if ($u.userPrincipalName) { [string]$u.userPrincipalName } elseif ($u.displayName) { [string]$u.displayName } else { $Id }
    if ($name -ne $Id) { $script:userById[$Id] = $name }
    return $name
}

function Resolve-GroupName {
    param([string]$Id)

    if (-not (Test-IsGuid $Id)) { return $Id }
    if ($script:groupById.ContainsKey($Id)) { return $script:groupById[$Id] }

    $g = Invoke-GraphGetSafe -Uri "https://graph.microsoft.com/v1.0/groups/$Id?`$select=id,displayName"
    if (-not $g) { $g = Invoke-GraphGetSafe -Uri "https://graph.microsoft.com/v1.0/directoryObjects/$Id" }

    $name = if ($g.displayName) { [string]$g.displayName } else { $Id }
    if ($name -ne $Id) { $script:groupById[$Id] = $name }
    return $name
}

function Resolve-RoleName {
    param([string]$Id)

    if (-not (Test-IsGuid $Id)) { return $Id }
    if ($script:roleById.ContainsKey($Id)) { return $script:roleById[$Id] }
    return $Id
}

function Resolve-AppName {
    param([string]$AppId)

    if (-not (Test-IsGuid $AppId)) { return $AppId }
    if ($script:appById.ContainsKey($AppId)) { return $script:appById[$AppId] }

    $spResp = Invoke-GraphGetSafe -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$AppId'&`$select=displayName,appId"
    $sp = if ($spResp -and $spResp.ContainsKey('value')) { @($spResp.value | Select-Object -First 1)[0] } else { $spResp }
    $name = if ($sp -and $sp.displayName) { [string]$sp.displayName } else { $AppId }
    if ($name -ne $AppId) { $script:appById[$AppId] = $name }
    return $name
}

function Resolve-LocationName {
    param([string]$Id)

    if (-not (Test-IsGuid $Id)) { return $Id }
    if ($script:locationById.ContainsKey($Id)) { return $script:locationById[$Id] }
    return $Id
}

function Resolve-IdListToNames {
    param(
        $Values,
        [Parameter(Mandatory)][ValidateSet('User', 'Group', 'Role', 'App', 'Location')][string]$Type
    )

    if ($null -eq $Values) { return @() }
    $items = @($Values)
    $result = New-Object System.Collections.Generic.List[string]

    foreach ($item in $items) {
        $raw = [string]$item
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }

        switch ($Type) {
            'User' { $resolved = Resolve-UserUpn -Id $raw }
            'Group' { $resolved = Resolve-GroupName -Id $raw }
            'Role' { $resolved = Resolve-RoleName -Id $raw }
            'App' { $resolved = Resolve-AppName -AppId $raw }
            'Location' { $resolved = Resolve-LocationName -Id $raw }
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$resolved)) {
            $result.Add([string]$resolved) | Out-Null
        }
    }

    return @($result | Select-Object -Unique)
}

function Resolve-TermsOfUseNames {
    param($Values)

    if ($null -eq $Values) { return @() }
    $items = @($Values)
    $result = New-Object System.Collections.Generic.List[string]

    foreach ($item in $items) {
        $id = [string]$item
        if ([string]::IsNullOrWhiteSpace($id)) { continue }

        if ($script:termsById -and $script:termsById.ContainsKey($id)) {
            $result.Add([string]$script:termsById[$id]) | Out-Null
        }
        else {
            $result.Add($id) | Out-Null
        }
    }

    return @($result | Select-Object -Unique)
}

function Format-GuestsExternalScope {
    param($Scope)

    if ($null -eq $Scope) { return '' }

    if ($Scope -is [string]) {
        try { $Scope = $Scope | ConvertFrom-Json -AsHashtable -Depth 20 } catch { return [string]$Scope }
    }

    $typeMap = @{
        'internalGuest'        = 'internalGuest'
        'b2bCollaborationGuest' = 'b2bGuest'
        'b2bCollaborationMember' = 'b2bMember'
        'b2bDirectConnectUser' = 'b2bDirect'
        'otherExternalUser'    = 'otherExternal'
        'serviceProvider'      = 'serviceProvider'
    }

    $parts = New-Object System.Collections.Generic.List[string]

    $typesRaw = [string]$Scope.guestOrExternalUserTypes
    if (-not [string]::IsNullOrWhiteSpace($typesRaw)) {
        $mapped = @(
            $typesRaw -split ',' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { if ($typeMap.ContainsKey($_)) { $typeMap[$_] } else { $_ } }
        )
        if ($mapped.Count -gt 0) {
            $parts.Add("types=" + ($mapped -join ', ')) | Out-Null
        }
    }

    if ($Scope.externalTenants) {
        $membership = [string]$Scope.externalTenants.membershipKind
        if (-not [string]::IsNullOrWhiteSpace($membership)) {
            $parts.Add("tenants=" + $membership) | Out-Null
        }

        if ($Scope.externalTenants.members) {
            $memberList = @($Scope.externalTenants.members)
            if ($memberList.Count -gt 0) {
                $parts.Add("tenantMembers=" + ($memberList -join ', ')) | Out-Null
            }
        }
    }

    return ($parts -join '; ')
}

function Initialize-LookupMaps {
    Write-Host 'Loading lookup maps (roles/locations/terms)...' -ForegroundColor DarkCyan

    $script:roleById = @{}
    $script:locationById = @{}
    $script:termsById = @{}

    $roleTemplatesResp = Invoke-GraphGetPaged -Uri "https://graph.microsoft.com/v1.0/directoryRoleTemplates?`$select=id,displayName"
    $roleTemplates = if ($roleTemplatesResp.value) { @($roleTemplatesResp.value) } else { @() }
    foreach ($r in $roleTemplates) {
        if ($r.id -and $r.displayName) {
            $script:roleById[[string]$r.id] = [string]$r.displayName
        }
    }

    $roleDefsResp = Invoke-GraphGetPaged -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?`$select=id,displayName"
    $roleDefs = if ($roleDefsResp.value) { @($roleDefsResp.value) } else { @() }
    foreach ($r in $roleDefs) {
        if ($r.id -and $r.displayName -and -not $script:roleById.ContainsKey([string]$r.id)) {
            $script:roleById[[string]$r.id] = [string]$r.displayName
        }
    }

    $namedLocationsResp = Invoke-GraphGetPaged -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations?`$select=id,displayName"
    $namedLocations = if ($namedLocationsResp.value) { @($namedLocationsResp.value) } else { @() }
    foreach ($l in $namedLocations) {
        if ($l.id -and $l.displayName) {
            $script:locationById[[string]$l.id] = [string]$l.displayName
        }
    }

    $termsResp = Invoke-GraphGetSafe -Uri "https://graph.microsoft.com/v1.0/identityGovernance/termsOfUse/agreements?`$select=id,displayName"
    if ($termsResp) {
        $terms = if ($termsResp.ContainsKey('value')) { @($termsResp.value) } else { @($termsResp) }
        foreach ($t in $terms) {
            if ($t.id -and $t.displayName) {
                $script:termsById[[string]$t.id] = [string]$t.displayName
            }
        }
    }

    Write-Host ("Lookup loaded: roles={0}, locations={1}, terms={2}" -f $script:roleById.Count, $script:locationById.Count, $script:termsById.Count) -ForegroundColor DarkGray
}

function Convert-PolicyToReportRow {
    param([Parameter(Mandatory)]$Policy)

    $row = [ordered]@{}

    foreach ($prop in $Policy.PSObject.Properties) {
        $name = [string]$prop.Name
        $value = $prop.Value

        if ($null -eq $value) { $row[$name] = $null; continue }

        if ($value -is [string] -or $value -is [bool] -or $value -is [int] -or $value -is [long] -or
            $value -is [double] -or $value -is [decimal] -or $value -is [datetime] -or $value -is [guid]) {
            $row[$name] = $value
        }
        else {
            $row[$name] = ($value | ConvertTo-Json -Depth 25 -Compress)
        }
    }

    if (-not $row.Contains('policyId') -and $row.Contains('id')) { $row['policyId'] = $row['id'] }
    if (-not $row.Contains('policyName') -and $row.Contains('displayName')) { $row['policyName'] = $row['displayName'] }

    $flattenColumns = @(
        'cond_UserRiskLevels', 'cond_SignInRiskLevels', 'cond_ClientAppTypes', 'cond_InsiderRiskLevels',
        'cond_Applications_Include', 'cond_Applications_Exclude', 'cond_Target_UserActions', 'cond_Target_AuthContext',
        'cond_Users_Include', 'cond_Users_Exclude', 'cond_Groups_Include', 'cond_Groups_Exclude',
        'cond_Roles_Include', 'cond_Roles_Exclude', 'cond_Assign_GuestsExternal_Include', 'cond_Assign_GuestsExternal_Exclude',
        'cond_Assign_Agents_Include', 'cond_Assign_Agents_Exclude', 'cond_Platforms_Include', 'cond_Platforms_Exclude',
        'cond_Locations_Include', 'cond_Locations_Exclude', 'cond_DeviceFilter', 'cond_AuthenticationFlows',
        'grant_Operator', 'grant_BuiltInControls', 'grant_CustomAuthenticationFactors', 'grant_TermsOfUse',
        'grant_AuthStrengthId', 'grant_AuthStrengthDisplayName',
        'session_AppEnforcedRestrictions_IsEnabled', 'session_PersistentBrowser_IsEnabled', 'session_PersistentBrowser_Mode',
        'session_SignInFrequency_IsEnabled', 'session_SignInFrequency_Type', 'session_SignInFrequency_Value',
        'session_SignInFrequency_AuthenticationType', 'session_DisableResilienceDefaults',
        'session_CloudAppSecurity_IsEnabled', 'session_CloudAppSecurity_Type'
    )
    foreach ($fc in $flattenColumns) {
        if (-not $row.Contains($fc)) { $row[$fc] = $null }
    }

    function Join-Values {
        param($Values)
        if ($null -eq $Values) { return '' }
        $items = @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($items.Count -eq 0) { return '' }
        return ($items -join '; ')
    }

    function To-FlatInline {
        param([string]$Text)
        if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
        return ($Text -replace ';\s*', ', ')
    }

    function Build-FlatSummary {
        param([Parameter(Mandatory)][hashtable]$Values)
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($k in $Values.Keys) {
            $v = $Values[$k]
            if ($null -eq $v) { continue }
            $text = [string]$v
            if ([string]::IsNullOrWhiteSpace($text)) { continue }
            $parts.Add(("{0}={1}" -f $k, $text)) | Out-Null
        }
        return ($parts -join ' | ')
    }

    $conditions = $Policy.conditions
    $grantControls = $Policy.grantControls
    $sessionControls = $Policy.sessionControls

    if ($conditions) {
        $row['cond_UserRiskLevels'] = Join-Values $conditions.userRiskLevels
        $row['cond_SignInRiskLevels'] = Join-Values $conditions.signInRiskLevels
        $row['cond_InsiderRiskLevels'] = Join-Values $conditions.insiderRiskLevels
        $row['cond_ClientAppTypes'] = Join-Values $conditions.clientAppTypes
        $row['cond_Applications_Include'] = Join-Values (Resolve-IdListToNames -Values $conditions.applications.includeApplications -Type App)
        $row['cond_Applications_Exclude'] = Join-Values (Resolve-IdListToNames -Values $conditions.applications.excludeApplications -Type App)
        $row['cond_Target_UserActions'] = Join-Values $conditions.applications.includeUserActions
        $row['cond_Target_AuthContext'] = Join-Values $conditions.applications.includeAuthenticationContextClassReferences
        $row['cond_Users_Include'] = Join-Values (Resolve-IdListToNames -Values $conditions.users.includeUsers -Type User)
        $row['cond_Users_Exclude'] = Join-Values (Resolve-IdListToNames -Values $conditions.users.excludeUsers -Type User)
        $row['cond_Groups_Include'] = Join-Values (Resolve-IdListToNames -Values $conditions.users.includeGroups -Type Group)
        $row['cond_Groups_Exclude'] = Join-Values (Resolve-IdListToNames -Values $conditions.users.excludeGroups -Type Group)
        $row['cond_Roles_Include'] = Join-Values (Resolve-IdListToNames -Values $conditions.users.includeRoles -Type Role)
        $row['cond_Roles_Exclude'] = Join-Values (Resolve-IdListToNames -Values $conditions.users.excludeRoles -Type Role)
        $row['cond_Assign_GuestsExternal_Include'] = Format-GuestsExternalScope -Scope $conditions.users.includeGuestsOrExternalUsers
        $row['cond_Assign_GuestsExternal_Exclude'] = Format-GuestsExternalScope -Scope $conditions.users.excludeGuestsOrExternalUsers
        $row['cond_Platforms_Include'] = Join-Values $conditions.platforms.includePlatforms
        $row['cond_Platforms_Exclude'] = Join-Values $conditions.platforms.excludePlatforms
        $row['cond_Locations_Include'] = Join-Values (Resolve-IdListToNames -Values $conditions.locations.includeLocations -Type Location)
        $row['cond_Locations_Exclude'] = Join-Values (Resolve-IdListToNames -Values $conditions.locations.excludeLocations -Type Location)

        if ($conditions.clientApplications) {
            $clientApps = $conditions.clientApplications
            $includeAgents = @()
            $excludeAgents = @()
            if ($clientApps.PSObject.Properties.Name -contains 'includeServicePrincipals') { $includeAgents = @($clientApps.includeServicePrincipals) }
            elseif ($clientApps.PSObject.Properties.Name -contains 'includeApplications') { $includeAgents = @($clientApps.includeApplications) }
            if ($clientApps.PSObject.Properties.Name -contains 'excludeServicePrincipals') { $excludeAgents = @($clientApps.excludeServicePrincipals) }
            elseif ($clientApps.PSObject.Properties.Name -contains 'excludeApplications') { $excludeAgents = @($clientApps.excludeApplications) }
            $row['cond_Assign_Agents_Include'] = Join-Values (Resolve-IdListToNames -Values $includeAgents -Type App)
            $row['cond_Assign_Agents_Exclude'] = Join-Values (Resolve-IdListToNames -Values $excludeAgents -Type App)
            if ([string]::IsNullOrWhiteSpace($row['cond_Assign_Agents_Include']) -and [string]::IsNullOrWhiteSpace($row['cond_Assign_Agents_Exclude'])) {
                $row['cond_Assign_Agents_Include'] = $clientApps | ConvertTo-Json -Depth 20 -Compress
            }
        }
        else {
            $row['cond_Assign_Agents_Include'] = ''
            $row['cond_Assign_Agents_Exclude'] = ''
        }

        if ($conditions.devices) {
            if ($conditions.devices.deviceFilter) { $row['cond_DeviceFilter'] = $conditions.devices.deviceFilter | ConvertTo-Json -Depth 20 -Compress }
            else { $row['cond_DeviceFilter'] = $conditions.devices | ConvertTo-Json -Depth 20 -Compress }
        }
        else { $row['cond_DeviceFilter'] = '' }
        $row['cond_AuthenticationFlows'] = if ($conditions.authenticationFlows) { $conditions.authenticationFlows | ConvertTo-Json -Depth 20 -Compress } else { '' }
    }

    $row['conditions_flat'] = @(
        "UsersIn=$((To-FlatInline $row['cond_Users_Include']))",
        "UsersEx=$((To-FlatInline $row['cond_Users_Exclude']))",
        "GroupsIn=$((To-FlatInline $row['cond_Groups_Include']))",
        "GroupsEx=$((To-FlatInline $row['cond_Groups_Exclude']))",
        "RolesIn=$((To-FlatInline $row['cond_Roles_Include']))",
        "AppsIn=$((To-FlatInline $row['cond_Applications_Include']))",
        "LocationsIn=$((To-FlatInline $row['cond_Locations_Include']))",
        "LocationsEx=$((To-FlatInline $row['cond_Locations_Exclude']))",
        "PlatformsIn=$((To-FlatInline $row['cond_Platforms_Include']))",
        "PlatformsEx=$((To-FlatInline $row['cond_Platforms_Exclude']))",
        "ClientApps=$((To-FlatInline $row['cond_ClientAppTypes']))",
        "UserActions=$((To-FlatInline $row['cond_Target_UserActions']))",
        "AuthContext=$((To-FlatInline $row['cond_Target_AuthContext']))",
        "InsiderRisk=$((To-FlatInline $row['cond_InsiderRiskLevels']))",
        "GuestsExtIncl=$((To-FlatInline $row['cond_Assign_GuestsExternal_Include']))",
        "GuestsExtExcl=$((To-FlatInline $row['cond_Assign_GuestsExternal_Exclude']))",
        "AgentsIncl=$((To-FlatInline $row['cond_Assign_Agents_Include']))",
        "AgentsExcl=$((To-FlatInline $row['cond_Assign_Agents_Exclude']))",
        "DeviceFilter=$((To-FlatInline $row['cond_DeviceFilter']))",
        "AuthFlows=$((To-FlatInline $row['cond_AuthenticationFlows']))"
    ) -join ' | '

    if ($grantControls) {
        $row['grant_Operator'] = [string]$grantControls.operator
        $row['grant_BuiltInControls'] = Join-Values $grantControls.builtInControls
        $row['grant_CustomAuthenticationFactors'] = Join-Values $grantControls.customAuthenticationFactors
        $row['grant_TermsOfUse'] = Join-Values (Resolve-TermsOfUseNames -Values $grantControls.termsOfUse)
        $row['grant_AuthStrengthId'] = [string]$grantControls.authenticationStrength.id
        $row['grant_AuthStrengthDisplayName'] = [string]$grantControls.authenticationStrength.displayName
    }
    $row['grantControls_flat'] = Build-FlatSummary -Values ([ordered]@{
            BuiltIn      = $row['grant_BuiltInControls']
            AuthStrength = $row['grant_AuthStrengthDisplayName']
            Terms        = $row['grant_TermsOfUse']
        })

    if ($sessionControls) {
        $row['session_AppEnforcedRestrictions_IsEnabled'] = if ($sessionControls.applicationEnforcedRestrictions) { $sessionControls.applicationEnforcedRestrictions.isEnabled } else { $null }
        $row['session_PersistentBrowser_IsEnabled'] = if ($sessionControls.persistentBrowser) { $sessionControls.persistentBrowser.isEnabled } else { $null }
        $row['session_PersistentBrowser_Mode'] = if ($sessionControls.persistentBrowser) { [string]$sessionControls.persistentBrowser.mode } else { $null }
        $row['session_SignInFrequency_IsEnabled'] = if ($sessionControls.signInFrequency) { $sessionControls.signInFrequency.isEnabled } else { $null }
        $row['session_SignInFrequency_Type'] = if ($sessionControls.signInFrequency) { [string]$sessionControls.signInFrequency.type } else { $null }
        $row['session_SignInFrequency_Value'] = if ($sessionControls.signInFrequency) { [string]$sessionControls.signInFrequency.value } else { $null }
        $row['session_SignInFrequency_AuthenticationType'] = if ($sessionControls.signInFrequency) { [string]$sessionControls.signInFrequency.authenticationType } else { $null }
        $row['session_DisableResilienceDefaults'] = $sessionControls.disableResilienceDefaults
        $row['session_CloudAppSecurity_IsEnabled'] = if ($sessionControls.cloudAppSecurity) { $sessionControls.cloudAppSecurity.isEnabled } else { $null }
        $row['session_CloudAppSecurity_Type'] = if ($sessionControls.cloudAppSecurity) { [string]$sessionControls.cloudAppSecurity.cloudAppSecurityType } else { $null }
    }
    $row['sessionControls_flat'] = Build-FlatSummary -Values ([ordered]@{
            SIFreqEnabled     = $row['session_SignInFrequency_IsEnabled']
            SIFreqType        = $row['session_SignInFrequency_Type']
            SIFreqValue       = $row['session_SignInFrequency_Value']
            PersistentBrowser = $row['session_PersistentBrowser_Mode']
            CloudAppSecurity  = $row['session_CloudAppSecurity_Type']
        })

    $row['controls_summary'] = @(
        "Grant: Operator=$($row['grant_Operator']); BuiltIn=$($row['grant_BuiltInControls']); AuthStrength=$($row['grant_AuthStrengthDisplayName']); Terms=$($row['grant_TermsOfUse'])",
        "Session: SIFreqEnabled=$($row['session_SignInFrequency_IsEnabled']); SIFreqType=$($row['session_SignInFrequency_Type']); SIFreqValue=$($row['session_SignInFrequency_Value']); PersistentBrowser=$($row['session_PersistentBrowser_Mode']); CloudAppSecurity=$($row['session_CloudAppSecurity_Type'])"
    ) -join ' | '

    if ($row.Contains('conditions')) { $row['conditions_json'] = $row['conditions'] }
    if ($row.Contains('grantControls')) { $row['grantControls_json'] = $row['grantControls'] }
    if ($row.Contains('sessionControls')) { $row['sessionControls_json'] = $row['sessionControls'] }

    $row['conditions'] = $row['conditions_flat']
    $row['grantControls'] = $row['grantControls_flat']
    $row['sessionControls'] = $row['sessionControls_flat']

    return [pscustomobject]$row
}

# Connect to Microsoft Graph
$connectParams = @{
    Scopes = @('Policy.Read.All', 'Directory.Read.All', 'Agreement.Read.All')
}
if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
    $connectParams['TenantId'] = $TenantId
}
Connect-MgGraph @connectParams -ErrorAction Stop | Out-Null

$mgContext = Get-MgContext
Write-Host "Connected: $($mgContext.Account) | Tenant: $($mgContext.TenantId)" -ForegroundColor DarkCyan

$script:userById = @{}
$script:groupById = @{}
$script:appById = @{}
$script:roleById = @{}
$script:locationById = @{}

Write-Host 'Fetching Conditional Access policies...' -ForegroundColor Cyan

$policiesResp = Invoke-GraphGetPaged -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies'
$policies = if ($policiesResp.value) { @($policiesResp.value) } else { @() }

if (-not $policies -or $policies.Count -eq 0) {
    Write-Warning 'No Conditional Access policies found.'
    return
}

Initialize-LookupMaps

$reportRows = @(
    $policies | ForEach-Object { Convert-PolicyToReportRow -Policy $_ }
)

$stateCounts = $policies | Group-Object state | Sort-Object Name

Write-Host "Total policies: $($reportRows.Count)" -ForegroundColor Green
foreach ($stateGroup in $stateCounts) {
    Write-Host ("State '{0}': {1}" -f $stateGroup.Name, $stateGroup.Count) -ForegroundColor DarkGray
}

$sortedRows = $reportRows | Sort-Object state, policyName

$finalSelectColumns = @(
    @{ Name = 'policyName'; Expression = { $_.policyName } },
    @{ Name = 'state'; Expression = { $_.state } },
    @{ Name = 'createdDateTime'; Expression = { $_.createdDateTime } },
    @{ Name = 'modifiedDateTime'; Expression = { $_.modifiedDateTime } },
    @{ Name = 'policyId'; Expression = { $_.policyId } },
    @{ Name = 'inclUsers'; Expression = { $_.cond_Users_Include } },
    @{ Name = 'exclUsers'; Expression = { $_.cond_Users_Exclude } },
    @{ Name = 'inclGuestsExternal'; Expression = { $_.cond_Assign_GuestsExternal_Include } },
    @{ Name = 'exclGuestsExternal'; Expression = { $_.cond_Assign_GuestsExternal_Exclude } },
    @{ Name = 'inclAgents'; Expression = { $_.cond_Assign_Agents_Include } },
    @{ Name = 'exclAgents'; Expression = { $_.cond_Assign_Agents_Exclude } },
    @{ Name = 'inclGroups'; Expression = { $_.cond_Groups_Include } },
    @{ Name = 'exclGroups'; Expression = { $_.cond_Groups_Exclude } },
    @{ Name = 'inclRoles'; Expression = { $_.cond_Roles_Include } },
    @{ Name = 'exclRoles'; Expression = { $_.cond_Roles_Exclude } },
    @{ Name = 'inclApps'; Expression = { $_.cond_Applications_Include } },
    @{ Name = 'exclApps'; Expression = { $_.cond_Applications_Exclude } },
    @{ Name = 'userActions'; Expression = { $_.cond_Target_UserActions } },
    @{ Name = 'authContext'; Expression = { $_.cond_Target_AuthContext } },
    @{ Name = 'inclPlatforms'; Expression = { $_.cond_Platforms_Include } },
    @{ Name = 'exclPlatforms'; Expression = { $_.cond_Platforms_Exclude } },
    @{ Name = 'clientAppTypes'; Expression = { $_.cond_ClientAppTypes } },
    @{ Name = 'deviceFilter'; Expression = { $_.cond_DeviceFilter } },
    @{ Name = 'authenticationFlows'; Expression = { $_.cond_AuthenticationFlows } },
    @{ Name = 'userRiskLevels'; Expression = { $_.cond_UserRiskLevels } },
    @{ Name = 'signInRiskLevels'; Expression = { $_.cond_SignInRiskLevels } },
    @{ Name = 'insiderRiskLevels'; Expression = { $_.cond_InsiderRiskLevels } },
    @{ Name = 'inclLocations'; Expression = { $_.cond_Locations_Include } },
    @{ Name = 'exclLocations'; Expression = { $_.cond_Locations_Exclude } },
    @{ Name = 'grantOperator'; Expression = { $_.grant_Operator } },
    @{ Name = 'grantControls'; Expression = { $_.grantControls_flat } },
    @{ Name = 'sessionControls'; Expression = { $_.sessionControls_flat } }
)

$baseRows = $sortedRows | Select-Object -Property $finalSelectColumns
$finalRows = [object[]]@($baseRows)

$doCsvExport = $ExportCsv.IsPresent -or (-not [string]::IsNullOrWhiteSpace($ExportCsvPath))
$doExcelExport = $ExportExcel.IsPresent -or (-not [string]::IsNullOrWhiteSpace($ExportExcelPath))
$doJsonExport = $ExportJson.IsPresent -or (-not [string]::IsNullOrWhiteSpace($ExportJsonPath))

if ($doCsvExport) {
    $csvOut = Resolve-ExportPath -RequestedPath $ExportCsvPath -DefaultBaseName 'ConditionalAccessPolicies' -Extension 'csv'
    $finalRows | Export-Csv -Path $csvOut -NoTypeInformation -Encoding UTF8
    Write-Host "CSV exported: $csvOut" -ForegroundColor Green
}

if ($doExcelExport) {
    $excelOut = Resolve-ExportPath -RequestedPath $ExportExcelPath -DefaultBaseName 'ConditionalAccessPolicies' -Extension 'xlsx'

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "ExportExcel requested but module 'ImportExcel' is not installed. Install with: Install-Module ImportExcel -Scope CurrentUser"
    }

    Import-Module ImportExcel -ErrorAction Stop

    $excelMultilineColumns = @(
        'inclUsers', 'exclUsers', 'inclGuestsExternal', 'exclGuestsExternal',
        'inclAgents', 'exclAgents', 'inclGroups', 'exclGroups', 'inclRoles', 'exclRoles',
        'inclApps', 'exclApps', 'userActions', 'authContext', 'inclPlatforms', 'exclPlatforms',
        'clientAppTypes', 'deviceFilter', 'authenticationFlows', 'inclLocations', 'exclLocations',
        'insiderRiskLevels', 'grantControls', 'sessionControls'
    )

    $excelRows = Convert-RowsForExcelLineBreaks -Rows $finalRows -Columns $excelMultilineColumns
    $excelRows | Export-Excel -Path $excelOut -WorksheetName 'Policies' -AutoSize -FreezeTopRow -ClearSheet -TableStyle Medium2 | Out-Null

    $excelPackage = Open-ExcelPackage -Path $excelOut
    $ws = $excelPackage.Workbook.Worksheets['Policies']
    if ($ws -and $ws.Dimension) {
        Set-WorksheetWrapByColumnName -Worksheet $ws -WrapColumns $excelMultilineColumns -NoWrapColumns @()
    }

    Close-ExcelPackage $excelPackage
    Write-Host "Excel exported: $excelOut" -ForegroundColor Green
}

if ($doJsonExport) {
    $jsonOut = Resolve-ExportPath -RequestedPath $ExportJsonPath -DefaultBaseName 'ConditionalAccessPolicies' -Extension 'json'
    $policies | ConvertTo-Json -Depth 50 | Set-Content -Path $jsonOut -Encoding UTF8
    Write-Host "JSON exported: $jsonOut" -ForegroundColor Green
}

if (-not $NoGridView) {
    $finalRows | Out-GridView -Title "Conditional Access Policies ($($mgContext.TenantId))"
}
