#Requires -Module Microsoft.Graph.Authentication
#Requires -Module ImportExcel
<#
.SYNOPSIS
    Produces a comprehensive report of Entra ID role assignments and PIM group memberships.
.DESCRIPTION
    Connects to Microsoft Graph and builds a unified report that includes:
        - Eligible Entra ID role assignments (direct user and group-based)
        - Active/direct Entra ID role assignments (user, group, PIM group, service principal)
        - PIM group eligible and active membership schedules
    Each row represents a single user-to-role or user-to-PIM-group assignment, with the
    assignment type (Eligible/Active/Direct), scope (Directory or Admin Unit), and
    company/group context. The report is exported to Excel and shown in a grid view.
.AUTHOR
    Daniel Petri
.EXAMPLE
    .\Get-RoleAndPIMGroupAssignments.ps1
    Connects interactively and exports the assignments report to the system temp folder.
.EXAMPLE
    .\Get-RoleAndPIMGroupAssignments.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -OutputPath "C:\Reports\Assignments.xlsx"
    Connects to the specified tenant and exports to the given path.
.NOTES
    Requires: Microsoft.Graph.Authentication, ImportExcel
    Version: 1.0.0
    Permissions required (delegated):
        RoleManagement.Read.Directory,
        RoleEligibilitySchedule.Read.Directory, RoleAssignmentSchedule.Read.Directory,
        PrivilegedEligibilitySchedule.Read.AzureADGroup,
        PrivilegedAssignmentSchedule.Read.AzureADGroup, Group.Read.All,
        User.Read.All, Directory.Read.All
#>

param(
    [string]$TenantId,

    [string]$OutputPath = (Join-Path $env:TEMP "$(Get-Date -Format 'yyyy-MM-dd')-AllAssignmentsReport.xlsx")
)

# Connect to Microsoft Graph
$connectParams = @{
    Scopes = @(
        'RoleManagement.Read.Directory',
        'RoleEligibilitySchedule.Read.Directory',
        'RoleAssignmentSchedule.Read.Directory',
        'PrivilegedEligibilitySchedule.Read.AzureADGroup',
        'PrivilegedAssignmentSchedule.Read.AzureADGroup',
        'Group.Read.All',
        'User.Read.All',
        'Directory.Read.All'
    )
}
if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
    $connectParams['TenantId'] = $TenantId
}
Connect-MgGraph @connectParams -ErrorAction Stop | Out-Null

# Helper: retrieve all pages from a Graph endpoint
function Invoke-GraphGetAllPages {
    param([Parameter(Mandatory)][string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
        if ($response.value) { foreach ($item in $response.value) { $items.Add($item) } }
        $Uri = $response.'@odata.nextLink'
    } while ($Uri)
    return $items
}

#region Load reference data
$allUsersRaw = Invoke-GraphGetAllPages -Uri 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,companyName'
$allUsers = @{}
foreach ($u in $allUsersRaw) { $allUsers[$u.id] = $u }
if ($allUsers.Count -eq 0) { Write-Warning 'No users retrieved from Microsoft Graph.' }
Write-Host "Users loaded: $($allUsers.Count)"

$allGroupsRaw = Invoke-GraphGetAllPages -Uri 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName'
$allGroups = @{}
foreach ($g in $allGroupsRaw) { $allGroups[$g.id] = $g }
if ($allGroups.Count -eq 0) { Write-Warning 'No groups retrieved from Microsoft Graph.' }
Write-Host "Groups loaded: $($allGroups.Count)"

$adminUnits = Invoke-GraphGetAllPages -Uri 'https://graph.microsoft.com/v1.0/directory/administrativeUnits?$select=id,displayName'
Write-Host "Administrative units loaded: $($adminUnits.Count)"

$pimGroups = Invoke-GraphGetAllPages -Uri 'https://graph.microsoft.com/v1.0/groups?$filter=isAssignableToRole eq true&$select=id,displayName'
Write-Host "PIM groups loaded: $($pimGroups.Count)"
#endregion

# Helper: resolve directory scope ID to a human-readable name
function Resolve-ScopeName {
    param([string]$ScopeId)
    if ($ScopeId -eq '/') { return 'Directory' }
    $auId = $ScopeId.Split('/')[-1]
    return ($adminUnits | Where-Object { $_.id -eq $auId } | Select-Object -First 1).displayName
}

#region PIM group eligible assignments
$pimGroupEligibleAssignments = [System.Collections.Generic.List[object]]::new()
$count = 0
Write-Host 'Retrieving PIM group eligible assignments...'
foreach ($group in $pimGroups) {
    $count++
    Write-Host "  Processing $count of $($pimGroups.Count): $($group.displayName)"
    $eligibleMembers = Invoke-GraphGetAllPages -Uri "https://graph.microsoft.com/beta/identityGovernance/privilegedAccess/group/eligibilityScheduleInstances?`$filter=groupId eq '$($group.id)'"
    foreach ($member in $eligibleMembers) {
        if ($allUsers.ContainsKey($member.principalId)) {
            $pimGroupEligibleAssignments.Add([PSCustomObject][Ordered]@{
                    GroupId          = $group.id
                    Group            = $group.displayName
                    NestedGroup      = ''
                    User             = $allUsers[$member.principalId].displayName
                    UPN              = $allUsers[$member.principalId].userPrincipalName
                    Company          = $allUsers[$member.principalId].companyName
                    ServicePrincipal = ''
                    StartDate        = $member.startDateTime
                    EndDate          = if ($null -eq $member.endDateTime) { 'Permanent' } else { $member.endDateTime }
                    Assignment       = 'Eligible'
                })
        }
        elseif ($allGroups.ContainsKey($member.principalId)) {
            $nestedMembers = Invoke-GraphGetAllPages -Uri "https://graph.microsoft.com/v1.0/groups/$($member.principalId)/transitiveMembers?`$select=id,displayName,userPrincipalName,companyName"
            foreach ($nm in $nestedMembers) {
                if ($allUsers.ContainsKey($nm.id)) {
                    $pimGroupEligibleAssignments.Add([PSCustomObject][Ordered]@{
                            GroupId          = $group.id
                            Group            = $group.displayName
                            NestedGroup      = $allGroups[$member.principalId].displayName
                            User             = $allUsers[$nm.id].displayName
                            UPN              = $allUsers[$nm.id].userPrincipalName
                            Company          = $allUsers[$nm.id].companyName
                            ServicePrincipal = ''
                            StartDate        = $member.startDateTime
                            EndDate          = if ($null -eq $member.endDateTime) { 'Permanent' } else { $member.endDateTime }
                            Assignment       = 'Eligible'
                        })
                }
                else {
                    Write-Warning "Unknown nested member '$($nm.id)' in PIM group '$($group.displayName)'"
                }
            }
        }
        else {
            Write-Warning "Unknown principal '$($member.principalId)' in PIM group '$($group.displayName)'"
        }
    }
}
Write-Host "PIM group eligible assignments: $($pimGroupEligibleAssignments.Count)"
#endregion

#region PIM group active assignments
$pimGroupActiveAssignments = [System.Collections.Generic.List[object]]::new()
$count = 0
Write-Host 'Retrieving PIM group active assignments...'
foreach ($group in $pimGroups) {
    $count++
    Write-Host "  Processing $count of $($pimGroups.Count): $($group.displayName)"
    $activeMembers = Invoke-GraphGetAllPages -Uri "https://graph.microsoft.com/beta/identityGovernance/privilegedAccess/group/assignmentScheduleInstances?`$filter=groupId eq '$($group.id)' and assignmentType eq 'Assigned'"
    foreach ($member in $activeMembers) {
        if ($allUsers.ContainsKey($member.principalId)) {
            $pimGroupActiveAssignments.Add([PSCustomObject][Ordered]@{
                    GroupId          = $group.id
                    Group            = $group.displayName
                    NestedGroup      = ''
                    User             = $allUsers[$member.principalId].displayName
                    UPN              = $allUsers[$member.principalId].userPrincipalName
                    Company          = $allUsers[$member.principalId].companyName
                    ServicePrincipal = ''
                    StartDate        = $member.startDateTime
                    EndDate          = if ($null -eq $member.endDateTime) { 'Permanent' } else { $member.endDateTime }
                    Assignment       = 'Active'
                })
        }
        elseif ($allGroups.ContainsKey($member.principalId)) {
            $nestedMembers = Invoke-GraphGetAllPages -Uri "https://graph.microsoft.com/v1.0/groups/$($member.principalId)/transitiveMembers?`$select=id,displayName,userPrincipalName,companyName"
            foreach ($nm in $nestedMembers) {
                if ($allUsers.ContainsKey($nm.id)) {
                    $pimGroupActiveAssignments.Add([PSCustomObject][Ordered]@{
                            GroupId          = $group.id
                            Group            = $group.displayName
                            NestedGroup      = $allGroups[$member.principalId].displayName
                            User             = $allUsers[$nm.id].displayName
                            UPN              = $allUsers[$nm.id].userPrincipalName
                            Company          = $allUsers[$nm.id].companyName
                            ServicePrincipal = ''
                            StartDate        = $member.startDateTime
                            EndDate          = if ($null -eq $member.endDateTime) { 'Permanent' } else { $member.endDateTime }
                            Assignment       = 'Active'
                        })
                }
                else {
                    Write-Warning "Unknown nested member '$($nm.id)' in PIM group '$($group.displayName)'"
                }
            }
        }
        else {
            Write-Warning "Unknown principal '$($member.principalId)' in PIM group '$($group.displayName)'"
        }
    }
}
Write-Host "PIM group active assignments: $($pimGroupActiveAssignments.Count)"
#endregion

#region Eligible role assignments
$allAssignmentsReport = [System.Collections.Generic.List[object]]::new()
$eligibleAssignments = Invoke-GraphGetAllPages -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?$expand=principal,roleDefinition'
$count = 0
Write-Host 'Processing eligible role assignments...'
foreach ($assignment in $eligibleAssignments) {
    $count++
    Write-Host "  Processing $count of $($eligibleAssignments.Count): $($assignment.roleDefinition.displayName)"
    $scopeName = Resolve-ScopeName -ScopeId $assignment.directoryScopeId
    if ($assignment.principal.'@odata.type' -eq '#microsoft.graph.user') {
        $allAssignmentsReport.Add([PSCustomObject][Ordered]@{
                Role             = $assignment.roleDefinition.displayName
                Target           = 'User'
                Group            = ''
                NestedGroup      = ''
                User             = $assignment.principal.displayName
                UPN              = $assignment.principal.userPrincipalName
                Company          = $assignment.principal.companyName
                ServicePrincipal = ''
                StartDate        = $assignment.startDateTime
                EndDate          = if ($null -eq $assignment.endDateTime) { 'Permanent' } else { $assignment.endDateTime }
                Assignment       = 'Eligible'
                Scope            = $scopeName
            })
    }
    elseif ($assignment.principal.'@odata.type' -eq '#microsoft.graph.group') {
        $groupMembers = Invoke-GraphGetAllPages -Uri "https://graph.microsoft.com/v1.0/groups/$($assignment.principalId)/transitiveMembers?`$select=id,displayName,userPrincipalName,companyName"
        foreach ($member in ($groupMembers | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' })) {
            $allAssignmentsReport.Add([PSCustomObject][Ordered]@{
                    Role             = $assignment.roleDefinition.displayName
                    Target           = 'Group'
                    Group            = $assignment.principal.displayName
                    NestedGroup      = ''
                    User             = $member.displayName
                    UPN              = $member.userPrincipalName
                    Company          = $member.companyName
                    ServicePrincipal = ''
                    StartDate        = $assignment.startDateTime
                    EndDate          = if ($null -eq $assignment.endDateTime) { 'Permanent' } else { $assignment.endDateTime }
                    Assignment       = 'Eligible'
                    Scope            = $scopeName
                })
        }
    }
    else {
        Write-Warning "Unhandled eligible assignment principal type: $($assignment.principal.'@odata.type')"
    }
}
Write-Host "Eligible role assignments processed: $count"
#endregion

#region Direct/active role assignments
$directAssignments = Invoke-GraphGetAllPages -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentSchedules?$expand=principal,roleDefinition'
$directAssignments = @($directAssignments | Where-Object { $_.assignmentType -ne 'Activated' })
$count = 0
Write-Host 'Processing direct role assignments...'
foreach ($assignment in $directAssignments) {
    $count++
    Write-Host "  Processing $count of $($directAssignments.Count): $($assignment.roleDefinition.displayName)"
    $scopeName = Resolve-ScopeName -ScopeId $assignment.directoryScopeId
    $principalType = $assignment.principal.'@odata.type'

    if ($principalType -eq '#microsoft.graph.user') {
        $allAssignmentsReport.Add([PSCustomObject][Ordered]@{
                Role             = $assignment.roleDefinition.displayName
                Target           = 'User'
                Group            = ''
                NestedGroup      = ''
                User             = $assignment.principal.displayName
                UPN              = $assignment.principal.userPrincipalName
                Company          = $assignment.principal.companyName
                ServicePrincipal = ''
                StartDate        = $assignment.startDateTime
                EndDate          = if ($null -eq $assignment.endDateTime) { 'Permanent' } else { $assignment.endDateTime }
                Assignment       = 'Direct'
                Scope            = $scopeName
            })
    }
    elseif ($principalType -eq '#microsoft.graph.servicePrincipal') {
        $allAssignmentsReport.Add([PSCustomObject][Ordered]@{
                Role             = $assignment.roleDefinition.displayName
                Target           = 'ServicePrincipal'
                Group            = ''
                NestedGroup      = ''
                User             = ''
                UPN              = ''
                Company          = ''
                ServicePrincipal = $assignment.principal.displayName
                StartDate        = $assignment.startDateTime
                EndDate          = if ($null -eq $assignment.endDateTime) { 'Permanent' } else { $assignment.endDateTime }
                Assignment       = 'Direct'
                Scope            = $scopeName
            })
    }
    elseif ($principalType -eq '#microsoft.graph.group') {
        # Check if this group is a PIM-enabled role-assignable group
        $eligibleFromGroup = @($pimGroupEligibleAssignments | Where-Object { $_.GroupId -eq $assignment.principalId })
        $activeFromGroup = @($pimGroupActiveAssignments   | Where-Object { $_.GroupId -eq $assignment.principalId })

        if ($eligibleFromGroup.Count -gt 0) {
            foreach ($member in $eligibleFromGroup) {
                $allAssignmentsReport.Add([PSCustomObject][Ordered]@{
                        Role             = $assignment.roleDefinition.displayName
                        Target           = 'PIM Group'
                        Group            = $assignment.principal.displayName
                        NestedGroup      = $member.NestedGroup
                        User             = $member.User
                        UPN              = $member.UPN
                        Company          = $member.Company
                        ServicePrincipal = ''
                        StartDate        = $member.StartDate
                        EndDate          = $member.EndDate
                        Assignment       = 'Eligible'
                        Scope            = $scopeName
                    })
            }
        }
        if ($activeFromGroup.Count -gt 0) {
            foreach ($member in $activeFromGroup) {
                $allAssignmentsReport.Add([PSCustomObject][Ordered]@{
                        Role             = $assignment.roleDefinition.displayName
                        Target           = 'PIM Group'
                        Group            = $assignment.principal.displayName
                        NestedGroup      = $member.NestedGroup
                        User             = $member.User
                        UPN              = $member.UPN
                        Company          = $member.Company
                        ServicePrincipal = ''
                        StartDate        = $member.StartDate
                        EndDate          = $member.EndDate
                        Assignment       = 'Active'
                        Scope            = $scopeName
                    })
            }
        }
        if ($eligibleFromGroup.Count -eq 0 -and $activeFromGroup.Count -eq 0) {
            # Regular group (not PIM-managed) - expand transitive members
            $groupMembers = Invoke-GraphGetAllPages -Uri "https://graph.microsoft.com/v1.0/groups/$($assignment.principalId)/transitiveMembers?`$select=id,displayName,userPrincipalName,companyName"
            foreach ($member in ($groupMembers | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' })) {
                $allAssignmentsReport.Add([PSCustomObject][Ordered]@{
                        Role             = $assignment.roleDefinition.displayName
                        Target           = 'Group'
                        Group            = $assignment.principal.displayName
                        NestedGroup      = ''
                        User             = $member.displayName
                        UPN              = $member.userPrincipalName
                        Company          = $member.companyName
                        ServicePrincipal = ''
                        StartDate        = $assignment.startDateTime
                        EndDate          = if ($null -eq $assignment.endDateTime) { 'Permanent' } else { $assignment.endDateTime }
                        Assignment       = 'Direct'
                        Scope            = $scopeName
                    })
            }
        }
    }
    else {
        Write-Warning "Unhandled direct assignment principal type: $principalType for role '$($assignment.roleDefinition.displayName)'"
    }
}
Write-Host "Direct role assignments processed: $count"
#endregion

#region Output
Write-Host "Total report rows: $($allAssignmentsReport.Count)" -ForegroundColor Green
$allAssignmentsReport | Out-GridView -Title 'Role and PIM Group Assignments'
$allAssignmentsReport | Export-Excel -Path $OutputPath -TableStyle Medium2 -AutoSize -Show
Write-Host "Exported to: $OutputPath" -ForegroundColor Green
#endregion
