# Connect to Microsoft Graph
try {
    Connect-MgGraph -TenantId <#Insert Tenant ID #> -Scopes User.Read.All, Organization.Read.All
} catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit
}

# Connect to SharePoint Online
try {
    Connect-SPOService -Url <#Insert SPO Admin URL #>
} catch {
    Write-Error "Failed to connect to SharePoint Online: $_"
    exit
}

# Initialize the list and counters
$list = @()
$i = 0
$j = 0

# Get licensed users
$users = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable licensedUserCount -All -Select UserPrincipalName,DisplayName,AssignedLicenses 
# Total licensed users count
$count = $users.Count

foreach ($u in $users) {
    $i++
    $j++
    Write-Host "$j/$count"

    $upn = $u.UserPrincipalName
    $list += $upn

    if ($i -eq 199) {
        # We reached the limit
        Write-Host "Batch limit reached, requesting provision for the current batch"
        Request-SPOPersonalSite -UserEmails $list -NoWait
        Start-Sleep -Milliseconds 655
        $list = @()
        $i = 0
    }
}

# Process any remaining users
if ($i -gt 0) {
    Request-SPOPersonalSite -UserEmails $list -NoWait
}

Write-Host "Completed OneDrive Provisioning for $j users"
