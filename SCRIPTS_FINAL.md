Problem reported: "I've changed usernames and I need any groups, I'm an owner of on my old UPN to be an owner on the new UPN in M365"

# =====================================================================
# Migrate-M365GroupOwnerUPN.ps1
# Inventory M365 Groups owned by old UPN, add new UPN as owner, validate
# Requires: Microsoft.Graph.Groups + Microsoft.Graph.Users modules
# =====================================================================

# --- 0. Connect to Microsoft Graph with the required scopes ----------
# Group.ReadWrite.All  -> list owners + add owner
# Directory.Read.All   -> resolve user objects
# User.ReadBasic.All   -> look up users by UPN
Connect-MgGraph -Scopes "Group.ReadWrite.All", "Directory.Read.All", "User.ReadBasic.All"

# --- 1. Parameters ----------------------------------------------------
$OldUPN = "jane.doe@olddomain.com"      # current UPN (source of inventory)
$NewUPN = "jane.doe@newdomain.com"     # new UPN (target to add as owner)

# --- 2. Resolve both user objects (by UPN) ---------------------------
try {
    $oldUser = Get-MgUser -Filter "userPrincipalName eq '$OldUPN'"
    if (-not $oldUser) { throw "Old UPN '$OldUPN' not found in the directory." }
}
catch {
    # Fallback: try as UserId (handles cases where UPN filter is finicky)
    $oldUser = Get-MgUser -UserId $OldUPN -ErrorAction Stop
}
$oldUserId = $oldUser.Id
Write-Host "Old UPN resolved: $OldUPN  ->  ObjectId: $oldUserId" -ForegroundColor Cyan

try {
    $newUser = Get-MgUser -Filter "userPrincipalName eq '$NewUPN'"
    if (-not $newUser) { throw "New UPN '$NewUPN' not found in the directory." }
}
catch {
    $newUser = Get-MgUser -UserId $NewUPN -ErrorAction Stop
}
$newUserId = $newUser.Id
Write-Host "New UPN resolved: $NewUPN  ->  ObjectId: $newUserId" -ForegroundColor Cyan

# --- 3. INVENTORY: M365 Groups where old UPN is an owner -------------
# M365 Groups have GroupTypes containing "Unified"
Write-Host "`n=== STEP 1: Inventory M365 Groups owned by $OldUPN ===" -ForegroundColor Yellow

$allGroups = Get-MgGroup -All -Property Id, DisplayName, GroupTypes, Mail
$m365Groups = $allGroups | Where-Object { $_.GroupTypes -contains "Unified" }

$ownedGroups = @()
foreach ($grp in $m365Groups) {
    # Get-MgGroupOwner returns directoryObject refs; compare Id to oldUserId
    $owners = Get-MgGroupOwner -GroupId $grp.Id -All
    if ($owners | Where-Object { $_.Id -eq $oldUserId }) {
        $ownedGroups += $grp
    }
}

if ($ownedGroups.Count -eq 0) {
    Write-Host "No M365 Groups found where $OldUPN is an owner. Exiting." -ForegroundColor Green
    return
}

Write-Host "Found $($ownedGroups.Count) M365 Group(s) owned by $OldUPN :" -ForegroundColor Green
$ownedGroups | Select-Object Id, DisplayName, Mail |
    Format-Table -AutoSize | Out-String | Write-Host

# Capture the inventory to a CSV for an audit trail
$inventoryCsv = "M365GroupOwnerInventory_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
$ownedGroups | Select-Object Id, DisplayName, Mail |
    Export-Csv -Path $inventoryCsv -NoTypeInformation
Write-Host "Inventory saved to: $inventoryCsv`n" -ForegroundColor DarkGray

# --- 4. ADD: new UPN as owner on each inventoried group ---------------
Write-Host "=== STEP 2: Adding $NewUPN as owner to each group ===" -ForegroundColor Yellow

# Per the Graph API: POST /groups/{id}/owners/$ref with @odata.id of the user
$addResults = @()
foreach ($grp in $ownedGroups) {
    $result = [PSCustomObject]@{
        GroupId      = $grp.Id
        DisplayName  = $grp.DisplayName
        NewOwnerUPN  = $NewUPN
        Added        = $false
        Error        = $null
    }

    try {
        $newOwnerRef = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/users/$newUserId"
        }
        New-MgGroupOwnerByRef -GroupId $grp.Id -BodyParameter $newOwnerRef -ErrorAction Stop
        $result.Added = $true
        Write-Host "  [OK]   $($grp.DisplayName)" -ForegroundColor Green
    }
    catch {
        # A common non-fatal error: user is already an owner (HTTP 400)
        if ($_.Exception.Message -match "already|existing|conflict") {
            $result.Added = $true
            $result.Error = "Already an owner (treated as success)"
            Write-Host "  [SKIP] $($grp.DisplayName) - already an owner" -ForegroundColor DarkYellow
        }
        else {
            $result.Error = $_.Exception.Message
            Write-Host "  [FAIL] $($grp.DisplayName) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    $addResults += $result
}

$addCsv = "M365GroupOwnerAddResults_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
$addResults | Export-Csv -Path $addCsv -NoTypeInformation
Write-Host "Add results saved to: $addCsv`n" -ForegroundColor DarkGray

# --- 5. VALIDATE: confirm new UPN now appears in each group's owners --
Write-Host "=== STEP 3: Validating $NewUPN is now an owner ===" -ForegroundColor Yellow

# Graph owner additions can take a few seconds to propagate; brief pause
Start-Sleep -Seconds 5

$validationResults = @()
foreach ($grp in $ownedGroups) {
    $vResult = [PSCustomObject]@{
        GroupId     = $grp.Id
        DisplayName = $grp.DisplayName
        NewOwnerUPN = $NewUPN
        IsOwnerNow  = $false
        OldUPNStill = $false
    }

    try {
        $owners = Get-MgGroupOwner -GroupId $grp.Id -All
        if ($owners | Where-Object { $_.Id -eq $newUserId }) {
            $vResult.IsOwnerNow = $true
        }
        if ($owners | Where-Object { $_.Id -eq $oldUserId }) {
            $vResult.OldUPNStill = $true
        }
    }
    catch {
        Write-Host "  [ERR] Could not read owners for $($grp.DisplayName): $($_.Exception.Message)" -ForegroundColor Red
    }

    if ($vResult.IsOwnerNow) {
        Write-Host "  [VALIDATED] $($grp.DisplayName) - new UPN is owner" -ForegroundColor Green
    } else {
        Write-Host "  [NOT FOUND] $($grp.DisplayName) - new UPN NOT in owners list" -ForegroundColor Red
    }
    $validationResults += $vResult
}

$validationCsv = "M365GroupOwnerValidation_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
$validationResults | Export-Csv -Path $validationCsv -NoTypeInformation
Write-Host "Validation results saved to: $validationCsv`n" -ForegroundColor DarkGray

# --- 6. Summary ------------------------------------------------------
Write-Host "=== SUMMARY ===" -ForegroundColor Cyan
Write-Host "Groups inventoried : $($ownedGroups.Count)"
Write-Host "Add attempts       : $($addResults.Count)"
Write-Host "Successfully added : $(($addResults | Where-Object Added).Count)"
Write-Host "Validated as owner : $(($validationResults | Where-Object IsOwnerNow).Count)"
Write-Host "Old UPN still owner: $(($validationResults | Where-Object OldUPNStill).Count) (expected - script does NOT remove the old UPN)"
Write-Host ""
Write-Host "Output files:" -ForegroundColor Cyan
Write-Host "  $inventoryCsv"
Write-Host "  $addCsv"
Write-Host "  $validationCsv"
