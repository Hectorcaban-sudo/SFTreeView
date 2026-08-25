<#
.SYNOPSIS
    Finds duplicate items in a SharePoint list based on a specified field, and removes
    duplicates that are NOT referenced by a lookup field in another (dependent) list.

.DESCRIPTION
    - Connects to a SharePoint site using PnP PowerShell.
    - Retrieves all items from the target list.
    - Groups items by the field you want to check for duplicates (e.g. "Title", "EmployeeID").
    - For each duplicate item, checks whether its ID is referenced anywhere in a lookup
      field on another list. If it IS referenced, the item is SKIPPED (kept, even if it's
      a duplicate) to avoid breaking the lookup relationship.
    - If it is NOT referenced anywhere, it's a removal candidate.
    - Writes ALL duplicates found (removed or kept) to a CSV, recording the Item ID and
      the duplicate field's value, plus status info.
    - $DryRun controls whether items are actually deleted or just reported.

.NOTES
    Requires the PnP.PowerShell module:
        Install-Module PnP.PowerShell -Scope CurrentUser

    Assumptions (adjust the CONFIG block below to match your environment):
    - Authentication uses interactive login (Connect-PnPOnline -Interactive).
      Swap this out for -ClientId/-Thumbprint/-Tenant app-only auth if running unattended.
    - The "duplicate check" field and the "lookup" field are both configurable below.
    - A lookup field's value in SharePoint (via PnP) typically comes back as an object
      with an .Id (or .LookupId) property, or a collection of such objects for multi-value
      lookups. The script handles both single-value and multi-value lookup fields.

.EXAMPLE
    Dry run (default) - just reports what would happen, no deletions, still writes CSV:
        .\Remove-SharePointDuplicates.ps1

    Live run - actually deletes unreferenced duplicates:
        .\Remove-SharePointDuplicates.ps1 -DryRun:$false
#>

[CmdletBinding()]
param(
    # ----------------- CONFIG: adjust these for your environment -----------------

    # Site containing the list you want to de-duplicate
    [string]$SiteUrl = "https://yourtenant.sharepoint.com/sites/YourSite",

    # The list to scan for duplicates
    [string]$TargetListName = "YourTargetList",

    # The internal field name on the target list used to detect duplicates
    # (e.g. "Title", "EmployeeID", "SKU"). Must be the INTERNAL name, not display name.
    [string]$DuplicateCheckField = "Title",

    # The list that might reference the target list's items via a lookup field
    [string]$LookupListName = "YourLookupList",

    # The internal name of the lookup field on $LookupListName that points back
    # to $TargetListName
    [string]$LookupFieldInternalName = "TargetListLookup",

    # Where to write the CSV of duplicates found
    [string]$OutputCsvPath = ".\SharePointDuplicates_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    # ----------------- SAFETY SWITCH -----------------
    # $true  = report only, no deletions (default - safe)
    # $false = actually delete unreferenced duplicates
    [bool]$DryRun = $true
)

# --------------------------------------------------------------------------------
# Setup
# --------------------------------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "PnP.PowerShell module not found. Install it with: Install-Module PnP.PowerShell -Scope CurrentUser"
    exit 1
}

Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

if ($DryRun) {
    Write-Host "*** DRY RUN MODE: No items will be deleted. ***" -ForegroundColor Yellow
} else {
    Write-Host "*** LIVE MODE: Unreferenced duplicates WILL BE DELETED. ***" -ForegroundColor Red
}

# --------------------------------------------------------------------------------
# Step 1: Pull all items from the target list
# --------------------------------------------------------------------------------

Write-Host "Retrieving items from '$TargetListName' ..." -ForegroundColor Cyan
$targetItems = Get-PnPListItem -List $TargetListName -PageSize 500 -Fields "Id", $DuplicateCheckField

Write-Host "Retrieved $($targetItems.Count) items from '$TargetListName'." -ForegroundColor Cyan

# --------------------------------------------------------------------------------
# Step 2: Pull all items from the lookup list and build a set of referenced IDs
# --------------------------------------------------------------------------------

Write-Host "Retrieving items from '$LookupListName' to determine which target IDs are referenced ..." -ForegroundColor Cyan
$lookupListItems = Get-PnPListItem -List $LookupListName -PageSize 500 -Fields $LookupFieldInternalName

$referencedIds = [System.Collections.Generic.HashSet[int]]::new()

foreach ($li in $lookupListItems) {
    $lookupValue = $li.FieldValues[$LookupFieldInternalName]

    if ($null -eq $lookupValue) { continue }

    # Multi-value lookup fields come back as a collection; single-value as one object
    if ($lookupValue -is [System.Collections.IEnumerable] -and -not ($lookupValue -is [string])) {
        foreach ($lv in $lookupValue) {
            $id = $lv.LookupId
            if (-not $id) { $id = $lv.Id }
            if ($id) { [void]$referencedIds.Add([int]$id) }
        }
    } else {
        $id = $lookupValue.LookupId
        if (-not $id) { $id = $lookupValue.Id }
        if ($id) { [void]$referencedIds.Add([int]$id) }
    }
}

Write-Host "Found $($referencedIds.Count) unique item IDs referenced via lookup in '$LookupListName'." -ForegroundColor Cyan

# --------------------------------------------------------------------------------
# Step 3: Group target list items by the duplicate-check field
# --------------------------------------------------------------------------------

$grouped = $targetItems | Group-Object { $_.FieldValues[$DuplicateCheckField] } | Where-Object { $_.Count -gt 1 }

Write-Host "Found $($grouped.Count) value(s) of '$DuplicateCheckField' with duplicates." -ForegroundColor Cyan

# --------------------------------------------------------------------------------
# Step 4: Evaluate each duplicate, decide keep/remove, log results
# --------------------------------------------------------------------------------

$csvRows = [System.Collections.Generic.List[object]]::new()
$removedCount = 0
$keptCount = 0

foreach ($group in $grouped) {

    foreach ($item in $group.Group) {

        $itemId       = $item.Id
        $fieldValue   = $item.FieldValues[$DuplicateCheckField]
        $isReferenced = $referencedIds.Contains([int]$itemId)

        if ($isReferenced) {
            $status = "KEPT - Referenced by lookup"
            $keptCount++
        }
        elseif ($DryRun) {
            $status = "WOULD REMOVE - Not referenced (dry run)"
            $removedCount++
        }
        else {
            try {
                Remove-PnPListItem -List $TargetListName -Identity $itemId -Force
                $status = "REMOVED - Not referenced"
                $removedCount++
            }
            catch {
                $status = "ERROR - Failed to remove: $($_.Exception.Message)"
            }
        }

        $csvRows.Add([PSCustomObject]@{
            ItemId              = $itemId
            DuplicateFieldName  = $DuplicateCheckField
            DuplicateFieldValue = $fieldValue
            Status              = $status
            DryRun              = $DryRun
        })

        Write-Host "  ID $itemId | $DuplicateCheckField = '$fieldValue' | $status"
    }
}

# --------------------------------------------------------------------------------
# Step 5: Export CSV
# --------------------------------------------------------------------------------

if ($csvRows.Count -gt 0) {
    $csvRows | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV written to: $OutputCsvPath" -ForegroundColor Green
} else {
    Write-Host "`nNo duplicates found - no CSV written." -ForegroundColor Green
}

Write-Host "`nSummary:" -ForegroundColor Cyan
Write-Host "  Duplicate records found : $($csvRows.Count)"
Write-Host "  Kept (referenced)       : $keptCount"
Write-Host "  Removed / would remove  : $removedCount"
Write-Host "  Mode                    : $(if ($DryRun) { 'DRY RUN' } else { 'LIVE' })"

Disconnect-PnPOnline
