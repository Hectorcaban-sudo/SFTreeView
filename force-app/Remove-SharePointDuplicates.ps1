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
    - SAFETY NET: within each group of duplicates, at least one record always
      survives. If none of the duplicates are lookup-referenced, the oldest
      (lowest ID) copy is kept automatically rather than deleting every copy.
    - Writes ALL duplicates found (removed or kept) to a CSV, recording the Item ID and
      the duplicate field's value, plus status info.
    - $DryRun controls whether items are actually deleted or just reported.
    - $UseBatch (optional) queues deletions with PnP's New-PnPBatch/Invoke-PnPBatch
      instead of deleting one item at a time, which is much faster for large numbers
      of deletions and reduces the chance of SharePoint throttling. CSV logging still
      happens per item either way.

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

    Narrow down to a record matching exact text before checking for duplicates
    (e.g. testing against one known record, or an exact subset like a specific customer name):
        .\Remove-SharePointDuplicates.ps1 -SearchText "Acme Corp"

    Search a different field than the duplicate-check field:
        .\Remove-SharePointDuplicates.ps1 -SearchText "12345" -SearchField "EmployeeID"

    Live run using batched deletions (faster, fewer requests, less throttling risk -
    recommended for large numbers of duplicates):
        .\Remove-SharePointDuplicates.ps1 -DryRun:$false -UseBatch:$true
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

    # ----------------- OPTIONAL: search for a particular record -----------------
    # If set, the script will only consider items in the target list whose
    # $SearchField exactly matches this text (case-insensitive, whitespace-trimmed,
    # but NOT a partial/contains match) before running the duplicate check.
    # Leave blank ("") to process the whole list.
    # Useful for testing against one known record, or narrowing down to a known subset.
    [string]$SearchText = "",

    # Which field to search against when $SearchText is provided.
    # Defaults to the same field used for duplicate detection.
    [string]$SearchField = $DuplicateCheckField,

    # ----------------- SAFETY SWITCH -----------------
    # $true  = report only, no deletions (default - safe)
    # $false = actually delete unreferenced duplicates
    [bool]$DryRun = $true,

    # ----------------- OPTIONAL: batch deletions -----------------
    # $false (default) = each deletion is sent to SharePoint immediately, one at a time.
    # $true             = deletions are queued with New-PnPBatch and sent together via
    #                     Invoke-PnPBatch after the loop, which is much faster for large
    #                     numbers of deletions and reduces the chance of throttling.
    #                     CSV logging and per-item status still happen for every row -
    #                     batched rows are marked accordingly and updated if the batch
    #                     as a whole fails.
    # Has no effect in dry run mode, since nothing is deleted either way.
    [bool]$UseBatch = $false
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
# Step 0: Validate that the lookup field actually points to the target list
# --------------------------------------------------------------------------------
# It's easy to mis-configure $LookupListName / $LookupFieldInternalName to point at
# the wrong list. If that happens, "referenced" IDs would be meaningless - either
# wrongly protecting duplicates that aren't really referenced, or wrongly clearing
# duplicates that are. So we check the field's schema before trusting it.

Write-Host "Validating that '$LookupFieldInternalName' on '$LookupListName' points to '$TargetListName' ..." -ForegroundColor Cyan

$targetList = Get-PnPList -Identity $TargetListName -Includes Id -ErrorAction Stop
$lookupField = Get-PnPField -List $LookupListName -Identity $LookupFieldInternalName -ErrorAction Stop

if ($lookupField.TypeAsString -notlike "Lookup*") {
    Write-Error "Field '$LookupFieldInternalName' on '$LookupListName' is not a Lookup field (type: $($lookupField.TypeAsString)). Aborting."
    Disconnect-PnPOnline
    exit 1
}

# The field's SchemaXml contains a List="{GUID}" attribute identifying the source list
$schemaXml = [xml]$lookupField.SchemaXml
$sourceListId = $schemaXml.Field.List

$targetListIdString = $targetList.Id.ToString("B").ToUpper()   # e.g. "{GUID}"
$sourceListIdString = $sourceListId.ToUpper()
if (-not $sourceListIdString.StartsWith("{")) { $sourceListIdString = "{$sourceListIdString}" }

if ($sourceListIdString -ne $targetListIdString) {
    Write-Error "Lookup field '$LookupFieldInternalName' on '$LookupListName' points to list ID $sourceListIdString, which does NOT match '$TargetListName' ($targetListIdString). Check `$LookupListName / `$LookupFieldInternalName. Aborting."
    Disconnect-PnPOnline
    exit 1
}

Write-Host "Confirmed: '$LookupFieldInternalName' points to '$TargetListName'." -ForegroundColor Green

# --------------------------------------------------------------------------------
# Step 1: Pull all items from the target list
# --------------------------------------------------------------------------------

Write-Host "Retrieving items from '$TargetListName' ..." -ForegroundColor Cyan

# Make sure we always pull both the duplicate-check field and the search field
$fieldsToRetrieve = @("Id", $DuplicateCheckField)
if ($SearchField -and ($fieldsToRetrieve -notcontains $SearchField)) {
    $fieldsToRetrieve += $SearchField
}

$targetItems = Get-PnPListItem -List $TargetListName -PageSize 500 -Fields $fieldsToRetrieve

Write-Host "Retrieved $($targetItems.Count) items from '$TargetListName'." -ForegroundColor Cyan

# --------------------------------------------------------------------------------
# Optional: filter down to a particular record / subset by search text
# --------------------------------------------------------------------------------

if (-not [string]::IsNullOrWhiteSpace($SearchText)) {
    Write-Host "Filtering to items where '$SearchField' exactly equals '$SearchText' ..." -ForegroundColor Cyan

    $beforeCount = $targetItems.Count
    $targetItems = $targetItems | Where-Object {
        $val = $_.FieldValues[$SearchField]
        $null -ne $val -and $val.ToString().Trim().Equals($SearchText.Trim(), [System.StringComparison]::OrdinalIgnoreCase)
    }

    Write-Host "Search matched $($targetItems.Count) of $beforeCount item(s)." -ForegroundColor Cyan

    if ($targetItems.Count -eq 0) {
        Write-Host "No items exactly matched '$SearchText' in field '$SearchField'. Exiting." -ForegroundColor Yellow
        Disconnect-PnPOnline
        exit 0
    }
}

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

# Total individual duplicate records across all groups (used for the progress counter)
$totalDuplicateItems = ($grouped | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum

# --------------------------------------------------------------------------------
# Step 4: Evaluate each duplicate, decide keep/remove, log results
# --------------------------------------------------------------------------------

$csvRows = [System.Collections.Generic.List[object]]::new()
$removedCount = 0
$keptCount = 0
$processedCount = 0

# When batching, deletions are queued here and sent together after the loop.
# We keep a reference to each row's CSV object so we can update its Status
# once we know whether the batch as a whole succeeded or failed.
$batch = $null
$batchQueuedRows = [System.Collections.Generic.List[object]]::new()
if ($UseBatch -and -not $DryRun) {
    $batch = New-PnPBatch
    Write-Host "Batch mode enabled: deletions will be queued and sent together." -ForegroundColor Cyan
}

foreach ($group in $grouped) {

    # Determine which items in this group are referenced by a lookup elsewhere
    $itemsInGroup = $group.Group
    $referencedItems   = @($itemsInGroup | Where-Object { $referencedIds.Contains([int]$_.Id) })
    $unreferencedItems = @($itemsInGroup | Where-Object { -not $referencedIds.Contains([int]$_.Id) })

    # Safety net: this group must end up with at least one surviving record.
    # If at least one item is already referenced, that's covered - referenced
    # items are never removed, so the group is safe.
    # If NONE are referenced, we must manually spare one from removal so the
    # group isn't wiped out entirely. We keep the lowest ID (oldest item) as
    # the "survivor" by convention.
    $survivorId = $null
    if ($referencedItems.Count -eq 0 -and $unreferencedItems.Count -gt 0) {
        $survivorId = ($unreferencedItems | Sort-Object Id | Select-Object -First 1).Id
    }

    foreach ($item in $itemsInGroup) {

        $processedCount++

        # Write-Progress pins a bar/status line at the top of the console window
        # so the counter stays visible while everything else scrolls beneath it.
        $percentComplete = if ($totalDuplicateItems -gt 0) { [math]::Round(($processedCount / $totalDuplicateItems) * 100) } else { 0 }
        Write-Progress -Activity "Processing SharePoint duplicates" `
            -Status "$processedCount out of $totalDuplicateItems" `
            -PercentComplete $percentComplete

        $itemId       = $item.Id
        $fieldValue   = $item.FieldValues[$DuplicateCheckField]
        $isReferenced = $referencedIds.Contains([int]$itemId)
        $isSurvivor   = ($null -ne $survivorId -and $itemId -eq $survivorId)

        if ($isReferenced) {
            $status = "KEPT - Referenced by lookup"
            $keptCount++
        }
        elseif ($isSurvivor) {
            $status = "KEPT - Last remaining copy (safety net, not referenced)"
            $keptCount++
        }
        elseif ($DryRun) {
            $status = "WOULD REMOVE - Not referenced (dry run)"
            $removedCount++
        }
        elseif ($UseBatch) {
            try {
                Remove-PnPListItem -List $TargetListName -Identity $itemId -Batch $batch -Force
                $status = "QUEUED - Not referenced (pending batch execution)"
                $removedCount++
            }
            catch {
                $status = "ERROR - Failed to queue for batch: $($_.Exception.Message)"
            }
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

        $csvRow = [PSCustomObject]@{
            ItemId              = $itemId
            DuplicateFieldName  = $DuplicateCheckField
            DuplicateFieldValue = $fieldValue
            Status              = $status
            DryRun              = $DryRun
        }
        $csvRows.Add($csvRow)

        # Keep a reference to rows that are pending batch execution so we can
        # update their Status once Invoke-PnPBatch runs, below.
        if ($UseBatch -and -not $DryRun -and $status -like "QUEUED*") {
            $batchQueuedRows.Add($csvRow)
        }

        Write-Host "  [$processedCount/$totalDuplicateItems] ID $itemId | $DuplicateCheckField = '$fieldValue' | $status"
    }
}

# Clear the progress bar now that processing is complete
Write-Progress -Activity "Processing SharePoint duplicates" -Completed

# --------------------------------------------------------------------------------
# Step 4b: Execute the batch (only runs when -UseBatch was set and this isn't a dry run)
# --------------------------------------------------------------------------------

if ($UseBatch -and -not $DryRun -and $batchQueuedRows.Count -gt 0) {
    Write-Host "`nSending batch of $($batchQueuedRows.Count) queued deletion(s) to SharePoint ..." -ForegroundColor Cyan
    try {
        Invoke-PnPBatch -Batch $batch
        foreach ($row in $batchQueuedRows) {
            $row.Status = "REMOVED - Not referenced (batched)"
        }
        Write-Host "Batch completed successfully." -ForegroundColor Green
    }
    catch {
        # If the batch call itself fails, we can't be sure which individual items
        # succeeded vs. failed, so mark all queued rows as errored for visibility
        # and prompt the user to re-run (ideally without -UseBatch) to confirm state.
        foreach ($row in $batchQueuedRows) {
            $row.Status = "ERROR - Batch deletion failed: $($_.Exception.Message)"
        }
        Write-Host "Batch FAILED: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Re-run the script (consider -UseBatch:`$false) to confirm which items were actually removed." -ForegroundColor Yellow
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