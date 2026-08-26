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
      happens per item either way. Since Invoke-PnPBatch has no built-in progress
      events, the script manually chunks deletions into groups of $BatchSize (default
      100, SharePoint's own cap) and sends/reports each chunk as it fills up, so you
      see progress batch-by-batch rather than one opaque call at the very end.
    - Timing: each batch chunk is timed individually and reported as "Batch X out of Y"
      with its own elapsed time, plus a total/average across all batches once done. The
      whole script run is also timed with a stopwatch, reported in the final summary.
    - $SendToRecycleBin (optional, default $true) sends removed items to the site's
      Recycle Bin instead of permanently deleting them, giving you a recovery window
      if something unexpected gets removed.

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

    Permanently delete instead of sending to the Recycle Bin (use with caution):
        .\Remove-SharePointDuplicates.ps1 -DryRun:$false -SendToRecycleBin:$false
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
    # $true             = deletions are queued and sent in chunks of $BatchSize via
    #                     New-PnPBatch/Invoke-PnPBatch, which is much faster for large
    #                     numbers of deletions and reduces the chance of throttling.
    #                     Each chunk is sent and reported on as soon as it fills up,
    #                     rather than waiting until the very end.
    #                     CSV logging and per-item status still happen for every row -
    #                     batched rows are marked accordingly and updated once their
    #                     chunk is sent (or if that chunk's send fails).
    # Has no effect in dry run mode, since nothing is deleted either way.
    [bool]$UseBatch = $false,

    # How many deletions to send per batch chunk when -UseBatch is enabled. SharePoint's
    # REST API caps batches at 100 requests, and Invoke-PnPBatch would silently split
    # anything larger into groups of 100 anyway. Chunking manually at this size lets the
    # script report progress after each chunk completes ("Batch 3 of 6 sent..."), since
    # Invoke-PnPBatch itself has no progress events/hooks to surface that internally.
    # Max effective value is 100 - anything higher is capped by SharePoint regardless.
    [int]$BatchSize = 100,

    # ----------------- OPTIONAL: recycle bin instead of permanent delete -----------------
    # $true  = removed items go to the site's Recycle Bin, recoverable for the normal
    #          retention window (default - safer, recommended).
    # $false = items are permanently deleted, bypassing the Recycle Bin entirely.
    # Has no effect in dry run mode, since nothing is deleted either way.
    [bool]$SendToRecycleBin = $true
)

# --------------------------------------------------------------------------------
# Setup
# --------------------------------------------------------------------------------

# Overall script stopwatch - reported at the very end of the run.
$scriptStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

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
# Step 3b: Pre-pass to work out how many batches we'll need (for "Batch X out of Y")
# --------------------------------------------------------------------------------
# This mirrors the keep/remove decision logic used in the main loop below (referenced
# items are kept, plus one survivor per group if none are referenced) but only counts,
# it doesn't delete anything. Only needed when batching a live run.

$totalBatches = 0
if ($UseBatch -and -not $DryRun) {
    $plannedRemovalCount = 0
    foreach ($group in $grouped) {
        $itemsInGroup = $group.Group
        $referencedItems   = @($itemsInGroup | Where-Object { $referencedIds.Contains([int]$_.Id) })
        $unreferencedItems = @($itemsInGroup | Where-Object { -not $referencedIds.Contains([int]$_.Id) })

        if ($referencedItems.Count -eq 0 -and $unreferencedItems.Count -gt 0) {
            # One survivor is spared, the rest of the unreferenced items would be removed
            $plannedRemovalCount += ($unreferencedItems.Count - 1)
        }
        else {
            # At least one referenced item exists, so ALL unreferenced items are removable
            $plannedRemovalCount += $unreferencedItems.Count
        }
    }
    $totalBatches = [math]::Max(1, [math]::Ceiling($plannedRemovalCount / $BatchSize))
    Write-Host "Planned removals: $plannedRemovalCount item(s) across an estimated $totalBatches batch(es)." -ForegroundColor Cyan
}

# --------------------------------------------------------------------------------
# Step 4: Evaluate each duplicate, decide keep/remove, log results
# --------------------------------------------------------------------------------

$csvRows = [System.Collections.Generic.List[object]]::new()
$removedCount = 0
$keptCount = 0
$processedCount = 0

# When batching, deletions are queued into chunks of $BatchSize (SharePoint's own
# limit is 100 requests/batch, and Invoke-PnPBatch would split silently at that size
# anyway - here we chunk manually ourselves so we can report progress after each
# chunk is actually sent, since Invoke-PnPBatch itself has no progress events/hooks).
$currentBatch      = $null
$currentBatchRows  = [System.Collections.Generic.List[object]]::new()
$currentBatchCount = 0
$batchNumber       = 0
$totalBatchedRemoved = 0
$batchTimings      = [System.Collections.Generic.List[double]]::new()

# Sends whatever is currently queued in $currentBatch, updates the CSV status for
# each row in that chunk based on success/failure, and resets for the next chunk.
# Times the send with its own stopwatch and reports "Batch X out of Y" plus elapsed time.
function Invoke-BatchChunk {
    param(
        [Parameter(Mandatory)] $BatchObject,
        [Parameter(Mandatory)] [System.Collections.Generic.List[object]]$Rows,
        [Parameter(Mandatory)] [int]$ChunkNumber,
        [Parameter(Mandatory)] [int]$TotalChunks,
        [Parameter(Mandatory)] [bool]$RecycleBin,
        [Parameter(Mandatory)] [System.Collections.Generic.List[double]]$TimingLog
    )

    if ($Rows.Count -eq 0) { return }

    $batchLabel = if ($TotalChunks -gt 0) { "Batch $ChunkNumber out of $TotalChunks" } else { "Batch $ChunkNumber" }

    Write-Host "`n$batchLabel - sending $($Rows.Count) item(s) to SharePoint ..." -ForegroundColor Cyan

    $batchStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        Invoke-PnPBatch -Batch $BatchObject
        $batchStopwatch.Stop()
        $TimingLog.Add($batchStopwatch.Elapsed.TotalSeconds)

        $recycleNote = if ($RecycleBin) { " (to Recycle Bin)" } else { "" }
        foreach ($row in $Rows) {
            $row.Status = "REMOVED - Not referenced (batch #$ChunkNumber)$recycleNote"
        }
        Write-Host "$batchLabel completed successfully in $($batchStopwatch.Elapsed.ToString('mm\:ss\.ff')) ($($Rows.Count) item(s) removed)" -ForegroundColor Green
    }
    catch {
        $batchStopwatch.Stop()
        $TimingLog.Add($batchStopwatch.Elapsed.TotalSeconds)

        # If the batch call itself fails, we can't be sure which individual items in
        # THIS chunk succeeded vs. failed, so mark all rows in this chunk as errored.
        # Earlier chunks that already completed successfully are unaffected.
        foreach ($row in $Rows) {
            $row.Status = "ERROR - Batch #$ChunkNumber failed: $($_.Exception.Message)"
        }
        Write-Host "$batchLabel FAILED after $($batchStopwatch.Elapsed.ToString('mm\:ss\.ff')): $($_.Exception.Message)" -ForegroundColor Red
    }
}

if ($UseBatch -and -not $DryRun) {
    $currentBatch = New-PnPBatch
    Write-Host "Batch mode enabled: deletions will be sent in chunks of $BatchSize." -ForegroundColor Cyan
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
                # NOTE: -Force is intentionally omitted here. Remove-PnPListItem's -Batch
                # parameter set does not accept -Force ("Parameter set cannot be resolved"
                # error) - batched deletions don't prompt for confirmation anyway, so it's
                # not needed. -Recycle IS supported alongside -Batch.
                if ($SendToRecycleBin) {
                    Remove-PnPListItem -List $TargetListName -Identity $itemId -Batch $currentBatch -Recycle
                } else {
                    Remove-PnPListItem -List $TargetListName -Identity $itemId -Batch $currentBatch
                }
                $status = "QUEUED - Not referenced (pending batch #$($batchNumber + 1))"
                $removedCount++
            }
            catch {
                $status = "ERROR - Failed to queue for batch: $($_.Exception.Message)"
            }
        }
        else {
            try {
                if ($SendToRecycleBin) {
                    Remove-PnPListItem -List $TargetListName -Identity $itemId -Force -Recycle
                } else {
                    Remove-PnPListItem -List $TargetListName -Identity $itemId -Force
                }
                $recycleNote = if ($SendToRecycleBin) { " (to Recycle Bin)" } else { "" }
                $status = "REMOVED - Not referenced$recycleNote"
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

        # Track rows pending in the current chunk. Once the chunk hits $BatchSize,
        # flush it immediately so progress is visible batch-by-batch rather than
        # all at once at the very end.
        if ($UseBatch -and -not $DryRun -and $status -like "QUEUED*") {
            $currentBatchRows.Add($csvRow)
            $currentBatchCount++

            if ($currentBatchCount -ge $BatchSize) {
                $batchNumber++
                Invoke-BatchChunk -BatchObject $currentBatch -Rows $currentBatchRows -ChunkNumber $batchNumber -TotalChunks $totalBatches -RecycleBin $SendToRecycleBin -TimingLog $batchTimings
                $totalBatchedRemoved += $currentBatchRows.Count

                # Reset for the next chunk
                $currentBatch      = New-PnPBatch
                $currentBatchRows  = [System.Collections.Generic.List[object]]::new()
                $currentBatchCount = 0
            }
        }

        Write-Host "  [$processedCount/$totalDuplicateItems] ID $itemId | $DuplicateCheckField = '$fieldValue' | $status"
    }
}

# Clear the progress bar now that processing is complete
Write-Progress -Activity "Processing SharePoint duplicates" -Completed

# --------------------------------------------------------------------------------
# Step 4b: Flush any remaining partial chunk (fewer than $BatchSize items left over)
# --------------------------------------------------------------------------------

if ($UseBatch -and -not $DryRun -and $currentBatchRows.Count -gt 0) {
    $batchNumber++
    Invoke-BatchChunk -BatchObject $currentBatch -Rows $currentBatchRows -ChunkNumber $batchNumber -TotalChunks $totalBatches -RecycleBin $SendToRecycleBin -TimingLog $batchTimings
    $totalBatchedRemoved += $currentBatchRows.Count
}

if ($UseBatch -and -not $DryRun -and $batchNumber -gt 0) {
    $totalBatchTime = ($batchTimings | Measure-Object -Sum).Sum
    $avgBatchTime   = ($batchTimings | Measure-Object -Average).Average
    $totalBatchTimeSpan = [TimeSpan]::FromSeconds($totalBatchTime)
    $avgBatchTimeSpan   = [TimeSpan]::FromSeconds($avgBatchTime)

    Write-Host "`nAll batches complete: $batchNumber batch(es) sent, $totalBatchedRemoved item(s) processed." -ForegroundColor Cyan
    Write-Host "  Total time in batch sends : $($totalBatchTimeSpan.ToString('mm\:ss\.ff'))" -ForegroundColor Cyan
    Write-Host "  Average time per batch    : $($avgBatchTimeSpan.ToString('mm\:ss\.ff'))" -ForegroundColor Cyan
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

$scriptStopwatch.Stop()
Write-Host "  Total script run time   : $($scriptStopwatch.Elapsed.ToString('hh\:mm\:ss\.ff'))" -ForegroundColor Cyan

Disconnect-PnPOnline