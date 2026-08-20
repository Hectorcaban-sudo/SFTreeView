# SharePoint Tree List (LWC + Apex)

A Lightning Web Component that mimics a Kendo React TreeList: hierarchical
rows with expand/collapse, per-column filters, paging, a loading spinner,
and file upload into the SharePoint library — all backed by your custom
Apex class calling your external REST API.

## Files

```
force-app/main/default/classes/
  SharePointTreeListController.cls        Apex controller (callouts)
  SharePointTreeListController.cls-meta.xml

force-app/main/default/lwc/sharePointTreeList/
  sharePointTreeList.js                   Tree/filter/paging/upload logic
  sharePointTreeList.html                 Markup
  sharePointTreeList.css                  Styling
  sharePointTreeList.js-meta.xml          Component exposure config
```

## Passing Project Number / Task Order Number from the current record

The component reads two values off the **current Salesforce record** and
sends them to the REST API on every fetch and every upload:

- It must be placed on a **Record Page** (any object) — `recordId` and
  `objectApiName` are then auto-injected by the platform.
- In the **Lightning App Builder**, when you drag the component onto a
  record page, its property panel has two fields: **"Project Number
  field API name"** and **"Task Order Number field API name"** — set
  these to the actual API names of the fields on that object (defaults
  are `Project_Number__c` / `Task_Order_Number__c`, almost certainly not
  your real field names — change them).
- The LWC uses `lightning/uiRecordApi`'s `getRecord` to read those two
  fields off the record, then calls Apex only once it has them.
- In Apex, `getLiveFileRecords()` currently sends them as query string
  params (`?projectNumber=...&taskOrderNumber=...`) on the GET request,
  and `uploadFile()` includes them in the upload JSON body. If your REST
  API expects them somewhere else (headers, a different body shape,
  different param names), that's the one place to change it.
- If the field API names are misconfigured (field doesn't exist / no
  access), the component shows an inline error instead of guessing.
- If the component is placed somewhere with no record context (a Home or
  App page), it simply loads without project/task order values.

## How the tree is actually built (important)

The SharePoint library itself has **no real folders** — every record the
REST API returns is a flat file. The tree you see is entirely virtual,
built like this in `SharePointTreeListController.cls`:

1. **`FOLDER_LIST`** is a hardcoded map of the fixed master folder set
   (taken from `Contracts_Folders.csv`: Base Contract, Closeout, CPARs,
   Modifications, Custom Folder, TBD, etc.). These always render as root
   rows, even at 0 files.
2. Each file record has a `Contracts_x0020_Subfolder` field; its `.Value`
   is matched (case-insensitively) against `FOLDER_LIST` to decide which
   root folder the file belongs under. Anything that doesn't match lands
   in a synthetic **Unmatched** bucket rather than being dropped.
3. **Special case — "Custom Folder"**: when a file's subfolder is
   literally "Custom Folder", a second field (assumed internal name
   `Custom_x0020_Folder` — **change `CUSTOM_FOLDER_FIELD` at the top of
   the class if yours differs**) supplies a free-text value. Distinct
   values become child folders nested under "Custom Folder", and the
   file files under that child instead of directly under the root.
4. Every folder (root or custom child) gets a **file count**, shown in
   the LWC as `FolderName (N)` — counts roll up, so "Custom Folder"
   itself shows the total across all its custom children.

If your API's field names differ from the sample payload (`ID`,
`{FilenameWithExtension}`, `{Link}`, `Modified`, `Editor.DisplayName`,
`Contracts_x0020_Subfolder.Value`, `Custom_x0020_Folder.Value`), adjust
`parseFileRecord()` — it's the single place that maps raw JSON to the
internal `FileRecord` shape. There's no file-size field in the sample
payload, so size shows blank unless you tell `parseFileRecord()` where
to find it in your actual response.

## Sample data mode (works out of the box)

`SharePointTreeListController.cls` has a `USE_SAMPLE_DATA = true` flag at
the top. While `true`, `getTreeListData()` runs a small hardcoded list of
sample file records (spanning several of the real folder names, plus two
different "Custom Folder" sub-values and one intentionally-unmatched
record) through the exact same tree-building logic real data would go
through, instead of calling your REST API. `uploadFile()` just simulates
a successful upload without persisting anything. This means you can
deploy the component and see it fully working — tree, folder counts,
paging, filters, upload UX — before your integration is wired up.

**Flip `USE_SAMPLE_DATA` to `false`** once your Named Credential and REST
API are ready, so the class starts making real callouts.

## Setup

1. **Named Credential** — Create one called `SharePoint_API` (Setup >
   Named Credentials) pointing at your custom REST API's base URL, with
   whatever auth it needs (OAuth 2.0, API key header, etc.). The Apex
   class calls it as `callout:SharePoint_API/...` so no secrets or raw
   URLs live in code.

2. **Adjust endpoints/JSON shape** — Open
   `SharePointTreeListController.cls` and:
   - Set `GET_ENDPOINT` / `UPLOAD_ENDPOINT` to your API's actual paths.
   - Update `parseItem()` if your API's JSON field names differ from the
     assumed shape (`id`, `parentId`, `name`, `isFolder`, `size`,
     `modifiedBy`, `modifiedDate`, `url`). A commented-out alternative
     shows how to derive hierarchy from a flat `serverRelativeUrl` list
     instead, if your API doesn't return parent/child IDs directly.
   - Update the upload request body shape in `uploadFile()` to match
     what your API expects (it currently POSTs JSON with a base64
     `contentB64` field, a `fileName`, and a `folderId`).

3. **Deploy** with SFDX/Salesforce CLI:
   ```
   sf project deploy start -d force-app
   ```
   or drag the `classes` and `lwc` folders into a Metadata API deploy /
   VS Code project as usual.

4. **Add the component** to a Lightning App, Record, or Home page via
   App Builder (labelled "SharePoint Tree List"), or add it to a
   Lightning Tab.

## How it behaves

- **Loading**: on load, and after any upload, it fetches the *entire*
  library listing in one call (per your requirement that the API
  returns everything at once), showing a spinner overlay the whole
  time.
- **Tree**: rows are built from `id`/`parentId` into a parent → children
  map; folders sort before files, then alphabetically. Click the
  chevron to expand/collapse a folder.
- **Paging**: pages the **top-level (root)** rows — a real TreeList
  paginates root items, not every nested row, so folder contents aren't
  arbitrarily split across pages. Page size is adjustable (10/25/50/100).
- **Filtering**: each column has its own filter input (Name, Type,
  Size, Modified By, Modified Date), all combined with AND logic. When
  any filter is active, matching items' ancestor folders are
  auto-expanded so you can see the path to each match; paging then
  applies to the filtered root set.
- **Upload**: pick a target folder from the dropdown (or leave as
  Library Root), click **Upload File**, choose a file — it's read as
  base64 in the browser and sent to Apex, which POSTs it to your REST
  API. The tree refreshes automatically on success.

## Notes / things you may want to extend

- Sorting isn't implemented (wasn't requested) — the row-flattening
  function is a natural place to add a comparator if you want it later.
- Large libraries: since the API returns everything in one shot, very
  large libraries (tens of thousands of items) may make the initial
  load slow — that cost is inherent to "the REST API returns all data"
  and isn't something the component can avoid; if that becomes an
  issue, the natural fix is asking the API owner for server-side
  paging/filtering instead.
- File size limit: base64-in-JSON uploads work well for typical
  documents; for very large files (100MB+) you'd want a chunked/
  streaming upload endpoint on the API side instead.
