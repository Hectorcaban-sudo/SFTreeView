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

## Sample data mode (works out of the box)

`SharePointTreeListController.cls` has a `USE_SAMPLE_DATA = true` flag at
the top. While `true`, `getTreeListData()` returns a hardcoded sample
folder/file hierarchy (Contracts, Marketing, HR, etc.) instead of calling
your REST API, and `uploadFile()` just simulates a successful upload
without persisting anything. This means you can deploy the component and
see it fully working — tree, paging, filters, upload UX — before your
integration is wired up.

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
