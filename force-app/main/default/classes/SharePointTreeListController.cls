/**
 * SharePointTreeListController
 * ------------------------------------------------------------
 * Calls a custom external REST API (fronting a SharePoint document
 * library) via a Named Credential, and exposes the data to the
 * sharePointTreeList LWC.
 *
 * IMPORTANT - THINGS YOU MUST CONFIGURE / ADJUST:
 * 1. Create a Named Credential called "SharePoint_API" pointing at your
 *    custom REST API's base URL (Setup > Named Credentials). Use it for
 *    auth (OAuth / API key / etc.) so no secrets live in this class.
 * 2. Adjust GET_ENDPOINT / UPLOAD_ENDPOINT to match your API's actual
 *    resource paths.
 * 3. Adjust parseItem() to match the actual JSON shape your API returns.
 *    The shape assumed below is a reasonable generic guess:
 *      [
 *        {
 *          "id": "unique-id",
 *          "parentId": "parent-unique-id-or-null-for-root",
 *          "name": "Folder or File name",
 *          "isFolder": true,
 *          "size": 12345,
 *          "modifiedBy": "Jane Doe",
 *          "modifiedDate": "2026-08-01T10:15:00Z",
 *          "url": "https://.../file.pdf"
 *        },
 *        ...
 *      ]
 *    If your API instead returns a flat list of server-relative URLs
 *    (no explicit parentId), derive parentId from the folder path in
 *    parseItem() - see the commented-out alternative there.
 *
 * Because the requirement states the external API returns ALL items in
 * one call (no native paging/filtering), this controller fetches the
 * full data set and hands it to the client once. Paging and per-column
 * filtering are then done client-side in the LWC for a fast, snappy
 * UI (no server round trip per keystroke/page).
 */
public with sharing class SharePointTreeListController {

    // ---- Adjust these to match your custom REST API ----
    private static final String GET_ENDPOINT    = 'callout:SharePoint_API/api/library/items';
    private static final String UPLOAD_ENDPOINT = 'callout:SharePoint_API/api/library/upload';

    // ---- Sample data mode ----
    // Flip this to false once your Named Credential + REST API are wired up.
    // While true, getTreeListData() returns the hardcoded sample below instead
    // of making a live callout, so you can deploy and demo the component
    // immediately. uploadFile() still just simulates success in this mode
    // (it does NOT persist anything) so you can test the upload UX end to end.
    private static final Boolean USE_SAMPLE_DATA = true;

    public class TreeItem {
        @AuraEnabled public String id;
        @AuraEnabled public String parentId;
        @AuraEnabled public String name;
        @AuraEnabled public Boolean isFolder;
        @AuraEnabled public Decimal size;
        @AuraEnabled public String modifiedBy;
        @AuraEnabled public String modifiedDate;
        @AuraEnabled public String fileUrl;
    }

    public class TreeListResult {
        @AuraEnabled public List<TreeItem> items;
        @AuraEnabled public Integer totalCount;
    }

    public class SharePointException extends Exception {}

    /**
     * Fetches the entire library listing from the external API.
     * Not marked cacheable because the caller (LWC) triggers this
     * explicitly (imperative call) so it can show a spinner and
     * control refresh-after-upload.
     */
    @AuraEnabled
    public static TreeListResult getTreeListData() {
        if (USE_SAMPLE_DATA) {
            return getSampleTreeListData();
        }

        TreeListResult result = new TreeListResult();
        result.items = new List<TreeItem>();

        HttpRequest req = new HttpRequest();
        req.setEndpoint(GET_ENDPOINT);
        req.setMethod('GET');
        req.setHeader('Accept', 'application/json');
        req.setTimeout(120000);

        Http http = new Http();
        HttpResponse res;
        try {
            res = http.send(req);
        } catch (Exception e) {
            throw new SharePointException('Callout to SharePoint API failed: ' + e.getMessage());
        }

        if (res.getStatusCode() < 200 || res.getStatusCode() >= 300) {
            throw new SharePointException(
                'SharePoint API returned status ' + res.getStatusCode() + ': ' + res.getBody()
            );
        }

        Object parsed = JSON.deserializeUntyped(res.getBody());
        List<Object> rawItems;

        // Support either a bare array response, or an object with an "items"/"value" wrapper
        if (parsed instanceof List<Object>) {
            rawItems = (List<Object>) parsed;
        } else if (parsed instanceof Map<String, Object>) {
            Map<String, Object> parsedMap = (Map<String, Object>) parsed;
            if (parsedMap.containsKey('items')) {
                rawItems = (List<Object>) parsedMap.get('items');
            } else if (parsedMap.containsKey('value')) {
                rawItems = (List<Object>) parsedMap.get('value');
            } else {
                throw new SharePointException('Unrecognized response shape from SharePoint API.');
            }
        } else {
            throw new SharePointException('Unrecognized response shape from SharePoint API.');
        }

        for (Object o : rawItems) {
            result.items.add(parseItem((Map<String, Object>) o));
        }
        result.totalCount = result.items.size();
        return result;
    }

    private static TreeItem parseItem(Map<String, Object> src) {
        TreeItem t = new TreeItem();
        t.id           = (String) src.get('id');
        t.parentId     = (String) src.get('parentId'); // null/blank => root item
        t.name         = (String) src.get('name');
        t.isFolder     = src.get('isFolder') == null ? false : (Boolean) src.get('isFolder');
        t.size         = src.get('size') == null ? null : Decimal.valueOf(String.valueOf(src.get('size')));
        t.modifiedBy   = (String) src.get('modifiedBy');
        t.modifiedDate = (String) src.get('modifiedDate');
        t.fileUrl      = (String) src.get('url');

        /* ---- Alternative parsing if your API only gives a flat server-relative path ----
         * String path = (String) src.get('serverRelativeUrl'); // e.g. /sites/x/Shared Documents/A/B/file.pdf
         * t.id = path;
         * List<String> parts = path.split('/');
         * parts.remove(parts.size() - 1);
         * t.parentId = String.join(parts, '/');
         */
        return t;
    }

    /**
     * Uploads a file (sent from the LWC as base64) into the SharePoint
     * library, under the given folder path/id, via the custom REST API.
     */
    @AuraEnabled
    public static String uploadFile(String base64Data, String fileName, String targetFolderId) {
        if (String.isBlank(base64Data) || String.isBlank(fileName)) {
            throw new SharePointException('File name and content are required.');
        }

        if (USE_SAMPLE_DATA) {
            // Demo mode: pretend the upload succeeded. Nothing is persisted.
            // Flip USE_SAMPLE_DATA to false to actually call the REST API below.
            return 'OK';
        }

        Map<String, Object> body = new Map<String, Object>{
            'fileName'   => fileName,
            'folderId'   => targetFolderId, // null/blank => upload to library root
            'contentB64' => base64Data
        };

        HttpRequest req = new HttpRequest();
        req.setEndpoint(UPLOAD_ENDPOINT);
        req.setMethod('POST');
        req.setHeader('Content-Type', 'application/json');
        req.setTimeout(120000);
        req.setBody(JSON.serialize(body));

        Http http = new Http();
        HttpResponse res;
        try {
            res = http.send(req);
        } catch (Exception e) {
            throw new SharePointException('Upload callout failed: ' + e.getMessage());
        }

        if (res.getStatusCode() < 200 || res.getStatusCode() >= 300) {
            throw new SharePointException(
                'Upload failed with status ' + res.getStatusCode() + ': ' + res.getBody()
            );
        }

        return 'OK';
    }

    /**
     * Hardcoded sample hierarchy that mimics what the external REST API
     * would return for a SharePoint document library. Handy for building
     * and demoing the LWC before the real integration is wired up.
     */
    private static TreeListResult getSampleTreeListData() {
        List<TreeItem> items = new List<TreeItem>();

        items.add(mkFolder('folder-contracts', null, 'Contracts', 'Jane Doe', '2026-07-01T09:15:00Z'));
        items.add(mkFolder('folder-contracts-2026', 'folder-contracts', '2026', 'Jane Doe', '2026-07-01T09:15:00Z'));
        items.add(mkFile('file-msa-acme', 'folder-contracts-2026', 'Acme_MSA_2026.pdf', 245000, 'Jane Doe', '2026-07-10T14:22:00Z'));
        items.add(mkFile('file-nda-globex', 'folder-contracts-2026', 'Globex_NDA.docx', 58000, 'Mike Chen', '2026-06-18T11:05:00Z'));
        items.add(mkFolder('folder-contracts-2025', 'folder-contracts', '2025', 'Jane Doe', '2025-12-20T08:00:00Z'));
        items.add(mkFile('file-msa-initech', 'folder-contracts-2025', 'Initech_MSA_2025.pdf', 198000, 'Priya Patel', '2025-11-02T16:40:00Z'));

        items.add(mkFolder('folder-marketing', null, 'Marketing', 'Sam Rivera', '2026-05-14T10:00:00Z'));
        items.add(mkFile('file-brand-guide', 'folder-marketing', 'Brand_Guidelines_v3.pdf', 3200000, 'Sam Rivera', '2026-05-14T10:02:00Z'));
        items.add(mkFolder('folder-marketing-assets', 'folder-marketing', 'Assets', 'Sam Rivera', '2026-05-14T10:03:00Z'));
        items.add(mkFile('file-logo-primary', 'folder-marketing-assets', 'logo-primary.png', 84000, 'Sam Rivera', '2026-05-14T10:04:00Z'));
        items.add(mkFile('file-logo-dark', 'folder-marketing-assets', 'logo-dark-mode.png', 91000, 'Sam Rivera', '2026-05-15T09:30:00Z'));

        items.add(mkFolder('folder-hr', null, 'HR', 'Angela Wu', '2026-01-10T08:30:00Z'));
        items.add(mkFile('file-handbook', 'folder-hr', 'Employee_Handbook_2026.pdf', 512000, 'Angela Wu', '2026-01-10T08:31:00Z'));
        items.add(mkFile('file-benefits', 'folder-hr', 'Benefits_Overview.pptx', 1450000, 'Angela Wu', '2026-02-01T13:10:00Z'));

        items.add(mkFile('file-readme-root', null, 'README.txt', 2100, 'System', '2026-01-01T00:00:00Z'));

        TreeListResult result = new TreeListResult();
        result.items = items;
        result.totalCount = items.size();
        return result;
    }

    private static TreeItem mkFolder(String id, String parentId, String name, String modifiedBy, String modifiedDate) {
        TreeItem t = new TreeItem();
        t.id = id;
        t.parentId = parentId;
        t.name = name;
        t.isFolder = true;
        t.size = null;
        t.modifiedBy = modifiedBy;
        t.modifiedDate = modifiedDate;
        t.fileUrl = null;
        return t;
    }

    private static TreeItem mkFile(String id, String parentId, String name, Integer sizeBytes, String modifiedBy, String modifiedDate) {
        TreeItem t = new TreeItem();
        t.id = id;
        t.parentId = parentId;
        t.name = name;
        t.isFolder = false;
        t.size = sizeBytes;
        t.modifiedBy = modifiedBy;
        t.modifiedDate = modifiedDate;
        t.fileUrl = 'https://contoso.sharepoint.com/sites/demo/Shared%20Documents/' + name;
        return t;
    }
}
