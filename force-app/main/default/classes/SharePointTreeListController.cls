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
}
