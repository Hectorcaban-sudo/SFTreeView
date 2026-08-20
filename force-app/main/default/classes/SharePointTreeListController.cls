/**
 * SharePointTreeListController
 * ------------------------------------------------------------
 * Calls a custom external REST API (a Power Automate / Logic-Apps style
 * "get items" endpoint fronting a SharePoint document LIBRARY THAT HAS
 * NO REAL FOLDERS) and reshapes the flat file list into a virtual tree
 * for the sharePointTreeList LWC.
 *
 * HOW THE TREE IS BUILT (read this before changing anything)
 * ------------------------------------------------------------
 * The SharePoint library itself is flat - every record the REST API
 * returns is a FILE, never a folder ({IsFolder} / folder metadata in the
 * raw JSON is ignored on purpose). The folder structure shown in the
 * LWC is entirely virtual, derived like this:
 *
 * 1. FOLDER_LIST (below) is the fixed master list of top-level folders,
 *    taken from Contracts_Folders.csv. These always appear as root rows
 *    in the tree, even when a folder currently has 0 files.
 *
 * 2. Every file record has a "Contracts_x0020_Subfolder" field (a
 *    SharePoint lookup column) whose .Value is matched, case-insensitively,
 *    against FOLDER_LIST to decide which root folder the file belongs to.
 *    A file whose value doesn't match any known folder is bucketed under
 *    a synthetic "Unmatched" folder appended at runtime, so nothing is
 *    silently dropped.
 *
 * 3. SPECIAL CASE - "Custom Folder": when a file's Contracts_x0020_Subfolder
 *    is "Custom Folder", a SECOND field (assumed name below:
 *    CUSTOM_FOLDER_FIELD = "Custom_x0020_Folder", matching SharePoint's
 *    "%20" -> "_x0020_" internal-name convention) supplies a free-text
 *    sub-folder name. Distinct values of that field become CHILD folders
 *    under the "Custom Folder" root, and the file is filed under that
 *    child rather than directly under "Custom Folder".
 *    >>> If your API's actual internal field name for "Custom Folder"
 *    >>> differs, change CUSTOM_FOLDER_FIELD below - that's the only
 *    >>> place it's referenced.
 *
 * 4. Every folder (root AND custom sub-folder) gets a fileCount = number
 *    of files nested under it (recursively, so "Custom Folder" itself
 *    shows the sum of all its children). The LWC renders this as
 *    "FolderName (N)".
 *
 * SETUP
 * ------------------------------------------------------------
 * 1. Create a Named Credential called "SharePoint_API" pointing at your
 *    custom REST API's base URL (Setup > Named Credentials).
 * 2. Adjust GET_ENDPOINT / UPLOAD_ENDPOINT to your API's actual paths.
 * 3. Adjust parseFileRecord() if your API's JSON field names differ from
 *    the sample payload this was built against (ID, {FilenameWithExtension}
 *    / {Name}, {Link}, Modified, Editor.DisplayName / Author.DisplayName,
 *    Contracts_x0020_Subfolder.Value, Custom_x0020_Folder.Value).
 * 4. There is no file-size field in the sample payload provided, so size
 *    is left blank unless you tell parseFileRecord() where to find it.
 */
public with sharing class SharePointTreeListController {

    // ---- Adjust these to match your custom REST API ----
    private static final String GET_ENDPOINT    = 'callout:SharePoint_API/api/library/items';
    private static final String UPLOAD_ENDPOINT = 'callout:SharePoint_API/api/library/upload';

    // ---- Sample data mode ----
    // Flip this to false once your Named Credential + REST API are wired up.
    // While true, getTreeListData() builds the tree from the hardcoded
    // sample file records below instead of making a live callout, so you
    // can deploy and demo the component immediately. uploadFile() still
    // just simulates success in this mode (does NOT persist anything).
    private static final Boolean USE_SAMPLE_DATA = true;

    // Internal SharePoint field name that holds the free-text custom
    // sub-folder value. Only read when Contracts_x0020_Subfolder = "Custom Folder".
    // >>> CHANGE THIS if your API uses a different internal name. <<<
    private static final String CUSTOM_FOLDER_FIELD = 'Custom_x0020_Folder';

    private static final String CUSTOM_FOLDER_NAME = 'Custom Folder';
    private static final String UNMATCHED_FOLDER_NAME = 'Unmatched';

    // ---- Fixed master folder list, from Contracts_Folders.csv ----
    // name -> csv id (id is only used to build a stable root folder id)
    private static final Map<String, String> FOLDER_LIST = new Map<String, String>{
        'Base Contract' => '1',
        'Closeout' => '2',
        'Contract Document Checklist' => '3',
        'Correspondence' => '4',
        'CPARs' => '5',
        'Deliverables' => '6',
        'Government Property' => '7',
        'Modifications' => '8',
        'Negotiations' => '9',
        'OCI' => '10',
        'PCI' => '11',
        'Risk' => '12',
        'SB Plan' => '13',
        'NDA -TA' => '15',
        'ECPs' => '16',
        'Program Control' => '17',
        'COI' => '18',
        CUSTOM_FOLDER_NAME => '19',
        'Employee Qualification Packages' => '20',
        'TBD' => '22'
    };

    public class TreeItem {
        @AuraEnabled public String id;
        @AuraEnabled public String parentId;
        @AuraEnabled public String name;
        @AuraEnabled public Boolean isFolder;
        @AuraEnabled public Decimal size;
        @AuraEnabled public String modifiedBy;
        @AuraEnabled public String modifiedDate;
        @AuraEnabled public String fileUrl;
        @AuraEnabled public Integer fileCount; // only populated for folder rows
    }

    public class TreeListResult {
        @AuraEnabled public List<TreeItem> items;
        @AuraEnabled public Integer totalCount;
    }

    public class SharePointException extends Exception {}

    // Flat record shape parsed out of each raw JSON file entry, before
    // it gets placed into the virtual folder tree.
    private class FileRecord {
        String id;
        String name;
        String fileUrl;
        String modifiedBy;
        String modifiedDate;
        Decimal size;
        String contractSubfolder;
        String customFolder;
    }

    @AuraEnabled
    public static TreeListResult getTreeListData() {
        List<FileRecord> files = USE_SAMPLE_DATA ? getSampleFileRecords() : getLiveFileRecords();
        return buildTree(files);
    }

    private static List<FileRecord> getLiveFileRecords() {
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

        // Support either a bare array response, or an OData-style wrapper
        if (parsed instanceof List<Object>) {
            rawItems = (List<Object>) parsed;
        } else if (parsed instanceof Map<String, Object>) {
            Map<String, Object> parsedMap = (Map<String, Object>) parsed;
            if (parsedMap.containsKey('value')) {
                rawItems = (List<Object>) parsedMap.get('value');
            } else if (parsedMap.containsKey('items')) {
                rawItems = (List<Object>) parsedMap.get('items');
            } else {
                // single-record response, same shape as the sample payload
                rawItems = new List<Object>{ parsedMap };
            }
        } else {
            throw new SharePointException('Unrecognized response shape from SharePoint API.');
        }

        List<FileRecord> files = new List<FileRecord>();
        for (Object o : rawItems) {
            files.add(parseFileRecord((Map<String, Object>) o));
        }
        return files;
    }

    /**
     * Parses one raw JSON file record (Power Automate SharePoint
     * "get items" shape) into a FileRecord. Adjust field names here if
     * your API differs from the sample payload.
     */
    private static FileRecord parseFileRecord(Map<String, Object> src) {
        FileRecord f = new FileRecord();

        f.id = src.get('ID') != null ? String.valueOf(src.get('ID')) : (String) src.get('ItemInternalId');

        String filenameWithExt = (String) src.get('{FilenameWithExtension}');
        String displayName = (String) src.get('{Name}');
        f.name = String.isNotBlank(filenameWithExt) ? filenameWithExt : displayName;

        f.fileUrl = (String) src.get('{Link}');
        f.modifiedDate = (String) src.get('Modified');

        f.modifiedBy = extractLookupValue(src.get('Editor'), 'DisplayName');
        if (String.isBlank(f.modifiedBy)) {
            f.modifiedBy = extractLookupValue(src.get('Author'), 'DisplayName');
        }

        f.contractSubfolder = extractLookupValue(src.get('Contracts_x0020_Subfolder'), 'Value');
        f.customFolder = extractLookupValue(src.get(CUSTOM_FOLDER_FIELD), 'Value');

        // No file-size field observed in the sample payload. If your API
        // exposes one (e.g. "{Length}" or "File_x0020_Size"), map it here:
        // f.size = src.get('{Length}') == null ? null : Decimal.valueOf(String.valueOf(src.get('{Length}')));
        f.size = null;

        return f;
    }

    // Pulls a named sub-key out of a SharePoint "expanded reference"
    // object shape: { "@odata.type": "...", "Id": 2, "Value": "Closeout" }
    // Falls back to treating the value itself as a plain string.
    private static String extractLookupValue(Object raw, String key) {
        if (raw == null) {
            return null;
        }
        if (raw instanceof Map<String, Object>) {
            Object val = ((Map<String, Object>) raw).get(key);
            return val == null ? null : String.valueOf(val);
        }
        return String.valueOf(raw);
    }

    /**
     * Turns the flat file list into a virtual folder tree:
     * FOLDER_LIST roots -> (Custom Folder's children only) -> files.
     */
    private static TreeListResult buildTree(List<FileRecord> files) {
        List<TreeItem> items = new List<TreeItem>();

        // 1) root folders, in FOLDER_LIST order, id = 'folder-<csvId>'
        Map<String, String> folderNameToId = new Map<String, String>(); // lowercased name -> item id
        Map<String, TreeItem> foldersById = new Map<String, TreeItem>();

        for (String folderName : FOLDER_LIST.keySet()) {
            TreeItem folder = new TreeItem();
            folder.id = 'folder-' + FOLDER_LIST.get(folderName);
            folder.parentId = null;
            folder.name = folderName;
            folder.isFolder = true;
            folder.fileCount = 0;
            items.add(folder);
            foldersById.put(folder.id, folder);
            folderNameToId.put(folderName.trim().toLowerCase(), folder.id);
        }

        String customFolderRootId = folderNameToId.get(CUSTOM_FOLDER_NAME.toLowerCase());

        // 2) custom sub-folders, discovered from file data, id = 'customfolder-<slug>'
        Map<String, TreeItem> customSubfoldersByValue = new Map<String, TreeItem>(); // lowercased value -> item

        // 3) unmatched bucket, only created if actually needed
        TreeItem unmatchedFolder;

        for (FileRecord rec : files) {
            String matchedFolderId = String.isNotBlank(rec.contractSubfolder)
                ? folderNameToId.get(rec.contractSubfolder.trim().toLowerCase())
                : null;

            String targetParentId;

            if (matchedFolderId == null) {
                // no known folder matched - file into the synthetic Unmatched bucket
                if (unmatchedFolder == null) {
                    unmatchedFolder = new TreeItem();
                    unmatchedFolder.id = 'folder-unmatched';
                    unmatchedFolder.parentId = null;
                    unmatchedFolder.name = UNMATCHED_FOLDER_NAME;
                    unmatchedFolder.isFolder = true;
                    unmatchedFolder.fileCount = 0;
                    items.add(unmatchedFolder);
                }
                unmatchedFolder.fileCount += 1;
                targetParentId = unmatchedFolder.id;

            } else if (matchedFolderId == customFolderRootId && String.isNotBlank(rec.customFolder)) {
                // Custom Folder with a specific sub-folder value: create/find the child
                String key = rec.customFolder.trim().toLowerCase();
                TreeItem sub = customSubfoldersByValue.get(key);
                if (sub == null) {
                    sub = new TreeItem();
                    sub.id = 'customfolder-' + key.replaceAll('[^a-z0-9]+', '-');
                    sub.parentId = customFolderRootId;
                    sub.name = rec.customFolder.trim();
                    sub.isFolder = true;
                    sub.fileCount = 0;
                    items.add(sub);
                    customSubfoldersByValue.put(key, sub);
                }
                sub.fileCount += 1;
                foldersById.get(customFolderRootId).fileCount += 1; // parent total includes children
                targetParentId = sub.id;

            } else {
                // normal match, straight into that root folder
                foldersById.get(matchedFolderId).fileCount += 1;
                targetParentId = matchedFolderId;
            }

            TreeItem fileItem = new TreeItem();
            fileItem.id = 'file-' + rec.id;
            fileItem.parentId = targetParentId;
            fileItem.name = rec.name;
            fileItem.isFolder = false;
            fileItem.size = rec.size;
            fileItem.modifiedBy = rec.modifiedBy;
            fileItem.modifiedDate = rec.modifiedDate;
            fileItem.fileUrl = rec.fileUrl;
            items.add(fileItem);
        }

        TreeListResult result = new TreeListResult();
        result.items = items;
        result.totalCount = items.size();
        return result;
    }

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
     * Hardcoded sample file records (same field shape as the real API)
     * spanning several folders - including a few different "Custom
     * Folder" sub-values - so the demo tree shows nesting and counts.
     */
    private static List<FileRecord> getSampleFileRecords() {
        List<FileRecord> files = new List<FileRecord>();

        files.add(mkFile('93339', 'Project 0996-02-0008 Closeout Bulletin.pdf', 'Closeout', null,
            'Swygert, Carol C (US)', '2021-06-01T21:09:43Z'));
        files.add(mkFile('93340', 'Closeout Checklist Final.docx', 'Closeout', null,
            'Swygert, Carol C (US)', '2021-06-05T14:12:00Z'));

        files.add(mkFile('93341', 'Base_Contract_Signed.pdf', 'Base Contract', null,
            'Jane Doe', '2026-07-01T09:15:00Z'));

        files.add(mkFile('93342', 'CPAR_2025_Q4.pdf', 'CPARs', null,
            'Mike Chen', '2026-01-15T10:00:00Z'));

        files.add(mkFile('93343', 'Modification_P00012.pdf', 'Modifications', null,
            'Priya Patel', '2025-11-02T16:40:00Z'));
        files.add(mkFile('93344', 'Modification_P00013.pdf', 'Modifications', null,
            'Priya Patel', '2025-12-11T08:05:00Z'));
        files.add(mkFile('93345', 'Modification_P00014.pdf', 'Modifications', null,
            'Priya Patel', '2026-02-20T13:45:00Z'));

        // Custom Folder examples - two distinct custom sub-folder values
        files.add(mkFile('93346', 'Site_Visit_Photos.zip', CUSTOM_FOLDER_NAME, 'Site Visits',
            'Sam Rivera', '2026-03-01T09:00:00Z'));
        files.add(mkFile('93347', 'Site_Visit_Notes.docx', CUSTOM_FOLDER_NAME, 'Site Visits',
            'Sam Rivera', '2026-03-02T11:20:00Z'));
        files.add(mkFile('93348', 'Vendor_Correspondence_Misc.pdf', CUSTOM_FOLDER_NAME, 'Vendor Misc',
            'Angela Wu', '2026-04-18T15:30:00Z'));

        // A record whose subfolder value doesn't match anything in FOLDER_LIST,
        // to demonstrate the "Unmatched" safety-net bucket
        files.add(mkFile('93349', 'Old_Legacy_Doc.pdf', 'Archive (legacy)', null,
            'System', '2020-05-01T00:00:00Z'));

        return files;
    }

    private static FileRecord mkFile(String id, String name, String contractSubfolder, String customFolder,
                                      String modifiedBy, String modifiedDate) {
        FileRecord f = new FileRecord();
        f.id = id;
        f.name = name;
        f.fileUrl = 'https://contoso.sharepoint.com/sites/demo/Contracts%20Library/' + name;
        f.modifiedBy = modifiedBy;
        f.modifiedDate = modifiedDate;
        f.contractSubfolder = contractSubfolder;
        f.customFolder = customFolder;
        f.size = null;
        return f;
    }
}
