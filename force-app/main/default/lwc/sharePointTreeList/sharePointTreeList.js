import { LightningElement, track, api, wire } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import { getRecord, getFieldValue } from 'lightning/uiRecordApi';
import getTreeListData from '@salesforce/apex/SharePointTreeListController.getTreeListData';
import uploadFile from '@salesforce/apex/SharePointTreeListController.uploadFile';

const PAGE_SIZE_OPTIONS = [10, 25, 50, 100];

export default class SharePointTreeList extends LightningElement {
    // ---- record context (auto-populated when placed on a Record Page) ----
    @api recordId;
    @api objectApiName;

    // ---- configurable in App Builder: field API names on the current
    // record that hold the project number / task order number. Change
    // the defaults below to match your object, or set them per-page in
    // the Lightning App Builder property panel. ----
    @api projectNumberFieldName = 'Project_Number__c';
    @api taskOrderNumberFieldName = 'Task_Order_Number__c';

    // resolved values, sent to Apex on every call
    projectNumber;
    taskOrderNumber;
    recordFieldsError;

    // ---- raw data ----
    allItems = [];                 // flat list from Apex
    childrenByParent = new Map();  // parentId ('' for root) -> [items]

    // ---- ui state ----
    @track isLoading = true;
    @track errorMessage;
    @track expandedIds = new Set();
    @track filters = {
        name: '',
        type: '',
        modifiedBy: '',
        modifiedDate: '',
        size: ''
    };

    @track currentPage = 1;
    pageSize = 25;
    pageSizeOptions = PAGE_SIZE_OPTIONS.map((n) => ({ label: String(n), value: String(n) }));

    @track visibleRows = [];       // rows currently rendered (after paging + filtering)
    totalRootCount = 0;

    // upload
    @track isUploading = false;
    selectedFolderId = '';         // '' = root of library
    folderOptions = [{ label: '(Library Root)', value: '' }];

    // Builds the qualified field references getRecord needs, e.g.
    // "Opportunity.Project_Number__c". Empty until objectApiName is known
    // (it's auto-injected by the platform when on a Record Page).
    get qualifiedRecordFields() {
        if (!this.objectApiName || !this.recordId) {
            return [];
        }
        return [
            `${this.objectApiName}.${this.projectNumberFieldName}`,
            `${this.objectApiName}.${this.taskOrderNumberFieldName}`
        ];
    }

    @wire(getRecord, { recordId: '$recordId', fields: '$qualifiedRecordFields' })
    wiredRecord({ data, error }) {
        if (data) {
            this.projectNumber = getFieldValue(data, `${this.objectApiName}.${this.projectNumberFieldName}`);
            this.taskOrderNumber = getFieldValue(data, `${this.objectApiName}.${this.taskOrderNumberFieldName}`);
            this.recordFieldsError = undefined;
            this.loadData();
        } else if (error) {
            this.recordFieldsError =
                `Could not read "${this.projectNumberFieldName}" / "${this.taskOrderNumberFieldName}" from this record. ` +
                `Check the field API names configured for this component. (${this.extractError(error)})`;
            this.isLoading = false;
        }
    }

    connectedCallback() {
        // If there's no record context (component placed on a Home/App
        // page rather than a Record Page), qualifiedRecordFields will be
        // empty and the wire above never fires - so load without project/
        // task order values in that case.
        if (!this.recordId) {
            this.loadData();
        }
    }

    get recordContextLabel() {
        if (!this.recordId) {
            return '';
        }
        const parts = [];
        if (this.projectNumber) parts.push(`Project #: ${this.projectNumber}`);
        if (this.taskOrderNumber) parts.push(`Task Order #: ${this.taskOrderNumber}`);
        return parts.join('   |   ');
    }

    // ------------------------------------------------------------------
    // Data loading
    // ------------------------------------------------------------------
    loadData() {
        this.isLoading = true;
        this.errorMessage = undefined;

        getTreeListData({
            projectNumber: this.projectNumber,
            taskOrderNumber: this.taskOrderNumber
        })
            .then((result) => {
                this.allItems = result.items || [];
                this.buildIndexes();
                this.buildFolderOptions();
                this.currentPage = 1;
                this.recomputeVisibleRows();
            })
            .catch((error) => {
                this.errorMessage = this.extractError(error);
                this.allItems = [];
                this.visibleRows = [];
            })
            .finally(() => {
                this.isLoading = false;
            });
    }

    handleRefresh() {
        this.loadData();
    }

    buildIndexes() {
        this.childrenByParent = new Map();
        for (const item of this.allItems) {
            const key = item.parentId ? item.parentId : '';
            if (!this.childrenByParent.has(key)) {
                this.childrenByParent.set(key, []);
            }
            this.childrenByParent.get(key).push(item);
        }
        // sort each bucket: folders first, then alpha by name
        for (const list of this.childrenByParent.values()) {
            list.sort((a, b) => {
                if (!!a.isFolder !== !!b.isFolder) {
                    return a.isFolder ? -1 : 1;
                }
                return (a.name || '').localeCompare(b.name || '');
            });
        }
        this.computeFolderFileCounts();
    }

    // Counts files under every folder directly from childrenByParent, so
    // the number shown is always consistent with the tree actually being
    // rendered - independent of the fileCount value Apex sends.
    computeFolderFileCounts() {
        this.folderFileCounts = new Map();
        const countFiles = (folderId) => {
            const kids = this.childrenByParent.get(folderId) || [];
            let total = 0;
            for (const kid of kids) {
                total += kid.isFolder ? countFiles(kid.id) : 1;
            }
            this.folderFileCounts.set(folderId, total);
            return total;
        };
        for (const item of this.allItems) {
            if (item.isFolder && !this.folderFileCounts.has(item.id)) {
                countFiles(item.id);
            }
        }
    }

    buildFolderOptions() {
        const opts = [{ label: '(Library Root)', value: '' }];
        for (const item of this.allItems) {
            if (item.isFolder) {
                opts.push({ label: item.name, value: item.id });
            }
        }
        this.folderOptions = opts;
    }

    hasChildren(itemId) {
        const key = itemId ? itemId : '';
        return this.childrenByParent.has(key) && this.childrenByParent.get(key).length > 0;
    }

    // ------------------------------------------------------------------
    // Filtering + tree flattening + paging
    // ------------------------------------------------------------------
    get isFilterActive() {
        return Object.values(this.filters).some((v) => v && v.trim().length > 0);
    }

    itemMatchesFilters(item) {
        const f = this.filters;
        if (f.name && !(item.name || '').toLowerCase().includes(f.name.toLowerCase())) {
            return false;
        }
        if (f.type) {
            const typeLabel = item.isFolder ? 'folder' : 'file';
            if (!typeLabel.includes(f.type.toLowerCase())) {
                return false;
            }
        }
        if (f.modifiedBy && !(item.modifiedBy || '').toLowerCase().includes(f.modifiedBy.toLowerCase())) {
            return false;
        }
        if (f.modifiedDate && !(item.modifiedDate || '').toLowerCase().includes(f.modifiedDate.toLowerCase())) {
            return false;
        }
        if (f.size && !String(item.size != null ? item.size : '').includes(f.size)) {
            return false;
        }
        return true;
    }

    // returns Set of item ids that either match the filter themselves,
    // or are an ancestor of a matching item (so the path stays visible)
    computeFilterVisibleIds() {
        const visible = new Set();
        const byId = new Map(this.allItems.map((i) => [i.id, i]));

        for (const item of this.allItems) {
            if (this.itemMatchesFilters(item)) {
                visible.add(item.id);
                // walk up ancestors
                let parentId = item.parentId;
                while (parentId && byId.has(parentId) && !visible.has(parentId)) {
                    visible.add(parentId);
                    parentId = byId.get(parentId).parentId;
                }
            }
        }
        return visible;
    }

    recomputeVisibleRows() {
        const filterMode = this.isFilterActive;
        const filterVisibleIds = filterMode ? this.computeFilterVisibleIds() : null;

        // root rows, filtered if needed
        let rootRows = this.childrenByParent.get('') || [];
        if (filterMode) {
            rootRows = rootRows.filter((r) => filterVisibleIds.has(r.id));
        }

        this.totalRootCount = rootRows.length;

        // page the root rows
        const start = (this.currentPage - 1) * this.pageSize;
        const pagedRoots = rootRows.slice(start, start + this.pageSize);

        // flatten each paged root with its visible descendants
        const rows = [];
        for (const root of pagedRoots) {
            this.flattenInto(root, 0, filterMode, filterVisibleIds, rows);
        }
        this.visibleRows = rows;
    }

    flattenInto(item, level, filterMode, filterVisibleIds, out) {
        const expanded = filterMode ? true : this.expandedIds.has(item.id);
        const childList = this.childrenByParent.get(item.id) || [];
        const hasKids = childList.length > 0;
        const fileCount = item.isFolder ? this.folderFileCounts.get(item.id) || 0 : null;

        out.push({
            id: item.id,
            name: item.name,
            countLabel: item.isFolder ? `(${fileCount})` : '',
            isFolder: item.isFolder,
            typeLabel: item.isFolder ? 'Folder' : 'File',
            size: item.isFolder ? '' : this.formatSize(item.size),
            modifiedBy: item.modifiedBy,
            modifiedDate: this.formatDate(item.modifiedDate),
            fileUrl: item.fileUrl,
            hasChildren: hasKids,
            isExpanded: expanded,
            level,
            indentStyle: `padding-left: ${level * 24}px`,
            toggleIcon: expanded ? 'utility:chevrondown' : 'utility:chevronright',
            rowClass: item.isFolder ? 'row-folder' : 'row-file'
        });

        if (hasKids && expanded) {
            let kids = childList;
            if (filterMode) {
                kids = kids.filter((k) => filterVisibleIds.has(k.id));
            }
            for (const kid of kids) {
                this.flattenInto(kid, level + 1, filterMode, filterVisibleIds, out);
            }
        }
    }

    formatSize(bytes) {
        if (bytes === null || bytes === undefined || bytes === '') {
            return '';
        }
        const num = Number(bytes);
        if (Number.isNaN(num)) {
            return '';
        }
        if (num < 1024) return `${num} B`;
        if (num < 1024 * 1024) return `${(num / 1024).toFixed(1)} KB`;
        return `${(num / (1024 * 1024)).toFixed(1)} MB`;
    }

    formatDate(dateStr) {
        if (!dateStr) return '';
        const d = new Date(dateStr);
        if (Number.isNaN(d.getTime())) return dateStr;
        return d.toLocaleString();
    }

    // ------------------------------------------------------------------
    // Expand / collapse
    // ------------------------------------------------------------------
    handleToggleRow(event) {
        const id = event.currentTarget.dataset.id;
        if (this.expandedIds.has(id)) {
            this.expandedIds.delete(id);
        } else {
            this.expandedIds.add(id);
        }
        this.recomputeVisibleRows();
    }

    // ------------------------------------------------------------------
    // Filters (per column)
    // ------------------------------------------------------------------
    handleFilterChange(event) {
        const field = event.currentTarget.dataset.field;
        this.filters = { ...this.filters, [field]: event.target.value };
        this.currentPage = 1;
        this.recomputeVisibleRows();
    }

    handleClearFilters() {
        this.filters = { name: '', type: '', modifiedBy: '', modifiedDate: '', size: '' };
        this.currentPage = 1;
        this.recomputeVisibleRows();
    }

    // ------------------------------------------------------------------
    // Paging
    // ------------------------------------------------------------------
    get totalPages() {
        return Math.max(1, Math.ceil(this.totalRootCount / this.pageSize));
    }

    get isFirstPage() {
        return this.currentPage <= 1;
    }

    get isLastPage() {
        return this.currentPage >= this.totalPages;
    }

    get pageInfoLabel() {
        if (this.totalRootCount === 0) {
            return 'No records';
        }
        return `Page ${this.currentPage} of ${this.totalPages} (${this.totalRootCount} top-level items)`;
    }

    handlePrevPage() {
        if (!this.isFirstPage) {
            this.currentPage -= 1;
            this.recomputeVisibleRows();
        }
    }

    handleNextPage() {
        if (!this.isLastPage) {
            this.currentPage += 1;
            this.recomputeVisibleRows();
        }
    }

    handlePageSizeChange(event) {
        this.pageSize = Number(event.detail.value);
        this.currentPage = 1;
        this.recomputeVisibleRows();
    }

    // ------------------------------------------------------------------
    // Upload
    // ------------------------------------------------------------------
    handleFolderSelect(event) {
        this.selectedFolderId = event.detail.value;
    }

    handleFileSelected(event) {
        const file = event.target.files && event.target.files[0];
        if (!file) {
            return;
        }
        this.isUploading = true;
        const reader = new FileReader();
        reader.onload = () => {
            // result looks like: data:<mime>;base64,AAAA...
            const base64 = reader.result.split(',')[1];
            uploadFile({
                base64Data: base64,
                fileName: file.name,
                targetFolderId: this.selectedFolderId,
                projectNumber: this.projectNumber,
                taskOrderNumber: this.taskOrderNumber
            })
                .then(() => {
                    this.dispatchEvent(
                        new ShowToastEvent({
                            title: 'Success',
                            message: `${file.name} uploaded.`,
                            variant: 'success'
                        })
                    );
                    // reset the file input and refresh the tree
                    const fileInput = this.template.querySelector('.file-input');
                    if (fileInput) fileInput.value = '';
                    this.loadData();
                })
                .catch((error) => {
                    this.dispatchEvent(
                        new ShowToastEvent({
                            title: 'Upload failed',
                            message: this.extractError(error),
                            variant: 'error'
                        })
                    );
                })
                .finally(() => {
                    this.isUploading = false;
                });
        };
        reader.onerror = () => {
            this.isUploading = false;
            this.dispatchEvent(
                new ShowToastEvent({
                    title: 'Upload failed',
                    message: 'Could not read the selected file.',
                    variant: 'error'
                })
            );
        };
        reader.readAsDataURL(file);
    }

    handleUploadButtonClick() {
        const fileInput = this.template.querySelector('.file-input');
        if (fileInput) fileInput.click();
    }

    // ------------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------------
    extractError(error) {
        if (error && error.body && error.body.message) {
            return error.body.message;
        }
        if (error && error.message) {
            return error.message;
        }
        return 'An unknown error occurred.';
    }
}
