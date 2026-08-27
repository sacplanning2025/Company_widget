(function () {
    "use strict";

    var tmpl = document.createElement("template");
    tmpl.innerHTML = `
        <style>
            :host{
                display:block;
                width:100%;
                height:100%;
                font-family:"72", Arial, Helvetica, sans-serif;
                color:#1f2d3d;
                --cw-primary:#0a6ed1;
                --cw-primary-dark:#085caf;
                --cw-border:#d9d9d9;
                --cw-soft:#f7f9fb;
                --cw-soft2:#fafcff;
                --cw-text:#1f2d3d;
                --cw-sub:#6a6d70;
                --cw-success-bg:#f6ffed;
                --cw-success-bd:#b7eb8f;
                --cw-success-tx:#237804;
                --cw-error-bg:#fff1f0;
                --cw-error-bd:#ffa39e;
                --cw-error-tx:#a8071a;
                --cw-warn-bg:#fffbe6;
                --cw-warn-bd:#ffe58f;
                --cw-warn-tx:#ad6800;
            }
            *{ box-sizing:border-box; }
            .excel-widget{ width:100%; height:100%; }
            .upload-card{
                border:1px solid var(--cw-border);
                border-radius:12px;
                background:linear-gradient(180deg, #ffffff 0%, #fbfdff 100%);
                box-shadow:0 6px 18px rgba(0,0,0,0.06);
                padding:10px;
                height:100%;
                display:flex;
                flex-direction:column;
                overflow:hidden;
                gap:8px;
            }
            .top-bar{
                display:flex;
                justify-content:space-between;
                align-items:center;
                gap:10px;
                flex-wrap:wrap;
            }
            .top-bar-right{
                display:flex;
                align-items:center;
                gap:8px;
                flex-wrap:wrap;
            }
            .title-wrap{ display:flex; flex-direction:column; gap:2px; }
            .upload-title{
                font-size:15px;
                font-weight:700;
                color:var(--cw-primary);
                line-height:1.15;
            }
            .upload-subtitle{
                font-size:11px;
                color:var(--cw-sub);
                line-height:1.3;
            }
            .status-badge{
                font-size:11px;
                font-weight:700;
                padding:6px 10px;
                border-radius:9px;
                background:#f5f6f7;
                color:#354a5f;
                border:1px solid var(--cw-border);
                white-space:nowrap;
                display:inline-flex;
                align-items:center;
                min-height:34px;
            }
            .status-ready{ background:#f5f6f7; color:#354a5f; border-color:var(--cw-border); }
            .status-processing{ background:#fff7e6; color:#8a5a00; border-color:#ffd591; }
            .status-completed{ background:var(--cw-success-bg); color:var(--cw-success-tx); border-color:var(--cw-success-bd); }
            .status-error{ background:var(--cw-error-bg); color:var(--cw-error-tx); border-color:var(--cw-error-bd); }
            .status-warning{ background:var(--cw-warn-bg); color:var(--cw-warn-tx); border-color:var(--cw-warn-bd); }

            .toolbar-btn{
                appearance:none;
                border:1px solid #c7ced4;
                background:#ffffff;
                color:var(--cw-text);
                border-radius:9px;
                padding:8px 14px;
                font-size:12px;
                font-weight:700;
                cursor:pointer;
                transition:all 0.2s ease;
                min-height:34px;
            }
            .toolbar-btn:hover{
                border-color:var(--cw-primary);
                color:var(--cw-primary);
                background:#f4f9ff;
            }
            .toolbar-btn.primary{
                background:var(--cw-primary);
                color:#ffffff;
                border-color:var(--cw-primary);
            }
            .toolbar-btn.primary:hover{
                background:var(--cw-primary-dark);
                color:#ffffff;
                border-color:var(--cw-primary-dark);
            }
            .toolbar-btn:disabled{
                opacity:0.55;
                cursor:not-allowed;
            }

            .upload-area{
                border:1px solid #e3eaf2;
                border-radius:12px;
                padding:12px;
                background:#fafcff;
                flex:0 0 auto;
            }
            .upload-top-row{
                display:flex;
                justify-content:space-between;
                align-items:center;
                gap:14px;
                flex-wrap:wrap;
            }
            .upload-left{
                display:flex;
                align-items:center;
                gap:10px;
                flex-wrap:nowrap;
                min-width:auto;
            }
            .upload-right{
                display:flex;
                align-items:center;
                justify-content:flex-end;
                flex:0 0 320px;
                max-width:100%;
            }
            .file-input{
                font-size:13px;
                padding:10px 12px;
                border:1px solid var(--cw-border);
                border-radius:10px;
                background:#fff;
                width:300px;
                min-width:300px;
                max-width:300px;
                height:42px;
                flex:0 0 300px;
                color:var(--cw-text);
            }
            .action-btn{
                min-width:100px;
                height:42px;
                padding:0 14px;
                font-size:13px;
                border-radius:10px;
            }
            .continue-btn{
                min-width:160px;
            }
            .search-input{
                width:100%;
                min-width:260px;
                padding:10px 12px;
                border:1px solid var(--cw-border);
                border-radius:10px;
                font-size:13px;
                outline:none;
                background:#fff;
                height:42px;
                color:var(--cw-text);
            }
            .search-input:focus{
                border-color:var(--cw-primary);
                box-shadow:0 0 0 2px rgba(10,110,209,0.12);
            }
            .mini-text{
                font-size:11px;
                color:var(--cw-sub);
                margin-top:8px;
                line-height:1.4;
            }
            .mini-text b{ color:var(--cw-text); }

            .summary-grid{
                display:none;
                grid-template-columns:repeat(6, minmax(0, 1fr));
                gap:8px;
                margin-top:10px;
            }
            .summary-grid.show{ display:grid; }
            .summary-item{
                border:1px solid #edf1f5;
                border-radius:10px;
                padding:8px 10px;
                background:#ffffff;
                min-height:52px;
            }
            .summary-item .k{
                font-size:11px;
                color:var(--cw-sub);
                margin-bottom:3px;
                font-weight:600;
            }
            .summary-item .v{
                font-size:15px;
                font-weight:700;
                color:var(--cw-text);
            }

            .progress-wrap{ display:none; }
            .progress-wrap.show{ display:block; }
            .progress-label-row{
                display:flex;
                justify-content:space-between;
                align-items:center;
                margin-bottom:6px;
                font-size:12px;
                color:#354a5f;
            }
            .progress-bar{
                width:100%;
                height:11px;
                background:#edf2f7;
                border-radius:10px;
                overflow:hidden;
                border:1px solid var(--cw-border);
            }
            .progress-fill{
                width:0%;
                height:100%;
                background:linear-gradient(90deg, var(--cw-primary), #4db1ff);
                transition:width 0.25s ease;
            }

            .preview-wrap{
                display:none;
                border:1px solid #dfe6ee;
                border-radius:12px;
                overflow:hidden;
                background:#ffffff;
                flex:1 1 auto;
                min-height:0;
            }
            .preview-wrap.show{
                display:flex;
                flex-direction:column;
            }
            .section-head{
                display:flex;
                justify-content:space-between;
                align-items:center;
                padding:8px 10px;
                background:var(--cw-soft);
                border-bottom:1px solid #e5e7eb;
                gap:8px;
                flex-wrap:wrap;
            }
            .section-title{
                font-size:12px;
                font-weight:700;
                color:#354a5f;
            }
            .section-meta{
                font-size:11px;
                color:var(--cw-sub);
            }
            .preview-grid{
                overflow:auto;
                flex:1 1 auto;
                min-height:0;
                height:100%;
            }
            table{
                width:max-content;
                min-width:100%;
                border-collapse:collapse;
            }
            th, td{
                border-bottom:1px solid #eef2f6;
                padding:7px 10px;
                text-align:left;
                vertical-align:top;
                font-size:12px;
                white-space:nowrap;
                color:var(--cw-text);
            }
            th{
                position:sticky;
                top:0;
                z-index:1;
                background:#f8fbff;
                color:#354a5f;
                font-weight:700;
            }
            tr:hover td{ background:#fafcff; }
            tr.invalid-row td{ background:#fff7f7; }
            td.invalid-cell{
                background:#fff1f0 !important;
                border-left:3px solid #ff4d4f;
            }
            .cell-error{
                display:block;
                margin-top:4px;
                color:#cf1322;
                font-size:11px;
                line-height:1.3;
                white-space:normal;
            }
            .row-status-valid{
                color:#237804;
                font-weight:700;
            }
            .row-status-invalid{
                color:#cf1322;
                font-weight:700;
            }

            .bottom-actions{
                display:flex;
                justify-content:space-between;
                align-items:center;
                gap:12px;
                flex-wrap:wrap;
                padding:10px 12px;
                border:1px solid #dfe6ee;
                border-radius:12px;
                background:#f9fbfd;
            }
            .bottom-summary{
                font-size:12px;
                color:#354a5f;
                font-weight:600;
            }
            .bottom-btns{
                display:flex;
                gap:8px;
                flex-wrap:wrap;
            }

            .log-box{
                display:none;
                border:1px solid #e5e7eb;
                border-radius:10px;
                background:#fafbfc;
                padding:8px 10px;
                max-height:100px;
                overflow:auto;
                font-size:11px;
                line-height:1.4;
                color:#354a5f;
                white-space:pre-wrap;
            }
            .log-box.show{ display:block; }
            .footer-note{
                font-size:11px;
                color:var(--cw-sub);
                white-space:pre-wrap;
            }

            @media (max-width: 980px){
                .upload-top-row{ flex-direction:column; align-items:stretch; }
                .upload-left{ min-width:100%; flex-wrap:wrap; }
                .upload-right{ flex:1 1 auto; width:100%; }
                .file-input{
                    min-width:100%;
                    max-width:100%;
                    width:100%;
                    flex:1 1 auto;
                }
                .search-input{ min-width:100%; }
                .summary-grid{ grid-template-columns:repeat(2, minmax(0, 1fr)); }
            }
        </style>

        <div class="excel-widget">
            <div class="upload-card">
                <div class="top-bar">
                    <div class="title-wrap">
                        <div class="upload-title" id="titleEl">Excel Upload</div>
                        <div class="upload-subtitle" id="subtitleEl">Preview, validate and upload Excel file</div>
                    </div>
                    <div class="top-bar-right">
                        <button type="button" class="toolbar-btn primary" id="downloadTemplateBtn">Download Template</button>
                        <div class="status-badge status-ready" id="statusBadge">Ready</div>
                        <button type="button" class="toolbar-btn" id="downloadErrorBtn" disabled>Download Error Log</button>
                        <button type="button" class="toolbar-btn" id="togglePreviewBtn">Toggle Preview</button>
                    </div>
                </div>

                <div class="upload-area" id="dropZone">
                    <div class="upload-top-row">
                        <div class="upload-left">
                            <input type="file" id="fileInput" class="file-input" />
                            <button type="button" class="toolbar-btn primary action-btn" id="previewBtn">Preview</button>
                            <button type="button" class="toolbar-btn primary action-btn continue-btn" id="continueBtn" disabled>Continue to Upload</button>
                            <button type="button" class="toolbar-btn action-btn" id="clearBtn">Clear</button>
                        </div>
                        <div class="upload-right">
                            <input type="text" id="searchInput" class="search-input" placeholder="Search in preview..." />
                        </div>
                    </div>

                    <div class="mini-text" id="templateInfo">Required columns will be validated automatically.</div>

                    <div class="summary-grid" id="summaryGrid">
                        <div class="summary-item"><div class="k">Rows Read</div><div class="v" id="sumRows">0</div></div>
                        <div class="summary-item"><div class="k">Valid Rows</div><div class="v" id="sumValid">0</div></div>
                        <div class="summary-item"><div class="k">Invalid Rows</div><div class="v" id="sumInvalid">0</div></div>
                        <div class="summary-item"><div class="k">Sheet</div><div class="v" id="sumSheet">-</div></div>
                        <div class="summary-item"><div class="k">Columns</div><div class="v" id="sumColumns">0</div></div>
                        <div class="summary-item"><div class="k">Validation</div><div class="v" id="sumValidation">-</div></div>
                    </div>
                </div>

                <div class="progress-wrap" id="progressWrap">
                    <div class="progress-label-row">
                        <span id="progressText">Preparing preview...</span>
                        <span id="progressPercent">0%</span>
                    </div>
                    <div class="progress-bar">
                        <div class="progress-fill" id="progressFill"></div>
                    </div>
                </div>

                <div class="preview-wrap" id="previewWrap">
                    <div class="section-head">
                        <div class="section-title">Preview</div>
                        <div class="section-meta" id="previewMeta">0 row(s) shown</div>
                    </div>
                    <div class="preview-grid">
                        <table>
                            <thead id="previewHead"></thead>
                            <tbody id="previewBody"></tbody>
                        </table>
                    </div>
                </div>

                <div class="bottom-actions" id="bottomActions">
                    <div class="bottom-summary" id="bottomSummary">Preview the file first. Validation will be shown before upload.</div>
                    <div class="bottom-btns">
                        <button type="button" class="toolbar-btn primary" id="continueBottomBtn" disabled>Continue to Upload</button>
                        <button type="button" class="toolbar-btn" id="clearBottomBtn">Clear</button>
                    </div>
                </div>

                <div class="log-box" id="logBox"></div>
                <div class="footer-note" id="footerNote">Supported template: Sheet1 with required business columns</div>
            </div>
        </div>
    `;

    var TEMPLATE_URL = "https://raw.githubusercontent.com/sacplanning2025/Company_widget/main/New_Position_Creation_V2.1.xlsm";

    class Excel extends HTMLElement {
        constructor() {
            super();
            this._shadowRoot = this.attachShadow({ mode: "open" });
            this._shadowRoot.appendChild(tmpl.content.cloneNode(true));

            this._export_settings = {
                title: "",
                subtitle: "",
                icon: "",
                unit: "",
                validpayload: "",
                footer: "",
                errorlogfilename: "Excel_Upload_Error_Log.csv",
                templatefilename: "New_Position_Creation_V2.1.xlsm",
                templateurl: TEMPLATE_URL,
                requiredcolumns: "ID,DESCRIPTION,H1,COMPANY,COSTCENTER,ASSET_CLASS,CAPITALIZED",
                keycolumn: "ID",
                maxrows: 2000,
                previewrows: 1000,
                allowcsv: true,
                autovalidate: true,
                showpreview: true,
                showlogs: true,
                stricttemplate: false,
                validationresult: "true",
                validationerrors: "[]",
                previewcompleted: "false",
                continueenabled: "false",
                lastevent: "",
                invalidcount: 0,
                validcount: 0,
                rowcount: 0
            };

            this._errorLog = [];
            this._previewRows = [];
            this._previewColumns = [];
            this._sheetName = "-";
            this._searchText = "";
            this._validationErrorsParsed = [];
            this._validationMap = {};
            this._uploadedHeaders = [];
            this._currentFileName = "";
            this._previewVisible = true;
            this._designMode = false;

            this._bindEvents();
        }

        connectedCallback() {
            this._setStatus("Ready", "ready");
            this._applyHeaderSettings();
            this._applyAcceptedTypes();
            this._applyVisibility();
            this._loadExcelLibrary();
            this._renderPreview();
            this._syncContinueButtons();
        }

        onCustomWidgetBeforeUpdate(changedProperties) {
            if ("designMode" in changedProperties) {
                this._designMode = changedProperties.designMode;
            }
        }

        onCustomWidgetAfterUpdate(changedProperties) {
            if ("title" in changedProperties) this.title = changedProperties.title;
            if ("subtitle" in changedProperties) this.subtitle = changedProperties.subtitle;
            if ("icon" in changedProperties) this.icon = changedProperties.icon;
            if ("unit" in changedProperties) this.unit = changedProperties.unit;
            if ("validpayload" in changedProperties) this.validpayload = changedProperties.validpayload;
            if ("footer" in changedProperties) this.footer = changedProperties.footer;
            if ("errorlogfilename" in changedProperties) this.errorlogfilename = changedProperties.errorlogfilename;
            if ("templatefilename" in changedProperties) this.templatefilename = changedProperties.templatefilename;
            if ("templateurl" in changedProperties) this.templateurl = changedProperties.templateurl;
            if ("requiredcolumns" in changedProperties) this.requiredcolumns = changedProperties.requiredcolumns;
            if ("keycolumn" in changedProperties) this.keycolumn = changedProperties.keycolumn;
            if ("maxrows" in changedProperties) this.maxrows = changedProperties.maxrows;
            if ("previewrows" in changedProperties) this.previewrows = changedProperties.previewrows;
            if ("allowcsv" in changedProperties) this.allowcsv = changedProperties.allowcsv;
            if ("autovalidate" in changedProperties) this.autovalidate = changedProperties.autovalidate;
            if ("showpreview" in changedProperties) this.showpreview = changedProperties.showpreview;
            if ("showlogs" in changedProperties) this.showlogs = changedProperties.showlogs;
            if ("stricttemplate" in changedProperties) this.stricttemplate = changedProperties.stricttemplate;
            if ("validationresult" in changedProperties) this.validationresult = changedProperties.validationresult;
            if ("validationerrors" in changedProperties) this.validationerrors = changedProperties.validationerrors;
            if ("previewcompleted" in changedProperties) this.previewcompleted = changedProperties.previewcompleted;
            if ("continueenabled" in changedProperties) this.continueenabled = changedProperties.continueenabled;
            if ("lastevent" in changedProperties) this.lastevent = changedProperties.lastevent;

            this._applyHeaderSettings();
            this._applyAcceptedTypes();
            this._applyVisibility();
            this._applySacValidation();
            this._syncContinueButtons();
        }

        static get observedAttributes() {
            return [
                "title","subtitle","icon","unit","validpayload","footer","errorlogfilename",
                "templatefilename","templateurl","requiredcolumns","keycolumn","maxrows",
                "previewrows","allowcsv","autovalidate","showpreview","showlogs","stricttemplate",
                "validationresult","validationerrors","previewcompleted","continueenabled",
                "lastevent","invalidcount","validcount","rowcount"
            ];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue !== newValue) this[name] = newValue;
        }

        get title() { return this._export_settings.title; }
        set title(v) { this._export_settings.title = v || ""; this._applyHeaderSettings(); }

        get subtitle() { return this._export_settings.subtitle; }
        set subtitle(v) { this._export_settings.subtitle = v || ""; this._applyHeaderSettings(); }

        get icon() { return this._export_settings.icon; }
        set icon(v) { this._export_settings.icon = v || ""; }

        get unit() { return this._export_settings.unit; }
        set unit(v) { this._export_settings.unit = v || ""; }

        get validpayload() { return this._export_settings.validpayload; }
        set validpayload(v) { this._export_settings.validpayload = v || ""; }

        get footer() { return this._export_settings.footer; }
        set footer(v) { this._export_settings.footer = v || ""; this._applyHeaderSettings(); }

        get errorlogfilename() { return this._export_settings.errorlogfilename; }
        set errorlogfilename(v) { this._export_settings.errorlogfilename = v || "Excel_Upload_Error_Log.csv"; }

        get templatefilename() { return this._export_settings.templatefilename; }
        set templatefilename(v) { this._export_settings.templatefilename = v || "New_Position_Creation_V2.1.xlsm"; }

        get templateurl() { return this._export_settings.templateurl; }
        set templateurl(v) { this._export_settings.templateurl = v || TEMPLATE_URL; }

        get requiredcolumns() { return this._export_settings.requiredcolumns; }
        set requiredcolumns(v) { this._export_settings.requiredcolumns = v || "ID,DESCRIPTION,H1,COMPANY,COSTCENTER,ASSET_CLASS,CAPITALIZED"; }

        get keycolumn() { return this._export_settings.keycolumn; }
        set keycolumn(v) { this._export_settings.keycolumn = v || "ID"; }

        get maxrows() { return this._export_settings.maxrows; }
        set maxrows(v) {
            var n = parseInt(v, 10);
            this._export_settings.maxrows = isNaN(n) || n <= 0 ? 2000 : n;
        }

        get previewrows() { return this._export_settings.previewrows; }
        set previewrows(v) {
            var n = parseInt(v, 10);
            this._export_settings.previewrows = isNaN(n) || n <= 0 ? 1000 : n;
            this._renderPreview();
        }

        get allowcsv() { return this._export_settings.allowcsv; }
        set allowcsv(v) { this._export_settings.allowcsv = this._toBoolean(v, true); }

        get autovalidate() { return this._export_settings.autovalidate; }
        set autovalidate(v) { this._export_settings.autovalidate = this._toBoolean(v, true); }

        get showpreview() { return this._export_settings.showpreview; }
        set showpreview(v) { this._export_settings.showpreview = this._toBoolean(v, true); this._applyVisibility(); }

        get showlogs() { return this._export_settings.showlogs; }
        set showlogs(v) { this._export_settings.showlogs = this._toBoolean(v, true); this._applyVisibility(); }

        get stricttemplate() { return this._export_settings.stricttemplate; }
        set stricttemplate(v) { this._export_settings.stricttemplate = this._toBoolean(v, false); }

        get validationresult() { return this._export_settings.validationresult; }
        set validationresult(v) { this._export_settings.validationresult = v || "true"; }

        get validationerrors() { return this._export_settings.validationerrors; }
        set validationerrors(v) { this._export_settings.validationerrors = v || "[]"; }

        get previewcompleted() { return this._export_settings.previewcompleted; }
        set previewcompleted(v) { this._export_settings.previewcompleted = v || "false"; }

        get continueenabled() { return this._export_settings.continueenabled; }
        set continueenabled(v) { this._export_settings.continueenabled = v || "false"; }

        get lastevent() { return this._export_settings.lastevent; }
        set lastevent(v) { this._export_settings.lastevent = v || ""; }

        _bindEvents() {
            var that = this;

            setTimeout(function () {
                that._shadowRoot.getElementById("downloadTemplateBtn").addEventListener("click", function () {
                    that._downloadTemplate();
                });

                that._shadowRoot.getElementById("downloadErrorBtn").addEventListener("click", function () {
                    that._downloadErrorLog();
                });

                that._shadowRoot.getElementById("previewBtn").addEventListener("click", function () {
                    that._processPreview();
                });

                that._shadowRoot.getElementById("continueBtn").addEventListener("click", function () {
                    that._continueUpload();
                });

                that._shadowRoot.getElementById("continueBottomBtn").addEventListener("click", function () {
                    that._continueUpload();
                });

                that._shadowRoot.getElementById("clearBtn").addEventListener("click", function () {
                    that.clear();
                });

                that._shadowRoot.getElementById("clearBottomBtn").addEventListener("click", function () {
                    that.clear();
                });

                that._shadowRoot.getElementById("togglePreviewBtn").addEventListener("click", function () {
                    that._previewVisible = !that._previewVisible;
                    that._applyVisibility();
                });

                that._shadowRoot.getElementById("searchInput").addEventListener("input", function (e) {
                    that._searchText = e.target.value || "";
                    that._renderPreview();
                });

                var dropZone = that._shadowRoot.getElementById("dropZone");

                dropZone.addEventListener("dragover", function (e) {
                    e.preventDefault();
                    dropZone.style.borderColor = "#0a6ed1";
                    dropZone.style.background = "#f0f8ff";
                });

                dropZone.addEventListener("dragleave", function () {
                    dropZone.style.borderColor = "#e3eaf2";
                    dropZone.style.background = "#fafcff";
                });

                dropZone.addEventListener("drop", function (e) {
                    e.preventDefault();
                    dropZone.style.borderColor = "#e3eaf2";
                    dropZone.style.background = "#fafcff";
                    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                        that._shadowRoot.getElementById("fileInput").files = e.dataTransfer.files;
                        that._log("File dropped: " + e.dataTransfer.files[0].name, false);
                    }
                });
            }, 0);
        }

        _applyHeaderSettings() {
            var titleEl = this._shadowRoot.getElementById("titleEl");
            var subtitleEl = this._shadowRoot.getElementById("subtitleEl");
            var footerEl = this._shadowRoot.getElementById("footerNote");
            var templateInfo = this._shadowRoot.getElementById("templateInfo");

            if (titleEl) titleEl.textContent = this._export_settings.title || "Excel Upload";
            if (subtitleEl) subtitleEl.textContent = this._export_settings.subtitle || "Preview, validate and upload Excel file";
            if (footerEl) footerEl.textContent = this._export_settings.footer || "Supported template: Sheet1 with required business columns";
            if (templateInfo) templateInfo.innerHTML = "Required Columns: <b>" + this._getRequiredColumns().join(", ") + "</b>";
        }

        _applyAcceptedTypes() {
            var input = this._shadowRoot.getElementById("fileInput");
            var accept = this._export_settings.allowcsv ? ".xls,.xlsx,.xlsm,.csv" : ".xls,.xlsx,.xlsm";
            if (input) input.setAttribute("accept", accept);
        }

        _applyVisibility() {
            var previewWrap = this._shadowRoot.getElementById("previewWrap");
            var logBox = this._shadowRoot.getElementById("logBox");

            if (previewWrap) {
                if (this._export_settings.showpreview && this._previewVisible) previewWrap.classList.add("show");
                else previewWrap.classList.remove("show");
            }

            if (logBox) {
                if (this._export_settings.showlogs && logBox.textContent !== "") logBox.classList.add("show");
                else logBox.classList.remove("show");
            }
        }

        _setStatus(text, type) {
            var badge = this._shadowRoot.getElementById("statusBadge");
            badge.textContent = text;
            badge.className = "status-badge";
            if (type === "processing") badge.classList.add("status-processing");
            else if (type === "completed") badge.classList.add("status-completed");
            else if (type === "error") badge.classList.add("status-error");
            else if (type === "warning") badge.classList.add("status-warning");
            else badge.classList.add("status-ready");
        }

        _setProgress(percent, text) {
            this._shadowRoot.getElementById("progressWrap").classList.add("show");
            this._shadowRoot.getElementById("progressFill").style.width = percent + "%";
            this._shadowRoot.getElementById("progressText").textContent = text || "";
            this._shadowRoot.getElementById("progressPercent").textContent = percent + "%";
        }

        _hideProgress() {
            this._shadowRoot.getElementById("progressWrap").classList.remove("show");
        }

        _setSummary(rows, valid, invalid, sheet, columns, validationText) {
            this._shadowRoot.getElementById("summaryGrid").classList.add("show");
            this._shadowRoot.getElementById("sumRows").textContent = rows || 0;
            this._shadowRoot.getElementById("sumValid").textContent = valid || 0;
            this._shadowRoot.getElementById("sumInvalid").textContent = invalid || 0;
            this._shadowRoot.getElementById("sumSheet").textContent = sheet || "-";
            this._shadowRoot.getElementById("sumColumns").textContent = columns || 0;
            this._shadowRoot.getElementById("sumValidation").textContent = validationText || "-";

            var summary = "Preview ready. Valid rows: " + (valid || 0) + ", Invalid rows: " + (invalid || 0) + ".";
            if ((invalid || 0) > 0) {
                summary += " Continue to Upload will upload only valid rows.";
            }
            this._shadowRoot.getElementById("bottomSummary").textContent = summary;
        }

        _log(message, reset) {
            var box = this._shadowRoot.getElementById("logBox");
            if (reset) box.textContent = "";
            if (!this._export_settings.showlogs) return;
            box.classList.add("show");
            box.textContent += (box.textContent ? "\n" : "") + message;
        }

        _enableErrorDownload(enable) {
            this._shadowRoot.getElementById("downloadErrorBtn").disabled = !enable;
        }

        _syncContinueButtons() {
            var enabled = this._export_settings.continueenabled === "true";
            this._shadowRoot.getElementById("continueBtn").disabled = !enabled;
            this._shadowRoot.getElementById("continueBottomBtn").disabled = !enabled;
        }

        _loadExcelLibrary() {
            var that = this;
            loadScriptOnce("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js", this._shadowRoot)
                .then(function () {
                    that._setStatus("Ready", "ready");
                    that._log("Excel library loaded successfully", true);
                })
                .catch(function () {
                    that._setStatus("Error", "error");
                    that._log("Failed to load Excel library", true);
                });
        }

        _downloadTemplate() {
            var that = this;
            var url = this._export_settings.templateurl;
            var fileName = this._export_settings.templatefilename || "Template.xlsm";
            var isSharePoint = /sharepoint\.com|sharepoint-df\.com/i.test(url);
            var isSacFileLink = /\/sap\/fpa\/ui\/app\.html#\/files/i.test(url);

            that._log("Downloading template...", false);
            that._setStatus("Downloading", "processing");

            if (isSharePoint || isSacFileLink) {
                try {
                    var aSp = document.createElement("a");
                    aSp.href = url;
                    aSp.target = "_blank";
                    aSp.rel = "noopener noreferrer";
                    document.body.appendChild(aSp);
                    aSp.click();
                    document.body.removeChild(aSp);
                    that._setStatus("Ready", "ready");
                    that._log("Opened template link: " + url);
                } catch (errSp) {
                    that._setStatus("Error", "error");
                    that._log("Template open failed: " + errSp.message, false);
                }
                return;
            }

            fetch(url)
                .then(function (response) {
                    if (!response.ok) throw new Error("HTTP " + response.status + " - Template file not found");
                    return response.blob();
                })
                .then(function (blob) {
                    var downloadUrl = URL.createObjectURL(blob);
                    var a = document.createElement("a");
                    a.href = downloadUrl;
                    a.download = fileName;
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(downloadUrl);
                    that._setStatus("Ready", "ready");
                    that._log("Template downloaded successfully: " + fileName);
                })
                .catch(function (err) {
                    that._setStatus("Error", "error");
                    that._log("Template download failed: " + err.message, false);
                });
        }

        _processPreview() {
            var that = this;
            var input = this._shadowRoot.getElementById("fileInput");
            var file = input.files && input.files[0] ? input.files[0] : null;

            if (!file) {
                this._setStatus("Warning", "warning");
                this._log("Please select a file before preview", true);
                return;
            }

            if (typeof XLSX === "undefined") {
                this._setStatus("Error", "error");
                this._log("Excel library is not loaded. Check internet/CDN access.", true);
                return;
            }

            this._currentFileName = file.name;
            this._setStatus("Processing", "processing");
            this._setProgress(10, "Reading file...");
            this._errorLog = [];
            this._previewRows = [];
            this._previewColumns = [];
            this._validationErrorsParsed = [];
            this._validationMap = {};
            this._uploadedHeaders = [];
            this._enableErrorDownload(false);
            this._export_settings.previewcompleted = "false";
            this._export_settings.continueenabled = "false";
            this._syncContinueButtons();
            this._log("File selected: " + file.name, true);

            var reader = new FileReader();

            reader.onload = function (e) {
                try {
                    that._setProgress(30, "Parsing workbook...");
                    var data = e.target.result;
                    var workbook = XLSX.read(data, { type: "binary" });
                    var actualSheet = workbook.SheetNames.indexOf("Sheet1") > -1 ? "Sheet1" : workbook.SheetNames[0];

                    if (!actualSheet) {
                        that._setStatus("Error", "error");
                        that._log("No sheet found in uploaded file", true);
                        return;
                    }

                    that._sheetName = actualSheet;
                    that._setProgress(45, "Reading rows...");
                    var sheet = workbook.Sheets[actualSheet];
                    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

                    if (!rows || rows.length <= 1) {
                        that._setStatus("Error", "error");
                        that._log("There is no record to preview", true);
                        return;
                    }

                    var header = rows[0] || [];
                    that._uploadedHeaders = [];
                    var h = 0;
                    for (h = 0; h < header.length; h++) {
                        that._uploadedHeaders.push(String(header[h] || "").trim());
                    }

                    that._setProgress(60, "Validating template...");
                    var prepare = that._prepareColumnMapFromHeaders();

                    if (!prepare.ok) {
                        that._setStatus("Error", "error");
                        that._log(prepare.message, true);
                        return;
                    }

                    that._setProgress(72, "Checking duplicates...");
                    var requiredColumns = that._getRequiredColumns();
                    var colMap = prepare.colMap;
                    var keyColumn = that._getKeyColumn(requiredColumns);
                    var idCount = {};
                    var r = 0;

                    for (r = 1; r < rows.length; r++) {
                        var row = rows[r] || [];
                        var id = String(row[colMap[keyColumn]] || "").trim();
                        if (!id) continue;
                        idCount[id] = (idCount[id] || 0) + 1;
                    }

                    that._setProgress(82, "Preparing preview...");
                    that._buildPreviewDataFromRows(rows, requiredColumns, colMap, idCount);

                    if (that._previewRows.length > that._export_settings.maxrows) {
                        that._setStatus("Error", "error");
                        that._log("Maximum rows are " + that._export_settings.maxrows, true);
                        return;
                    }

                    var allRowsForSac = [];
                    var validRowsForSac = [];
                    var pr = 0;
                    var pc = 0;

                    for (pr = 0; pr < that._previewRows.length; pr++) {
                        var srcRow = that._previewRows[pr];
                        var outRow = {};
                        for (pc = 0; pc < that._previewColumns.length; pc++) {
                            var colName = that._previewColumns[pc];
                            outRow[colName] = srcRow[colName] !== undefined && srcRow[colName] !== null ? String(srcRow[colName]) : "";
                        }
                        allRowsForSac.push(outRow);

                        if (!srcRow._localInvalid) {
                            validRowsForSac.push(outRow);
                        }
                    }

                    that.unit = JSON.stringify(allRowsForSac);
                    that.validpayload = JSON.stringify(validRowsForSac);

                    that._export_settings.rowcount = that._previewRows.length;
                    that._export_settings.validcount = validRowsForSac.length;
                    that._export_settings.invalidcount = that._getLocalInvalidRowCount();
                    that._export_settings.previewcompleted = "true";
                    that._export_settings.continueenabled = "true";

                    that._buildValidationMap();
                    that._renderPreview();
                    that._enableErrorDownload(that._errorLog.length > 0);
                    that._setSummary(
                        that._previewRows.length,
                        that._export_settings.validcount,
                        that._export_settings.invalidcount,
                        actualSheet,
                        requiredColumns.length,
                        that._export_settings.invalidcount > 0 ? "Invalid" : "Valid"
                    );

                    that._setProgress(100, "Preview ready");
                    that._setStatus("Preview Ready", "completed");
                    that._syncContinueButtons();
                    that._firePropertiesChanged("previewCompleted");

                    that.dispatchEvent(new CustomEvent("onPreview", {
                        detail: {
                            rowCount: that._export_settings.rowcount,
                            validCount: that._export_settings.validcount,
                            invalidCount: that._export_settings.invalidcount,
                            fileName: file.name,
                            sheetName: actualSheet
                        }
                    }));

                    that._log("Preview completed");
                    that._log("Rows read: " + that._previewRows.length);
                    that._log("Local valid rows: " + validRowsForSac.length);
                    that._log("Local invalid rows: " + that._getLocalInvalidRowCount());
                    that._log("Detected headers: " + that._uploadedHeaders.join(", "));
                } catch (err) {
                    that._setStatus("Error", "error");
                    that._log("Processing failed: " + err.message, true);
                }
            };

            reader.readAsBinaryString(file);
        }

        _continueUpload() {
            if (this._export_settings.previewcompleted !== "true") {
                this._log("Preview must be completed before continue upload", false);
                return;
            }

            this._export_settings.lastevent = "continueUpload";
            this._setStatus("Upload Requested", "processing");
            this._firePropertiesChanged("continueUpload");

            this.dispatchEvent(new CustomEvent("onContinueUpload", {
                detail: {
                    rowCount: this._export_settings.rowcount,
                    validCount: this._export_settings.validcount,
                    invalidCount: this._export_settings.invalidcount,
                    fileName: this._currentFileName,
                    sheetName: this._sheetName
                }
            }));
        }

        _prepareColumnMapFromHeaders() {
            var requiredColumns = this._getRequiredColumns();
            var colMap = {};
            var missing = [];
            var i = 0;
            var j = 0;

            for (i = 0; i < requiredColumns.length; i++) {
                var req = requiredColumns[i];
                colMap[req] = -1;

                for (j = 0; j < this._uploadedHeaders.length; j++) {
                    if (String(this._uploadedHeaders[j]).trim().toUpperCase() === String(req).trim().toUpperCase()) {
                        colMap[req] = j;
                        break;
                    }
                }

                if (colMap[req] === -1) {
                    missing.push(req);
                }
            }

            if (missing.length > 0) {
                return { ok: false, message: "Invalid template. Missing required columns: " + missing.join(", ") };
            }

            return { ok: true, colMap: colMap };
        }

        _buildPreviewDataFromRows(rows, requiredColumns, colMap, idCount) {
            var previewRows = [];
            var errorRows = [];
            var keyColumn = this._getKeyColumn(requiredColumns);
            var r = 0;
            var rc = 0;

            this._validationErrorsParsed = [];

            for (r = 1; r < rows.length; r++) {
                var row2 = rows[r] || [];
                var rowNumber = r + 1;
                var previewRowObj = { _rowNumber: rowNumber, _localInvalid: false };
                var isBlankRow = true;
                var errors = [];
                var cellErrors = [];

                for (rc = 0; rc < requiredColumns.length; rc++) {
                    var colName = requiredColumns[rc];
                    var cellValue = String(row2[colMap[colName]] || "").trim();
                    previewRowObj[colName] = cellValue;

                    if (cellValue !== "") {
                        isBlankRow = false;
                    }
                }

                if (isBlankRow) {
                    continue;
                }

                for (rc = 0; rc < requiredColumns.length; rc++) {
                    var reqCol = requiredColumns[rc];
                    if (!previewRowObj[reqCol]) {
                        errors.push(reqCol + " is mandatory");
                        cellErrors.push({
                            rowIndex: previewRows.length,
                            field: reqCol,
                            message: "Blank cell not allowed"
                        });
                    }
                }

                if (previewRowObj[keyColumn] && idCount[previewRowObj[keyColumn]] > 1) {
                    errors.push("Duplicate " + keyColumn + " '" + previewRowObj[keyColumn] + "'");
                    cellErrors.push({
                        rowIndex: previewRows.length,
                        field: keyColumn,
                        message: "Duplicate value not allowed"
                    });
                }

                if (errors.length > 0) {
                    previewRowObj._localInvalid = true;
                    errorRows.push({
                        RowNumber: rowNumber,
                        ErrorMessage: errors.join(" | ")
                    });

                    for (rc = 0; rc < cellErrors.length; rc++) {
                        this._validationErrorsParsed.push(cellErrors[rc]);
                    }
                }

                previewRows.push(previewRowObj);
            }

            this._previewRows = previewRows;
            this._previewColumns = requiredColumns.slice(0);
            this._errorLog = errorRows;
        }

        _getLocalInvalidRowCount() {
            var count = 0;
            var i = 0;
            for (i = 0; i < this._previewRows.length; i++) {
                if (this._previewRows[i]._localInvalid === true) count = count + 1;
            }
            return count;
        }

        _getRequiredColumns() {
            var raw = String(this._export_settings.requiredcolumns || "");
            var parts = raw.split(",");
            var out = [];
            var i = 0;

            for (i = 0; i < parts.length; i++) {
                var v = String(parts[i] || "").trim();
                if (v !== "") out.push(v);
            }

            return out.length ? out : ["ID", "DESCRIPTION", "H1", "COMPANY", "COSTCENTER", "ASSET_CLASS", "CAPITALIZED"];
        }

        _getKeyColumn(requiredColumns) {
            var key = String(this._export_settings.keycolumn || "").trim();
            if (key !== "") return key;
            if (requiredColumns.indexOf("ID") > -1) return "ID";
            return requiredColumns[0];
        }

        _applySacValidation() {
            this._validationErrorsParsed = this._safeParseArray(this._export_settings.validationerrors);
            this._buildValidationMap();
            this._renderPreview();

            var invalidRowMap = {};
            var i = 0;

            for (i = 0; i < this._validationErrorsParsed.length; i++) {
                var err = this._validationErrorsParsed[i];
                if (err.rowIndex !== undefined && err.rowIndex !== null) {
                    invalidRowMap[String(err.rowIndex)] = true;
                }
            }

            var invalidRowsFinal = 0;
            for (var k in invalidRowMap) {
                if (Object.prototype.hasOwnProperty.call(invalidRowMap, k)) invalidRowsFinal = invalidRowsFinal + 1;
            }

            var totalRowsFinal = this._previewRows.length;
            var validRowsFinal = totalRowsFinal - invalidRowsFinal;
            if (validRowsFinal < 0) validRowsFinal = 0;

            this._export_settings.rowcount = totalRowsFinal;
            this._export_settings.validcount = validRowsFinal;
            this._export_settings.invalidcount = invalidRowsFinal;

            this._setSummary(
                totalRowsFinal,
                validRowsFinal,
                invalidRowsFinal,
                this._sheetName,
                this._previewColumns.length,
                this._export_settings.validationresult === "false" ? "Invalid" : "Valid"
            );

            if (this._previewRows.length > 0) {
                if (this._export_settings.validationresult === "false") {
                    this._setStatus("Validation Error", "error");
                } else {
                    this._setStatus("Validated", "completed");
                }
                this._export_settings.continueenabled = "true";
            } else {
                this._setStatus("Ready", "ready");
                this._export_settings.continueenabled = "false";
            }

            this._syncContinueButtons();
        }

        _buildValidationMap() {
            this._validationMap = {};
            var i = 0;

            for (i = 0; i < this._validationErrorsParsed.length; i++) {
                var err = this._validationErrorsParsed[i];
                var rowIndex = err.rowIndex !== undefined ? String(err.rowIndex) : "";
                var field = err.field || "";
                var key = rowIndex + "|" + field;

                if (!this._validationMap[key]) this._validationMap[key] = [];
                this._validationMap[key].push(err.message || "Invalid value");

                var rowKey = rowIndex + "|__row__";
                if (!this._validationMap[rowKey]) this._validationMap[rowKey] = [];
                this._validationMap[rowKey].push(err.message || "Invalid value");
            }
        }

        _renderPreview() {
            var wrap = this._shadowRoot.getElementById("previewWrap");
            var head = this._shadowRoot.getElementById("previewHead");
            var body = this._shadowRoot.getElementById("previewBody");
            var meta = this._shadowRoot.getElementById("previewMeta");

            if (!this._export_settings.showpreview || !this._previewVisible) {
                wrap.classList.remove("show");
                return;
            }

            wrap.classList.add("show");
            head.innerHTML = "";
            body.innerHTML = "";

            if (!this._previewRows || this._previewRows.length === 0) {
                meta.textContent = "0 row(s) shown";
                body.innerHTML = '<tr><td colspan="50">No preview available</td></tr>';
                return;
            }

            var cols = this._previewColumns && this._previewColumns.length ? this._previewColumns.slice(0) : [];
            var filteredRows = [];
            var sText = String(this._searchText || "").toLowerCase().trim();
            var i = 0;
            var j = 0;

            for (i = 0; i < this._previewRows.length; i++) {
                var row = this._previewRows[i];
                if (sText === "") {
                    filteredRows.push(row);
                } else {
                    var matched = false;
                    for (j = 0; j < cols.length; j++) {
                        var val = String(row[cols[j]] || "").toLowerCase();
                        if (val.indexOf(sText) > -1) {
                            matched = true;
                            break;
                        }
                    }
                    if (matched) filteredRows.push(row);
                }
            }

            var limit = this._export_settings.previewrows;
            var totalMatched = filteredRows.length;
            if (filteredRows.length > limit) filteredRows = filteredRows.slice(0, limit);

            meta.textContent = filteredRows.length + " row(s) shown" + (totalMatched > limit ? " of " + totalMatched : "");

            var headHtml = "<tr><th>#</th>";
            for (j = 0; j < cols.length; j++) {
                headHtml += "<th>" + this._escapeHtml(cols[j]) + "</th>";
            }
            headHtml += "<th>STATUS</th><th>ERRORS</th></tr>";
            head.innerHTML = headHtml;

            var bodyHtml = "";
            for (i = 0; i < filteredRows.length; i++) {
                var rowObj = filteredRows[i];
                var originalIndex = this._getOriginalPreviewIndex(rowObj);
                var rowKey = String(originalIndex) + "|__row__";
                var rowErrors = this._validationMap[rowKey] || [];
                var rowClass = rowErrors.length > 0 ? "invalid-row" : "";
                var statusText = rowErrors.length > 0 ? "Invalid" : "Valid";
                var statusClass = rowErrors.length > 0 ? "row-status-invalid" : "row-status-valid";

                bodyHtml += '<tr class="' + rowClass + '">';
                bodyHtml += "<td>" + this._escapeHtml(String(rowObj._rowNumber || (originalIndex + 2))) + "</td>";

                for (j = 0; j < cols.length; j++) {
                    var field = cols[j];
                    var cellKey = String(originalIndex) + "|" + field;
                    var invalidClass = this._validationMap[cellKey] ? "invalid-cell" : "";
                    var cellVal = rowObj[field] === undefined || rowObj[field] === null ? "" : String(rowObj[field]);

                    bodyHtml += '<td class="' + invalidClass + '">';
                    bodyHtml += this._escapeHtml(cellVal);

                    if (this._validationMap[cellKey]) {
                        for (var m = 0; m < this._validationMap[cellKey].length; m++) {
                            bodyHtml += '<span class="cell-error">' + this._escapeHtml(this._validationMap[cellKey][m]) + '</span>';
                        }
                    }
                    bodyHtml += "</td>";
                }

                bodyHtml += '<td><span class="' + statusClass + '">' + statusText + "</span></td>";
                bodyHtml += "<td>" + this._escapeHtml(rowErrors.join(" | ")) + "</td>";
                bodyHtml += "</tr>";
            }

            body.innerHTML = bodyHtml;
        }

        _getOriginalPreviewIndex(rowObj) {
            var i = 0;
            for (i = 0; i < this._previewRows.length; i++) {
                if (this._previewRows[i] === rowObj) return i;
                if (this._previewRows[i]._rowNumber === rowObj._rowNumber) return i;
            }
            return 0;
        }

        clear() {
            this._shadowRoot.getElementById("fileInput").value = "";
            this._shadowRoot.getElementById("searchInput").value = "";
            this._searchText = "";
            this.unit = "";
            this.validpayload = "";
            this._errorLog = [];
            this._previewRows = [];
            this._previewColumns = [];
            this._validationErrorsParsed = [];
            this._validationMap = {};
            this._uploadedHeaders = [];
            this._export_settings.validationresult = "true";
            this._export_settings.validationerrors = "[]";
            this._export_settings.previewcompleted = "false";
            this._export_settings.continueenabled = "false";
            this._sheetName = "-";
            this._currentFileName = "";
            this._enableErrorDownload(false);
            this._setStatus("Ready", "ready");
            this._hideProgress();
            this._setSummary(0, 0, 0, "-", 0, "-");
            this._renderPreview();
            this._log("Cleared previous file and output", true);
            this._export_settings.lastevent = "clear";
            this._syncContinueButtons();
            this._firePropertiesChanged("clear");

            this.dispatchEvent(new CustomEvent("onClear", { detail: {} }));
        }

        _downloadErrorLog() {
            if (!this._errorLog || this._errorLog.length === 0) {
                this._log("No error log available to download");
                return;
            }

            var rows = ["RowNumber,ErrorMessage"];
            var i = 0;
            for (i = 0; i < this._errorLog.length; i++) {
                var item = this._errorLog[i];
                rows.push([this._escapeCsv(item.RowNumber), this._escapeCsv(item.ErrorMessage)].join(","));
            }

            this._downloadBlob(rows.join("\n"), "text/csv;charset=utf-8;", this._export_settings.errorlogfilename);
            this._log("Error log downloaded successfully");
        }

        _safeParseArray(text) {
            try {
                var parsed = JSON.parse(text || "[]");
                return Array.isArray(parsed) ? parsed : [];
            } catch (e) {
                return [];
            }
        }

        _downloadBlob(content, mimeType, fileName) {
            var blob = new Blob([content], { type: mimeType });
            var url = URL.createObjectURL(blob);
            var a = document.createElement("a");
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        _escapeCsv(value) {
            var str = value == null ? "" : String(value);
            if (str.indexOf(",") > -1 || str.indexOf('"') > -1 || str.indexOf("\n") > -1) {
                str = '"' + str.replace(/"/g, '""') + '"';
            }
            return str;
        }

        _firePropertiesChanged(eventName) {
            this._export_settings.lastevent = eventName || "";

            this.dispatchEvent(new CustomEvent("propertiesChanged", {
                detail: {
                    properties: {
                        unit: this.unit,
                        validpayload: this.validpayload,
                        validationresult: this._export_settings.validationresult,
                        validationerrors: this._export_settings.validationerrors,
                        previewcompleted: this._export_settings.previewcompleted,
                        continueenabled: this._export_settings.continueenabled,
                        lastevent: this._export_settings.lastevent,
                        invalidcount: this._export_settings.invalidcount,
                        validcount: this._export_settings.validcount,
                        rowcount: this._export_settings.rowcount
                    }
                }
            }));
        }

        _toBoolean(v, defaultValue) {
            if (v === true || v === "true") return true;
            if (v === false || v === "false") return false;
            return defaultValue;
        }

        _escapeHtml(str) {
            return String(str)
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;");
        }
    }

    customElements.define("com-fd-djaja-sap-sac-excelcom", Excel);

    function loadScriptOnce(src, shadowRoot) {
        return new Promise(function (resolve, reject) {
            if (typeof XLSX !== "undefined") {
                resolve();
                return;
            }

            var existing = shadowRoot.querySelector('script[src="' + src + '"]');
            if (existing) {
                existing.addEventListener("load", resolve);
                existing.addEventListener("error", reject);
                return;
            }

            var script = document.createElement("script");
            script.src = src;
            script.onload = resolve;
            script.onerror = reject;
            shadowRoot.appendChild(script);
        });
    }
})();




/*(function () {
    "use strict";

    var tmpl = document.createElement("template");
    tmpl.innerHTML = `
        <style>
            :host{
                display:block;
                width:100%;
                height:100%;
                font-family:"72", Arial, Helvetica, sans-serif;
                color:#1f2d3d;
                --cw-primary:#0a6ed1;
                --cw-primary-dark:#085caf;
                --cw-border:#d9d9d9;
                --cw-soft:#f7f9fb;
                --cw-soft2:#fafcff;
                --cw-text:#1f2d3d;
                --cw-sub:#6a6d70;
                --cw-success-bg:#f6ffed;
                --cw-success-bd:#b7eb8f;
                --cw-success-tx:#237804;
                --cw-error-bg:#fff1f0;
                --cw-error-bd:#ffa39e;
                --cw-error-tx:#a8071a;
                --cw-warn-bg:#fffbe6;
                --cw-warn-bd:#ffe58f;
                --cw-warn-tx:#ad6800;
                --cw-info-bg:#f0f7ff;
                --cw-info-bd:#bae0ff;
                --cw-info-tx:#0958d9;
            }

            *{
                box-sizing:border-box;
            }

            .excel-widget{
                width:100%;
                height:100%;
            }

            .upload-card{
                border:1px solid var(--cw-border);
                border-radius:12px;
                background:linear-gradient(180deg, #ffffff 0%, #fbfdff 100%);
                box-shadow:0 6px 18px rgba(0,0,0,0.06);
                padding:10px;
                height:100%;
                display:flex;
                flex-direction:column;
                overflow:hidden;
                gap:8px;
            }

            .upload-header,
            .toolbar-row{
                display:none;
            }

            .top-bar{
                display:flex;
                justify-content:space-between;
                align-items:center;
                gap:10px;
                flex-wrap:wrap;
                padding-bottom:6px;
                border-bottom:none;
            }

            .top-bar-left{
                display:flex;
                align-items:flex-start;
                min-width:auto;
            }

            .top-bar-right{
                display:flex;
                align-items:center;
                justify-content:flex-end;
                gap:8px;
                flex-wrap:wrap;
            }

            .title-wrap{
                display:flex;
                flex-direction:column;
                gap:2px;
            }

            .upload-title{
                font-size:15px;
                font-weight:700;
                color:var(--cw-primary);
                line-height:1.15;
                letter-spacing:0;
            }

            .upload-subtitle{
                font-size:11px;
                color:var(--cw-sub);
                line-height:1.3;
            }

            .status-badge{
                font-size:11px;
                font-weight:700;
                padding:6px 10px;
                border-radius:9px;
                background:#f5f6f7;
                color:#354a5f;
                border:1px solid var(--cw-border);
                white-space:nowrap;
                display:inline-flex;
                align-items:center;
                justify-content:center;
                min-height:34px;
                box-shadow:none;
            }

            .status-ready{
                background:#f5f6f7;
                color:#354a5f;
                border-color:var(--cw-border);
            }

            .status-processing{
                background:#fff7e6;
                color:#8a5a00;
                border-color:#ffd591;
            }

            .status-completed{
                background:var(--cw-success-bg);
                color:var(--cw-success-tx);
                border-color:var(--cw-success-bd);
            }

            .status-error{
                background:var(--cw-error-bg);
                color:var(--cw-error-tx);
                border-color:var(--cw-error-bd);
            }

            .status-warning{
                background:var(--cw-warn-bg);
                color:var(--cw-warn-tx);
                border-color:var(--cw-warn-bd);
            }

            .toolbar-btn{
                appearance:none;
                border:1px solid #c7ced4;
                background:#ffffff;
                color:var(--cw-text);
                border-radius:9px;
                padding:8px 14px;
                font-size:12px;
                font-weight:700;
                cursor:pointer;
                transition:all 0.2s ease;
                min-height:34px;
                min-width:auto;
                box-shadow:none;
            }

            .toolbar-btn:hover{
                border-color:var(--cw-primary);
                color:var(--cw-primary);
                background:#f4f9ff;
            }

            .toolbar-btn.primary{
                background:var(--cw-primary);
                color:#ffffff;
                border-color:var(--cw-primary);
            }

            .toolbar-btn.primary:hover{
                background:var(--cw-primary-dark);
                color:#ffffff;
                border-color:var(--cw-primary-dark);
            }

            .toolbar-btn:disabled{
                opacity:0.55;
                cursor:not-allowed;
            }

            .upload-area{
                border:1px solid #e3eaf2;
                border-radius:12px;
                padding:12px;
                background:#fafcff;
                margin-bottom:0;
                flex:0 0 auto;
            }

            .upload-top-row{
                display:flex;
                justify-content:space-between;
                align-items:center;
                gap:14px;
                flex-wrap:wrap;
                padding-bottom:0;
                border-bottom:none;
            }

           .upload-left{
            display:flex;
            align-items:center;
            gap:10px;
            flex-wrap:nowrap;
            flex:0 1 auto;
            min-width:auto;
        }


            .upload-right{
                display:flex;
                align-items:center;
                justify-content:flex-end;
                flex:0 0 320px;
                max-width:100%;
            }

            .file-input{
            font-size:13px;
            padding:10px 12px;
            border:1px solid var(--cw-border);
            border-radius:10px;
            background:#fff;
            width:300px;
            min-width:300px;
            max-width:300px;
            height:42px;
            flex:0 0 300px;
            color:var(--cw-text);
        }


           .action-btn{
            min-width:92px;
            width:92px;
            height:42px;
            padding:0 14px;
            font-size:13px;
            border-radius:10px;
            flex:0 0 92px;
        }


            .search-input{
                width:100%;
                min-width:260px;
                padding:10px 12px;
                border:1px solid var(--cw-border);
                border-radius:10px;
                font-size:13px;
                outline:none;
                background:#fff;
                height:42px;
                color:var(--cw-text);
            }

            .search-input:focus{
                border-color:var(--cw-primary);
                box-shadow:0 0 0 2px rgba(10,110,209,0.12);
            }

            .mini-text{
                font-size:11px;
                color:var(--cw-sub);
                margin-top:8px;
                white-space:normal;
                line-height:1.4;
            }

            .mini-text b{
                color:var(--cw-text);
            }

            .summary-grid{
                display:none;
                grid-template-columns:repeat(6, minmax(0, 1fr));
                gap:8px;
                margin-top:10px;
                flex:0 0 auto;
            }

            .summary-grid.show{
                display:grid;
            }

            .summary-item{
                border:1px solid #edf1f5;
                border-radius:10px;
                padding:8px 10px;
                background:#ffffff;
                min-height:52px;
                box-shadow:none;
            }

            .summary-item .k{
                font-size:11px;
                color:var(--cw-sub);
                margin-bottom:3px;
                font-weight:600;
            }

            .summary-item .v{
                font-size:15px;
                font-weight:700;
                color:var(--cw-text);
                line-height:1.2;
            }

            .progress-wrap{
                display:none;
                margin-top:0;
                flex:0 0 auto;
            }

            .progress-wrap.show{
                display:block;
            }

            .progress-label-row{
                display:flex;
                justify-content:space-between;
                align-items:center;
                margin-bottom:6px;
                font-size:12px;
                color:#354a5f;
            }

            .progress-bar{
                width:100%;
                height:11px;
                background:#edf2f7;
                border-radius:10px;
                overflow:hidden;
                border:1px solid var(--cw-border);
            }

            .progress-fill{
                width:0%;
                height:100%;
                background:linear-gradient(90deg, var(--cw-primary), #4db1ff);
                transition:width 0.25s ease;
            }

            .msg{
                display:none !important;
            }

            .msg.show{
                display:none !important;
            }

            .msg.error{
                background:var(--cw-error-bg);
                border:1px solid var(--cw-error-bd);
                color:var(--cw-error-tx);
            }

            .msg.warn{
                background:var(--cw-warn-bg);
                border:1px solid var(--cw-warn-bd);
                color:#8a6d1d;
            }

            .msg.success{
                background:var(--cw-success-bg);
                border:1px solid var(--cw-success-bd);
                color:var(--cw-success-tx);
            }

            .msg.info{
                background:var(--cw-info-bg);
                border:1px solid var(--cw-info-bd);
                color:var(--cw-info-tx);
            }

            .preview-wrap{
                display:none;
                margin-top:0;
                border:1px solid #dfe6ee;
                border-radius:12px;
                overflow:hidden;
                background:#ffffff;
                flex:1 1 auto;
                min-height:0;
            }

            .preview-wrap.show{
                display:flex;
                flex-direction:column;
            }

            .section-head{
                display:flex;
                justify-content:space-between;
                align-items:center;
                padding:8px 10px;
                background:var(--cw-soft);
                border-bottom:1px solid #e5e7eb;
                gap:8px;
                flex-wrap:wrap;
                flex:0 0 auto;
            }

            .section-title{
                font-size:12px;
                font-weight:700;
                color:#354a5f;
            }

            .section-meta{
                font-size:11px;
                color:var(--cw-sub);
            }

            .preview-grid{
                overflow:auto;
                flex:1 1 auto;
                min-height:0;
                height:100%;
                max-height:none;
                padding-top:0;
            }

            table{
                width:max-content;
                min-width:100%;
                border-collapse:collapse;
                border:none;
                border-radius:0;
            }

            th,
            td{
                border-bottom:1px solid #eef2f6;
                padding:7px 10px;
                text-align:left;
                vertical-align:top;
                font-size:12px;
                white-space:nowrap;
                color:var(--cw-text);
            }

            th{
                position:sticky;
                top:0;
                z-index:1;
                background:#f8fbff;
                color:#354a5f;
                font-weight:700;
            }

            tr:hover td{
                background:#fafcff;
            }

            tr.invalid-row td{
                background:#fff7f7;
            }

            td.invalid-cell{
                background:#fff1f0 !important;
                border-left:3px solid #ff4d4f;
                color:inherit;
                font-weight:inherit;
            }

            .cell-error{
                display:block;
                margin-top:4px;
                color:#cf1322;
                font-size:11px;
                line-height:1.3;
                white-space:normal;
                font-weight:400;
            }

            .log-box{
                display:none;
                margin-top:0;
                border:1px solid #e5e7eb;
                border-radius:10px;
                background:#fafbfc;
                padding:8px 10px;
                max-height:100px;
                overflow:auto;
                font-size:11px;
                line-height:1.4;
                color:#354a5f;
                white-space:pre-wrap;
                flex:0 0 auto;
            }

            .log-box.show{
                display:block;
            }

            .footer-note{
                margin-top:0;
                font-size:11px;
                color:var(--cw-sub);
                white-space:pre-wrap;
                flex:0 0 auto;
            }

@media (max-width: 980px){
    .upload-top-row{
        flex-direction:column;
        align-items:stretch;
    }

    .upload-left{
        min-width:100%;
        flex-wrap:wrap;
    }

    .upload-right{
        flex:1 1 auto;
        width:100%;
    }

    .file-input{
        min-width:100%;
        max-width:100%;
        width:100%;
        flex:1 1 auto;
    }

    .search-input{
        min-width:100%;
    }

    .toolbar-btn,
    .action-btn{
        min-width:auto;
    }
}

        </style>

        <div class="excel-widget">
            <div class="upload-card">
                <div class="top-bar">
                    <div class="top-bar-left">
                        <div class="title-wrap">
                            <div class="upload-title" id="titleEl">Excel Upload</div>
                            <div class="upload-subtitle" id="subtitleEl">Upload and validate Excel file</div>
                        </div>
                    </div>

                    <div class="top-bar-right">
                        <button type="button" class="toolbar-btn primary" id="downloadTemplateBtn">Download Template</button>
                        <div class="status-badge status-ready" id="statusBadge">Ready</div>
                        <button type="button" class="toolbar-btn" id="downloadErrorBtn" disabled>Download Error Log</button>
                        <button type="button" class="toolbar-btn" id="togglePreviewBtn">Toggle Preview</button>
                    </div>
                </div>

                <div class="upload-area" id="dropZone">
                    <div class="upload-top-row">
                        <div class="upload-left">
                            <input type="file" id="fileInput" class="file-input" />
                            <button type="button" class="toolbar-btn primary action-btn" id="uploadBtn">Upload</button>
                            <button type="button" class="toolbar-btn action-btn" id="clearBtn">Clear</button>
                        </div>

                        <div class="upload-right">
                            <input type="text" id="searchInput" class="search-input" placeholder="Search in preview..." />
                        </div>
                    </div>

                    <div class="mini-text" id="templateInfo">Required columns will be validated automatically.</div>

                    <div class="summary-grid" id="summaryGrid">
                        <div class="summary-item"><div class="k">Rows Read</div><div class="v" id="sumRows">0</div></div>
                        <div class="summary-item"><div class="k">Valid Rows</div><div class="v" id="sumValid">0</div></div>
                        <div class="summary-item"><div class="k">Invalid Rows</div><div class="v" id="sumInvalid">0</div></div>
                        <div class="summary-item"><div class="k">Sheet</div><div class="v" id="sumSheet">-</div></div>
                        <div class="summary-item"><div class="k">Columns</div><div class="v" id="sumColumns">0</div></div>
                        <div class="summary-item"><div class="k">Validation</div><div class="v" id="sumValidation">-</div></div>
                    </div>
                </div>

                <div class="progress-wrap" id="progressWrap">
                    <div class="progress-label-row">
                        <span id="progressText">Preparing upload...</span>
                        <span id="progressPercent">0%</span>
                    </div>
                    <div class="progress-bar">
                        <div class="progress-fill" id="progressFill"></div>
                    </div>
                </div>

                <div class="msg" id="messageBox"></div>

                <div class="preview-wrap" id="previewWrap">
                    <div class="section-head">
                        <div class="section-title">Preview</div>
                        <div class="section-meta" id="previewMeta">0 row(s) shown</div>
                    </div>
                    <div class="preview-grid">
                        <table>
                            <thead id="previewHead"></thead>
                            <tbody id="previewBody"></tbody>
                        </table>
                    </div>
                </div>

                <div class="log-box" id="logBox"></div>
                <div class="footer-note" id="footerNote">Supported template: Sheet1 with required business columns</div>
            </div>
        </div>
    `;

    var TEMPLATE_URL = "https://raw.githubusercontent.com/sacplanning2025/Company_widget/main/New_Position_Creation_V2.1.xlsm";

    class Excel extends HTMLElement {
        constructor() {
            super();

            this._shadowRoot = this.attachShadow({ mode: "open" });
            this._shadowRoot.appendChild(tmpl.content.cloneNode(true));

            this._export_settings = {
                title: "",
                subtitle: "",
                icon: "",
                unit: "",
                footer: "",
                errorlogfilename: "Excel_Upload_Error_Log.csv",
                templatefilename: "New_Position_Creation_V2.1.xlsm",
                templateurl: TEMPLATE_URL,
                requiredcolumns: "ID,DESCRIPTION,H1,COMPANY,COSTCENTER,ASSET_CLASS,CAPITALIZED",
                keycolumn: "ID",
                maxrows: 2000,
                previewrows: 1000,
                allowcsv: true,
                autovalidate: true,
                showpreview: true,
                showlogs: true,
                stricttemplate: false,
                validationresult: "true",
                validationerrors: "[]",
                lastevent: "",
                invalidcount: 0,
                validcount: 0,
                rowcount: 0
            };

            this._errorLog = [];
            this._validData = [];
            this._previewRows = [];
            this._previewColumns = [];
            this._sheetName = "-";
            this._searchText = "";
            this._validationErrorsParsed = [];
            this._validationMap = {};
            this._uploadedHeaders = [];
            this._currentFileName = "";
            this._previewVisible = true;
            this._designMode = false;

            this._bindEvents();
        }

        connectedCallback() {
            this._setStatus("Ready", "ready");
            this._applyHeaderSettings();
            this._applyAcceptedTypes();
            this._applyVisibility();
            this._loadExcelLibrary();
            this._renderPreview();
        }

        onCustomWidgetBeforeUpdate(changedProperties) {
            if ("designMode" in changedProperties) {
                this._designMode = changedProperties.designMode;
            }
        }

        onCustomWidgetAfterUpdate(changedProperties) {
            if ("title" in changedProperties) this.title = changedProperties.title;
            if ("subtitle" in changedProperties) this.subtitle = changedProperties.subtitle;
            if ("icon" in changedProperties) this.icon = changedProperties.icon;
            if ("unit" in changedProperties) this.unit = changedProperties.unit;
            if ("footer" in changedProperties) this.footer = changedProperties.footer;
            if ("errorlogfilename" in changedProperties) this.errorlogfilename = changedProperties.errorlogfilename;
            if ("templatefilename" in changedProperties) this.templatefilename = changedProperties.templatefilename;
            if ("templateurl" in changedProperties) this.templateurl = changedProperties.templateurl;
            if ("requiredcolumns" in changedProperties) this.requiredcolumns = changedProperties.requiredcolumns;
            if ("keycolumn" in changedProperties) this.keycolumn = changedProperties.keycolumn;
            if ("maxrows" in changedProperties) this.maxrows = changedProperties.maxrows;
            if ("previewrows" in changedProperties) this.previewrows = changedProperties.previewrows;
            if ("allowcsv" in changedProperties) this.allowcsv = changedProperties.allowcsv;
            if ("autovalidate" in changedProperties) this.autovalidate = changedProperties.autovalidate;
            if ("showpreview" in changedProperties) this.showpreview = changedProperties.showpreview;
            if ("showlogs" in changedProperties) this.showlogs = changedProperties.showlogs;
            if ("stricttemplate" in changedProperties) this.stricttemplate = changedProperties.stricttemplate;
            if ("validationresult" in changedProperties) this.validationresult = changedProperties.validationresult;
            if ("validationerrors" in changedProperties) this.validationerrors = changedProperties.validationerrors;
            if ("lastevent" in changedProperties) this.lastevent = changedProperties.lastevent;

            this._applyHeaderSettings();
            this._applyAcceptedTypes();
            this._applyVisibility();
            this._applySacValidation();
        }

        static get observedAttributes() {
            return [
                "title",
                "subtitle",
                "icon",
                "unit",
                "footer",
                "errorlogfilename",
                "templatefilename",
                "templateurl",
                "requiredcolumns",
                "keycolumn",
                "maxrows",
                "previewrows",
                "allowcsv",
                "autovalidate",
                "showpreview",
                "showlogs",
                "stricttemplate",
                "validationresult",
                "validationerrors",
                "lastevent",
                "invalidcount",
                "validcount",
                "rowcount"
            ];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue !== newValue) {
                this[name] = newValue;
            }
        }

        get title() { return this._export_settings.title; }
        set title(v) { this._export_settings.title = v || ""; this._applyHeaderSettings(); }

        get subtitle() { return this._export_settings.subtitle; }
        set subtitle(v) { this._export_settings.subtitle = v || ""; this._applyHeaderSettings(); }

        get icon() { return this._export_settings.icon; }
        set icon(v) { this._export_settings.icon = v || ""; }

        get unit() { return this._export_settings.unit; }
        set unit(v) { this._export_settings.unit = v || ""; }

        get footer() { return this._export_settings.footer; }
        set footer(v) { this._export_settings.footer = v || ""; this._applyHeaderSettings(); }

        get errorlogfilename() { return this._export_settings.errorlogfilename; }
        set errorlogfilename(v) { this._export_settings.errorlogfilename = v || "Excel_Upload_Error_Log.csv"; }

        get templatefilename() { return this._export_settings.templatefilename; }
        set templatefilename(v) { this._export_settings.templatefilename = v || "New_Position_Creation_V2.1.xlsm"; }

        get templateurl() { return this._export_settings.templateurl; }
        set templateurl(v) { this._export_settings.templateurl = v || TEMPLATE_URL; }

        get requiredcolumns() { return this._export_settings.requiredcolumns; }
        set requiredcolumns(v) { this._export_settings.requiredcolumns = v || "ID,DESCRIPTION,H1,COMPANY,COSTCENTER,ASSET_CLASS,CAPITALIZED"; }

        get keycolumn() { return this._export_settings.keycolumn; }
        set keycolumn(v) { this._export_settings.keycolumn = v || "ID"; }

        get maxrows() { return this._export_settings.maxrows; }
        set maxrows(v) {
            var n = parseInt(v, 10);
            this._export_settings.maxrows = isNaN(n) || n <= 0 ? 2000 : n;
        }

        get previewrows() { return this._export_settings.previewrows; }
        set previewrows(v) {
            var n = parseInt(v, 10);
            this._export_settings.previewrows = isNaN(n) || n <= 0 ? 1000 : n;
            this._renderPreview();
        }

        get allowcsv() { return this._export_settings.allowcsv; }
        set allowcsv(v) { this._export_settings.allowcsv = this._toBoolean(v, true); }

        get autovalidate() { return this._export_settings.autovalidate; }
        set autovalidate(v) { this._export_settings.autovalidate = this._toBoolean(v, true); }

        get showpreview() { return this._export_settings.showpreview; }
        set showpreview(v) { this._export_settings.showpreview = this._toBoolean(v, true); this._applyVisibility(); }

        get showlogs() { return this._export_settings.showlogs; }
        set showlogs(v) { this._export_settings.showlogs = this._toBoolean(v, true); this._applyVisibility(); }

        get stricttemplate() { return this._export_settings.stricttemplate; }
        set stricttemplate(v) { this._export_settings.stricttemplate = this._toBoolean(v, false); }

        get validationresult() { return this._export_settings.validationresult; }
        set validationresult(v) { this._export_settings.validationresult = v || "true"; this._applySacValidation(); }

        get validationerrors() { return this._export_settings.validationerrors; }
        set validationerrors(v) { this._export_settings.validationerrors = v || "[]"; this._applySacValidation(); }

        get lastevent() { return this._export_settings.lastevent; }
        set lastevent(v) { this._export_settings.lastevent = v || ""; }

        _bindEvents() {
            var that = this;

            setTimeout(function () {
                that._shadowRoot.getElementById("downloadTemplateBtn").addEventListener("click", function () {
                    that._downloadTemplate();
                });

                that._shadowRoot.getElementById("downloadErrorBtn").addEventListener("click", function () {
                    that._downloadErrorLog();
                });

                that._shadowRoot.getElementById("uploadBtn").addEventListener("click", function () {
                    that._processUpload();
                });

                that._shadowRoot.getElementById("clearBtn").addEventListener("click", function () {
                    that.clear();
                });

                that._shadowRoot.getElementById("togglePreviewBtn").addEventListener("click", function () {
                    that._previewVisible = !that._previewVisible;
                    that._applyVisibility();
                });

                that._shadowRoot.getElementById("searchInput").addEventListener("input", function (e) {
                    that._searchText = e.target.value || "";
                    that._renderPreview();
                });

                var dropZone = that._shadowRoot.getElementById("dropZone");

                dropZone.addEventListener("dragover", function (e) {
                    e.preventDefault();
                    dropZone.style.borderColor = "#0a6ed1";
                    dropZone.style.background = "#f0f8ff";
                });

                dropZone.addEventListener("dragleave", function () {
                    dropZone.style.borderColor = "#b8c4d1";
                    dropZone.style.background = "";
                });

                dropZone.addEventListener("drop", function (e) {
                    e.preventDefault();
                    dropZone.style.borderColor = "#b8c4d1";
                    dropZone.style.background = "";
                    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                        that._shadowRoot.getElementById("fileInput").files = e.dataTransfer.files;
                        that._log("File dropped: " + e.dataTransfer.files[0].name, false);
                    }
                });
            }, 0);
        }

        _applyHeaderSettings() {
            var titleEl = this._shadowRoot.getElementById("titleEl");
            var subtitleEl = this._shadowRoot.getElementById("subtitleEl");
            var footerEl = this._shadowRoot.getElementById("footerNote");
            var templateInfo = this._shadowRoot.getElementById("templateInfo");

            if (titleEl) titleEl.textContent = this._export_settings.title || "Excel Upload";
            if (subtitleEl) subtitleEl.textContent = this._export_settings.subtitle || "Upload and validate Excel file";
            if (footerEl) footerEl.textContent = this._export_settings.footer || "Supported template: Sheet1 with required business columns";
            if (templateInfo) templateInfo.innerHTML = "Required Columns: <b>" + this._getRequiredColumns().join(", ") + "</b>";
        }

        _applyAcceptedTypes() {
            var input = this._shadowRoot.getElementById("fileInput");
            var accept = this._export_settings.allowcsv ? ".xls,.xlsx,.xlsm,.csv" : ".xls,.xlsx,.xlsm";
            if (input) input.setAttribute("accept", accept);
        }

        _applyVisibility() {
            var previewWrap = this._shadowRoot.getElementById("previewWrap");
            var logBox = this._shadowRoot.getElementById("logBox");

            if (previewWrap) {
                if (this._export_settings.showpreview && this._previewVisible) previewWrap.classList.add("show");
                else previewWrap.classList.remove("show");
            }

            if (logBox) {
                if (this._export_settings.showlogs && logBox.textContent !== "") logBox.classList.add("show");
                else logBox.classList.remove("show");
            }
        }

        _setStatus(text, type) {
            var badge = this._shadowRoot.getElementById("statusBadge");
            badge.textContent = text;
            badge.className = "status-badge";
            if (type === "processing") badge.classList.add("status-processing");
            else if (type === "completed") badge.classList.add("status-completed");
            else if (type === "error") badge.classList.add("status-error");
            else if (type === "warning") badge.classList.add("status-warning");
            else badge.classList.add("status-ready");
        }

        _setProgress(percent, text) {
            this._shadowRoot.getElementById("progressWrap").classList.add("show");
            this._shadowRoot.getElementById("progressFill").style.width = percent + "%";
            this._shadowRoot.getElementById("progressText").textContent = text || "";
            this._shadowRoot.getElementById("progressPercent").textContent = percent + "%";
        }

        _hideProgress() {
            this._shadowRoot.getElementById("progressWrap").classList.remove("show");
        }

        _setSummary(rows, valid, invalid, sheet, columns, validationText) {
            this._shadowRoot.getElementById("summaryGrid").classList.add("show");
            this._shadowRoot.getElementById("sumRows").textContent = rows || 0;
            this._shadowRoot.getElementById("sumValid").textContent = valid || 0;
            this._shadowRoot.getElementById("sumInvalid").textContent = invalid || 0;
            this._shadowRoot.getElementById("sumSheet").textContent = sheet || "-";
            this._shadowRoot.getElementById("sumColumns").textContent = columns || 0;
            this._shadowRoot.getElementById("sumValidation").textContent = validationText || "-";
        }

        _showMessage(type, text) {
            var box = this._shadowRoot.getElementById("messageBox");
            box.className = "msg show " + type;
            box.textContent = text || "";
        }

        _hideMessage() {
            var box = this._shadowRoot.getElementById("messageBox");
            box.className = "msg";
            box.textContent = "";
        }

        _log(message, reset) {
            var box = this._shadowRoot.getElementById("logBox");
            if (reset) box.textContent = "";
            if (!this._export_settings.showlogs) return;
            box.classList.add("show");
            box.textContent += (box.textContent ? "\n" : "") + message;
        }

        _enableErrorDownload(enable) {
            this._shadowRoot.getElementById("downloadErrorBtn").disabled = !enable;
        }

        _loadExcelLibrary() {
            var that = this;
            loadScriptOnce(
                "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
                this._shadowRoot
            ).then(function () {
                that._setStatus("Ready", "ready");
                that._log("Excel library loaded successfully", true);
            }).catch(function () {
                that._setStatus("Error", "error");
                that._showMessage("error", "Failed to load Excel library");
                that._log("Failed to load Excel library", true);
            });
        }

        _downloadTemplate() {
            var that = this;
            var url = this._export_settings.templateurl;
            var fileName = this._export_settings.templatefilename || "Template.xlsm";

            var isSharePoint = /sharepoint\.com|sharepoint-df\.com/i.test(url);
            var isSacFileLink = /\/sap\/fpa\/ui\/app\.html#\/files/i.test(url);

            that._log("Downloading template...", false);
            that._setStatus("Downloading", "processing");

            if (isSharePoint || isSacFileLink) {
                try {
                    var aSp = document.createElement("a");
                    aSp.href = url;
                    aSp.target = "_blank";
                    aSp.rel = "noopener noreferrer";
                    document.body.appendChild(aSp);
                    aSp.click();
                    document.body.removeChild(aSp);

                    that._setStatus("Ready", "ready");
                    that._showMessage("info", "Template link opened. If file does not download, check access/permissions.");
                    that._log("Opened template link: " + url);
                } catch (errSp) {
                    that._setStatus("Error", "error");
                    that._showMessage("error", "Template open failed: " + errSp.message);
                    that._log("Template open failed: " + errSp.message, false);
                }
                return;
            }

            fetch(url)
                .then(function (response) {
                    if (!response.ok) {
                        throw new Error("HTTP " + response.status + " - Template file not found");
                    }
                    return response.blob();
                })
                .then(function (blob) {
                    var downloadUrl = URL.createObjectURL(blob);
                    var a = document.createElement("a");
                    a.href = downloadUrl;
                    a.download = fileName;
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(downloadUrl);

                    that._setStatus("Ready", "ready");
                    that._showMessage("success", "Template downloaded successfully: " + fileName);
                    that._log("Template downloaded successfully: " + fileName);
                })
                .catch(function (err) {
                    that._setStatus("Error", "error");
                    that._showMessage("error", "Template download failed: " + err.message);
                    that._log("Template download failed: " + err.message, false);
                });
        }

        _processUpload() {
            var that = this;
            var input = this._shadowRoot.getElementById("fileInput");
            var file = input.files && input.files[0] ? input.files[0] : null;

            this._hideMessage();

            if (!file) {
                this._setStatus("Warning", "warning");
                this._showMessage("warn", "Please select a file before upload");
                this._log("Please select a file before upload", true);
                return;
            }

            if (typeof XLSX === "undefined") {
                this._setStatus("Error", "error");
                this._showMessage("error", "Excel library is not loaded. Check internet/CDN access.");
                this._log("Excel library is not loaded. Check internet/CDN access.", true);
                return;
            }

            this._currentFileName = file.name;
            this._setStatus("Processing", "processing");
            this._setProgress(10, "Reading file...");
            this._errorLog = [];
            this._validData = [];
            this._previewRows = [];
            this._previewColumns = [];
            this._validationErrorsParsed = [];
            this._validationMap = {};
            this._uploadedHeaders = [];
            this._enableErrorDownload(false);
            this._log("File selected: " + file.name, true);

            var reader = new FileReader();

            reader.onload = function (e) {
                try {
                    that._setProgress(30, "Parsing workbook...");

                    var data = e.target.result;
                    var workbook = XLSX.read(data, { type: "binary" });

                    var actualSheet = workbook.SheetNames.indexOf("Sheet1") > -1 ? "Sheet1" : workbook.SheetNames[0];

                    if (!actualSheet) {
                        that._setStatus("Error", "error");
                        that._showMessage("error", "No sheet found in uploaded file");
                        that._log("No sheet found in uploaded file", true);
                        return;
                    }

                    that._sheetName = actualSheet;

                    that._setProgress(45, "Reading rows...");
                    var sheet = workbook.Sheets[actualSheet];
                    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

                    if (!rows || rows.length <= 1) {
                        that._setStatus("Error", "error");
                        that._showMessage("error", "There is no record to be uploaded");
                        that._log("There is no record to be uploaded", true);
                        return;
                    }

                    var header = rows[0] || [];
                    that._uploadedHeaders = [];
                    var h = 0;
                    for (h = 0; h < header.length; h++) {
                        that._uploadedHeaders.push(String(header[h] || "").trim());
                    }

                    that._setProgress(60, "Validating template...");
                    var prepare = that._prepareColumnMapFromHeaders();

                    if (!prepare.ok) {
                        that._setStatus("Error", "error");
                        that._showMessage("error", prepare.message);
                        that._log(prepare.message, true);
                        return;
                    }

                    that._setProgress(72, "Scanning duplicates...");
                    var requiredColumns = that._getRequiredColumns();
                    var colMap = prepare.colMap;
                    var keyColumn = that._getKeyColumn(requiredColumns);
                    var idCount = {};

                    var r = 0;
                    for (r = 1; r < rows.length; r++) {
                        var row = rows[r] || [];
                        var id = String(row[colMap[keyColumn]] || "").trim();
                        if (!id) continue;
                        idCount[id] = (idCount[id] || 0) + 1;
                    }

                    that._setProgress(82, "Validating rows...");
                    that._buildDataFromRows(rows, requiredColumns, colMap, idCount);

                    if (that._validData.length > that._export_settings.maxrows) {
                        that._setStatus("Error", "error");
                        that._showMessage("error", "Maximum valid records are " + that._export_settings.maxrows);
                        that._log("Maximum valid records are " + that._export_settings.maxrows, true);
                        return;
                    }

                    var allRowsForSac = [];
                    var pr = 0;
                    var pc = 0;

                    for (pr = 0; pr < that._previewRows.length; pr++) {
                        var rowObjForSac = {};
                        for (pc = 0; pc < that._previewColumns.length; pc++) {
                            var colNameForSac = that._previewColumns[pc];
                            if (that._previewRows[pr][colNameForSac] !== undefined && that._previewRows[pr][colNameForSac] !== null) {
                                rowObjForSac[colNameForSac] = String(that._previewRows[pr][colNameForSac]);
                            } else {
                                rowObjForSac[colNameForSac] = "";
                            }
                        }
                        allRowsForSac.push(rowObjForSac);
                    }

                    that.unit = JSON.stringify(allRowsForSac);
                    that._export_settings.rowcount = that._previewRows.length;
                    that._export_settings.validcount = that._validData.length;
                    that._export_settings.invalidcount = that._errorLog.length;

                    that._buildValidationMap();
                    that._renderPreview();
                    that._enableErrorDownload(that._errorLog.length > 0);

                    that._firePropertiesChanged("uploadCompleted");

                    that.dispatchEvent(new CustomEvent("onStart", {
                        detail: {
                            settings: {},
                            rowCount: that._validData.length,
                            invalidCount: that._errorLog.length,
                            fileName: file.name,
                            sheetName: actualSheet
                        }
                    }));

                    that._setSummary(
                        that._previewRows.length,
                        that._validData.length,
                        that._errorLog.length,
                        actualSheet,
                        requiredColumns.length,
                        that._errorLog.length > 0 ? "Invalid" : "Valid"
                    );

                    that._setProgress(100, "Completed");

                    if (that._errorLog.length > 0) {
                        that._setStatus("Completed with Errors", "warning");
                        that._hideMessage();
                    } else {
                        that._setStatus("Completed", "completed");
                        that._hideMessage();
                    }

                    that._log("Valid rows: " + that._validData.length);
                    that._log("Invalid rows: " + that._errorLog.length);
                    that._log("Detected headers: " + that._uploadedHeaders.join(", "));

                    var dupIds = Object.keys(idCount).filter(function (id2) {
                        return idCount[id2] > 1;
                    });

                    if (dupIds.length > 0) {
                        that._log("Duplicate key values rejected: " + dupIds.join(", "));
                    }

                } catch (err) {
                    that._setStatus("Error", "error");
                    that._showMessage("error", "Processing failed: " + err.message);
                    that._log("Processing failed: " + err.message, true);
                }
            };

            reader.readAsBinaryString(file);
        }

        _prepareColumnMapFromHeaders() {
            var requiredColumns = this._getRequiredColumns();
            var colMap = {};
            var missing = [];
            var i = 0;
            var j = 0;

            for (i = 0; i < requiredColumns.length; i++) {
                var req = requiredColumns[i];
                colMap[req] = -1;

                for (j = 0; j < this._uploadedHeaders.length; j++) {
                    if (String(this._uploadedHeaders[j]).trim().toUpperCase() === String(req).trim().toUpperCase()) {
                        colMap[req] = j;
                        break;
                    }
                }

                if (colMap[req] === -1) {
                    missing.push(req);
                }
            }

            if (missing.length > 0) {
                return {
                    ok: false,
                    message: "Invalid template. Missing required columns: " + missing.join(", ")
                };
            }

            return {
                ok: true,
                colMap: colMap
            };
        }

        _buildDataFromRows(rows, requiredColumns, colMap, idCount) {
            var validRows = [];
            var errorRows = [];
            var previewRows = [];
            var keyColumn = this._getKeyColumn(requiredColumns);
            var r = 0;
            var rc = 0;

            for (r = 1; r < rows.length; r++) {
                var row2 = rows[r] || [];
                var rowNumber = r + 1;
                var rowObj = {};
                var previewRowObj = { _rowNumber: rowNumber };
                var isBlankRow = true;
                var errors = [];
                var cellErrors = [];

                for (rc = 0; rc < requiredColumns.length; rc++) {
                    var colName = requiredColumns[rc];
                    var cellValue = String(row2[colMap[colName]] || "").trim();
                    rowObj[colName] = cellValue;
                    previewRowObj[colName] = cellValue;

                    if (cellValue !== "") {
                        isBlankRow = false;
                    }
                }

                if (isBlankRow) {
                    continue;
                }

                previewRows.push(previewRowObj);

                for (rc = 0; rc < requiredColumns.length; rc++) {
                    var reqCol = requiredColumns[rc];
                    if (!rowObj[reqCol]) {
                        errors.push(reqCol + " is mandatory");
                        cellErrors.push({
                            rowIndex: previewRows.length - 1,
                            field: reqCol,
                            message: "Blank cell not allowed"
                        });
                    }
                }

                if (rowObj[keyColumn] && idCount[rowObj[keyColumn]] > 1) {
                    errors.push("Duplicate " + keyColumn + " '" + rowObj[keyColumn] + "' - all " + idCount[rowObj[keyColumn]] + " occurrences rejected");
                    cellErrors.push({
                        rowIndex: previewRows.length - 1,
                        field: keyColumn,
                        message: "Duplicate value not allowed"
                    });
                }

                if (errors.length > 0) {
                    errorRows.push({
                        RowNumber: rowNumber,
                        ErrorMessage: errors.join(" | ")
                    });

                    for (rc = 0; rc < cellErrors.length; rc++) {
                        this._validationErrorsParsed.push(cellErrors[rc]);
                    }
                } else {
                    validRows.push(rowObj);
                }
            }

            this._validData = validRows;
            this._errorLog = errorRows;
            this._previewRows = previewRows;
            this._previewColumns = requiredColumns.slice(0);
        }

        _getRequiredColumns() {
            var raw = String(this._export_settings.requiredcolumns || "");
            var parts = raw.split(",");
            var out = [];
            var i = 0;

            for (i = 0; i < parts.length; i++) {
                var v = String(parts[i] || "").trim();
                if (v !== "") {
                    out.push(v);
                }
            }

            return out.length ? out : ["ID", "DESCRIPTION", "H1", "COMPANY", "COSTCENTER", "ASSET_CLASS", "CAPITALIZED"];
        }

        _getKeyColumn(requiredColumns) {
            var key = String(this._export_settings.keycolumn || "").trim();
            if (key !== "") {
                return key;
            }
            if (requiredColumns.indexOf("ID") > -1) {
                return "ID";
            }
            return requiredColumns[0];
        }

        _applySacValidation() {
            var parsed = this._safeParseArray(this._export_settings.validationerrors);

            this._validationErrorsParsed = parsed;
            this._buildValidationMap();
            this._renderPreview();

            var invalidRowMap = {};
            var i = 0;

            for (i = 0; i < this._validationErrorsParsed.length; i++) {
                var err = this._validationErrorsParsed[i];
                if (err.rowIndex !== undefined && err.rowIndex !== null) {
                    invalidRowMap[String(err.rowIndex)] = true;
                }
            }

            var invalidRowsFinal = 0;
            for (var k in invalidRowMap) {
                if (Object.prototype.hasOwnProperty.call(invalidRowMap, k)) {
                    invalidRowsFinal = invalidRowsFinal + 1;
                }
            }

            var totalRowsFinal = this._previewRows.length;
            var validRowsFinal = totalRowsFinal - invalidRowsFinal;

            if (validRowsFinal < 0) {
                validRowsFinal = 0;
            }

            this._setSummary(
                totalRowsFinal,
                validRowsFinal,
                invalidRowsFinal,
                this._sheetName,
                this._previewColumns.length,
                this._export_settings.validationresult === "false" ? "Invalid" : "Valid"
            );

            if (this._export_settings.validationresult === "false") {
                this._hideMessage();
                this._setStatus("Validation Error", "error");
            } else {
                this._hideMessage();
                if (this._previewRows.length > 0) {
                    this._setStatus("Completed", "completed");
                } else {
                    this._setStatus("Ready", "ready");
                }
            }
        }

        _buildValidationSummaryText() {
            if (!this._validationErrorsParsed || this._validationErrorsParsed.length === 0) {
                return "Validation failed.";
            }

            var lines = [];
            var i = 0;

            for (i = 0; i < this._validationErrorsParsed.length; i++) {
                var err = this._validationErrorsParsed[i];
                var rowText = "Row " + (err.rowIndex !== undefined ? (parseInt(err.rowIndex, 10) + 1) : "?");
                var fieldText = err.field ? " / " + err.field : "";
                var msgText = err.message || "Invalid value";
                lines.push(rowText + fieldText + ": " + msgText);
            }

            return lines.join("\n");
        }

        _buildValidationMap() {
            this._validationMap = {};
            var i = 0;

            for (i = 0; i < this._validationErrorsParsed.length; i++) {
                var err = this._validationErrorsParsed[i];
                var rowIndex = err.rowIndex !== undefined ? String(err.rowIndex) : "";
                var field = err.field || "";
                var key = rowIndex + "|" + field;

                if (!this._validationMap[key]) this._validationMap[key] = [];
                this._validationMap[key].push(err.message || "Invalid value");

                var rowKey = rowIndex + "|__row__";
                if (!this._validationMap[rowKey]) this._validationMap[rowKey] = [];
                this._validationMap[rowKey].push(err.message || "Invalid value");
            }
        }

        _renderPreview() {
            var wrap = this._shadowRoot.getElementById("previewWrap");
            var head = this._shadowRoot.getElementById("previewHead");
            var body = this._shadowRoot.getElementById("previewBody");
            var meta = this._shadowRoot.getElementById("previewMeta");

            if (!this._export_settings.showpreview || !this._previewVisible) {
                wrap.classList.remove("show");
                return;
            }

            wrap.classList.add("show");
            head.innerHTML = "";
            body.innerHTML = "";

            if (!this._previewRows || this._previewRows.length === 0) {
                meta.textContent = "0 row(s) shown";
                body.innerHTML = '<tr><td colspan="50">No preview available</td></tr>';
                return;
            }

            var cols = this._previewColumns && this._previewColumns.length ? this._previewColumns.slice(0) : [];
            if (cols.length === 0) {
                var sample = this._previewRows[0];
                for (var k in sample) {
                    if (k !== "_rowNumber") cols.push(k);
                }
            }

            var filteredRows = [];
            var sText = String(this._searchText || "").toLowerCase().trim();
            var i = 0;
            var j = 0;

            for (i = 0; i < this._previewRows.length; i++) {
                var row = this._previewRows[i];
                if (sText === "") {
                    filteredRows.push(row);
                } else {
                    var matched = false;
                    for (j = 0; j < cols.length; j++) {
                        var val = String(row[cols[j]] || "").toLowerCase();
                        if (val.indexOf(sText) > -1) {
                            matched = true;
                            break;
                        }
                    }
                    if (matched) filteredRows.push(row);
                }
            }

            var limit = this._export_settings.previewrows;
            var totalMatched = filteredRows.length;
            if (filteredRows.length > limit) {
                filteredRows = filteredRows.slice(0, limit);
            }

            meta.textContent = filteredRows.length + " row(s) shown" + (totalMatched > limit ? " of " + totalMatched : "");

            var headHtml = "<tr><th>#</th>";
            for (j = 0; j < cols.length; j++) {
                headHtml += "<th>" + this._escapeHtml(cols[j]) + "</th>";
            }
            headHtml += "</tr>";
            head.innerHTML = headHtml;

            var bodyHtml = "";
            for (i = 0; i < filteredRows.length; i++) {
                var rowObj = filteredRows[i];
                var originalIndex = this._getOriginalPreviewIndex(rowObj);
                var rowKey = String(originalIndex) + "|__row__";
                var rowClass = this._validationMap[rowKey] ? "invalid-row" : "";

                bodyHtml += '<tr class="' + rowClass + '">';
                bodyHtml += "<td>" + this._escapeHtml(String(rowObj._rowNumber || (originalIndex + 2))) + "</td>";

                for (j = 0; j < cols.length; j++) {
                    var field = cols[j];
                    var cellKey = String(originalIndex) + "|" + field;
                    var invalidClass = this._validationMap[cellKey] ? "invalid-cell" : "";
                    var cellVal = rowObj[field] === undefined || rowObj[field] === null ? "" : String(rowObj[field]);

                    bodyHtml += '<td class="' + invalidClass + '">';
                    bodyHtml += this._escapeHtml(cellVal);

                    if (this._validationMap[cellKey]) {
                        for (var m = 0; m < this._validationMap[cellKey].length; m++) {
                            bodyHtml += '<span class="cell-error">' + this._escapeHtml(this._validationMap[cellKey][m]) + '</span>';
                        }
                    }

                    bodyHtml += "</td>";
                }

                bodyHtml += "</tr>";
            }

            body.innerHTML = bodyHtml;
        }

        _getOriginalPreviewIndex(rowObj) {
            var i = 0;
            for (i = 0; i < this._previewRows.length; i++) {
                if (this._previewRows[i] === rowObj) return i;
                if (this._previewRows[i]._rowNumber === rowObj._rowNumber) return i;
            }
            return 0;
        }

        clear() {
            this._shadowRoot.getElementById("fileInput").value = "";
            this._shadowRoot.getElementById("searchInput").value = "";
            this._searchText = "";
            this.unit = "";
            this._errorLog = [];
            this._validData = [];
            this._previewRows = [];
            this._previewColumns = [];
            this._validationErrorsParsed = [];
            this._validationMap = {};
            this._uploadedHeaders = [];
            this._export_settings.validationresult = "true";
            this._export_settings.validationerrors = "[]";
            this._sheetName = "-";
            this._currentFileName = "";
            this._enableErrorDownload(false);
            this._setStatus("Ready", "ready");
            this._hideProgress();
            this._hideMessage();
            this._setSummary(0, 0, 0, "-", 0, "-");
            this._renderPreview();
            this._log("Cleared previous file and output", true);
            this._export_settings.lastevent = "clear";
            this._firePropertiesChanged("clear");

            this.dispatchEvent(new CustomEvent("onClear", { detail: {} }));
        }

        _downloadErrorLog() {
            if (!this._errorLog || this._errorLog.length === 0) {
                this._showMessage("info", "No error log available to download");
                this._log("No error log available to download");
                return;
            }

            var rows = ["RowNumber,ErrorMessage"];
            var i = 0;

            for (i = 0; i < this._errorLog.length; i++) {
                var item = this._errorLog[i];
                rows.push([
                    this._escapeCsv(item.RowNumber),
                    this._escapeCsv(item.ErrorMessage)
                ].join(","));
            }

            this._downloadBlob(rows.join("\n"), "text/csv;charset=utf-8;", this._export_settings.errorlogfilename);
            this._showMessage("success", "Error log downloaded successfully");
            this._log("Error log downloaded successfully");
        }

        _safeParseArray(text) {
            try {
                var parsed = JSON.parse(text || "[]");
                return Array.isArray(parsed) ? parsed : [];
            } catch (e) {
                return [];
            }
        }

        _downloadBlob(content, mimeType, fileName) {
            var blob = new Blob([content], { type: mimeType });
            var url = URL.createObjectURL(blob);
            var a = document.createElement("a");
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        _escapeCsv(value) {
            var str = value == null ? "" : String(value);
            if (str.indexOf(",") > -1 || str.indexOf('"') > -1 || str.indexOf("\n") > -1) {
                str = '"' + str.replace(/"/g, '""') + '"';
            }
            return str;
        }

        _firePropertiesChanged(eventName) {
            this._export_settings.lastevent = eventName || "";

            this.dispatchEvent(new CustomEvent("propertiesChanged", {
                detail: {
                    properties: {
                        unit: this.unit,
                        validationresult: this._export_settings.validationresult,
                        validationerrors: this._export_settings.validationerrors,
                        lastevent: this._export_settings.lastevent,
                        invalidcount: this._export_settings.invalidcount,
                        validcount: this._export_settings.validcount,
                        rowcount: this._export_settings.rowcount
                    }
                }
            }));
        }

        _toBoolean(v, defaultValue) {
            if (v === true || v === "true") return true;
            if (v === false || v === "false") return false;
            return defaultValue;
        }

        _escapeHtml(str) {
            return String(str)
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;");
        }
    }

    customElements.define("com-fd-djaja-sap-sac-excelcom", Excel);

    function loadScriptOnce(src, shadowRoot) {
        return new Promise(function (resolve, reject) {
            if (typeof XLSX !== "undefined") {
                resolve();
                return;
            }

            var existing = shadowRoot.querySelector('script[src="' + src + '"]');
            if (existing) {
                existing.addEventListener("load", resolve);
                existing.addEventListener("error", reject);
                return;
            }

            var script = document.createElement("script");
            script.src = src;
            script.onload = resolve;
            script.onerror = reject;
            shadowRoot.appendChild(script);
        });
    }
})();
*/

