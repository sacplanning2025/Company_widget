/*
(function () {
    "use strict";

    var tmpl = document.createElement("template");
    tmpl.innerHTML = `
        <style>
            :host{
                display:block;
                width:100%;
                font-family:"72", Arial, Helvetica, sans-serif;
                color:#1f2d3d;
            }
            .excel-widget{
                width:100%;
                box-sizing:border-box;
            }
            .upload-card{
                border:1px solid #d9d9d9;
                border-radius:10px;
                background:#ffffff;
                box-shadow:0 2px 8px rgba(0,0,0,0.08);
                padding:14px;
            }
            .upload-header{
                display:flex;
                align-items:center;
                justify-content:space-between;
                margin-bottom:12px;
                gap:12px;
                flex-wrap:wrap;
            }
            .upload-title{
                font-size:15px;
                font-weight:700;
                color:#0a6ed1;
            }
            .status-badge{
                font-size:11px;
                font-weight:600;
                padding:5px 10px;
                border-radius:14px;
                background:#f5f6f7;
                color:#354a5f;
                border:1px solid #d9d9d9;
            }
            .toolbar-actions{
                display:flex;
                gap:8px;
                flex-wrap:wrap;
                margin-bottom:12px;
            }
            .toolbar-btn{
                appearance:none;
                border:1px solid #c7ced4;
                background:#ffffff;
                color:#1f2d3d;
                border-radius:8px;
                padding:8px 12px;
                font-size:12px;
                font-weight:600;
                cursor:pointer;
                transition:all 0.2s ease;
            }
            .toolbar-btn:hover{
                border-color:#0a6ed1;
                color:#0a6ed1;
            }
            .toolbar-btn.primary{
                background:#0a6ed1;
                color:#ffffff;
                border-color:#0a6ed1;
            }
            .toolbar-btn.primary:hover{
                background:#085caf;
                color:#ffffff;
                border-color:#085caf;
            }
            .toolbar-btn:disabled{
                opacity:0.55;
                cursor:not-allowed;
            }
            .upload-area{
                border:1px dashed #b8c4d1;
                border-radius:10px;
                padding:14px;
                background:#fafcff;
            }
            .upload-row{
                display:flex;
                gap:10px;
                align-items:center;
                flex-wrap:wrap;
            }
            .file-input{
                font-size:12px;
                padding:8px;
                border:1px solid #d9d9d9;
                border-radius:8px;
                background:#fff;
            }
            .progress-wrap{
                display:none;
                margin-top:12px;
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
                height:10px;
                background:#edf2f7;
                border-radius:10px;
                overflow:hidden;
                border:1px solid #d9d9d9;
            }
            .progress-fill{
                width:0%;
                height:100%;
                background:linear-gradient(90deg, #0a6ed1, #4db1ff);
                transition:width 0.25s ease;
            }
            .summary-grid{
                display:none;
                grid-template-columns:repeat(4, minmax(0, 1fr));
                gap:8px;
                margin-top:12px;
            }
            .summary-grid.show{
                display:grid;
            }
            .summary-item{
                border:1px solid #e5e7eb;
                border-radius:8px;
                padding:10px;
                background:#fafbfc;
            }
            .summary-item .k{
                font-size:11px;
                color:#6a6d70;
                margin-bottom:4px;
            }
            .summary-item .v{
                font-size:15px;
                font-weight:700;
                color:#1f2d3d;
            }
            .log-box{
                display:none;
                margin-top:12px;
                border:1px solid #e5e7eb;
                border-radius:8px;
                background:#fafbfc;
                padding:10px;
                max-height:160px;
                overflow:auto;
                font-size:12px;
                line-height:1.45;
                color:#354a5f;
                white-space:pre-wrap;
            }
            .log-box.show{
                display:block;
            }
            .footer-note{
                margin-top:10px;
                font-size:11px;
                color:#6a6d70;
            }
        </style>

        <div class="excel-widget">
            <div class="upload-card">
                <div class="upload-header">
                    <div>
                        <div class="upload-title" id="titleEl">Excel Upload</div>
                    </div>
                    <div class="status-badge" id="statusBadge">Ready</div>
                </div>

                <div class="toolbar-actions">
                    <button type="button" class="toolbar-btn primary" id="downloadTemplateBtn">Download Template</button>
                    <button type="button" class="toolbar-btn" id="downloadErrorBtn" disabled>Download Error Log</button>
                </div>

                <div class="upload-area">
                    <div class="upload-row">
                        <input type="file" id="fileInput" class="file-input" accept=".xls,.xlsx,.xlsm,.csv" />
                        <button type="button" class="toolbar-btn primary" id="uploadBtn">Upload</button>
                        <button type="button" class="toolbar-btn" id="clearBtn">Clear</button>
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

                <div class="summary-grid" id="summaryGrid">
                    <div class="summary-item">
                        <div class="k">Rows Read</div>
                        <div class="v" id="sumRows">0</div>
                    </div>
                    <div class="summary-item">
                        <div class="k">Valid Rows</div>
                        <div class="v" id="sumValid">0</div>
                    </div>
                    <div class="summary-item">
                        <div class="k">Invalid Rows</div>
                        <div class="v" id="sumInvalid">0</div>
                    </div>
                    <div class="summary-item">
                        <div class="k">Sheet</div>
                        <div class="v" id="sumSheet">-</div>
                    </div>
                </div>

                <div class="log-box" id="logBox"></div>
                <div class="footer-note" id="footerNote">Supported template: Sheet1 with columns ID, DESCRIPTION, H1, costcenter</div>
            </div>
        </div>
    `;

    var _shadowRoot;
    var _result = "";
    var TEMPLATE_URL = "https://raw.githubusercontent.com/sacplanning2025/Company_widget/main/New_Position_Creation_V2.1.xlsm";

    class Excel extends HTMLElement {
        constructor() {
            super();
            _shadowRoot = this.attachShadow({ mode: "open" });
            _shadowRoot.appendChild(tmpl.content.cloneNode(true));

            this._export_settings = {
                title: "",
                subtitle: "",
                icon: "",
                unit: "",
                footer: "",
                errorlogfilename: "Excel_Upload_Error_Log.csv",
                templatefilename: "New_Position_Creation_V2.1.xlsm",
                templateurl: TEMPLATE_URL
            };

            this._errorLog = [];
            this._validData = [];
            this._designMode = false;

            this._bindEvents();
        }

        connectedCallback() {
            this._setStatus("Ready");
            this._applyHeaderSettings();
            this._loadExcelLibrary();
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

            this._applyHeaderSettings();
        }

        static get observedAttributes() {
            return ["title", "subtitle", "icon", "unit", "footer", "errorlogfilename", "templatefilename", "templateurl"];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue !== newValue) {
                this[name] = newValue;
            }
        }

        get title() {
            return this._export_settings.title;
        }
        set title(v) {
            this._export_settings.title = v || "";
            this._applyHeaderSettings();
        }

        get subtitle() {
            return this._export_settings.subtitle;
        }
        set subtitle(v) {
            this._export_settings.subtitle = v || "";
        }

        get icon() {
            return this._export_settings.icon;
        }
        set icon(v) {
            this._export_settings.icon = v || "";
        }

        get unit() {
            return this._export_settings.unit;
        }
        set unit(v) {
            this._export_settings.unit = _result || v || "";
        }

        get footer() {
            return this._export_settings.footer;
        }
        set footer(v) {
            this._export_settings.footer = v || "";
            this._applyHeaderSettings();
        }

        get errorlogfilename() {
            return this._export_settings.errorlogfilename;
        }
        set errorlogfilename(v) {
            this._export_settings.errorlogfilename = v || "Excel_Upload_Error_Log.csv";
        }

        get templatefilename() {
            return this._export_settings.templatefilename;
        }
        set templatefilename(v) {
            this._export_settings.templatefilename = v || "New_Position_Creation_V2.1.xlsm";
        }

        get templateurl() {
            return this._export_settings.templateurl;
        }
        set templateurl(v) {
            this._export_settings.templateurl = v || TEMPLATE_URL;
        }

        _bindEvents() {
            var that = this;
            setTimeout(function () {
                _shadowRoot.getElementById("downloadTemplateBtn").addEventListener("click", function () {
                    that._downloadTemplate();
                });
                _shadowRoot.getElementById("downloadErrorBtn").addEventListener("click", function () {
                    that._downloadErrorLog();
                });
                _shadowRoot.getElementById("uploadBtn").addEventListener("click", function () {
                    that._processUpload();
                });
                _shadowRoot.getElementById("clearBtn").addEventListener("click", function () {
                    that._clearAll();
                });
            }, 0);
        }

        _applyHeaderSettings() {
            var titleEl = _shadowRoot.getElementById("titleEl");
            var footerEl = _shadowRoot.getElementById("footerNote");
            if (titleEl) {
                titleEl.textContent = this._export_settings.title || "Excel Upload";
            }
            if (footerEl) {
                footerEl.textContent = this._export_settings.footer || "Supported template: Sheet1 with columns ID, DESCRIPTION, H1, costcenter";
            }
        }

        _setStatus(text) {
            _shadowRoot.getElementById("statusBadge").textContent = text;
        }

        _setProgress(percent, text) {
            _shadowRoot.getElementById("progressWrap").classList.add("show");
            _shadowRoot.getElementById("progressFill").style.width = percent + "%";
            _shadowRoot.getElementById("progressText").textContent = text || "";
            _shadowRoot.getElementById("progressPercent").textContent = percent + "%";
        }

        _hideProgress() {
            _shadowRoot.getElementById("progressWrap").classList.remove("show");
        }

        _setSummary(rows, valid, invalid, sheet) {
            _shadowRoot.getElementById("summaryGrid").classList.add("show");
            _shadowRoot.getElementById("sumRows").textContent = rows || 0;
            _shadowRoot.getElementById("sumValid").textContent = valid || 0;
            _shadowRoot.getElementById("sumInvalid").textContent = invalid || 0;
            _shadowRoot.getElementById("sumSheet").textContent = sheet || "-";
        }

        _log(message, reset) {
            var box = _shadowRoot.getElementById("logBox");
            if (reset) {
                box.textContent = "";
            }
            box.classList.add("show");
            box.textContent += (box.textContent ? "\n" : "") + message;
        }

        _enableErrorDownload(enable) {
            _shadowRoot.getElementById("downloadErrorBtn").disabled = !enable;
        }

        _loadExcelLibrary() {
            var that = this;
            loadScriptOnce(
                "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
                _shadowRoot
            ).then(function () {
                that._setStatus("Ready");
                that._log("Excel library loaded successfully", true);
            }).catch(function () {
                that._setStatus("Error");
                that._log("Failed to load Excel library", true);
            });
        }

        _downloadTemplate() {
            var that = this;
            var url = this._export_settings.templateurl;
            var fileName = this._export_settings.templatefilename || "New_Position_Creation_V2.1.xlsm";

            that._log("Downloading template...", false);
            that._setStatus("Downloading");

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

                    that._setStatus("Ready");
                    that._log("Template downloaded successfully: " + fileName);
                })
                .catch(function (err) {
                    that._setStatus("Error");
                    that._log("Template download failed: " + err.message, false);
                });
        }

        _processUpload() {
            var that = this;
            var input = _shadowRoot.getElementById("fileInput");
            var file = input.files && input.files[0] ? input.files[0] : null;

            if (!file) {
                this._setStatus("Warning");
                this._log("Please select a file before upload", true);
                return;
            }

            if (typeof XLSX === "undefined") {
                this._setStatus("Error");
                this._log("Excel library is not loaded. Check internet/CDN access.", true);
                return;
            }

            this._setStatus("Processing");
            this._setProgress(10, "Reading file...");
            this._errorLog = [];
            this._validData = [];
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
                        that._setStatus("Error");
                        that._log("No sheet found in uploaded file", true);
                        return;
                    }

                    that._setProgress(50, "Reading rows...");
                    var sheet = workbook.Sheets[actualSheet];
                    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

                    if (!rows || rows.length <= 1) {
                        that._setStatus("Error");
                        that._log("There is no record to be uploaded", true);
                        return;
                    }

                    var header = rows[0];
                    var colMap = {
                        ID: -1,
                        DESCRIPTION: -1,
                        H1: -1,
                        costcenter: -1
                    };

                    for (var i = 0; i < header.length; i++) {
                        var h = String(header[i] || "").trim();
                        if (h === "ID") {
                            colMap.ID = i;
                        } else if (h === "DESCRIPTION") {
                            colMap.DESCRIPTION = i;
                        } else if (h === "H1") {
                            colMap.H1 = i;
                        } else if (h.toLowerCase() === "costcenter") {
                            colMap.costcenter = i;
                        }
                    }

                    if (colMap.ID === -1 || colMap.DESCRIPTION === -1 || colMap.H1 === -1 || colMap.costcenter === -1) {
                        that._setStatus("Error");
                        that._log("Invalid template. Required columns: ID, DESCRIPTION, H1, costcenter", true);
                        return;
                    }

                    that._setProgress(65, "Scanning for duplicates...");

                    var idCount = {};
                    for (var r = 1; r < rows.length; r++) {
                        var row = rows[r] || [];
                        var id = String(row[colMap.ID] || "").trim();
                        if (!id) {
                            continue;
                        }
                        idCount[id] = (idCount[id] || 0) + 1;
                    }

                    that._setProgress(80, "Validating rows...");

                    var validRows = [];
                    var errorRows = [];

                    for (var r2 = 1; r2 < rows.length; r2++) {
                        var row2 = rows[r2] || [];
                        var rowNumber = r2 + 1;

                        var rowObj = {
                            ID: String(row2[colMap.ID] || "").trim(),
                            DESCRIPTION: String(row2[colMap.DESCRIPTION] || "").trim(),
                            H1: String(row2[colMap.H1] || "").trim(),
                            costcenter: String(row2[colMap.costcenter] || "").trim()
                        };

                        if (!rowObj.ID && !rowObj.DESCRIPTION && !rowObj.H1 && !rowObj.costcenter) {
                            continue;
                        }

                        var errors = [];

                        if (!rowObj.ID) {
                            errors.push("ID is mandatory");
                        }
                        if (!rowObj.DESCRIPTION) {
                            errors.push("DESCRIPTION is mandatory");
                        }
                        if (!rowObj.H1) {
                            errors.push("H1 is mandatory");
                        }
                        if (!rowObj.costcenter) {
                            errors.push("costcenter is mandatory");
                        }

                        if (rowObj.ID && idCount[rowObj.ID] > 1) {
                            errors.push("Duplicate ID '" + rowObj.ID + "' - all " + idCount[rowObj.ID] + " occurrences rejected");
                        }

                        if (errors.length > 0) {
                            errorRows.push({
                                RowNumber: rowNumber,
                                ID: rowObj.ID,
                                DESCRIPTION: rowObj.DESCRIPTION,
                                H1: rowObj.H1,
                                costcenter: rowObj.costcenter,
                                ErrorMessage: errors.join(" | ")
                            });
                        } else {
                            validRows.push(rowObj);
                        }
                    }

                    if (validRows.length > 2000) {
                        that._setStatus("Error");
                        that._log("Maximum records are 2000", true);
                        return;
                    }

                    that._validData = validRows;
                    that._errorLog = errorRows;
                    that._enableErrorDownload(errorRows.length > 0);

                    _result = JSON.stringify(validRows);
                    that._firePropertiesChanged();

                    that.dispatchEvent(new CustomEvent("onStart", {
                        detail: {
                            settings: {},
                            rowCount: validRows.length,
                            invalidCount: errorRows.length,
                            fileName: file.name,
                            sheetName: actualSheet
                        }
                    }));

                    that._setSummary(rows.length - 1, validRows.length, errorRows.length, actualSheet);
                    that._setProgress(100, "Completed");
                    that._setStatus("Completed");
                    that._log("Valid rows: " + validRows.length);
                    that._log("Invalid rows: " + errorRows.length);

                    var dupIds = Object.keys(idCount).filter(function (id) {
                        return idCount[id] > 1;
                    });

                    if (dupIds.length > 0) {
                        that._log("Duplicate IDs rejected (all occurrences): " + dupIds.join(", "));
                    }

                } catch (err) {
                    that._setStatus("Error");
                    that._log("Processing failed: " + err.message, true);
                }
            };

            reader.readAsBinaryString(file);
        }

        _clearAll() {
            _shadowRoot.getElementById("fileInput").value = "";
            _result = "";
            this._errorLog = [];
            this._validData = [];
            this._enableErrorDownload(false);
            this._setStatus("Ready");
            this._hideProgress();
            this._setSummary(0, 0, 0, "-");
            this._log("Cleared previous file and output", true);
            this._firePropertiesChanged();
        }

        _downloadErrorLog() {
            if (!this._errorLog || this._errorLog.length === 0) {
                this._log("No error log available to download");
                return;
            }

            var rows = ["RowNumber,ID,DESCRIPTION,H1,costcenter,ErrorMessage"];
            for (var i = 0; i < this._errorLog.length; i++) {
                var item = this._errorLog[i];
                rows.push([
                    this._escapeCsv(item.RowNumber),
                    this._escapeCsv(item.ID),
                    this._escapeCsv(item.DESCRIPTION),
                    this._escapeCsv(item.H1),
                    this._escapeCsv(item.costcenter),
                    this._escapeCsv(item.ErrorMessage)
                ].join(","));
            }

            this._downloadBlob(rows.join("\n"), "text/csv;charset=utf-8;", this._export_settings.errorlogfilename);
            this._log("Error log downloaded successfully");
        }

        _escapeCsv(value) {
            var str = value == null ? "" : String(value);
            if (str.indexOf(",") > -1 || str.indexOf('"') > -1 || str.indexOf("\n") > -1) {
                str = '"' + str.replace(/"/g, '""') + '"';
            }
            return str;
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

        _firePropertiesChanged() {
            this.unit = _result;
            this.dispatchEvent(new CustomEvent("propertiesChanged", {
                detail: {
                    properties: {
                        unit: this.unit
                    }
                }
            }));
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
(function () {
    "use strict";

    var tmpl = document.createElement("template");
    tmpl.innerHTML = `
        <style>
            :host{
                display:block;
                width:100%;
                font-family:"72", Arial, Helvetica, sans-serif;
                color:#1f2d3d;
            }
            .excel-widget{
                width:100%;
                box-sizing:border-box;
            }
            .upload-card{
                border:1px solid #d9d9d9;
                border-radius:14px;
                background:linear-gradient(180deg, #ffffff 0%, #fbfdff 100%);
                box-shadow:0 8px 24px rgba(0,0,0,0.08);
                padding:16px;
            }
            .upload-header{
                display:flex;
                align-items:flex-start;
                justify-content:space-between;
                margin-bottom:14px;
                gap:12px;
                flex-wrap:wrap;
            }
            .title-wrap{
                display:flex;
                flex-direction:column;
                gap:4px;
            }
            .upload-title{
                font-size:18px;
                font-weight:700;
                color:#0a6ed1;
                line-height:1.2;
            }
            .upload-subtitle{
                font-size:12px;
                color:#6a6d70;
            }
            .status-badge{
                font-size:11px;
                font-weight:700;
                padding:6px 10px;
                border-radius:14px;
                background:#f5f6f7;
                color:#354a5f;
                border:1px solid #d9d9d9;
                white-space:nowrap;
            }
            .status-ready{ background:#f5f6f7; color:#354a5f; border-color:#d9d9d9; }
            .status-processing{ background:#fff7e6; color:#8a5a00; border-color:#ffd591; }
            .status-completed{ background:#f6ffed; color:#237804; border-color:#b7eb8f; }
            .status-error{ background:#fff1f0; color:#a8071a; border-color:#ffa39e; }
            .status-warning{ background:#fffbe6; color:#ad6800; border-color:#ffe58f; }

            .toolbar-actions{
                display:flex;
                gap:8px;
                flex-wrap:wrap;
                margin-bottom:14px;
            }
            .toolbar-btn{
                appearance:none;
                border:1px solid #c7ced4;
                background:#ffffff;
                color:#1f2d3d;
                border-radius:9px;
                padding:8px 12px;
                font-size:12px;
                font-weight:700;
                cursor:pointer;
                transition:all 0.2s ease;
            }
            .toolbar-btn:hover{
                border-color:#0a6ed1;
                color:#0a6ed1;
                background:#f4f9ff;
            }
            .toolbar-btn.primary{
                background:#0a6ed1;
                color:#ffffff;
                border-color:#0a6ed1;
            }
            .toolbar-btn.primary:hover{
                background:#085caf;
                color:#ffffff;
                border-color:#085caf;
            }
            .toolbar-btn.warn{
                background:#fff7e6;
                color:#8a5a00;
                border-color:#ffd591;
            }
            .toolbar-btn:disabled{
                opacity:0.55;
                cursor:not-allowed;
            }

            .upload-area{
                border:1px dashed #b8c4d1;
                border-radius:12px;
                padding:16px;
                background:#fafcff;
                margin-bottom:12px;
            }
            .upload-row{
                display:flex;
                gap:10px;
                align-items:center;
                flex-wrap:wrap;
            }
            .file-input{
                font-size:12px;
                padding:8px;
                border:1px solid #d9d9d9;
                border-radius:8px;
                background:#fff;
                min-width:280px;
                max-width:100%;
            }
            .mini-text{
                font-size:11px;
                color:#6a6d70;
                margin-top:8px;
            }

            .search-row{
                margin-top:12px;
                display:flex;
                gap:10px;
                flex-wrap:wrap;
                align-items:center;
            }
            .search-input{
                min-width:260px;
                max-width:100%;
                flex:1;
                padding:8px 10px;
                border:1px solid #d9d9d9;
                border-radius:8px;
                font-size:12px;
                outline:none;
            }
            .search-input:focus{
                border-color:#0a6ed1;
                box-shadow:0 0 0 2px rgba(10,110,209,0.12);
            }

            .progress-wrap{
                display:none;
                margin-top:12px;
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
                height:12px;
                background:#edf2f7;
                border-radius:10px;
                overflow:hidden;
                border:1px solid #d9d9d9;
            }
            .progress-fill{
                width:0%;
                height:100%;
                background:linear-gradient(90deg, #0a6ed1, #4db1ff);
                transition:width 0.25s ease;
            }

            .summary-grid{
                display:none;
                grid-template-columns:repeat(6, minmax(0, 1fr));
                gap:10px;
                margin-top:14px;
            }
            .summary-grid.show{
                display:grid;
            }
            .summary-item{
                border:1px solid #e5e7eb;
                border-radius:10px;
                padding:10px;
                background:#fafbfc;
            }
            .summary-item .k{
                font-size:11px;
                color:#6a6d70;
                margin-bottom:4px;
            }
            .summary-item .v{
                font-size:16px;
                font-weight:700;
                color:#1f2d3d;
            }

            .msg{
                display:none;
                margin-top:12px;
                padding:10px 12px;
                border-radius:10px;
                font-size:12px;
                line-height:1.45;
                white-space:pre-wrap;
            }
            .msg.show{ display:block; }
            .msg.error{ background:#fff2f0; border:1px solid #ffccc7; color:#a8071a; }
            .msg.warn{ background:#fffbe6; border:1px solid #ffe58f; color:#8a6d1d; }
            .msg.success{ background:#f6ffed; border:1px solid #b7eb8f; color:#237804; }
            .msg.info{ background:#f0f7ff; border:1px solid #bae0ff; color:#0958d9; }

            .preview-wrap{
                display:none;
                margin-top:14px;
                border:1px solid #e5e7eb;
                border-radius:12px;
                overflow:hidden;
                background:#ffffff;
            }
            .preview-wrap.show{
                display:block;
            }
            .preview-header{
                display:flex;
                justify-content:space-between;
                align-items:center;
                padding:10px 12px;
                background:#f7f9fb;
                border-bottom:1px solid #e5e7eb;
                gap:8px;
                flex-wrap:wrap;
            }
            .preview-title{
                font-size:13px;
                font-weight:700;
                color:#354a5f;
            }
            .preview-meta{
                font-size:11px;
                color:#6a6d70;
            }
            .preview-grid{
                overflow:auto;
                max-height:360px;
            }
            table{
                width:100%;
                border-collapse:collapse;
                min-width:700px;
            }
            th, td{
                border-bottom:1px solid #eef2f6;
                padding:8px 10px;
                text-align:left;
                vertical-align:top;
                font-size:12px;
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
            }
            .cell-error{
                display:block;
                margin-top:4px;
                color:#cf1322;
                font-size:11px;
                line-height:1.3;
                white-space:normal;
            }

            .log-box{
                display:none;
                margin-top:12px;
                border:1px solid #e5e7eb;
                border-radius:10px;
                background:#fafbfc;
                padding:12px;
                max-height:180px;
                overflow:auto;
                font-size:12px;
                line-height:1.45;
                color:#354a5f;
                white-space:pre-wrap;
            }
            .log-box.show{
                display:block;
            }

            .footer-note{
                margin-top:12px;
                font-size:11px;
                color:#6a6d70;
            }
        </style>

        <div class="excel-widget">
            <div class="upload-card">
                <div class="upload-header">
                    <div class="title-wrap">
                        <div class="upload-title" id="titleEl">Excel Upload</div>
                        <div class="upload-subtitle" id="subtitleEl">Upload and validate Excel file</div>
                    </div>
                    <div class="status-badge status-ready" id="statusBadge">Ready</div>
                </div>

                <div class="toolbar-actions">
                    <button type="button" class="toolbar-btn primary" id="downloadTemplateBtn">Download Template</button>
                    <button type="button" class="toolbar-btn" id="downloadErrorBtn" disabled>Download Error Log</button>
                    <button type="button" class="toolbar-btn warn" id="revalidateBtn">Refresh Validation</button>
                </div>

                <div class="upload-area">
                    <div class="upload-row">
                        <input type="file" id="fileInput" class="file-input" />
                        <button type="button" class="toolbar-btn primary" id="uploadBtn">Upload</button>
                        <button type="button" class="toolbar-btn" id="clearBtn">Clear</button>
                    </div>

                    <div class="mini-text" id="acceptText">Accepted: .xls, .xlsx, .xlsm, .csv</div>

                    <div class="search-row">
                        <input type="text" id="searchInput" class="search-input" placeholder="Search in preview..." />
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

                <div class="summary-grid" id="summaryGrid">
                    <div class="summary-item"><div class="k">Rows Read</div><div class="v" id="sumRows">0</div></div>
                    <div class="summary-item"><div class="k">Valid Rows</div><div class="v" id="sumValid">0</div></div>
                    <div class="summary-item"><div class="k">Invalid Rows</div><div class="v" id="sumInvalid">0</div></div>
                    <div class="summary-item"><div class="k">Sheet</div><div class="v" id="sumSheet">-</div></div>
                    <div class="summary-item"><div class="k">Columns</div><div class="v" id="sumColumns">0</div></div>
                    <div class="summary-item"><div class="k">Validation</div><div class="v" id="sumValidation">-</div></div>
                </div>

                <div class="msg" id="messageBox"></div>

                <div class="preview-wrap" id="previewWrap">
                    <div class="preview-header">
                        <div class="preview-title">Preview</div>
                        <div class="preview-meta" id="previewMeta">0 row(s)</div>
                    </div>
                    <div class="preview-grid">
                        <table>
                            <thead id="previewHead"></thead>
                            <tbody id="previewBody"></tbody>
                        </table>
                    </div>
                </div>

                <div class="log-box" id="logBox"></div>
                <div class="footer-note" id="footerNote">Supported template: Sheet1 with columns ID, DESCRIPTION, H1, costcenter</div>
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
                requiredcolumns: "ID,DESCRIPTION,H1,costcenter",
                maxrows: 2000,
                previewrows: 50,
                allowcsv: true,
                autovalidate: true,
                showpreview: true,
                showlogs: true,
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
            this._designMode = false;
            this._searchText = "";
            this._validationErrorsParsed = [];
            this._validationMap = {};

            this._bindEvents();
        }

        connectedCallback() {
            this._setStatus("Ready", "ready");
            this._applyHeaderSettings();
            this._applyAcceptedTypes();
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
            if ("maxrows" in changedProperties) this.maxrows = changedProperties.maxrows;
            if ("previewrows" in changedProperties) this.previewrows = changedProperties.previewrows;
            if ("allowcsv" in changedProperties) this.allowcsv = changedProperties.allowcsv;
            if ("autovalidate" in changedProperties) this.autovalidate = changedProperties.autovalidate;
            if ("showpreview" in changedProperties) this.showpreview = changedProperties.showpreview;
            if ("showlogs" in changedProperties) this.showlogs = changedProperties.showlogs;
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
                "maxrows",
                "previewrows",
                "allowcsv",
                "autovalidate",
                "showpreview",
                "showlogs",
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
        set requiredcolumns(v) { this._export_settings.requiredcolumns = v || "ID,DESCRIPTION,H1,costcenter"; }

        get maxrows() { return this._export_settings.maxrows; }
        set maxrows(v) {
            var n = parseInt(v, 10);
            this._export_settings.maxrows = isNaN(n) || n <= 0 ? 2000 : n;
        }

        get previewrows() { return this._export_settings.previewrows; }
        set previewrows(v) {
            var n = parseInt(v, 10);
            this._export_settings.previewrows = isNaN(n) || n <= 0 ? 50 : n;
            this._renderPreview();
        }

        get allowcsv() { return this._export_settings.allowcsv; }
        set allowcsv(v) { this._export_settings.allowcsv = this._toBoolean(v, true); this._applyAcceptedTypes(); }

        get autovalidate() { return this._export_settings.autovalidate; }
        set autovalidate(v) { this._export_settings.autovalidate = this._toBoolean(v, true); }

        get showpreview() { return this._export_settings.showpreview; }
        set showpreview(v) { this._export_settings.showpreview = this._toBoolean(v, true); this._applyVisibility(); }

        get showlogs() { return this._export_settings.showlogs; }
        set showlogs(v) { this._export_settings.showlogs = this._toBoolean(v, true); this._applyVisibility(); }

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
                that._shadowRoot.getElementById("revalidateBtn").addEventListener("click", function () {
                    that._applySacValidation();
                    that._dispatchValidate();
                });
                that._shadowRoot.getElementById("searchInput").addEventListener("input", function (e) {
                    that._searchText = e.target.value || "";
                    that._renderPreview();
                });
            }, 0);
        }

        _applyHeaderSettings() {
            var titleEl = this._shadowRoot.getElementById("titleEl");
            var subtitleEl = this._shadowRoot.getElementById("subtitleEl");
            var footerEl = this._shadowRoot.getElementById("footerNote");

            if (titleEl) titleEl.textContent = this._export_settings.title || "Excel Upload";
            if (subtitleEl) subtitleEl.textContent = this._export_settings.subtitle || "Upload and validate Excel file";
            if (footerEl) footerEl.textContent = this._export_settings.footer || "Supported template: Sheet1 with columns ID, DESCRIPTION, H1, costcenter";
        }

        _applyAcceptedTypes() {
            var input = this._shadowRoot.getElementById("fileInput");
            var acceptText = this._shadowRoot.getElementById("acceptText");
            var accept = this._export_settings.allowcsv ? ".xls,.xlsx,.xlsm,.csv" : ".xls,.xlsx,.xlsm";
            if (input) input.setAttribute("accept", accept);
            if (acceptText) acceptText.textContent = "Accepted: " + accept;
        }

        _applyVisibility() {
            var previewWrap = this._shadowRoot.getElementById("previewWrap");
            var logBox = this._shadowRoot.getElementById("logBox");

            if (previewWrap) {
                if (this._export_settings.showpreview) previewWrap.classList.add("show");
                else previewWrap.classList.remove("show");
            }

            if (logBox) {
                if (this._export_settings.showlogs && logBox.textContent !== "") logBox.classList.add("show");
                else if (!this._export_settings.showlogs) logBox.classList.remove("show");
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

            that._log("Downloading template...", false);
            that._setStatus("Downloading", "processing");

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

            this._setStatus("Processing", "processing");
            this._setProgress(10, "Reading file...");
            this._errorLog = [];
            this._validData = [];
            this._previewRows = [];
            this._previewColumns = [];
            this._validationErrorsParsed = [];
            this._validationMap = {};
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
                    that._setProgress(50, "Reading rows...");
                    var sheet = workbook.Sheets[actualSheet];
                    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

                    if (!rows || rows.length <= 1) {
                        that._setStatus("Error", "error");
                        that._showMessage("error", "There is no record to be uploaded");
                        that._log("There is no record to be uploaded", true);
                        return;
                    }

                    var header = rows[0];
                    that._previewColumns = header.slice(0);

                    var requiredColumns = that._getRequiredColumns();
                    var colMap = {};
                    var hc = 0;
                    for (hc = 0; hc < requiredColumns.length; hc++) {
                        colMap[requiredColumns[hc]] = -1;
                    }

                    for (var i = 0; i < header.length; i++) {
                        var h = String(header[i] || "").trim();
                        for (hc = 0; hc < requiredColumns.length; hc++) {
                            if (String(requiredColumns[hc]).toLowerCase() === h.toLowerCase()) {
                                colMap[requiredColumns[hc]] = i;
                            }
                        }
                    }

                    var missingColumns = [];
                    for (hc = 0; hc < requiredColumns.length; hc++) {
                        if (colMap[requiredColumns[hc]] === -1) {
                            missingColumns.push(requiredColumns[hc]);
                        }
                    }

                    if (missingColumns.length > 0) {
                        that._setStatus("Error", "error");
                        that._showMessage("error", "Invalid template. Missing columns: " + missingColumns.join(", "));
                        that._log("Invalid template. Missing columns: " + missingColumns.join(", "), true);
                        return;
                    }

                    that._setProgress(65, "Scanning for duplicates...");

                    var idKey = requiredColumns.indexOf("ID") > -1 ? "ID" : requiredColumns[0];
                    var idCount = {};
                    for (var r = 1; r < rows.length; r++) {
                        var row = rows[r] || [];
                        var id = String(row[colMap[idKey]] || "").trim();
                        if (!id) continue;
                        idCount[id] = (idCount[id] || 0) + 1;
                    }

                    that._setProgress(80, "Validating rows...");

                    var validRows = [];
                    var errorRows = [];
                    var previewRows = [];

                    for (var r2 = 1; r2 < rows.length; r2++) {
                        var row2 = rows[r2] || [];
                        var rowNumber = r2 + 1;
                        var rowObj = {};
                        var isBlankRow = true;

                        for (hc = 0; hc < requiredColumns.length; hc++) {
                            var colName = requiredColumns[hc];
                            rowObj[colName] = String(row2[colMap[colName]] || "").trim();
                            if (rowObj[colName] !== "") isBlankRow = false;
                        }

                        if (isBlankRow) {
                            continue;
                        }

                        var errors = [];
                        var cellErrors = [];

                        for (hc = 0; hc < requiredColumns.length; hc++) {
                            var reqCol = requiredColumns[hc];
                            if (!rowObj[reqCol]) {
                                errors.push(reqCol + " is mandatory");
                                cellErrors.push({
                                    rowIndex: previewRows.length,
                                    field: reqCol,
                                    message: "Blank cell not allowed"
                                });
                            }
                        }

                        if (rowObj[idKey] && idCount[rowObj[idKey]] > 1) {
                            errors.push("Duplicate " + idKey + " '" + rowObj[idKey] + "' - all " + idCount[rowObj[idKey]] + " occurrences rejected");
                            cellErrors.push({
                                rowIndex: previewRows.length,
                                field: idKey,
                                message: "Duplicate value not allowed"
                            });
                        }

                        var previewRowObj = {
                            _rowNumber: rowNumber
                        };
                        for (hc = 0; hc < requiredColumns.length; hc++) {
                            previewRowObj[requiredColumns[hc]] = rowObj[requiredColumns[hc]];
                        }
                        previewRows.push(previewRowObj);

                        if (errors.length > 0) {
                            errorRows.push({
                                RowNumber: rowNumber,
                                ErrorMessage: errors.join(" | ")
                            });
                            for (var ce = 0; ce < cellErrors.length; ce++) {
                                that._validationErrorsParsed.push(cellErrors[ce]);
                            }
                        } else {
                            validRows.push(rowObj);
                        }
                    }

                    if (validRows.length > that._export_settings.maxrows) {
                        that._setStatus("Error", "error");
                        that._showMessage("error", "Maximum valid records are " + that._export_settings.maxrows);
                        that._log("Maximum valid records are " + that._export_settings.maxrows, true);
                        return;
                    }

                    that._validData = validRows;
                    that._errorLog = errorRows;
                    that._previewRows = previewRows;
                    that._enableErrorDownload(errorRows.length > 0);

                    that.unit = JSON.stringify(validRows);
                    that._export_settings.rowcount = previewRows.length;
                    that._export_settings.validcount = validRows.length;
                    that._export_settings.invalidcount = errorRows.length;

                    that._buildValidationMap();
                    that._renderPreview();
                    that._firePropertiesChanged("uploadCompleted");

                    that.dispatchEvent(new CustomEvent("onStart", {
                        detail: {
                            settings: {},
                            rowCount: validRows.length,
                            invalidCount: errorRows.length,
                            fileName: file.name,
                            sheetName: actualSheet
                        }
                    }));

                    that._setSummary(
                        previewRows.length,
                        validRows.length,
                        errorRows.length,
                        actualSheet,
                        requiredColumns.length,
                        errorRows.length > 0 ? "Invalid" : "Valid"
                    );

                    that._setProgress(100, "Completed");

                    if (errorRows.length > 0) {
                        that._setStatus("Completed with Errors", "warning");
                        that._showMessage("warn", "Upload parsed with validation issues. Review highlighted rows/cells.");
                    } else {
                        that._setStatus("Completed", "completed");
                        that._showMessage("success", "Upload completed successfully.");
                    }

                    that._log("Valid rows: " + validRows.length);
                    that._log("Invalid rows: " + errorRows.length);

                    var dupIds = Object.keys(idCount).filter(function (id) {
                        return idCount[id] > 1;
                    });
                    if (dupIds.length > 0) {
                        that._log("Duplicate IDs rejected: " + dupIds.join(", "));
                    }

                } catch (err) {
                    that._setStatus("Error", "error");
                    that._showMessage("error", "Processing failed: " + err.message);
                    that._log("Processing failed: " + err.message, true);
                }
            };

            reader.readAsBinaryString(file);
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
            for (var i = 0; i < this._errorLog.length; i++) {
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

        _getRequiredColumns() {
            var raw = String(this._export_settings.requiredcolumns || "");
            var parts = raw.split(",");
            var out = [];
            for (var i = 0; i < parts.length; i++) {
                var v = String(parts[i] || "").trim();
                if (v !== "") out.push(v);
            }
            return out.length ? out : ["ID", "DESCRIPTION", "H1", "costcenter"];
        }

        _applySacValidation() {
            var parsed = this._safeParseArray(this._export_settings.validationerrors);
            this._validationErrorsParsed = parsed.length ? parsed : this._validationErrorsParsed;
            this._buildValidationMap();
            this._renderPreview();

            if (this._export_settings.validationresult === "false") {
                this._showMessage("error", this._buildValidationSummaryText());
                this._setStatus("Validation Error", "error");
            }
        }

        _buildValidationSummaryText() {
            if (!this._validationErrorsParsed || this._validationErrorsParsed.length === 0) {
                return "Validation failed.";
            }

            var lines = [];
            for (var i = 0; i < this._validationErrorsParsed.length; i++) {
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
            for (var i = 0; i < this._validationErrorsParsed.length; i++) {
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

            if (!this._export_settings.showpreview) {
                wrap.classList.remove("show");
                return;
            }

            wrap.classList.add("show");
            head.innerHTML = "";
            body.innerHTML = "";

            if (!this._previewRows || this._previewRows.length === 0) {
                meta.textContent = "0 row(s)";
                body.innerHTML = '<tr><td colspan="20">No preview available</td></tr>';
                return;
            }

            var cols = [];
            var sample = this._previewRows[0];
            for (var k in sample) {
                if (k !== "_rowNumber") cols.push(k);
            }

            var filteredRows = [];
            var sText = String(this._searchText || "").toLowerCase().trim();
            for (var i = 0; i < this._previewRows.length; i++) {
                var row = this._previewRows[i];
                if (sText === "") {
                    filteredRows.push(row);
                } else {
                    var matched = false;
                    for (var j = 0; j < cols.length; j++) {
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
            if (filteredRows.length > limit) {
                filteredRows = filteredRows.slice(0, limit);
            }

            meta.textContent = filteredRows.length + " row(s) shown";

            var headHtml = "<tr><th>#</th>";
            for (var c = 0; c < cols.length; c++) {
                headHtml += "<th>" + this._escapeHtml(cols[c]) + "</th>";
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

                for (c = 0; c < cols.length; c++) {
                    var field = cols[c];
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
            for (var i = 0; i < this._previewRows.length; i++) {
                if (this._previewRows[i] === rowObj) return i;
                if (this._previewRows[i]._rowNumber === rowObj._rowNumber) return i;
            }
            return 0;
        }

        _dispatchValidate() {
            this._export_settings.lastevent = "validate";
            this.dispatchEvent(new CustomEvent("onValidate", {
                detail: {}
            }));
            this._firePropertiesChanged("validate");
        }

        _safeParseArray(text) {
            try {
                var parsed = JSON.parse(text || "[]");
                return Array.isArray(parsed) ? parsed : [];
            } catch (e) {
                return [];
            }
        }

        _escapeCsv(value) {
            var str = value == null ? "" : String(value);
            if (str.indexOf(",") > -1 || str.indexOf('"') > -1 || str.indexOf("\n") > -1) {
                str = '"' + str.replace(/"/g, '""') + '"';
            }
            return str;
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

