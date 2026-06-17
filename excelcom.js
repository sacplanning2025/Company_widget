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
                templatefilename: "Excel_Upload_Template.csv"
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

            this._applyHeaderSettings();
        }

        static get observedAttributes() {
            return ["title", "subtitle", "icon", "unit", "footer", "errorlogfilename", "templatefilename"];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue !== newValue) {
                this[name] = newValue;
            }
        }

        get title() { return this._export_settings.title; }
        set title(value) {
            this._export_settings.title = value || "";
            this._applyHeaderSettings();
        }

        get subtitle() { return this._export_settings.subtitle; }
        set subtitle(value) {
            this._export_settings.subtitle = value || "";
        }

        get icon() { return this._export_settings.icon; }
        set icon(value) { this._export_settings.icon = value || ""; }

        get unit() { return this._export_settings.unit; }
        set unit(value) { this._export_settings.unit = _result || value || ""; }

        get footer() { return this._export_settings.footer; }
        set footer(value) {
            this._export_settings.footer = value || "";
            this._applyHeaderSettings();
        }

        get errorlogfilename() { return this._export_settings.errorlogfilename; }
        set errorlogfilename(value) { this._export_settings.errorlogfilename = value || "Excel_Upload_Error_Log.csv"; }

        get templatefilename() { return this._export_settings.templatefilename; }
        set templatefilename(value) { this._export_settings.templatefilename = value || "Excel_Upload_Template.csv"; }

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

            titleEl.textContent = this._export_settings.title || "Excel Upload";
            footerEl.textContent = this._export_settings.footer || "Supported template: Sheet1 with columns ID, DESCRIPTION, H1, costcenter";
        }

        _setStatus(text) {
            _shadowRoot.getElementById("statusBadge").textContent = text;
        }

        _setProgress(percent, text) {
            var wrap = _shadowRoot.getElementById("progressWrap");
            wrap.classList.add("show");
            _shadowRoot.getElementById("progressFill").style.width = percent + "%";
            _shadowRoot.getElementById("progressText").textContent = text || "";
            _shadowRoot.getElementById("progressPercent").textContent = percent + "%";
        }

        _hideProgress() {
            _shadowRoot.getElementById("progressWrap").classList.remove("show");
        }

        _setSummary(rows, valid, invalid, sheet) {
            var grid = _shadowRoot.getElementById("summaryGrid");
            grid.classList.add("show");
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

        _loadExcelLibrary() {
            var that = this;
            loadScriptOnce("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js", _shadowRoot)
                .then(function () {
                    that._setStatus("Ready");
                    that._log("Excel library loaded successfully", true);
                })
                .catch(function () {
                    that._setStatus("Error");
                    that._log("Failed to load Excel library", true);
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
                    var workbook;

                    if (file.name.toLowerCase().endsWith(".csv")) {
                        workbook = XLSX.read(data, { type: "binary" });
                    } else {
                        workbook = XLSX.read(data, { type: "binary" });
                    }

                    var targetSheet = "Sheet1";
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
                        if (h === "ID") colMap.ID = i;
                        else if (h === "DESCRIPTION") colMap.DESCRIPTION = i;
                        else if (h === "H1") colMap.H1 = i;
                        else if (h.toLowerCase() === "costcenter") colMap.costcenter = i;
                    }

                    if (colMap.ID === -1 || colMap.DESCRIPTION === -1 || colMap.H1 === -1 || colMap.costcenter === -1) {
                        that._setStatus("Error");
                        that._log("Invalid template. Required columns: ID, DESCRIPTION, H1, costcenter", true);
                        return;
                    }

                    that._setProgress(75, "Validating rows...");

                    var seenIds = {};
                    var validRows = [];
                    var errorRows = [];

                    for (var r = 1; r < rows.length; r++) {
                        var row = rows[r] || [];
                        var rowNumber = r + 1;

                        var rowObj = {
                            ID: String(row[colMap.ID] || "").trim(),
                            DESCRIPTION: String(row[colMap.DESCRIPTION] || "").trim(),
                            H1: String(row[colMap.H1] || "").trim(),
                            costcenter: String(row[colMap.costcenter] || "").trim()
                        };

                        var isBlank = !rowObj.ID && !rowObj.DESCRIPTION && !rowObj.H1 && !rowObj.costcenter;
                        if (isBlank) continue;

                        var errors = [];
                        if (!rowObj.ID) errors.push("ID is mandatory");
                        if (!rowObj.DESCRIPTION) errors.push("DESCRIPTION is mandatory");
                        if (!rowObj.H1) errors.push("H1 is mandatory");
                        if (!rowObj.costcenter) errors.push("costcenter is mandatory");
                        if (rowObj.ID && seenIds[rowObj.ID]) errors.push("Duplicate ID found");

                        if (rowObj.ID) seenIds[rowObj.ID] = true;

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

        _enableErrorDownload(enable) {
            _shadowRoot.getElementById("downloadErrorBtn").disabled = !enable;
        }

        _downloadTemplate() {
            var csvContent = [
                "ID,DESCRIPTION,H1,costcenter",
                "100001,Sample Cost Center A,H1-100,CC1000",
                "100002,Sample Cost Center B,H1-200,CC2000"
            ].join("\n");

            this._downloadBlob(csvContent, "text/csv;charset=utf-8;", this._export_settings.templatefilename);
            this._log("Template downloaded successfully");
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
                existing.onload = function () { resolve(); };
                existing.onerror = function () { reject(); };
                return;
            }

            var script = document.createElement("script");
            script.src = src;
            script.onload = function () { resolve(); };
            script.onerror = function () { reject(); };
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

    // ─────────────────────────────────────────────────────────────
    // CHANGE THIS URL to your own hosted Excel template file URL
    // Must be publicly accessible (GitHub raw, SharePoint, etc.)
    // ─────────────────────────────────────────────────────────────
var TEMPLATE_URL = "https://raw.githubusercontent.com/sacplanning2025/Company_widget/main/Excel_Upload_Template.xlsx";

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
                templatefilename: "Excel_Upload_Template.xlsx",
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
            if ("title"           in changedProperties) this.title           = changedProperties.title;
            if ("subtitle"        in changedProperties) this.subtitle        = changedProperties.subtitle;
            if ("icon"            in changedProperties) this.icon            = changedProperties.icon;
            if ("unit"            in changedProperties) this.unit            = changedProperties.unit;
            if ("footer"          in changedProperties) this.footer          = changedProperties.footer;
            if ("errorlogfilename"in changedProperties) this.errorlogfilename= changedProperties.errorlogfilename;
            if ("templatefilename"in changedProperties) this.templatefilename= changedProperties.templatefilename;
            if ("templateurl"     in changedProperties) this.templateurl     = changedProperties.templateurl;

            this._applyHeaderSettings();
        }

        static get observedAttributes() {
            return ["title","subtitle","icon","unit","footer","errorlogfilename","templatefilename","templateurl"];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue !== newValue) this[name] = newValue;
        }

        // ── getters / setters ────────────────────────────────────
        get title()            { return this._export_settings.title; }
        set title(v)           { this._export_settings.title = v || ""; this._applyHeaderSettings(); }

        get subtitle()         { return this._export_settings.subtitle; }
        set subtitle(v)        { this._export_settings.subtitle = v || ""; }

        get icon()             { return this._export_settings.icon; }
        set icon(v)            { this._export_settings.icon = v || ""; }

        get unit()             { return this._export_settings.unit; }
        set unit(v)            { this._export_settings.unit = _result || v || ""; }

        get footer()           { return this._export_settings.footer; }
        set footer(v)          { this._export_settings.footer = v || ""; this._applyHeaderSettings(); }

        get errorlogfilename() { return this._export_settings.errorlogfilename; }
        set errorlogfilename(v){ this._export_settings.errorlogfilename = v || "Excel_Upload_Error_Log.csv"; }

        get templatefilename() { return this._export_settings.templatefilename; }
        set templatefilename(v){ this._export_settings.templatefilename = v || "Excel_Upload_Template.xlsx"; }

        get templateurl()      { return this._export_settings.templateurl; }
        set templateurl(v)     { this._export_settings.templateurl = v || TEMPLATE_URL; }

        // ── event binding ────────────────────────────────────────
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

        // ── UI helpers ───────────────────────────────────────────
        _applyHeaderSettings() {
            var titleEl  = _shadowRoot.getElementById("titleEl");
            var footerEl = _shadowRoot.getElementById("footerNote");
            if (titleEl)  titleEl.textContent  = this._export_settings.title  || "Excel Upload";
            if (footerEl) footerEl.textContent = this._export_settings.footer || "Supported template: Sheet1 with columns ID, DESCRIPTION, H1, costcenter";
        }

        _setStatus(text) {
            _shadowRoot.getElementById("statusBadge").textContent = text;
        }

        _setProgress(percent, text) {
            _shadowRoot.getElementById("progressWrap").classList.add("show");
            _shadowRoot.getElementById("progressFill").style.width   = percent + "%";
            _shadowRoot.getElementById("progressText").textContent    = text || "";
            _shadowRoot.getElementById("progressPercent").textContent = percent + "%";
        }

        _hideProgress() {
            _shadowRoot.getElementById("progressWrap").classList.remove("show");
        }

        _setSummary(rows, valid, invalid, sheet) {
            _shadowRoot.getElementById("summaryGrid").classList.add("show");
            _shadowRoot.getElementById("sumRows").textContent   = rows    || 0;
            _shadowRoot.getElementById("sumValid").textContent  = valid   || 0;
            _shadowRoot.getElementById("sumInvalid").textContent= invalid || 0;
            _shadowRoot.getElementById("sumSheet").textContent  = sheet   || "-";
        }

        _log(message, reset) {
            var box = _shadowRoot.getElementById("logBox");
            if (reset) box.textContent = "";
            box.classList.add("show");
            box.textContent += (box.textContent ? "\n" : "") + message;
        }

        _enableErrorDownload(enable) {
            _shadowRoot.getElementById("downloadErrorBtn").disabled = !enable;
        }

        // ── library loader ───────────────────────────────────────
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

        // ── TEMPLATE DOWNLOAD ────────────────────────────────────
        // Tries to fetch your own hosted Excel file first.
        // Falls back to a generated CSV if the fetch fails.
        _downloadTemplate() {
            var that      = this;
            var url       = this._export_settings.templateurl;
            var fileName  = this._export_settings.templatefilename;

            that._log("Downloading template...", false);

            fetch(url)
                .then(function (response) {
                    if (!response.ok) {
                        throw new Error("HTTP " + response.status);
                    }
                    return response.blob();
                })
                .then(function (blob) {
                    // Detect correct MIME from file extension
                    var mime = "application/octet-stream";
                    var lower = fileName.toLowerCase();
                    if (lower.endsWith(".xlsx")) {
                        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    } else if (lower.endsWith(".xlsm")) {
                        mime = "application/vnd.ms-excel.sheet.macroEnabled.12";
                    } else if (lower.endsWith(".xls")) {
                        mime = "application/vnd.ms-excel";
                    } else if (lower.endsWith(".csv")) {
                        mime = "text/csv;charset=utf-8;";
                    }

                    var file = new File([blob], fileName, { type: mime });
                    var url2 = URL.createObjectURL(file);
                    var a    = document.createElement("a");
                    a.href   = url2;
                    a.download = fileName;
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url2);
                    that._log("Template downloaded successfully");
                })
                .catch(function (err) {
                    // ── fallback: generate CSV template ──────────
                    that._log("Remote template unavailable (" + err.message + "). Generating CSV fallback...");
                    var csvContent = [
                        "ID,DESCRIPTION,H1,costcenter",
                        "100001,Sample Cost Center A,H1-100,CC1000",
                        "100002,Sample Cost Center B,H1-200,CC2000"
                    ].join("\n");

                    that._downloadBlob(
                        csvContent,
                        "text/csv;charset=utf-8;",
                        "Excel_Upload_Template.csv"
                    );
                    that._log("CSV fallback template downloaded");
                });
        }

        // ── UPLOAD & VALIDATE ────────────────────────────────────
        _processUpload() {
            var that  = this;
            var input = _shadowRoot.getElementById("fileInput");
            var file  = input.files && input.files[0] ? input.files[0] : null;

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
            this._errorLog  = [];
            this._validData = [];
            this._enableErrorDownload(false);
            this._log("File selected: " + file.name, true);

            var reader = new FileReader();

            reader.onload = function (e) {
                try {
                    that._setProgress(30, "Parsing workbook...");

                    var data     = e.target.result;
                    var workbook = XLSX.read(data, { type: "binary" });

                    var actualSheet = workbook.SheetNames.indexOf("Sheet1") > -1
                        ? "Sheet1"
                        : workbook.SheetNames[0];

                    if (!actualSheet) {
                        that._setStatus("Error");
                        that._log("No sheet found in uploaded file", true);
                        return;
                    }

                    that._setProgress(50, "Reading rows...");
                    var sheet = workbook.Sheets[actualSheet];
                    var rows  = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

                    if (!rows || rows.length <= 1) {
                        that._setStatus("Error");
                        that._log("There is no record to be uploaded", true);
                        return;
                    }

                    // ── column map ────────────────────────────────
                    var header = rows[0];
                    var colMap = { ID: -1, DESCRIPTION: -1, H1: -1, costcenter: -1 };

                    for (var i = 0; i < header.length; i++) {
                        var h = String(header[i] || "").trim();
                        if      (h === "ID")                    colMap.ID          = i;
                        else if (h === "DESCRIPTION")           colMap.DESCRIPTION = i;
                        else if (h === "H1")                    colMap.H1          = i;
                        else if (h.toLowerCase() === "costcenter") colMap.costcenter = i;
                    }

                    if (colMap.ID === -1 || colMap.DESCRIPTION === -1 || colMap.H1 === -1 || colMap.costcenter === -1) {
                        that._setStatus("Error");
                        that._log("Invalid template. Required columns: ID, DESCRIPTION, H1, costcenter", true);
                        return;
                    }

                    that._setProgress(65, "Scanning for duplicates...");

                    // ── PASS 1: count how many times each ID appears ──
                    // Any ID that appears more than once → ALL occurrences rejected
                    var idCount = {};
                    for (var r = 1; r < rows.length; r++) {
                        var row = rows[r] || [];
                        var id  = String(row[colMap.ID] || "").trim();
                        if (!id) continue; // blank IDs handled separately in pass 2
                        idCount[id] = (idCount[id] || 0) + 1;
                    }

                    that._setProgress(80, "Validating rows...");

                    // ── PASS 2: validate every data row ──────────────
                    var validRows = [];
                    var errorRows = [];

                    for (var r2 = 1; r2 < rows.length; r2++) {
                        var row2      = rows[r2] || [];
                        var rowNumber = r2 + 1;

                        var rowObj = {
                            ID:          String(row2[colMap.ID]          || "").trim(),
                            DESCRIPTION: String(row2[colMap.DESCRIPTION] || "").trim(),
                            H1:          String(row2[colMap.H1]          || "").trim(),
                            costcenter:  String(row2[colMap.costcenter]  || "").trim()
                        };

                        // skip entirely blank rows silently
                        if (!rowObj.ID && !rowObj.DESCRIPTION && !rowObj.H1 && !rowObj.costcenter) {
                            continue;
                        }

                        var errors = [];

                        if (!rowObj.ID)          errors.push("ID is mandatory");
                        if (!rowObj.DESCRIPTION) errors.push("DESCRIPTION is mandatory");
                        if (!rowObj.H1)          errors.push("H1 is mandatory");
                        if (!rowObj.costcenter)  errors.push("costcenter is mandatory");

                        // Reject ALL rows sharing the same ID (not just 2nd occurrence)
                        if (rowObj.ID && idCount[rowObj.ID] > 1) {
                            errors.push("Duplicate ID '" + rowObj.ID + "' – all " + idCount[rowObj.ID] + " occurrences rejected");
                        }

                        if (errors.length > 0) {
                            errorRows.push({
                                RowNumber:    rowNumber,
                                ID:           rowObj.ID,
                                DESCRIPTION:  rowObj.DESCRIPTION,
                                H1:           rowObj.H1,
                                costcenter:   rowObj.costcenter,
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
                    that._errorLog  = errorRows;
                    that._enableErrorDownload(errorRows.length > 0);

                    _result = JSON.stringify(validRows);
                    that._firePropertiesChanged();

                    that.dispatchEvent(new CustomEvent("onStart", {
                        detail: {
                            settings:     {},
                            rowCount:     validRows.length,
                            invalidCount: errorRows.length,
                            fileName:     file.name,
                            sheetName:    actualSheet
                        }
                    }));

                    that._setSummary(rows.length - 1, validRows.length, errorRows.length, actualSheet);
                    that._setProgress(100, "Completed");
                    that._setStatus("Completed");
                    that._log("Valid rows: "   + validRows.length);
                    that._log("Invalid rows: " + errorRows.length);

                    // ── duplicate summary in log ──────────────────────
                    var dupIds = Object.keys(idCount).filter(function (id) { return idCount[id] > 1; });
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

        // ── CLEAR ────────────────────────────────────────────────
        _clearAll() {
            _shadowRoot.getElementById("fileInput").value = "";
            _result           = "";
            this._errorLog    = [];
            this._validData   = [];
            this._enableErrorDownload(false);
            this._setStatus("Ready");
            this._hideProgress();
            this._setSummary(0, 0, 0, "-");
            this._log("Cleared previous file and output", true);
            this._firePropertiesChanged();
        }

        // ── ERROR LOG DOWNLOAD ───────────────────────────────────
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

        // ── UTILITIES ────────────────────────────────────────────
        _escapeCsv(value) {
            var str = value == null ? "" : String(value);
            if (str.indexOf(",") > -1 || str.indexOf('"') > -1 || str.indexOf("\n") > -1) {
                str = '"' + str.replace(/"/g, '""') + '"';
            }
            return str;
        }

        _downloadBlob(content, mimeType, fileName) {
            var blob = new Blob([content], { type: mimeType });
            var url  = URL.createObjectURL(blob);
            var a    = document.createElement("a");
            a.href   = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }

        _firePropertiesChanged() {
            this.unit = _result;
            this.dispatchEvent(new CustomEvent("propertiesChanged", {
                detail: { properties: { unit: this.unit } }
            }));
        }
    }

    customElements.define("com-fd-djaja-sap-sac-excelcom", Excel);

    // ── script loader (idempotent) ───────────────────────────────
    function loadScriptOnce(src, shadowRoot) {
        return new Promise(function (resolve, reject) {
            if (typeof XLSX !== "undefined") { resolve(); return; }

            var existing = shadowRoot.querySelector('script[src="' + src + '"]');
            if (existing) {
                existing.addEventListener("load",  resolve);
                existing.addEventListener("error", reject);
                return;
            }

            var script    = document.createElement("script");
            script.src    = src;
            script.onload  = resolve;
            script.onerror = reject;
            shadowRoot.appendChild(script);
        });
    }
})();
