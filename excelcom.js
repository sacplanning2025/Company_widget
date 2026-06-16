(function () {
    "use strict";

    // =========================================================
    // MAIN WIDGET TEMPLATE
    // =========================================================
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

            .upload-subtitle{
                font-size:12px;
                color:#6a6d70;
                margin-top:2px;
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
                margin-bottom:10px;
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
                box-shadow:0 0 0 2px rgba(10,110,209,0.08);
            }

            .toolbar-btn.primary{
                background:#0a6ed1;
                color:#ffffff;
                border-color:#0a6ed1;
            }

            .toolbar-btn.primary:hover{
                background:#085caf;
                border-color:#085caf;
                color:#ffffff;
            }

            .toolbar-btn:disabled{
                opacity:0.55;
                cursor:not-allowed;
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

            @media (max-width: 700px){
                .summary-grid{
                    grid-template-columns:repeat(2, minmax(0, 1fr));
                }
            }
        </style>
        <div class="excel-widget">
            <div class="upload-card">
                <div class="upload-header">
                    <div>
                        <div class="upload-title">Excel Upload</div>
                        <div class="upload-subtitle">Upload and validate Excel data before sending it to SAC script</div>
                    </div>
                    <div class="status-badge" id="statusBadge">Ready</div>
                </div>

                <div class="toolbar-actions">
                    <button type="button" class="toolbar-btn primary" id="downloadTemplateBtn">Download Template</button>
                    <button type="button" class="toolbar-btn" id="downloadErrorBtn" disabled>Download Error Log</button>
                </div>

                <div id="ui5_host"></div>

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

                <div class="footer-note" id="footerNote">
                    Supported template: Sheet1 with columns ID, DESCRIPTION, H1, costcenter
                </div>
            </div>
        </div>
    `;

    var _shadowRoot;
    var _result = "";
    var _widgetId = "";
    var _ui5ViewPlaced = false;
    var _xlsxLoaded = false;

    class Excel extends HTMLElement {
        constructor() {
            super();

            _shadowRoot = this.attachShadow({ mode: "open" });
            _shadowRoot.appendChild(tmpl.content.cloneNode(true));

            _widgetId = createGuid();

            this._export_settings = {
                title: "",
                subtitle: "",
                icon: "",
                unit: "",
                footer: "",
                widgetName: ""
            };

            this._subscription = null;
            this._firstConnection = 0;
            this._designMode = false;
            this.metadata = "";
            this._errorLog = [];
            this._validData = [];

            this._bindToolbarActions();
        }

        connectedCallback() {
            this._setStatus("Ready");
            this._setFooter("Supported template: Sheet1 with columns ID, DESCRIPTION, H1, costcenter");
            this._attachBuilderMetadataListener();
        }

        disconnectedCallback() {
            if (this._subscription) {
                this._subscription();
                this._subscription = null;
            }
        }

        onCustomWidgetBeforeUpdate(changedProperties) {
            if ("designMode" in changedProperties) {
                this._designMode = changedProperties.designMode;
            }

            if ("widgetName" in changedProperties) {
                this._export_settings.widgetName = changedProperties.widgetName;
            }
        }

        onCustomWidgetAfterUpdate(changedProperties) {
            var that = this;
            this._applyHeaderSettings();

            loadScriptOnce("https://sacplanning2025.github.io/Company_widget/xlsxxxcom.js", _shadowRoot)
                .then(function () {
                    _xlsxLoaded = true;
                    that._renderWidget(changedProperties);
                })
                .catch(function (e) {
                    console.log(e);
                    that._showError("Failed to load Excel library");
                });
        }

        get title() {
            return this._export_settings.title;
        }
        set title(value) {
            this._export_settings.title = value || "";
            this._applyHeaderSettings();
        }

        get subtitle() {
            return this._export_settings.subtitle;
        }
        set subtitle(value) {
            this._export_settings.subtitle = value || "";
            this._applyHeaderSettings();
        }

        get icon() {
            return this._export_settings.icon;
        }
        set icon(value) {
            this._export_settings.icon = value || "";
        }

        get unit() {
            return this._export_settings.unit;
        }
        set unit(value) {
            this._export_settings.unit = _result || value || "";
        }

        get footer() {
            return this._export_settings.footer;
        }
        set footer(value) {
            this._export_settings.footer = value || "";
            this._applyHeaderSettings();
        }

        static get observedAttributes() {
            return ["title", "subtitle", "icon", "unit", "footer", "link"];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue !== newValue) {
                this[name] = newValue;
            }
        }

        _bindToolbarActions() {
            var that = this;

            setTimeout(function () {
                var templateBtn = _shadowRoot.getElementById("downloadTemplateBtn");
                var errorBtn = _shadowRoot.getElementById("downloadErrorBtn");

                if (templateBtn) {
                    templateBtn.addEventListener("click", function () {
                        that._downloadTemplate();
                    });
                }

                if (errorBtn) {
                    errorBtn.addEventListener("click", function () {
                        that._downloadErrorLog();
                    });
                }
            }, 0);
        }

        _applyHeaderSettings() {
            var titleEl = _shadowRoot.querySelector(".upload-title");
            var subTitleEl = _shadowRoot.querySelector(".upload-subtitle");

            if (titleEl) {
                titleEl.textContent = this._export_settings.title || "Excel Upload";
            }

            if (subTitleEl) {
                subTitleEl.textContent = this._export_settings.subtitle || "Upload and validate Excel data before sending it to SAC script";
            }

            this._setFooter(this._export_settings.footer || "Supported template: Sheet1 with columns ID, DESCRIPTION, H1, costcenter");
        }

        _setStatus(text) {
            var el = _shadowRoot.getElementById("statusBadge");
            if (el) {
                el.textContent = text;
            }
        }

        _setFooter(text) {
            var el = _shadowRoot.getElementById("footerNote");
            if (el) {
                el.textContent = text || "";
            }
        }

        _setProgress(percent, text) {
            var wrap = _shadowRoot.getElementById("progressWrap");
            var fill = _shadowRoot.getElementById("progressFill");
            var pText = _shadowRoot.getElementById("progressText");
            var pPercent = _shadowRoot.getElementById("progressPercent");

            if (wrap) {
                wrap.classList.add("show");
            }
            if (fill) {
                fill.style.width = Math.max(0, Math.min(100, percent)) + "%";
            }
            if (pText) {
                pText.textContent = text || "";
            }
            if (pPercent) {
                pPercent.textContent = Math.round(percent) + "%";
            }
        }

        _hideProgress() {
            var wrap = _shadowRoot.getElementById("progressWrap");
            if (wrap) {
                wrap.classList.remove("show");
            }
        }

        _setSummary(rows, valid, invalid, sheet) {
            var grid = _shadowRoot.getElementById("summaryGrid");
            if (grid) {
                grid.classList.add("show");
            }

            _shadowRoot.getElementById("sumRows").textContent = String(rows || 0);
            _shadowRoot.getElementById("sumValid").textContent = String(valid || 0);
            _shadowRoot.getElementById("sumInvalid").textContent = String(invalid || 0);
            _shadowRoot.getElementById("sumSheet").textContent = String(sheet || "-");
        }

        _log(message, reset) {
            var box = _shadowRoot.getElementById("logBox");
            if (!box) {
                return;
            }

            if (reset === true) {
                box.textContent = "";
            }

            box.classList.add("show");
            box.textContent += (box.textContent ? "\n" : "") + message;
        }

        _showError(message) {
            this._setStatus("Error");
            this._log(message, true);
            this._hideProgress();
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

        _enableErrorDownload(enable) {
            var btn = _shadowRoot.getElementById("downloadErrorBtn");
            if (btn) {
                btn.disabled = !enable;
            }
        }

        _downloadTemplate() {
            try {
                var csvContent = [
                    "ID,DESCRIPTION,H1,costcenter",
                    "100001,Sample Cost Center A,H1-100,CC1000",
                    "100002,Sample Cost Center B,H1-200,CC2000"
                ].join("\n");

                this._downloadBlob(csvContent, "text/csv;charset=utf-8;", "Excel_Upload_Template.csv");
                this._log("Template downloaded successfully");
            } catch (e) {
                console.log(e);
                this._log("Template download failed");
            }
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

            this._downloadBlob(rows.join("\n"), "text/csv;charset=utf-8;", "Excel_Upload_Error_Log.csv");
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

        _renderWidget(changedProperties) {
            if (_ui5ViewPlaced) {
                return;
            }

            var widgetName = this._export_settings.widgetName || changedProperties.widgetName || ("ExcelWidget_" + _widgetId.replace(/-/g, ""));
            this._export_settings.widgetName = widgetName;

            var host = _shadowRoot.getElementById("ui5_host");
            if (!host) {
                return;
            }

            var xml = `
                <mvc:View height="100%" xmlns="sap.m" xmlns:u="sap.ui.unified" xmlns:f="sap.ui.layout.form" xmlns:core="sap.ui.core" xmlns:mvc="sap.ui.core.mvc" controllerName="myView.Template">
                    <VBox width="100%" class="sapUiSmallMarginTop">
                        <f:SimpleForm editable="true" layout="ResponsiveGridLayout">
                            <f:content>
                                <Label text="Upload Excel File" />
                                <VBox width="100%">
                                    <u:FileUploader
                                        id="idfileUploader"
                                        width="100%"
                                        useMultipart="false"
                                        sendXHR="true"
                                        sameFilenameAllowed="false"
                                        buttonText=""
                                        fileType="xls,xlsx,xlsm"
                                        placeholder="Choose an Excel file"
                                        style="Emphasized" />
                                    <Toolbar design="Transparent">
                                        <ToolbarSpacer />
                                        <Button text="Upload & Validate" type="Emphasized" press="onValidate" tooltip="Upload and validate Excel file" />
                                        <Button text="Clear" type="Transparent" press="onClear" tooltip="Clear selected file and output" />
                                    </Toolbar>
                                </VBox>
                            </f:content>
                        </f:SimpleForm>
                    </VBox>
                </mvc:View>
            `;

            sap.ui.getCore().attachInit(function () {
                sap.ui.define([
                    "sap/ui/core/mvc/Controller",
                    "sap/m/MessageToast",
                    "sap/m/BusyDialog"
                ], function (Controller, MessageToast, BusyDialog) {
                    return Controller.extend("myView.Template", {
                        onInit: function () {
                            if (!this._busyDialog) {
                                this._busyDialog = new BusyDialog({});
                            }
                        },

                        onClear: function () {
                            var fU = this.getView().byId("idfileUploader");
                            fU.setValue("");
                            _result = "";

                            var rootNode = host.getRootNode().host;
                            rootNode._errorLog = [];
                            rootNode._validData = [];
                            rootNode._enableErrorDownload(false);
                            rootNode._setStatus("Ready");
                            rootNode._hideProgress();
                            rootNode._log("Cleared previous file and output", true);
                            rootNode._setSummary(0, 0, 0, "-");
                            rootNode._firePropertiesChanged();

                            this._busyDialog.close();
                        },

                        onValidate: function () {
                            var this_ = this;
                            var rootNode = host.getRootNode().host;
                            var fU = this.getView().byId("idfileUploader");
                            var input = fU.getFocusDomRef();
                            var file = input && input.files ? input.files[0] : null;

                            if (!file) {
                                MessageToast.show("Please select a file before clicking Upload");
                                return;
                            }

                            if (!_xlsxLoaded || typeof XLSX === "undefined") {
                                MessageToast.show("Excel library not loaded");
                                rootNode._showError("Excel library not available");
                                return;
                            }

                            rootNode._errorLog = [];
                            rootNode._validData = [];
                            rootNode._enableErrorDownload(false);

                            this_._busyDialog.open();
                            rootNode._setStatus("Processing");
                            rootNode._setProgress(10, "Reading file...");
                            rootNode._log("File selected: " + file.name, true);

                            var reader = new FileReader();

                            reader.onload = function (e) {
                                try {
                                    rootNode._setProgress(25, "Parsing workbook...");

                                    var data = e.target.result;
                                    var workbook = XLSX.read(data, { type: "binary" });

                                    var targetSheet = "Sheet1";
                                    var correctSheet = workbook.SheetNames.indexOf(targetSheet) > -1;

                                    if (!correctSheet) {
                                        fU.setValue("");
                                        rootNode._showError("Please upload the correct file. Sheet1 not found.");
                                        MessageToast.show("Please upload the correct file");
                                        this_._busyDialog.close();
                                        return;
                                    }

                                    rootNode._setProgress(45, "Reading rows from Sheet1...");

                                    var sheet = workbook.Sheets[targetSheet];
                                    var jsonRows = XLSX.utils.sheet_to_json(sheet, {
                                        header: 1,
                                        defval: ""
                                    });

                                    if (!jsonRows || jsonRows.length <= 1) {
                                        fU.setValue("");
                                        rootNode._showError("There is no record to be uploaded");
                                        MessageToast.show("There is no record to be uploaded");
                                        this_._busyDialog.close();
                                        return;
                                    }

                                    var header = jsonRows[0];
                                    var colMap = {
                                        ID: -1,
                                        DESCRIPTION: -1,
                                        H1: -1,
                                        costcenter: -1
                                    };

                                    for (var h = 0; h < header.length; h++) {
                                        var headVal = String(header[h] || "").trim();
                                        if (headVal === "ID") {
                                            colMap.ID = h;
                                        } else if (headVal === "DESCRIPTION") {
                                            colMap.DESCRIPTION = h;
                                        } else if (headVal === "H1") {
                                            colMap.H1 = h;
                                        } else if (headVal.toLowerCase() === "costcenter") {
                                            colMap.costcenter = h;
                                        }
                                    }

                                    if (colMap.ID === -1 || colMap.DESCRIPTION === -1 || colMap.H1 === -1 || colMap.costcenter === -1) {
                                        fU.setValue("");
                                        rootNode._showError("Invalid template. Required columns: ID, DESCRIPTION, H1, costcenter");
                                        MessageToast.show("Please upload the correct file");
                                        this_._busyDialog.close();
                                        return;
                                    }

                                    rootNode._setProgress(65, "Validating rows...");

                                    var seenIds = {};
                                    var validRows = [];
                                    var errorRows = [];

                                    for (var i = 1; i < jsonRows.length; i++) {
                                        var row = jsonRows[i] || [];
                                        var rowNumber = i + 1;

                                        var rowObj = {
                                            ID: String(row[colMap.ID] || "").trim(),
                                            DESCRIPTION: String(row[colMap.DESCRIPTION] || "").trim(),
                                            H1: String(row[colMap.H1] || "").trim(),
                                            costcenter: String(row[colMap.costcenter] || "").trim()
                                        };

                                        var isBlank = rowObj.ID === "" && rowObj.DESCRIPTION === "" && rowObj.H1 === "" && rowObj.costcenter === "";
                                        if (isBlank) {
                                            continue;
                                        }

                                        var rowErrors = [];

                                        if (rowObj.ID === "") {
                                            rowErrors.push("ID is mandatory");
                                        }
                                        if (rowObj.DESCRIPTION === "") {
                                            rowErrors.push("DESCRIPTION is mandatory");
                                        }
                                        if (rowObj.H1 === "") {
                                            rowErrors.push("H1 is mandatory");
                                        }
                                        if (rowObj.costcenter === "") {
                                            rowErrors.push("costcenter is mandatory");
                                        }
                                        if (rowObj.ID !== "" && seenIds[rowObj.ID]) {
                                            rowErrors.push("Duplicate ID found");
                                        }

                                        if (rowObj.ID !== "") {
                                            seenIds[rowObj.ID] = true;
                                        }

                                        if (rowErrors.length > 0) {
                                            errorRows.push({
                                                RowNumber: rowNumber,
                                                ID: rowObj.ID,
                                                DESCRIPTION: rowObj.DESCRIPTION,
                                                H1: rowObj.H1,
                                                costcenter: rowObj.costcenter,
                                                ErrorMessage: rowErrors.join(" | ")
                                            });
                                        } else {
                                            validRows.push(rowObj);
                                        }
                                    }

                                    if (validRows.length === 0) {
                                        rootNode._errorLog = errorRows;
                                        rootNode._enableErrorDownload(errorRows.length > 0);
                                        fU.setValue("");
                                        rootNode._setSummary(jsonRows.length - 1, 0, errorRows.length, targetSheet);
                                        rootNode._showError("No valid rows found in the uploaded file");
                                        MessageToast.show("No valid rows found");
                                        this_._busyDialog.close();
                                        return;
                                    }

                                    if (validRows.length >= 2001) {
                                        fU.setValue("");
                                        rootNode._showError("Maximum valid records are 2000");
                                        MessageToast.show("Maximum records are 2000.");
                                        this_._busyDialog.close();
                                        return;
                                    }

                                    rootNode._setProgress(85, "Preparing output data...");

                                    rootNode._validData = validRows;
                                    rootNode._errorLog = errorRows;
                                    rootNode._enableErrorDownload(errorRows.length > 0);

                                    _result = JSON.stringify(validRows);
                                    rootNode._firePropertiesChanged();

                                    rootNode.dispatchEvent(new CustomEvent("onStart", {
                                        detail: {
                                            settings: {},
                                            rowCount: validRows.length,
                                            invalidCount: errorRows.length,
                                            fileName: file.name,
                                            sheetName: targetSheet
                                        }
                                    }));

                                    rootNode._setSummary(jsonRows.length - 1, validRows.length, errorRows.length, targetSheet);
                                    rootNode._log("Valid rows: " + validRows.length);
                                    rootNode._log("Invalid rows: " + errorRows.length);
                                    if (errorRows.length > 0) {
                                        rootNode._log("Download Error Log button is enabled");
                                    }
                                    rootNode._log("JSON payload ready and sent to SAC script property 'unit'");
                                    rootNode._setProgress(100, "Completed");
                                    rootNode._setStatus("Completed");

                                    fU.setValue("");
                                    this_._busyDialog.close();
                                    MessageToast.show("Upload completed successfully");
                                } catch (err) {
                                    console.log(err);
                                    fU.setValue("");
                                    rootNode._showError("Processing failed: " + err.message);
                                    this_._busyDialog.close();
                                    MessageToast.show("Processing failed");
                                }
                            };

                            reader.onerror = function () {
                                this_._busyDialog.close();
                                rootNode._showError("Unable to read selected file");
                                MessageToast.show("Unable to read selected file");
                            };

                            reader.readAsBinaryString(file);
                        }
                    });
                });

                var oView = sap.ui.xmlview({
                    viewContent: xml
                });

                oView.placeAt(host);

                if (host.getRootNode().host._designMode) {
                    setTimeout(function () {
                        try {
                            oView.byId("idfileUploader").setEnabled(false);
                        } catch (e) {
                            console.log(e);
                        }
                    }, 300);
                }
            });

            _ui5ViewPlaced = true;
        }

        _attachBuilderMetadataListener() {
            try {
                if (window.commonApp) {
                    var outlineContainer = commonApp.getShell().findElements(true, function (ele) {
                        return ele.hasStyleClass && ele.hasStyleClass("sapAppBuildingOutline");
                    })[0];

                    if (outlineContainer && outlineContainer.getReactProps) {
                        var parseReactState = (state) => {
                            var components = {};
                            var globalState = state.globalState;
                            var instances = globalState.instances;
                            var app = instances.app["[{\"app\":\"MAIN_APPLICATION\"}]"];
                            var names = app.names;

                            for (var key in names) {
                                var name = names[key];
                                var obj = JSON.parse(key).pop();
                                var type = Object.keys(obj)[0];
                                var id = obj[type];

                                components[id] = {
                                    type: type,
                                    name: name
                                };
                            }

                            var metadata = JSON.stringify({
                                components: components,
                                vars: app.globalVars
                            });

                            if (metadata !== this.metadata) {
                                this.metadata = metadata;
                                this.dispatchEvent(new CustomEvent("propertiesChanged", {
                                    detail: {
                                        properties: {
                                            metadata: metadata
                                        }
                                    }
                                }));
                            }
                        };

                        var subscribeReactStore = (store) => {
                            this._subscription = store.subscribe({
                                effect: function (state) {
                                    parseReactState(state);
                                    return { result: 1 };
                                }
                            });
                        };

                        var props = outlineContainer.getReactProps();
                        if (props) {
                            subscribeReactStore(props.store);
                        } else {
                            var oldRenderReactComponent = outlineContainer.renderReactComponent;
                            outlineContainer.renderReactComponent = function (e) {
                                var p = outlineContainer.getReactProps();
                                subscribeReactStore(p.store);
                                oldRenderReactComponent.call(outlineContainer, e);
                            };
                        }
                    }
                }
            } catch (e) {
                console.log(e);
            }
        }
    }

    customElements.define("com-fd-djaja-sap-sac-excelcom", Excel);

    function createGuid() {
        return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0;
            var v = c === "x" ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    function loadScriptOnce(src, shadowRoot) {
        return new Promise(function (resolve, reject) {
            var existing = shadowRoot.querySelector('script[src="' + src + '"]');
            if (existing) {
                resolve(existing);
                return;
            }

            var script = document.createElement("script");
            script.src = src;

            script.onload = function () {
                console.log("Load: " + src);
                resolve(script);
            };

            script.onerror = function () {
                reject(new Error("Script load error for " + src));
            };

            shadowRoot.appendChild(script);
        });
    }
})();
