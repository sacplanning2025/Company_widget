/*(function () {
    "use strict";

    var tmpl = document.createElement("template");
    tmpl.innerHTML = `
        <style>
            :host{
                display:block;
                font-family:"72", Arial, Helvetica, sans-serif;
                color:#1f2d3d;
            }

            .panel{
                border:1px solid #d9d9d9;
                border-radius:10px;
                background:#ffffff;
                padding:14px;
                box-shadow:0 2px 8px rgba(0,0,0,0.06);
            }

            .header{
                margin-bottom:12px;
            }

            .title{
                font-size:15px;
                font-weight:700;
                color:#0a6ed1;
                margin-bottom:4px;
            }

            .subtitle{
                font-size:12px;
                color:#6a6d70;
            }

            fieldset{
                margin:0;
                border:1px solid #d9d9d9;
                border-radius:8px;
                padding:12px;
                background:#fafbfc;
            }

            legend{
                padding:0 6px;
                font-size:12px;
                font-weight:700;
                color:#354a5f;
            }

            table{
                width:100%;
                border-collapse:collapse;
            }

            td{
                padding:6px 4px;
                vertical-align:middle;
            }

            .label{
                width:120px;
                font-size:12px;
                font-weight:600;
                color:#354a5f;
            }

            input{
                font-family:"72", Arial, Helvetica, sans-serif;
                width:100%;
                padding:8px 10px;
                box-sizing:border-box;
                border:1px solid #c7ced4;
                border-radius:6px;
                background:#ffffff;
                font-size:13px;
                color:#1f2d3d;
                outline:none;
            }

            input:focus{
                border-color:#0a6ed1;
                box-shadow:0 0 0 2px rgba(10,110,209,0.12);
            }

            .help{
                margin-top:12px;
                padding:10px 12px;
                border-radius:8px;
                background:#f5faff;
                border:1px solid #d0e7ff;
                font-size:12px;
                color:#354a5f;
                line-height:1.5;
            }
        </style>

        <div class="panel">
            <div class="header">
                <div class="title">Excel Upload Widget Settings</div>
                <div class="subtitle">Configure runtime title, subtitle, payload field and footer information</div>
            </div>

            <form id="form" autocomplete="off">
                <fieldset>
                    <legend>General</legend>
                    <table>
                        <tr>
                            <td class="label"><label for="title">Title</label></td>
                            <td><input id="title" name="title" type="text" placeholder="Excel Upload"></td>
                        </tr>
                        <tr>
                            <td class="label"><label for="subtitle">Sub Title</label></td>
                            <td><input id="subtitle" name="subtitle" type="text" placeholder="Upload and validate Excel file"></td>
                        </tr>
                        <tr>
                            <td class="label"><label for="icon">Icon</label></td>
                            <td><input id="icon" name="icon" type="text" placeholder="Optional icon name"></td>
                        </tr>
                        <tr>
                            <td class="label"><label for="unit">Unit</label></td>
                            <td><input id="unit" name="unit" type="text" placeholder="Payload output property"></td>
                        </tr>
                        <tr>
                            <td class="label"><label for="footer">Footer</label></td>
                            <td><input id="footer" name="footer" type="text" placeholder="Template guidance or note"></td>
                        </tr>
                    </table>
                </fieldset>
                <button type="submit" hidden>Submit</button>
            </form>

            <div class="help">
                Recommended footer example: Supported template: Sheet1 with columns ID, DESCRIPTION, H1, costcenter
            </div>
        </div>
    `;

    class ExcelAps extends HTMLElement {
        constructor() {
            super();
            this._shadowRoot = this.attachShadow({ mode: "open" });
            this._shadowRoot.appendChild(tmpl.content.cloneNode(true));

            var form = this._shadowRoot.getElementById("form");
            form.addEventListener("submit", this._submit.bind(this));
            form.addEventListener("change", this._change.bind(this));
            form.addEventListener("input", this._change.bind(this));
        }

        _submit(e) {
            e.preventDefault();

            var properties = {};
            for (var i = 0; i < ExcelAps.observedAttributes.length; i++) {
                var name = ExcelAps.observedAttributes[i];
                properties[name] = this[name];
            }

            this._firePropertiesChanged(properties);
            return false;
        }

        _change(e) {
            this._changeProperty(e.target.name);
        }

        _changeProperty(name) {
            var properties = {};
            properties[name] = this[name];
            this._firePropertiesChanged(properties);
        }

        _firePropertiesChanged(properties) {
            this.dispatchEvent(new CustomEvent("propertiesChanged", {
                detail: {
                    properties: properties
                }
            }));
        }

        get title() {
            return this._getValue("title");
        }
        set title(value) {
            this._setValue("title", value);
        }

        get subtitle() {
            return this._getValue("subtitle");
        }
        set subtitle(value) {
            this._setValue("subtitle", value);
        }

        get icon() {
            return this._getValue("icon");
        }
        set icon(value) {
            this._setValue("icon", value);
        }

        get unit() {
            return this._getValue("unit");
        }
        set unit(value) {
            this._setValue("unit", value);
        }

        get footer() {
            return this._getValue("footer");
        }
        set footer(value) {
            this._setValue("footer", value);
        }

        _getValue(id) {
            var el = this._shadowRoot.getElementById(id);
            return el ? el.value : "";
        }

        _setValue(id, value) {
            var el = this._shadowRoot.getElementById(id);
            if (el) {
                el.value = value || "";
            }
        }

        static get observedAttributes() {
            return ["title", "subtitle", "icon", "unit", "footer"];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue !== newValue) {
                this[name] = newValue;
            }
        }
    }

    customElements.define("com-fd-djaja-sap-sac-excelcom-aps", ExcelAps);
})();*/

(function () {
    "use strict";

    var DEFAULTS = {
        title: "Excel Upload",
        subtitle: "Upload and validate Excel file",
        icon: "",
        unit: "",
        footer: "Supported template: Sheet1 with required business columns",
        templatefilename: "New_Position_Creation_V2.1.xlsm",
        templateurl: "https://raw.githubusercontent.com/sacplanning2025/Company_widget/main/New_Position_Creation_V2.1.xlsm",
        errorlogfilename: "Excel_Upload_Error_Log.csv",
        requiredcolumns: "ID,DESCRIPTION,H1,COMPANY,COSTCENTER,ASSET_CLASS,CAPITALIZED",
        keycolumn: "ID",
        maxrows: "2000",
        previewrows: "300",
        allowcsv: "true",
        autovalidate: "true",
        showpreview: "true",
        showlogs: "true",
        stricttemplate: "false",
        validationresult: "true",
        validationerrors: "[]",
        lastevent: "",
        rowcount: "",
        validcount: "",
        invalidcount: ""
    };

    var tmpl = document.createElement("template");
    tmpl.innerHTML = `
        <style>
            :host{
                display:block;
                font-family:"72", Arial, Helvetica, sans-serif;
                color:#1f2d3d;
                --aps-primary:#0a6ed1;
                --aps-primary-soft:#f4f9ff;
                --aps-border:#d9d9d9;
                --aps-border-soft:#e8edf2;
                --aps-bg:#ffffff;
                --aps-bg-soft:#fafbfc;
                --aps-bg-info:#f5faff;
                --aps-info-border:#d0e7ff;
                --aps-text:#1f2d3d;
                --aps-sub:#6a6d70;
                --aps-label:#354a5f;
                --aps-shadow:0 6px 18px rgba(0,0,0,0.06);
                --aps-radius:12px;
            }

            *{
                box-sizing:border-box;
            }

            :host,
            .panel,
            .layout,
            fieldset,
            table,
            tr,
            td,
            .field,
            .field-wrap{
                min-width:0;
            }

            .panel{
                width:100%;
                max-width:100%;
                overflow:hidden;
                border:1px solid var(--aps-border);
                border-radius:var(--aps-radius);
                background:var(--aps-bg);
                padding:12px;
                box-shadow:var(--aps-shadow);
            }

            .header{
                margin-bottom:12px;
            }

            .title{
                font-size:15px;
                font-weight:700;
                color:var(--aps-primary);
                margin-bottom:4px;
                line-height:1.2;
            }

            .subtitle{
                font-size:11px;
                color:var(--aps-sub);
                line-height:1.4;
            }

            .layout{
                display:grid;
                grid-template-columns:1fr;
                gap:12px;
                width:100%;
            }

            fieldset{
                margin:0;
                width:100%;
                max-width:100%;
                border:1px solid var(--aps-border);
                border-radius:10px;
                padding:10px;
                background:var(--aps-bg-soft);
                overflow:hidden;
            }

            legend{
                padding:0 6px;
                font-size:12px;
                font-weight:700;
                color:var(--aps-label);
            }

            table{
                width:100%;
                max-width:100%;
                border-collapse:collapse;
                table-layout:fixed;
            }

            tr{
                vertical-align:top;
            }

            td{
                padding:6px 4px;
                vertical-align:top;
            }

            .label{
                width:88px;
                min-width:88px;
                max-width:88px;
                font-size:12px;
                font-weight:600;
                color:var(--aps-label);
                padding-top:11px;
                word-break:break-word;
            }

            .field{
                width:auto;
            }

            .field-wrap{
                width:100%;
                max-width:100%;
                overflow:hidden;
            }

            input, select, textarea{
                display:block;
                font-family:"72", Arial, Helvetica, sans-serif;
                width:100%;
                max-width:100%;
                min-width:0;
                padding:8px 10px;
                border:1px solid #c7ced4;
                border-radius:6px;
                background:#ffffff;
                font-size:13px;
                color:var(--aps-text);
                outline:none;
                transition:border-color 0.15s ease, box-shadow 0.15s ease, background 0.15s ease;
            }

            input:focus, select:focus, textarea:focus{
                border-color:var(--aps-primary);
                box-shadow:0 0 0 2px rgba(10,110,209,0.12);
            }

            input[readonly]{
                background:#f8f9fa;
                color:#5f6b7a;
            }

            textarea{
                resize:vertical;
                min-height:86px;
                line-height:1.4;
            }

            .wide-input,
            .super-wide-input{
                width:100%;
                max-width:100%;
                min-width:0;
            }

            .hint{
                margin-top:4px;
                font-size:11px;
                color:var(--aps-sub);
                line-height:1.35;
                word-break:break-word;
            }

            .help{
                margin-top:12px;
                padding:10px 12px;
                border-radius:8px;
                background:var(--aps-bg-info);
                border:1px solid var(--aps-info-border);
                font-size:11px;
                color:var(--aps-label);
                line-height:1.5;
                word-break:break-word;
            }

            .section-note{
                font-size:11px;
                color:var(--aps-sub);
                margin-top:8px;
                line-height:1.4;
                word-break:break-word;
            }

            .toolbar{
                display:flex;
                justify-content:space-between;
                align-items:center;
                gap:10px;
                flex-wrap:wrap;
                margin-top:12px;
                padding-top:10px;
                border-top:1px solid var(--aps-border-soft);
            }

            .chip-row{
                display:flex;
                gap:8px;
                flex-wrap:wrap;
            }

            .chip{
                display:inline-flex;
                align-items:center;
                min-height:24px;
                padding:4px 8px;
                border:1px solid var(--aps-border);
                border-radius:999px;
                background:#fff;
                font-size:10px;
                font-weight:600;
                color:#415466;
            }

            .btn{
                appearance:none;
                border:1px solid #c7ced4;
                background:#fff;
                color:var(--aps-text);
                border-radius:8px;
                padding:8px 12px;
                font-size:12px;
                font-weight:700;
                cursor:pointer;
                transition:all 0.2s ease;
            }

            .btn:hover{
                border-color:var(--aps-primary);
                color:var(--aps-primary);
                background:var(--aps-primary-soft);
            }

            .btn.primary{
                background:var(--aps-primary);
                color:#fff;
                border-color:var(--aps-primary);
            }

            .btn.primary:hover{
                background:#085caf;
                border-color:#085caf;
                color:#fff;
            }

            .btn.secondary{
                background:#fff;
            }

            .hidden-submit{
                display:none;
            }

            @media (max-width: 900px){
                .label{
                    width:78px;
                    min-width:78px;
                    max-width:78px;
                }

                .wide-input,
                .super-wide-input{
                    min-width:0;
                    max-width:100%;
                }
            }
        </style>

        <div class="panel">
            <div class="header">
                <div class="title">Advanced Excel Upload Widget</div>
                <div class="subtitle">Configure title, template, validation, preview and runtime behavior</div>
            </div>

            <form id="form" autocomplete="off">
                <div class="layout">
                    <fieldset>
                        <legend>General</legend>
                        <table>
                            <tr>
                                <td class="label"><label for="title">Title</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="title" name="title" class="super-wide-input" type="text" placeholder="Excel Upload">
                                        <div class="hint">Widget header title shown in runtime.</div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="subtitle">Subtitle</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="subtitle" name="subtitle" class="super-wide-input" type="text" placeholder="Upload and validate Excel file">
                                        <div class="hint">Short helper text below the main title.</div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="icon">Icon</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="icon" name="icon" class="wide-input" type="text" placeholder="Optional icon">
                                        <div class="hint">Optional future-use icon name or semantic marker.</div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="footer">Footer</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="footer" name="footer" class="super-wide-input" type="text" placeholder="Footer note">
                                        <div class="hint">Shown at the bottom of the runtime widget.</div>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </fieldset>

                    <fieldset>
                        <legend>Template & Output</legend>
                        <table>
                            <tr>
                                <td class="label"><label for="unit">Output</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="unit" name="unit" class="super-wide-input" type="text" placeholder="Payload output property">
                                        <div class="hint">Usually leave as configured by widget runtime integration.</div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="templatefilename">Template File</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="templatefilename" name="templatefilename" class="super-wide-input" type="text" placeholder="New_Position_Creation_V2.1.xlsm">
                                        <div class="hint">Downloaded filename for the template button.</div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="templateurl">Template URL</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="templateurl" name="templateurl" class="super-wide-input" type="text" placeholder="https://...">
                                        <div class="hint">Public URL, SharePoint link, or SAC file link.</div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="errorlogfilename">Error Log File</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="errorlogfilename" name="errorlogfilename" class="super-wide-input" type="text" placeholder="Excel_Upload_Error_Log.csv">
                                        <div class="hint">Name of the downloadable validation error file.</div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="requiredcolumns">Required Columns</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="requiredcolumns" name="requiredcolumns" class="super-wide-input" type="text" placeholder="ID,DESCRIPTION,H1,COMPANY,COSTCENTER,ASSET_CLASS,CAPITALIZED">
                                        <div class="hint">Comma-separated business columns expected in upload.</div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="keycolumn">Key Column</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="keycolumn" name="keycolumn" class="wide-input" type="text" placeholder="ID">
                                        <div class="hint">Used for duplicate detection inside the uploaded file.</div>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </fieldset>

                    <fieldset>
                        <legend>Validation & Preview</legend>
                        <table>
                            <tr>
                                <td class="label"><label for="maxrows">Max Rows</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="maxrows" name="maxrows" class="wide-input" type="number" min="1" step="1" placeholder="2000">
                                        <div class="hint">Maximum allowed upload rows.</div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="previewrows">Preview Rows</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="previewrows" name="previewrows" class="wide-input" type="number" min="1" step="1" placeholder="300">
                                        <div class="hint">Maximum rows rendered in the runtime preview table.</div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="allowcsv">Allow CSV</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <select id="allowcsv" name="allowcsv" class="wide-input">
                                            <option value="true">true</option>
                                            <option value="false">false</option>
                                        </select>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="autovalidate">Auto Validate</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <select id="autovalidate" name="autovalidate" class="wide-input">
                                            <option value="true">true</option>
                                            <option value="false">false</option>
                                        </select>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="showpreview">Show Preview</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <select id="showpreview" name="showpreview" class="wide-input">
                                            <option value="true">true</option>
                                            <option value="false">false</option>
                                        </select>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="showlogs">Show Logs</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <select id="showlogs" name="showlogs" class="wide-input">
                                            <option value="true">true</option>
                                            <option value="false">false</option>
                                        </select>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="stricttemplate">Strict Template</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <select id="stricttemplate" name="stricttemplate" class="wide-input">
                                            <option value="false">false</option>
                                            <option value="true">true</option>
                                        </select>
                                        <div class="hint">If true, runtime can enforce stricter template behavior.</div>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </fieldset>

                    <fieldset>
                        <legend>Runtime State (Read / Bind)</legend>
                        <table>
                            <tr>
                                <td class="label"><label for="validationresult">Validation Result</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="validationresult" name="validationresult" class="wide-input" type="text" placeholder="true / false">
                                        <div class="hint">Can be updated by SAC script to control final status.</div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="validationerrors">Validation Errors</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <textarea id="validationerrors" name="validationerrors" class="super-wide-input" rows="4" placeholder='[{"rowIndex":0,"field":"ID","message":"Blank cell not allowed"}]'></textarea>
                                        <div class="hint">JSON array used by runtime preview highlighting.</div>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="lastevent">Last Event</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="lastevent" name="lastevent" class="wide-input" type="text" placeholder="uploadCompleted / clear">
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="rowcount">Row Count</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="rowcount" name="rowcount" class="wide-input" type="text" readonly>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="validcount">Valid Count</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="validcount" name="validcount" class="wide-input" type="text" readonly>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="label"><label for="invalidcount">Invalid Count</label></td>
                                <td class="field">
                                    <div class="field-wrap">
                                        <input id="invalidcount" name="invalidcount" class="wide-input" type="text" readonly>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </fieldset>
                </div>

                <button type="submit" class="hidden-submit">Submit</button>
            </form>

            <div class="help">
                SAC validation format example:<br>
                [{"rowIndex":0,"field":"ID","message":"Blank cell not allowed"},{"rowIndex":2,"field":"COSTCENTER","message":"Invalid cost center"}]
            </div>

            <div class="section-note">
                Tip: validationresult and validationerrors can be updated from SAC script so the runtime widget highlights row and cell errors correctly.
            </div>

            <div class="toolbar">
                <div class="chip-row">
                    <div class="chip">Supports XLS / XLSX / XLSM</div>
                    <div class="chip">CSV optional</div>
                    <div class="chip">Preview driven validation</div>
                </div>
                <div>
                    <button type="button" class="btn secondary" id="resetBtn">Reset Fields</button>
                    <button type="button" class="btn primary" id="applyBtn">Apply Changes</button>
                </div>
            </div>
        </div>
    `;

    class ExcelAps extends HTMLElement {
        constructor() {
            super();
            this._shadowRoot = this.attachShadow({ mode: "open" });
            this._shadowRoot.appendChild(tmpl.content.cloneNode(true));

            var form = this._shadowRoot.getElementById("form");
            form.addEventListener("submit", this._submit.bind(this));
            form.addEventListener("change", this._change.bind(this));
            form.addEventListener("input", this._change.bind(this));

            this._shadowRoot.getElementById("applyBtn").addEventListener("click", this._submit.bind(this));
            this._shadowRoot.getElementById("resetBtn").addEventListener("click", this._reset.bind(this));
        }

        connectedCallback() {
            this._applyDefaultsIfEmpty();
        }

        _submit(e) {
            if (e) {
                e.preventDefault();
            }

            var properties = {};
            for (var i = 0; i < ExcelAps.observedAttributes.length; i++) {
                var name = ExcelAps.observedAttributes[i];
                properties[name] = this[name];
            }

            this._firePropertiesChanged(properties);
            return false;
        }

        _change(e) {
            this._changeProperty(e.target.name);
        }

        _changeProperty(name) {
            var properties = {};
            properties[name] = this[name];
            this._firePropertiesChanged(properties);
        }

        _reset() {
            var keys = Object.keys(DEFAULTS);
            for (var i = 0; i < keys.length; i++) {
                this[keys[i]] = DEFAULTS[keys[i]];
            }
            this._firePropertiesChanged(this._collectProperties());
        }

        _applyDefaultsIfEmpty() {
            var keys = Object.keys(DEFAULTS);
            for (var i = 0; i < keys.length; i++) {
                var key = keys[i];
                if (this[key] === "") {
                    this[key] = DEFAULTS[key];
                }
            }
        }

        _collectProperties() {
            var properties = {};
            for (var i = 0; i < ExcelAps.observedAttributes.length; i++) {
                var name = ExcelAps.observedAttributes[i];
                properties[name] = this[name];
            }
            return properties;
        }

        _firePropertiesChanged(properties) {
            this.dispatchEvent(new CustomEvent("propertiesChanged", {
                detail: {
                    properties: properties
                }
            }));
        }

        _getValue(id) {
            var el = this._shadowRoot.getElementById(id);
            return el ? el.value : "";
        }

        _setValue(id, value) {
            var el = this._shadowRoot.getElementById(id);
            if (el) {
                el.value = value === undefined || value === null ? "" : value;
            }
        }

        get title() { return this._getValue("title"); }
        set title(v) { this._setValue("title", v); }

        get subtitle() { return this._getValue("subtitle"); }
        set subtitle(v) { this._setValue("subtitle", v); }

        get icon() { return this._getValue("icon"); }
        set icon(v) { this._setValue("icon", v); }

        get unit() { return this._getValue("unit"); }
        set unit(v) { this._setValue("unit", v); }

        get footer() { return this._getValue("footer"); }
        set footer(v) { this._setValue("footer", v); }

        get templatefilename() { return this._getValue("templatefilename"); }
        set templatefilename(v) { this._setValue("templatefilename", v); }

        get templateurl() { return this._getValue("templateurl"); }
        set templateurl(v) { this._setValue("templateurl", v); }

        get errorlogfilename() { return this._getValue("errorlogfilename"); }
        set errorlogfilename(v) { this._setValue("errorlogfilename", v); }

        get requiredcolumns() { return this._getValue("requiredcolumns"); }
        set requiredcolumns(v) { this._setValue("requiredcolumns", v); }

        get keycolumn() { return this._getValue("keycolumn"); }
        set keycolumn(v) { this._setValue("keycolumn", v); }

        get maxrows() { return this._getValue("maxrows"); }
        set maxrows(v) { this._setValue("maxrows", v); }

        get previewrows() { return this._getValue("previewrows"); }
        set previewrows(v) { this._setValue("previewrows", v); }

        get allowcsv() { return this._getValue("allowcsv"); }
        set allowcsv(v) { this._setValue("allowcsv", v); }

        get autovalidate() { return this._getValue("autovalidate"); }
        set autovalidate(v) { this._setValue("autovalidate", v); }

        get showpreview() { return this._getValue("showpreview"); }
        set showpreview(v) { this._setValue("showpreview", v); }

        get showlogs() { return this._getValue("showlogs"); }
        set showlogs(v) { this._setValue("showlogs", v); }

        get stricttemplate() { return this._getValue("stricttemplate"); }
        set stricttemplate(v) { this._setValue("stricttemplate", v); }

        get validationresult() { return this._getValue("validationresult"); }
        set validationresult(v) { this._setValue("validationresult", v); }

        get validationerrors() { return this._getValue("validationerrors"); }
        set validationerrors(v) { this._setValue("validationerrors", v); }

        get lastevent() { return this._getValue("lastevent"); }
        set lastevent(v) { this._setValue("lastevent", v); }

        get rowcount() { return this._getValue("rowcount"); }
        set rowcount(v) { this._setValue("rowcount", v); }

        get validcount() { return this._getValue("validcount"); }
        set validcount(v) { this._setValue("validcount", v); }

        get invalidcount() { return this._getValue("invalidcount"); }
        set invalidcount(v) { this._setValue("invalidcount", v); }

        static get observedAttributes() {
            return [
                "title",
                "subtitle",
                "icon",
                "unit",
                "footer",
                "templatefilename",
                "templateurl",
                "errorlogfilename",
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
                "rowcount",
                "validcount",
                "invalidcount"
            ];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue !== newValue) {
                this[name] = newValue;
            }
        }
    }

    customElements.define("com-fd-djaja-sap-sac-excelcom-aps", ExcelAps);
})();
