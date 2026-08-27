(function () {
    "use strict";

    var DEFAULTS = {
        title: "Excel Upload",
        subtitle: "Preview, validate and upload Excel file",
        icon: "",
        unit: "",
        validpayload: "",
        footer: "Supported template: Sheet1 with required business columns",
        templatefilename: "New_Position_Creation_V2.1.xlsm",
        templateurl: "https://raw.githubusercontent.com/sacplanning2025/Company_widget/main/New_Position_Creation_V2.1.xlsm",
        errorlogfilename: "Excel_Upload_Error_Log.csv",
        requiredcolumns: "ID,DESCRIPTION,H1,COMPANY,COSTCENTER,ASSET_CLASS,CAPITALIZED",
        keycolumn: "ID",
        maxrows: "2000",
        previewrows: "1000",
        allowcsv: "true",
        autovalidate: "true",
        showpreview: "true",
        showlogs: "true",
        stricttemplate: "false",
        validationresult: "true",
        validationerrors: "[]",
        previewcompleted: "false",
        continueenabled: "false",
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
                width:100%;
                max-width:100%;
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
            *{ box-sizing:border-box; }
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
            .header{ margin-bottom:12px; }
            .title{
                font-size:15px;
                font-weight:700;
                color:var(--aps-primary);
                margin-bottom:4px;
                line-height:1.2;
                word-break:break-word;
            }
            .subtitle{
                font-size:11px;
                color:var(--aps-sub);
                line-height:1.4;
                word-break:break-word;
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
            .form-grid{
                display:grid;
                grid-template-columns:110px minmax(0, 1fr);
                gap:8px 10px;
                width:100%;
                align-items:start;
            }
            .label{
                font-size:12px;
                font-weight:600;
                color:var(--aps-label);
                padding-top:9px;
                word-break:break-word;
            }
            .field{
                min-width:0;
                width:100%;
            }
            input, select, textarea{
                display:block;
                width:100%;
                font-family:"72", Arial, Helvetica, sans-serif;
                padding:8px 10px;
                border:1px solid #c7ced4;
                border-radius:6px;
                background:#ffffff;
                font-size:13px;
                color:var(--aps-text);
                outline:none;
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
            .hint{
                margin-top:4px;
                font-size:11px;
                color:var(--aps-sub);
                line-height:1.35;
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
            }
            .btn.primary{
                background:var(--aps-primary);
                color:#fff;
                border-color:var(--aps-primary);
            }
            .hidden-submit{ display:none; }
            @media (max-width: 900px){
                .form-grid{
                    grid-template-columns:1fr;
                    gap:6px 0;
                }
                .label{ padding-top:0; }
            }
        </style>

        <div class="panel">
            <div class="header">
                <div class="title">Advanced Excel Upload Widget</div>
                <div class="subtitle">Configure preview, validation and continue-to-upload behavior</div>
            </div>

            <form id="form" autocomplete="off">
                <div class="layout">
                    <fieldset>
                        <legend>General</legend>
                        <div class="form-grid">
                            <div class="label"><label for="title">Title</label></div>
                            <div class="field"><input id="title" name="title" type="text"></div>

                            <div class="label"><label for="subtitle">Subtitle</label></div>
                            <div class="field"><input id="subtitle" name="subtitle" type="text"></div>

                            <div class="label"><label for="icon">Icon</label></div>
                            <div class="field"><input id="icon" name="icon" type="text"></div>

                            <div class="label"><label for="footer">Footer</label></div>
                            <div class="field"><input id="footer" name="footer" type="text"></div>
                        </div>
                    </fieldset>

                    <fieldset>
                        <legend>Template & Output</legend>
                        <div class="form-grid">
                            <div class="label"><label for="unit">Preview JSON</label></div>
                            <div class="field"><input id="unit" name="unit" type="text"></div>

                            <div class="label"><label for="validpayload">Valid JSON</label></div>
                            <div class="field"><input id="validpayload" name="validpayload" type="text"></div>

                            <div class="label"><label for="templatefilename">Template File</label></div>
                            <div class="field"><input id="templatefilename" name="templatefilename" type="text"></div>

                            <div class="label"><label for="templateurl">Template URL</label></div>
                            <div class="field"><input id="templateurl" name="templateurl" type="text"></div>

                            <div class="label"><label for="errorlogfilename">Error Log</label></div>
                            <div class="field"><input id="errorlogfilename" name="errorlogfilename" type="text"></div>

                            <div class="label"><label for="requiredcolumns">Columns</label></div>
                            <div class="field"><input id="requiredcolumns" name="requiredcolumns" type="text"></div>

                            <div class="label"><label for="keycolumn">Key</label></div>
                            <div class="field"><input id="keycolumn" name="keycolumn" type="text"></div>
                        </div>
                    </fieldset>

                    <fieldset>
                        <legend>Validation & Preview</legend>
                        <div class="form-grid">
                            <div class="label"><label for="maxrows">Max Rows</label></div>
                            <div class="field"><input id="maxrows" name="maxrows" type="number" min="1" step="1"></div>

                            <div class="label"><label for="previewrows">Preview Rows</label></div>
                            <div class="field"><input id="previewrows" name="previewrows" type="number" min="1" step="1"></div>

                            <div class="label"><label for="allowcsv">Allow CSV</label></div>
                            <div class="field">
                                <select id="allowcsv" name="allowcsv">
                                    <option value="true">true</option>
                                    <option value="false">false</option>
                                </select>
                            </div>

                            <div class="label"><label for="autovalidate">Auto Validate</label></div>
                            <div class="field">
                                <select id="autovalidate" name="autovalidate">
                                    <option value="true">true</option>
                                    <option value="false">false</option>
                                </select>
                            </div>

                            <div class="label"><label for="showpreview">Show Preview</label></div>
                            <div class="field">
                                <select id="showpreview" name="showpreview">
                                    <option value="true">true</option>
                                    <option value="false">false</option>
                                </select>
                            </div>

                            <div class="label"><label for="showlogs">Show Logs</label></div>
                            <div class="field">
                                <select id="showlogs" name="showlogs">
                                    <option value="true">true</option>
                                    <option value="false">false</option>
                                </select>
                            </div>

                            <div class="label"><label for="stricttemplate">Strict</label></div>
                            <div class="field">
                                <select id="stricttemplate" name="stricttemplate">
                                    <option value="false">false</option>
                                    <option value="true">true</option>
                                </select>
                            </div>
                        </div>
                    </fieldset>

                    <fieldset>
                        <legend>Runtime State</legend>
                        <div class="form-grid">
                            <div class="label"><label for="validationresult">Result</label></div>
                            <div class="field"><input id="validationresult" name="validationresult" type="text"></div>

                            <div class="label"><label for="validationerrors">Errors</label></div>
                            <div class="field"><textarea id="validationerrors" name="validationerrors"></textarea></div>

                            <div class="label"><label for="previewcompleted">Preview Done</label></div>
                            <div class="field"><input id="previewcompleted" name="previewcompleted" type="text"></div>

                            <div class="label"><label for="continueenabled">Continue</label></div>
                            <div class="field"><input id="continueenabled" name="continueenabled" type="text"></div>

                            <div class="label"><label for="lastevent">Event</label></div>
                            <div class="field"><input id="lastevent" name="lastevent" type="text"></div>

                            <div class="label"><label for="rowcount">Rows</label></div>
                            <div class="field"><input id="rowcount" name="rowcount" type="text" readonly></div>

                            <div class="label"><label for="validcount">Valid</label></div>
                            <div class="field"><input id="validcount" name="validcount" type="text" readonly></div>

                            <div class="label"><label for="invalidcount">Invalid</label></div>
                            <div class="field"><input id="invalidcount" name="invalidcount" type="text" readonly></div>
                        </div>
                    </fieldset>
                </div>

                <button type="submit" class="hidden-submit">Submit</button>
            </form>

            <div class="help">
                SAC validation format example:<br>
                [{"rowIndex":0,"field":"ID","message":"Blank cell not allowed"},{"rowIndex":2,"field":"COSTCENTER","message":"Invalid cost center"}]
            </div>

            <div class="toolbar">
                <div class="chip-row">
                    <div class="chip">Preview first</div>
                    <div class="chip">Validate in SAC</div>
                    <div class="chip">Continue to upload</div>
                </div>
                <div>
                    <button type="button" class="btn" id="resetBtn">Reset Fields</button>
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
            if (e) e.preventDefault();
            this._firePropertiesChanged(this._collectProperties());
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
                if (this[key] === "") this[key] = DEFAULTS[key];
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
                detail: { properties: properties }
            }));
        }

        _getValue(id) {
            var el = this._shadowRoot.getElementById(id);
            return el ? el.value : "";
        }

        _setValue(id, value) {
            var el = this._shadowRoot.getElementById(id);
            if (el) el.value = value === undefined || value === null ? "" : value;
        }

        get title() { return this._getValue("title"); }
        set title(v) { this._setValue("title", v); }

        get subtitle() { return this._getValue("subtitle"); }
        set subtitle(v) { this._setValue("subtitle", v); }

        get icon() { return this._getValue("icon"); }
        set icon(v) { this._setValue("icon", v); }

        get unit() { return this._getValue("unit"); }
        set unit(v) { this._setValue("unit", v); }

        get validpayload() { return this._getValue("validpayload"); }
        set validpayload(v) { this._setValue("validpayload", v); }

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

        get previewcompleted() { return this._getValue("previewcompleted"); }
        set previewcompleted(v) { this._setValue("previewcompleted", v); }

        get continueenabled() { return this._getValue("continueenabled"); }
        set continueenabled(v) { this._setValue("continueenabled", v); }

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
                "title","subtitle","icon","unit","validpayload","footer","templatefilename","templateurl",
                "errorlogfilename","requiredcolumns","keycolumn","maxrows","previewrows","allowcsv",
                "autovalidate","showpreview","showlogs","stricttemplate","validationresult",
                "validationerrors","previewcompleted","continueenabled","lastevent",
                "rowcount","validcount","invalidcount"
            ];
        }

        attributeChangedCallback(name, oldValue, newValue) {
            if (oldValue !== newValue) this[name] = newValue;
        }
    }

    customElements.define("com-fd-djaja-sap-sac-excelcom-aps", ExcelAps);
})();

/*(function () {
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
                width:100%;
                max-width:100%;
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
                word-break:break-word;
            }

            .subtitle{
                font-size:11px;
                color:var(--aps-sub);
                line-height:1.4;
                word-break:break-word;
            }

            .layout{
                display:grid;
                grid-template-columns:1fr;
                gap:12px;
                width:100%;
                max-width:100%;
            }

            fieldset{
                margin:0;
                width:100%;
                max-width:100%;
                min-width:0;
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

            .form-grid{
                display:grid;
                grid-template-columns:88px minmax(0, 1fr);
                gap:8px 10px;
                width:100%;
                max-width:100%;
                align-items:start;
            }

            .label{
                font-size:12px;
                font-weight:600;
                color:var(--aps-label);
                padding-top:9px;
                min-width:0;
                word-break:break-word;
            }

            .field{
                min-width:0;
                width:100%;
                max-width:100%;
                overflow:hidden;
            }

            input, select, textarea{
                display:block;
                width:100%;
                max-width:100%;
                min-width:0;
                font-family:"72", Arial, Helvetica, sans-serif;
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
                min-width:0;
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

            .hidden-submit{
                display:none;
            }

            @media (max-width: 900px){
                .form-grid{
                    grid-template-columns:1fr;
                    gap:6px 0;
                }

                .label{
                    padding-top:0;
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
                        <div class="form-grid">
                            <div class="label"><label for="title">Title</label></div>
                            <div class="field">
                                <input id="title" name="title" type="text" placeholder="Excel Upload">
                                <div class="hint">Widget header title shown in runtime.</div>
                            </div>

                            <div class="label"><label for="subtitle">Subtitle</label></div>
                            <div class="field">
                                <input id="subtitle" name="subtitle" type="text" placeholder="Upload and validate Excel file">
                                <div class="hint">Short helper text below the main title.</div>
                            </div>

                            <div class="label"><label for="icon">Icon</label></div>
                            <div class="field">
                                <input id="icon" name="icon" type="text" placeholder="Optional icon">
                                <div class="hint">Optional future-use icon name or semantic marker.</div>
                            </div>

                            <div class="label"><label for="footer">Footer</label></div>
                            <div class="field">
                                <input id="footer" name="footer" type="text" placeholder="Footer note">
                                <div class="hint">Shown at the bottom of the runtime widget.</div>
                            </div>
                        </div>
                    </fieldset>

                    <fieldset>
                        <legend>Template & Output</legend>
                        <div class="form-grid">
                            <div class="label"><label for="unit">Output</label></div>
                            <div class="field">
                                <input id="unit" name="unit" type="text" placeholder="Payload output property">
                                <div class="hint">Usually leave as configured by widget runtime integration.</div>
                            </div>

                            <div class="label"><label for="templatefilename">Template File</label></div>
                            <div class="field">
                                <input id="templatefilename" name="templatefilename" type="text" placeholder="New_Position_Creation_V2.1.xlsm">
                                <div class="hint">Downloaded filename for the template button.</div>
                            </div>

                            <div class="label"><label for="templateurl">Template URL</label></div>
                            <div class="field">
                                <input id="templateurl" name="templateurl" type="text" placeholder="https://...">
                                <div class="hint">Public URL, SharePoint link, or SAC file link.</div>
                            </div>

                            <div class="label"><label for="errorlogfilename">Error Log</label></div>
                            <div class="field">
                                <input id="errorlogfilename" name="errorlogfilename" type="text" placeholder="Excel_Upload_Error_Log.csv">
                                <div class="hint">Name of the downloadable validation error file.</div>
                            </div>

                            <div class="label"><label for="requiredcolumns">Columns</label></div>
                            <div class="field">
                                <input id="requiredcolumns" name="requiredcolumns" type="text" placeholder="ID,DESCRIPTION,H1,COMPANY,COSTCENTER,ASSET_CLASS,CAPITALIZED">
                                <div class="hint">Comma-separated business columns expected in upload.</div>
                            </div>

                            <div class="label"><label for="keycolumn">Key</label></div>
                            <div class="field">
                                <input id="keycolumn" name="keycolumn" type="text" placeholder="ID">
                                <div class="hint">Used for duplicate detection inside the uploaded file.</div>
                            </div>
                        </div>
                    </fieldset>

                    <fieldset>
                        <legend>Validation & Preview</legend>
                        <div class="form-grid">
                            <div class="label"><label for="maxrows">Max Rows</label></div>
                            <div class="field">
                                <input id="maxrows" name="maxrows" type="number" min="1" step="1" placeholder="2000">
                                <div class="hint">Maximum allowed upload rows.</div>
                            </div>

                            <div class="label"><label for="previewrows">Preview</label></div>
                            <div class="field">
                                <input id="previewrows" name="previewrows" type="number" min="1" step="1" placeholder="300">
                                <div class="hint">Maximum rows rendered in the runtime preview table.</div>
                            </div>

                            <div class="label"><label for="allowcsv">CSV</label></div>
                            <div class="field">
                                <select id="allowcsv" name="allowcsv">
                                    <option value="true">true</option>
                                    <option value="false">false</option>
                                </select>
                            </div>

                            <div class="label"><label for="autovalidate">Validate</label></div>
                            <div class="field">
                                <select id="autovalidate" name="autovalidate">
                                    <option value="true">true</option>
                                    <option value="false">false</option>
                                </select>
                            </div>

                            <div class="label"><label for="showpreview">Show Prev</label></div>
                            <div class="field">
                                <select id="showpreview" name="showpreview">
                                    <option value="true">true</option>
                                    <option value="false">false</option>
                                </select>
                            </div>

                            <div class="label"><label for="showlogs">Show Logs</label></div>
                            <div class="field">
                                <select id="showlogs" name="showlogs">
                                    <option value="true">true</option>
                                    <option value="false">false</option>
                                </select>
                            </div>

                            <div class="label"><label for="stricttemplate">Strict</label></div>
                            <div class="field">
                                <select id="stricttemplate" name="stricttemplate">
                                    <option value="false">false</option>
                                    <option value="true">true</option>
                                </select>
                                <div class="hint">If true, runtime can enforce stricter template behavior.</div>
                            </div>
                        </div>
                    </fieldset>

                    <fieldset>
                        <legend>Runtime State (Read / Bind)</legend>
                        <div class="form-grid">
                            <div class="label"><label for="validationresult">Result</label></div>
                            <div class="field">
                                <input id="validationresult" name="validationresult" type="text" placeholder="true / false">
                                <div class="hint">Can be updated by SAC script to control final status.</div>
                            </div>

                            <div class="label"><label for="validationerrors">Errors</label></div>
                            <div class="field">
                                <textarea id="validationerrors" name="validationerrors" rows="4" placeholder='[{"rowIndex":0,"field":"ID","message":"Blank cell not allowed"}]'></textarea>
                                <div class="hint">JSON array used by runtime preview highlighting.</div>
                            </div>

                            <div class="label"><label for="lastevent">Event</label></div>
                            <div class="field">
                                <input id="lastevent" name="lastevent" type="text" placeholder="uploadCompleted / clear">
                            </div>

                            <div class="label"><label for="rowcount">Rows</label></div>
                            <div class="field">
                                <input id="rowcount" name="rowcount" type="text" readonly>
                            </div>

                            <div class="label"><label for="validcount">Valid</label></div>
                            <div class="field">
                                <input id="validcount" name="validcount" type="text" readonly>
                            </div>

                            <div class="label"><label for="invalidcount">Invalid</label></div>
                            <div class="field">
                                <input id="invalidcount" name="invalidcount" type="text" readonly>
                            </div>
                        </div>
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
*/
