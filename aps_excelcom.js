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
                border-radius:12px;
                background:#ffffff;
                padding:16px;
                box-shadow:0 4px 14px rgba(0,0,0,0.06);
            }

            .header{
                margin-bottom:14px;
            }

            .title{
                font-size:16px;
                font-weight:700;
                color:#0a6ed1;
                margin-bottom:4px;
            }

            .subtitle{
                font-size:12px;
                color:#6a6d70;
            }

            fieldset{
                margin:0 0 14px 0;
                border:1px solid #d9d9d9;
                border-radius:10px;
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
                padding:7px 4px;
                vertical-align:middle;
            }

            .label{
                width:180px;
                font-size:12px;
                font-weight:600;
                color:#354a5f;
            }

            input, select{
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

            input:focus, select:focus{
                border-color:#0a6ed1;
                box-shadow:0 0 0 2px rgba(10,110,209,0.12);
            }

            .help{
                margin-top:12px;
                padding:12px;
                border-radius:8px;
                background:#f5faff;
                border:1px solid #d0e7ff;
                font-size:12px;
                color:#354a5f;
                line-height:1.5;
            }

            .section-note{
                font-size:11px;
                color:#6a6d70;
                margin-top:8px;
            }
        </style>

        <div class="panel">
            <div class="header">
                <div class="title">Advanced Excel Upload Widget</div>
                <div class="subtitle">Configure title, template, validation, preview and runtime behavior</div>
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
                            <td class="label"><label for="subtitle">Subtitle</label></td>
                            <td><input id="subtitle" name="subtitle" type="text" placeholder="Upload and validate Excel file"></td>
                        </tr>
                        <tr>
                            <td class="label"><label for="icon">Icon</label></td>
                            <td><input id="icon" name="icon" type="text" placeholder="Optional icon"></td>
                        </tr>
                        <tr>
                            <td class="label"><label for="footer">Footer</label></td>
                            <td><input id="footer" name="footer" type="text" placeholder="Footer note"></td>
                        </tr>
                    </table>
                </fieldset>

                <fieldset>
                    <legend>Template & Output</legend>
                    <table>
                        <tr>
                            <td class="label"><label for="unit">Output Property</label></td>
                            <td><input id="unit" name="unit" type="text" placeholder="Payload output property"></td>
                        </tr>
                        <tr>
                            <td class="label"><label for="templatefilename">Template Filename</label></td>
                            <td><input id="templatefilename" name="templatefilename" type="text" placeholder="New_Position_Creation_V2.1.xlsm"></td>
                        </tr>
                        <tr>
                            <td class="label"><label for="templateurl">Template URL</label></td>
                            <td><input id="templateurl" name="templateurl" type="text" placeholder="https://..."></td>
                        </tr>
                        <tr>
                            <td class="label"><label for="errorlogfilename">Error Log Filename</label></td>
                            <td><input id="errorlogfilename" name="errorlogfilename" type="text" placeholder="Excel_Upload_Error_Log.csv"></td>
                        </tr>
                        <tr>
                            <td class="label"><label for="requiredcolumns">Required Columns</label></td>
                            <td><input id="requiredcolumns" name="requiredcolumns" type="text" placeholder="ID,DESCRIPTION,H1,costcenter"></td>
                        </tr>
                    </table>
                </fieldset>

                <fieldset>
                    <legend>Validation & Preview</legend>
                    <table>
                        <tr>
                            <td class="label"><label for="maxrows">Max Rows</label></td>
                            <td><input id="maxrows" name="maxrows" type="number" min="1" step="1" placeholder="2000"></td>
                        </tr>
                        <tr>
                            <td class="label"><label for="previewrows">Preview Rows</label></td>
                            <td><input id="previewrows" name="previewrows" type="number" min="1" step="1" placeholder="50"></td>
                        </tr>
                        <tr>
                            <td class="label"><label for="allowcsv">Allow CSV</label></td>
                            <td>
                                <select id="allowcsv" name="allowcsv">
                                    <option value="true">true</option>
                                    <option value="false">false</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <td class="label"><label for="autovalidate">Auto Validate</label></td>
                            <td>
                                <select id="autovalidate" name="autovalidate">
                                    <option value="true">true</option>
                                    <option value="false">false</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <td class="label"><label for="showpreview">Show Preview</label></td>
                            <td>
                                <select id="showpreview" name="showpreview">
                                    <option value="true">true</option>
                                    <option value="false">false</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <td class="label"><label for="showlogs">Show Logs</label></td>
                            <td>
                                <select id="showlogs" name="showlogs">
                                    <option value="true">true</option>
                                    <option value="false">false</option>
                                </select>
                            </td>
                        </tr>
                    </table>
                </fieldset>

                <button type="submit" hidden>Submit</button>
            </form>

            <div class="help">
                SAC validation format example:<br>
                [{"rowIndex":0,"field":"ID","message":"Blank cell not allowed"},{"rowIndex":2,"field":"costcenter","message":"Invalid cost center"}]
            </div>
            <div class="section-note">
                Tip: validationresult + validationerrors can be set from SAC script so the widget visually highlights errors.
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
                "maxrows",
                "previewrows",
                "allowcsv",
                "autovalidate",
                "showpreview",
                "showlogs"
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

