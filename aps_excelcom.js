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
})();
