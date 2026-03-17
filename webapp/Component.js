sap.ui.define([
    "sap/ui/core/UIComponent",
    "sap/ui/Device",
    "sap/ui/model/json/JSONModel"
], function (UIComponent, Device, JSONModel) {
    "use strict";

    return UIComponent.extend("exportdoc.exportexcel.Component", {

        metadata: {
            manifest: "json"
        },

        init: function () {
            // Call parent init FIRST — always required
            UIComponent.prototype.init.apply(this, arguments);

            // Device model for responsive behaviour
            var oDeviceModel = new JSONModel(Device);
            oDeviceModel.setDefaultBindingMode("OneWay");
            this.setModel(oDeviceModel, "device");

            // Initialize the router
            this.getRouter().initialize();
        }
    });
});