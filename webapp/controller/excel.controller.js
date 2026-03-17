sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/m/Column",
    "sap/m/ColumnListItem",
    "sap/m/Text",
    "sap/m/Label",
    "sap/ui/export/Spreadsheet",
    "sap/ui/export/library"
], function (Controller, JSONModel, MessageToast, MessageBox,
             Column, ColumnListItem, Text, Label, Spreadsheet, exportLibrary) {
    "use strict";

    var EdmType = exportLibrary.EdmType;

    return Controller.extend("exportdoc.exportexcel.controller.excel", {

        // ─── Lifecycle ───────────────────────────────────────────────

        onInit: function () {
            // Excel model
            this.getView().setModel(new JSONModel({
                rows: [], columns: [], rawPreview: "",
                fileInfo: "", fileType: "", totalRows: 0, totalColumns: 0
            }), "documentModel");

            // Media model
            this.getView().setModel(new JSONModel({
                imageName: "", imageUploaded: false,
                videoName: "", videoUploaded: false
            }), "mediaModel");

            this._allRows       = [];
            this._fileData      = null;
            this._imageFile     = null;
            this._videoFile     = null;
            this._imageDataUrl  = null;
            this._videoObjectUrl = null;
        },

        // ════════════════════════════════════════════════════════════
        //  EXCEL SECTION
        // ════════════════════════════════════════════════════════════

        onFileChange: function (oEvent) {
            var oFile = oEvent.getParameter("files")[0];
            if (!oFile) return;
            this._fileData = oFile;
            this.getView().getModel("documentModel")
                .setProperty("/fileInfo", "Selected: " + oFile.name +
                    " | Size: " + this._formatFileSize(oFile.size));
            MessageToast.show("File selected. Click 'Load Document' to parse.");
        },

        onTypeMismatch: function (oEvent) {
            MessageBox.error("Invalid file: " + oEvent.getParameter("fileName") +
                "\nOnly .xlsx files are supported.");
        },

        onLoadDocument: function () {
            if (!this._fileData) { MessageBox.warning("Please select an Excel file first."); return; }
            if (!this._fileData.name.toLowerCase().endsWith(".xlsx")) {
                MessageBox.error("Only .xlsx files are supported."); return;
            }
            this._parseXlsx(this._fileData);
        },

        _parseXlsx: function (oFile) {
            var that = this;
            var reader = new FileReader();
            reader.onload = function (e) {
                try {
                    var uint8 = new Uint8Array(e.target.result);
                    if (uint8[0] !== 0x50 || uint8[1] !== 0x4B) {
                        MessageBox.error("Not a valid .xlsx file."); return;
                    }
                    that._extractZipEntries(uint8).then(function (entries) {
                        var ss = entries["xl/sharedStrings.xml"]
                            ? that._parseSharedStrings(entries["xl/sharedStrings.xml"]) : [];
                        var sheetXml = entries["xl/worksheets/sheet1.xml"] ||
                            entries[Object.keys(entries).find(function (k) {
                                return k.startsWith("xl/worksheets/sheet") && k.endsWith(".xml");
                            })];
                        if (!sheetXml) { MessageBox.error("No worksheet found."); return; }
                        var result = that._parseSheetXml(sheetXml, ss);
                        if (!result || result.rows.length === 0) {
                            MessageBox.warning("Excel file appears empty."); return;
                        }
                        that._updateTableData(result.headers, result.rows, "Excel (.xlsx)", "");
                        that._setRawPreview(result.rows.slice(0, 30)
                            .map(function (r) { return Object.values(r).join(" | "); }).join("\n"));
                    }).catch(function (err) { MessageBox.error("Failed: " + err.message); });
                } catch (err) { MessageBox.error("Error: " + err.message); }
            };
            reader.onerror = function () { MessageBox.error("Could not read the file."); };
            reader.readAsArrayBuffer(oFile);
        },

        _extractZipEntries: function (uint8) {
            var entries = {}, promises = [], i = 0;
            while (i < uint8.length - 30) {
                if (uint8[i] === 0x50 && uint8[i+1] === 0x4B &&
                    uint8[i+2] === 0x03 && uint8[i+3] === 0x04) {
                    var compression = uint8[i+8] | (uint8[i+9] << 8);
                    var compSize    = uint8[i+18] | (uint8[i+19] << 8) | (uint8[i+20] << 16) | (uint8[i+21] << 24);
                    var fnLen       = uint8[i+26] | (uint8[i+27] << 8);
                    var exLen       = uint8[i+28] | (uint8[i+29] << 8);
                    var name        = new TextDecoder("utf-8").decode(uint8.slice(i+30, i+30+fnLen));
                    var dataStart   = i + 30 + fnLen + exLen;
                    var compData    = uint8.slice(dataStart, dataStart + compSize);
                    if (name.endsWith(".xml") || name.endsWith(".rels")) {
                        if (compression === 0) {
                            entries[name] = new TextDecoder("utf-8").decode(compData);
                        } else if (compression === 8) {
                            promises.push(this._inflate(compData).then(
                                function (n, t) { entries[n] = t; }.bind(null, name)));
                        }
                    }
                    i = dataStart + compSize;
                } else { i++; }
            }
            return Promise.all(promises).then(function () { return entries; });
        },

        _inflate: function (compData) {
            return new Promise(function (resolve, reject) {
                try {
                    var ds = new DecompressionStream("deflate-raw");
                    var writer = ds.writable.getWriter();
                    var reader = ds.readable.getReader();
                    var chunks = [];
                    function read() {
                        reader.read().then(function (r) {
                            if (r.done) {
                                var total = 0, offset = 0;
                                chunks.forEach(function (c) { total += c.length; });
                                var out = new Uint8Array(total);
                                chunks.forEach(function (c) { out.set(c, offset); offset += c.length; });
                                resolve(new TextDecoder("utf-8").decode(out));
                            } else { chunks.push(r.value); read(); }
                        }).catch(reject);
                    }
                    read(); writer.write(compData); writer.close();
                } catch (err) { reject(err); }
            });
        },

        _parseSharedStrings: function (xml) {
            var strings = [];
            (xml.match(/<si>([\s\S]*?)<\/si>/g) || []).forEach(function (si) {
                var val = (si.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || [])
                    .map(function (t) { return t.replace(/<\/?t[^>]*>/g, ""); }).join("");
                strings.push(val.replace(/&amp;/g,"&").replace(/&lt;/g,"<")
                    .replace(/&gt;/g,">").replace(/&quot;/g,'"'));
            });
            return strings;
        },

        _parseSheetXml: function (xml, ss) {
            var rows = [];
            (xml.match(/<row[^>]*>([\s\S]*?)<\/row>/g) || []).forEach(function (rowXml) {
                var rowData = {};
                (rowXml.match(/<c[^>]*>[\s\S]*?<\/c>/g) || []).forEach(function (cellXml) {
                    var ref   = (cellXml.match(/r="([A-Z]+)\d+"/) || [])[1];
                    var type  = (cellXml.match(/t="([^"]*)"/) || [])[1] || "";
                    var raw   = (cellXml.match(/<v[^>]*>([\s\S]*?)<\/v>/) || [])[1] || "";
                    var inl   = (cellXml.match(/<is>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/is>/) || [])[1] || "";
                    if (!ref) return;
                    var val = type === "s" ? (ss[parseInt(raw)] || raw)
                            : type === "inlineStr" ? inl
                            : type === "b" ? (raw === "1" ? "TRUE" : "FALSE")
                            : raw;
                    rowData[ref] = val.replace(/&amp;/g,"&").replace(/&lt;/g,"<")
                        .replace(/&gt;/g,">").replace(/&quot;/g,'"');
                });
                rows.push(rowData);
            });
            if (!rows.length) return { headers: [], rows: [] };
            var allCols = [];
            rows.forEach(function (r) {
                Object.keys(r).forEach(function (c) { if (allCols.indexOf(c) < 0) allCols.push(c); });
            });
            allCols.sort(function (a, b) {
                return a.length !== b.length ? a.length - b.length : a < b ? -1 : 1;
            });
            var hMap = {};
            allCols.forEach(function (c) {
                hMap[c] = rows[0][c] && rows[0][c].trim() ? rows[0][c] : c;
            });
            return {
                headers: allCols.map(function (c) { return hMap[c]; }),
                rows: rows.slice(1).map(function (r) {
                    var o = {};
                    allCols.forEach(function (c) { o[hMap[c]] = r[c] || ""; });
                    return o;
                })
            };
        },

        _updateTableData: function (aHeaders, aRows, sType, sExtra) {
            var oView = this.getView(), oModel = oView.getModel("documentModel"),
                oTable = oView.byId("dataTable");
            oTable.destroyColumns(); oTable.unbindItems();
            aHeaders.forEach(function (h) {
                oTable.addColumn(new Column({
                    header: new Label({ text: h, design: "Bold" }),
                    width: "auto", minScreenWidth: "Tablet", demandPopin: true
                }));
            });
            this._allRows = aRows;
            oModel.setProperty("/rows", aRows);
            oModel.setProperty("/columns", aHeaders);
            oModel.setProperty("/fileType", sType);
            oModel.setProperty("/totalRows", aRows.length);
            oModel.setProperty("/totalColumns", aHeaders.length);
            oTable.bindItems({
                path: "documentModel>/rows",
                template: new ColumnListItem({
                    cells: aHeaders.map(function (h) {
                        return new Text({ text: "{documentModel>" + h + "}" });
                    })
                })
            });
            if (sExtra) {
                oModel.setProperty("/fileInfo", oModel.getProperty("/fileInfo") + " | " + sExtra);
            }
            MessageToast.show(aRows.length + " rows loaded with " + aHeaders.length + " columns.");
        },

        _setRawPreview: function (s) {
            this.getView().getModel("documentModel").setProperty("/rawPreview",
                s + (s.length >= 2000 ? "\n\n... [truncated]" : ""));
        },

        onSearch: function (oEvent) {
            var q = (oEvent.getParameter("newValue") || oEvent.getParameter("query") || "").toLowerCase();
            var oModel = this.getView().getModel("documentModel");
            var aResult = q ? this._allRows.filter(function (r) {
                return Object.values(r).some(function (v) { return String(v).toLowerCase().includes(q); });
            }) : this._allRows;
            oModel.setProperty("/rows", aResult);
            oModel.setProperty("/totalRows", aResult.length);
            if (q) MessageToast.show(aResult.length + " records match.");
        },

        onExportToExcel: function () {
            var oModel = this.getView().getModel("documentModel");
            var aRows = oModel.getProperty("/rows");
            var aColumns = oModel.getProperty("/columns");
            if (!aRows || !aRows.length) { MessageBox.warning("No data to export."); return; }
            var oSheet = new Spreadsheet({
                workbook: { columns: aColumns.map(function (c) { return { label: c, property: c, type: EdmType.String }; }) },
                dataSource: aRows,
                fileName: "ExportedData_" + new Date().toISOString().slice(0, 10) + ".xlsx",
                worker: false
            });
            oSheet.build()
                .then(function () { MessageToast.show("Exported!"); })
                .catch(function (e) { MessageBox.error("Export failed: " + e); })
                .finally(function () { oSheet.destroy(); });
        },

        // ════════════════════════════════════════════════════════════
        //  IMAGE SECTION
        // ════════════════════════════════════════════════════════════

        onImageFileChange: function (oEvent) {
            var oFile = oEvent.getParameter("files")[0];
            if (!oFile) return;
            this._imageFile = oFile;
            MessageToast.show("Image selected: " + oFile.name);
        },

        onImageTypeMismatch: function () {
            MessageBox.error("Unsupported image format.\nUse: jpg, jpeg, png, gif, webp, bmp");
        },

        onPreviewImage: function () {
            if (!this._imageFile) { MessageBox.warning("Please select an image first."); return; }
            var that = this, oView = this.getView();
            var reader = new FileReader();
            reader.onload = function (e) {
                that._imageDataUrl = e.target.result;
                // Set info labels before opening
                oView.byId("imagePreviewName").setText(that._imageFile.name);
                oView.byId("imagePreviewSize").setText(that._formatFileSize(that._imageFile.size));
                // Open dialog — image injected in afterOpen
                oView.byId("imagePreviewDialog").open();
            };
            reader.readAsDataURL(this._imageFile);
        },

        // Called after dialog is fully rendered in DOM
        onImageDialogAfterOpen: function () {
            var oContainer = this.getView().byId("imagePreviewContainer");
            if (oContainer && oContainer.getDomRef()) {
                oContainer.getDomRef().innerHTML =
                    '<img src="' + this._imageDataUrl + '" ' +
                    'style="max-width:100%;max-height:420px;border-radius:8px;' +
                    'box-shadow:0 4px 20px rgba(0,0,0,0.2);object-fit:contain;" />';
            }
        },

        onConfirmImageUpload: function () {
            var oView = this.getView();
            var oMediaModel = oView.getModel("mediaModel");

            // Show the display panel
            oMediaModel.setProperty("/imageUploaded", true);
            oMediaModel.setProperty("/imageName",
                this._imageFile.name + "  (" + this._formatFileSize(this._imageFile.size) + ")");

            // Close dialog
            oView.byId("imagePreviewDialog").close();

            // Inject image into display container after panel renders
            setTimeout(function () {
                var oContainer = oView.byId("imageDisplayContainer");
                if (oContainer && oContainer.getDomRef()) {
                    oContainer.getDomRef().innerHTML =
                        '<img src="' + this._imageDataUrl + '" ' +
                        'style="max-width:100%;max-height:520px;border-radius:10px;' +
                        'box-shadow:0 6px 24px rgba(0,0,0,0.18);object-fit:contain;' +
                        'display:block;margin:0 auto;" />';
                }
            }.bind(this), 400);

            MessageToast.show("Image displayed on page!");
        },

        onCloseImagePreview: function () {
            this.getView().byId("imagePreviewDialog").close();
        },

        onClearImage: function () {
            var oView = this.getView();
            var oMediaModel = oView.getModel("mediaModel");
            oMediaModel.setProperty("/imageUploaded", false);
            oMediaModel.setProperty("/imageName", "");
            this._imageFile    = null;
            this._imageDataUrl = null;
            // Clear the display container
            var oContainer = oView.byId("imageDisplayContainer");
            if (oContainer && oContainer.getDomRef()) {
                oContainer.getDomRef().innerHTML = "";
            }
            oView.byId("imageUploader").clear();
            MessageToast.show("Image cleared.");
        },

        // ════════════════════════════════════════════════════════════
        //  VIDEO SECTION
        // ════════════════════════════════════════════════════════════

        onVideoFileChange: function (oEvent) {
            var oFile = oEvent.getParameter("files")[0];
            if (!oFile) return;
            this._videoFile = oFile;
            MessageToast.show("Video selected: " + oFile.name);
        },

        onVideoTypeMismatch: function () {
            MessageBox.error("Unsupported video format.\nUse: mp4, webm, ogg, mov");
        },

        onPreviewVideo: function () {
            if (!this._videoFile) { MessageBox.warning("Please select a video first."); return; }

            // Revoke previous object URL to avoid memory leak
            if (this._videoObjectUrl) {
                URL.revokeObjectURL(this._videoObjectUrl);
            }
            this._videoObjectUrl = URL.createObjectURL(this._videoFile);

            var oView = this.getView();
            oView.byId("videoPreviewName").setText(this._videoFile.name);
            oView.byId("videoPreviewSize").setText(this._formatFileSize(this._videoFile.size));

            // Open dialog — video injected in afterOpen
            oView.byId("videoPreviewDialog").open();
        },

        // Called after dialog is fully rendered in DOM
        onVideoDialogAfterOpen: function () {
            var oContainer = this.getView().byId("videoPreviewContainer");
            if (oContainer && oContainer.getDomRef()) {
                oContainer.getDomRef().innerHTML =
                    '<video controls style="max-width:100%;max-height:420px;border-radius:8px;' +
                    'box-shadow:0 4px 20px rgba(0,0,0,0.2);">' +
                    '<source src="' + this._videoObjectUrl + '" />' +
                    '</video>';
            }
        },

        onConfirmVideoUpload: function () {
            var oView = this.getView();
            var oMediaModel = oView.getModel("mediaModel");

            // Pause preview video
            var oPreviewContainer = oView.byId("videoPreviewContainer");
            if (oPreviewContainer && oPreviewContainer.getDomRef()) {
                var previewVideo = oPreviewContainer.getDomRef().querySelector("video");
                if (previewVideo) previewVideo.pause();
            }

            // Show display panel
            oMediaModel.setProperty("/videoUploaded", true);
            oMediaModel.setProperty("/videoName",
                this._videoFile.name + "  (" + this._formatFileSize(this._videoFile.size) + ")");

            oView.byId("videoPreviewDialog").close();

            // Inject video into display container after panel renders
            var that = this;
            setTimeout(function () {
                var oContainer = oView.byId("videoDisplayContainer");
                if (oContainer && oContainer.getDomRef()) {
                    oContainer.getDomRef().innerHTML =
                        '<video controls style="max-width:100%;max-height:520px;border-radius:10px;' +
                        'box-shadow:0 6px 24px rgba(0,0,0,0.18);display:block;margin:0 auto;">' +
                        '<source src="' + that._videoObjectUrl + '" />' +
                        '</video>';
                }
            }, 400);

            MessageToast.show("Video displayed on page!");
        },

        onCloseVideoPreview: function () {
            var oView = this.getView();
            // Pause the preview video
            var oContainer = oView.byId("videoPreviewContainer");
            if (oContainer && oContainer.getDomRef()) {
                var vid = oContainer.getDomRef().querySelector("video");
                if (vid) vid.pause();
            }
            oView.byId("videoPreviewDialog").close();
        },

        onClearVideo: function () {
            var oView = this.getView();
            var oMediaModel = oView.getModel("mediaModel");
            oMediaModel.setProperty("/videoUploaded", false);
            oMediaModel.setProperty("/videoName", "");

            // Clear display container
            var oContainer = oView.byId("videoDisplayContainer");
            if (oContainer && oContainer.getDomRef()) {
                var vid = oContainer.getDomRef().querySelector("video");
                if (vid) vid.pause();
                oContainer.getDomRef().innerHTML = "";
            }

            this._videoFile = null;
            if (this._videoObjectUrl) {
                URL.revokeObjectURL(this._videoObjectUrl);
                this._videoObjectUrl = null;
            }
            oView.byId("videoUploader").clear();
            MessageToast.show("Video cleared.");
        },

        // ════════════════════════════════════════════════════════════
        //  CLEAR ALL
        // ════════════════════════════════════════════════════════════

        onClearAll: function () {
            // Excel
            this.getView().getModel("documentModel").setData({
                rows: [], columns: [], rawPreview: "",
                fileInfo: "", fileType: "", totalRows: 0, totalColumns: 0
            });
            this._allRows = []; this._fileData = null;
            var oTable = this.getView().byId("dataTable");
            oTable.destroyColumns(); oTable.unbindItems();
            this.getView().byId("fileUploader").clear();
            // Image + Video
            this.onClearImage();
            this.onClearVideo();
            MessageToast.show("All sections cleared.");
        },

        // ─── Helpers ─────────────────────────────────────────────────

        _formatFileSize: function (n) {
            if (n < 1024) return n + " B";
            if (n < 1048576) return (n / 1024).toFixed(1) + " KB";
            return (n / 1048576).toFixed(1) + " MB";
        }

    });
});