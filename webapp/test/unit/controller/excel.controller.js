/*global QUnit*/

sap.ui.define([
	"exportdoc/exportexcel/controller/excel.controller"
], function (Controller) {
	"use strict";

	QUnit.module("excel Controller");

	QUnit.test("I should test the excel controller", function (assert) {
		var oAppController = new Controller();
		oAppController.onInit();
		assert.ok(oAppController);
	});

});
