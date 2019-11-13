/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/taskpane/taskpane.js");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/taskpane/taskpane.js":
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.js ***!
  \**********************************/
/*! no static exports found */
/***/ (function(module, exports) {

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
// Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.

/*if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
  console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
}*/
Office.context.requirements.isSetSupported('ExcelApi', '1.9'); // Assign event handlers and other initialization logic.

function createTable() {
  // Excel.context.workbook.worksheets.getActiveWorksheet().getRange()
  Excel.run(function (context) {
    // TODO1: Queue table creation logic here.
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true
    /*hasHeaders*/
    );
    expensesTable.name = "ExpensesTable"; // TODO2: Queue commands to populate the table with data.

    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];
    expensesTable.rows.add(null
    /*add at the end*/
    , [["1/1/2017", "The Phone Company", "Communications", "120"], ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"], ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"], ["1/10/2017", "Coho Vineyard", "Restaurant", "33"], ["1/11/2017", "Bellows College", "Education", "350.1"], ["1/12/2017", "Trey Research", "Other", "135"], ["2/11/2017", "Best For You Organics Company", "Groceries", "97.88"]]); // TODO3: Queue commands to format the table.

    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    return context.sync();
  })["catch"](function (error) {
    console.log("Error: " + error);
  });
}

document.getElementById("create-table").onclick = createTable;
document.getElementById("filter-table").onclick = filterTable;

function filterTable() {
  Excel.run(function (context) {
    // TODO1: Queue commands to filter out all expense categories except
    //        Groceries and Education.
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(['Education', 'Groceries']);
    return context.sync();
  })["catch"](function (error) {
    console.log("Error: " + error);
  });
}

document.getElementById("sort-table").onclick = sortTable;

function sortTable() {
  Excel.run(function (context) {
    // TODO1: Queue commands to sort the table by Merchant name.
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var sortFields = [{
      key: 1,
      // Merchant column
      ascending: false
    }];
    expensesTable.sort.apply(sortFields);
    return context.sync();
  })["catch"](function (error) {
    console.log("Error: " + error);
  });
}

document.getElementById("helloworld").onclick = helloword;

function helloword() {
  Excel.run(function (context) {
    var R = context.workbook.getSelectedRange();
    R.load();
    R.values = "Hello World";
    return context.sync();
  })["catch"](function (error) {
    console.log("Error: " + error);
  });
}

document.getElementById("action1").onclick = action1;

function action1() {
  Excel.run(function (context) {
    var Sheets = context.workbook.worksheets;
    var bool = true;
    document.getElementById("Sheets").innerHTML = bool + " Planning 1"; //document.getElementById("Sheets").innerHTML=  Sheets.names

    var NB = Sheets.getCount().value; // for(var i=0; i<NB.value;i++){
    //   document.getElementById("Sheets").innerHTML+= "|" + i.toString();

    /*  if(Sheets.getItemAt(i).name == "Planning"){
          bool=false
      }*/
    // }

    document.getElementById("Sheets").innerHTML += "|" + NB;
    document.getElementById("Sheets").innerHTML += bool + " Planning 2";

    if (bool == true) {
      context.workbook.worksheets.add("Planning");
      var SheetPlanning = context.workbook.worksheets.getItem("Planning");
      SheetPlanning.getRange("A1").values = "Date de Création";
      SheetPlanning.getRange("B1").values = "Nom de la Task";
      SheetPlanning.getRange("C1").values = "Date Start";
      SheetPlanning.getRange("D1").values = "Duration All (days)";
      SheetPlanning.getRange("E1").values = "Duration Work (hours)";
      SheetPlanning.getRange("F1").values = "Dependance";
    }

    bool = true;

    for (var j in Sheets.items) {
      if (Sheets.items[j].name == "CRA") {
        bool = false;
      }
    }

    if (bool == true) {
      context.workbook.worksheets.add("CRA");
      var SheetCRA = context.workbook.worksheets.getItem("CRA");
      SheetCRA.getRange("A1").values = "Date de Création";
      SheetCRA.getRange("B1").values = "Nom de la Task";
      SheetCRA.getRange("C1").values = "Duration Work (hours)";
    }

    bool = true;

    for (var k in Sheets.items) {
      if (Sheets.items[k].name == "CRA") {
        bool = false;
      }
    }

    if (bool == true) {
      context.workbook.worksheets.add("INFO");
      var SheetINFO = context.workbook.worksheets.getItem("INFO");
      SheetINFO.getRange("A1").values = "GANTT By Renaud HENRY";
      SheetINFO.getRange("A2").values = "Version 0.0.1";
    }

    return context.sync();
  })["catch"](function (error) {
    console.log("Error: " + error);
  });
}

/***/ })

/******/ });
//# sourceMappingURL=taskpane.js.map