/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
/*if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
  console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
}*/
Office.context.requirements.isSetSupported('ExcelApi', '1.9');
// Assign event handlers and other initialization logic.




function createTable() {
   // Excel.context.workbook.worksheets.getActiveWorksheet().getRange()
  Excel.run(function (context) {

      // TODO1: Queue table creation logic here.
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";

      // TODO2: Queue commands to populate the table with data.
expensesTable.getHeaderRowRange().values =
    [["Date", "Merchant", "Category", "Amount"]];

expensesTable.rows.add(null /*add at the end*/, [
    ["1/1/2017", "The Phone Company", "Communications", "120"],
    ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
    ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
    ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
    ["1/11/2017", "Bellows College", "Education", "350.1"],
    ["1/12/2017", "Trey Research", "Other", "135"],
    ["2/11/2017", "Best For You Organics Company", "Groceries", "97.88"]
]);
      // TODO3: Queue commands to format the table.
      expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
      expensesTable.getRange().format.autofitColumns();
      expensesTable.getRange().format.autofitRows();
      return context.sync();
  })
  .catch(function (error) {
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
    })
    .catch(function (error) {
        console.log("Error: " + error);
    });
}
document.getElementById("sort-table").onclick = sortTable;
function sortTable() {
    Excel.run(function (context) {

        // TODO1: Queue commands to sort the table by Merchant name.
        var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
        var sortFields = [
            {
                key: 1,            // Merchant column
                ascending: false,
            }
        ];
        
        expensesTable.sort.apply(sortFields);
        return context.sync();
    })
    .catch(function (error) {
        console.log("Error: " + error);
    });
}


document.getElementById("helloworld").onclick = helloword;
function helloword() {
    Excel.run(function (context) {
       var R= context.workbook.getSelectedRange();
       R.load();

      R.values = "Hello World";
        return context.sync();
    })
    .catch(function (error) {
        console.log("Error: " + error);
    });
}