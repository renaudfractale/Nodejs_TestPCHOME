/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("run2").onclick = run2;
  }
});

export async function run() {
    await Excel.run(async context => {
      var sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      for(var i =0 ; i<sheets.items.length;i++){
        document.getElementById("log").innerHTML+= "<br>"+sheets.items[i].name;
      }
      let bool : boolean = true;
      let NameSheets : string = "Planning";

      sheets.load("items/name");
      await context.sync();
      for( i =0 ; i<sheets.items.length;i++){
        if(sheets.items[i].name==NameSheets){
          bool=false;
        }
      }
      if(bool==true){
        context.workbook.worksheets.add(NameSheets);
        let SheetPlanning= context.workbook.worksheets.getItem(NameSheets);

        let expensesTable = SheetPlanning.tables.add("A1:F1", true /*hasHeaders*/);
        expensesTable.name = "TablePlanning";

        expensesTable.getHeaderRowRange().values =[["Date de Création", "Nom de la Task", "Date Start", "Duration All (days)","Duration Work (hours)","Dependance"]];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
    }
    await context.sync();



      bool  = true;
      NameSheets = "CRA";

      sheets.load("items/name");
      await context.sync();
      for( i =0 ; i<sheets.items.length;i++){
        if(sheets.items[i].name==NameSheets){
          bool=false;
        }
      }
      if(bool==true){
        context.workbook.worksheets.add(NameSheets);
        let SheetCRA= context.workbook.worksheets.getItem(NameSheets);
        SheetCRA.getRange("A1").values=[["Date de Création"]];
        SheetCRA.getRange("B1").values=[["Nom de la Task"]];
        SheetCRA.getRange("C1").values=[["Duration Work (hours)"]];
    }


      bool  = true;
      NameSheets = "INFO";

    sheets.load("items/name");
    await context.sync();
    for( i =0 ; i<sheets.items.length;i++){
      if(sheets.items[i].name==NameSheets){
        bool=false;
      }
    }
    if(bool==true){
      context.workbook.worksheets.add(NameSheets);
      let SheetCRA= context.workbook.worksheets.getItem(NameSheets);
      SheetCRA.getRange("A1").values=[["GANTT By Renaud HENRY"]];
      SheetCRA.getRange("A2").values=[["Version 0.0.1"]];
    }



     // const Sheets =  context.workbook.worksheets.items;
    //  Sheets.load("count")
      //document.getElementById("log").innerHTML=  "5555";

     document.getElementById("log").innerHTML+= "<br>"+sheets.items.length;//+Sheets.length.toString();
     //context.sync()
     const range = context.workbook.getSelectedRange();

     // Read the range address
     range.load("address");

     // Update the fill color
     range.format.fill.color = "yellow";


      await context.sync();
    });
}


export async function run2() {
  await Excel.run(async context => {
    document.getElementById("log").innerHTML+= "<br>"+"888888888888"
    let NameSheets : string = "Planning";


 
    await context.sync();
   
    let SheetPlanning= context.workbook.worksheets.getItem(NameSheets);
    document.getElementById("log").innerHTML+= "<br>"+ "5555555555555";
    SheetPlanning.load("tables");
    await context.sync();
    let expensesTable = SheetPlanning.tables.getItem("TablePlanning")
    document.getElementById("log").innerHTML+= "<br>"+ "666666666";
    expensesTable.load("rows");
    await context.sync();
    document.getElementById("log").innerHTML+= "<br>"+ "7777777777";
   let  rowRange=  expensesTable.rows.getItemAt(1).load("values");
    await context.sync();
    document.getElementById("log").innerHTML+= "<br>"+ rowRange.values

   await context.sync();



  });
}