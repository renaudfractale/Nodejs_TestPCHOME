/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
class ConfTable {
  NameSheet : string
  NameTable : string
  constructor(nameSheet : string, nameTable: string){
    this.NameSheet=nameSheet;
    this.NameTable=nameTable;
  }
   async Exist( context : Excel.RequestContext) : Promise<boolean>{
      
      try {
        //chargemenet des Sheets
        let sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        let sheet :  Excel.Worksheet = null;
        for( let i =0 ; i<sheets.items.length;i++){
          if(sheets.items[i].name==this.NameSheet){
            sheet= sheets.items[i];
          }
        }
        if(sheet==null) {
          return false;
        }
        //chargemenet des Tables
        let tables : Excel.TableCollection = sheet.tables;
        tables.load("items/name");
        await context.sync();
        let table : Excel.Table = null;
        for( let i =0 ; i<tables.items.length;i++){
          if(tables.items[i].name==this.NameTable){
            table = tables.items[i];
          }
        }
        if(table==null) {
          return false;
        } else {
          return true;
        }
      } catch (error) {
        return false;
      }
  }
}
Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("run2").onclick = run2;
    document.getElementById("run3").onclick = run3;
    document.getElementById("ok").onclick = ok;
  }
});
let ArrayData : Array<string>= ["1", "2", "aaaaa"];
//document.getElementById("sync").onclick = sync;
sync();
export function sync(){
  let dataist : HTMLElement =  document.getElementById("ListeTask");
  dataist.innerHTML = "";
  for( let i =0;i<ArrayData.length;i++ ){
    let txt = ArrayData[i]
    let option : HTMLOptionElement = document.createElement("option")
    option.value = txt;
    option.dataset.id=i.toString();
    dataist.appendChild(option)
  }
}

export async function ok(){

  let bool : boolean = true
  let txtHtml : HTMLInputElement = <HTMLInputElement>document.getElementById("TaskName");
  for( let i =0;i<ArrayData.length;i++ ){
    let txt = ArrayData[i]
    if( txt==txtHtml.value){
      bool=false
       ArrayData.splice(i,1)
      break
    }
  }
  document.getElementById("log").innerHTML+= "<br>" +"9999999999999"
  if(bool==true){
    ArrayData.push(txtHtml.value)
  }
  sync();
}

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
    let  rowRange=  expensesTable.rows.getItemAt(0).load("values");
    await context.sync();

    document.getElementById("log").innerHTML+= "<br>"+ rowRange.values

   await context.sync();



  });
}



export async function run3() {
  await Excel.run(async context => {
    document.getElementById("log").innerHTML+= "<br>"+"0000000000000"
    let NameSheets : string = "Planning";
    let NameTable : string = "TablePlanning"
    document.getElementById("log").innerHTML+= "<br>"+"111111111"
    let conf = new ConfTable(NameSheets,NameTable);
    let status = await conf.Exist(context)
    document.getElementById("log").innerHTML+= "<br>"+ status

   await context.sync();



  });
}