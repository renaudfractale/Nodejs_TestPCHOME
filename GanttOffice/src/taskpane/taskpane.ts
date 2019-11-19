/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    //document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run_tableGenerator").onclick = run_tableGenerator;
  }
});


export async function log_RH(txt : string) {
  let DateNow = new Date();
  document.getElementById("log").innerHTML+= DateNow.getFullYear().toString()+"/"
                +DateNow.getMonth().toString()
                +"/"+DateNow.getDate().toString()
                +" "+DateNow.getHours().toString() 
                +":"+DateNow.getMinutes().toString() 
                +":"+DateNow.getSeconds().toString() 
                +" "+DateNow.getMilliseconds().toString() 
                +" : "
                +txt+"<br>"
}

class ConfTable {
  NameSheet : string
  NameTable : string
  constructor(nameSheet : string, nameTable: string){
    this.NameSheet=nameSheet;
    this.NameTable=nameTable;
  }
  async ExistSheet( context : Excel.RequestContext) : Promise<boolean>{
    try {
      //chargemenet des Sheets
      let sheets: Excel.WorksheetCollection= context.workbook.worksheets;
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
      } else {
        return true;
      }
    } catch (error) {
      return false;
    }
  }
  async ExistTable( context : Excel.RequestContext) : Promise<boolean>{
    try {
      //chargemenet des Sheets
      let sheets: Excel.WorksheetCollection= context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      let sheet= sheets.items[0];
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
  async Exist( context : Excel.RequestContext) : Promise<boolean>{
    let statusSheet: boolean = await this.ExistSheet(context) 
    let statusTable: boolean = await this.ExistTable(context)
    return (statusTable && statusSheet)

  }
}


export async function run_tableGenerator() {
  try {
    await Excel.run(async context => {
    let NameSheets : string = "Planning";
    let NameTable : string = "TablePlanning"
    let conf = new ConfTable(NameSheets,NameTable);
    let statusAll: boolean = await conf.Exist(context)
    log_RH("Exits "+NameSheets+"."+NameTable+"("+status.toString()+")");

    let statusSeeht: boolean = await conf.Exist(context)
    log_RH("Exits "+NameSheets+"."+NameTable+"("+status.toString()+")");

    let status: boolean = await conf.Exist(context)
    log_RH("Exits "+NameSheets+"."+NameTable+"("+status.toString()+")");

    if(status==false){

    }


      await context.sync();
    
    });
  } catch (error) {
    console.error(error);
  }
}


export async function run() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
