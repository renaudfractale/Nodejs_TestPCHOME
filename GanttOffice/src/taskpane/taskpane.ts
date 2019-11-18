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

export async function run_tableGenerator() {
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
      log_RH(`The range address was ${range.address}.`);
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
