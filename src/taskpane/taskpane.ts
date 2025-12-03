/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady((info) => {
  console.log("Taskpane Office.onReady called, host:", info.host);
  
  if (info.host === Office.HostType.Excel) {
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    const runBtn = document.getElementById("run");
    
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";
    if (runBtn) runBtn.onclick = run;
    
    console.log("Taskpane initialized successfully");
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "yellow";
      await context.sync();
      console.log(`The range address was ${range.address}.`);  // Fixed syntax
    });
  } catch (error) {
    console.error(error);
  }
}