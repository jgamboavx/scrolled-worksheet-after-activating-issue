/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run-scrolled").onclick = runScrolled;
    document.getElementById("run-not-scrolled").onclick = runNotScrolled;
  }
});

export async function runScrolled() {
  try {
    await Excel.run(async (context) => {
      const newWorksheet = context.workbook.worksheets.add();

      newWorksheet.getRange("A1:A2").merge();

      newWorksheet.activate();

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function runNotScrolled() {
  try {
    await Excel.run(async (context) => {
      const newWorksheet = context.workbook.worksheets.add();

      newWorksheet.activate();

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
