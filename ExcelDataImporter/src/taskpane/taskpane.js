/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    // document.getElementById("status").textContent = "Fetching data...";
    const data = await fetchDataFromApi();

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1").getResizedRange(data.length - 1, data[0].length - 1);
      range.values = data;
      await context.sync();
      // document.getElementById("status").textContent = "Data inserted!";
    });
 
  } catch (error) {
    console.error(error);
  }
}

async function fetchDataFromApi() {
  return new Promise((resolve) => {
    setTimeout(() => {
      resolve([
        ["ID", "Name", "Email"],
        [1, "Alice", "alice@example.com"],
        [2, "Bob", "bob@example.com"],
        [3, "Charlie", "charlie@example.com"],
      ]);
    }, 1000);
  });
}
