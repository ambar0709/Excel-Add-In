/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = getData;
    document.getElementById("sheet-data").onclick = getSheetData;
  }
});

export async function getSheetData() {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let data = sheet.toJSON();
      await context.sync();
      console.log(data);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function run() {
  try {
    await Excel.run(async (context) => {
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

export async function getData() {
  try {
    const userName = document.getElementById("userName").value;
    fetch(`https://api.github.com/users/${userName}`)
      .then((response) => {
        return response.json();
      })
      .then((data) => {
        document.getElementById("api-data").innerHTML = JSON.stringify(data);
        setSheetData(data);
      });
  } catch (error) {
    console.log(error);
  }
}

export async function setSheetData(data) {
  setHeader(data);
  //await setData(productData);
}

var rowIndex = 1;
export async function setHeader(data) {
  Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    // // Create the headers and format them to stand out.
    let headers = [Object.keys(data)];
    let headerRange = sheet.getRangeByIndexes(0, 0, 1, headers[0].length);

    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "black";
    headerRange.format.font.bold = true;

    // Create the product data rows.
    let productData = [[]];

    headers[0].forEach((header) => {
      productData[0].push(data[header]);
    });

    let dataRange = sheet.getRangeByIndexes(rowIndex, 0, 1, productData[0].length);
    rowIndex++;
    dataRange.values = productData;
    dataRange.format.font.color = "black";
    dataRange.format.autofitColumns();
    await context.sync();

    return productData;
  });
}

export async function setData(productData) {
  Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    //let usedRange = sheet.getUsedRange();

    // Find the next empty row
    // let nextRow = usedRange.rowIndex + usedRange.rowCount;

    let dataRange = sheet.getRangeByIndexes(1, 0, 1, productData[0].length);
    dataRange.values = productData;
    dataRange.format.font.color = "black";
    dataRange.format.autofitColumns();
    await context.sync();
  });
}
