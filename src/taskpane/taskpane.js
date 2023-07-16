/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = fetchData;
    document.getElementById("sheet-data").onclick = readSheetData;
    document.getElementById("userName").onkeyup = resetError;
  }
});

export async function readSheetData() {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = sheet.getUsedRange();
      // Load the values from the range
      range.load("values");

      // Synchronize the document state by executing the queued commands
      await context.sync();

      // Access the values
      let values = range.values;

      // Get the headers from the first row
      let headers = values[0];

      // Create an array of JSON objects
      let jsonArray = [];

      // Iterate over the remaining rows
      for (let i = 1; i < values.length; i++) {
        let row = values[i];
        let jsonObject = {};

        // Iterate over the cells in the row and assign values to the corresponding headers
        for (let j = 0; j < headers.length; j++) {
          let header = headers[j];
          let value = row[j];
          jsonObject[header] = value;
        }

        jsonArray.push(jsonObject);
      }

      // Print the resulting JSON object array
      document.getElementById("api-data").innerHTML = JSON.stringify(jsonArray);
      console.log(JSON.stringify(jsonArray));
    });
  } catch (error) {
    console.error(error);
  }
}

export async function fetchData() {
  try {
    Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();

      var usedRange = sheet.getUsedRange(true);

      usedRange.load("rowCount");
      await context.sync();

      var nameRange = sheet.getRangeByIndexes(1, 0, usedRange.rowCount, 1);

      nameRange.load("values");
      await context.sync();

      const userName = document.getElementById("userName").value;
      var values = nameRange.values;

      if (values.flatMap(value => value[0]).includes(userName)) {
       document.getElementById("userNameError").innerText = "user already added";
        return;
      }
      fetch(`https://api.github.com/users/${userName}`)
        .then((response) => {
          return response.json();
        })
        .then((data) => {
          setSheetData(data);
        });

      context.sync();
    });
  } catch (error) {
    console.log(error);
  }
}

function resetError(){
  document.getElementById("userNameError").innerText = "";
}

export async function setSheetData(data) {
  let productData = [[]];
  let headers = [Object.keys(data)];
  Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    // // Create the headers and format them to stand out.
    let headerRange = sheet.getRangeByIndexes(0, 0, 1, headers[0].length);

    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "black";
    headerRange.format.font.bold = true;

    // Create the product data rows.

    headers[0].forEach((header) => {
      productData[0].push(data[header]);
    });

    await context.sync();

    let usedRange = sheet.getUsedRange();
    usedRange.load("rowCount");

    await context.sync();

    let dataRange = sheet.getRangeByIndexes(usedRange.rowCount, 0, 1, productData[0].length);

    dataRange.values = productData;
    dataRange.format.font.color = "black";
    dataRange.format.autofitColumns();
    await context.sync();

    // sheet.onChanged("cellValueChanged", (cell) => {
    //   const updatedValue = cell.value;
    //   const rowNumber = cell.row.number;
    //   const columnNumber = cell.column.number;
    //   // Perform necessary actions to sync the updated value with your database
    //   console.log(`Value updated: ${updatedValue}, row number: ${rowNumber}`);
    // });
  });
}
