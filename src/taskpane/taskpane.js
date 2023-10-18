/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = () => tryCatch(createTable);
    document.getElementById("filter-table").onclick = () => tryCatch(filterTable);
    document.getElementById("sort-table").onclick = () => tryCatch(sortTable);
    document.getElementById("create-chart").onclick = () => tryCatch(createChart);
    document.getElementById("freeze-header").onclick = () => tryCatch(freezeHeader);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
  }
});

// export async function run() {
//   try {
//     await Excel.run(async (context) => {
//       /**
//        * Insert your Excel code here
//        */
//       const range = context.workbook.getSelectedRange();

//       // Read the range address
//       range.load("address");

//       // Update the fill color
//       range.format.fill.color = "yellow";

//       await context.sync();
//       console.log(`The range address was ${range.address}.`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }

async function createTable() {
  await Excel.run(async (context) => {
    // TODO1: Queue table creation logic here.
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    // TODO2: Queue commands to populate the table with data.
    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
      ["1/1/2017", "The Phone Company", "Communications", "120"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
      ["1/11/2017", "Bellows College", "Education", "350.1"],
      ["1/15/2017", "Trey Research", "Other", "135"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"],
    ]);

    // TODO3: Queue commands to format the table.
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u0024#,##0.00"]];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();

    await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

// async function filterTable() {
//   await Excel.run(async (context) => {
//     // TODO1: Queue commands to filter out all expense categories except
//     //        Groceries and Education.
//     const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
//     const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
//     const categoryFilter = expensesTable.columns.getItem("Category").filter;
//     categoryFilter.applyValuesFilter(["Education", "Groceries"]);

//     await context.sync();
//   });
// }

// Modified function
async function filterTable() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
    const categoryColumn = expensesTable.columns.getItem("Category");

    // Load the filter property of the category column
    categoryColumn.load("filter");

    await context.sync();

    const categoryFilter = categoryColumn.filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);

    await context.sync();
  });
}

async function sortTable() {
  await Excel.run(async (context) => {
    // TODO1: Queue commands to sort the table by Merchant name.
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
    const sortFields = [
      {
        key: 1, // Merchant column
        ascending: false,
      },
    ];

    expensesTable.sort.apply(sortFields);

    await context.sync();
  });
}

async function createChart() {
  await Excel.run(async (context) => {
    // TODO1: Queue commands to get the range of data to be charted.
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
    const dataRange = expensesTable.getDataBodyRange();

    // TODO2: Queue command to create the chart and define its type.
    // First parameter to add method specifies the type of chart, second parameter specifies the range of data to include, third parameter determines whether a series of data points from the table should be charted row-wise or column-wise. The option auto tells Excel to decide the best method.
    const chart = currentWorksheet.charts.add("ColumnClustered", dataRange, "Auto");

    // TODO3: Queue commands to position and format the chart.
    // Parameters to setPosition specify upper left nad lower right cells of worksheet area that should contain the chart. Excel automatically adjusts things like line width.
    // A 'series' = set of data points from column of table. Since there is only one non-strong column in the table, Excel infers that the column is the only column of data points to chart. It interprets the other columns as chart labels. So there will be just one series in the chart and it will have index 0. The is the one to label with 'Value in €'
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "Right";
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = "Value in \u0024";

    await context.sync();
  });
}

async function freezeHeader() {
  await Excel.run(async (context) => {
    // TODO1: Queue commands to keep the header visible when the user scrolls.
    // Worksheet.freezePanes collection is a set of panes in the worksheet that are pinned, or frozen in place when the worksheet is scrolled.
    // The freezeRows method takes as a parameter the number of rows, from the top, that are to be pinned in place. We pass 1 to pin the first row in place.
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);

    await context.sync();
  });
}
