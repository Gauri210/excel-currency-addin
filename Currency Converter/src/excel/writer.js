

export const CONVERSIONS_SHEET = "Conversions";
export const CONVERSIONS_TABLE = "ConversionTable";
export const HISTORY_SHEET = "Rate History";
export const HISTORY_TABLE = "RateHistoryTable";

//                      A         B       C     D       E            F
export const HEADERS = ["Amount", "From", "To", "Rate", "Rate date", "Result"];

/** Excel formula for the Result cell of a given row: amount × rate, 2 decimals. */
export function resultFormula(row) {
  return `=ROUND(A${row}*D${row},2)`;
}

/** Get a worksheet by name, creating it if it is not there yet. */
async function getOrCreateSheet(context, name) {
  let sheet = context.workbook.worksheets.getItemOrNullObject(name);
  await context.sync();
  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(name);
  }
  return sheet;
}

/** Apply number formats to one data row (columns A to F). */
function formatRow(sheet, row) {
  sheet.getRange(`A${row}`).numberFormat = [["#,##0.00"]];
  sheet.getRange(`D${row}`).numberFormat = [["0.0000"]];
  sheet.getRange(`E${row}`).numberFormat = [["yyyy-mm-dd"]];
  sheet.getRange(`F${row}`).numberFormat = [["#,##0.00"]];
}

/**
 * Add one conversion to the Conversions table.
 *
 * @param {{amount: number, from: string, to: string, rate: number, rateDate: string, result: number}} c
 *        the object returned by convert() in converter.js
 * @returns {Promise<string>} address of the new row, e.g. "Conversions!A3:F3"
 */
export async function writeConversion(c) {
  return Excel.run(async (context) => {
    const sheet = await getOrCreateSheet(context, CONVERSIONS_SHEET);
    const table = sheet.tables.getItemOrNullObject(CONVERSIONS_TABLE);
    await context.sync();

    let row; // the Excel row number (1-based) the new data goes into

    if (table.isNullObject) {
      // First conversion: write headers in row 1 and data in row 2, then turn A1:F2 into a table.
      row = 2;
      sheet.getRange("A1:F1").values = [HEADERS];
      sheet.getRange("A2:E2").values = [[c.amount, c.from, c.to, c.rate, c.rateDate]];
      sheet.getRange("F2").formulas = [[resultFormula(2)]];

      const newTable = sheet.tables.add("A1:F2", true);
      newTable.name = CONVERSIONS_TABLE;
      newTable.style = "TableStyleMedium2";
    } else {
      // Later conversions: add a row at the bottom of the table.
      const tableRange = table.getRange();
      tableRange.load("rowIndex, rowCount");
      await context.sync();

      row = tableRange.rowIndex + tableRange.rowCount + 1; // rowIndex is 0-based
      table.rows.add(null, [[c.amount, c.from, c.to, c.rate, c.rateDate, resultFormula(row)]]);
    }

    formatRow(sheet, row);
    sheet.getRange("A1:F1").format.font.bold = true;
    sheet.getRange("A:F").format.autofitColumns();

    const newRow = sheet.getRange(`A${row}:F${row}`);
    newRow.select();
    sheet.activate();

    newRow.load("address");
    await context.sync();
    return newRow.address;
  });
}

/**
 * Write the daily rates to the "Rate History" sheet and draw a line chart.
 * The sheet is cleared each time so it always shows the latest request.
 *
 * @param {string} from e.g. "USD"
 * @param {string} to   e.g. "INR"
 * @param {Array<{date: string, rate: number}>} points oldest first
 * @param {{activate?: boolean}} [options]
 */
export async function writeRateHistory(from, to, points, options = {}) {
  if (!points || points.length === 0) return;

  await Excel.run(async (context) => {
    const sheet = await getOrCreateSheet(context, HISTORY_SHEET);

    // Start from a clean sheet: remove the old chart, table and values.
    sheet.charts.load("items");
    const oldTable = sheet.tables.getItemOrNullObject(HISTORY_TABLE);
    await context.sync();
    sheet.charts.items.forEach((chart) => chart.delete());
    if (!oldTable.isNullObject) oldTable.delete();
    sheet.getRange().clear();

    // Data in A:B.
    const lastRow = points.length + 1;
    const header = [["Date", `${from} → ${to}`]];
    const body = points.map((p) => [p.date, p.rate]);
    sheet.getRange("A1:B1").values = header;
    sheet.getRange(`A2:B${lastRow}`).values = body;
    sheet.getRange(`A2:A${lastRow}`).numberFormat = body.map(() => ["yyyy-mm-dd"]);
    sheet.getRange(`B2:B${lastRow}`).numberFormat = body.map(() => ["0.0000"]);

    const table = sheet.tables.add(`A1:B${lastRow}`, true);
    table.name = HISTORY_TABLE;
    table.style = "TableStyleLight9";
    sheet.getRange("A:B").format.autofitColumns();

    // Line chart next to the data.
    const chart = sheet.charts.add(Excel.ChartType.line, sheet.getRange(`A1:B${lastRow}`), Excel.ChartSeriesBy.columns);
    chart.title.text = `${from} to ${to}: ${points[0].date} to ${points[points.length - 1].date}`;
    chart.legend.visible = false;
    chart.axes.valueAxis.title.text = `${to} per 1 ${from}`;
    chart.setPosition("D2", "L20");

    if (options.activate) sheet.activate();
    await context.sync();
  });
}
