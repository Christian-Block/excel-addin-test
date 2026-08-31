import { DEMO_TABLE_HEADERS, DEMO_TABLE_NAME } from "./types";

const DEFAULT_SHEET_RANGE = "A1:C1";

const SAMPLE_ROWS: (string | number)[][] = [
  ["Pencil", 120, 0.5],
  ["Notebook", 45, 2.75],
  ["Backpack", 12, 24.99],
];

/**
 * Creates the demo table on the active worksheet and fills it with sample data.
 * The table is created only once; calling the function again keeps the
 * existing table.
 */
export async function createDemoTable(): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.add(DEFAULT_SHEET_RANGE, true /* hasHeaders */);
    table.name = DEMO_TABLE_NAME;
    table.getHeaderRowRange().values = [DEMO_TABLE_HEADERS];
    table.rows.add(undefined /* append */, SAMPLE_ROWS);
    sheet.getUsedRange().format.autofitColumns();

    await context.sync();
  });
}

/**
 * Reads the value of a single cell.
 *
 * @param address Cell address, for example "A2".
 */
export async function readCell(address: string): Promise<string | number | boolean> {
  let value: string | number | boolean = "";

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.load("values");

    await context.sync();

    value = range.values[0][0];
  });

  return value;
}

/**
 * Writes a value into a single cell.
 *
 * @param address Cell address, for example "A2".
 * @param value Value to write.
 */
export async function writeCell(address: string, value: string | number | boolean): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.values = [[value]];

    await context.sync();
  });
}

/**
 * Reads the data body of the demo table.
 */
export async function readTableRows(): Promise<(string | number)[][]> {
  let rows: (string | number)[][] = [];

  await Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(DEMO_TABLE_NAME);
    const bodyRange = table.getDataBodyRange();
    bodyRange.load("values");

    await context.sync();

    rows = bodyRange.values as (string | number)[][];
  });

  return rows;
}

/**
 * Replaces the data body of the demo table with the given rows.
 */
export async function writeTableRows(rows: (string | number)[][]): Promise<void> {
  await Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(DEMO_TABLE_NAME);
    table.getDataBodyRange().delete(Excel.DeleteShiftDirection.up);
    await context.sync();

    table.rows.add(undefined /* append */, rows);

    await context.sync();
  });
}
