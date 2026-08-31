import { installExcelMock } from "./mocks/excelMock";
import { createDemoTable, readCell, readTableRows, writeCell, writeTableRows } from "../src/core/tableService";
import { DEMO_TABLE_HEADERS, DEMO_TABLE_NAME } from "../src/core/types";

describe("tableService", () => {
  it("creates the demo table with headers and sample rows", async () => {
    const excelMock = installExcelMock();

    await createDemoTable();

    const sheet = excelMock.context.workbook.worksheets.worksheet;
    expect(sheet.tables.addedAddress).toBe("A1:C1");
    expect(sheet.tables.addedHasHeaders).toBe(true);
    expect(sheet.tables.table.name).toBe(DEMO_TABLE_NAME);
    expect(sheet.tables.table.headerRowRange.values).toEqual([DEMO_TABLE_HEADERS]);
    expect(sheet.tables.table.rows.addedRowCount).toBe(3);
    expect(sheet.usedRange.format.autofitCalled).toBe(true);
  });

  it("reads a single cell", async () => {
    const excelMock = installExcelMock();
    excelMock.context.workbook.worksheets.worksheet.range.values = [["Pencil"]];

    const value = await readCell("A2");

    expect(value).toBe("Pencil");
    expect(excelMock.context.workbook.worksheets.worksheet.requestedAddress).toBe("A2");
  });

  it("writes a single cell", async () => {
    const excelMock = installExcelMock();

    await writeCell("B3", 42);

    expect(excelMock.context.workbook.worksheets.worksheet.range.values).toEqual([[42]]);
  });

  it("reads the rows of the demo table", async () => {
    installExcelMock([["Pencil", 120, 0.5]]);

    await expect(readTableRows()).resolves.toEqual([["Pencil", 120, 0.5]]);
  });

  it("replaces the rows of the demo table", async () => {
    const excelMock = installExcelMock();

    await writeTableRows([["Eraser", 3, 1.5]]);

    const table = excelMock.context.workbook.tables.table;
    expect(excelMock.context.workbook.tables.requestedName).toBe(DEMO_TABLE_NAME);
    expect(table.bodyRange.wasDeleted).toBe(true);
    expect(table.rows.lastAddedRows).toEqual([["Eraser", 3, 1.5]]);
  });
});
