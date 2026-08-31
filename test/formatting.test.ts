import { installExcelMock } from "./mocks/excelMock";
import {
  addColumnConditionalFormat,
  addRowConditionalFormat,
  clearConditionalFormats,
} from "../src/core/formatting";

describe("conditional formatting", () => {
  it("adds a custom format that highlights rows below the threshold", async () => {
    const excelMock = installExcelMock();

    await addRowConditionalFormat(30);

    const conditionalFormats = excelMock.context.workbook.tables.table.bodyRange.conditionalFormats;
    expect(conditionalFormats.wasAdded).toBe(true);
    expect(conditionalFormats.format.custom.rule.formula).toBe("=$B2<30");
    expect(conditionalFormats.format.custom.format.fill.color).toBe("#FFF2CC");
  });

  it("adds a colour scale to a single column", async () => {
    const excelMock = installExcelMock();

    await addColumnConditionalFormat("Price");

    const table = excelMock.context.workbook.tables.table;
    expect(table.columns.requestedName).toBe("Price");
    expect(table.columns.column.bodyRange.conditionalFormats.wasAdded).toBe(true);
    expect(table.columns.column.bodyRange.conditionalFormats.format.colorScale.criteria.maximum.color).toBe("#63BE7B");
  });

  it("clears all conditional formats of the table", async () => {
    const excelMock = installExcelMock();

    await clearConditionalFormats();

    expect(excelMock.context.workbook.tables.table.range.conditionalFormats.wasCleared).toBe(true);
  });
});
