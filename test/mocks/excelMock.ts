import { OfficeMockObject } from "office-addin-mock";

/**
 * Mock data for the Excel JavaScript API, built with the official
 * `office-addin-mock` library. See
 * https://learn.microsoft.com/office/dev/add-ins/testing/unit-testing
 */

function conditionalFormatsMock() {
  return {
    wasAdded: false,
    wasCleared: false,
    type: "",
    format: {
      custom: {
        rule: { formula: "" },
        format: { fill: { color: "" }, font: { color: "" } },
      },
      // `criteria` is intentionally not part of the mock data: object
      // properties are restored on every sync, which would undo the
      // assignment made by the add-in.
      colorScale: {},
    },
    add: function (type: string) {
      this.wasAdded = true;
      this.type = type;
      return this.format;
    },
    clearAll: function () {
      this.wasCleared = true;
    },
  };
}

function rangeMock(values: unknown[][]) {
  return {
    values,
    wasDeleted: false,
    shift: "",
    conditionalFormats: conditionalFormatsMock(),
    format: {
      autofitCalled: false,
      autofitColumns: function () {
        this.autofitCalled = true;
      },
    },
    delete: function (shift: string) {
      this.wasDeleted = true;
      this.shift = shift;
    },
  };
}

function tableMock(rows: (string | number)[][]) {
  return {
    name: "",
    headerRowRange: rangeMock([[]]),
    bodyRange: rangeMock(rows),
    range: rangeMock(rows),
    rows: {
      addedRowCount: 0,
      lastAddedRows: [] as (string | number)[][],
      add: function (index: number | undefined, values: (string | number)[][]) {
        this.addedRowCount = values.length;
        this.lastAddedRows = values;
      },
    },
    columns: {
      requestedName: "",
      column: {
        bodyRange: rangeMock(rows),
        getDataBodyRange: function () {
          return this.bodyRange;
        },
      },
      getItem: function (name: string) {
        this.requestedName = name;
        return this.column;
      },
    },
    getHeaderRowRange: function () {
      return this.headerRowRange;
    },
    getDataBodyRange: function () {
      return this.bodyRange;
    },
    getRange: function () {
      return this.range;
    },
  };
}

export function createExcelMockData(rows: (string | number)[][] = [["Pencil", 120, 0.5]]) {
  return {
    ConditionalFormatType: {
      custom: "Custom",
      colorScale: "ColorScale",
    },
    ConditionalFormatColorCriterionType: {
      lowestValue: "LowestValue",
      highestValue: "HighestValue",
      percentile: "Percentile",
    },
    DeleteShiftDirection: {
      up: "Up",
    },
    context: {
      workbook: {
        tables: {
          requestedName: "",
          table: tableMock(rows),
          getItem: function (name: string) {
            this.requestedName = name;
            return this.table;
          },
        },
        worksheets: {
          worksheet: {
            requestedAddress: "",
            range: rangeMock([[""]]),
            usedRange: rangeMock(rows),
            tables: {
              addedAddress: "",
              addedHasHeaders: false,
              table: tableMock(rows),
              add: function (address: string, hasHeaders: boolean) {
                this.addedAddress = address;
                this.addedHasHeaders = hasHeaders;
                return this.table;
              },
            },
            getRange: function (address: string) {
              this.requestedAddress = address;
              return this.range;
            },
            getUsedRange: function () {
              return this.usedRange;
            },
          },
          getActiveWorksheet: function () {
            return this.worksheet;
          },
        },
      },
    },
    run: async function (callback: (context: unknown) => Promise<unknown>) {
      await callback(this.context);
    },
  };
}

/**
 * Creates the Excel mock and registers it as the global `Excel` object.
 */
export function installExcelMock(rows?: (string | number)[][]): any {
  const excelMock = new OfficeMockObject(createExcelMockData(rows)) as any;

  (globalThis as any).Excel = excelMock;

  return excelMock;
}
