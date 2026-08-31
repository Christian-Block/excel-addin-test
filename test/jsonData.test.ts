import { installExcelMock } from "./mocks/excelMock";
import { downloadJson, exportTable, importJson, parseJson, toJson } from "../src/core/jsonData";
import { DEMO_TABLE_HEADERS, DEMO_TABLE_NAME, EXPORT_VERSION } from "../src/core/types";

describe("JSON export and import", () => {
  it("exports the table content", async () => {
    installExcelMock([["Pencil", 120, 0.5]]);

    await expect(exportTable()).resolves.toEqual({
      version: EXPORT_VERSION,
      tableName: DEMO_TABLE_NAME,
      headers: DEMO_TABLE_HEADERS,
      rows: [["Pencil", 120, 0.5]],
    });
  });

  it("serialises and parses the export payload", () => {
    const payload = { version: 1, tableName: "DemoTable", headers: ["A"], rows: [["x"]] };

    expect(parseJson(toJson(payload))).toEqual(payload);
  });

  it("rejects malformed JSON", () => {
    expect(() => parseJson("{")).toThrow(/valid JSON/);
  });

  it("writes imported rows into the table", async () => {
    const excelMock = installExcelMock();
    const json = toJson({ version: 1, tableName: DEMO_TABLE_NAME, headers: DEMO_TABLE_HEADERS, rows: [["Eraser", 3, 1.5]] });

    const imported = await importJson(json);

    expect(imported.rows).toEqual([["Eraser", 3, 1.5]]);
    expect(excelMock.context.workbook.tables.table.rows.lastAddedRows).toEqual([["Eraser", 3, 1.5]]);
  });

  it("triggers a download of the JSON file", () => {
    const createObjectURL = jest.fn().mockReturnValue("blob:demo");
    const revokeObjectURL = jest.fn();
    (URL as unknown as { createObjectURL: unknown }).createObjectURL = createObjectURL;
    (URL as unknown as { revokeObjectURL: unknown }).revokeObjectURL = revokeObjectURL;
    const click = jest.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);

    const url = downloadJson("{}", "demo.json");

    expect(url).toBe("blob:demo");
    expect(click).toHaveBeenCalled();
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:demo");
    expect(document.querySelectorAll("a")).toHaveLength(0);

    click.mockRestore();
  });
});
