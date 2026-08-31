import { isValidPrice, isValidProduct, isValidQuantity, validateRow, validateRows } from "../src/core/validation";
import { assertValidExport } from "../src/core/validation";

describe("value validation", () => {
  it("accepts valid values", () => {
    expect(isValidProduct("Pencil")).toBe(true);
    expect(isValidQuantity(10)).toBe(true);
    expect(isValidQuantity("10")).toBe(true);
    expect(isValidPrice(1.25)).toBe(true);
  });

  it("rejects invalid values", () => {
    expect(isValidProduct("   ")).toBe(false);
    expect(isValidQuantity(-1)).toBe(false);
    expect(isValidQuantity(1.5)).toBe(false);
    expect(isValidPrice("abc")).toBe(false);
    expect(isValidPrice(-0.01)).toBe(false);
  });

  it("reports one issue per invalid cell", () => {
    const issues = validateRow(["", -3, "abc"], 2);

    expect(issues).toHaveLength(3);
    expect(issues[0]).toEqual({ row: 2, column: "Product", message: "Product must not be empty." });
  });

  it("validates several rows", () => {
    expect(validateRows([["Pencil", 1, 1], ["", 1, 1]])).toHaveLength(1);
  });
});

describe("assertValidExport", () => {
  const valid = { version: 1, tableName: "DemoTable", headers: ["A", "B"], rows: [["x", "y"]] };

  it("accepts a valid payload", () => {
    expect(() => assertValidExport(valid)).not.toThrow();
  });

  it("rejects payloads with a wrong shape", () => {
    expect(() => assertValidExport(null)).toThrow(/object/);
    expect(() => assertValidExport({ ...valid, version: "1" })).toThrow(/version/);
    expect(() => assertValidExport({ ...valid, headers: [] })).toThrow(/headers/);
    expect(() => assertValidExport({ ...valid, rows: undefined })).toThrow(/rows/);
    expect(() => assertValidExport({ ...valid, rows: [["only-one"]] })).toThrow(/same length/);
  });
});
