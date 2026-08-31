import { DEMO_TABLE_HEADERS, ValidationIssue, WorkbookExport } from "./types";

/**
 * Pure validation helpers. They contain no Office API calls so they can be
 * unit tested without a host application.
 */

export function isValidProduct(value: unknown): boolean {
  return typeof value === "string" && value.trim().length > 0;
}

export function isValidQuantity(value: unknown): boolean {
  const numeric = typeof value === "string" ? Number(value) : value;
  return typeof numeric === "number" && Number.isInteger(numeric) && numeric >= 0;
}

export function isValidPrice(value: unknown): boolean {
  const numeric = typeof value === "string" ? Number(value) : value;
  return typeof numeric === "number" && Number.isFinite(numeric) && numeric >= 0;
}

/**
 * Validates a single table row and returns all issues found.
 *
 * @param row Row values in the order of {@link DEMO_TABLE_HEADERS}.
 * @param rowIndex Zero based index of the row, used for reporting.
 */
export function validateRow(row: unknown[], rowIndex: number): ValidationIssue[] {
  const issues: ValidationIssue[] = [];

  if (!isValidProduct(row[0])) {
    issues.push({ row: rowIndex, column: DEMO_TABLE_HEADERS[0], message: "Product must not be empty." });
  }
  if (!isValidQuantity(row[1])) {
    issues.push({ row: rowIndex, column: DEMO_TABLE_HEADERS[1], message: "Quantity must be a non-negative integer." });
  }
  if (!isValidPrice(row[2])) {
    issues.push({ row: rowIndex, column: DEMO_TABLE_HEADERS[2], message: "Price must be a non-negative number." });
  }

  return issues;
}

export function validateRows(rows: unknown[][]): ValidationIssue[] {
  return rows.reduce<ValidationIssue[]>((all, row, index) => all.concat(validateRow(row, index)), []);
}

/**
 * Validates the payload of an imported JSON file.
 *
 * @throws Error when the payload does not match the expected shape.
 */
export function assertValidExport(payload: unknown): asserts payload is WorkbookExport {
  if (typeof payload !== "object" || payload === null) {
    throw new Error("Imported JSON must be an object.");
  }

  const candidate = payload as Partial<WorkbookExport>;

  if (typeof candidate.version !== "number") {
    throw new Error("Imported JSON is missing a numeric 'version' property.");
  }
  if (!Array.isArray(candidate.headers) || candidate.headers.length === 0) {
    throw new Error("Imported JSON is missing the 'headers' array.");
  }
  if (!Array.isArray(candidate.rows)) {
    throw new Error("Imported JSON is missing the 'rows' array.");
  }
  if (candidate.rows.some((row) => !Array.isArray(row) || row.length !== candidate.headers!.length)) {
    throw new Error("Every imported row must have the same length as the headers.");
  }
}
