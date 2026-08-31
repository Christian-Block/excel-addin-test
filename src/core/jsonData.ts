import { readTableRows, writeTableRows } from "./tableService";
import { DEMO_TABLE_HEADERS, DEMO_TABLE_NAME, EXPORT_VERSION, WorkbookExport } from "./types";
import { assertValidExport } from "./validation";

/**
 * Reads the demo table and returns its content as a serialisable object.
 */
export async function exportTable(): Promise<WorkbookExport> {
  const rows = await readTableRows();

  return {
    version: EXPORT_VERSION,
    tableName: DEMO_TABLE_NAME,
    headers: [...DEMO_TABLE_HEADERS],
    rows,
  };
}

/**
 * Serialises the export payload to a formatted JSON string.
 */
export function toJson(data: WorkbookExport): string {
  return JSON.stringify(data, null, 2);
}

/**
 * Parses and validates JSON produced by {@link toJson}.
 *
 * @throws Error when the JSON is malformed or does not match the schema.
 */
export function parseJson(json: string): WorkbookExport {
  let payload: unknown;

  try {
    payload = JSON.parse(json);
  } catch {
    throw new Error("The selected file does not contain valid JSON.");
  }

  assertValidExport(payload);

  return payload;
}

/**
 * Triggers a browser download for the given JSON string.
 *
 * @returns The object URL that was created, so callers can revoke it in tests.
 */
export function downloadJson(json: string, fileName = "excel-addin-export.json"): string {
  const blob = new Blob([json], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);

  return url;
}

/**
 * Imports previously exported JSON back into the demo table.
 */
export async function importJson(json: string): Promise<WorkbookExport> {
  const payload = parseJson(json);

  await writeTableRows(payload.rows);

  return payload;
}
