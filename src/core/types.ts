/**
 * Shared types used by the task pane services.
 */

export interface DemoRow {
  product: string;
  quantity: number;
  price: number;
}

export interface WorkbookExport {
  version: number;
  tableName: string;
  headers: string[];
  rows: (string | number)[][];
}

export interface ValidationIssue {
  row: number;
  column: string;
  message: string;
}

export const DEMO_TABLE_NAME = "DemoTable";

export const DEMO_TABLE_HEADERS = ["Product", "Quantity", "Price"];

export const EXPORT_VERSION = 1;
