import { createDemoTable } from "../core/tableService";

/* global Office */

/**
 * Ribbon command that creates the demo table without opening the task pane.
 */
export async function createTable(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await createDemoTable();
  } finally {
    event.completed();
  }
}

Office.onReady(() => {
  Office.actions.associate("createTable", createTable);
});
