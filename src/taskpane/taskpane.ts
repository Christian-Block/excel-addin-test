import { AuthService } from "../auth/authService";
import { addColumnConditionalFormat, addRowConditionalFormat } from "../core/formatting";
import { downloadJson, exportTable, importJson, toJson } from "../core/jsonData";
import { DEFAULT_USER_SETTINGS, loadUserSettings, saveUserSettings, UserSettings } from "../core/settings";
import { createDemoTable, readCell, readTableRows, writeCell } from "../core/tableService";
import { validateRows } from "../core/validation";
import { getValueMap } from "../functions/valueMap";

const PREMIUM_ELEMENT_IDS = ["column-format", "import-file", "display-name", "quantity-threshold", "save-settings"];

function byId<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);

  if (!element) {
    throw new Error(`Element '${id}' is missing in the task pane.`);
  }

  return element as T;
}

function setStatus(message: string): void {
  byId("status").textContent = message;
}

function showError(error: unknown): void {
  setStatus(error instanceof Error ? error.message : String(error));
}

/**
 * Wraps a click handler so that every failure is reported in the status line.
 */
function onClick(id: string, handler: () => Promise<void> | void): void {
  byId(id).addEventListener("click", async () => {
    try {
      await handler();
    } catch (error) {
      showError(error);
    }
  });
}

export function updateFeatureAvailability(auth: AuthService): void {
  for (const id of PREMIUM_ELEMENT_IDS) {
    (byId(id) as HTMLButtonElement | HTMLInputElement).disabled = !auth.isSignedIn;
  }

  byId("account-state").textContent = auth.isSignedIn ? "Signed in." : "Not signed in.";
}

function renderValidationIssues(rows: (string | number)[][]): void {
  const list = byId<HTMLUListElement>("validation-issues");
  const issues = validateRows(rows);

  list.textContent = "";

  if (issues.length === 0) {
    setStatus("All values are valid.");
    return;
  }

  for (const issue of issues) {
    const item = document.createElement("li");
    item.textContent = `Row ${issue.row + 1}, ${issue.column}: ${issue.message}`;
    list.appendChild(item);
  }

  setStatus(`${issues.length} invalid value(s) found.`);
}

function readSettingsFromForm(): UserSettings {
  const threshold = Number(byId<HTMLInputElement>("quantity-threshold").value);

  return {
    ...DEFAULT_USER_SETTINGS,
    displayName: byId<HTMLInputElement>("display-name").value,
    quantityThreshold: Number.isFinite(threshold) ? threshold : DEFAULT_USER_SETTINGS.quantityThreshold,
  };
}

function applySettingsToForm(settings: UserSettings): void {
  byId<HTMLInputElement>("display-name").value = settings.displayName;
  byId<HTMLInputElement>("quantity-threshold").value = String(settings.quantityThreshold);
}

/**
 * Binds every task pane control. Exported for unit testing.
 */
export function initializeUi(auth: AuthService = new AuthService()): AuthService {
  byId("app").classList.remove("hidden");
  byId("map-size").textContent = `${Object.keys(getValueMap()).length} entries in the built in value map.`;

  onClick("sign-in", async () => {
    await auth.signIn();
    updateFeatureAvailability(auth);
    applySettingsToForm(loadUserSettings());
    setStatus("Signed in.");
  });

  onClick("sign-out", () => {
    auth.signOut();
    updateFeatureAvailability(auth);
    setStatus("Signed out.");
  });

  onClick("create-table", async () => {
    await createDemoTable();
    setStatus("Demo table created.");
  });

  onClick("read-cell", async () => {
    const address = byId<HTMLInputElement>("cell-address").value;
    const value = await readCell(address);
    byId<HTMLInputElement>("cell-value").value = String(value);
    setStatus(`Read ${address}.`);
  });

  onClick("write-cell", async () => {
    const address = byId<HTMLInputElement>("cell-address").value;
    await writeCell(address, byId<HTMLInputElement>("cell-value").value);
    setStatus(`Wrote ${address}.`);
  });

  onClick("validate", async () => {
    renderValidationIssues(await readTableRows());
  });

  onClick("row-format", async () => {
    await addRowConditionalFormat(loadUserSettings().quantityThreshold);
    setStatus("Rows highlighted.");
  });

  onClick("column-format", async () => {
    auth.assertFeatureAvailable("column-conditional-format");
    await addColumnConditionalFormat(loadUserSettings().highlightColumn);
    setStatus("Colour scale added.");
  });

  onClick("export-json", async () => {
    downloadJson(toJson(await exportTable()));
    setStatus("JSON downloaded.");
  });

  byId("import-file").addEventListener("change", async (event) => {
    try {
      auth.assertFeatureAvailable("json-import");

      const file = (event.target as HTMLInputElement).files?.[0];

      if (!file) {
        return;
      }

      const imported = await importJson(await file.text());
      setStatus(`Imported ${imported.rows.length} row(s).`);
    } catch (error) {
      showError(error);
    }
  });

  onClick("save-settings", async () => {
    auth.assertFeatureAvailable("user-settings");
    await saveUserSettings(readSettingsFromForm());
    setStatus("Settings saved.");
  });

  updateFeatureAvailability(auth);

  return auth;
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initializeUi();
  }
});
