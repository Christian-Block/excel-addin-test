import { expect, test } from "@playwright/test";
import * as path from "path";
import {
  createBlankWorkbook,
  readCredentials,
  sideloadAddin,
  signIn,
  taskpaneFrame,
} from "./helpers/excelOnline";

const credentials = readCredentials();
const manifestPath = path.join(__dirname, "..", "dist", "manifest.xml");

test.describe("Excel for the web", () => {
  test.skip(
    credentials === undefined,
    "Set E2E_M365_USERNAME, E2E_M365_PASSWORD and (optionally) E2E_M365_TOTP_SECRET to run the integration test."
  );

  test.beforeEach(async ({ page }) => {
    await page.goto("https://www.office.com/launch/excel");

    if (await page.getByRole("textbox", { name: /email|phone|username/i }).isVisible().catch(() => false)) {
      await signIn(page, credentials!);
    }

    await createBlankWorkbook(page);
    await sideloadAddin(page, manifestPath);
  });

  test("runs the free and the protected features of the task pane", async ({ page }) => {
    const taskpane = taskpaneFrame(page);

    await expect(taskpane.getByTestId("status")).toHaveText("Ready.");

    // Free features.
    await taskpane.getByTestId("create-table").click();
    await expect(taskpane.getByTestId("status")).toHaveText("Demo table created.");

    await taskpane.getByTestId("cell-address").fill("A2");
    await taskpane.getByTestId("read-cell").click();
    await expect(taskpane.getByTestId("cell-value")).toHaveValue("Pencil");

    await taskpane.getByTestId("cell-value").fill("Marker");
    await taskpane.getByTestId("write-cell").click();
    await expect(taskpane.getByTestId("status")).toHaveText("Wrote A2.");

    await taskpane.getByTestId("validate").click();
    await expect(taskpane.getByTestId("status")).toHaveText("All values are valid.");

    await taskpane.getByTestId("row-format").click();
    await expect(taskpane.getByTestId("status")).toHaveText("Rows highlighted.");

    // Protected features are disabled until the user signs in.
    await expect(taskpane.getByTestId("column-format")).toBeDisabled();

    await taskpane.getByTestId("sign-in").click();
    await expect(taskpane.getByTestId("account-state")).toHaveText("Signed in.");
    await expect(taskpane.getByTestId("column-format")).toBeEnabled();

    await taskpane.getByTestId("column-format").click();
    await expect(taskpane.getByTestId("status")).toHaveText("Colour scale added.");
  });

  test("downloads the table as JSON", async ({ page }) => {
    const taskpane = taskpaneFrame(page);

    await taskpane.getByTestId("create-table").click();

    const download = page.waitForEvent("download");
    await taskpane.getByTestId("export-json").click();

    expect((await download).suggestedFilename()).toBe("excel-addin-export.json");
  });

  test("replaces values with the custom function", async ({ page }) => {
    const grid = page.frameLocator("iframe[name='WacFrame_Excel_0']");

    await grid.getByRole("textbox", { name: /name box/i }).fill("E1");
    await grid.getByRole("textbox", { name: /name box/i }).press("Enter");
    await page.keyboard.type('=DEMO.REPLACE("DE")');
    await page.keyboard.press("Enter");

    await expect(grid.getByRole("gridcell", { name: "Germany" })).toBeVisible();
  });
});
