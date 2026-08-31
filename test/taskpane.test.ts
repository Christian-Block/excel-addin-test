import * as fs from "fs";
import * as path from "path";
import { installExcelMock } from "./mocks/excelMock";
import { installOfficeMock } from "./mocks/officeMock";
import { AuthService } from "../src/auth/authService";

const TASKPANE_HTML = fs.readFileSync(path.join(__dirname, "..", "src", "taskpane", "taskpane.html"), "utf8");

function loadTaskpaneDom(): void {
  const body = TASKPANE_HTML.substring(TASKPANE_HTML.indexOf("<main"), TASKPANE_HTML.indexOf("</body>"));
  document.body.innerHTML = body;
}

function byId<T extends HTMLElement>(id: string): T {
  return document.getElementById(id) as T;
}

async function flush(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

function importTaskpane(): typeof import("../src/taskpane/taskpane") {
  let taskpane: typeof import("../src/taskpane/taskpane");
  jest.isolateModules(() => {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    taskpane = require("../src/taskpane/taskpane");
  });
  return taskpane!;
}

describe("task pane", () => {
  let auth: AuthService;

  beforeEach(() => {
    installExcelMock();
    installOfficeMock(null);
    loadTaskpaneDom();
    auth = new AuthService();
    importTaskpane().initializeUi(auth);
  });

  it("keeps the premium controls disabled while signed out", () => {
    expect(byId<HTMLButtonElement>("column-format").disabled).toBe(true);
    expect(byId<HTMLInputElement>("import-file").disabled).toBe(true);
    expect(byId("account-state").textContent).toBe("Not signed in.");
  });

  it("enables the premium controls after signing in", async () => {
    byId("sign-in").click();
    await flush();

    expect(auth.isSignedIn).toBe(true);
    expect(byId<HTMLButtonElement>("column-format").disabled).toBe(false);
    expect(byId<HTMLButtonElement>("save-settings").disabled).toBe(false);
    expect(byId("status").textContent).toBe("Signed in.");
  });

  it("disables the premium controls again after signing out", async () => {
    byId("sign-in").click();
    await flush();
    byId("sign-out").click();
    await flush();

    expect(byId<HTMLButtonElement>("column-format").disabled).toBe(true);
    expect(byId("status").textContent).toBe("Signed out.");
  });

  it("creates the demo table", async () => {
    byId("create-table").click();
    await flush();

    expect(byId("status").textContent).toBe("Demo table created.");
  });

  it("reads and writes a cell", async () => {
    const excelMock = installExcelMock();
    excelMock.context.workbook.worksheets.worksheet.range.values = [["Pencil"]];

    byId<HTMLInputElement>("cell-address").value = "A2";
    byId("read-cell").click();
    await flush();

    expect(byId<HTMLInputElement>("cell-value").value).toBe("Pencil");

    byId<HTMLInputElement>("cell-value").value = "Eraser";
    byId("write-cell").click();
    await flush();

    expect(excelMock.context.workbook.worksheets.worksheet.range.values).toEqual([["Eraser"]]);
  });

  it("lists the validation issues of the table", async () => {
    installExcelMock([["", -1, 2]]);

    byId("validate").click();
    await flush();

    expect(byId("validation-issues").children).toHaveLength(2);
    expect(byId("status").textContent).toBe("2 invalid value(s) found.");
  });

  it("reports errors in the status line", async () => {
    (globalThis as any).Excel.run = async () => {
      throw new Error("Excel is not available.");
    };

    byId("create-table").click();
    await flush();

    expect(byId("status").textContent).toBe("Excel is not available.");
  });
});
