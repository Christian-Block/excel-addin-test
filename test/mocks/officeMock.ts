import { OfficeMockObject } from "office-addin-mock";

/**
 * Mock data for the Office common API used to store user data.
 */
export function installOfficeMock(storedValue: unknown = null): any {
  const officeMock = new OfficeMockObject({
    AsyncResultStatus: {
      Succeeded: "succeeded",
      Failed: "failed",
    },
    HostType: {
      Excel: "Excel",
    },
    context: {
      document: {
        settings: {
          value: "" as unknown,
          saved: false,
          shouldFail: false,
          lastGetKey: "",
          lastSetKey: "",
          get: function (key: string) {
            this.lastGetKey = key;
            return this.value;
          },
          set: function (key: string, value: unknown) {
            this.lastSetKey = key;
            this.value = value;
          },
          saveAsync: function (callback: (result: unknown) => void) {
            this.saved = true;
            callback({
              status: this.shouldFail ? "failed" : "succeeded",
              error: { message: "Save failed." },
            });
          },
        },
      },
    },
    auth: {
      token: "access-token",
      getAccessToken: async function () {
        return this.token;
      },
    },
    onReady: function (callback: (info: { host: string }) => void) {
      this.readyCallback = callback;
    },
    actions: {
      associate: function (name: string, handler: unknown) {
        this.associatedName = name;
        this.associatedHandler = handler;
      },
    },
  }) as any;

  // Scalar values of the mock have to be assigned explicitly, otherwise the
  // library reports them as "not loaded".
  officeMock.AsyncResultStatus.Succeeded = "succeeded";
  officeMock.AsyncResultStatus.Failed = "failed";
  officeMock.HostType.Excel = "Excel";
  officeMock.auth.token = "access-token";
  officeMock.context.document.settings.shouldFail = false;
  officeMock.context.document.settings.value = storedValue;
  (globalThis as any).Office = officeMock;

  return officeMock;
}
