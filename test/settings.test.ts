import { installOfficeMock } from "./mocks/officeMock";
import {
  DEFAULT_USER_SETTINGS,
  loadUserSettings,
  saveUserSettings,
  USER_SETTINGS_KEY,
} from "../src/core/settings";

describe("user settings", () => {
  it("returns the defaults when nothing is stored", () => {
    installOfficeMock(null);

    expect(loadUserSettings()).toEqual(DEFAULT_USER_SETTINGS);
  });

  it("merges the stored values with the defaults", () => {
    installOfficeMock({ displayName: "Ada" });

    expect(loadUserSettings()).toEqual({ ...DEFAULT_USER_SETTINGS, displayName: "Ada" });
  });

  it("stores and persists the settings", async () => {
    const officeMock = installOfficeMock(null);
    const settings = { ...DEFAULT_USER_SETTINGS, displayName: "Grace", quantityThreshold: 10 };

    await saveUserSettings(settings);

    expect(officeMock.context.document.settings.lastSetKey).toBe(USER_SETTINGS_KEY);
    expect(officeMock.context.document.settings.value).toEqual(settings);
    expect(officeMock.context.document.settings.saved).toBe(true);
  });

  it("rejects when the host cannot save the settings", async () => {
    const officeMock = installOfficeMock(null);
    officeMock.context.document.settings.shouldFail = true;

    await expect(saveUserSettings(DEFAULT_USER_SETTINGS)).rejects.toThrow("Save failed.");
  });
});
