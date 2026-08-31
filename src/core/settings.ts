/**
 * Persists user data in the document using the Office settings API.
 */

export const USER_SETTINGS_KEY = "excel-addin-test:user-settings";

export interface UserSettings {
  displayName: string;
  quantityThreshold: number;
  highlightColumn: string;
}

export const DEFAULT_USER_SETTINGS: UserSettings = {
  displayName: "",
  quantityThreshold: 50,
  highlightColumn: "Price",
};

/**
 * Reads the stored settings, falling back to the defaults for missing values.
 */
export function loadUserSettings(): UserSettings {
  const stored = Office.context.document.settings.get(USER_SETTINGS_KEY) as Partial<UserSettings> | null;

  if (!stored || typeof stored !== "object") {
    return { ...DEFAULT_USER_SETTINGS };
  }

  return { ...DEFAULT_USER_SETTINGS, ...stored };
}

/**
 * Stores the settings in the document and persists them asynchronously.
 */
export async function saveUserSettings(settings: UserSettings): Promise<void> {
  Office.context.document.settings.set(USER_SETTINGS_KEY, settings);

  await new Promise<void>((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result && result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error ? result.error.message : "Could not save the settings."));
        return;
      }

      resolve();
    });
  });
}
