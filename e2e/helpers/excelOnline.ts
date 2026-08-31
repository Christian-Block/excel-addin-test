import { expect, FrameLocator, Page } from "@playwright/test";
import { generateTotp } from "./totp";

export interface M365Credentials {
  username: string;
  password: string;
  totpSecret?: string;
}

/**
 * Reads the credentials of the test account from the environment. The tests
 * are skipped when the credentials are not configured.
 */
export function readCredentials(): M365Credentials | undefined {
  const username = process.env.E2E_M365_USERNAME;
  const password = process.env.E2E_M365_PASSWORD;

  if (!username || !password) {
    return undefined;
  }

  return { username, password, totpSecret: process.env.E2E_M365_TOTP_SECRET };
}

/**
 * Signs in to Microsoft 365 including the multi factor authentication step.
 *
 * Three MFA variants are handled:
 * - a verification code from an authenticator app (`E2E_M365_TOTP_SECRET`)
 * - the "number matching" prompt of the Microsoft Authenticator push
 *   notification, which only has to be confirmed on the phone
 * - accounts without MFA, where the prompt is simply not shown
 */
export async function signIn(page: Page, credentials: M365Credentials): Promise<void> {
  await page.getByRole("textbox", { name: /email|phone|username/i }).fill(credentials.username);
  await page.getByRole("button", { name: /next/i }).click();

  await page.getByRole("textbox", { name: /password/i }).fill(credentials.password);
  await page.getByRole("button", { name: /sign in/i }).click();

  await completeMultiFactorAuthentication(page, credentials);

  // "Stay signed in?" - answering with yes keeps the storage state usable for
  // the following tests.
  const staySignedIn = page.getByRole("button", { name: /^yes$/i });

  if (await staySignedIn.isVisible({ timeout: 15_000 }).catch(() => false)) {
    await staySignedIn.click();
  }
}

async function completeMultiFactorAuthentication(page: Page, credentials: M365Credentials): Promise<void> {
  const codeInput = page.getByRole("textbox", { name: /code/i });
  const numberMatching = page.locator("#idRichContext_DisplaySign");

  await Promise.race([
    codeInput.waitFor({ state: "visible", timeout: 30_000 }).catch(() => undefined),
    numberMatching.waitFor({ state: "visible", timeout: 30_000 }).catch(() => undefined),
    page.waitForURL(/office|sharepoint/i, { timeout: 30_000 }).catch(() => undefined),
  ]);

  if (await numberMatching.isVisible().catch(() => false)) {
    const numberToApprove = await numberMatching.textContent();
    throw new Error(
      `The test account uses push based MFA. Approve the number ${numberToApprove} in the Microsoft Authenticator app ` +
        `or configure 'E2E_M365_TOTP_SECRET' to use verification codes instead.`
    );
  }

  if (await codeInput.isVisible().catch(() => false)) {
    if (!credentials.totpSecret) {
      throw new Error("The account requires a verification code; set 'E2E_M365_TOTP_SECRET' to generate it.");
    }

    await codeInput.fill(generateTotp(credentials.totpSecret));
    await page.getByRole("button", { name: /verify|sign in|next/i }).click();
  }
}

/**
 * Creates a new, empty workbook in Excel for the web.
 */
export async function createBlankWorkbook(page: Page): Promise<void> {
  await page.goto("https://www.office.com/launch/excel");
  await page.getByRole("button", { name: /blank workbook|new blank workbook/i }).click();
  await expect(page.frameLocator("iframe[name='WacFrame_Excel_0']").locator("body")).toBeVisible({
    timeout: 120_000,
  });
}

function excelFrame(page: Page): FrameLocator {
  return page.frameLocator("iframe[name='WacFrame_Excel_0']");
}

/**
 * Side loads the add-in through "Home > Add-ins > Upload My Add-in".
 *
 * @param manifestPath Absolute path of the manifest that is uploaded.
 */
export async function sideloadAddin(page: Page, manifestPath: string): Promise<void> {
  const excel = excelFrame(page);

  await excel.getByRole("tab", { name: /home/i }).click();
  await excel.getByRole("button", { name: /add-ins/i }).click();
  await page.getByRole("link", { name: /upload my add-in/i }).click();

  const fileChooser = page.waitForEvent("filechooser");
  await page.getByRole("button", { name: /browse/i }).click();
  await (await fileChooser).setFiles(manifestPath);
  await page.getByRole("button", { name: /upload/i }).click();
}

/**
 * Returns the frame of the task pane of the add-in.
 */
export function taskpaneFrame(page: Page): FrameLocator {
  return excelFrame(page).frameLocator("iframe[title*='Excel add-in'], iframe[name*='taskpane']");
}
