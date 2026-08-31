import { defineConfig, devices } from "@playwright/test";

/**
 * Playwright configuration for the integration test that runs the add-in in
 * Excel for the web. See e2e/README.md for the required environment
 * variables.
 */
export default defineConfig({
  testDir: "./e2e",
  timeout: 5 * 60 * 1000,
  expect: { timeout: 60 * 1000 },
  fullyParallel: false,
  workers: 1,
  retries: process.env.CI ? 1 : 0,
  reporter: [["list"], ["html", { open: "never" }]],
  use: {
    baseURL: "https://www.office.com",
    trace: "retain-on-failure",
    screenshot: "only-on-failure",
    video: "retain-on-failure",
    storageState: process.env.E2E_STORAGE_STATE,
  },
  projects: [
    {
      name: "chromium",
      use: { ...devices["Desktop Chrome"] },
    },
  ],
});
