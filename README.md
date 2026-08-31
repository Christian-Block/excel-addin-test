# excel-addin-test

Demo of an Excel web add-in with a task pane, custom functions, unit tests and a
GitHub Pages deployment.

## Features

### Task pane

| Feature | Sign in required | Implementation |
| --- | --- | --- |
| Create the demo table | no | `src/core/tableService.ts` |
| Read and write single cells | no | `src/core/tableService.ts` |
| Validate the table values | no | `src/core/validation.ts` |
| Highlight rows below a quantity threshold (conditional format of a row range) | no | `src/core/formatting.ts` |
| Colour scale for a single column (conditional format of a column) | yes | `src/core/formatting.ts` |
| Download the table as JSON | no | `src/core/jsonData.ts` |
| Import a JSON file back into the table | yes | `src/core/jsonData.ts` |
| Store user data (display name, threshold) in the document | yes | `src/core/settings.ts` |

The free feature area is always usable. The features of the second area are
only enabled after signing in; the sign in uses single sign-on
(`Office.auth.getAccessToken`) and is implemented in `src/auth/authService.ts`.

### Custom functions

The add-in ships a key/value map (`src/functions/valueMap.ts`) that is used to
replace values:

| Function | Description |
| --- | --- |
| `=DEMO.REPLACE("DE")` | Replaces a key with the value of the map. |
| `=DEMO.REPLACERANGE(A1:A10; "n/a")` | Replaces every key of a range. |
| `=DEMO.MAPSIZE()` | Number of entries in the current map. |
| `=DEMO.LOADMAP("https://example.com/map.json")` | Loads the map from a REST endpoint. |

The built in list is the default. `loadValueMapFromEndpoint` already implements
the planned REST based loading, so the list can be replaced at runtime without
changing the custom functions. Both `{ "DE": "Germany" }` and
`[{ "key": "DE", "value": "Germany" }]` are accepted as response payload.

## Development

```bash
npm install     # install the dependencies
npm start       # dev server on https://localhost:3000
npm run build   # production build into dist/
npm run lint    # ESLint
npm run typecheck
npm test        # Jest unit tests
```

Side load `dist/manifest.xml` (development build: it points to
`https://localhost:3000`) to try the add-in locally, see
[Sideload an add-in](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins).

## Tests

### Unit tests

The unit tests use [Jest](https://jestjs.io) together with the official Office
mocking library
[`office-addin-mock`](https://learn.microsoft.com/office/dev/add-ins/testing/unit-testing#install-the-tool).
The mock data for Excel and for the Office common API lives in
`test/mocks/`; every service, the custom functions and the task pane UI are
covered.

```bash
npm test
```

### Integration test with Excel for the web

`e2e/excelOnline.spec.ts` runs the add-in in Excel for the web with
[Playwright](https://playwright.dev): it signs in, creates a workbook, side
loads `dist/manifest.xml` and exercises the task pane and a custom function.

```bash
npm run build
npx playwright install chromium
E2E_M365_USERNAME=user@contoso.com \
E2E_M365_PASSWORD='...' \
E2E_M365_TOTP_SECRET='JBSWY3DPEHPK3PXP' \
npm run test:e2e
```

| Variable | Purpose |
| --- | --- |
| `E2E_M365_USERNAME` | Microsoft 365 account used for the test. |
| `E2E_M365_PASSWORD` | Password of that account. |
| `E2E_M365_TOTP_SECRET` | Base32 secret of the authenticator app of the account. It is used to generate the verification code of the multi factor authentication (`e2e/helpers/totp.ts`). |
| `E2E_STORAGE_STATE` | Optional Playwright storage state file with an existing session, which skips the sign in completely. |

The tests are skipped automatically when no credentials are configured. Accounts
that use push notifications ("number matching") cannot be automated; use an
authenticator app secret or a stored session instead. Never commit the
credentials - store them as repository secrets, they are read by
`.github/workflows/e2e.yml`.

## Deployment

`.github/workflows/deploy.yml` builds the add-in on every push to `main` and
publishes `dist/` to GitHub Pages. Enable Pages with the source "GitHub Actions"
once in the repository settings. The production manifest points to
`https://christian-block.github.io/excel-addin-test/`.
