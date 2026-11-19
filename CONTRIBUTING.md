# Contributing

## Rules that keep "no network" true
- Third-party code goes only into the webview/worker bundle, never into `src/host`.
- Do not loosen the CSP in `src/host/webviewHtml.ts` (a unit test pins it).
- New hits in `scripts/allowed-hits.json` need a written review reason.

Run commands one per line: Windows PowerShell 5.1 does not support `&&` (use `;` or separate lines).

## Upgrading a dependency
1. `npm diff --diff=<pkg>@<old> --diff=<pkg>@<new>` and read the diff.
2. `npm audit` (the exceljs `uuid` advisory does not apply: exceljs ships a prebuilt browser bundle and `uuid` is only called without a buffer).
3. Install with `npm install --save-exact <pkg>@<new>` (scripts are disabled by `.npmrc`).
4. `npm run check` and `npm run test:integration`; review new bundle-scan hits.
5. `npm run notices` and commit `THIRD_PARTY_NOTICES.md`.

## Tests
`npm test` (unit, network stubbed), `npm run test:integration` (real VS Code). `npm run fixtures` rewrites `test/fixtures/` from `test/helpers/fixtures.ts`. Real Excel/Word/LibreOffice files are part of the manual checklist before a release: open, edit, save, reopen in the original application.

## Running tests on Windows

Use `npm test`, not `npx vitest`. The launcher `scripts/vitest.mjs` fixes the drive-letter case of the working directory; plain `npx vitest` fails with "Cannot read properties of undefined (reading 'config')" when the terminal is in `d:\...` (lower case).
