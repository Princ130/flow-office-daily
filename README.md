# CashFlow Office

A browser-based, local-first cash register prototype for GitHub Pages. The Excel workbook is the persistent data store; no database or backend server is used.

## Run locally

Because the app uses JavaScript ES modules, do not open `index.html` directly as a `file://` URL. Use a local static server, for example:

```bash
python -m http.server 8000
```

Then visit `http://localhost:8000`.

## Deploy to GitHub Pages

1. Create a GitHub repository.
2. Upload the project files and folders (do not upload real office workbooks).
3. In repository Settings → Pages, deploy from the `main` branch and `/ (root)`.
4. Open the published Pages URL.

## Workbook workflow

- Choose **Open existing Excel workbook** to load a `.xlsx` file.
- Choose **Create a new workbook** to start with an empty in-browser workbook.
- Start each day by confirming physical opening cash/coins and entering the three platform limits.
- Record transactions. Cash movement is applied only when Payment Mode is `Cash`.
- Save after each transaction. On supported Chrome/Edge browsers, the selected workbook can be saved directly if opened through the File System Access API. Otherwise, the app downloads an updated `.xlsx`; replace the previous workbook with that download before continuing.

## Important prototype notes

- Keep the workbook in a private office folder. Never commit it to GitHub.
- This version is intended for one active operator/workbook at a time. Do not let multiple computers edit separate copies simultaneously.
- The platform-limit rule currently uses the required amount for customer receipts and payout amount for send transactions. Confirm this rule with the employer before using live accounts.
- The app has no authentication or encryption. Do not expose it to public/shared devices with sensitive data.
- The daily serial is assigned in insertion order. The transaction ID is unique across workbook records.
- A closed day is intended to be read-only; the current prototype displays the closed state but does not yet implement a full admin reopening workflow.
- Back up the workbook daily to a separate drive.

## Folder structure

- `css/` stylesheets
- `js/` ES modules and application logic
- `assets/` optional images, icons and fonts
- `templates/` workbook template
- `docs/` operating and workbook documentation
