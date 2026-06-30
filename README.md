# Reimbursement Helper

Reimbursement Helper prepares VisionNav USA and Korea reimbursement workbooks from receipt screenshots, PDFs, and USA payment proof images.

There are two ways to use it:

- **Desktop app**: best for daily work on a Windows computer.
- **Web app**: runs from GitHub Pages and uses the API key typed by the user in the browser.

Open the web app here:

https://snowey1110.github.io/Reimbursement-Helper/

## Quick Start For Users

1. Download the latest release zip from the GitHub **Releases** page.
2. Extract the zip.
3. Double-click `run_reimbursement_helper.bat`.
4. Click the blue suggested button and follow the flow:
   - `Select Files`
   - USA only: `Select Payment Proof`
   - Korea only: `Select 汇率 Image` for the exchange-rate screenshot
   - `Generate All`
   - `Generate Excel`

The app leaves `申请人` / applicant blank and does not include personal defaults in the shared project.

The app supports:

- bulk receipt image/PDF selection
- additive uploads so more files can be selected later without clearing the list
- one combined receipt/details manager section
- receipt preview beside editable fields
- receipt screenshot rotation for sideways images
- manual receipt image perspective cropping with four independent corner points and revert
- USA payment proof images for the proof-of-payment column
- manual entry when AI is not needed
- one-click AI extraction when an OpenAI API key is available locally
- live USD/RMB and Korea original-currency-to-KRW/RMB conversion using editable rate fields
- currency correction for Korea receipts when AI detects the wrong receipt currency
- multi-row editing for highlighted receipts, useful for shared fields like project number
- USA and Korea Excel outputs from stored blank templates
- local error logs for debugging without sharing private files
- saved working session with restore prompt on reopen

## Privacy Defaults

The shared project does not store personal applicant information. The USA template keeps the applicant field blank by default, and the app exports it blank unless a local private config is added later.

Local private files are ignored by git:

- `config/user_settings.json`
- `config/api_key.txt`
- `.env.local`
- `outputs/`
- `receipt_uploads/`
- `work/`
- `logs/`
- `Unprocessed/`
- `Processed/`

Copy `config/user_settings.example.json` to `config/user_settings.json` if you want private defaults on your own computer.

## Run

```bash
python -m pip install -r requirements.txt
python reimbursement_helper.py
```

Or double-click:

```text
run_reimbursement_helper.bat
```

## Web Version

A static GitHub Pages version is in `web/`. It runs in the browser, uses blank templates from `web/public/templates/`, and asks each user to enter their own OpenAI API key. The key is not committed; by default it is kept only for the browser session, with an optional checkbox to remember it on that device.

Live site:

https://snowey1110.github.io/Reimbursement-Helper/

```bash
cd web
pnpm install
pnpm dev
```

Build and test:

```bash
cd web
pnpm test
pnpm build
```

GitHub Pages deployment is handled by `.github/workflows/deploy-web.yml`. On `main` or `master`, changes under `web/` are tested, built, and published from `web/dist`.

## GitHub Pages Setup

After pushing this repository to GitHub:

1. Open the GitHub repository.
2. Go to **Settings**.
3. Go to **Pages**.
4. Under **Build and deployment**, set **Source** to **GitHub Actions**.
5. Push to `master` or run the **Deploy web app to GitHub Pages** workflow manually from the **Actions** tab.

The repository README is the main GitHub project page for coworkers. The GitHub Pages site is the browser app.

## Release Download

To publish a desktop download on GitHub:

1. Create a zip from the project files or use the prepared release zip.
2. Open the GitHub repository.
3. Go to **Releases**.
4. Click **Draft a new release**.
5. Use tag `v0.1.0` or the next version number.
6. Upload the zip file.
7. Publish the release.

## File Workflow

1. Open the app and click `Select Files`.
2. Choose one or more receipt images or PDFs.
3. Review or generate details.
4. Click `Select Files` again later to add more receipts without clearing the current list.
5. Each PDF is imported as one combined image with page 1, page 2, and later pages stacked together.
6. For USA reports, click `Select Payment Proof` separately to add card or bank proof screenshots/PDFs.

## AI Extraction

AI extraction looks for an API key in this order:

1. `OPENAI_API_KEY` environment variable
2. `.env.local`
3. `.env`
4. `config/api_key.txt`
5. the saved Daily Logger API key on this PC

Do not commit real keys. The files above are ignored by `.gitignore`.

## Rates

The app has editable exchange-rate fields. Current defaults are:

- USD to RMB: `6.8175`
- KRW to RMB: `0.004433`

For Korea, `Original amount` plus `Currency` fills the KRW and RMB charged amounts. Change `Currency` to the receipt's real currency, such as `USD`, and the converted fields update immediately. Update the visible rate field before export whenever the reimbursement department needs a different rate.

## Bulk Edits

Highlight multiple receipt rows, then edit a field such as `Project number`, `Category`, or `Currency`. The edited field applies to every highlighted row while each row keeps its own receipt image and other details. Click inside the receipt list and press `Ctrl+A` to select every row.

## USA Payment Proof

For the USA form, use `Select Files` for receipts and `Select Payment Proof` for card or bank proof images. `Generate All` reads both groups, merges receipt screenshots that have the same date and USD amount, and matches one proof image by date and charge amount.

The preview shows receipt screenshots on the left. After payment proof files are loaded, a Payment proof section appears on the right. Use the divider swap icon to switch proof images when another proof is available. If a payment proof was added as a receipt by mistake, drag that screenshot tile into the Payment proof section. Select any screenshot tile and use the delete icon to remove only that image. In the receipt list, press `Delete` or `Backspace` to remove selected receipt rows after confirmation.

## Receipt Cropping

Use the rotate icons when a screenshot is sideways. Each preview image also has four independent draggable crop points. Drag the corners to match the receipt edges; the app straightens that shape into a rectangle for the generated Excel workbook. The revert icon resets the selected screenshot back to its original crop and rotation.

## Logs

Errors are written to `logs/app.log` with rotating backups. The `logs/` folder is ignored by git, so coworkers do not receive your local debugging history.

## Restore

The app saves the current working session when it closes. On the next launch it asks whether to restore the previous session. The recovery file is `config/session_state.json`, and it is ignored by git.

## Templates

Blank templates are stored in `templates/`.

- `usa_expense_report_template.xlsx`
- `korea_cover_receipts_template.xlsx`
- `korea_details_template.xlsx`

The app copies these templates on export and fills the copy only. Original templates are not modified during normal use.

## Notes

- Receipt upload currently supports PNG, JPG, JPEG, WEBP, BMP, GIF, and PDF.
- Korea output is one workbook with three sheets: cover, details, and receipts.
- Korea receipts are grouped by payment on the `发票` sheet, with a compact payment label directly above each larger full-width receipt image. Optional Korea `汇率` screenshots are placed in the first `发票` image slot; overflow continues on the next printed page below.
- USA output preserves the provided workbook format and puts matched payment proof images in column E.
