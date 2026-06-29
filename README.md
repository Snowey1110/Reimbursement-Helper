# Reimbursement Helper

Lightweight Windows desktop app for preparing VisionNav reimbursement workbooks from receipt screenshots.

The app supports:

- bulk receipt image upload
- folder intake from `Unprocessed/` with finished files moved to `Processed/`
- one combined receipt/details manager section
- receipt preview beside editable fields
- manual receipt image cropping with four draggable corner points and revert
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

## Folder Workflow

1. Put receipt images into `Unprocessed/`.
2. Open the app and click `Upload Folder`.
3. Review or generate details.
4. Files that successfully finish AI processing move to `Processed/`.
5. Any remaining loaded files also move to `Processed/` after a successful Excel export.

If `Unprocessed/` is empty, the helper shows the folder path and can open the folder for you. Check `Do not show this again` on that prompt to keep it quiet on your computer.

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

Highlight multiple receipt rows, then edit a field such as `Project number`, `Category`, or `Currency`. The edited field applies to every highlighted row while each row keeps its own receipt image and other details.

## Receipt Cropping

Each preview image has four draggable crop points. Drag the corners to keep only the receipt area before export. `Revert crop` resets the selected receipt, and cropped images are used in the generated Excel workbook.

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

- Receipt upload currently supports image files: PNG, JPG, JPEG, WEBP, BMP, and GIF.
- Korea output is one workbook with three sheets: cover, details, and receipts.
- USA output preserves the provided workbook format and leaves proof-of-payment cells blank.
