# Reimbursement Helper

Lightweight Windows desktop app for preparing VisionNav reimbursement workbooks from receipt screenshots.

The app supports:

- bulk receipt image upload
- one combined receipt/details manager section
- receipt preview beside editable fields
- manual entry when AI is not needed
- one-click AI extraction when an OpenAI API key is available locally
- live USD/RMB and KRW/RMB conversion using editable rate fields
- USA and Korea Excel outputs from stored blank templates
- local error logs for debugging without sharing private files

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

Update the visible rate field before export whenever the reimbursement department needs a different rate.

## Logs

Errors are written to `logs/app.log` with rotating backups. The `logs/` folder is ignored by git, so coworkers do not receive your local debugging history.

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
