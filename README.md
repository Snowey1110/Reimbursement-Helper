# Reimbursement Helper

Lightweight Windows desktop app for preparing VisionNav reimbursement workbooks from receipt screenshots.

The app supports:

- bulk receipt image upload
- receipt preview beside editable fields
- manual entry when AI is not needed
- one-click AI extraction when an OpenAI API key is available locally
- USA and Korea Excel outputs from stored blank templates
- private local defaults kept out of GitHub

## Privacy Defaults

The shared project does not store personal applicant information. The USA template keeps `申请人` blank by default, and the app writes `Employee: / 申请人：` without a name.

Local private files are ignored by git:

- `config/user_settings.json`
- `config/api_key.txt`
- `.env.local`
- `outputs/`
- `receipt_uploads/`

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

## Templates

Blank templates are stored in `templates/`.

- `usa_expense_report_template.xlsx`
- `korea_cover_receipts_template.xlsx`
- `korea_details_template.xlsx`

The app copies these templates on export and fills the copy only. Original templates are not modified during normal use.

## Notes

- Receipt upload currently supports image files: PNG, JPG, JPEG, WEBP, BMP, and GIF.
- Korea output is one workbook with three sheets: `境外同事报销使用`, `报销明细`, and `发票`.
- USA output preserves the provided workbook format and leaves proof-of-payment cells blank.
