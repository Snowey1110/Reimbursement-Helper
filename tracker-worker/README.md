# Report Tracker Worker

This optional Cloudflare Worker counts successful web report exports. It records only:

- form version: `USA` or `Korea`
- UI language: `en`, `zh`, or `ko`
- server-side day
- count

It does not receive receipt images, workbook files, amounts, vendors, API keys, applicant names, or project numbers.

## Free Tier Fit

The tracker is designed for Cloudflare's free tier. One generated report creates one Worker request and one D1 row write/update.

## Deploy

1. Create or log in to a free Cloudflare account.
2. From this folder:

```bash
npm install
npx wrangler login
npx wrangler d1 create reimbursement_report_tracker
```

3. Copy `wrangler.toml.example` to `wrangler.toml`.
4. Put the returned D1 `database_id` into `wrangler.toml`.
5. Apply the database migration and deploy:

```bash
npx wrangler d1 migrations apply reimbursement_report_tracker --remote
npx wrangler deploy
```

6. Copy the deployed Worker URL, usually like:

```text
https://reimbursement-helper-tracker.<your-account>.workers.dev
```

7. In `web/public/tracker-config.json`, set:

```json
{
  "enabled": true,
  "endpoint": "https://reimbursement-helper-tracker.<your-account>.workers.dev"
}
```

8. Commit and push the updated `tracker-config.json`.

## Endpoints

- `POST /track`
- `GET /stats`
- `GET /health`

Example stats:

```json
{
  "total": 12,
  "byForm": { "USA": 7, "Korea": 5 },
  "byLanguage": { "en": 4, "zh": 8 },
  "recentDays": [{ "day": "2026-07-02", "count": 12 }]
}
```
