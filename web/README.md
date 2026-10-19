# Reimbursement Helper Web

Static GitHub Pages app for building USA and Korea reimbursement Excel reports in the browser.

## Privacy

- API keys are typed by the user in the browser.
- Session storage is the default.
- The optional "Remember on this device" checkbox stores the key in that browser's local storage.
- No applicant name, personal default, or API key is stored in this project.

## Report Recovery

- The current report, receipt/payment-proof images, exchange-rate images, and edited rates are autosaved in this browser's IndexedDB after a short pause in editing. API keys are not part of the report draft.
- Refreshing restores the saved report's form and exact edited amounts. URL form presets apply when starting without a saved draft.
- Report controls wait for recovery to finish so a slow read cannot overwrite new work. If recovery fails, the existing draft is left untouched; reload to retry, or choose **Clear** to discard it and start again.
- **Clear** removes the stored report. Removing all report contents also removes the draft.
- Drafts stay on this browser and site address; they do not sync across devices or between the custom domain and GitHub Pages. Clearing site data, browser storage eviction, or private browsing can remove them. Export important reports rather than relying on autosave as a backup.

## Run Locally

```bash
pnpm install
pnpm dev
```

## Test And Build

```bash
pnpm test
pnpm build
```

## GitHub Pages

The repository workflow builds this folder and deploys `web/dist` to Pages. In GitHub, enable Pages with "GitHub Actions" as the source.
