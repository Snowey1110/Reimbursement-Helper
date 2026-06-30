# Reimbursement Helper Web

Static GitHub Pages app for building USA and Korea reimbursement Excel reports in the browser.

## Privacy

- API keys are typed by the user in the browser.
- Session storage is the default.
- The optional "Remember on this device" checkbox stores the key in that browser's local storage.
- No applicant name, personal default, or API key is stored in this project.

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
