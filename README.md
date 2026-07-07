# Risk Analysis

Offline Cardinal QCx equipment risk analyzer.

## Run locally

```bash
npm install
npm run serve
```

Open `http://127.0.0.1:4179`.

The app processes workbooks in the browser. Workbook data is not uploaded. The only network call is a one-time GitHub Releases update check when the app opens or refreshes.

## Release updates

The update modal checks:

```text
https://api.github.com/repos/noahfgarrett/Risk-Analysis/releases/latest
```

Attach a downloadable `.html` asset to each release for in-app updates.

Build those assets with:

```bash
npm run build:standalone
```
