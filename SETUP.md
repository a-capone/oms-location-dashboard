# Setup automatic refresh and GitHub Pages deploy

This repository intentionally does not store operational data or secrets.

The tracked `tmp/index.html` is a safe template with an empty payload. GitHub Actions refreshes it at runtime and deploys the generated file to GitHub Pages.

## 1. Enable GitHub Pages

Use GitHub Pages with GitHub Actions as the publishing source.

Default site:

```text
Site URL: https://a-capone.github.io/oms-location-dashboard/
```

In GitHub:

1. Open `a-capone/oms-location-dashboard`.
2. Go to Settings.
3. Go to Pages.
4. Set Source to `GitHub Actions`.

The repository is private. GitHub Pages for private repositories requires a GitHub plan that supports private Pages. If this is not available, make the repository public or use Cloudflare Pages with the same generated `tmp/` folder.

## 2. Add GitHub Actions secrets

In GitHub:

1. Open `a-capone/oms-location-dashboard`.
2. Go to Settings.
3. Go to Secrets and variables.
4. Go to Actions.
5. Add the repository secrets.

Required secrets:

```text
GCP_SERVICE_ACCOUNT_JSON
GEODIS_SFTP_PASSWORD
INVAPP_CLIENT_ID
INVAPP_CLIENT_SECRET
```

`GCP_SERVICE_ACCOUNT_JSON` must contain the full JSON service account content, not a local file path.

No Netlify secrets are required.

## 3. Optional dashboard URL variable

The workflow downloads the previously deployed `index.html` before each refresh so the inventory tab can preserve GXO history.

By default it reads from:

```text
https://a-capone.github.io/oms-location-dashboard/index.html
```

If the dashboard is published on a custom domain or a different URL, add this repository variable:

```text
DASHBOARD_PUBLIC_URL
```

Example value:

```text
https://dashboard.example.com
```

## 4. Run first deploy

In GitHub:

1. Go to Actions.
2. Open `Refresh OMS dashboard`.
3. Select `Run workflow`.
4. Select branch `main`.
5. Run the workflow.

The workflow will:

1. Query BigQuery/Kibo.
2. Download the latest Geodis `Uscite_*.xlsx`.
3. Call the Inventory API for `TLTEMP0048`, `TLITWH0048`, `TLITGX0048`.
4. Generate `tmp/index.html`.
5. Deploy `tmp/` to GitHub Pages.

## 5. Schedule

The workflow already runs every 2 hours:

```yaml
schedule:
  - cron: "0 */2 * * *"
```

This is UTC time.
