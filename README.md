# OMS Location Dashboard

Static GitHub Pages dashboard for OMS/Kibo shipments, Geodis outbound status and WMS inventory for:

- `TLTEMP0048`
- `TLITWH0048`
- `TLITGX0048`

The published site is the single file in `tmp/index.html`.

The Shipments view exposes dedicated tabs for `TLTEMP0048`, `TLITWH0048`, `TLITGX0048`, plus an all-locations view.
Each tab can be downloaded as CSV from the table header.

Default Pages URL:

```text
https://a-capone.github.io/oms-location-dashboard/
```

## Refresh

The GitHub Actions workflow refreshes the dashboard every 2 hours and can also be run manually from the Actions tab.

Workflow:

1. Reads OMS/Kibo data from BigQuery view `tlg-business-intelligence-prd.til.v_oms_active_location_dashboard`.
2. Reads the latest Geodis `Uscite_*.xlsx` from SFTP.
3. Calls `InventoryWmsMonitoring` for the three dashboard locations.
4. Downloads the previous deployed `index.html` from GitHub Pages when available, so the inventory tab can preserve GXO history.
5. Writes the refreshed mono-file to `tmp/index.html`.
6. Deploys `tmp/` to GitHub Pages.

## Inventory Reading

The inventory tab is intentionally not a total stock view. It tracks the temporary migration from TEMP to GXO:

- `TLTEMP0048`: frozen TEMP baseline.
- `TLITGX0048`: quantity entering GXO, the primary value to monitor.
- `TLITWH0048`: supporting warehouse quantity.

The dashboard shows GXO current quantity, TEMP frozen quantity, GXO delta from the previous refresh, GXO delta from the first retained history point, and a compact GXO trend chart.

## Required GitHub Secrets

Configure these repository secrets before relying on the scheduled workflow:

- `GCP_SERVICE_ACCOUNT_JSON`: JSON service account with BigQuery read access.
- `GEODIS_SFTP_PASSWORD`: Geodis SFTP password.
- `INVAPP_CLIENT_ID`: OAuth client id for the inventory API.
- `INVAPP_CLIENT_SECRET`: OAuth client secret for the inventory API.

Optional repository variable:

- `DASHBOARD_PUBLIC_URL`: deployed dashboard base URL, only needed if the site uses a custom domain or a non-default Pages URL.

## Local Build

```powershell
$env:GOOGLE_APPLICATION_CREDENTIALS=(Resolve-Path '.\path-to-service-account.json').Path
$env:GEODIS_SFTP_PASSWORD='...'
$env:INVAPP_CLIENT_ID='...'
$env:INVAPP_CLIENT_SECRET='...'
python scripts\build_oms_location_dashboard.py --output tmp\index.html
```

Open `tmp/index.html` in a browser after the build.
