# Setup automatic refresh and Netlify deploy

This repository intentionally does not store operational data or secrets.

The tracked `tmp/index.html` is a safe template with an empty payload. GitHub Actions refreshes it at runtime and deploys the generated file to Netlify.

## 1. Create or choose the Netlify site

Use a Netlify site that will be deployed by GitHub Actions.

Current site:

```text
Site URL: https://oms-location-dashboard-a-capone.netlify.app
Admin URL: https://app.netlify.com/projects/oms-location-dashboard-a-capone
Site ID: 89bfda26-6337-4034-8e57-a6575f466c41
```

Recommended setup:

- Do not rely on Netlify build from Git pushes.
- GitHub Actions is the deployer.
- Netlify is only the hosting target.

If the site is connected to the GitHub repository in Netlify, disable Netlify auto builds or make sure they do not overwrite the GitHub Actions production deploy with the empty template.

## 2. Create a Netlify personal access token

In Netlify:

1. User settings.
2. Applications.
3. Personal access tokens.
4. New access token.
5. Copy the token once.

GitHub secret name:

```text
NETLIFY_AUTH_TOKEN
```

## 3. Find the Netlify site id

In Netlify:

1. Open the site.
2. Site configuration.
3. General.
4. Site details.
5. Copy the Site ID.

GitHub secret name:

```text
NETLIFY_SITE_ID
```

## 4. Add GitHub Actions secrets

In GitHub:

1. Open `a-capone/oms-location-dashboard`.
2. Settings.
3. Secrets and variables.
4. Actions.
5. New repository secret.

Required secrets:

```text
GCP_SERVICE_ACCOUNT_JSON
GEODIS_SFTP_PASSWORD
INVAPP_CLIENT_ID
INVAPP_CLIENT_SECRET
NETLIFY_AUTH_TOKEN
NETLIFY_SITE_ID
```

`GCP_SERVICE_ACCOUNT_JSON` must contain the full JSON service account content, not a local file path.

## 5. Run first deploy

In GitHub:

1. Actions.
2. Refresh OMS dashboard.
3. Run workflow.
4. Select branch `main`.
5. Run workflow.

The workflow will:

1. Query BigQuery/Kibo.
2. Download the latest Geodis `Uscite_*.xlsx`.
3. Call the Inventory API for `TLTEMP0048`, `TLITWH0048`, `TLITGX0048`.
4. Generate `tmp/index.html`.
5. Deploy `tmp/` to Netlify production.

## 6. Schedule

The workflow already runs every 2 hours:

```yaml
schedule:
  - cron: "0 */2 * * *"
```

This is UTC time.
