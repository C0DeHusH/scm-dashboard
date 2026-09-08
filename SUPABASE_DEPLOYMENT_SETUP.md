# SCM Executive Control Tower — Persistent Online Deployment

## What changed

This version keeps the latest uploaded SCM Excel workbook in **Supabase Storage**.
The workbook is automatically loaded when a new Streamlit session starts, so
your data can survive browser refreshes, app restarts, and redeployments.

A local `persistent_scm_data.xlsx` cache is still used as a fallback for local
development, but it is not treated as durable cloud storage.

## 1. Create a Supabase project

1. Sign in to Supabase and create/select a project.
2. Open **Storage**.
3. Create a bucket named:

   `scm-dashboard`

4. A private bucket is recommended because the Streamlit backend will use a
   server-side secret key.

## 2. Get the server-side Supabase key

In Supabase, open **Settings -> API Keys** and use a **Secret key**
(`sb_secret_...`) for the Streamlit backend.

Do not put the secret key in Python source code, GitHub, screenshots, chat,
browser JavaScript, or a public repository.

Legacy `service_role` JWT keys are also supported by the app for compatibility,
but the current Supabase secret-key format is preferred.

## 3. Configure Streamlit secrets

For local development, create:

`.streamlit/secrets.toml`

Use:

```toml
SUPABASE_URL = "https://YOUR_PROJECT_REF.supabase.co"
SUPABASE_SECRET_KEY = "sb_secret_YOUR_SECRET_KEY"
SUPABASE_BUCKET = "scm-dashboard"
SUPABASE_OBJECT = "persistent_scm_data.xlsx"
```

For Streamlit Community Cloud:

1. Open your deployed app's **Settings**.
2. Open the **Secrets** section.
3. Paste the TOML values above.
4. Save/reboot the app if required.

Never commit the real `.streamlit/secrets.toml` file to GitHub.

## 4. Repository files

Your repository should contain at minimum:

```text
scm_executive_control_tower_professional_v4.py
requirements.txt
```

Set the Python file as the Streamlit entrypoint.

## 5. How persistence works

### On initial app load
1. The app checks Supabase Storage.
2. If `persistent_scm_data.xlsx` exists, it downloads it.
3. The dashboard loads that workbook automatically.
4. A local copy is refreshed as a fallback.

### When an administrator uploads a new Excel file
1. The workbook is validated first.
2. If valid, it replaces the local cache.
3. If Supabase is configured, it upserts the workbook to the cloud object.
4. The active dashboard data changes immediately.
5. The same uploaded file is not repeatedly saved when filters trigger reruns.

### If Supabase is not configured
The dashboard still works using the local cache, but hosted platforms may delete
or replace local runtime files during restart or redeployment.

## 6. Rounding rule

The dashboard now uses standard ROUND_HALF_UP behavior:

- 8.0 -> 8
- 8.4 -> 8
- 8.49 -> 8
- 8.5 -> 9
- 8.6 -> 9
- 9.5 -> 10

This applies to displayed whole-number KPI values, percentages, Days of
Inventory, Inventory, Suggested Transfer, area averages, and graph labels.

## Recommended production control

Because the upload replaces the dashboard's persistent source workbook, restrict
upload access to trusted/admin users before making a public deployment.
