# Mist Org Audit Log

Exports a Juniper Mist organisation's **full** audit trail to Excel — not just
the last 90 days the Mist portal shows.

## How it works

1. Reads `org_id`, `api_token` and `cloud` from `mist_org_audit_log.ini`.
2. Connects to the Mist cloud set in `cloud` (required) — use the API host
   matching the portal you log in to, e.g. `api.eu.mist.com` for
   `manage.eu.mist.com`. Short forms like `eu` / `gc3` and the portal URL
   itself also work.
3. Looks up the **first** and **last** audit records and shows them with the
   total record count.
4. Asks whether to export that full range. Answer `n` to enter your own start
   and/or end date (dd/mm/yyyy); press Enter to keep either boundary.
5. Pages through the logs (cursor-based) and writes an `.xlsx` to your home
   directory, with site names resolved.

## Setup

```bash
pip install -r requirements.txt
cp mist_org_audit_log.ini.example mist_org_audit_log.ini
# edit org_id, api_token and cloud
python3 mist_org_audit_log.py
```

## Running Behind a TLS-Inspecting Proxy (e.g. Zscaler)

If your network intercepts and re-signs HTTPS traffic, TLS verification against
the Mist API will fail with a certificate error. Fill in the `[network]`
section of `mist_org_audit_log.ini`:

```ini
[network]
ca_bundle = /path/to/zscaler-root-ca.pem
verify_ssl = true
http_proxy =
https_proxy =
```

- `ca_bundle` — PEM file with your proxy's root CA certificate (or a bundle
  that includes it). Your IT team or the Zscaler client app can usually export it.
- `verify_ssl = false` disables certificate verification entirely. Only use it
  if you truly can't get the proxy's CA certificate.
- `http_proxy` / `https_proxy` — only needed if the standard `HTTP_PROXY` /
  `HTTPS_PROXY` / `NO_PROXY` environment variables aren't already set.

With no `[network]` section (or blank values), the tool behaves exactly as on a
normal, non-intercepted connection.

## Output columns

Timestamp (UTC), Timestamp (Epoch), Admin Name, Admin ID, Site Name, Site ID,
For Site, Message, Source IP, Before, After, Log ID — newest first.
