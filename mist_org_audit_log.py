#!/usr/bin/env python3
"""Export the full Mist org audit trail (beyond the GUI's 90-day limit) to Excel.

Only org_id and api_token are read from mist_org_audit_log.ini. The Mist cloud
is auto-detected, the earliest and latest audit records are looked up, and the
user can accept that full range or enter their own start/end dates.
"""
import configparser
import json
import os
import sys
import time
from datetime import datetime, timezone
from urllib.parse import urlparse

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
config = configparser.ConfigParser()
config.read(os.path.join(SCRIPT_DIR, "mist_org_audit_log.ini"))

ORG_ID = config.get("mist", "org_id")
API_TOKEN = config.get("mist", "api_token")
OUTPUT_DIR = os.path.expanduser("~")

# Mist clouds to probe for the org - first one that returns the org wins.
MIST_CLOUDS = [
    "https://api.mist.com",
    "https://api.eu.mist.com",
    "https://api.gc1.mist.com",
    "https://api.gc2.mist.com",
    "https://api.gc3.mist.com",
    "https://api.gc4.mist.com",
    "https://api.gc5.mist.com",
    "https://api.gc6.mist.com",
    "https://api.gc7.mist.com",
    "https://api.ac2.mist.com",
    "https://api.ac5.mist.com",
    "https://api.ac6.mist.com",
]

# Lower bound for "all history" searches. start=0 is silently ignored by the
# logs API (it falls back to a short default window), so use a real date that
# predates any Mist org.
HISTORY_START = int(datetime(2015, 1, 1, tzinfo=timezone.utc).timestamp())

REQUEST_TIMEOUT = 60
API_BASE = None
API_HOST = None


def build_session():
    """Shared HTTP session, wired up for a TLS-inspecting proxy (e.g. Zscaler)
    when the optional [network] section is present in the .ini. With no
    [network] section, behaviour is unchanged (certifi CA store, and the
    HTTP_PROXY / HTTPS_PROXY / NO_PROXY environment variables if set)."""
    session = requests.Session()
    session.headers.update({"Authorization": f"Token {API_TOKEN}"})
    if "network" not in config:
        return session

    section = config["network"]
    ca_bundle = section.get("ca_bundle", "").strip()
    if ca_bundle:
        ca_path = os.path.expanduser(ca_bundle)
        if not os.path.isfile(ca_path):
            sys.exit(f"[network] ca_bundle does not exist: {ca_path}\n"
                     f"This should be a PEM file containing your proxy's root CA "
                     f"certificate (e.g. exported from Zscaler).")
        session.verify = ca_path
    elif not section.getboolean("verify_ssl", fallback=True):
        session.verify = False
        print("WARNING: TLS certificate verification is DISABLED ([network] verify_ssl = false). "
              "Only use this as a last resort.")
        requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)

    for scheme in ("http", "https"):
        proxy = section.get(f"{scheme}_proxy", "").strip()
        if proxy:
            session.proxies[scheme] = proxy
    return session


SESSION = build_session()


def http_get(url, timeout=REQUEST_TIMEOUT):
    """GET with friendly errors for the usual proxy / Zscaler failure modes."""
    try:
        return SESSION.get(url, timeout=timeout)
    except requests.exceptions.SSLError as e:
        sys.exit(f"\nTLS certificate verification failed for {url}: {e}\n"
                 f"If you're behind a TLS-inspecting proxy (e.g. Zscaler), set [network] ca_bundle "
                 f"in mist_org_audit_log.ini to the path of its root CA certificate (PEM format).")
    except requests.exceptions.ProxyError as e:
        sys.exit(f"\nCould not reach the proxy for {url}: {e}\n"
                 f"Check [network] http_proxy / https_proxy in mist_org_audit_log.ini "
                 f"or your HTTP(S)_PROXY environment variables.")


def parse_ddmmyyyy(date_str, end_of_day=False):
    dt = datetime.strptime(date_str, "%d/%m/%Y")
    if end_of_day:
        dt = dt.replace(hour=23, minute=59, second=59)
    return int(dt.timestamp())


def fmt_ts(ts):
    return datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")


def fmt_ddmmyyyy(ts):
    return datetime.fromtimestamp(ts).strftime("%d/%m/%Y")


def format_eta(seconds):
    if seconds < 60:
        return f"{int(seconds)}s"
    m, s = divmod(int(seconds), 60)
    return f"{m}m {s}s"


def request_with_retry(url, max_retries=5):
    for attempt in range(1, max_retries + 1):
        resp = http_get(url)
        if resp.status_code == 429:
            retry_after = int(resp.headers.get("Retry-After", 30))
            print(f"\n  Rate limited - waiting {retry_after}s (attempt {attempt}/{max_retries})...")
            time.sleep(retry_after)
            continue
        resp.raise_for_status()
        return resp
    resp.raise_for_status()
    return resp


def detect_cloud():
    """Return (api_host, org_info) for the first Mist cloud that knows this org."""
    for host in MIST_CLOUDS:
        try:
            resp = http_get(f"{host}/api/v1/orgs/{ORG_ID}", timeout=15)
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout):
            continue
        if resp.status_code == 200:
            return host, resp.json()
    return None, None


def fetch_boundary_record(sort):
    """Fetch the single oldest (sort=timestamp) or newest (sort=-timestamp) audit record."""
    now = int(time.time())
    url = f"{API_BASE}/orgs/{ORG_ID}/logs?limit=1&start={HISTORY_START}&end={now}&sort={sort}"
    data = request_with_retry(url).json()
    results = data.get("results", [])
    return (results[0] if results else None), data.get("total")


def prompt_date(label, default_str):
    while True:
        value = input(f"  {label} date (dd/mm/yyyy) [Enter = {default_str}]: ").strip()
        if not value:
            return None
        try:
            datetime.strptime(value, "%d/%m/%Y")
            return value
        except ValueError:
            print("    Invalid date - use dd/mm/yyyy, e.g. 01/02/2024")


def choose_range(first_ts, last_ts):
    """Let the user accept the full first->last range or enter their own dates.

    Returns (start_epoch, end_epoch, start_label, end_label).
    """
    first_str = fmt_ddmmyyyy(first_ts)
    last_str = fmt_ddmmyyyy(last_ts)
    while True:
        accept = input(f"\n  Export all records from {first_str} to {last_str}? (y/n): ").strip().lower()
        if accept == "y":
            # Use exact record timestamps so the first and last entries are both included.
            return first_ts, last_ts + 1, first_str, last_str
        if accept != "n":
            continue

        print("\n  Enter a new range (press Enter to keep the first/last record date).")
        start_in = prompt_date("Start", first_str)
        end_in = prompt_date("End", last_str)
        start = parse_ddmmyyyy(start_in) if start_in else first_ts
        end = parse_ddmmyyyy(end_in, end_of_day=True) if end_in else last_ts + 1
        if start >= end:
            print("    Start date must be before end date - try again.")
            continue
        return start, end, start_in or first_str, end_in or last_str


def fetch_site_names():
    site_map = {}
    page = 1
    api_calls = 0
    while True:
        url = f"{API_BASE}/orgs/{ORG_ID}/sites?limit=1000&page={page}"
        resp = request_with_retry(url)
        api_calls += 1
        data = resp.json()
        if not data:
            break
        for site in data:
            site_map[site["id"]] = site.get("name", site["id"])
        if len(data) < 1000:
            break
        page += 1
    return site_map, api_calls


def fetch_audit_logs(start, end):
    all_logs = []
    api_calls = 0
    page_times = []
    total = None
    url = f"{API_BASE}/orgs/{ORG_ID}/logs?limit=1000&start={start}&end={end}&sort=-timestamp"
    page = 0
    empty_first_page_retried = False
    while url:
        page += 1
        t0 = time.time()
        resp = request_with_retry(url)
        page_time = time.time() - t0
        page_times.append(page_time)
        api_calls += 1
        data = resp.json()
        if total is None:
            total = data.get("total", 0)
        results = data.get("results", [])
        if not results:
            # Mist's audit-log search occasionally returns an empty first
            # page for a large historical window even when data exists -
            # confirmed by re-running the identical request moments later
            # and getting real results. Give it one retry before giving up.
            if page == 1 and not empty_first_page_retried:
                empty_first_page_retried = True
                page -= 1
                page_times.pop()
                total = None
                print("\r  First page came back empty - retrying once in case it was a transient hiccup...")
                time.sleep(5)
                continue
            break
        all_logs.extend(results)
        avg_time = sum(page_times) / len(page_times)
        pct = min(99, len(all_logs) / max(total, 1) * 100) if total else 0
        remaining = max(0, total - len(all_logs)) / 1000 * avg_time
        sys.stdout.write(f"\r  Page {page} | {len(all_logs)}/{total} entries ({pct:.0f}%) | {avg_time:.1f}s/page | ETA: {format_eta(remaining)}   ")
        sys.stdout.flush()
        next_path = data.get("next")
        if next_path:
            url = f"{API_HOST}{next_path}"
        else:
            break
    print(f"\r  Audit log fetch complete - {len(all_logs)} entries in {format_eta(sum(page_times))}                    ")
    return all_logs, api_calls


def style_header(ws, headers, fill_color):
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"


def auto_width(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 3, 60)


def main():
    global API_BASE, API_HOST

    if API_TOKEN in ("", "YOUR_API_TOKEN_HERE") or ORG_ID in ("", "YOUR_ORG_ID_HERE"):
        print("Set org_id and api_token in mist_org_audit_log.ini before running.")
        return

    print("Locating organisation on Mist cloud...")
    host, org_info = detect_cloud()
    if not host:
        print("  Org not found on any Mist cloud - check org_id and api_token.")
        return
    API_HOST = host
    API_BASE = f"{host}/api/v1"
    org_name = org_info.get("name", "Unknown")

    print("Looking up first and last audit records...")
    first, total_all = fetch_boundary_record("timestamp")
    last, _ = fetch_boundary_record("-timestamp")
    if not first or not last:
        print("  No audit log entries found for this organisation.")
        return
    first_ts, last_ts = first["timestamp"], last["timestamp"]

    print(f"\n{'='*55}")
    print(f"  Mist Org Audit Log")
    print(f"{'='*55}")
    print(f"  Organisation:  {org_name}")
    print(f"  Org ID:        {ORG_ID}")
    print(f"  Cloud:         {urlparse(API_HOST).netloc}")
    print(f"  Total records: {total_all}")
    print(f"  First record:  {fmt_ts(first_ts)}  ({first.get('admin_name', '')}: {first.get('message', '')})")
    print(f"  Last record:   {fmt_ts(last_ts)}  ({last.get('admin_name', '')}: {last.get('message', '')})")
    print(f"{'='*55}")

    start, end, start_label, end_label = choose_range(first_ts, last_ts)

    print(f"\n  Date range: {start_label} - {end_label}")
    confirm = input("  Proceed? (y/n): ").strip().lower()
    if confirm != "y":
        print("  Aborted.")
        return

    start_time = time.time()

    print("\nFetching audit log entries from Mist API...")
    logs, log_api_calls = fetch_audit_logs(start, end)
    print(f"Total entries fetched: {len(logs)}")

    site_ids = set(l.get("site_id") for l in logs if l.get("site_id"))
    site_map = {}
    site_api_calls = 0
    if site_ids:
        print(f"Fetching names for {len(site_ids)} sites...")
        site_map, site_api_calls = fetch_site_names()

    print("Building Excel spreadsheet...")

    headers = [
        "Timestamp (UTC)", "Timestamp (Epoch)", "Admin Name", "Admin ID",
        "Site Name", "Site ID", "For Site", "Message", "Source IP",
        "Before", "After", "Log ID",
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Log"
    style_header(ws, headers, "37474F")

    for r, log in enumerate(logs, 2):
        ts = log.get("timestamp")
        site_id = log.get("site_id", "") or ""
        before = log.get("before")
        after = log.get("after")
        row = [
            fmt_ts(ts) if ts else "",
            ts,
            log.get("admin_name", ""),
            log.get("admin_id", ""),
            site_map.get(site_id, ""),
            site_id,
            log.get("for_site", ""),
            log.get("message", ""),
            log.get("src_ip", ""),
            json.dumps(before) if before else "",
            json.dumps(after) if after else "",
            log.get("id", ""),
        ]
        for c, value in enumerate(row, 1):
            ws.cell(row=r, column=c, value=value)

    auto_width(ws)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    safe_org_name = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in org_name).strip().replace(" ", "_")
    safe_start = start_label.replace("/", "-")
    safe_end = end_label.replace("/", "-")
    filename = f"Mist_Org_Audit_Log_{safe_org_name}_{safe_start}_to_{safe_end}_{timestamp}.xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    wb.save(filepath)

    total_api = log_api_calls + site_api_calls
    elapsed = time.time() - start_time
    print(f"\n{'='*55}")
    print(f"  Mist Org Audit Log Summary - {org_name}")
    print(f"{'='*55}")
    print(f"  Entries fetched:      {len(logs)}")
    print(f"  Unique sites seen:    {len(site_ids)}")
    print(f"  API calls - logs:     {log_api_calls}")
    print(f"  API calls - sites:    {site_api_calls}")
    print(f"  Total API calls:      {total_api}")
    print(f"  Total elapsed time:   {format_eta(elapsed)}")
    print(f"{'='*55}")
    print(f"  Report saved: {filepath}")


if __name__ == "__main__":
    main()
