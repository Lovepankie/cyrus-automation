"""
Download and merge the audit Google Sheets for all regions.
Writes: merged_output_audit/DDMMYYYY_HHMMSS/audit_merged_DDMMYYYY_HHMMSS.xlsx

Sheets must be shared with:
    download-sheets@drive-399813.iam.gserviceaccount.com
"""

import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

# ---------------------------------------------------------------------------
# Service account — reads from env var in GitHub Actions, falls back to embedded
# ---------------------------------------------------------------------------
_ENV_SA = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
SERVICE_ACCOUNT_INFO = json.loads(_ENV_SA) if _ENV_SA else {
    "type": "service_account",
    "project_id": "drive-399813",
    "private_key_id": "be807337b7d6b3391199a4421ffbb1739bf7430f",
    "private_key": (
        "-----BEGIN PRIVATE KEY-----\n"
        "MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDhyt1lrbccnrqE\n"
        "Sb+o/6do/I1p/MvrdjP/ZZBgTgn7765uE13rWrKc/bkUBmEwyCTVXGOuB4bTCmsx\n"
        "z/NuU+Qea/0cT339zGCNj2c9Fam3HIneL9e/vSC0yj8KAjFNBbzzRr1BJL8S+pML\n"
        "lMouH9QqJWsW5+ldakSV/g8bxk52INitXmqTizC9aonrMfXIK0C9+uYk0HFesWX1\n"
        "JxfG2PeZ8s4yUpxDFbDGtUjRI5Xpv2GXo9ywjRzm52EuRQ6PVmAcbyEeybG9Zhix\n"
        "+z0K5hAVSnvJl7WTY4xCLNYEmk32E8pBFceUwITWfsyLHiNtMxjLAJqyFZCP3a31\n"
        "mMyPbgJRAgMBAAECggEACIGhB6ScwmpEcplorBA9mVnyZezNLei9GsETyF0ISUKF\n"
        "WsZGAojfM8TnRRbccH8JzDP32WWVMbwxv6Wq8Rwd+vtshLWF9JrPhGXDqOx5AVoE\n"
        "46b/xMx2P/limJjujI1Lygp/NMYSoL9p7MTFPSmbz56rQyinhQps3Q2+OfilmhOe\n"
        "lW744GYZz4DgFqEwEWQN/JYU4K4Rp2DMZ1Vxou5RhRMFKDg/QVU9yfdWwfAIb56x\n"
        "Nsj+fYI8bXCAC+wGLAgSSBndQzAV96MyMx//PhnWGOWu1vkcEDb6rfhYVwMKB7/n\n"
        "FDhkr9j3pGles2Sr2JR3tQx5pllFVcui/P6vOOeshQKBgQD5fPPzV1Pj/178+QVB\n"
        "4VV4sf5eZPSXijjRMaWsfDsr3glJkGhtrbXYv4U8IBCzFwght/TuYZ9hk1EBJV9/\n"
        "mp/pQ5L9B74+X5hpFK/9ont6+snhBqmbX/EF8+DBn/XJVn3yxlSGutQHcpCwuBZq\n"
        "hRwKcCNhTnsW8i+v7gkmVYd9YwKBgQDnr5SiOe/xOpqt3Z3AYz7GoeitnpTKpT9D\n"
        "+pQV61V67BHueICUp109NCyNi1qZCO4d0LsjIslyx7kPOEYUf1C3Ul1LsUr+cAkC\n"
        "1zvdNRVvmNB4RQ+WzCwHgOSsQXQyNxO74uj5fZNffVT0CJjbgvbHsQIQWQkj9vS6\n"
        "dvQZJwpZuwKBgQC7jQ6LCUQcataKj90+6FlrkUsqxPQGk7cgtBTatM16rcEHl1KZ\n"
        "2POSTG+pgmVrbE4FoxeyuJqrLKbBmMnQ0HmUTuNJ6i8/DngxzoZ6wlHXYn6u9mY8\n"
        "UoSLOAnnJQwNXkLyZdwXKI3KR3q0Dr9zNjudMS23sdrgphHpTKV+Nt/TBwKBgHSG\n"
        "6NjXDwljEl0UM72JrQ82a9K3CdsKVdGY/FYx8OJMTZCFZxZdPxYDYc0nI8AIr8qr\n"
        "KxQ28N8b+MXg5c51YmFxuZ7SYwepzb5yBpfxlQB1+ZQkF/0eX56+g0Tn/ssqzHAZ\n"
        "ZlflgvPqE4pRsJ/nNLunGYSjY2eFU/1cytTDv/71AoGACmCAabda8YQzHwovE9ar\n"
        "4K5GT6m7L6T4357FtnacoE17j9YcMJrZNeTaddYshIL/8HzpFJtZQXxmTk/XfVpm\n"
        "DyxCj1BUjH5H4MjnDc6efS5Jz5CkljEiWPupTAMRyOg+LV085j+sg/zvZSewfAdw\n"
        "o80f1uB4a2fQb77Kw4YLc+E=\n"
        "-----END PRIVATE KEY-----\n"
    ),
    "client_email": "download-sheets@drive-399813.iam.gserviceaccount.com",
    "client_id": "110748723405822534910",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/download-sheets%40drive-399813.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com",
}

# ---------------------------------------------------------------------------
# Region sheets
# ---------------------------------------------------------------------------
REGIONS = [
    {"name": "World Bank Audit App - CSTW",    "tab": "Quality and Material Checklist", "url": "https://docs.google.com/spreadsheets/d/1FMKX30Jv0CEONxfeqm8vplYrdrin7GzdcgquPTdKyJs/edit?gid=85391464#gid=85391464"},
    {"name": "World Bank Audit App - CSMETRO", "tab": "Quality and Material Checklist", "url": "https://docs.google.com/spreadsheets/d/1zuMYMGY4cUcEYLxup335mGAytqV6Bsu2jxU22dICUiw/edit?gid=85391464#gid=85391464"},
    {"name": "World Bank Audit App - Western", "tab": "Quality and Material Checklist", "url": "https://docs.google.com/spreadsheets/d/11BCoyxmjy_Ukuo4TteY7B8sZLxHq8sWNbXfrMlhU_is/edit?gid=85391464#gid=85391464"},
    {"name": "World Bank Audit App - NORTHE",  "tab": "Quality and Material Checklist", "url": "https://docs.google.com/spreadsheets/d/1I-ALJDEDjPzUaPmAVqOOirzMcS2Vjc4sI-_L3UG6FLw/edit?gid=85391464#gid=85391464"},
    {"name": "World Bank Audit App - CSTSOUT", "tab": "Quality and Material Checklist", "url": "https://docs.google.com/spreadsheets/d/1F8nB8n72kHa9dGRHZWQo9iDajOa6IqPA967w-ctkj1o/edit?gid=85391464#gid=85391464"},
    {"name": "World Bank Audit App - CSTN",    "tab": "Quality and Material Checklist", "url": "https://docs.google.com/spreadsheets/d/1C-foadJzRhiNSHNLrb8ByAvp8ujrSW6WbryZLbLnbBY/edit?gid=85391464#gid=85391464"},
    {"name": "EASP Audit App - Western 2",     "tab": "Quality and Material Checklist", "url": "https://docs.google.com/spreadsheets/d/1GliAdOV-G0MnOnonDBDE23DpimXPqtUxl8zMghWjMhc/edit?gid=236169048#gid=236169048"},
    {"name": "World Bank Audit App Extra",     "tab": "Quality and Material Checklist", "url": "https://docs.google.com/spreadsheets/d/1_PKQP8uiTySYqGKv2tzObi2BBzkqQXmxYwnWbywWg5Y/edit?gid=85391464#gid=85391464"},
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

HERE       = Path(__file__).parent
OUTPUT_DIR = HERE / "merged_output_audit"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_client():
    creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=SCOPES)
    return gspread.authorize(creds)


def parse_sheet_id_and_gid(url: str):
    m = re.search(r"/spreadsheets/d/([^/?\s]+)", url)
    if not m:
        raise ValueError(f"Cannot extract sheet ID from: {url}")
    sheet_id = m.group(1)
    gid_m = re.search(r"gid=(\d+)", url)
    gid = int(gid_m.group(1)) if gid_m else 0
    return sheet_id, gid


def download_tab(gc, region: str, url: str) -> pd.DataFrame:
    sheet_id, gid = parse_sheet_id_and_gid(url)
    spreadsheet = gc.open_by_key(sheet_id)
    worksheet = next((ws for ws in spreadsheet.worksheets() if ws.id == gid), None)
    if worksheet is None:
        raise ValueError(f"Tab with gid={gid} not found in spreadsheet {sheet_id}")

    try:
        records = worksheet.get_all_records(numericise_ignore=["all"])
        if not records:
            return pd.DataFrame({"Region": pd.Series([], dtype=str)})
        df = pd.DataFrame(records)
    except gspread.exceptions.APIError:
        values = worksheet.get_all_values()
        if len(values) < 2:
            return pd.DataFrame({"Region": pd.Series([], dtype=str)})
        df = pd.DataFrame(values[1:], columns=values[0])

    df["Region"] = region
    df = df[["Region"] + [c for c in df.columns if c != "Region"]]
    return df


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def save_individual(frames_map: dict, run_dir: Path, timestamp: str):
    for name, df in frames_map.items():
        safe = re.sub(r'[\\/*?:\[\]]', "_", name)
        path = run_dir / f"{safe}_{timestamp}.xlsx"
        df.to_excel(path, index=False)
        print(f"    Saved {name} -> {path.name}")


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"Authenticating as {SERVICE_ACCOUNT_INFO['client_email']} ...")
    gc = get_client()
    print(f"OK — downloading {len(REGIONS)} sheets\n")

    frames_map, errors = {}, []

    for row in REGIONS:
        name, tab, url = row["name"], row["tab"], row["url"]
        print(f"  [{name}] {tab} ... ", end="", flush=True)
        try:
            df = download_tab(gc, name, url)
            frames_map[name] = df
            print(f"{len(df)} rows, {len(df.columns)} columns")
        except Exception as exc:
            print(f"FAILED - {exc}")
            errors.append((name, str(exc)))

    if not frames_map:
        print("\nNo data downloaded. Make sure each sheet is shared with:")
        print(f"  {SERVICE_ACCOUNT_INFO['client_email']}")
        sys.exit(1)

    timestamp = datetime.now().strftime("%d%m%Y_%H%M%S")
    run_dir   = OUTPUT_DIR / timestamp
    run_dir.mkdir(exist_ok=True)

    merged   = pd.concat(frames_map.values(), ignore_index=True, sort=False)
    out_path = run_dir / f"audit_merged_{timestamp}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        merged.to_excel(writer, index=False, sheet_name="All Regions")
        for name, df in frames_map.items():
            safe = re.sub(r'[\\/*?:\[\]]', "_", name)[:31]
            df.to_excel(writer, index=False, sheet_name=safe)

    print(f"\nMerged {len(merged)} total rows from {len(frames_map)} sheets.")
    print(f"Saved -> {out_path}")

    print("\nSaving individual files:")
    save_individual(frames_map, run_dir, timestamp)

    if errors:
        print(f"\nFailed ({len(errors)}):")
        for name, msg in errors:
            print(f"  {name}: {msg}")

    # Write the merged file path so the email script can find it
    (HERE / ".last_merged_file").write_text(str(out_path))


if __name__ == "__main__":
    main()
