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
# Service account — loaded from GOOGLE_SERVICE_ACCOUNT_JSON GitHub Secret
# ---------------------------------------------------------------------------
_ENV_SA = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
if not _ENV_SA:
    print("ERROR: GOOGLE_SERVICE_ACCOUNT_JSON environment variable is not set.")
    sys.exit(1)
SERVICE_ACCOUNT_INFO = json.loads(_ENV_SA)

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
