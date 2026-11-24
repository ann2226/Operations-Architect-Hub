"""Export the "Levrly Monthly Client and VA Consult" Google Sheet into this repo.

Every tab is written to data/client-va-consult/<tab-name>.csv, and a README.md
index for that folder is regenerated. Run by .github/workflows/sync-client-va-consult.yml
once a day; can also be run by hand.

Auth: a Google service account whose JSON key is in the GOOGLE_SERVICE_ACCOUNT_JSON
environment variable. The sheet must be shared (Viewer) with that account's email.
"""

import csv
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build

SPREADSHEET_ID = os.environ.get(
    "SHEET_ID", "11KlOJOYk1RtAueViHBb0g_L4Wx56p9LLBG_3tRNOuPA"
)
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit"
OUT_DIR = Path(__file__).resolve().parent.parent / "data" / "client-va-consult"


def slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-") or "sheet"


def trim(rows: list[list[str]]) -> list[list[str]]:
    """Drop trailing empty rows and pad every row to the same width."""
    while rows and not any(c.strip() for c in rows[-1]):
        rows.pop()
    width = max((len(r) for r in rows), default=0)
    return [r + [""] * (width - len(r)) for r in rows]


def main() -> int:
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        print("GOOGLE_SERVICE_ACCOUNT_JSON is not set", file=sys.stderr)
        return 1
    creds = service_account.Credentials.from_service_account_info(
        json.loads(raw),
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)

    meta = sheets.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    title = meta["properties"]["title"]
    tabs = [s["properties"]["title"] for s in meta["sheets"]]

    result = (
        sheets.spreadsheets()
        .values()
        .batchGet(
            spreadsheetId=SPREADSHEET_ID,
            ranges=[f"'{t}'" for t in tabs],
            valueRenderOption="FORMATTED_VALUE",
        )
        .execute()
    )

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    written = set()
    index_rows = []
    for tab, vr in zip(tabs, result["valueRanges"]):
        rows = trim([[str(c) for c in r] for r in vr.get("values", [])])
        fname = f"{slug(tab)}.csv"
        with open(OUT_DIR / fname, "w", newline="", encoding="utf-8") as f:
            csv.writer(f, lineterminator="\n").writerows(rows)
        written.add(fname)
        index_rows.append(f"| [{tab}]({fname}) | {max(len(rows) - 1, 0)} |")

    # Remove CSVs for tabs that no longer exist in the sheet.
    for old in OUT_DIR.glob("*.csv"):
        if old.name not in written:
            old.unlink()

    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    (OUT_DIR / "README.md").write_text(
        f"# {title}\n\n"
        f"Daily copy of the [Google Sheet]({SHEET_URL}). The sheet is the source of "
        "truth: edit it there, not here. These files are overwritten by the "
        "`Sync Client & VA Consult sheet` GitHub Action every day.\n\n"
        f"Contains client and VA names. Keep this repository private.\n\n"
        "| Tab | Data rows |\n| :---- | ----: |\n"
        + "\n".join(index_rows)
        + f"\n\nLast pulled: {now}\n",
        encoding="utf-8",
    )
    print(f"Wrote {len(written)} tabs to {OUT_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
