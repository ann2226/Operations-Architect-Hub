"""Copy Levrly's tracking spreadsheets from Google Drive into this repo as CSV.

Each source in SOURCES is written to data/<folder>/<tab-name>.csv (one file per
tab), plus a README.md index for that folder. Run daily by
.github/workflows/sync-sheets.yml; can also be run by hand.

Two kinds of source are supported:
  - "sheet": a native Google Sheet, read with the Sheets API
  - "xlsx":  an Excel file stored in Drive, downloaded with the Drive API and
             read with openpyxl (formula cells use the values last saved in the file)

Auth: a Google service account whose JSON key is in the GOOGLE_SERVICE_ACCOUNT_JSON
environment variable. Every source file must be shared (Viewer) with that account.
"""

import csv
import io
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SOURCES = [
    {
        "name": "Levrly Monthly Client and VA Consult",
        "kind": "sheet",
        "id": "11KlOJOYk1RtAueViHBb0g_L4Wx56p9LLBG_3tRNOuPA",
        "url": "https://docs.google.com/spreadsheets/d/11KlOJOYk1RtAueViHBb0g_L4Wx56p9LLBG_3tRNOuPA/edit",
        "folder": "client-va-consult",
    },
    {
        "name": "Levrly Escalation Tracker",
        "kind": "xlsx",
        "id": "1t2rKmmxya7lKrGfd2sLYZsXH7n5TINuC",
        "url": "https://docs.google.com/spreadsheets/d/1t2rKmmxya7lKrGfd2sLYZsXH7n5TINuC/edit",
        "folder": "escalation-tracker",
    },
]

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-") or "sheet"


def clean(rows: list[list]) -> list[list[str]]:
    """Stringify cells, drop trailing empty rows/columns, pad rows to one width."""
    rows = [["" if c is None else str(c) for c in r] for r in rows]
    rows = [r[: max((i + 1 for i, c in enumerate(r) if c.strip()), default=0)] for r in rows]
    while rows and not rows[-1]:
        rows.pop()
    width = max((len(r) for r in rows), default=0)
    return [r + [""] * (width - len(r)) for r in rows]


def read_sheet(sheets, file_id: str) -> list[tuple[str, list[list[str]]]]:
    meta = sheets.spreadsheets().get(spreadsheetId=file_id).execute()
    tabs = [s["properties"]["title"] for s in meta["sheets"]]
    res = (
        sheets.spreadsheets()
        .values()
        .batchGet(
            spreadsheetId=file_id,
            ranges=[f"'{t}'" for t in tabs],
            valueRenderOption="FORMATTED_VALUE",
        )
        .execute()
    )
    return [(t, clean(vr.get("values", []))) for t, vr in zip(tabs, res["valueRanges"])]


def read_xlsx(drive, file_id: str) -> list[tuple[str, list[list[str]]]]:
    from openpyxl import load_workbook

    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, drive.files().get_media(fileId=file_id))
    done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    wb = load_workbook(buf, data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        rows = []
        for r in ws.iter_rows(values_only=True):
            rows.append(
                [c.strftime("%Y-%m-%d") if hasattr(c, "strftime") else c for c in r]
            )
        out.append((ws.title, clean(rows)))
    return out


def write_source(src: dict, tabs: list[tuple[str, list[list[str]]]], now: str) -> None:
    out_dir = DATA_DIR / src["folder"]
    out_dir.mkdir(parents=True, exist_ok=True)
    written, index = set(), []
    for tab, rows in tabs:
        fname = f"{slug(tab)}.csv"
        with open(out_dir / fname, "w", newline="", encoding="utf-8") as f:
            csv.writer(f, lineterminator="\n").writerows(rows)
        written.add(fname)
        index.append(f"| [{tab}]({fname}) | {max(len(rows) - 1, 0)} |")
    for old in out_dir.glob("*.csv"):
        if old.name not in written:
            old.unlink()
    (out_dir / "README.md").write_text(
        f"# {src['name']}\n\n"
        f"Daily copy of the [source spreadsheet]({src['url']}). The spreadsheet is "
        "the source of truth: edit it there, not here. These files are overwritten "
        "by the `Sync spreadsheets` GitHub Action every day.\n\n"
        "Contains client and team names. Keep this repository private.\n\n"
        "| Tab | Rows (after header) |\n| :---- | ----: |\n"
        + "\n".join(index)
        + f"\n\nLast pulled: {now}\n",
        encoding="utf-8",
    )
    print(f"{src['name']}: {len(written)} tabs -> {out_dir}")


def main() -> int:
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        print("GOOGLE_SERVICE_ACCOUNT_JSON is not set", file=sys.stderr)
        return 1
    creds = service_account.Credentials.from_service_account_info(
        json.loads(raw), scopes=SCOPES
    )
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    failed = []
    for src in SOURCES:
        try:
            tabs = read_sheet(sheets, src["id"]) if src["kind"] == "sheet" else read_xlsx(drive, src["id"])
            write_source(src, tabs, now)
        except Exception as e:  # keep syncing the other sources
            print(f"::error::{src['name']}: {e}", file=sys.stderr)
            failed.append(src["name"])
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
