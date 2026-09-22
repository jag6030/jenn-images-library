"""
Jenn Gabel Source Image Library -- Cloudinary Uploader + Sheet Writer
----------------------------------------------------------------------
Scans UPLOAD_DIR for new source images, uploads each to Cloudinary using
its full filename (minus extension) as the Public ID, and writes the
resulting Public ID into the matching row of the "image-library-tracker"
Google Sheet -- matched by Source Image ID, extracted from the filename.

Rows are NEVER created by this script -- only updated. If a file's
Source Image ID has no matching row, the file is left in place (not
archived) and a warning is logged. The script never crashes on a bad or
unmatched file; it skips that file and keeps going.

FILENAME CONVENTION (required):
  {SourceImageID}-{landscape|vertical}-{descriptive-slug}.jpg

  e.g. si0001-landscape-margate-drone1.jpg
       si0001-vertical-margate-drone1.jpg

  Source Image ID is everything before the first hyphen. Orientation is
  the token right after it and must be exactly "landscape" or "vertical".
  Everything after that (the descriptive slug) is free text and may
  itself contain hyphens -- it's never parsed, just carried through as
  part of the Public ID.

  Cloudinary Public ID = the full filename with its extension stripped,
  e.g. "si0001-landscape-margate-drone1" -- this script doesn't need to
  know about the 8 named transformations already built in Cloudinary;
  it just uploads the raw image and stores whatever Public ID comes
  back.

SHEET COLUMNS TOUCHED (by header name, not position -- a column reorder
on the sheet won't misalign anything):
  landscape file -> "Landscape Public ID"
  vertical file   -> "Vertical Public ID"
  either           -> "Uploaded Date" (today's date, stamped every time
                       this script successfully writes a Public ID)

No other column (Status, Tags, Notes, etc.) is ever read or written.

CREDENTIALS: loaded from .env via python-dotenv -- CLOUDINARY_CLOUD_NAME,
CLOUDINARY_API_KEY, CLOUDINARY_API_SECRET, SERVICE_ACCOUNT_FILE, SHEET_ID,
UPLOAD_DIR, ARCHIVE_DIR. No secrets hardcoded in this file.

SETUP:
  pip install cloudinary gspread google-auth python-dotenv --break-system-packages

USAGE:
  python upload_to_cloudinary.py            # live upload + sheet write
  python upload_to_cloudinary.py --dry-run  # preview only, no upload/write/move
"""

import os
import sys
from datetime import date

import cloudinary
import cloudinary.uploader
import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

load_dotenv()

# ---------------- CONFIG — from .env, no hardcoded secrets ----------------

CLOUDINARY_CLOUD_NAME = os.getenv("CLOUDINARY_CLOUD_NAME")
CLOUDINARY_API_KEY = os.getenv("CLOUDINARY_API_KEY")
CLOUDINARY_API_SECRET = os.getenv("CLOUDINARY_API_SECRET")
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE")
SHEET_ID = os.getenv("SHEET_ID")
UPLOAD_DIR = os.getenv("UPLOAD_DIR")
ARCHIVE_DIR = os.getenv("ARCHIVE_DIR")

REQUIRED_ENV_VARS = {
    "CLOUDINARY_CLOUD_NAME": CLOUDINARY_CLOUD_NAME,
    "CLOUDINARY_API_KEY": CLOUDINARY_API_KEY,
    "CLOUDINARY_API_SECRET": CLOUDINARY_API_SECRET,
    "SERVICE_ACCOUNT_FILE": SERVICE_ACCOUNT_FILE,
    "SHEET_ID": SHEET_ID,
    "UPLOAD_DIR": UPLOAD_DIR,
    "ARCHIVE_DIR": ARCHIVE_DIR,
}

TRACKER_TAB = "image-library-tracker"  # sheet has this as its only/main tab

SOURCE_ID_COL = "Source Image ID"
UPLOADED_DATE_COL = "Uploaded Date"

# orientation -> sheet column that gets the Public ID
ORIENTATION_TO_COLUMN = {
    "landscape": "Landscape Public ID",
    "vertical": "Vertical Public ID",
}

VALID_EXTENSIONS = (".jpg", ".jpeg", ".png")

# -------------------------------------------------------


def check_env():
    missing = [name for name, val in REQUIRED_ENV_VARS.items() if not val]
    if missing:
        print(f"ERROR: missing required .env value(s): {', '.join(missing)}")
        sys.exit(1)


def parse_filename(fname):
    """
    Returns (source_image_id, orientation) or None if fname doesn't
    match {SourceImageID}-{landscape|vertical}-{descriptive-slug}.jpg.

    Only the first two hyphen-separated tokens are parsed -- Source
    Image ID (everything before the first hyphen) and orientation
    (the token right after it, must be exactly "landscape" or
    "vertical"). Whatever comes after that (the descriptive slug) is
    never parsed, even if it contains more hyphens.
    """
    base = os.path.splitext(fname)[0]
    parts = base.split("-", 2)  # source_id, orientation, rest-of-slug
    if len(parts) < 2:
        return None

    source_image_id, orientation = parts[0], parts[1].lower()
    if orientation not in ORIENTATION_TO_COLUMN:
        return None

    return source_image_id, orientation


def connect_sheet():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_ID)


def load_tracker_tab(sheet):
    """
    Loads the tracker tab once and returns (worksheet, header_index,
    source_id_to_row). header_index maps column name -> 0-based index,
    read fresh from the sheet's actual header row rather than assumed
    positions, so a column reorder never misaligns a write.
    """
    ws = sheet.worksheet(TRACKER_TAB)
    all_values = ws.get_all_values()
    if not all_values:
        print(f"ERROR: '{TRACKER_TAB}' tab is empty -- no header row found.")
        sys.exit(1)

    header = all_values[0]
    header_index = {name: i for i, name in enumerate(header)}

    for required_col in (SOURCE_ID_COL, UPLOADED_DATE_COL, *ORIENTATION_TO_COLUMN.values()):
        if required_col not in header_index:
            print(f"ERROR: expected column '{required_col}' not found in '{TRACKER_TAB}' header.")
            sys.exit(1)

    source_id_to_row = {}
    id_col_idx = header_index[SOURCE_ID_COL]
    for i, row in enumerate(all_values[1:], start=2):  # sheet rows are 1-indexed; row 1 is header
        if len(row) > id_col_idx and row[id_col_idx].strip():
            source_id_to_row[row[id_col_idx].strip()] = i

    print(f"Loaded '{TRACKER_TAB}' -- {len(source_id_to_row)} Source Image ID row(s).")
    return ws, header_index, source_id_to_row


def process_one_file(fname, dry_run, ws, header_index, source_id_to_row):
    """
    Returns True if the file was fully handled and should be archived
    (or, in a dry run, would be), False if it should be left in place.
    """
    parsed = parse_filename(fname)
    if parsed is None:
        print(f"\n[{fname}]\n  SKIPPED: filename doesn't match "
              f"{{SourceImageID}}-{{landscape|vertical}}-{{descriptive-slug}}.jpg -- "
              f"can't extract Source Image ID / orientation.")
        return False

    source_image_id, orientation = parsed
    public_id = os.path.splitext(fname)[0]
    target_column = ORIENTATION_TO_COLUMN[orientation]

    print(f"\n[{fname}]\n  Source Image ID: {source_image_id} | Orientation: {orientation} | "
          f"Public ID: {public_id}")

    row_num = source_id_to_row.get(source_image_id)
    if row_num is None:
        print(f"  WARNING: Source Image ID '{source_image_id}' not found in '{TRACKER_TAB}' -- "
              f"file left in place, nothing uploaded.")
        return False

    if dry_run:
        print(f"  Would upload to Cloudinary Public ID: {public_id}")
        print(f"  Would write '{public_id}' to column '{target_column}', row {row_num}")
        print(f"  Would stamp today's date ({date.today().strftime('%m/%d/%Y')}) into "
              f"'{UPLOADED_DATE_COL}', row {row_num}")
        return True

    local_path = os.path.join(UPLOAD_DIR, fname)
    try:
        result = cloudinary.uploader.upload(
            local_path, public_id=public_id, overwrite=True, resource_type="image"
        )
        uploaded_public_id = result["public_id"]
    except Exception as e:
        print(f"  FAILED to upload to Cloudinary: {e} -- file left in place.")
        return False

    print(f"  Uploaded. Public ID: {uploaded_public_id}")

    try:
        col_num = header_index[target_column] + 1  # gspread is 1-indexed
        date_col_num = header_index[UPLOADED_DATE_COL] + 1
        ws.update_cell(row_num, col_num, uploaded_public_id)
        ws.update_cell(row_num, date_col_num, date.today().strftime("%m/%d/%Y"))
        print(f"  Wrote '{uploaded_public_id}' to '{target_column}', row {row_num}. "
              f"Stamped '{UPLOADED_DATE_COL}'.")
    except Exception as e:
        print(f"  FAILED to write to sheet: {e} -- "
              f"image was uploaded to Cloudinary but the sheet was not updated. File left in place.")
        return False

    return True


def archive_file(fname):
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    src = os.path.join(UPLOAD_DIR, fname)
    dst = os.path.join(ARCHIVE_DIR, fname)
    os.replace(src, dst)


def main():
    dry_run = "--dry-run" in sys.argv

    check_env()

    cloudinary.config(
        cloud_name=CLOUDINARY_CLOUD_NAME,
        api_key=CLOUDINARY_API_KEY,
        api_secret=CLOUDINARY_API_SECRET,
        secure=True,
    )

    if not os.path.isdir(UPLOAD_DIR):
        print(f"ERROR: UPLOAD_DIR not found: {UPLOAD_DIR}")
        return

    files = [
        f for f in sorted(os.listdir(UPLOAD_DIR))
        if os.path.isfile(os.path.join(UPLOAD_DIR, f)) and f.lower().endswith(VALID_EXTENSIONS)
    ]

    if not files:
        print(f"No image files found in {UPLOAD_DIR}")
        return

    print(f"Found {len(files)} image file(s) in {UPLOAD_DIR}.")
    if dry_run:
        print("DRY RUN -- no uploads, sheet writes, or file moves will happen.\n")

    print("Connecting to sheet...")
    sheet = connect_sheet()
    ws, header_index, source_id_to_row = load_tracker_tab(sheet)

    success_count = 0
    left_in_place_count = 0

    for fname in files:
        handled = process_one_file(fname, dry_run, ws, header_index, source_id_to_row)
        if handled:
            success_count += 1
            if not dry_run:
                archive_file(fname)
        else:
            left_in_place_count += 1

    print()
    if dry_run:
        print(f"Dry run complete. {success_count} file(s) would be processed, "
              f"{left_in_place_count} would be left in place. "
              f"Re-run without --dry-run to actually upload and write to the sheet.")
    else:
        print(f"Done. {success_count} file(s) uploaded, written to the sheet, and archived to "
              f"{ARCHIVE_DIR}. {left_in_place_count} left in place -- see warnings above.")


if __name__ == "__main__":
    main()