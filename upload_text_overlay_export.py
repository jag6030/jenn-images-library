"""
Jenn Gabel Text Overlay Export -- Cloudinary Uploader + Sheet Writer
----------------------------------------------------------------------
Companion to upload_to_cloudinary.py, for TextOverlay_TEMPLATE.html's ZIP
export instead of raw source photos.

WHY THIS IS A SEPARATE SCRIPT, NOT A REUSE OF upload_to_cloudinary.py:
  upload_to_cloudinary.py handles ONE base photo per orientation
  (landscape/vertical) and stores only its Public ID -- the sheet's
  "<size> URL" columns are CONCATENATE formulas that build all 7 URLs from
  that one Public ID + Cloudinary's named transformations.

  A text-overlay export has no single croppable source -- the template
  bakes text into 7 INDEPENDENTLY exported files (text placement/size
  differs per crop), each already at final resolution. So this script:
    - uploads each of the 7 files as its own Cloudinary asset
    - writes the returned secure_url DIRECTLY into that row's own
      "<sizeCode> URL" column (no transformation string involved)
    - never touches "Landscape Public ID" / "Vertical Public ID" -- those
      columns don't apply to a text-output row and are deliberately left
      blank/untouched

FILENAME CONVENTION (required, matches TextOverlay_TEMPLATE.html's export):
  {Source Image ID}-{sizeCode}.jpg
  e.g. si0003-l2160x1080.jpg, si0003-v1080x1920.jpg

  sizeCode must be one of the 7 the template produces (square s1080x1080 is
  never produced by this template and is not handled here):
    l2160x1080, l1440x1080, l1200x900, l1920x1080,
    v1080x1350, v1080x1440, v1080x1920

  Cloudinary Public ID = the full filename with its extension stripped,
  e.g. "si0003-l2160x1080".

MATCHING: the row is matched by Source Image ID EXACTLY (e.g. si0003) --
never by BD ID. This script only ever writes to the row whose OWN images
these are; BD ID is only read by TextOverlay_WebApp.gs at preview time.

Rows are NEVER created by this script -- only updated. If a file's Source
Image ID has no matching row, or its size code isn't recognized, the file
is left in place (not archived) and a warning is logged. The script never
crashes on a bad or unmatched file; it skips that file and keeps going.

SHEET COLUMNS TOUCHED (by header name, not position):
  each matched file -> "<sizeCode> URL" (e.g. "l2160x1080 URL")
  either             -> "Uploaded Date" (today's date, stamped every time
                         this script successfully writes a URL)

No other column (Status, Tags, Notes, Landscape/Vertical Public ID, etc.)
is ever read or written.

CREDENTIALS: loaded from .env via python-dotenv -- reuses the same
CLOUDINARY_CLOUD_NAME, CLOUDINARY_API_KEY, CLOUDINARY_API_SECRET,
SERVICE_ACCOUNT_FILE, SHEET_ID as upload_to_cloudinary.py, plus two NEW
.env values specific to this script:
  TEXT_OVERLAY_EXPORT_DIR  -- where you unzip the template's download
  TEXT_OVERLAY_ARCHIVE_DIR -- where successfully-processed files move to

SETUP:
  pip install cloudinary gspread google-auth python-dotenv --break-system-packages

USAGE:
  python upload_text_overlay_export.py            # live upload + sheet write
  python upload_text_overlay_export.py --dry-run  # preview only, no upload/write/move
"""

import os
import re
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
EXPORT_DIR = os.getenv("TEXT_OVERLAY_EXPORT_DIR")
ARCHIVE_DIR = os.getenv("TEXT_OVERLAY_ARCHIVE_DIR")

REQUIRED_ENV_VARS = {
    "CLOUDINARY_CLOUD_NAME": CLOUDINARY_CLOUD_NAME,
    "CLOUDINARY_API_KEY": CLOUDINARY_API_KEY,
    "CLOUDINARY_API_SECRET": CLOUDINARY_API_SECRET,
    "SERVICE_ACCOUNT_FILE": SERVICE_ACCOUNT_FILE,
    "SHEET_ID": SHEET_ID,
    "TEXT_OVERLAY_EXPORT_DIR": EXPORT_DIR,
    "TEXT_OVERLAY_ARCHIVE_DIR": ARCHIVE_DIR,
}

TRACKER_TAB = "image-library-tracker"

SOURCE_ID_COL = "Source Image ID"
UPLOADED_DATE_COL = "Uploaded Date"

# The 7 size codes this template produces (square is deliberately excluded —
# text overlay isn't used on that crop), mapped to their sheet URL column.
SIZE_CODE_TO_COLUMN = {
    "l2160x1080": "l2160x1080 URL",
    "l1440x1080": "l1440x1080 URL",
    "l1200x900": "l1200x900 URL",
    "l1920x1080": "l1920x1080 URL",
    "v1080x1350": "v1080x1350 URL",
    "v1080x1440": "v1080x1440 URL",
    "v1080x1920": "v1080x1920 URL",
}

VALID_EXTENSIONS = (".jpg", ".jpeg", ".png")

# {SourceImageID}-{sizeCode}.ext — source ID is everything before the LAST
# hyphen-group that matches a known size code, so a source ID containing its
# own hyphens (not currently the case, but future-proofed) wouldn't break.
FILENAME_RE = re.compile(
    r"^(?P<source_id>.+)-(?P<size_code>l\d+x\d+|v\d+x\d+)$",
    re.IGNORECASE,
)

# -------------------------------------------------------


def check_env():
    missing = [name for name, val in REQUIRED_ENV_VARS.items() if not val]
    if missing:
        print(f"ERROR: missing required .env value(s): {', '.join(missing)}")
        print("(TEXT_OVERLAY_EXPORT_DIR / TEXT_OVERLAY_ARCHIVE_DIR are new — "
              "add them to .env alongside the existing upload_to_cloudinary.py values.)")
        sys.exit(1)


def parse_filename(fname):
    """
    Returns (source_image_id, size_code) or None if fname doesn't match
    {SourceImageID}-{sizeCode}.jpg, or the size code isn't one of the 7
    this template produces.
    """
    base = os.path.splitext(fname)[0]
    m = FILENAME_RE.match(base)
    if not m:
        return None

    source_image_id = m.group("source_id")
    size_code = m.group("size_code").lower()
    if size_code not in SIZE_CODE_TO_COLUMN:
        return None

    return source_image_id, size_code


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

    for required_col in (SOURCE_ID_COL, UPLOADED_DATE_COL, *SIZE_CODE_TO_COLUMN.values()):
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
              f"{{SourceImageID}}-{{sizeCode}}.jpg with a recognized size code -- "
              f"can't extract Source Image ID / size code.")
        return False

    source_image_id, size_code = parsed
    public_id = os.path.splitext(fname)[0]
    target_column = SIZE_CODE_TO_COLUMN[size_code]

    print(f"\n[{fname}]\n  Source Image ID: {source_image_id} | Size: {size_code} | "
          f"Public ID: {public_id}")

    row_num = source_id_to_row.get(source_image_id)
    if row_num is None:
        print(f"  WARNING: Source Image ID '{source_image_id}' not found in '{TRACKER_TAB}' -- "
              f"file left in place, nothing uploaded.")
        return False

    if dry_run:
        print(f"  Would upload to Cloudinary Public ID: {public_id}")
        print(f"  Would write the resulting URL to column '{target_column}', row {row_num}")
        print(f"  Would stamp today's date ({date.today().strftime('%m/%d/%Y')}) into "
              f"'{UPLOADED_DATE_COL}', row {row_num}")
        return True

    local_path = os.path.join(EXPORT_DIR, fname)
    try:
        result = cloudinary.uploader.upload(
            local_path, public_id=public_id, overwrite=True, resource_type="image"
        )
        uploaded_url = result["secure_url"]
    except Exception as e:
        print(f"  FAILED to upload to Cloudinary: {e} -- file left in place.")
        return False

    print(f"  Uploaded. URL: {uploaded_url}")

    try:
        col_num = header_index[target_column] + 1  # gspread is 1-indexed
        date_col_num = header_index[UPLOADED_DATE_COL] + 1
        ws.update_cell(row_num, col_num, uploaded_url)
        ws.update_cell(row_num, date_col_num, date.today().strftime("%m/%d/%Y"))
        print(f"  Wrote URL to '{target_column}', row {row_num}. Stamped '{UPLOADED_DATE_COL}'.")
    except Exception as e:
        print(f"  FAILED to write to sheet: {e} -- "
              f"image was uploaded to Cloudinary but the sheet was not updated. File left in place.")
        return False

    return True


def archive_file(fname):
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    src = os.path.join(EXPORT_DIR, fname)
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

    if not os.path.isdir(EXPORT_DIR):
        print(f"ERROR: TEXT_OVERLAY_EXPORT_DIR not found: {EXPORT_DIR}")
        return

    files = [
        f for f in sorted(os.listdir(EXPORT_DIR))
        if os.path.isfile(os.path.join(EXPORT_DIR, f)) and f.lower().endswith(VALID_EXTENSIONS)
    ]

    if not files:
        print(f"No image files found in {EXPORT_DIR}")
        return

    print(f"Found {len(files)} image file(s) in {EXPORT_DIR}.")
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