#!/usr/bin/env python3
"""Regenerates images/previews/*.png for the Vendor Sizing Template Panel.

Cross-platform (Mac and Windows) — uses Pillow instead of macOS's `sips`,
so it works the same on both. Install Pillow once if needed:

    pip install Pillow

Run this after adding or replacing any photo in images/. To add a preview
photo for a new item: name the file exactly

    <MAIN_CATEGORY>__<SUB_CATEGORY>.png

(sanitized the same way the panel does: uppercase, any run of
non-alphanumeric characters squashed to a single underscore — e.g. CSV
category "GENERAL-PRODUCTS" + item "Bracelet_Add_Shot" -> file
"GENERAL_PRODUCTS__BRACELET_ADD_SHOT.png"), drop it directly in images/
(not images/previews/), and re-run this script. The crop ratio (1:1, 4:5,
whatever) is looked up automatically from BCOM_VENDOR_SIZINGS.csv by that
same key — nothing else needs editing.
"""
import csv
import os
import re

from PIL import Image

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGES_DIR = os.path.join(SCRIPT_DIR, "images")
PREVIEWS_DIR = os.path.join(IMAGES_DIR, "previews")
CSV_PATH = os.path.join(SCRIPT_DIR, "BCOM_VENDOR_SIZINGS.csv")
PREVIEW_TARGET_WIDTH = 220  # must match PREVIEW_TARGET_WIDTH in the .jsx panel exactly
PAD_COLOR = (255, 255, 255)  # white, matching the vendor's own white-ground photography

# Column indices — same fixed layout the panel's own CSV parser uses (the
# source file's headers have typos like "Widh" / "Horizotal Allignment",
# so this reads by position rather than by header name matching).
COL_MAIN_CATEGORY = 0
COL_ITEM_NAME = 1
COL_WIDTH = 2
COL_HEIGHT = 4


def sanitize_for_filename(text):
    text = text.strip().upper()
    text = re.sub(r"[^A-Z0-9]+", "_", text)
    return text.strip("_")


def load_csv_ratios():
    with open(CSV_PATH, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))

    ratios = {}
    for row in rows[1:]:  # skip header
        if len(row) <= COL_HEIGHT or not row[COL_MAIN_CATEGORY].strip():
            continue
        try:
            width = float(row[COL_WIDTH])
            height = float(row[COL_HEIGHT])
        except ValueError:
            continue  # blank/non-numeric width or height — no crop ratio to derive

        key = sanitize_for_filename(row[COL_MAIN_CATEGORY]) + "__" + sanitize_for_filename(row[COL_ITEM_NAME])
        ratios[key] = (width, height)
    return ratios


def generate_one(key, csv_w, csv_h, src, dest):
    preview_w = PREVIEW_TARGET_WIDTH
    preview_h = round(PREVIEW_TARGET_WIDTH * csv_h / csv_w)

    with Image.open(src) as im:
        # Flatten any transparency onto the same white ground the vendor
        # shoots against, so a photo with an alpha channel doesn't end up
        # with a black (or undefined) matte around it in the preview.
        if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
            im = im.convert("RGBA")
            flattened = Image.new("RGB", im.size, PAD_COLOR)
            flattened.paste(im, mask=im.split()[-1])
            im = flattened
        else:
            im = im.convert("RGB")

        src_w, src_h = im.size
        scale = min(preview_w / src_w, preview_h / src_h)
        resized_w = max(1, round(src_w * scale))
        resized_h = max(1, round(src_h * scale))
        resized = im.resize((resized_w, resized_h), Image.LANCZOS)

    canvas = Image.new("RGB", (preview_w, preview_h), PAD_COLOR)
    offset_x = (preview_w - resized_w) // 2
    offset_y = (preview_h - resized_h) // 2
    canvas.paste(resized, (offset_x, offset_y))
    canvas.save(dest, "PNG")

    print("OK: %s -> %dx%d (source %dx%d, CSV ratio %gx%g)" % (
        key, preview_w, preview_h, src_w, src_h, csv_w, csv_h
    ))


def main():
    ratios = load_csv_ratios()
    os.makedirs(PREVIEWS_DIR, exist_ok=True)

    source_files = sorted(
        f for f in os.listdir(IMAGES_DIR)
        if f.lower().endswith(".png") and os.path.isfile(os.path.join(IMAGES_DIR, f))
    )

    if not source_files:
        print("No .png files found directly in " + IMAGES_DIR)
        return

    for filename in source_files:
        key = filename[:-4]  # strip ".png"
        src = os.path.join(IMAGES_DIR, filename)
        dest = os.path.join(PREVIEWS_DIR, filename)

        if key not in ratios:
            print("SKIP (no matching CSV row for key \"%s\"): %s" % (key, filename))
            continue

        csv_w, csv_h = ratios[key]
        generate_one(key, csv_w, csv_h, src, dest)


if __name__ == "__main__":
    main()
