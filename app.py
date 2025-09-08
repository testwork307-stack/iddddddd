# -*- coding: utf-8 -*-
"""
Streamlit – HR ID Card Generator (Arabic-aware, ZIP-only / repo-folder support)

Changes in this version (built for Streamlit Cloud):
- Removed RAR support (Streamlit Cloud cannot reliably run system `unrar`).
- Supports two photo sources:
    1) Upload a ZIP (extracted to a temporary directory)
    2) Use a photos folder bundled in the app repository (e.g. `photos/`)
- Simpler, robust extraction and cleanup logic.
- `requirements.txt` content is shown below (copy into a separate file in your repo):

# === requirements.txt ===
# Web / UI
streamlit>=1.36.0

# Data + Excel
pandas>=2.2.2
openpyxl>=3.1.3

# Images
pillow>=10.3.0
opencv-python>=4.10.0.84

# Barcode
python-barcode>=0.15.1

# Arabic support
arabic-reshaper>=3.0.0
python-bidi>=0.6.0

# Note: rarfile removed (not needed for Streamlit Cloud deploy)
# =========================

Place any photos folder you want to use in the app repository (for example `photos/`) or use the ZIP upload in the UI.
"""

import os
import io
import shutil
import zipfile
import tempfile
from pathlib import Path
from typing import Optional

import cv2
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from barcode import Code128
from barcode.writer import ImageWriter

import arabic_reshaper
from bidi.algorithm import get_display

# ===================== CONFIG =====================
PHOTO_POS = (111, 168)
PHOTO_SIZE = (300, 300)
BARCODE_POS = (570, 465)
BARCODE_SIZE = (390, 120)

# Fine-tune these two to move the *name* relative to original point:
NAME_OFFSET_X = -40  # negative = left
NAME_OFFSET_Y = -20  # negative = up

# ===================== UI =========================
st.set_page_config(page_title="HR ID Card Generator", page_icon="🎫", layout="wide")
st.title("🎫 HR ID Card Generator")

with st.sidebar:
    st.markdown("**Tips**")
    st.markdown("- Excel columns: **الاسم**, **الوظيفة**, **الرقم**, **الرقم القومي**, **الصورة**.")
    st.markdown("- Archive (ZIP) can have nested folders; the app searches recursively.")
    st.markdown("- Or place a `photos/` folder in your app repository and choose 'Use app folder'.")
    st.markdown("- Upload an Arabic TTF/OTF font for best Arabic rendering.")

    # Optional custom fonts
    font_ar_file = st.file_uploader("Arabic font (TTF/OTF, e.g., Amiri)", type=["ttf", "otf"], key="ar_font")
    font_en_file = st.file_uploader("English font (TTF/OTF)", type=["ttf", "otf"], key="en_font")

# Main inputs
excel_file = st.file_uploader("📂 Upload Excel (.xlsx)", type=["xlsx"], key="xlsx")
template_file = st.file_uploader("🖼 Upload Card Template (PNG/JPG)", type=["png", "jpg", "jpeg"], key="tpl")

# Photo source selector: either upload ZIP or use folder in repo
photos_source = st.selectbox("مصدر الصور (Photos source)", ["Upload ZIP", "Use app folder"]) 

photos_zip = None
photos_folder_in_repo = "photos"

if photos_source == "Upload ZIP":
    photos_zip = st.file_uploader("📦 Upload Photos (ZIP)", type=["zip"], key="archive_zip")
else:
    photos_folder_in_repo = st.text_input("Path to photos folder inside app (relative to repo root)", value="photos")

# ================== Helpers =======================

def load_font_from_upload(upload, fallback_name: str, size: int):
    """Load a font from an uploaded file; otherwise try common local fonts; otherwise PIL default."""
    if upload is not None:
        try:
            return ImageFont.truetype(io.BytesIO(upload.read()), size)
        except Exception:
            st.warning(f"⚠️ Failed to load uploaded font for {fallback_name}. Falling back to default.")

    # Fallbacks – try common installed fonts; finally PIL default
    for candidate in [
        "Amiri-Regular.ttf",
        "Amiri.ttf",
        "NotoNaskhArabic-Regular.ttf",
        "Arial.ttf",
        "Tahoma.ttf",
    ]:
        try:
            return ImageFont.truetype(candidate, size)
        except Exception:
            continue

    return ImageFont.load_default()


def prepare_text(text: str) -> str:
    """Arabic reshape + bidi for correct display."""
    if not text:
        return ""
    reshaped = arabic_reshaper.reshape(str(text))
    return get_display(reshaped)


def draw_aligned_text(draw: ImageDraw.ImageDraw, xy, text, font, fill="black", anchor="rt"):
    """Anchored text; multi-line supported line-by-line."""
    if not text:
        return
    lines = str(text).split("\n")
    x, y = xy
    for i, line in enumerate(lines):
        if i > 0:
            bbox = draw.textbbox((0, 0), line, font=font)
            y += (bbox[3] - bbox[1])
        draw.text((x, y), line, font=font, fill=fill, anchor=anchor)


def draw_bold_text(draw, xy, text, font, fill="black", anchor="rt"):
    """Fake-bold by layering 1px offsets (PIL-friendly)."""
    for dx, dy in [(0, 0), (1, 0), (0, 1), (1, 1)]:
        draw_aligned_text(draw, (xy[0] + dx, xy[1] + dy), text, font, fill=fill, anchor=anchor)


def find_photo_path(root_dir: str, requested: str) -> Optional[str]:
    """Find photo by stem match (case/ext-insensitive), search recursively."""
    if not requested:
        return None
    req_stem = Path(str(requested)).stem.lower()
    candidates = []
    for dirpath, _, filenames in os.walk(root_dir):
        for fn in filenames:
            if Path(fn).stem.lower() == req_stem:
                candidates.append(os.path.join(dirpath, fn))
    if candidates:
        order = {".png": 0, ".jpg": 1, ".jpeg": 2, ".bmp": 3}
        candidates.sort(key=lambda p: order.get(Path(p).suffix.lower(), 9))
        return candidates[0]
    return None


def crop_face_and_shoulders(image_path: str) -> Optional[Image.Image]:
    """Optional: crop around the first detected face area."""
    img = cv2.imread(image_path)
    if img is None:
        return None
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    faces = face_cascade.detectMultiScale(gray, 1.1, 5)
    if len(faces) == 0:
        return None
    x, y, w, h = faces[0]
    y_start = max(0, y - int(0.3 * h))
    y_end = min(img.shape[0], y + int(2.0 * h))
    x_start = max(0, x - int(0.3 * w))
    x_end = min(img.shape[1], x + int(1.3 * w))
    cropped = img[y_start:y_end, x_start:x_end]
    return Image.fromarray(cv2.cvtColor(cropped, cv2.COLOR_BGR2RGB))


# ================== Main logic ====================
if (excel_file is None) or (template_file is None):
    st.info("👆 Upload at least the Excel file and the template image to start. Choose your photos source too.")
else:
    # Fonts
    font_ar = load_font_from_upload(font_ar_file, "Arabic", 36)
    font_en = load_font_from_upload(font_en_file, "English", 30)

    # Read Excel
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"❌ Failed to read Excel: {e}")
        st.stop()

    # Read template
    try:
        template = Image.open(template_file).convert("RGB")
    except Exception as e:
        st.error(f"❌ Failed to read template image: {e}")
        st.stop()

    # Prepare photos source
    tmpdir = None
    photos_root = None

    if photos_source == "Upload ZIP":
        if photos_zip is None:
            st.error("⚠️ You chose 'Upload ZIP' but didn't upload a ZIP file.")
            st.stop()

        tmpdir = tempfile.mkdtemp(prefix="idcards_photos_")
        archive_path = os.path.join(tmpdir, photos_zip.name)
        with open(archive_path, "wb") as f:
            f.write(photos_zip.getbuffer())

        try:
            if archive_path.lower().endswith(".zip"):
                with zipfile.ZipFile(archive_path, "r") as zf:
                    zf.extractall(tmpdir)
                photos_root = tmpdir
            else:
                st.error("❌ Unsupported archive type. Upload a ZIP file.")
                shutil.rmtree(tmpdir, ignore_errors=True)
                st.stop()
        except Exception as e:
            st.error(f"❌ Failed to extract ZIP: {e}")
            shutil.rmtree(tmpdir, ignore_errors=True)
            st.stop()

    else:  # Use app folder
        candidate = Path.cwd() / photos_folder_in_repo
        if not candidate.exists() or not candidate.is_dir():
            st.error(f"⚠️ Photos folder not found in app repo: {candidate} . Put your images inside your repository (e.g. `photos/`).")
            st.stop()
        photos_root = str(candidate)

    # Process each employee
    output_cards: list[Image.Image] = []
    progress = st.progress(0)
    status = st.empty()

    for idx, row in df.iterrows():
        status.info(f"Processing {idx+1}/{len(df)} – {row.get('الاسم', '')}")
        card = template.copy()
        draw = ImageDraw.Draw(card)

        # Prepare texts (Arabic shaping + bidi)
        name = prepare_text(str(row.get("الاسم", "")).strip())
        job = prepare_text(str(row.get("الوظيفة", "")).strip())
        num = str(row.get("الرقم", "")).strip()
        national_id = str(row.get("الرقم القومي", "")).strip()
        photo_filename = str(row.get("الصورة", "")).strip()

        # ---- TEXT PLACEMENT ----
        base_name_xy = (915, 240)  # anchor reference from the design
        name_xy = (base_name_xy[0] + NAME_OFFSET_X, base_name_xy[1] + NAME_OFFSET_Y)
        draw_bold_text(draw, name_xy, name, font_ar, fill="black", anchor="rt")

        # JOB under name (+10)
        name_bbox = draw.textbbox((0, 0), name, font=font_ar)
        name_height = (name_bbox[3] - name_bbox[1]) + 20
        job_xy = (name_xy[0], name_xy[1] + name_height)
        draw_aligned_text(draw, job_xy, job, font=font_ar, fill="black", anchor="rt")

        # EMPLOYEE NUMBER under job (+15)
        job_bbox = draw.textbbox((0, 0), job, font=font_ar)
        job_height = (job_bbox[3] - job_bbox[1]) + 25
        job_id_label = prepare_text(f"الرقم الوظيفي: {num}")
        id_xy = (name_xy[0], job_xy[1] + job_height)
        draw_aligned_text(draw, id_xy, job_id_label, font=font_ar, fill="black", anchor="rt")

        # ---- PHOTO ----
        photo_path = find_photo_path(photos_root, photo_filename) if photos_root else None
        if photo_path and os.path.exists(photo_path):
            try:
                cropped = crop_face_and_shoulders(photo_path)
                img = cropped if cropped is not None else Image.open(photo_path)
                img = img.convert("RGB").resize(PHOTO_SIZE)
                card.paste(img, PHOTO_POS)
            except Exception as e:
                st.warning(f"⚠️ Failed to place photo for '{row.get('الاسم', '')}': {e}")
        else:
            st.warning(f"📷 Photo not found for '{row.get('الاسم', '')}'. Requested: {photo_filename}")

        # ---- BARCODE (using national ID) ----
        try:
            if national_id:
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_barcode:
                    out_noext = tmp_barcode.name[:-4]
                    barcode = Code128(national_id, writer=ImageWriter())
                    barcode_path = barcode.save(out_noext, {"write_text": False})
                with Image.open(barcode_path) as bimg:
                    bimg = bimg.convert("RGB").resize(BARCODE_SIZE)
                    card.paste(bimg, BARCODE_POS)
                for p in [out_noext + ".png", out_noext + ".svg"]:
                    try:
                        os.remove(p)
                    except Exception:
                        pass
            else:
                st.warning(f"🧾 National ID missing for '{row.get('الاسم', '')}'. Skipped barcode.")
        except Exception as e:
            st.warning(f"⚠️ Failed to generate barcode for '{row.get('الاسم', '')}': {e}")

        output_cards.append(card)
        progress.progress(int(((idx + 1) / max(len(df), 1)) * 100))

    status.empty()

    # ---- EXPORT PDF ----
    if output_cards:
        try:
            pdf_path = os.path.join(tempfile.gettempdir(), "All_ID_Cards.pdf")
            output_cards[0].save(pdf_path, save_all=True, append_images=output_cards[1:])
            with open(pdf_path, "rb") as f:
                st.download_button("⬇️ Download All ID Cards (PDF)", f, file_name="All_ID_Cards.pdf")
            st.success(f"✅ Generated {len(output_cards)} cards")
            st.image(output_cards[0], caption="Preview", width=320)
        except Exception as e:
            st.error(f"❌ Failed to write PDF: {e}")
    else:
        st.warning("No cards generated.")

    # Cleanup temp dir if used
    if tmpdir:
        try:
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass
