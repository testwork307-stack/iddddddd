# -*- coding: utf-8 -*-
"""
Streamlit HR ID Generator - Streamlit Cloud friendly (ZIP / repo-folder / ZIP URL)
Notes:
- Removed RAR support (not reliable on Streamlit Cloud).
- Removed OpenCV face-cropping to avoid heavy binary deps.
- Photos: upload ZIP, or provide ZIP URL, or use repo folder path (default: ./photos).
"""
import io
import os
import zipfile
import tempfile
from pathlib import Path

import requests
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from barcode import Code128
from barcode.writer import ImageWriter
import arabic_reshaper
from bidi.algorithm import get_display

# ---------------- Config ----------------
PHOTO_POS = (111, 168)
PHOTO_SIZE = (300, 300)

BARCODE_POS = (570, 465)
BARCODE_SIZE = (390, 120)

# Nudges for name
NAME_OFFSET_X = -40
NAME_OFFSET_Y = -20

# ---------------- UI ----------------
st.set_page_config(page_title="HR ID Card Generator", page_icon="🎫", layout="wide")
st.title("🎫 HR ID Card Generator (Streamlit Cloud)")

with st.sidebar:
    st.markdown("**خيارات مصدر الصور**")
    photos_source = st.radio(
        "اختر مصدر الصور",
        ("Upload ZIP", "Use repo folder", "ZIP URL"),
        index=0,
    )
    st.markdown("---")
    st.markdown("**خط (اختياري)**")
    font_ar_file = st.file_uploader("Arabic font (TTF/OTF) — اختياري", type=["ttf", "otf"])
    font_en_file = st.file_uploader("English font (TTF/OTF) — اختياري", type=["ttf", "otf"])

    st.markdown("---")
    st.markdown("ملاحظات:")
    st.markdown("- Streamlit Cloud: يفضل ZIP أو مجلد داخل الريبو (مثل `photos/`).")
    st.markdown("- تأكد أن أسماء الصور في الإكسل تتطابق مع أسماء الملفات أو الـ stems.")

# main inputs
excel_file = st.file_uploader("📂 Upload Excel (.xlsx)", type=["xlsx"])
template_file = st.file_uploader("🖼 Upload Card Template (PNG/JPG)", type=["png", "jpg", "jpeg"])

# photos inputs depending on choice
zip_upload = None
zip_url = None
repo_photos_path = None
if photos_source == "Upload ZIP":
    zip_upload = st.file_uploader("📦 Upload Photos ZIP", type=["zip"], key="zip")
elif photos_source == "ZIP URL":
    zip_url = st.text_input("ZIP URL (http/https)")
else:  # repo folder
    repo_photos_path = st.text_input("Repo photos folder (relative to app root)", value="photos")


# ---------------- Helpers ----------------
def load_font_from_upload(upload, fallback_name: str, size: int):
    if upload is not None:
        try:
            return ImageFont.truetype(io.BytesIO(upload.read()), size)
        except Exception:
            st.warning(f"⚠️ Failed to load uploaded font for {fallback_name}. Falling back.")
    # common fallback tries
    for candidate in ("Amiri-Regular.ttf", "Amiri.ttf", "NotoNaskhArabic-Regular.ttf", "Arial.ttf", "Tahoma.ttf"):
        try:
            return ImageFont.truetype(candidate, size)
        except Exception:
            continue
    return ImageFont.load_default()


def prepare_text(text: str) -> str:
    if not text:
        return ""
    reshaped = arabic_reshaper.reshape(str(text))
    return get_display(reshaped)


def draw_aligned_text(draw: ImageDraw.ImageDraw, xy, text, font, fill="black", anchor="rt"):
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
    for dx, dy in ((0, 0), (1, 0), (0, 1), (1, 1)):
        draw_aligned_text(draw, (xy[0] + dx, xy[1] + dy), text, font, fill=fill, anchor=anchor)


def resize_and_center_crop(pil_img: Image.Image, target_size: tuple[int, int]) -> Image.Image:
    tw, th = target_size
    img = pil_img.convert("RGB")
    iw, ih = img.size
    # scale so that the image fully covers the target (like CSS cover)
    scale = max(tw / iw, th / ih)
    new_w = int(iw * scale)
    new_h = int(ih * scale)
    img = img.resize((new_w, new_h), Image.LANCZOS)
    left = (new_w - tw) // 2
    top = (new_h - th) // 2
    img = img.crop((left, top, left + tw, top + th))
    return img


def find_photo_path(root_dir: str, requested: str):
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


def extract_zip_bytes_to(zip_bytes: bytes, dest: str):
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        zf.extractall(dest)


# ---------------- Main ----------------
if not excel_file or not template_file:
    st.info("👆 ارفع Excel و Template لبدء توليد البطاقات. اختر مصدر الصور في الشريط الجانبي.")
    st.stop()

# load fonts
font_ar = load_font_from_upload(font_ar_file, "Arabic", 36)
font_en = load_font_from_upload(font_en_file, "English", 30)

# read excel & template
try:
    df = pd.read_excel(excel_file)
except Exception as e:
    st.error(f"❌ Failed to read Excel: {e}")
    st.stop()

try:
    template = Image.open(template_file).convert("RGB")
except Exception as e:
    st.error(f"❌ Failed to read template image: {e}")
    st.stop()

# prepare photos root
tmpdir = None
created_tmpdir = False
photos_root = None
try:
    if photos_source == "Upload ZIP":
        if not zip_upload:
            st.error("📦 ارفع ملف ZIP للصور أو غيّر مصدر الصور.")
            st.stop()
        tmpdir = tempfile.mkdtemp(prefix="idcards_")
        created_tmpdir = True
        extract_zip_bytes_to(zip_upload.getbuffer(), tmpdir)
        photos_root = tmpdir
    elif photos_source == "ZIP URL":
        if not zip_url:
            st.error("🔗 اكتب رابط ZIP صحيح.")
            st.stop()
        tmpdir = tempfile.mkdtemp(prefix="idcards_")
        created_tmpdir = True
        try:
            r = requests.get(zip_url, timeout=30)
            r.raise_for_status()
            extract_zip_bytes_to(r.content, tmpdir)
            photos_root = tmpdir
        except Exception as e:
            st.error(f"❌ Failed to download/extract ZIP from URL: {e}")
            if created_tmpdir:
                try: shutil.rmtree(tmpdir) 
                except: pass
            st.stop()
    else:  # repo folder
        # Resolve relative to app root
        # Path(__file__).parent may not always be available in some runtimes; fallback to cwd
        try:
            app_root = Path(__file__).parent
        except Exception:
            app_root = Path.cwd()
        photos_root_candidate = app_root.joinpath(repo_photos_path)
        if not photos_root_candidate.exists():
            st.error(f"📁 Repo photos folder not found: {photos_root_candidate}")
            st.stop()
        photos_root = str(photos_root_candidate.resolve())
except Exception as e:
    st.error(f"❌ Error preparing photos: {e}")
    st.stop()

# Process rows
output_cards: list[Image.Image] = []
progress = st.progress(0)
status = st.empty()
total = len(df)

for idx, row in df.iterrows():
    status.info(f"Processing {idx+1}/{total} – {row.get('الاسم', '')}")
    card = template.copy()
    draw = ImageDraw.Draw(card)

    name = prepare_text(str(row.get("الاسم", "") or "").strip())
    job = prepare_text(str(row.get("الوظيفة", "") or "").strip())
    num = str(row.get("الرقم", "") or "").strip()
    national_id = str(row.get("الرقم القومي", "") or "").strip()
    photo_filename = str(row.get("الصورة", "") or "").strip()

    # NAME
    base_name_xy = (915, 240)
    name_xy = (base_name_xy[0] + NAME_OFFSET_X, base_name_xy[1] + NAME_OFFSET_Y)
    draw_bold_text(draw, name_xy, name, font_ar, fill="black", anchor="rt")

    # JOB (spacing +10)
    name_bbox = draw.textbbox((0, 0), name, font=font_ar)
    name_height = (name_bbox[3] - name_bbox[1]) + 20
    job_xy = (name_xy[0], name_xy[1] + name_height)
    draw_aligned_text(draw, job_xy, job, font=font_ar, fill="black", anchor="rt")

    # EMPLOYEE NUMBER (spacing +15)
    job_bbox = draw.textbbox((0, 0), job, font=font_ar)
    job_height = (job_bbox[3] - job_bbox[1]) + 25
    job_id_label = prepare_text(f"الرقم الوظيفي: {num}")
    id_xy = (name_xy[0], job_xy[1] + job_height)
    draw_aligned_text(draw, id_xy, job_id_label, font=font_ar, fill="black", anchor="rt")

    # PHOTO
    placed_photo = False
    if photos_root:
        photo_path = find_photo_path(photos_root, photo_filename)
        if photo_path and os.path.exists(photo_path):
            try:
                with Image.open(photo_path) as pimg:
                    pimg = resize_and_center_crop(pimg, PHOTO_SIZE)
                    card.paste(pimg, PHOTO_POS)
                    placed_photo = True
            except Exception as e:
                st.warning(f"⚠️ Failed to place photo for '{row.get('الاسم', '')}': {e}")
        else:
            st.warning(f"📷 Photo not found for '{row.get('الاسم', '')}'. Requested: {photo_filename}")
    else:
        st.warning("⚠️ photos_root not available.")

    # BARCODE (national id)
    try:
        if national_id:
            buf = io.BytesIO()
            barcode = Code128(national_id, writer=ImageWriter())
            barcode.write(buf, {"write_text": False})
            buf.seek(0)
            with Image.open(buf) as bimg:
                bimg = bimg.convert("RGB").resize(BARCODE_SIZE)
                card.paste(bimg, BARCODE_POS)
            buf.close()
        else:
            st.warning(f"🧾 National ID missing for '{row.get('الاسم', '')}'. Skipped barcode.")
    except Exception as e:
        st.warning(f"⚠️ Failed to generate barcode for '{row.get('الاسم', '')}': {e}")

    output_cards.append(card)
    progress.progress(int(((idx + 1) / max(total, 1)) * 100))

status.empty()

# Export
if output_cards:
    try:
        # Save multi-page PDF
        pdf_bytes = io.BytesIO()
        output_cards[0].save(pdf_bytes, format="PDF", save_all=True, append_images=output_cards[1:])
        pdf_bytes.seek(0)
        st.download_button("⬇️ Download All ID Cards (PDF)", pdf_bytes, file_name="All_ID_Cards.pdf")
        st.success(f"✅ Generated {len(output_cards)} cards")
        st.image(output_cards[0], caption="Preview", width=320)
    except Exception as e:
        st.error(f"❌ Failed to create/download PDF: {e}")
else:
    st.warning("No cards generated.")

# cleanup temp dir if created
if created_tmpdir and tmpdir:
    try:
        import shutil
        shutil.rmtree(tmpdir)
    except Exception:
        pass
