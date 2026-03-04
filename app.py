import os
import io
from datetime import date

import streamlit as st
import numpy as np
import pandas as pd
import cv2
from PIL import Image
import openpyxl

from docx import Document  # python-docx

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image as RLImage, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors

from rust_analyzer import analyze_rust_bgr

# ============================================================
# Page setup + header
# ============================================================
st.set_page_config(page_title="Rust Quantification – One Report", layout="wide")

if os.path.exists("logo.png"):
    c1, c2 = st.columns([1, 6])
    with c1:
        st.image("logo.png", use_container_width=True)
    with c2:
        st.title("Rust Quantification – One Report")
else:
    st.title("Rust Quantification – One Report")

# ============================================================
# Upload format support
# ============================================================
# Excel formats that openpyxl can load (zip-based)
EXCEL_EXTS = {"xlsx", "xlsm", "xltx", "xltm"}
# Word formats that python-docx can load (zip-based)
WORD_EXTS = {"docx", "docm", "dotx", "dotm"}

# Legacy binary formats (not supported for embedded-image extraction)
LEGACY_UNSUPPORTED = {"xls", "doc"}

# Option B: common raster formats PIL can read well
PHOTO_EXTS = {"png", "jpg", "jpeg", "bmp", "tif", "tiff", "webp"}

st.caption(
    "Option A: Upload Excel/Word report with embedded photos → ONE report. "
    "Option B: Upload one or multiple photos → ONE report. "
    "PDF includes thumbnails (Original / Overlay / Mask)."
)
st.divider()

MAX_PHOTOS_IN_PDF = 30  # hard cap as requested

# ============================================================
# Helpers
# ============================================================
def get_ext(filename: str) -> str:
    if not filename or "." not in filename:
        return ""
    return filename.rsplit(".", 1)[-1].lower().strip()

def get_severity(rust_pct, minor_thr, moderate_thr):
    if rust_pct < minor_thr:
        return "Minor"
    if rust_pct < moderate_thr:
        return "Moderate"
    return "Severe"

def extract_images_from_excel(xlsx_bytes):
    """
    Extract embedded images from Excel (xlsx/xlsm/xltx/xltm).
    """
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
    out = []
    for ws in wb.worksheets:
        for img in getattr(ws, "_images", []):
            try:
                out.append((f"Excel:{ws.title}", img._data()))
            except Exception:
                continue
    return out

def extract_images_from_docx(docx_bytes):
    """
    Extract embedded images from DOCX/DOCM/DOTX/DOTM.
    Returns list of (source_label, img_bytes)
    """
    doc = Document(io.BytesIO(docx_bytes))
    out = []
    seen = set()

    for idx, shape in enumerate(doc.inline_shapes, start=1):
        try:
            rid = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
            part = doc.part.related_parts[rid]
            img_bytes = part.blob
            key = (rid, len(img_bytes))
            if key not in seen:
                seen.add(key)
                out.append((f"WORD:Image {idx}", img_bytes))
        except Exception:
            continue

    return out

def pil_to_png_bytes(pil_img):
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    return buf.getvalue()

def make_thumb(pil_img, max_w=900):
    if pil_img.width <= max_w:
        return pil_img
    scale = max_w / pil_img.width
    return pil_img.resize(
        (int(pil_img.width * scale), int(pil_img.height * scale)),
        Image.LANCZOS
    )

def generate_pdf_with_thumbnails(report_meta, totals, per_photo_rows, photo_panels):
    buf = io.BytesIO()
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=28, leftMargin=28)
    story = []

    story.append(Paragraph("<b>Rust Inspection Report</b>", styles["Title"]))
    story.append(Spacer(1, 8))

    for k, v in report_meta.items():
        story.append(Paragraph(f"<b>{k}:</b> {v or '-'}", styles["Normal"]))
    story.append(Spacer(1, 12))

    story.append(Paragraph("<b>Summary</b>", styles["Heading2"]))
    story.append(Paragraph(f"Total Rust Area: <b>{totals['rust_pct']:.2f}%</b>", styles["Normal"]))
    story.append(Paragraph(f"Severity: <b>{totals['severity']}</b>", styles["Normal"]))
    story.append(Paragraph(f"Rust Pixels: {totals['rust_pixels']:,}", styles["Normal"]))
    story.append(Paragraph(f"Valid Pixels: {totals['valid_pixels']:,}", styles["Normal"]))
    story.append(Spacer(1, 12))

    if per_photo_rows:
        story.append(Paragraph("<b>Per-Photo Breakdown</b>", styles["Heading2"]))
        df = pd.DataFrame(per_photo_rows)
        tbl = Table([df.columns.tolist()] + df.values.tolist(), repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        story.append(tbl)
        story.append(Spacer(1, 14))

    story.append(Paragraph("<b>Photo Evidence</b>", styles["Heading2"]))
    story.append(Paragraph("Overlay: rust in red | Mask: white = rust", styles["Normal"]))
    story.append(Spacer(1, 10))

    capped = photo_panels[:MAX_PHOTOS_IN_PDF]
    for i, p in enumerate(capped, 1):
        story.append(Paragraph(f"<b>{p['title']}</b>", styles["Heading3"]))
        imgs = [
            RLImage(io.BytesIO(p["orig"]), 170, 120),
            RLImage(io.BytesIO(p["overlay"]), 170, 120),
            RLImage(io.BytesIO(p["mask"]), 170, 120),
        ]
        grid = Table(
            [["Original", "Overlay", "Mask"], imgs],
            colWidths=[170, 170, 170]
        )
        grid.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ]))
        story.append(grid)
        story.append(Spacer(1, 12))
        if i % 3 == 0:
            story.append(PageBreak())

    if len(photo_panels) > MAX_PHOTOS_IN_PDF:
        story.append(Spacer(1, 10))
        story.append(Paragraph(
            f"Note: PDF includes only first {MAX_PHOTOS_IN_PDF} photos (limit).",
            styles["Italic"]
        ))

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()

# ============================================================
# Sidebar settings
# ============================================================
with st.sidebar:
    st.header("Analysis Settings")
    exclude_dark = st.checkbox("Exclude dark pixels", True)
    min_v = st.slider("Shadow threshold (V)", 0, 255, 35)
    kernel_size = st.slider("Kernel size", 1, 21, 5, 2)
    open_iters = st.slider("Open iters", 0, 5, 1)
    close_iters = st.slider("Close iters", 0, 5, 2)
    minor_thr = st.number_input("Minor < (%)", value=5.0)
    moderate_thr = st.number_input("Moderate < (%)", value=15.0)

# ============================================================
# New Inspection (ONE REPORT)
# ============================================================
st.subheader("New Inspection (One Report)")

with st.form("inspect"):
    c1, c2, c3 = st.columns(3)
    with c1:
        insp_date = st.date_input("Inspection Date", date.today())
        vessel = st.text_input("Vessel Name")
    with c2:
        tank = st.text_input("Tank / Hold")
        location = st.text_input("Location")
    with c3:
        inspector = st.text_input("Inspector")
        remarks = st.text_area("Remarks")

    st.info(
        "✅ Choose ONE option:\n\n"
        "• Option A: Upload Excel/Word report with embedded photos\n"
        "• Option B: Upload one OR multiple photos\n\n"
        "⚠️ Do NOT use both options at the same time."
    )

    report_file = st.file_uploader(
        "Option A – Upload report file (Excel/Word with embedded photos)",
        type=sorted(list(EXCEL_EXTS | WORD_EXTS | LEGACY_UNSUPPORTED)),
        key="report_upl"
    )

    st.info(
        "📸 Option B: Upload one OR multiple photos.\n"
        "Supported: PNG/JPG/JPEG/BMP/TIF/TIFF/WEBP\n"
        "Use Ctrl/Shift to select multiple files."
    )
    photos = st.file_uploader(
        "Option B – Photo(s)",
        type=sorted(list(PHOTO_EXTS)),
        accept_multiple_files=True,
        key="photos_upl"
    )

    run = st.form_submit_button("Analyze (One Report)")

# ============================================================
# Analysis
# ============================================================
if run:
    use_report = report_file is not None
    use_photos = photos is not None and len(photos) > 0

    if not use_report and not use_photos:
        st.error("Upload a report file (Excel/Word) OR at least one photo.")
        st.stop()

    if use_report and use_photos:
        st.error("Please use only ONE option: report file OR photo(s).")
        st.stop()

    rust_px_total = 0
    valid_px_total = 0
    per_photo_rows = []
    photo_panels = []

    # -----------------------------
    # Option A: Report file (Excel or Word)
    # -----------------------------
    if use_report:
        fname = report_file.name
        ext = get_ext(fname)
        b = report_file.getvalue()

        if ext in LEGACY_UNSUPPORTED:
            st.error(
                f"'{ext.upper()}' format is not supported for embedded-image extraction.\n\n"
                f"Please open the file and **Save As**:\n"
                f"• Excel: .xlsx / .xlsm / .xltx / .xltm\n"
                f"• Word:  .docx / .docm / .dotx / .dotm"
            )
            st.stop()

        if ext in EXCEL_EXTS:
            extracted = extract_images_from_excel(b)
        elif ext in WORD_EXTS:
            extracted = extract_images_from_docx(b)
        else:
            st.error("Unsupported report type.")
            st.stop()

        if not extracted:
            st.error("No embedded images found. Please ensure photos are inserted/embedded in the report.")
            st.stop()

        for i, (src, img_bytes) in enumerate(extracted, 1):
            try:
                pil = Image.open(io.BytesIO(img_bytes)).convert("RGB")
            except Exception:
                # Some embedded objects may not be standard raster images
                continue

            bgr = cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)
            pct, rp, vp, mask, overlay = analyze_rust_bgr(
                bgr,
                exclude_dark_pixels=exclude_dark,
                min_v_for_valid=min_v,
                kernel_size=kernel_size,
                open_iters=open_iters,
                close_iters=close_iters,
            )

            rust_px_total += int(rp)
            valid_px_total += int(vp)

            per_photo_rows.append({
                "Photo": i,
                "Source": src,
                "Rust %": round(float(pct), 2)
            })

            photo_panels.append({
                "title": f"Photo {i} ({src})",
                "orig": pil_to_png_bytes(make_thumb(pil)),
                "overlay": pil_to_png_bytes(make_thumb(Image.fromarray(cv2.cvtColor(overlay, cv2.COLOR_BGR2RGB)))),
                "mask": pil_to_png_bytes(make_thumb(Image.fromarray(mask).convert("RGB"))),
            })

        input_type = f"Report file: {report_file.name}"

    # -----------------------------
    # Option B: Multiple photos -> ONE report
    # -----------------------------
    else:
        photos_sorted = sorted(photos, key=lambda f: f.name.lower())

        for i, f in enumerate(photos_sorted, 1):
            try:
                pil = Image.open(f).convert("RGB")
            except Exception:
                st.warning(f"Skipped unsupported/invalid image: {f.name}")
                continue

            bgr = cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR)

            pct, rp, vp, mask, overlay = analyze_rust_bgr(
                bgr,
                exclude_dark_pixels=exclude_dark,
                min_v_for_valid=min_v,
                kernel_size=kernel_size,
                open_iters=open_iters,
                close_iters=close_iters,
            )

            rust_px_total += int(rp)
            valid_px_total += int(vp)

            per_photo_rows.append({
                "Photo": i,
                "Source": f.name,
                "Rust %": round(float(pct), 2)
            })

            photo_panels.append({
                "title": f"Photo {i} ({f.name})",
                "orig": pil_to_png_bytes(make_thumb(pil)),
                "overlay": pil_to_png_bytes(make_thumb(Image.fromarray(cv2.cvtColor(overlay, cv2.COLOR_BGR2RGB)))),
                "mask": pil_to_png_bytes(make_thumb(Image.fromarray(mask).convert("RGB"))),
            })

        input_type = f"Photos uploaded: {len(photos)} files"

    if valid_px_total <= 0:
        st.error("Valid pixels total is 0. Photos may be too dark or invalid.")
        st.stop()

    rust_pct_total = 100 * rust_px_total / max(valid_px_total, 1)
    severity = get_severity(rust_pct_total, float(minor_thr), float(moderate_thr))

    st.success(f"TOTAL Rust: {rust_pct_total:.2f}% | Severity: {severity}")

    report_meta = {
        "Vessel": vessel,
        "Tank / Hold": tank,
        "Location": location,
        "Inspection Date": str(insp_date),
        "Inspector": inspector,
        "Remarks": remarks,
        "Input Type": input_type
    }

    totals = {
        "rust_pct": float(rust_pct_total),
        "severity": severity,
        "rust_pixels": int(rust_px_total),
        "valid_pixels": int(rust_px_total) if False else int(valid_px_total),  # keep valid_px_total
    }
    totals["valid_pixels"] = int(valid_px_total)

    pdf_bytes = generate_pdf_with_thumbnails(
        report_meta=report_meta,
        totals=totals,
        per_photo_rows=per_photo_rows,
        photo_panels=photo_panels
    )

    safe_vessel = (vessel or "inspection").replace(" ", "_")
    st.download_button(
        "📄 Download Inspection Report (PDF)",
        data=pdf_bytes,
        file_name=f"rust_report_{safe_vessel}.pdf",
        mime="application/pdf"
    )
