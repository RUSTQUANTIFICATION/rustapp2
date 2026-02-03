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
# HARD-LOCKED COMPANY BASELINE SETTINGS (crew cannot change)
# ============================================================
exclude_dark = True          # Always exclude shadows
min_v = 40                   # Shadow threshold (V)
kernel_size = 5              # Morphology kernel
open_iters = 1               # Remove small noise
close_iters = 2              # Fill small gaps

minor_thr = 5.0              # Severity thresholds
moderate_thr = 15.0

alert_rust_pct = 10.0        # Fleet alert thresholds
alert_increase_pct = 2.0

MAX_PHOTOS_IN_PDF = 30       # PDF photo cap

# ============================================================
# Fleet master data (EDIT these 25 vessels)
# ============================================================
FLEET_MASTER = [
    {"Vessel": "ASIA UNITY", "PaintManufacturer": "Jotun", "Yard": "COSCO Guangzhou", "CoatingAppliedDate": "2025-03-15", "CoatingNotes": "Ballast tank / touch-up"},
    {"Vessel": "ASIA LIBERTY", "PaintManufacturer": "Hempel", "Yard": "Dalian", "CoatingAppliedDate": "2024-11-10", "CoatingNotes": "Cargo hold coating"},
    {"Vessel": "ASIA EVERGREEN", "PaintManufacturer": "International", "Yard": "Zhoushan", "CoatingAppliedDate": "2024-07-05", "CoatingNotes": "Ballast tank full coat"},
    {"Vessel": "ASIA ASPARA", "PaintManufacturer": "Jotun", "Yard": "COSCO", "CoatingAppliedDate": "2023-12-20", "CoatingNotes": "Spot repair"},
    {"Vessel": "ASIA INSPIRE", "PaintManufacturer": "Hempel", "Yard": "Guangzhou", "CoatingAppliedDate": "2025-01-18", "CoatingNotes": "Hold coating"},
]

# Auto-fill placeholders up to 25 if not provided yet
while len(FLEET_MASTER) < 25:
    idx = len(FLEET_MASTER) + 1
    FLEET_MASTER.append({
        "Vessel": f"VESSEL-{idx:02d}",
        "PaintManufacturer": "TBD",
        "Yard": "TBD",
        "CoatingAppliedDate": "2025-01-01",
        "CoatingNotes": ""
    })

master_df_default = pd.DataFrame(FLEET_MASTER)

# ============================================================
# Page setup + branding
# ============================================================
st.set_page_config(page_title="Fleet Rust Monitoring – One Report", layout="wide")

if os.path.exists("logo.png"):
    c1, c2 = st.columns([1, 6])
    with c1:
        st.image("logo.png", use_container_width=True)
    with c2:
        st.title("Fleet Rust Monitoring – One Report")
else:
    st.title("Fleet Rust Monitoring – One Report")

st.caption(
    "Objective: (a) react proactively when rust increases, "
    "(b) identify poor-performing paint manufacturers or yards using trend data."
)
st.divider()

# ============================================================
# Sidebar (READ-ONLY, locked settings)
# ============================================================
with st.sidebar:
    st.header("Analysis Settings (Locked)")
    st.info("Company baseline settings are locked for consistency across fleet.")
    st.write(f"Exclude dark pixels: {exclude_dark}")
    st.write(f"Shadow threshold (V): {min_v}")
    st.write(f"Kernel size: {kernel_size}")
    st.write(f"Open iters: {open_iters}")
    st.write(f"Close iters: {close_iters}")
    st.divider()
    st.write(f"Severity: Minor < {minor_thr}%, Moderate < {moderate_thr}%")
    st.write(f"Alerts: Rust ≥ {alert_rust_pct}% OR Increase ≥ {alert_increase_pct}%")

# ============================================================
# Helpers
# ============================================================
def season_from_month(m: int) -> str:
    if m in (12, 1, 2): return "Winter"
    if m in (3, 4, 5): return "Spring"
    if m in (6, 7, 8): return "Summer"
    return "Autumn"

def parse_date_safe(s):
    try:
        return pd.to_datetime(s).date()
    except Exception:
        return None

def months_between(d1: date, d2: date) -> float:
    return max((d1 - d2).days / 30.44, 0.0)

def get_severity(rust_pct, minor_thr_, moderate_thr_):
    if rust_pct < minor_thr_:
        return "Minor"
    if rust_pct < moderate_thr_:
        return "Moderate"
    return "Severe"

def extract_images_from_excel(xlsx_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
    out = []
    for ws in wb.worksheets:
        for img in getattr(ws, "_images", []):
            out.append((f"Excel:{ws.title}", img._data()))
    return out

def extract_images_from_docx(docx_bytes):
    doc = Document(io.BytesIO(docx_bytes))
    out, seen = [], set()
    for idx, shape in enumerate(doc.inline_shapes, start=1):
        try:
            rid = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
            part = doc.part.related_parts[rid]
            img_bytes = part.blob
            key = (rid, len(img_bytes))
            if key not in seen:
                seen.add(key)
                out.append((f"DOCX:Image {idx}", img_bytes))
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
        grid = Table([["Original", "Overlay", "Mask"], imgs], colWidths=[170, 170, 170])
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

def to_excel_bytes(df: pd.DataFrame, sheet_name="Sheet1") -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    out.seek(0)
    return out.getvalue()

# ============================================================
# State init
# ============================================================
if "master_df" not in st.session_state:
    st.session_state.master_df = master_df_default.copy()

if "log_df" not in st.session_state:
    st.session_state.log_df = pd.DataFrame(columns=[
        "InspectionDate", "Vessel", "Tank/Hold", "Location", "Inspector",
        "PaintManufacturer", "Yard", "CoatingAppliedDate", "CoatingSeason", "CoatingNotes",
        "InputType", "NumPhotosAnalyzed",
        "TotalRustPct", "Severity",
        "RustPixels", "ValidPixels",
        "Remarks"
    ])

# ============================================================
# Tabs
# ============================================================
tab1, tab2 = st.tabs(["New Inspection (One Report)", "Fleet Dashboard"])

# ============================================================
# TAB 1: New Inspection
# ============================================================
with tab1:
    st.subheader("New Inspection (One Report)")

    master_df = st.session_state.master_df
    vessel_list = master_df["Vessel"].tolist()
# --- Vessel selection OUTSIDE the form (important) ---
vessel = st.selectbox("Vessel", vessel_list, key="vessel_select")

# Read master data for selected vessel
row = master_df[master_df["Vessel"] == vessel].iloc[0]

default_paint = str(row["PaintManufacturer"])
default_yard = str(row["Yard"])
default_coating_applied = str(row["CoatingAppliedDate"])
default_notes = str(row.get("CoatingNotes", ""))
with st.form("inspect"):
    c1, c2, c3 = st.columns(3)
    with c1:
        insp_date = st.date_input("Inspection Date", date.today())
    with c2:
        tank = st.text_input("Tank / Hold")
        location = st.text_input("Location")
    with c3:
        inspector = st.text_input("Inspector")
        remarks = st.text_area("Remarks", height=90)

    # Auto-filled from Fleet Master (updates when vessel changes)
    paint_maker = st.text_input("Paint Manufacturer", value=default_paint)
    yard = st.text_input("Yard (coating applied)", value=default_yard)
    coating_applied = st.text_input(
        "Coating Applied Date (YYYY-MM-DD)",
        value=default_coating_applied
    )
    coating_notes = st.text_input(
        "Coating notes/system",
        value=default_notes
    )   
        with c2:
            tank = st.text_input("Tank / Hold")
            location = st.text_input("Location")
        with c3:
            inspector = st.text_input("Inspector")
            remarks = st.text_area("Remarks", height=90)

        # Auto-fill from master
        row = master_df[master_df["Vessel"] == vessel].iloc[0]
        paint_maker = st.text_input("Paint Manufacturer", value=str(row["PaintManufacturer"]))
        yard = st.text_input("Yard (coating applied)", value=str(row["Yard"]))
        coating_applied = st.text_input("Coating Applied Date (YYYY-MM-DD)", value=str(row["CoatingAppliedDate"]))
        coating_notes = st.text_input("Coating notes/system", value=str(row.get("CoatingNotes", "")))

        st.info(
            "Choose ONE input option:\n"
            "• Option A: Upload report file (.xlsx or .docx) with embedded photos\n"
            "• Option B: Upload one OR multiple photos (PNG/JPG/JPEG)\n\n"
            "Do NOT use both options together."
        )

        report_file = st.file_uploader("Option A – Report file (.xlsx or .docx)", ["xlsx", "docx"])
        photos = st.file_uploader(
            "Option B – Photo(s) (PNG/JPG/JPEG)",
            ["png", "jpg", "jpeg"],
            accept_multiple_files=True
        )

        run = st.form_submit_button("Analyze + Add to Fleet Log")

    if run:
        use_report = report_file is not None
        use_photos = photos is not None and len(photos) > 0

        if not use_report and not use_photos:
            st.error("Upload a report file OR at least one photo.")
            st.stop()
        if use_report and use_photos:
            st.error("Please use only ONE option: report file OR photos.")
            st.stop()

        coating_dt = parse_date_safe(coating_applied)
        coating_season = season_from_month(coating_dt.month) if coating_dt else "Unknown"

        rust_px_total = 0
        valid_px_total = 0
        per_photo_rows = []
        photo_panels = []
        num_photos = 0

        if use_report:
            b = report_file.getvalue()
            if report_file.name.lower().endswith(".xlsx"):
                extracted = extract_images_from_excel(b)
            else:
                extracted = extract_images_from_docx(b)

            if not extracted:
                st.error("No embedded images found in the report.")
                st.stop()

            for i, (src, img_bytes) in enumerate(extracted, 1):
                try:
                    pil = Image.open(io.BytesIO(img_bytes)).convert("RGB")
                except Exception:
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
                num_photos += 1

                per_photo_rows.append({"Photo": i, "Source": src, "Rust %": round(float(pct), 2)})
                photo_panels.append({
                    "title": f"Photo {i} ({src})",
                    "orig": pil_to_png_bytes(make_thumb(pil)),
                    "overlay": pil_to_png_bytes(make_thumb(Image.fromarray(cv2.cvtColor(overlay, cv2.COLOR_BGR2RGB)))),
                    "mask": pil_to_png_bytes(make_thumb(Image.fromarray(mask).convert("RGB"))),
                })

            input_type = f"Report: {report_file.name}"

        else:
            photos_sorted = sorted(photos, key=lambda f: f.name.lower())
            for i, f in enumerate(photos_sorted, 1):
                pil = Image.open(f).convert("RGB")
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
                num_photos += 1

                per_photo_rows.append({"Photo": i, "Source": f.name, "Rust %": round(float(pct), 2)})
                photo_panels.append({
                    "title": f"Photo {i} ({f.name})",
                    "orig": pil_to_png_bytes(make_thumb(pil)),
                    "overlay": pil_to_png_bytes(make_thumb(Image.fromarray(cv2.cvtColor(overlay, cv2.COLOR_BGR2RGB)))),
                    "mask": pil_to_png_bytes(make_thumb(Image.fromarray(mask).convert("RGB"))),
                })

            input_type = f"Photos uploaded ({len(photos)} files)"

        if valid_px_total <= 0:
            st.error("Valid pixels total is 0 (image too dark/invalid).")
            st.stop()

        rust_pct_total = 100 * rust_px_total / max(valid_px_total, 1)
        severity = get_severity(rust_pct_total, minor_thr, moderate_thr)

        st.success(f"TOTAL Rust: {rust_pct_total:.2f}% | Severity: {severity}")

        report_meta = {
            "Vessel": vessel,
            "Tank / Hold": tank,
            "Location": location,
            "Inspection Date": str(insp_date),
            "Inspector": inspector,
            "Paint Manufacturer": paint_maker,
            "Yard": yard,
            "Coating Applied Date": coating_applied,
            "Coating Season": coating_season,
            "Coating Notes": coating_notes,
            "Input Type": input_type,
            "No. of Photos Analyzed": str(num_photos),
            "Remarks": remarks,
        }

        totals = {
            "rust_pct": float(rust_pct_total),
            "severity": severity,
            "rust_pixels": int(rust_px_total),
            "valid_pixels": int(valid_px_total),
        }

        pdf_bytes = generate_pdf_with_thumbnails(report_meta, totals, per_photo_rows, photo_panels)
        st.download_button(
            "📄 Download Inspection Report (PDF)",
            data=pdf_bytes,
            file_name=f"rust_report_{vessel.replace(' ', '_')}_{insp_date}.pdf",
            mime="application/pdf"
        )

        new_row = {
            "InspectionDate": str(insp_date),
            "Vessel": vessel,
            "Tank/Hold": tank,
            "Location": location,
            "Inspector": inspector,
            "PaintManufacturer": paint_maker,
            "Yard": yard,
            "CoatingAppliedDate": coating_applied,
            "CoatingSeason": coating_season,
            "CoatingNotes": coating_notes,
            "InputType": input_type,
            "NumPhotosAnalyzed": int(num_photos),
            "TotalRustPct": float(rust_pct_total),
            "Severity": severity,
            "RustPixels": int(rust_px_total),
            "ValidPixels": int(valid_px_total),
            "Remarks": remarks,
        }

        st.session_state.log_df = pd.concat([st.session_state.log_df, pd.DataFrame([new_row])], ignore_index=True)
        st.info("Added this inspection to Fleet Log (session). Download the Fleet Log in the Dashboard tab to keep history.")

# ============================================================
# TAB 2: Fleet Dashboard
# ============================================================
with tab2:
    st.subheader("Fleet Dashboard")

    st.write("### 1) Fleet Master (25 vessels)")
    st.dataframe(st.session_state.master_df, use_container_width=True)

    colA, colB = st.columns(2)
    with colA:
        st.download_button(
            "⬇️ Download Fleet Master (Excel)",
            data=to_excel_bytes(st.session_state.master_df, "FleetMaster"),
            file_name="fleet_master.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with colB:
        st.download_button(
            "⬇️ Download Fleet Master (CSV)",
            data=st.session_state.master_df.to_csv(index=False).encode("utf-8"),
            file_name="fleet_master.csv",
            mime="text/csv"
        )

    st.write("### 2) Upload Fleet Log (continue month-to-month)")
    up = st.file_uploader("Upload existing Fleet Log (CSV or Excel)", ["csv", "xlsx"])

    if up is not None:
        if up.name.lower().endswith(".csv"):
            df_up = pd.read_csv(up)
        else:
            df_up = pd.read_excel(up)
        st.session_state.log_df = df_up.copy()
        st.success("Fleet Log loaded.")

    log_df = st.session_state.log_df.copy()

    st.write("### 3) Fleet Log")
    st.dataframe(log_df.tail(200), use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "⬇️ Download Fleet Log (Excel)",
            data=to_excel_bytes(log_df, "FleetLog"),
            file_name="fleet_rust_log.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col2:
        st.download_button(
            "⬇️ Download Fleet Log (CSV)",
            data=log_df.to_csv(index=False).encode("utf-8"),
            file_name="fleet_rust_log.csv",
            mime="text/csv"
        )

    if log_df.empty:
        st.warning("No log data yet. Run at least one inspection in the first tab, then download the log.")
        st.stop()

    log_df["InspectionDate_dt"] = pd.to_datetime(log_df["InspectionDate"], errors="coerce")
    log_df["CoatingAppliedDate_dt"] = pd.to_datetime(log_df["CoatingAppliedDate"], errors="coerce")
    log_df = log_df.dropna(subset=["InspectionDate_dt"])

    log_df["AgeMonths"] = log_df.apply(
        lambda r: months_between(r["InspectionDate_dt"].date(), r["CoatingAppliedDate_dt"].date())
        if pd.notna(r["CoatingAppliedDate_dt"]) else np.nan,
        axis=1
    )
    log_df["Month"] = log_df["InspectionDate_dt"].dt.to_period("M").astype(str)

    st.divider()
    st.write("### 4) Filters")
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        vessel_f = st.selectbox("Vessel", ["(All)"] + sorted(log_df["Vessel"].dropna().unique().tolist()))
    with f2:
        maker_f = st.selectbox("Paint Manufacturer", ["(All)"] + sorted(log_df["PaintManufacturer"].dropna().unique().tolist()))
    with f3:
        yard_f = st.selectbox("Yard", ["(All)"] + sorted(log_df["Yard"].dropna().unique().tolist()))
    with f4:
        tank_f = st.selectbox("Tank/Hold", ["(All)"] + sorted(log_df["Tank/Hold"].dropna().unique().tolist()))

    df = log_df.copy()
    if vessel_f != "(All)":
        df = df[df["Vessel"] == vessel_f]
    if maker_f != "(All)":
        df = df[df["PaintManufacturer"] == maker_f]
    if yard_f != "(All)":
        df = df[df["Yard"] == yard_f]
    if tank_f != "(All)":
        df = df[df["Tank/Hold"] == tank_f]

    st.divider()
    st.write("### 5) Alerts")
    df_sorted = df.sort_values(["Vessel", "Tank/Hold", "InspectionDate_dt"])
    df_sorted["PrevRust"] = df_sorted.groupby(["Vessel", "Tank/Hold"])["TotalRustPct"].shift(1)
    df_sorted["IncreaseVsLast"] = df_sorted["TotalRustPct"] - df_sorted["PrevRust"]

    alerts = df_sorted[
        (df_sorted["TotalRustPct"] >= float(alert_rust_pct)) |
        (df_sorted["IncreaseVsLast"] >= float(alert_increase_pct))
    ].copy()

    if alerts.empty:
        st.success("No alerts triggered by current thresholds.")
    else:
        st.warning("Alerts triggered — review these items for proactive action.")
        st.dataframe(
            alerts[["InspectionDate", "Vessel", "Tank/Hold", "TotalRustPct", "PrevRust", "IncreaseVsLast",
                   "PaintManufacturer", "Yard", "AgeMonths"]],
            use_container_width=True
        )

    st.divider()
    st.write("### 6) Vessel Trend (Monthly)")
    trend = df.groupby(["Vessel", "Month"], as_index=False)["TotalRustPct"].mean()
    if vessel_f == "(All)":
        fleet_trend = df.groupby("Month", as_index=False)["TotalRustPct"].mean().sort_values("Month")
        st.line_chart(fleet_trend.set_index("Month")["TotalRustPct"])
        st.caption("Fleet average rust% trend by month (filtered).")
    else:
        vt = trend[trend["Vessel"] == vessel_f].sort_values("Month")
        if not vt.empty:
            st.line_chart(vt.set_index("Month")["TotalRustPct"])
            st.caption(f"{vessel_f} average rust% trend by month.")

    st.divider()
    st.write("### 7) Worst performers (latest per tank/hold)")
    latest = df.sort_values("InspectionDate_dt").groupby(["Vessel", "Tank/Hold"], as_index=False).tail(1)
    worst = latest.sort_values("TotalRustPct", ascending=False).head(10)
    st.dataframe(
        worst[["InspectionDate", "Vessel", "Tank/Hold", "TotalRustPct", "PaintManufacturer", "Yard", "AgeMonths"]],
        use_container_width=True
    )

    st.divider()
    st.write("### 8) Compare Performance: Paint Manufacturer & Yard")

    maker_perf = df.groupby("PaintManufacturer", as_index=False).agg(
        AvgRustPct=("TotalRustPct", "mean"),
        MedianRustPct=("TotalRustPct", "median"),
        Count=("TotalRustPct", "count"),
        AvgAgeMonths=("AgeMonths", "mean")
    ).sort_values("AvgRustPct", ascending=False)

    st.write("**Paint Manufacturer Performance (higher rust% = poorer)**")
    st.dataframe(maker_perf, use_container_width=True)
    if len(maker_perf) > 1:
        st.bar_chart(maker_perf.set_index("PaintManufacturer")["AvgRustPct"])

    yard_perf = df.groupby("Yard", as_index=False).agg(
        AvgRustPct=("TotalRustPct", "mean"),
        MedianRustPct=("TotalRustPct", "median"),
        Count=("TotalRustPct", "count"),
        AvgAgeMonths=("AgeMonths", "mean")
    ).sort_values("AvgRustPct", ascending=False)

    st.write("**Yard Performance (higher rust% = poorer)**")
    st.dataframe(yard_perf, use_container_width=True)
    if len(yard_perf) > 1:
        st.bar_chart(yard_perf.set_index("Yard")["AvgRustPct"])

    st.caption(
        "Tip: For fair comparison, keep coating age comparable. We capture AgeMonths for later normalization "
        "(next step: rate-of-rust per month within age bands)."
    )

