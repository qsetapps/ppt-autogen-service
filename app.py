import io
import os
import datetime as dt
import re

from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from openpyxl import load_workbook

app = Flask(__name__)

MONTHS_SV = [
    "Januari","Februari","Mars","April","Maj","Juni",
    "Juli","Augusti","September","Oktober","November","December"
]

def to_period_label(excel_value):
    if isinstance(excel_value, (dt.datetime, dt.date)):
        d = excel_value.date() if isinstance(excel_value, dt.datetime) else excel_value
        return f"{MONTHS_SV[d.month-1]} {d.year}"

    s = str(excel_value).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%y", "%m/%d/%y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            d = dt.datetime.strptime(s, fmt).date()
            return f"{MONTHS_SV[d.month-1]} {d.year}"
        except ValueError:
            pass
    return s  # fallback

def norm_lines(txt: str):
    return (txt or "").replace("\x0b", "\n").split("\n")

def find_shape_by_first_line(slide, first_line_text: str):
    target = first_line_text.strip()
    for shape in slide.shapes:
        if shape.has_text_frame:
            lines = norm_lines(shape.text)
            if len(lines) >= 1 and lines[0].strip() == target:
                return shape
    return None

def set_value_under_label(slide, label, value):
    shape = find_shape_by_first_line(slide, label)
    if not shape:
        raise ValueError(f"Hittade ingen ruta med label '{label}'")
    shape.text = f"{label}\n{value}"

def set_slide1_period(slide1, period_label: str):
    # I din mall finns texten "Månader År" på slide 1 (från din senaste version).
    # Vi byter ut den rutan till t.ex. "December 2025".
    for shape in slide1.shapes:
        if shape.has_text_frame:
            t = (shape.text or "").strip()
            if t in ("Månader År", "December 2025") or re.search(r"\b20\d{2}\b", t):
                shape.text = period_label
                return
    raise ValueError("Hittade ingen period-ruta på slide 1 att uppdatera.")

def update_ppt(excel_bytes: bytes, ppt_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    if "Auto" not in wb.sheetnames:
        raise ValueError("Saknar flik 'Auto' i Excel.")
    ws = wb["Auto"]

    period_label = to_period_label(ws["A1"].value)

    prs = Presentation(io.BytesIO(ppt_bytes))
    if len(prs.slides) < 2:
        raise ValueError("PPT måste ha minst 2 slides.")

    slide1 = prs.slides[0]
    slide2 = prs.slides[1]

    # Slide 1
    set_slide1_period(slide1, period_label)

    # Slide 2: läs Auto!B4:F6
    oms, tg1, tg2, tg3, ebita = [ws[c].value for c in ["B4","C4","D4","E4","F4"]]
    oms_g, tg1_g, tg2_g, tg3_g, ebita_g = [ws[c].value for c in ["B5","C5","D5","E5","F5"]]
    oms_gsek, tg1_gsek, tg2_gsek, tg3_gsek, ebita_gsek = [ws[c].value for c in ["B6","C6","D6","E6","F6"]]

    # Budget
    set_value_under_label(slide2, "Omsättning", oms)
    set_value_under_label(slide2, "TG 1", tg1)
    set_value_under_label(slide2, "TG 2", tg2)
    set_value_under_label(slide2, "TG 3", tg3)
    set_value_under_label(slide2, "EBITA", ebita)

    # Tillväxt %
    set_value_under_label(slide2, "Tillväxt Omsättning", oms_g)
    set_value_under_label(slide2, "Tillväxt TG 1", tg1_g)
    set_value_under_label(slide2, "Tillväxt TG 2", tg2_g)
    set_value_under_label(slide2, "Tillväxt TG 3", tg3_g)
    set_value_under_label(slide2, "Tillväxt EBITA", ebita_g)

    # Tillväxt SEK
    set_value_under_label(slide2, "Tillväxt Sek Omsättning", oms_gsek)
    set_value_under_label(slide2, "Tillväxt Sek TG 1", tg1_gsek)
    set_value_under_label(slide2, "Tillväxt Sek TG 2", tg2_gsek)
    set_value_under_label(slide2, "Tillväxt Sek TG 3", tg3_gsek)
    set_value_under_label(slide2, "Tillväxt Sek EBITA", ebita_gsek)

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()

def require_token():
    expected = os.environ.get("API_TOKEN")
    if not expected:
        return
    got = request.headers.get("X-API-Token", "")
    if got != expected:
        raise PermissionError("Unauthorized")

@app.route("/health", methods=["GET"])
def health():
    return {"ok": True}

@app.route("/update", methods=["POST"])
def update():
    try:
        require_token()

        if "excel" not in request.files or "ppt" not in request.files:
            return jsonify({"error": "Skicka multipart/form-data med 'excel' och 'ppt'."}), 400

        excel_bytes = request.files["excel"].read()
        ppt_bytes = request.files["ppt"].read()

        out_bytes = update_ppt(excel_bytes, ppt_bytes)

        return send_file(
            io.BytesIO(out_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="Ledningsmote_autogen.pptx",
        )

    except PermissionError as e:
        return jsonify({"error": str(e)}), 401
    except Exception as e:
        return jsonify({"error": str(e)}), 500
