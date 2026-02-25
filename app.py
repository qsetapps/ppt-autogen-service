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
    lines = norm_lines(shape.text)
    shape.text = f"{label}\n{value}"

def update_ppt(excel_bytes: bytes, ppt_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb["Auto"]

    period_label = to_period_label(ws["A1"].value)

    prs = Presentation(io.BytesIO(ppt_bytes))

    slide1 = prs.slides[0]
    slide2 = prs.slides[1]

    # --- Slide 1 ---
    for shape in slide1.shapes:
        if shape.has_text_frame and "År" in shape.text:
            shape.text = period_label
            break

    # --- Slide 2 Budget ---
    oms, tg1, tg2, tg3, ebita = [ws[c].value for c in ["B4","C4","D4","E4","F4"]]
    oms_g, tg1_g, tg2_g, tg3_g, ebita_g = [ws[c].value for c in ["B5","C5","D5","E5","F5"]]
    oms_gsek, tg1_gsek, tg2_gsek, tg3_gsek, ebita_gsek = [ws[c].value for c in ["B6","C6","D6","E6","F6"]]

    # Budget-raden
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
