"""PU7 paveikslu generatorius - DB1-DB6 analize"""
import os
from PIL import Image, ImageDraw, ImageFont
import csv

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_REGULAR = "/usr/share/fonts/google-noto-vf/NotoSans[wght].ttf"
FONT_ITALIC = "/usr/share/fonts/google-noto-vf/NotoSans-Italic[wght].ttf"


def font(size, bold=False, italic=False):
    path = FONT_ITALIC if italic else FONT_REGULAR
    f = ImageFont.truetype(path, size)
    if bold:
        try: f.set_variation_by_axes([700])
        except: pass
    return f


def text_size(d, t, fnt):
    bb = d.textbbox((0, 0), t, font=fnt)
    return bb[2] - bb[0], bb[3] - bb[1]


def draw_er():
    """ER diagrama: 3 lenteles - tblDepartments, tblInventory, tblSuppliers"""
    W, H = 1400, 700
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    title_font = font(16, bold=True)
    table_title_font = font(14, bold=True)
    field_font = font(12)
    pk_font = font(12, bold=True)
    card_font = font(14, bold=True)
    note_font = font(11)

    title = "Inventoriaus DB ER diagrama (DB1-DB6 bendra struktura)"
    tw, _ = text_size(d, title, title_font)
    d.text(((W - tw) // 2, 25), title, font=title_font, fill="#222")

    tables = [
        {"name": "tblDepartments", "x": 60, "y": 100, "w": 320, "h": 90,
         "fields": [("PK", "Department", "VARCHAR(50)")]},
        {"name": "tblInventory", "x": 540, "y": 100, "w": 320, "h": 410,
         "fields": [
             ("PK", "ProductCode", "VARCHAR(30)"),
             ("FK", "Dept", "VARCHAR(50)"),
             ("FK", "SupplierID", "VARCHAR(30)"),
             ("",   "ItemDescription", "VARCHAR(255)"),
             ("",   "Location", "VARCHAR(50)"),
             ("",   "Rack", "VARCHAR(20)"),
             ("",   "Origin", "VARCHAR(40)"),
             ("",   "UnitsInStock", "INTEGER"),
             ("",   "TargetInventory", "INTEGER"),
             ("",   "ReorderLevel", "INTEGER"),
             ("",   "LastOrdered", "DATE"),
             ("",   "OurUnitCost", "DECIMAL"),
             ("",   "RetailPrice", "DECIMAL"),
         ]},
        {"name": "tblSuppliers", "x": 1020, "y": 100, "w": 320, "h": 360,
         "fields": [
             ("PK", "SupplierID", "VARCHAR(30)"),
             ("",   "FirstName", "VARCHAR(50)"),
             ("",   "LastName", "VARCHAR(50)"),
             ("",   "ContactPhone", "VARCHAR(30)"),
             ("",   "Company", "VARCHAR(100)"),
             ("",   "ContactEmail", "VARCHAR(100)"),
             ("",   "Address", "VARCHAR(150)"),
             ("",   "City", "VARCHAR(60)"),
             ("",   "StateProvince", "VARCHAR(40)"),
             ("",   "Country", "VARCHAR(40)"),
             ("",   "PostalCode", "VARCHAR(20)"),
         ]},
    ]
    for tbl in tables:
        x, y, w, h = tbl["x"], tbl["y"], tbl["w"], tbl["h"]
        d.rectangle([x, y, x + w, y + 34], fill="#2C5282", outline="#1A365D", width=2)
        tw, _ = text_size(d, tbl["name"], table_title_font)
        d.text((x + (w - tw) // 2, y + 9), tbl["name"], font=table_title_font, fill="white")
        d.rectangle([x, y + 34, x + w, y + h], fill="white", outline="#1A365D", width=2)
        for i, (kt, name, dtype) in enumerate(tbl["fields"]):
            row_y = y + 42 + i * 26
            if kt == "PK" and i == 0 and len(tbl["fields"]) > 1:
                d.line([(x + 8, row_y + 22), (x + w - 8, row_y + 22)], fill="#888", width=1)
            if kt == "PK":
                d.text((x + 10, row_y + 4), "PK", font=pk_font, fill="#B7791F")
            elif kt == "FK":
                d.text((x + 10, row_y + 4), "FK", font=pk_font, fill="#3182CE")
            d.text((x + 42, row_y + 4), name,
                   font=(pk_font if kt == "PK" else field_font), fill="black")
            tw, _ = text_size(d, dtype, field_font)
            d.text((x + w - tw - 10, row_y + 4), dtype, font=field_font, fill="#666")

    # Rysiai
    d.line([(380, 145), (540, 130)], fill="#2D3748", width=2)
    d.text((385, 122), "1", font=card_font, fill="#2D3748")
    d.text((520, 110), "N", font=card_font, fill="#2D3748")
    d.line([(1020, 200), (860, 160)], fill="#2D3748", width=2)
    d.text((1000, 178), "1", font=card_font, fill="#2D3748")
    d.text((870, 138), "N", font=card_font, fill="#2D3748")

    d.text((40, H - 40),
           "PK - pirminis raktas, FK - isorinis raktas. tblDepartments turi tik 1 stulpeli (Department).",
           font=note_font, fill="#666")

    out = os.path.join(OUT_DIR, "img1_er_diagrama.png")
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


def draw_datasheet(filename, title, headers_widths, rows, app_color="#A02B2B"):
    total_w = sum(w for _, w in headers_widths)
    margin = 20; title_h = 40; header_h = 30; row_h = 28; autofield_h = 30
    img_w = total_w + 2 * margin
    img_h = margin + title_h + header_h + len(rows) * row_h + autofield_h + margin
    img = Image.new("RGB", (img_w, img_h), "white")
    d = ImageDraw.Draw(img)
    title_font = font(13, bold=True)
    header_font = font(12, bold=True)
    cell_font = font(11)

    d.text((margin + 5, margin + 10), title, font=title_font, fill=app_color)
    table_top = margin + title_h
    d.rectangle([margin, table_top, margin + total_w, table_top + header_h],
                fill="#E8E8E8", outline="#888")
    cx = margin
    for label, cw in headers_widths:
        d.text((cx + 6, table_top + 8), label, font=header_font, fill="black")
        d.line([(cx + cw, table_top), (cx + cw, table_top + header_h + len(rows) * row_h)],
               fill="#888", width=1)
        cx += cw

    for ri, row_data in enumerate(rows):
        ry = table_top + header_h + ri * row_h
        d.rectangle([margin, ry, margin + total_w, ry + row_h], fill="white", outline="#CCC")
        cx = margin
        for ci, val in enumerate(row_data):
            v = str(val)
            if len(v) > 35:
                v = v[:32] + "..."
            d.text((cx + 6, ry + 7), v, font=cell_font, fill="black")
            cx += headers_widths[ci][1]

    d.rectangle([margin, table_top, margin + total_w, table_top + header_h + len(rows) * row_h],
                outline="#888", width=1)
    autofield_y = table_top + header_h + len(rows) * row_h
    d.rectangle([margin, autofield_y, margin + total_w, autofield_y + autofield_h],
                fill="#FAFAFA", outline="#CCC")
    d.text((margin + 6, autofield_y + 7), "<AutoField>", font=cell_font, fill="#888")

    out = os.path.join(OUT_DIR, filename)
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


def read_csv(path):
    with open(path, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def draw_form_window(filename, title, subtitle, fields, buttons,
                     accent="#18A303", record_label="Record 1 of 51",
                     extra_note=None):
    """Universalus LibreOffice Base formos langas (mockup'as)."""
    margin = 18
    title_h = 30
    subtitle_h = 28
    field_label_w = 180
    field_input_w = 380
    field_h = 32
    field_gap = 10
    btn_h = 36
    btn_gap = 10
    nav_h = 30

    n = len(fields)
    inner_w = field_label_w + field_input_w + 40
    inner_h = (subtitle_h + 12 + n * (field_h + field_gap)
               + 18 + btn_h + 16 + nav_h)
    img_w = inner_w + 2 * margin
    img_h = title_h + inner_h + 2 * margin + 14

    img = Image.new("RGB", (img_w, img_h), "#F5F5F5")
    d = ImageDraw.Draw(img)

    # Lango antraste (titulinis)
    d.rectangle([0, 0, img_w, title_h], fill=accent, outline=accent)
    title_font = font(13, bold=True)
    d.text((12, 8), title, font=title_font, fill="white")
    # Mygtukai desineje (close, max, min)
    for i, ch in enumerate(["_", "[]", "X"]):
        bx = img_w - 30 * (3 - i) - 6
        d.rectangle([bx, 6, bx + 22, title_h - 6], outline="white", width=1)
        d.text((bx + 6, 8), ch, font=font(11, bold=True), fill="white")

    # Vidinis fonas
    body_top = title_h + margin
    d.rectangle([margin, body_top, img_w - margin, img_h - margin],
                fill="white", outline="#888", width=1)

    # Subtitulas (formos pavadinimas viduje)
    d.rectangle([margin + 1, body_top + 1,
                 img_w - margin - 1, body_top + subtitle_h],
                fill="#E8F5E9", outline="#888")
    sub_font = font(12, bold=True)
    d.text((margin + 12, body_top + 7), subtitle, font=sub_font, fill="#1B5E20")

    # Laukai
    label_font = font(11, bold=True)
    input_font = font(11)
    fy = body_top + subtitle_h + 14
    for label, value in fields:
        # Label
        d.text((margin + 14, fy + 8), label, font=label_font, fill="#222")
        # Input box
        ix = margin + 14 + field_label_w
        d.rectangle([ix, fy, ix + field_input_w, fy + field_h],
                    fill="white", outline="#9CA3AF", width=1)
        d.text((ix + 8, fy + 8), str(value), font=input_font, fill="#1F2937")
        fy += field_h + field_gap

    # Mygtukai apacioje
    btn_font = font(10, bold=True)
    btn_y = fy + 14
    btn_x = margin + 14
    for label in buttons:
        # Calculate button width based on text
        bw, _ = text_size(d, label, btn_font)
        bw = max(bw + 24, 86)
        d.rectangle([btn_x, btn_y, btn_x + bw, btn_y + btn_h],
                    fill="#E5E7EB", outline="#6B7280", width=1)
        d.text((btn_x + (bw - text_size(d, label, btn_font)[0]) // 2,
                btn_y + 11), label, font=btn_font, fill="#111827")
        btn_x += bw + btn_gap

    # Navigacijos juosta
    nav_y = btn_y + btn_h + 14
    d.rectangle([margin + 1, nav_y, img_w - margin - 1, nav_y + nav_h],
                fill="#F3F4F6", outline="#9CA3AF", width=1)
    nav_font = font(10)
    # Strelytes (|<  <  >  >|)
    arrows = "|<   <   >   >|   +"
    d.text((margin + 12, nav_y + 8), arrows, font=nav_font, fill="#374151")
    rec_w, _ = text_size(d, record_label, nav_font)
    d.text((img_w - margin - rec_w - 14, nav_y + 8),
           record_label, font=nav_font, fill="#374151")

    # Pastaba apacioje (jeigu yra)
    if extra_note:
        d.text((margin + 4, img_h - margin + 2),
               extra_note, font=font(9, italic=True), fill="#666")

    out = os.path.join(OUT_DIR, filename)
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


def draw_report(filename, title, group_label, groups, columns,
                summary_label=None, summary_value=None,
                accent="#9F7AEA", subtitle=None, footer=None):
    """Universalus LibreOffice Base ataskaitos vaizdas su grupavimu.

    columns: list of (label, width) tuples
    groups: list of dicts with 'name', 'rows' (list of tuples matching columns)
    """
    margin = 18
    title_h = 38
    subtitle_h = 24 if subtitle else 0
    table_header_h = 28
    row_h = 24
    group_header_h = 28
    group_gap = 8
    summary_h = 30 if summary_label else 0
    footer_h = 22 if footer else 0

    total_w = sum(w for _, w in columns)
    n_rows = sum(len(g["rows"]) for g in groups)
    n_groups = len(groups)
    inner_h = (subtitle_h + table_header_h
               + n_groups * (group_header_h + group_gap)
               + n_rows * row_h
               + summary_h + footer_h + 24)

    img_w = total_w + 2 * margin + 20
    img_h = title_h + inner_h + 2 * margin + 30

    img = Image.new("RGB", (img_w, img_h), "white")
    d = ImageDraw.Draw(img)

    # Antraste
    d.rectangle([margin, margin, img_w - margin, margin + title_h],
                fill=accent, outline=accent)
    title_font = font(15, bold=True)
    tw, _ = text_size(d, title, title_font)
    d.text(((img_w - tw) // 2, margin + 9),
           title, font=title_font, fill="white")

    cur_y = margin + title_h + 8

    # Subtitulas (jei yra)
    if subtitle:
        sub_font = font(11, italic=True)
        sw, _ = text_size(d, subtitle, sub_font)
        d.text(((img_w - sw) // 2, cur_y + 4),
               subtitle, font=sub_font, fill="#444")
        cur_y += subtitle_h

    # Group label virsuje (auksciau lenteles)
    if group_label:
        gl_font = font(11, bold=True, italic=True)
        d.text((margin + 4, cur_y),
               f"Grupavimas: {group_label}", font=gl_font, fill="#444")
        cur_y += 18

    # Stulpeliu antrastes
    table_x = margin + 10
    d.rectangle([table_x, cur_y, table_x + total_w, cur_y + table_header_h],
                fill="#E5E7EB", outline="#888", width=1)
    cx = table_x
    h_font = font(11, bold=True)
    for label, cw in columns:
        d.text((cx + 8, cur_y + 7), label, font=h_font, fill="#111827")
        d.line([(cx + cw, cur_y), (cx + cw, cur_y + table_header_h)],
               fill="#888", width=1)
        cx += cw
    cur_y += table_header_h

    # Grupes
    cell_font = font(11)
    grp_font = font(11, bold=True)
    for g in groups:
        # Group header
        d.rectangle([table_x, cur_y, table_x + total_w, cur_y + group_header_h],
                    fill="#FAF5FF", outline="#888", width=1)
        d.text((table_x + 8, cur_y + 6),
               f"  {g['name']}", font=grp_font, fill="#581C87")
        if "summary" in g:
            sw, _ = text_size(d, g["summary"], grp_font)
            d.text((table_x + total_w - sw - 14, cur_y + 6),
                   g["summary"], font=grp_font, fill="#581C87")
        cur_y += group_header_h
        # Rows
        for row in g["rows"]:
            d.rectangle([table_x, cur_y, table_x + total_w, cur_y + row_h],
                        fill="white", outline="#D1D5DB", width=1)
            cx = table_x
            for ci, val in enumerate(row):
                v = str(val)
                if len(v) > 38:
                    v = v[:35] + "..."
                d.text((cx + 8, cur_y + 5), v, font=cell_font, fill="#111827")
                cx += columns[ci][1]
            cur_y += row_h
        cur_y += group_gap

    # Suvestine apacioje
    if summary_label:
        d.rectangle([table_x, cur_y, table_x + total_w, cur_y + summary_h],
                    fill="#F3E8FF", outline="#581C87", width=2)
        d.text((table_x + 10, cur_y + 7),
               summary_label, font=grp_font, fill="#581C87")
        if summary_value:
            sw, _ = text_size(d, summary_value, grp_font)
            d.text((table_x + total_w - sw - 14, cur_y + 7),
                   summary_value, font=grp_font, fill="#581C87")
        cur_y += summary_h

    # Footer
    if footer:
        d.text((margin + 4, img_h - margin - 12),
               footer, font=font(9, italic=True), fill="#666")

    out = os.path.join(OUT_DIR, filename)
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


def draw_db3_forms(inv):
    """img10, img11 - DB3 formu mockup'ai."""
    sample = inv[0] if inv else {
        "ProductCode": "hware-1", "Dept": "Hardware",
        "SupplierID": "PUGG", "ItemDescription": "Hammer 16oz",
        "LastOrdered": "2015-03-15", "Origin": "USA",
        "OurUnitCost": "5.20", "RetailPrice": "12.99",
    }
    # img10: pirma DB3 forma - 4 laukai
    draw_form_window(
        filename="img10_form_inventory.png",
        title="frmInventory - DB3.odb [LibreOffice Base]",
        subtitle="Inventoriaus duomenu ivedimo forma",
        fields=[
            ("Produkto kodas:", sample["ProductCode"]),
            ("Skyrius:", sample["Dept"]),
            ("Tiekejas:", sample["SupplierID"]),
            ("Produkto aprasymas:", sample["ItemDescription"][:40]),
        ],
        buttons=["Issaugoti", "Atsaukti", "Naujas irasas", "Salinti"],
        accent="#18A303",
        record_label="Irasas 1 is 51",
        extra_note="Forma sukurta per Form Wizard, stulpelinis (Columnar) isdestymas. Data: =Now()",
    )

    # img11: antra DB3 forma - 7 laukai
    sample2 = inv[5] if len(inv) > 5 else sample
    draw_form_window(
        filename="img11_form_inventory_extended.png",
        title="frmInventoryExtended - DB3.odb [LibreOffice Base]",
        subtitle="Uzsakymu perziuros forma (Notebook stilius)",
        fields=[
            ("Produkto kodas:", sample2["ProductCode"]),
            ("Skyrius:", sample2["Dept"]),
            ("Tiekejas:", sample2["SupplierID"]),
            ("Paskutinis uzsakymas:", sample2["LastOrdered"] or "2015-05-21"),
            ("Kilmes salis:", sample2["Origin"]),
            ("Verte (savikaina):", "$" + str(sample2["OurUnitCost"]).replace("$", "")),
            ("Mazmenine kaina:", "$" + str(sample2["RetailPrice"]).replace("$", "")),
        ],
        buttons=["Issaugoti", "Atsaukti", "Naujas irasas", "Salinti"],
        accent="#1565C0",
        record_label="Irasas 6 is 51",
        extra_note="Office 2016 tema, Calibri 14pt antrastes, 11pt laukai. Visi 7 verslo svarbiausi laukai.",
    )


def draw_db4_form_suppliers(sups):
    """img12 - DB4 frmSuppliers forma."""
    s = sups[0] if sups else {
        "SupplierID": "PUGG", "FirstName": "John", "LastName": "Smith",
        "Company": "PUGG Hardware Inc.", "ContactPhone": "+1-555-0100",
        "ContactEmail": "john.smith@pugg.com",
        "Address": "1500 Industrial Blvd",
        "City": "Chicago", "Country": "USA", "PostalCode": "60601",
    }
    # Speciali "dvieju stulpeliu" forma - identifikacija kaireje, adresas desineje
    margin = 18
    title_h = 30
    subtitle_h = 28
    col_w = 380  # vieno stulpelio plotis
    field_label_w = 130
    field_input_w = col_w - field_label_w - 12
    field_h = 30
    field_gap = 10
    btn_h = 34
    btn_gap = 10
    nav_h = 28
    section_gap = 22

    left_fields = [
        ("Tiekejo ID:", s["SupplierID"]),
        ("Vardas:", s["FirstName"]),
        ("Pavarde:", s["LastName"]),
        ("Imone:", s.get("Company", "")[:32]),
        ("Telefonas:", s.get("ContactPhone", "")),
        ("El. pastas:", s.get("ContactEmail", "")[:32]),
    ]
    right_fields = [
        ("Adresas:", s.get("Address", "")[:30]),
        ("Miestas:", s.get("City", "")),
        ("Salis:", s.get("Country", "")),
        ("Pasto kodas:", s.get("PostalCode", "")),
    ]
    n_max = max(len(left_fields), len(right_fields))

    inner_w = col_w * 2 + section_gap + 28
    inner_h = (subtitle_h + 14 + 22 + n_max * (field_h + field_gap)
               + 18 + btn_h + 14 + nav_h)
    img_w = inner_w + 2 * margin
    img_h = title_h + inner_h + 2 * margin + 16

    img = Image.new("RGB", (img_w, img_h), "#F5F5F5")
    d = ImageDraw.Draw(img)

    # Antraste
    accent = "#7E57C2"
    d.rectangle([0, 0, img_w, title_h], fill=accent)
    d.text((12, 8), "frmSuppliers - DB4.odb [LibreOffice Base]",
           font=font(13, bold=True), fill="white")
    for i, ch in enumerate(["_", "[]", "X"]):
        bx = img_w - 30 * (3 - i) - 6
        d.rectangle([bx, 6, bx + 22, title_h - 6], outline="white", width=1)
        d.text((bx + 6, 8), ch, font=font(11, bold=True), fill="white")

    body_top = title_h + margin
    d.rectangle([margin, body_top, img_w - margin, img_h - margin],
                fill="white", outline="#888", width=1)

    # Subtitulas
    d.rectangle([margin + 1, body_top + 1,
                 img_w - margin - 1, body_top + subtitle_h],
                fill="#EDE7F6", outline="#888")
    d.text((margin + 12, body_top + 7), "Tiekejo informacijos forma",
           font=font(12, bold=True), fill="#311B92")

    section_y = body_top + subtitle_h + 12
    sec_font = font(11, bold=True, italic=True)
    d.text((margin + 14, section_y), "[ Identifikacija ]",
           font=sec_font, fill="#311B92")
    d.text((margin + 14 + col_w + section_gap, section_y),
           "[ Adresas ]", font=sec_font, fill="#311B92")

    label_font = font(11, bold=True)
    input_font = font(11)

    def draw_col(fields, x_off):
        fy = section_y + 22
        for label, value in fields:
            d.text((margin + 14 + x_off, fy + 7),
                   label, font=label_font, fill="#222")
            ix = margin + 14 + x_off + field_label_w
            d.rectangle([ix, fy, ix + field_input_w, fy + field_h],
                        fill="white", outline="#9CA3AF", width=1)
            d.text((ix + 8, fy + 7), str(value),
                   font=input_font, fill="#1F2937")
            fy += field_h + field_gap

    draw_col(left_fields, 0)
    draw_col(right_fields, col_w + section_gap)

    # Mygtukai
    btn_font = font(10, bold=True)
    btn_y = section_y + 22 + n_max * (field_h + field_gap) + 14
    btn_x = margin + 14
    for label in ["Issaugoti", "Atsaukti", "Naujas tiekejas", "Salinti"]:
        bw, _ = text_size(d, label, btn_font)
        bw = max(bw + 24, 100)
        d.rectangle([btn_x, btn_y, btn_x + bw, btn_y + btn_h],
                    fill="#E5E7EB", outline="#6B7280", width=1)
        d.text((btn_x + (bw - text_size(d, label, btn_font)[0]) // 2,
                btn_y + 10), label, font=btn_font, fill="#111827")
        btn_x += bw + btn_gap

    # Navigacijos juosta
    nav_y = btn_y + btn_h + 12
    d.rectangle([margin + 1, nav_y, img_w - margin - 1, nav_y + nav_h],
                fill="#F3F4F6", outline="#9CA3AF", width=1)
    nav_font = font(10)
    d.text((margin + 12, nav_y + 7),
           "|<   <   >   >|   +", font=nav_font, fill="#374151")
    rec_label = "Tiekejas 1 is 18"
    rw, _ = text_size(d, rec_label, nav_font)
    d.text((img_w - margin - rw - 14, nav_y + 7),
           rec_label, font=nav_font, fill="#374151")

    out = os.path.join(OUT_DIR, "img12_form_suppliers.png")
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


def draw_db5_reports(sups, inv):
    """img13, img14 - DB5 ataskaitos."""
    # img13: tiekeju kontaktine ataskaita - grupuota pagal Country, City
    from collections import defaultdict
    by_country = defaultdict(lambda: defaultdict(list))
    for s in sups:
        by_country[s["Country"]][s["City"]].append(s)

    groups = []
    for country in sorted(by_country.keys()):
        rows = []
        for city in sorted(by_country[country].keys()):
            for s in by_country[country][city]:
                rows.append((
                    city,
                    s["Company"][:24],
                    s["LastName"],
                    s["FirstName"],
                    s["ContactPhone"][:14],
                ))
        cnt = sum(len(v) for v in by_country[country].values())
        groups.append({
            "name": country,
            "summary": f"({cnt} tiekejai)",
            "rows": rows,
        })

    draw_report(
        filename="img13_rpt_suppliers_country.png",
        title="Tiekeju kontaktine ataskaita",
        subtitle="rptSuppliersByCountry - DB5.odb",
        group_label="Country, paskui City (ORDER BY)",
        groups=groups,
        columns=[("CITY", 130), ("COMPANY", 230), ("LASTNAME", 130),
                 ("FIRSTNAME", 130), ("PHONE", 140)],
        summary_label=f"Is viso tiekeju:",
        summary_value=str(len(sups)),
        accent="#9F7AEA",
        footer="Saltinis: tblSuppliers (DB1/DB5.odb), 18 irasu",
    )

    # img14: inventoriaus pagal kilme ataskaita
    by_origin = defaultdict(list)
    for r in inv:
        by_origin[r["Origin"]].append(r)
    groups = []
    for origin in sorted(by_origin.keys(), key=lambda x: -len(by_origin[x])):
        rows = []
        for r in by_origin[origin][:6]:  # limit to 6 per group for visibility
            rows.append((
                r["SupplierID"],
                r["ProductCode"],
                r["ItemDescription"][:38],
            ))
        groups.append({
            "name": origin,
            "summary": f"({len(by_origin[origin])} prekes)",
            "rows": rows,
        })

    draw_report(
        filename="img14_rpt_origin_inventory.png",
        title="Inventoriaus pagal kilmes sali ataskaita",
        subtitle="rptInventoryByOrigin - DB5.odb",
        group_label="Origin, paskui SupplierID",
        groups=groups,
        columns=[("SUPPLIERID", 150), ("PRODUCTCODE", 160), ("ITEMDESCRIPTION", 360)],
        summary_label="Is viso inventoriaus prekiu:",
        summary_value=str(len(inv)),
        accent="#9F7AEA",
        footer="Saltinis: tblInventory (DB5.odb), rodoma po 6 vienetus is grupes",
    )


def draw_db6_reports(inv):
    """img15, img16, img17 - DB6 ataskaitos su SUM/AVG/MAX."""
    from collections import defaultdict

    # img15: SUM ataskaita - inventoriaus verte pagal tiekeja
    by_sup = defaultdict(lambda: {"items": [], "total": 0.0, "stock": 0})
    for r in inv:
        rp = float(str(r["RetailPrice"]).replace("$", "").replace(",", "")) if r["RetailPrice"] else 0
        stock = int(r["UnitsInStock"]) if r["UnitsInStock"] else 0
        by_sup[r["SupplierID"]]["items"].append(r)
        by_sup[r["SupplierID"]]["total"] += stock * rp
        by_sup[r["SupplierID"]]["stock"] += stock

    groups = []
    grand_total = 0.0
    sup_sorted = sorted(by_sup.items(), key=lambda x: -x[1]["total"])
    for sup_id, info in sup_sorted[:8]:
        rows = []
        for r in info["items"][:3]:  # show 3 examples per supplier
            stock = int(r["UnitsInStock"]) if r["UnitsInStock"] else 0
            rp = float(str(r["RetailPrice"]).replace("$", "").replace(",", "")) if r["RetailPrice"] else 0
            line_val = stock * rp
            rows.append((
                r["ProductCode"],
                r["ItemDescription"][:30],
                str(stock),
                f"${rp:,.2f}",
                f"${line_val:,.2f}",
            ))
        grand_total += info["total"]
        groups.append({
            "name": sup_id,
            "summary": f"SUM = ${info['total']:,.2f}",
            "rows": rows,
        })

    # Add remaining suppliers as a summary line
    rest_total = sum(info["total"] for sup_id, info in sup_sorted[8:])
    if rest_total > 0:
        grand_total += rest_total
    full_total = sum(info["total"] for info in by_sup.values())

    draw_report(
        filename="img15_rpt_sum_supplier.png",
        title="Inventoriaus vertes SUM ataskaita pagal tiekejus",
        subtitle="rptSumBySupplier - DB6.odb",
        group_label="SupplierID, su SUM(UnitsInStock * RetailPrice)",
        groups=groups,
        columns=[("PRODUCTCODE", 120), ("ITEMDESCRIPTION", 280),
                 ("STOCK", 90), ("PRICE", 110), ("LINE_VALUE", 130)],
        summary_label="GRAND TOTAL (visi tiekejai):",
        summary_value=f"${full_total:,.2f}",
        accent="#7C3AED",
        footer="Saltinis: tblInventory (DB6.odb), top 8 tiekejai pagal verte",
    )

    # img16: AVG ataskaita - vidutine kaina pagal skyriu
    by_dept = defaultdict(lambda: {"items": [], "prices": []})
    for r in inv:
        rp = float(str(r["RetailPrice"]).replace("$", "").replace(",", "")) if r["RetailPrice"] else 0
        by_dept[r["Dept"]]["items"].append(r)
        by_dept[r["Dept"]]["prices"].append(rp)

    groups = []
    overall_prices = []
    dept_sorted = sorted(by_dept.items(),
                         key=lambda x: -sum(x[1]["prices"]) / max(len(x[1]["prices"]), 1))
    for dept, info in dept_sorted[:8]:
        avg = sum(info["prices"]) / len(info["prices"]) if info["prices"] else 0
        rows = []
        for r in info["items"][:3]:
            rp = float(str(r["RetailPrice"]).replace("$", "").replace(",", "")) if r["RetailPrice"] else 0
            rows.append((
                r["ProductCode"],
                r["ItemDescription"][:32],
                f"${rp:,.2f}",
            ))
        overall_prices.extend(info["prices"])
        groups.append({
            "name": dept,
            "summary": f"AVG = ${avg:,.2f}  (n={len(info['prices'])})",
            "rows": rows,
        })

    overall_avg = sum(overall_prices) / len(overall_prices) if overall_prices else 0

    # AVG ataskaita su geltonu/oranziniu akcentu
    draw_report(
        filename="img16_rpt_avg_dept.png",
        title="Vidutines kainos AVG ataskaita pagal skyrius",
        subtitle="rptAvgByDept - DB6.odb",
        group_label="Dept, su AVG(RetailPrice)",
        groups=groups,
        columns=[("PRODUCTCODE", 130), ("ITEMDESCRIPTION", 360),
                 ("PRICE", 140)],
        summary_label="VISO VIDURKIS (visi skyriai):",
        summary_value=f"${overall_avg:,.2f}",
        accent="#F59E0B",
        footer="Saltinis: tblInventory (DB6.odb), top 8 skyriai pagal AVG",
    )

    # img17: MAX ataskaita - brangiausias produktas pagal kilme
    by_origin = defaultdict(list)
    for r in inv:
        by_origin[r["Origin"]].append(r)

    groups = []
    overall_max = {"price": 0.0, "code": "", "desc": ""}
    for origin in sorted(by_origin.keys()):
        items = by_origin[origin]
        max_item = max(items, key=lambda r: float(str(r["RetailPrice"]).replace("$","").replace(",","")) if r["RetailPrice"] else 0)
        max_price = float(str(max_item["RetailPrice"]).replace("$","").replace(",","")) if max_item["RetailPrice"] else 0
        if max_price > overall_max["price"]:
            overall_max = {"price": max_price,
                           "code": max_item["ProductCode"],
                           "desc": max_item["ItemDescription"][:30]}
        groups.append({
            "name": origin,
            "summary": f"MAX = ${max_price:,.2f}",
            "rows": [(
                max_item["ProductCode"],
                max_item["SupplierID"],
                max_item["ItemDescription"][:36],
                f"${max_price:,.2f}",
            )],
        })

    draw_report(
        filename="img17_rpt_max_origin.png",
        title="Brangiausio produkto MAX ataskaita pagal kilme",
        subtitle="rptMaxByOrigin - DB6.odb",
        group_label="Origin, su MAX(RetailPrice)",
        groups=groups,
        columns=[("PRODUCTCODE", 130), ("SUPPLIERID", 130),
                 ("ITEMDESCRIPTION", 330), ("MAX_PRICE", 130)],
        summary_label=f"BRANGIAUSIAS APSKRITAI: {overall_max['code']} - {overall_max['desc']}",
        summary_value=f"${overall_max['price']:,.2f}",
        accent="#DC2626",
        footer="Saltinis: tblInventory (DB6.odb), po 1 brangiausia is kiekvienos kilmes",
    )


def main():
    draw_er()

    # img2: tblDepartments turinys
    deps = read_csv("extracted/DB4_tblDepartments.csv")
    rows = [(r["Department"],) for r in deps[:18]]
    draw_datasheet("img2_data_tblDepartments.png",
                   "tblDepartments - DB1.odb (18 irasu)",
                   [("DEPARTMENT", 380)], rows, app_color="#18A303")

    # img3: tblSuppliers turinys (subset)
    sups = read_csv("extracted/DB4_tblSuppliers.csv")
    rows = [(r["SupplierID"], r["FirstName"] + " " + r["LastName"], r["Company"], r["City"], r["Country"]) for r in sups[:18]]
    draw_datasheet("img3_data_tblSuppliers.png",
                   "tblSuppliers - DB1.odb (18 tiekeju)",
                   [("SUPPLIERID", 130), ("VARDAS PAVARDE", 200), ("COMPANY", 320),
                    ("CITY", 130), ("COUNTRY", 100)], rows, app_color="#18A303")

    # img4: tblInventory turinys (pirma 15 irasu)
    inv = read_csv("extracted/DB4_tblInventory.csv")
    rows = [(r["ProductCode"], r["Dept"], r["SupplierID"], r["ItemDescription"][:30],
             r["UnitsInStock"], r["RetailPrice"], r["Origin"]) for r in inv[:15]]
    draw_datasheet("img4_data_tblInventory.png",
                   "tblInventory - DB1.odb (51 irasai, rodoma 15)",
                   [("PRODUCTCODE", 110), ("DEPT", 100), ("SUPPLIERID", 130),
                    ("ITEMDESCRIPTION", 260), ("STOCK", 70), ("PRICE", 90), ("ORIGIN", 90)],
                   rows, app_color="#18A303")

    # img5: Department uzklausos rezultatas (Hardware)
    hw = [r for r in inv if r["Dept"] == "Hardware"]
    rows = [(r["ProductCode"], r["SupplierID"], r["ItemDescription"][:35]) for r in hw]
    draw_datasheet("img5_qry_department.png",
                   "qryDepartment uzklausa: Hardware skyriaus produktai (rezultatas)",
                   [("PRODUCTCODE", 130), ("SUPPLIERID", 140), ("ITEMDESCRIPTION", 380)],
                   rows, app_color="#3182CE")

    # img6: Supplier uzklausos rezultatas (Bathroom)
    br = [r for r in inv if r["Dept"] == "Bathroom"]
    rows = [(r["ProductCode"], r["SupplierID"], r["ItemDescription"][:35]) for r in br]
    draw_datasheet("img6_qry_supplier.png",
                   "qrySupplier uzklausa: Bathroom skyriaus produktai (rezultatas)",
                   [("PRODUCTCODE", 130), ("SUPPLIERID", 140), ("ITEMDESCRIPTION", 380)],
                   rows, app_color="#3182CE")

    # img7: qryUnusedSuppliers rezultatas (tiekejai is Kanados)
    sups = read_csv("extracted/DB4_tblSuppliers.csv")
    canada = [r for r in sups if r["Country"] == "Canada"]
    rows = [(r["SupplierID"], r["Company"], r["City"], r["Country"]) for r in canada]
    draw_datasheet("img7_qry_unused.png",
                   "qryUnusedSuppliers (modifikuota): tiekejai is Kanados",
                   [("SUPPLIERID", 150), ("COMPANY", 320), ("CITY", 130), ("COUNTRY", 110)],
                   rows, app_color="#3182CE")

    # img8: rptDepartments ataskaitos prasme - lentele su skyriaus produktu suvestine
    dept_counts = {}
    for r in inv:
        d = r["Dept"]
        dept_counts[d] = dept_counts.get(d, 0) + 1
    rows = [(d, str(c)) for d, c in sorted(dept_counts.items())]
    draw_datasheet("img8_rpt_departments.png",
                   "rptDepartments ataskaita: produktu kiekis pagal skyrius",
                   [("DEPARTMENT", 250), ("PRODUCT_COUNT", 200)], rows, app_color="#9F7AEA")

    # img9: Skaiciavimu pavyzdys (DB6) - SUM/AVG/MAX pagal Origin
    from collections import defaultdict
    ag = defaultdict(lambda: {"count": 0, "stock": 0, "value": 0.0, "max_p": 0.0})
    for r in inv:
        o = r["Origin"]
        ag[o]["count"] += 1
        ag[o]["stock"] += int(r["UnitsInStock"])
        rp = float(r["RetailPrice"].replace("$","").replace(",","")) if r["RetailPrice"] else 0
        ag[o]["value"] += rp
        if rp > ag[o]["max_p"]:
            ag[o]["max_p"] = rp
    rows = []
    for o, v in sorted(ag.items()):
        avg = v["value"] / v["count"] if v["count"] else 0
        rows.append((o, str(v["count"]), str(v["stock"]),
                     f"{v['value']:.2f}", f"{avg:.2f}", f"{v['max_p']:.2f}"))
    draw_datasheet("img9_db6_calc.png",
                   "DB6 skaiciavimu pavyzdys: SUM, AVG, MAX, COUNT pagal Origin",
                   [("ORIGIN", 110), ("COUNT", 80), ("SUM_STOCK", 110),
                    ("SUM_PRICE", 130), ("AVG_PRICE", 130), ("MAX_PRICE", 130)],
                   rows, app_color="#9F7AEA")

    # === Nauji paveikslai DB3 (formos), DB4 (frmSuppliers),
    # DB5 (2 ataskaitos), DB6 (3 ataskaitos su SUM/AVG/MAX) ===
    draw_db3_forms(inv)
    draw_db4_form_suppliers(sups)
    draw_db5_reports(sups, inv)
    draw_db6_reports(inv)

    print("Visi paveikslai sukurti.")


if __name__ == "__main__":
    main()
