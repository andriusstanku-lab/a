"""
PU5 paveikslu generatorius

Sukuria 6 PNG paveikslus:
  img1_er_eprekyba.png   - ER diagrama: e-prekybos sistema
  img2_er_ligonine.png   - ER diagrama: ligonines valdymo sistema
  img3_er_biblioteka.png - ER diagrama: bibliotekos valdymo sistema
  img4_data_eprekyba.png - LibreOffice Base datasheet: Uzsakymai lentele
  img5_data_ligonine.png - LibreOffice Base datasheet: Vizitai lentele
  img6_data_biblioteka.png - LibreOffice Base datasheet: Skolinimasi lentele
"""

import os
import sys
from PIL import Image, ImageDraw, ImageFont

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_REGULAR = "/usr/share/fonts/google-noto-vf/NotoSans[wght].ttf"


def font(size, bold=False):
    f = ImageFont.truetype(FONT_REGULAR, size)
    if bold:
        try:
            f.set_variation_by_axes([700])
        except Exception:
            pass
    return f


def text_size(draw, text, fnt):
    bbox = draw.textbbox((0, 0), text, font=fnt)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


# ============================================================
# ER diagramos generatorius
# ============================================================

def create_er_diagram(filename, title, tables):
    """Sukuria ER diagrama su 3 lentelemis ir rysiais."""
    W, H = 1400, 700
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    title_font = font(16, bold=True)
    table_title_font = font(14, bold=True)
    field_font = font(13)
    pk_font = font(13, bold=True)
    note_font = font(11)
    cardinality_font = font(14, bold=True)

    tw, _ = text_size(d, title, title_font)
    d.text(((W - tw) // 2, 30), title, font=title_font, fill="#222222")

    for tbl in tables:
        x, y, w, h = tbl["x"], tbl["y"], tbl["w"], tbl["h"]
        d.rectangle([x, y, x + w, y + 40],
                    fill="#2C5282", outline="#1A365D", width=2)
        tw, _ = text_size(d, tbl["name"], table_title_font)
        d.text((x + (w - tw) // 2, y + 11),
               tbl["name"], font=table_title_font, fill="white")
        d.rectangle([x, y + 40, x + w, y + h],
                    fill="white", outline="#1A365D", width=2)
        for i, (key_type, name, dtype) in enumerate(tbl["fields"]):
            row_y = y + 50 + i * 30
            if key_type == "PK" and i == 0 and len(tbl["fields"]) > 1:
                d.line([(x + 10, row_y + 26),
                        (x + w - 10, row_y + 26)],
                       fill="#888888", width=1)
            if key_type == "PK":
                d.text((x + 12, row_y + 4), "PK",
                       font=pk_font, fill="#B7791F")
            elif key_type == "FK":
                d.text((x + 12, row_y + 4), "FK",
                       font=pk_font, fill="#3182CE")
            d.text((x + 50, row_y + 4), name,
                   font=(pk_font if key_type == "PK" else field_font),
                   fill="black")
            tw, _ = text_size(d, dtype, field_font)
            d.text((x + w - tw - 12, row_y + 4),
                   dtype, font=field_font, fill="#666666")

    s_x = tables[0]["x"] + tables[0]["w"]
    s_y = tables[0]["y"] + 80
    e_x = tables[1]["x"]
    e_y = tables[1]["y"] + 110
    d.line([(s_x, s_y), (e_x, e_y)], fill="#2D3748", width=2)
    d.text((s_x + 10, s_y - 22), "1",
           font=cardinality_font, fill="#2D3748")
    d.text((e_x - 22, e_y - 22), "N",
           font=cardinality_font, fill="#2D3748")

    c_x = tables[2]["x"]
    c_y = tables[2]["y"] + 80
    e2_x = tables[1]["x"] + tables[1]["w"]
    e2_y = tables[1]["y"] + 140
    d.line([(c_x, c_y), (e2_x, e2_y)], fill="#2D3748", width=2)
    d.text((c_x - 22, c_y - 22), "1",
           font=cardinality_font, fill="#2D3748")
    d.text((e2_x + 10, e2_y - 22), "N",
           font=cardinality_font, fill="#2D3748")

    legend_y = 480
    d.rectangle([60, legend_y, W - 60, legend_y + 130],
                fill="#F7FAFC", outline="#CBD5E0", width=1)
    d.text((80, legend_y + 15),
           "Paaiskinimai:", font=table_title_font, fill="#2D3748")
    legend_lines = [
        ("PK", "#B7791F", "Pirminis raktas"),
        ("FK", "#3182CE", "Isorinis raktas"),
        ("1:N", "#2D3748", "Vienas-prie-daug rysys"),
    ]
    for i, (label, color, text) in enumerate(legend_lines):
        ly = legend_y + 45 + i * 25
        d.text((80, ly), label, font=pk_font, fill=color)
        d.text((150, ly), text, font=note_font, fill="#2D3748")

    out = os.path.join(OUT_DIR, filename)
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


def create_er_eprekyba():
    tables = [
        {"name": "KLIENTAI", "x": 60, "y": 110, "w": 320, "h": 180,
         "fields": [
             ("PK", "KlientoID", "INTEGER"),
             ("",   "Vardas", "VARCHAR(100)"),
             ("",   "ElPastas", "VARCHAR(100)"),
         ]},
        {"name": "UZSAKYMAI", "x": 540, "y": 110, "w": 320, "h": 270,
         "fields": [
             ("PK", "UzsakymoID", "INTEGER"),
             ("FK", "KlientoID", "INTEGER"),
             ("FK", "ProduktoID", "INTEGER"),
             ("",   "UzsakymoData", "DATE"),
             ("",   "Kiekis", "INTEGER"),
             ("",   "PristatymoAdresas", "VARCHAR(150)"),
         ]},
        {"name": "PRODUKTAI", "x": 1020, "y": 110, "w": 320, "h": 210,
         "fields": [
             ("PK", "ProduktoID", "INTEGER"),
             ("",   "Pavadinimas", "VARCHAR(100)"),
             ("",   "Kaina", "DECIMAL(8,2)"),
             ("",   "Kategorija", "VARCHAR(50)"),
         ]},
    ]
    create_er_diagram("img1_er_eprekyba.png",
                      "E-prekybos sistemos ER diagrama", tables)


def create_er_ligonine():
    tables = [
        {"name": "PACIENTAI", "x": 60, "y": 110, "w": 320, "h": 150,
         "fields": [
             ("PK", "PacientoID", "INTEGER"),
             ("",   "PacientoVardas", "VARCHAR(100)"),
         ]},
        {"name": "VIZITAI", "x": 540, "y": 110, "w": 320, "h": 270,
         "fields": [
             ("PK", "VizitoID", "INTEGER"),
             ("FK", "PacientoID", "INTEGER"),
             ("FK", "GydytojoID", "INTEGER"),
             ("",   "VizitoData", "DATE"),
             ("",   "Diagnoze", "VARCHAR(150)"),
             ("",   "GydymoPastabos", "VARCHAR(255)"),
         ]},
        {"name": "GYDYTOJAI", "x": 1020, "y": 110, "w": 320, "h": 180,
         "fields": [
             ("PK", "GydytojoID", "INTEGER"),
             ("",   "GydytojoVardas", "VARCHAR(100)"),
             ("",   "Specializacija", "VARCHAR(80)"),
         ]},
    ]
    create_er_diagram("img2_er_ligonine.png",
                      "Ligonines valdymo sistemos ER diagrama", tables)


def create_er_biblioteka():
    tables = [
        {"name": "KNYGOS", "x": 60, "y": 110, "w": 320, "h": 210,
         "fields": [
             ("PK", "KnygosID", "INTEGER"),
             ("",   "Pavadinimas", "VARCHAR(150)"),
             ("",   "Autorius", "VARCHAR(100)"),
             ("",   "ISBN", "VARCHAR(20)"),
         ]},
        {"name": "SKOLINIMASI", "x": 540, "y": 110, "w": 320, "h": 240,
         "fields": [
             ("PK", "SkolinimoID", "INTEGER"),
             ("FK", "KnygosID", "INTEGER"),
             ("FK", "NarioID", "INTEGER"),
             ("",   "SkolinimosiData", "DATE"),
             ("",   "GrazinimoData", "DATE"),
         ]},
        {"name": "NARIAI", "x": 1020, "y": 110, "w": 320, "h": 180,
         "fields": [
             ("PK", "NarioID", "INTEGER"),
             ("",   "NarioVardas", "VARCHAR(100)"),
             ("",   "NarystesData", "DATE"),
         ]},
    ]
    create_er_diagram("img3_er_biblioteka.png",
                      "Bibliotekos valdymo sistemos ER diagrama", tables)


def create_datasheet(filename, headers_widths, rows, title=None):
    total_w = sum(w for _, w in headers_widths)
    n_rows = len(rows)
    header_h = 30
    row_h = 30
    autofield_h = 30
    margin = 20
    title_h = 40 if title else 0

    img_w = total_w + 2 * margin
    img_h = margin + title_h + header_h + n_rows * row_h + autofield_h + margin

    img = Image.new("RGB", (img_w, img_h), "white")
    d = ImageDraw.Draw(img)

    title_font = font(13, bold=True)
    header_font = font(13, bold=True)
    cell_font = font(13)

    if title:
        d.text((margin + 5, margin + 10), title,
               font=title_font, fill="#18A303")

    table_top = margin + title_h
    d.rectangle([margin, table_top, margin + total_w, table_top + header_h],
                fill="#E8E8E8", outline="#888888")
    cx = margin
    for label, cw in headers_widths:
        d.text((cx + 8, table_top + 7), label, font=header_font, fill="black")
        d.line([(cx + cw, table_top),
                (cx + cw, table_top + header_h + n_rows * row_h)],
               fill="#888888", width=1)
        cx += cw

    for ri, row_data in enumerate(rows):
        ry = table_top + header_h + ri * row_h
        d.rectangle([margin, ry, margin + total_w, ry + row_h],
                    fill="white", outline="#CCCCCC")
        cx = margin
        for ci, val in enumerate(row_data):
            d.text((cx + 8, ry + 7), str(val), font=cell_font, fill="black")
            cx += headers_widths[ci][1]

    d.rectangle([margin, table_top, margin + total_w,
                 table_top + header_h + n_rows * row_h],
                outline="#888888", width=1)

    autofield_y = table_top + header_h + n_rows * row_h
    d.rectangle([margin, autofield_y, margin + total_w,
                 autofield_y + autofield_h],
                fill="#FAFAFA", outline="#CCCCCC")
    d.text((margin + 8, autofield_y + 7),
           "<AutoField>", font=cell_font, fill="#888888")

    out = os.path.join(OUT_DIR, filename)
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


def create_data_eprekyba():
    headers = [
        ("UZSAKYMOID", 110),
        ("KLIENTOID", 100),
        ("PRODUKTOID", 110),
        ("UZSAKYMODATA", 130),
        ("KIEKIS", 80),
        ("PRISTATYMOADRESAS", 280),
    ]
    rows = [
        ("1", "1", "1", "2025-01-15", "2", "Vilnius g. 1, Vilnius"),
        ("2", "2", "2", "2025-01-20", "1", "K. Donelaicio g. 5, Kaunas"),
        ("3", "1", "3", "2025-02-01", "5", "Vilnius g. 1, Vilnius"),
        ("4", "3", "4", "2025-02-10", "2", "Taikos pr. 10, Klaipeda"),
        ("5", "4", "5", "2025-02-15", "1", "Tilzes g. 8, Siauliai"),
    ]
    create_datasheet("img4_data_eprekyba.png", headers, rows,
                     title="UZSAKYMAI - eprekyba.odb: Table Data View")


def create_data_ligonine():
    headers = [
        ("VIZITOID", 100),
        ("PACIENTOID", 110),
        ("GYDYTOJOID", 110),
        ("VIZITODATA", 120),
        ("DIAGNOZE", 200),
        ("GYDYMOPASTABOS", 260),
    ]
    rows = [
        ("1", "1", "1", "2025-01-10", "Hipertenzija", "Skirtas gydymas vaistais"),
        ("2", "2", "3", "2025-01-12", "Apendicitas",  "Skubi operacija"),
        ("3", "3", "2", "2025-01-15", "Sloga",        "Vaistai 7 dienoms"),
        ("4", "4", "5", "2025-02-01", "Sanario nudegimas", "Konsultacija"),
        ("5", "5", "4", "2025-02-05", "Migrenos priepuolis", "Vaistai"),
    ]
    create_datasheet("img5_data_ligonine.png", headers, rows,
                     title="VIZITAI - ligonine.odb: Table Data View")


def create_data_biblioteka():
    headers = [
        ("SKOLINIMOID", 120),
        ("KNYGOSID", 100),
        ("NARIOID", 100),
        ("SKOLINIMOSIDATA", 160),
        ("GRAZINIMODATA", 150),
    ]
    rows = [
        ("1", "1", "1", "2025-03-01", "2025-03-15"),
        ("2", "2", "2", "2025-03-05", "2025-03-20"),
        ("3", "3", "1", "2025-03-10", "2025-03-24"),
        ("4", "4", "4", "2025-03-15", "2025-03-29"),
        ("5", "5", "5", "2025-03-20", "2025-04-03"),
    ]
    create_datasheet("img6_data_biblioteka.png", headers, rows,
                     title="SKOLINIMASI - biblioteka.odb: Table Data View")


def main():
    print("Generuojami PU5 paveikslai...")
    create_er_eprekyba()
    create_er_ligonine()
    create_er_biblioteka()
    create_data_eprekyba()
    create_data_ligonine()
    create_data_biblioteka()
    print("Baigta. 6 paveikslai sugeneruoti.")


if __name__ == "__main__":
    main()
