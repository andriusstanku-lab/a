"""
PU4 paveikslu generatorius

Sukuria 5 PNG paveikslus, naudojamus PU4 darbe:
  img1_er_diagrama.png       - ER diagrama (Students, Courses, Enrollments)
  img2_access_design.png     - MS Access Design View (Students lentele)
  img3_access_data.png       - MS Access Datasheet View su duomenimis
  img4_libre_design.png      - LibreOffice Base Table Design (Students)
  img5_libre_data.png        - LibreOffice Base Datasheet View su duomenimis

Naudojamos bibliotekos: Pillow (PIL).
Paveikslai yra realisticni mockup'ai, naudojami akademiniam darbui kaip
iliustracijos. Nera tikros ekrano nuotraukos.
"""

from PIL import Image, ImageDraw, ImageFont
import os

OUT_DIR = os.path.dirname(os.path.abspath(__file__))

FONT_REGULAR = "/usr/share/fonts/google-noto-vf/NotoSans[wght].ttf"


def font(size, bold=False):
    """Grazina TrueType srifta nurodyto dydzio."""
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
# 1 paveikslas: ER diagrama
# ============================================================

def create_er_diagram():
    W, H = 1200, 600
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    title_font = font(20, bold=True)
    table_title_font = font(15, bold=True)
    field_font = font(13)
    pk_font = font(13, bold=True)
    note_font = font(11)

    tables = [
        {
            "name": "Students",
            "x": 80, "y": 80, "w": 280, "h": 220,
            "fields": [
                ("PK", "StudentID", "INTEGER"),
                ("",   "FirstName", "VARCHAR(50)"),
                ("",   "LastName",  "VARCHAR(50)"),
                ("",   "Email",     "VARCHAR(100)"),
                ("",   "Major",     "VARCHAR(50)"),
                ("",   "EnrollmentYear", "INTEGER"),
            ],
        },
        {
            "name": "Enrollments",
            "x": 460, "y": 80, "w": 280, "h": 220,
            "fields": [
                ("PK", "EnrollmentID", "INTEGER"),
                ("FK", "StudentID",    "INTEGER"),
                ("FK", "CourseID",     "INTEGER"),
                ("",   "Semester",     "VARCHAR(20)"),
                ("",   "Grade",        "DECIMAL(3,1)"),
            ],
        },
        {
            "name": "Courses",
            "x": 840, "y": 80, "w": 280, "h": 220,
            "fields": [
                ("PK", "CourseID",   "INTEGER"),
                ("",   "CourseName", "VARCHAR(100)"),
                ("",   "Credits",    "INTEGER"),
                ("",   "Department", "VARCHAR(50)"),
            ],
        },
    ]

    # Lenteles
    for tbl in tables:
        x, y, w, h = tbl["x"], tbl["y"], tbl["w"], tbl["h"]
        d.rectangle([x, y, x + w, y + 36], fill="#4a6fa5", outline="black")
        tw, _ = text_size(d, tbl["name"], table_title_font)
        d.text((x + (w - tw) // 2, y + 8),
               tbl["name"], font=table_title_font, fill="white")
        d.rectangle([x, y + 36, x + w, y + h], fill="white", outline="black")
        for i, (key_type, name, dtype) in enumerate(tbl["fields"]):
            row_y = y + 44 + i * 28
            if key_type == "PK":
                d.text((x + 10, row_y), "[PK]", font=pk_font, fill="#b8860b")
                fld_font = pk_font
            elif key_type == "FK":
                d.text((x + 10, row_y), "[FK]", font=pk_font, fill="#5a5aff")
                fld_font = field_font
            else:
                fld_font = field_font
            d.text((x + 60, row_y), name, font=fld_font, fill="black")
            tw, _ = text_size(d, dtype, field_font)
            d.text((x + w - tw - 10, row_y),
                   dtype, font=field_font, fill="#555555")

    # Rysiai
    s_x, s_y = tables[0]["x"] + tables[0]["w"], tables[0]["y"] + 80
    e_x, e_y = tables[1]["x"], tables[1]["y"] + 105
    d.line([(s_x, s_y), (e_x, e_y)], fill="black", width=2)
    d.text((s_x + 5, s_y - 18), "1", font=table_title_font, fill="black")
    d.text((e_x - 18, e_y - 18), "N", font=table_title_font, fill="black")

    c_x, c_y = tables[2]["x"], tables[2]["y"] + 80
    e2_x, e2_y = tables[1]["x"] + tables[1]["w"], tables[1]["y"] + 130
    d.line([(c_x, c_y), (e2_x, e2_y)], fill="black", width=2)
    d.text((c_x - 18, c_y - 18), "1", font=table_title_font, fill="black")
    d.text((e2_x + 5, e2_y - 18), "N", font=table_title_font, fill="black")

    # Legenda
    legend_y = 380
    d.text((80, legend_y), "Legenda:", font=table_title_font, fill="black")
    d.text((80, legend_y + 30),
           "[PK] - pirminis raktas (Primary Key)",
           font=note_font, fill="black")
    d.text((80, legend_y + 50),
           "[FK] - isorinis raktas (Foreign Key)",
           font=note_font, fill="black")
    d.text((80, legend_y + 70),
           "1:N - vienas-prie-daug rysys",
           font=note_font, fill="black")
    d.text((80, legend_y + 90),
           "Enrollments yra siejanti lentele tarp Students ir Courses",
           font=note_font, fill="black")

    out = os.path.join(OUT_DIR, "img1_er_diagrama.png")
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


# ============================================================
# 2 paveikslas: MS Access Design View
# ============================================================

def create_access_design():
    W, H = 1200, 700
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    title_font = font(13, bold=True)
    ribbon_font = font(12)
    grid_header_font = font(13, bold=True)
    grid_font = font(13)
    prop_font = font(12)
    prop_label_font = font(12, bold=True)

    # Lango antraste (Access raudonas)
    d.rectangle([0, 0, W, 32], fill="#A02B2B", outline=None)
    d.text((10, 8),
           "Students - Table   |   Database1 - Microsoft Access",
           font=title_font, fill="white")

    # Juostele
    d.rectangle([0, 32, W, 92], fill="#F0F0F0", outline="#CCCCCC")
    tabs = ["File", "Home", "Create", "External Data", "Database Tools",
            "Table Design"]
    x = 12
    for t in tabs:
        is_active = (t == "Table Design")
        if is_active:
            d.rectangle([x - 4, 40, x + 100, 92],
                        fill="white", outline="#CCCCCC")
        d.text((x, 50), t, font=ribbon_font,
               fill="black" if not is_active else "#A02B2B")
        x += 110

    # Lauku grid
    grid_top = 110
    d.rectangle([20, grid_top, W - 20, grid_top + 32],
                fill="#E8E8E8", outline="#888888")
    d.text((50, grid_top + 8), "Field Name",
           font=grid_header_font, fill="black")
    d.text((360, grid_top + 8), "Data Type",
           font=grid_header_font, fill="black")
    d.text((620, grid_top + 8), "Description (Optional)",
           font=grid_header_font, fill="black")

    fields = [
        ("StudentID",      "AutoNumber", "Pirminis raktas, automatinis numeris"),
        ("FirstName",      "Short Text", "Studento vardas (privalomas)"),
        ("LastName",       "Short Text", "Studento pavarde (privaloma)"),
        ("Email",          "Short Text", "El. pasto adresas (unikalus)"),
        ("Major",          "Short Text", "Studiju programa"),
        ("EnrollmentYear", "Number",     "Imatrikuliacijos metai"),
    ]
    for i, (name, dtype, desc) in enumerate(fields):
        row_y = grid_top + 32 + i * 30
        if i == 0:
            d.text((10, row_y + 6), ">", font=grid_font, fill="black")
            d.ellipse([28, row_y + 9, 38, row_y + 19],
                      outline="#B8860B", fill="#FFD700")
        d.text((50, row_y + 6), name, font=grid_font, fill="black")
        d.text((360, row_y + 6), dtype, font=grid_font, fill="black")
        d.text((620, row_y + 6), desc, font=grid_font, fill="#444444")
        d.line([(20, row_y + 30), (W - 20, row_y + 30)],
               fill="#DDDDDD", width=1)

    # Field Properties panele
    prop_top = grid_top + 32 + 6 * 30 + 30
    d.rectangle([20, prop_top, W - 20, H - 30],
                fill="#F8F8F8", outline="#888888")
    d.text((30, prop_top + 8),
           "Field Properties (StudentID)",
           font=grid_header_font, fill="black")

    properties = [
        ("Field Size",   "Long Integer"),
        ("New Values",   "Increment"),
        ("Format",       ""),
        ("Caption",      "Studento ID"),
        ("Indexed",      "Yes (No Duplicates)"),
        ("Required",     "Yes"),
    ]
    for i, (label, value) in enumerate(properties):
        prop_y = prop_top + 38 + i * 22
        d.text((40, prop_y), label, font=prop_label_font, fill="black")
        d.text((220, prop_y), value, font=prop_font, fill="#222222")

    # Statuso juosta
    d.rectangle([0, H - 24, W, H], fill="#A02B2B", outline=None)
    d.text((10, H - 18),
           "Design view.   F6 = Switch panes.   F1 = Help.",
           font=ribbon_font, fill="white")

    out = os.path.join(OUT_DIR, "img2_access_design.png")
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


# ============================================================
# 3 paveikslas: MS Access Datasheet su duomenimis
# ============================================================

def create_access_datasheet():
    W, H = 1200, 700
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    title_font = font(13, bold=True)
    ribbon_font = font(12)
    header_font = font(12, bold=True)
    cell_font = font(12)

    # Lango antraste
    d.rectangle([0, 0, W, 32], fill="#A02B2B", outline=None)
    d.text((10, 8),
           "Students - Table (Datasheet View)   |   Database1 - Microsoft Access",
           font=title_font, fill="white")

    # Juostele
    d.rectangle([0, 32, W, 92], fill="#F0F0F0", outline="#CCCCCC")
    d.text((12, 52),
           "Datasheet View - Studentu lentele su 6 irasais",
           font=ribbon_font, fill="black")

    # Stulpeliu antrastes
    headers = [
        ("StudentID", 70),
        ("FirstName", 110),
        ("LastName", 130),
        ("Email", 250),
        ("Major", 220),
        ("EnrollmentYear", 140),
    ]
    grid_top = 110
    d.rectangle([30, grid_top, W - 30, grid_top + 32],
                fill="#D0D0E0", outline="#888888")
    cx = 30
    for label, cw in headers:
        d.text((cx + 8, grid_top + 8), label, font=header_font, fill="black")
        d.line([(cx + cw, grid_top), (cx + cw, grid_top + 32 + 6 * 32)],
               fill="#888888", width=1)
        cx += cw

    # Duomenys
    rows = [
        ("1", "Andrius",  "Vargonas",      "andrius.vargonas@knf.vu.lt", "Marketingo technologijos", "2024"),
        ("2", "Egle",     "Kazlauskaite",  "egle.k@knf.vu.lt",           "Verslo informatika",       "2023"),
        ("3", "Tomas",    "Petrauskas",    "tomas.p@knf.vu.lt",          "Marketingo technologijos", "2024"),
        ("4", "Ruta",     "Jonaityte",     "ruta.j@knf.vu.lt",           "Finansu valdymas",         "2025"),
        ("5", "Mantas",   "Bagdonas",      "mantas.b@knf.vu.lt",         "Verslo informatika",       "2023"),
        ("6", "Lina",     "Sakalauskaite", "lina.s@knf.vu.lt",           "Marketingo technologijos", "2025"),
    ]
    for ri, row in enumerate(rows):
        ry = grid_top + 32 + ri * 32
        bg = "#FFFFFF" if ri % 2 == 0 else "#F5F5F5"
        d.rectangle([30, ry, W - 30, ry + 32], fill=bg, outline="#DDDDDD")
        cx = 30
        for ci, val in enumerate(row):
            d.text((cx + 8, ry + 8), val, font=cell_font, fill="black")
            cx += headers[ci][1]

    # Navigacijos juosta
    nav_y = grid_top + 32 + 6 * 32 + 12
    d.rectangle([30, nav_y, W - 30, nav_y + 28],
                fill="#F0F0F0", outline="#888888")
    d.text((40, nav_y + 6),
           "Record:  |<  <   1 of 6   >  >|     Search: __________",
           font=cell_font, fill="black")

    # Statuso juosta
    d.rectangle([0, H - 24, W, H], fill="#A02B2B", outline=None)
    d.text((10, H - 18),
           "Datasheet view. 6 records.",
           font=ribbon_font, fill="white")

    out = os.path.join(OUT_DIR, "img3_access_data.png")
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


# ============================================================
# 4 paveikslas: LibreOffice Base Table Design
# ============================================================

def create_libre_design():
    W, H = 1200, 700
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    title_font = font(13, bold=True)
    menu_font = font(12)
    grid_header_font = font(13, bold=True)
    grid_font = font(13)
    prop_font = font(12)
    prop_label_font = font(12, bold=True)

    # Lango antraste (LibreOffice melsva)
    d.rectangle([0, 0, W, 32], fill="#18A303", outline=None)
    d.text((10, 8),
           "Students - LibreOffice Base: Table Design",
           font=title_font, fill="white")

    # Meniu juosta
    d.rectangle([0, 32, W, 60], fill="#F4F4F4", outline="#CCCCCC")
    menus = ["File", "Edit", "View", "Insert", "Tools", "Window", "Help"]
    x = 12
    for m in menus:
        d.text((x, 38), m, font=menu_font, fill="black")
        x += 60

    # Toolbar
    d.rectangle([0, 60, W, 92], fill="#EFEFEF", outline="#CCCCCC")
    d.text((12, 68),
           "[Save] [Copy] [Paste] [Primary Key] [Index Design] [Edit Foreign Key]",
           font=menu_font, fill="black")

    # Lauku grid
    grid_top = 110
    d.rectangle([20, grid_top, W - 20, grid_top + 32],
                fill="#DCEDDC", outline="#888888")
    d.text((50, grid_top + 8), "Field Name",
           font=grid_header_font, fill="black")
    d.text((360, grid_top + 8), "Field Type",
           font=grid_header_font, fill="black")
    d.text((620, grid_top + 8), "Description",
           font=grid_header_font, fill="black")

    fields = [
        ("StudentID",      "Integer [INTEGER]",         "Pirminis raktas, auto reiksmes"),
        ("FirstName",      "Text [VARCHAR]",            "Studento vardas (privalomas)"),
        ("LastName",       "Text [VARCHAR]",            "Studento pavarde (privaloma)"),
        ("Email",          "Text [VARCHAR]",            "El. pasto adresas (unikalus)"),
        ("Major",          "Text [VARCHAR]",            "Studiju programa"),
        ("EnrollmentYear", "Small Integer [SMALLINT]",  "Imatrikuliacijos metai"),
    ]
    for i, (name, dtype, desc) in enumerate(fields):
        row_y = grid_top + 32 + i * 30
        if i == 0:
            # Raktelio simbolis
            d.ellipse([28, row_y + 9, 38, row_y + 19],
                      outline="#B8860B", fill="#FFD700")
        d.text((50, row_y + 6), name, font=grid_font, fill="black")
        d.text((360, row_y + 6), dtype, font=grid_font, fill="black")
        d.text((620, row_y + 6), desc, font=grid_font, fill="#444444")
        d.line([(20, row_y + 30), (W - 20, row_y + 30)],
               fill="#DDDDDD", width=1)

    # Field Properties panele
    prop_top = grid_top + 32 + 6 * 30 + 30
    d.rectangle([20, prop_top, W - 20, H - 30],
                fill="#F4F8F4", outline="#888888")
    d.text((30, prop_top + 8),
           "Field Properties - StudentID",
           font=grid_header_font, fill="black")

    properties = [
        ("AutoValue",        "Yes"),
        ("Length",           "10"),
        ("Default value",    "(none)"),
        ("Format example",   "1, 2, 3, ..."),
        ("Entry required",   "Yes"),
    ]
    for i, (label, value) in enumerate(properties):
        prop_y = prop_top + 38 + i * 22
        d.text((40, prop_y), label, font=prop_label_font, fill="black")
        d.text((220, prop_y), value, font=prop_font, fill="#222222")

    # Statuso juosta
    d.rectangle([0, H - 24, W, H], fill="#18A303", outline=None)
    d.text((10, H - 18),
           "Table Design.   Press F1 for help.",
           font=menu_font, fill="white")

    out = os.path.join(OUT_DIR, "img4_libre_design.png")
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


# ============================================================
# 5 paveikslas: LibreOffice Base Datasheet
# ============================================================

def create_libre_datasheet():
    W, H = 1200, 700
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)

    title_font = font(13, bold=True)
    menu_font = font(12)
    header_font = font(12, bold=True)
    cell_font = font(12)

    # Lango antraste
    d.rectangle([0, 0, W, 32], fill="#18A303", outline=None)
    d.text((10, 8),
           "Courses - LibreOffice Base: Table Data View",
           font=title_font, fill="white")

    # Meniu
    d.rectangle([0, 32, W, 60], fill="#F4F4F4", outline="#CCCCCC")
    menus = ["File", "Edit", "View", "Insert", "Tools", "Window", "Help"]
    x = 12
    for m in menus:
        d.text((x, 38), m, font=menu_font, fill="black")
        x += 60

    # Toolbar
    d.rectangle([0, 60, W, 92], fill="#EFEFEF", outline="#CCCCCC")
    d.text((12, 68),
           "[Save] [Refresh] [Sort] [Filter] [Find] [New Record]",
           font=menu_font, fill="black")

    # Stulpeliu antrastes - Courses lentele
    headers = [
        ("CourseID", 100),
        ("CourseName", 380),
        ("Credits", 100),
        ("Department", 320),
    ]
    grid_top = 110
    total_w = sum(c[1] for c in headers)
    d.rectangle([30, grid_top, 30 + total_w, grid_top + 32],
                fill="#DCEDDC", outline="#888888")
    cx = 30
    for label, cw in headers:
        d.text((cx + 8, grid_top + 8), label, font=header_font, fill="black")
        d.line([(cx + cw, grid_top), (cx + cw, grid_top + 32 + 6 * 32)],
               fill="#888888", width=1)
        cx += cw

    # Duomenys - kursai
    rows = [
        ("101", "Skaitmeninio marketingo pagrindai",    "6", "Marketingo katedra"),
        ("102", "Duomenu bazes ir informacijos sistemos","6", "Informatikos katedra"),
        ("103", "Vartotoju elgsenos analize",           "5", "Marketingo katedra"),
        ("104", "Verslo statistika",                    "6", "Vadybos katedra"),
        ("105", "Programavimo pagrindai",               "5", "Informatikos katedra"),
        ("106", "Marketingo strategija",                "6", "Marketingo katedra"),
    ]
    for ri, row in enumerate(rows):
        ry = grid_top + 32 + ri * 32
        bg = "#FFFFFF" if ri % 2 == 0 else "#F0F8F0"
        d.rectangle([30, ry, 30 + total_w, ry + 32],
                    fill=bg, outline="#DDDDDD")
        cx = 30
        for ci, val in enumerate(row):
            d.text((cx + 8, ry + 8), val, font=cell_font, fill="black")
            cx += headers[ci][1]

    # Navigacijos juosta
    nav_y = grid_top + 32 + 6 * 32 + 12
    d.rectangle([30, nav_y, 30 + total_w, nav_y + 28],
                fill="#F0F0F0", outline="#888888")
    d.text((40, nav_y + 6),
           "Record  1 of 6     |<   <   >   >|   New Record",
           font=cell_font, fill="black")

    # Statuso juosta
    d.rectangle([0, H - 24, W, H], fill="#18A303", outline=None)
    d.text((10, H - 18),
           "Table Data View. 6 records.",
           font=menu_font, fill="white")

    out = os.path.join(OUT_DIR, "img5_libre_data.png")
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


# ============================================================
# Pagrindine funkcija
# ============================================================

if __name__ == "__main__":
    create_er_diagram()
    create_access_design()
    create_access_datasheet()
    create_libre_design()
    create_libre_datasheet()
    print("\nVisi paveikslai sugeneruoti.")
