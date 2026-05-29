"""PU7 paveikslu generatorius - DB1-DB6 analize"""
import os
from PIL import Image, ImageDraw, ImageFont
import csv

OUT_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_REGULAR = "/usr/share/fonts/google-noto-vf/NotoSans[wght].ttf"


def font(size, bold=False):
    f = ImageFont.truetype(FONT_REGULAR, size)
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

    print("Visi paveikslai sukurti.")


if __name__ == "__main__":
    main()
