"""
PU6 paveikslu generatorius - duomenu normalizavimas

9 paveikslai:
  img1-img2: Studentu registracija (UNF + 3NF)
  img3-img4: Pardavimai (UNF + 3NF)
  img5-img6: Sandeliai (UNF + 3NF)
  img7-img9: 3 datasheet view nuotraukos
"""
import os
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


def draw_unnormalized(filename, title, headers_widths, rows, dups=None):
    total_w = sum(w for _, w in headers_widths)
    margin = 30
    title_h = 40
    header_h = 36
    row_h = 32
    note_h = 50
    img_w = total_w + 2 * margin
    img_h = margin + title_h + header_h + len(rows) * row_h + note_h + margin
    img = Image.new("RGB", (img_w, img_h), "white")
    d = ImageDraw.Draw(img)
    title_font = font(14, bold=True)
    header_font = font(11, bold=True)
    cell_font = font(11)
    note_font = font(10)
    d.text((margin, margin), title, font=title_font, fill="#A02B2B")
    table_top = margin + title_h
    d.rectangle([margin, table_top, margin + total_w, table_top + header_h],
                fill="#FED7D7", outline="#A02B2B", width=2)
    cx = margin
    for label, cw in headers_widths:
        d.text((cx + 6, table_top + 10), label,
               font=header_font, fill="#742A2A")
        d.line([(cx + cw, table_top),
                (cx + cw, table_top + header_h + len(rows) * row_h)],
               fill="#A02B2B", width=1)
        cx += cw
    for ri, row_data in enumerate(rows):
        ry = table_top + header_h + ri * row_h
        d.rectangle([margin, ry, margin + total_w, ry + row_h],
                    fill="white", outline="#FED7D7")
        cx = margin
        for ci, val in enumerate(row_data):
            is_dup = (dups and (ri, ci) in dups)
            color = "#C53030" if is_dup else "black"
            d.text((cx + 6, ry + 8), str(val), font=cell_font, fill=color)
            cx += headers_widths[ci][1]
    d.rectangle([margin, table_top, margin + total_w,
                 table_top + header_h + len(rows) * row_h],
                outline="#A02B2B", width=2)
    note_y = table_top + header_h + len(rows) * row_h + 12
    d.text((margin, note_y),
           "Raudoni dublikatai - duomenu pertekliaus anomalija "
           "(reikia normalizavimo)", font=note_font, fill="#A02B2B")
    out = os.path.join(OUT_DIR, filename)
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


def draw_3nf(filename, title, tables, relationships):
    W, H = 1400, 720
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    title_font = font(15, bold=True)
    table_title_font = font(13, bold=True)
    field_font = font(12)
    pk_font = font(12, bold=True)
    card_font = font(13, bold=True)
    tw, _ = text_size(d, title, title_font)
    d.text(((W - tw) // 2, 20), title, font=title_font, fill="#2C5282")
    for tbl in tables:
        x, y, w, h = tbl["x"], tbl["y"], tbl["w"], tbl["h"]
        d.rectangle([x, y, x + w, y + 34],
                    fill="#2C5282", outline="#1A365D", width=2)
        tw, _ = text_size(d, tbl["name"], table_title_font)
        d.text((x + (w - tw) // 2, y + 9),
               tbl["name"], font=table_title_font, fill="white")
        d.rectangle([x, y + 34, x + w, y + h],
                    fill="white", outline="#1A365D", width=2)
        for i, (key_type, name, dtype) in enumerate(tbl["fields"]):
            row_y = y + 42 + i * 26
            if key_type == "PK" and i == 0 and len(tbl["fields"]) > 1:
                d.line([(x + 8, row_y + 22), (x + w - 8, row_y + 22)],
                       fill="#888888", width=1)
            if key_type == "PK":
                d.text((x + 10, row_y + 4), "PK",
                       font=pk_font, fill="#B7791F")
            elif key_type == "FK":
                d.text((x + 10, row_y + 4), "FK",
                       font=pk_font, fill="#3182CE")
            d.text((x + 42, row_y + 4), name,
                   font=(pk_font if key_type == "PK" else field_font),
                   fill="black")
            tw, _ = text_size(d, dtype, field_font)
            d.text((x + w - tw - 10, row_y + 4),
                   dtype, font=field_font, fill="#666666")
    for rel in relationships:
        x1, y1, x2, y2 = rel["from_x"], rel["from_y"], rel["to_x"], rel["to_y"]
        d.line([(x1, y1), (x2, y2)], fill="#2D3748", width=2)
        d.text((x1 + 5, y1 - 18), rel["card_from"],
               font=card_font, fill="#2D3748")
        d.text((x2 - 18, y2 - 18), rel["card_to"],
               font=card_font, fill="#2D3748")
    note_y = H - 40
    d.text((40, note_y),
           "PK - pirminis raktas, FK - isorinis raktas, 1:N - vienas-prie-daug rysys",
           font=field_font, fill="#666666")
    out = os.path.join(OUT_DIR, filename)
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


def draw_datasheet(filename, table_name, headers_widths, rows):
    total_w = sum(w for _, w in headers_widths)
    margin = 20
    title_h = 40
    header_h = 30
    row_h = 30
    autofield_h = 30
    img_w = total_w + 2 * margin
    img_h = margin + title_h + header_h + len(rows) * row_h + autofield_h + margin
    img = Image.new("RGB", (img_w, img_h), "white")
    d = ImageDraw.Draw(img)
    title_font = font(13, bold=True)
    header_font = font(13, bold=True)
    cell_font = font(13)
    d.text((margin + 5, margin + 10),
           f"{table_name} - LibreOffice Base: Table Data View",
           font=title_font, fill="#18A303")
    table_top = margin + title_h
    d.rectangle([margin, table_top, margin + total_w, table_top + header_h],
                fill="#E8E8E8", outline="#888888")
    cx = margin
    for label, cw in headers_widths:
        d.text((cx + 8, table_top + 7), label, font=header_font, fill="black")
        d.line([(cx + cw, table_top),
                (cx + cw, table_top + header_h + len(rows) * row_h)],
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
                 table_top + header_h + len(rows) * row_h],
                outline="#888888", width=1)
    autofield_y = table_top + header_h + len(rows) * row_h
    d.rectangle([margin, autofield_y, margin + total_w,
                 autofield_y + autofield_h],
                fill="#FAFAFA", outline="#CCCCCC")
    d.text((margin + 8, autofield_y + 7),
           "<AutoField>", font=cell_font, fill="#888888")
    out = os.path.join(OUT_DIR, filename)
    img.save(out, "PNG", optimize=True)
    print(f"Sukurta: {out}")


# ============================================================
# I sritis: Studentai
# ============================================================
def studentai_unf():
    headers = [("StudentID", 90), ("StudentName", 130), ("CourseID", 90),
               ("CourseName", 130), ("Instructor", 130),
               ("InstructorPhone", 150), ("Grade", 80)]
    rows = [
        (1, "Alice",   "C101", "Database", "Dr. Smith", "123-456-7890", "A"),
        (2, "Bob",     "C102", "Networks", "Dr. Jones", "234-567-8901", "B"),
        (1, "Alice",   "C103", "AI",       "Dr. Brown", "345-678-9012", "A"),
        (3, "Charlie", "C101", "Database", "Dr. Smith", "123-456-7890", "B"),
        (2, "Bob",     "C101", "Database", "Dr. Smith", "123-456-7890", "A"),
    ]
    dups = {(2, 1), (4, 1),
            (3, 3), (4, 3), (3, 4), (4, 4),
            (3, 5), (4, 5), (3, 6), (4, 6)}
    draw_unnormalized("img1_studentai_unf.png",
        "I sritis - Studentu registracija (NENORMALIZUOTA lentele)",
        headers, rows, dups)


def studentai_3nf():
    tables = [
        {"name": "STUDENTS", "x": 50, "y": 80, "w": 280, "h": 130,
         "fields": [("PK", "StudentID", "INTEGER"),
                    ("",   "StudentName", "VARCHAR(50)")]},
        {"name": "INSTRUCTORS", "x": 50, "y": 320, "w": 280, "h": 160,
         "fields": [("PK", "InstructorID", "INTEGER"),
                    ("",   "InstructorName", "VARCHAR(50)"),
                    ("",   "InstructorPhone", "VARCHAR(20)")]},
        {"name": "COURSES", "x": 540, "y": 200, "w": 280, "h": 160,
         "fields": [("PK", "CourseID", "VARCHAR(10)"),
                    ("",   "CourseName", "VARCHAR(50)"),
                    ("FK", "InstructorID", "INTEGER")]},
        {"name": "ENROLLMENTS", "x": 1020, "y": 200, "w": 280, "h": 180,
         "fields": [("PK", "EnrollmentID", "INTEGER"),
                    ("FK", "StudentID", "INTEGER"),
                    ("FK", "CourseID", "VARCHAR(10)"),
                    ("",   "Grade", "VARCHAR(2)")]},
    ]
    rels = [
        {"from_x": 330, "from_y": 145, "to_x": 1020, "to_y": 250,
         "card_from": "1", "card_to": "N"},
        {"from_x": 330, "from_y": 400, "to_x": 540, "to_y": 280,
         "card_from": "1", "card_to": "N"},
        {"from_x": 820, "from_y": 280, "to_x": 1020, "to_y": 290,
         "card_from": "1", "card_to": "N"},
    ]
    draw_3nf("img2_studentai_3nf.png",
        "I sritis - Studentu registracija po normalizavimo (3NF)",
        tables, rels)


def studentai_data():
    headers = [("ENROLLMENTID", 130), ("STUDENTID", 110),
               ("COURSEID", 110), ("GRADE", 100)]
    rows = [(1, 1, "C101", "A"), (2, 2, "C102", "B"),
            (3, 1, "C103", "A"), (4, 3, "C101", "B"), (5, 2, "C101", "A")]
    draw_datasheet("img7_data_studentai.png",
                   "ENROLLMENTS (studentai.odb)", headers, rows)


# ============================================================
# II sritis: Pardavimai
# ============================================================
def pardavimai_unf():
    headers = [("TransID", 70), ("CustID", 70), ("CustName", 110),
               ("ProdID", 70), ("ProdName", 110), ("Salesperson", 100),
               ("SalesPhone", 130), ("Qty", 60), ("Price", 70)]
    rows = [
        (1, 101, "John Doe",   "P001", "Laptop",   "Alice",   "123-456-7890", 1, 1000),
        (2, 102, "Jane Smith", "P002", "Mouse",    "Bob",     "234-567-8901", 2, 50),
        (3, 101, "John Doe",   "P003", "Keyboard", "Alice",   "123-456-7890", 1, 75),
        (4, 103, "Mike Brown", "P001", "Laptop",   "Charlie", "345-678-9012", 1, 1000),
        (5, 102, "Jane Smith", "P001", "Laptop",   "Alice",   "123-456-7890", 1, 1000),
    ]
    dups = {(2, 1), (2, 2), (4, 1), (4, 2),
            (3, 3), (4, 3), (3, 4), (4, 4),
            (2, 5), (4, 5), (2, 6), (4, 6),
            (3, 8), (4, 8)}
    draw_unnormalized("img3_pardavimai_unf.png",
        "II sritis - Pardavimai (NENORMALIZUOTA lentele)",
        headers, rows, dups)


def pardavimai_3nf():
    tables = [
        {"name": "CUSTOMERS", "x": 50, "y": 80, "w": 260, "h": 130,
         "fields": [("PK", "CustomerID", "INTEGER"),
                    ("",   "CustomerName", "VARCHAR(50)")]},
        {"name": "PRODUCTS", "x": 50, "y": 280, "w": 260, "h": 160,
         "fields": [("PK", "ProductID", "VARCHAR(10)"),
                    ("",   "ProductName", "VARCHAR(50)"),
                    ("",   "Price", "DECIMAL(8,2)")]},
        {"name": "SALESPEOPLE", "x": 50, "y": 510, "w": 260, "h": 160,
         "fields": [("PK", "SalespersonID", "INTEGER"),
                    ("",   "SalespersonName", "VARCHAR(50)"),
                    ("",   "SalespersonPhone", "VARCHAR(20)")]},
        {"name": "SALES", "x": 750, "y": 230, "w": 280, "h": 200,
         "fields": [("PK", "TransactionID", "INTEGER"),
                    ("FK", "CustomerID", "INTEGER"),
                    ("FK", "ProductID", "VARCHAR(10)"),
                    ("FK", "SalespersonID", "INTEGER"),
                    ("",   "Quantity", "INTEGER")]},
    ]
    rels = [
        {"from_x": 310, "from_y": 145, "to_x": 750, "to_y": 280,
         "card_from": "1", "card_to": "N"},
        {"from_x": 310, "from_y": 360, "to_x": 750, "to_y": 320,
         "card_from": "1", "card_to": "N"},
        {"from_x": 310, "from_y": 590, "to_x": 750, "to_y": 360,
         "card_from": "1", "card_to": "N"},
    ]
    draw_3nf("img4_pardavimai_3nf.png",
        "II sritis - Pardavimai po normalizavimo (3NF)", tables, rels)


def pardavimai_data():
    headers = [("TRANSACTIONID", 130), ("CUSTOMERID", 110),
               ("PRODUCTID", 100), ("SALESPERSONID", 130), ("QUANTITY", 100)]
    rows = [(1, 101, "P001", 1, 1), (2, 102, "P002", 2, 2),
            (3, 101, "P003", 1, 1), (4, 103, "P001", 3, 1),
            (5, 102, "P001", 1, 1)]
    draw_datasheet("img8_data_pardavimai.png",
                   "SALES (pardavimai.odb)", headers, rows)


# ============================================================
# III sritis: Sandeliai
# ============================================================
def sandeliai_unf():
    headers = [("WhID", 60), ("WhLocation", 110), ("ProdID", 70),
               ("ProdName", 100), ("SuppID", 70), ("SuppName", 110),
               ("SuppPhone", 130), ("Qty", 60), ("DeliveryDate", 110)]
    rows = [
        (1, "New York",    "P001", "Laptop",   "S001", "TechCorp",   "123-456-7890",  50, "2025-04-01"),
        (2, "Los Angeles", "P002", "Mouse",    "S002", "SupplyCo",   "234-567-8901", 200, "2025-04-02"),
        (1, "New York",    "P003", "Keyboard", "S003", "KeyMasters", "345-678-9012", 100, "2025-04-03"),
        (3, "Chicago",     "P001", "Laptop",   "S001", "TechCorp",   "123-456-7890",  75, "2025-04-04"),
        (2, "Los Angeles", "P001", "Laptop",   "S001", "TechCorp",   "123-456-7890",  50, "2025-04-05"),
    ]
    dups = {(2, 1), (4, 1),
            (3, 2), (4, 2), (3, 3), (4, 3),
            (3, 4), (4, 4), (3, 5), (4, 5), (3, 6), (4, 6)}
    draw_unnormalized("img5_sandeliai_unf.png",
        "III sritis - Sandeliai ir tiekejai (NENORMALIZUOTA lentele)",
        headers, rows, dups)


def sandeliai_3nf():
    tables = [
        {"name": "WAREHOUSES", "x": 50, "y": 80, "w": 280, "h": 130,
         "fields": [("PK", "WarehouseID", "INTEGER"),
                    ("",   "WarehouseLocation", "VARCHAR(80)")]},
        {"name": "PRODUCTS", "x": 50, "y": 280, "w": 280, "h": 130,
         "fields": [("PK", "ProductID", "VARCHAR(10)"),
                    ("",   "ProductName", "VARCHAR(50)")]},
        {"name": "SUPPLIERS", "x": 50, "y": 480, "w": 280, "h": 160,
         "fields": [("PK", "SupplierID", "VARCHAR(10)"),
                    ("",   "SupplierName", "VARCHAR(50)"),
                    ("",   "SupplierPhone", "VARCHAR(20)")]},
        {"name": "DELIVERIES", "x": 750, "y": 230, "w": 320, "h": 230,
         "fields": [("PK", "DeliveryID", "INTEGER"),
                    ("FK", "WarehouseID", "INTEGER"),
                    ("FK", "ProductID", "VARCHAR(10)"),
                    ("FK", "SupplierID", "VARCHAR(10)"),
                    ("",   "Quantity", "INTEGER"),
                    ("",   "DeliveryDate", "DATE")]},
    ]
    rels = [
        {"from_x": 330, "from_y": 145, "to_x": 750, "to_y": 290,
         "card_from": "1", "card_to": "N"},
        {"from_x": 330, "from_y": 340, "to_x": 750, "to_y": 330,
         "card_from": "1", "card_to": "N"},
        {"from_x": 330, "from_y": 560, "to_x": 750, "to_y": 370,
         "card_from": "1", "card_to": "N"},
    ]
    draw_3nf("img6_sandeliai_3nf.png",
        "III sritis - Sandeliai ir tiekejai po normalizavimo (3NF)",
        tables, rels)


def sandeliai_data():
    headers = [("DELIVERYID", 110), ("WAREHOUSEID", 130),
               ("PRODUCTID", 100), ("SUPPLIERID", 100),
               ("QUANTITY", 90), ("DELIVERYDATE", 130)]
    rows = [(1, 1, "P001", "S001", 50,  "2025-04-01"),
            (2, 2, "P002", "S002", 200, "2025-04-02"),
            (3, 1, "P003", "S003", 100, "2025-04-03"),
            (4, 3, "P001", "S001", 75,  "2025-04-04"),
            (5, 2, "P001", "S001", 50,  "2025-04-05")]
    draw_datasheet("img9_data_sandeliai.png",
                   "DELIVERIES (sandeliai.odb)", headers, rows)


def main():
    print("Generuojami PU6 paveikslai...")
    studentai_unf(); studentai_3nf(); studentai_data()
    pardavimai_unf(); pardavimai_3nf(); pardavimai_data()
    sandeliai_unf(); sandeliai_3nf(); sandeliai_data()
    print("Baigta. 9 paveikslai.")


if __name__ == "__main__":
    main()
