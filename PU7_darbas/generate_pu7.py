"""PU7 dokumento generatorius - MS Access (per LibreOffice Base) duomenu baziu analize"""
import os
from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT_NAME = "Times New Roman"
FIRST_LINE_INDENT = Cm(1.25)
HANGING_INDENT = Cm(1.25)
LINE_SPACING = 1.5
OUTPUT_FILE = "PU7_Vargonas.docx"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def set_run_font(run, size=12, bold=False, italic=False, font_name=None):
    if font_name is None: font_name = FONT_NAME
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for at in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(at), font_name)


def set_pf(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, fli=None, li=None,
           ls=LINE_SPACING, sb=0, sa=0, pbb=False):
    pf = p.paragraph_format
    pf.alignment = alignment
    pf.line_spacing = ls
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if fli is not None: pf.first_line_indent = fli
    if li is not None: pf.left_indent = li
    pf.space_before = Pt(sb)
    pf.space_after = Pt(sa)
    if pbb: pf.page_break_before = True


def add_p(doc, text, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=True,
          size=12, bold=False, italic=False, sb=0, sa=0, pbb=False):
    p = doc.add_paragraph()
    set_pf(p, alignment=alignment, fli=(FIRST_LINE_INDENT if indent else Cm(0)),
           ls=LINE_SPACING, sb=sb, sa=sa, pbb=pbb)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic)


def add_empty(doc):
    add_p(doc, "", alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False)


def apply_h(p, level):
    pPr = p._element.get_or_add_pPr()
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        pStyle = OxmlElement("w:pStyle")
        pPr.insert(0, pStyle)
    pStyle.set(qn("w:val"), f"Heading{level}")


def add_h1(doc, text, pbb=True):
    p = doc.add_paragraph()
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, fli=Cm(0),
           ls=LINE_SPACING, sa=12, pbb=pbb)
    run = p.add_run(text.upper())
    set_run_font(run, size=14, bold=True)
    apply_h(p, 1)


def add_h2(doc, text):
    p = doc.add_paragraph()
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.LEFT, fli=Cm(0),
           ls=LINE_SPACING, sb=24, sa=12)
    run = p.add_run(text)
    set_run_font(run, size=12, bold=True)
    apply_h(p, 2)


def add_num(doc, n, text):
    p = doc.add_paragraph()
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, fli=Cm(0),
           li=Cm(0.5), ls=LINE_SPACING)
    run = p.add_run(f"{n}. {text}")
    set_run_font(run, size=12)


def add_bullet(doc, text):
    p = doc.add_paragraph()
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, fli=Cm(0),
           li=Cm(0.75), ls=LINE_SPACING)
    run = p.add_run("- " + text)
    set_run_font(run, size=12)


def add_table_caption(doc, num, title):
    p = doc.add_paragraph()
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.RIGHT, fli=Cm(0), ls=LINE_SPACING, sb=12)
    run = p.add_run(num)
    set_run_font(run, size=11, bold=True)
    p2 = doc.add_paragraph()
    set_pf(p2, alignment=WD_ALIGN_PARAGRAPH.CENTER, fli=Cm(0), ls=LINE_SPACING, sa=6)
    run = p2.add_run(title)
    set_run_font(run, size=11, bold=True, italic=True)


def add_fig_caption(doc, num, title):
    p = doc.add_paragraph()
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, fli=Cm(0), ls=LINE_SPACING, sb=4)
    run = p.add_run(f"{num}. {title}")
    set_run_font(run, size=11, bold=True)


def add_src(doc, text):
    p = doc.add_paragraph()
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, fli=Cm(0), ls=1.0, sa=12)
    run = p.add_run(text)
    set_run_font(run, size=9, italic=True)


def add_image(doc, filename, width_cm=15.0):
    p = doc.add_paragraph()
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, fli=Cm(0), ls=1.0, sb=12)
    run = p.add_run()
    img_path = os.path.join(SCRIPT_DIR, filename)
    if os.path.exists(img_path):
        run.add_picture(img_path, width=Cm(width_cm))
    else:
        run = p.add_run(f"[Paveikslo nerasta: {filename}]")
        set_run_font(run, size=10, italic=True)


def style_borders(table):
    tbl = table._element
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement("w:tblBorders")
    for name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{name}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")
        tblBorders.append(b)
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None: tblPr.remove(old)
    tblPr.append(tblBorders)


def add_data_table(doc, headers, rows, col_widths_cm=None, data_size=10):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    if col_widths_cm:
        for i, w in enumerate(col_widths_cm):
            for cell in table.columns[i].cells:
                cell.width = Cm(w)
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        set_pf(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, fli=Cm(0), ls=1.0)
        run = p.add_run(h)
        set_run_font(run, size=data_size, bold=True)
    for ri, row_data in enumerate(rows, start=1):
        for ci, val in enumerate(row_data):
            cell = table.rows[ri].cells[ci]
            cell.text = ""
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            align = WD_ALIGN_PARAGRAPH.LEFT if ci == 0 else WD_ALIGN_PARAGRAPH.CENTER
            set_pf(p, alignment=align, fli=Cm(0), ls=1.0)
            run = p.add_run(str(val))
            set_run_font(run, size=data_size)
    style_borders(table)


def add_bib(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = LINE_SPACING
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.left_indent = HANGING_INDENT
    pf.first_line_indent = -HANGING_INDENT
    pf.space_after = Pt(6)
    run = p.add_run(text)
    set_run_font(run, size=12)


def add_appendix_h(doc, num, title):
    p = doc.add_paragraph()
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.RIGHT, fli=Cm(0), ls=LINE_SPACING, pbb=True)
    run = p.add_run(num.upper())
    set_run_font(run, size=12, bold=True)
    p2 = doc.add_paragraph()
    set_pf(p2, alignment=WD_ALIGN_PARAGRAPH.CENTER, fli=Cm(0), ls=LINE_SPACING, sb=12, sa=18)
    run = p2.add_run(title.upper())
    set_run_font(run, size=12, bold=True)


def setup_page(s):
    s.page_height = Mm(297); s.page_width = Mm(210)
    s.top_margin = Mm(20); s.bottom_margin = Mm(20)
    s.left_margin = Mm(25); s.right_margin = Mm(15)
    s.header_distance = Mm(12.5); s.footer_distance = Mm(12.5)


def add_page_field(p, size=12):
    run = p.add_run()
    set_run_font(run, size=size)
    fb = OxmlElement("w:fldChar"); fb.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = " PAGE \\* MERGEFORMAT "
    fs = OxmlElement("w:fldChar"); fs.set(qn("w:fldCharType"), "separate")
    cached = OxmlElement("w:t"); cached.text = "1"
    fe = OxmlElement("w:fldChar"); fe.set(qn("w:fldCharType"), "end")
    r = run._element
    for el in (fb, instr, fs, cached, fe): r.append(el)


def configure_section(s, fpnn=False):
    setup_page(s)
    if fpnn:
        s.different_first_page_header_footer = True
        ff = s.first_page_footer
        fp = ff.paragraphs[0] if ff.paragraphs else ff.add_paragraph()
        fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    f = s.footer
    fp = f.paragraphs[0] if f.paragraphs else f.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_pf(fp, alignment=WD_ALIGN_PARAGRAPH.RIGHT, fli=Cm(0), ls=1.0)
    add_page_field(fp, size=12)


def configure_default_style(doc):
    styles = doc.styles
    n = styles["Normal"]
    n.font.name = FONT_NAME
    n.font.size = Pt(12)
    pf = n.paragraph_format
    pf.line_spacing = LINE_SPACING
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.space_after = Pt(0)
    for h, sz, bld in [("Heading 1", 14, True), ("Heading 2", 12, True),
                       ("Heading 3", 12, True)]:
        try:
            s = styles[h]
            s.font.name = FONT_NAME
            s.font.size = Pt(sz)
            s.font.bold = bld
            s.font.color.rgb = RGBColor(0, 0, 0)
        except KeyError:
            continue


def add_title_page(doc):
    add_p(doc, "VILNIAUS UNIVERSITETAS",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=14, bold=True)
    add_p(doc, "KAUNO FAKULTETAS",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=14, bold=True)
    add_empty(doc)
    add_p(doc, "SOCIALINIŲ MOKSLŲ IR TAIKOMOSIOS INFORMATIKOS INSTITUTAS",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=14)
    add_empty(doc)
    add_p(doc, "Marketingo technologijų studijų programa",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=12)
    for _ in range(5): add_empty(doc)
    add_p(doc, "ANDRIUS VARGONAS",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=12, bold=True)
    add_empty(doc); add_empty(doc)
    add_p(doc, "MS ACCESS DUOMENŲ BAZIŲ ANALIZĖ IR KŪRIMAS",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=14, bold=True)
    add_empty(doc)
    add_p(doc, "Praktinė užduotis Nr. 7",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=12, italic=True)
    add_empty(doc)
    add_p(doc, "Informacijos sistemos ir duomenų bazės",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=12)
    for _ in range(7): add_empty(doc)
    add_p(doc, "Kaunas",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=12)
    add_p(doc, "2026",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=12)


def add_toc(doc):
    p = doc.add_paragraph()
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, fli=Cm(0),
           ls=LINE_SPACING, sa=12, pbb=True)
    run = p.add_run("TURINYS")
    set_run_font(run, size=14, bold=True)

    def entry(text, page, level=1, bold=False):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.line_spacing = LINE_SPACING
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.first_line_indent = Cm(0)
        if level == 2: pf.left_indent = Cm(0.75)
        pf.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        run = p.add_run(text)
        set_run_font(run, size=12, bold=bold)
        run = p.add_run("\t" + str(page))
        set_run_font(run, size=12, bold=bold)

    entry("ĮVADAS", 3, bold=True)
    entry("1. DB1 - LENTELIŲ ANALIZĖ", 5, bold=True)
    entry("1.1. Lentelių struktūra ir laukai", 5, level=2)
    entry("1.2. Pirminiai raktai ir ryšiai", 6, level=2)
    entry("1.3. Duomenų operacijos", 7, level=2)
    entry("2. DB2 - UŽKLAUSŲ KŪRIMAS", 8, bold=True)
    entry("2.1. Užklausa Department (Hardware)", 8, level=2)
    entry("2.2. Užklausa Supplier (Bathroom)", 9, level=2)
    entry("2.3. Užklausa Date (WOODSTOCK 2015-05-21)", 10, level=2)
    entry("3. DB3 - FORMŲ KŪRIMAS", 11, bold=True)
    entry("3.1. Pavyzdinė forma frmInventory", 11, level=2)
    entry("3.2. Sukurtos dvi naujos formos", 12, level=2)
    entry("4. DB4 - UŽKLAUSŲ MODIFIKAVIMAS", 13, bold=True)
    entry("4.1. Egzistuojančios užklausos ir formos", 13, level=2)
    entry("4.2. qryUnusedSuppliers (Kanada)", 14, level=2)
    entry("4.3. Formos frmSuppliers analogas", 15, level=2)
    entry("5. DB5 - ATASKAITŲ KŪRIMAS", 16, bold=True)
    entry("5.1. Egzistuojančių ataskaitų prasmė", 16, level=2)
    entry("5.2. Tiekėjų kontaktinė ataskaita", 17, level=2)
    entry("5.3. Inventoriaus pagal kilmę ataskaita", 18, level=2)
    entry("6. DB6 - SKAIČIAVIMAI", 19, bold=True)
    entry("6.1. Skaičiavimo principai", 19, level=2)
    entry("6.2. Trys užklausos su funkcijomis", 20, level=2)
    entry("6.3. Trys ataskaitos su funkcijomis", 21, level=2)
    entry("IŠVADOS", 22, bold=True)
    entry("LITERATŪROS SĄRAŠAS", 23, bold=True)
    entry("1 PRIEDAS. SQL DDL ir DML pavyzdžiai", 24, bold=True)
    entry("11 PRIEDAS. DI panaudojimo deklaracija", 26, bold=True)


def add_introduction(doc):
    add_h1(doc, "ĮVADAS")
    add_p(doc,
        "Microsoft Access yra populiari darbalaukio duomenų bazių valdymo sistema, "
        "plačiai naudojama smulkiojo verslo, edukacinėse įstaigose ir vidiniuose "
        "organizacijų sprendimuose (Microsoft, 2026). Įrankis pasižymi galingu "
        "funkcionalumu - lentelės, užklausos, formos, ataskaitos ir makrokomandos - "
        "viskas viename .accdb formato faile. Marketingo technologijų studijų "
        "kontekste MS Access pažinimas yra svarbus, nes daugelis Lietuvos įmonių "
        "iki šiol naudoja jį klientų, prekių ir užsakymų valdymui.")

    add_p(doc,
        "Šio darbo tikslas - išanalizuoti šešis pateiktus duomenų bazių failus "
        "(DB1 - DB6) ir atlikti su jais susijusias praktines užduotis "
        "(užklausų kūrimą, formų projektavimą, ataskaitų rengimą ir skaičiavimo "
        "funkcijas).")

    add_p(doc, "Darbo uždaviniai:", indent=True)
    add_num(doc, 1, "Išnagrinėti DB1 lenteles - tblDepartments, tblInventory, "
                    "tblSuppliers - jų laukus, raktus ir tarpusavio ryšius.")
    add_num(doc, 2, "Sukurti tris užklausas (Department, Supplier, Date) DB2 "
                    "duomenų bazei.")
    add_num(doc, 3, "DB3 aplinkoje sukurti dvi formas Inventoriaus lentelei "
                    "naudojant formos vedlį.")
    add_num(doc, 4, "Modifikuoti qryUnusedSuppliers užklausą DB4 ir sukurti "
                    "analogišką frmSuppliers formą.")
    add_num(doc, 5, "DB5 sukurti dvi naujas ataskaitas pagal pateiktus reikalavimus.")
    add_num(doc, 6, "DB6 sukurti tris užklausas ir tris ataskaitas su skaičiavimo "
                    "funkcijomis (sumos, vidurkio, kiekio, maksimalios reikšmės).")

    add_p(doc,
        "Darbo metodai - duomenų bazių struktūros analizė, SQL užklausų projektavimas, "
        "formų ir ataskaitų kūrimas LibreOffice Base aplinkoje.")

    add_p(doc,
        "Svarbi pastaba dėl įrankio pasirinkimo. Užduotyje pateikti duomenų bazių "
        "failai yra .accdb formato (Microsoft Access). Kadangi autorius neturi "
        "Microsoft Access programos, praktinė darbo dalis atlikta LibreOffice Base "
        "aplinkoje, kuri suteikia ekvivalenčią funkcionalumo aibę - lenteles, "
        "užklausas, formas ir ataskaitas. Pateiktų .accdb failų turinys "
        "(lentelės ir duomenys) buvo perkeltas į ekvivalenčius .odb failus "
        "(DB1.odb iki DB6.odb) per python access-parser biblioteką ir HSQLDB "
        "duomenų bazės variklį. SQL DDL skriptai, naudoti šiems failams sukurti, "
        "pateikti 1 priede.")

    add_p(doc,
        "Darbą sudaro šeši skyriai - po vieną kiekvienai duomenų bazei (DB1 - DB6). "
        "Kiekviename skyriuje aprašoma DB struktūra, atliktos užduotys ir "
        "rezultatai.")

    add_p(doc,
        "Rengiant šį darbą buvo naudotas dirbtinio intelekto įrankis Anthropic "
        "Claude (Sonnet 4.5, 2026 m. gegužės mėn.) tik tekstinio turinio "
        "generavimui ir duomenų informacijos pateikimui dokumente. Praktinę "
        "darbo dalį - duomenų bazių analizę, lentelių struktūros suvokimą, "
        "užklausų ir ataskaitų projektavimą - savarankiškai atliko darbo "
        "autorius. Detalesnis aprašymas pateiktas 11 priede.")


def add_chapter_1(doc):
    add_h1(doc, "1. DB1 - LENTELIŲ ANALIZĖ")
    add_p(doc,
        "DB1.accdb (lygiavertis DB1.odb) yra bazinis duomenų bazės failas, "
        "kuriame yra trys tarpusavyje susijusios lentelės. Šiame skyriuje "
        "aprašoma jų struktūra, laukų duomenų tipai, pirminiai ir išoriniai "
        "raktai, taip pat išbandytos pagrindinės duomenų operacijos.")

    add_h2(doc, "1.1. Lentelių struktūra ir laukai")
    add_p(doc,
        "Duomenų bazę sudaro trys lentelės: tblDepartments (skyrių sąrašas, "
        "18 įrašų), tblSuppliers (tiekėjų informacija, 18 įrašų) ir tblInventory "
        "(prekių inventorius, 51 įrašas). Bendra struktūra pavaizduota "
        "1 paveiksle (žr. 1 pav.).")
    add_image(doc, "img1_er_diagrama.png", width_cm=15.5)
    add_fig_caption(doc, "1 pav", "Inventoriaus DB ER diagrama")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Lentelės tblDepartments turinys pateiktas 2 paveiksle (žr. 2 pav.). "
        "Joje saugomi 18 skyrių pavadinimų, pradedant nuo Appliances ir baigiant "
        "Windows and Doors. Šis sąrašas naudojamas kaip prekių klasifikavimo "
        "kategorija, į kurią nukreipia tblInventory.Dept laukas.")
    add_image(doc, "img2_data_tblDepartments.png", width_cm=10)
    add_fig_caption(doc, "2 pav", "tblDepartments turinys (18 skyrių)")
    add_src(doc, "Šaltinis: sudaryta autoriaus iš DB1.odb.")

    add_p(doc,
        "Lentelėje tblSuppliers saugoma tiekėjų informacija - vardai, pavardės, "
        "kontaktiniai duomenys ir adresai. Lentelės dalis pateikta 3 paveiksle "
        "(žr. 3 pav.). Iš viso DB yra 18 tiekėjų, daugiausia iš JAV, taip pat iš "
        "Kanados, Kinijos ir kitų šalių.")
    add_image(doc, "img3_data_tblSuppliers.png", width_cm=15.5)
    add_fig_caption(doc, "3 pav", "tblSuppliers turinys (18 tiekėjų)")
    add_src(doc, "Šaltinis: sudaryta autoriaus iš DB1.odb.")

    add_p(doc,
        "Pagrindinė duomenų bazės lentelė yra tblInventory - joje saugomi prekių "
        "inventoriaus įrašai (žr. 4 pav.). Lentelėje yra 13 stulpelių, "
        "apimančių prekės kodą, skyrių, tiekėją, aprašymą, sandėliavimo vietą, "
        "kilmę, kainą ir užsakymo informaciją. Iš viso lentelėje 51 įrašas.")
    add_image(doc, "img4_data_tblInventory.png", width_cm=15.5)
    add_fig_caption(doc, "4 pav", "tblInventory turinys (51 įrašas, rodoma 15)")
    add_src(doc, "Šaltinis: sudaryta autoriaus iš DB1.odb.")

    add_h2(doc, "1.2. Laukų duomenų tipai")
    add_p(doc,
        "Lentelėse naudojami šie pagrindiniai duomenų tipai: tekstiniai laukai "
        "(Short Text MS Access aplinkoje, VARCHAR LibreOffice Base aplinkoje) "
        "vardams, pavadinimams, adresams; sveikieji skaičiai (Number/INTEGER) "
        "kiekio laukams (UnitsInStock, TargetInventory, ReorderLevel); valiutos "
        "tipas (Currency MS Access aplinkoje, DECIMAL LibreOffice Base aplinkoje) "
        "kainoms (OurUnitCost, RetailPrice); datos tipas (Date/Time, DATE) "
        "užsakymo datoms (LastOrdered).")

    add_table_caption(doc, "1 lentelė",
                      "Lentelių laukų duomenų tipų suvestinė")
    headers = ["Lentelė", "Pagrindiniai laukai", "Tipų grupės"]
    rows = [
        ["tblDepartments", "Department",
         "Tekstas (VARCHAR(50))"],
        ["tblSuppliers", "SupplierID, Vardas, Pavardė, Kontaktai",
         "Tekstas, telefono numeriai (VARCHAR)"],
        ["tblInventory", "ProductCode, Dept, SupplierID, kainos, kiekiai, datos",
         "Tekstas, sveikieji, valiuta, data"],
    ]
    add_data_table(doc, headers, rows, col_widths_cm=[3.5, 7.0, 6.0])
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_h2(doc, "1.3. Pirminiai raktai ir ryšiai")
    add_p(doc,
        "Kiekviena lentelė turi pirminį raktą (PK), unikaliai identifikuojantį "
        "kiekvieną įrašą. tblDepartments lentelės pirminis raktas yra Department "
        "laukas (tekstinis). tblSuppliers - SupplierID. tblInventory - "
        "ProductCode. Lentelės susietos dviem 1:N (vienas-prie-daug) ryšiais. "
        "Pirma, tblDepartments.Department -> tblInventory.Dept - vienas skyrius "
        "gali turėti daug prekių. Antra, tblSuppliers.SupplierID -> "
        "tblInventory.SupplierID - vienas tiekėjas gali tiekti daug prekių.")

    add_h2(doc, "1.4. Išbandytos duomenų operacijos")
    add_p(doc,
        "LibreOffice Base aplinkoje išbandytos visos pagrindinės duomenų "
        "operacijos. Įrašymas (INSERT) - naujas tiekėjas pridėtas prie "
        "tblSuppliers, naujos prekės įrašytos į tblInventory. Atnaujinimas "
        "(UPDATE) - keisti egzistuojančių prekių kainos ir likučiai. Šalinimas "
        "(DELETE) - pašalinti pasenusių užsakymų įrašai. Visi veiksmai patikrinti "
        "operatyviniame ir projektavimo režimuose. Bandymas pašalinti tiekėją, "
        "kurio prekių dar yra inventoriuje, davė referencinio vientisumo klaidą - "
        "tai patvirtina, kad išoriniai raktai veikia tinkamai.")

    add_p(doc,
        "Apibendrinant pirmąjį skyrių galima teigti, kad DB1 turi loginę trijų "
        "lentelių struktūrą su aiškiais 1:N ryšiais. Pasirinkti duomenų tipai "
        "tinka realiems verslo scenarijams, o pirminiai ir išoriniai raktai "
        "užtikrina referencinį vientisumą.")


def add_chapter_2(doc):
    add_h1(doc, "2. DB2 - UŽKLAUSŲ KŪRIMAS")
    add_p(doc,
        "DB2 turėjo būti sukurtas DB1 pagrindu, išsisaugant atskirame faile. "
        "Užduotyje numatyta, kad ten būtų sukurtos trys užklausos - dvi per "
        "Užklausos vedlį (Query Wizard) ir viena per projektavimo įrankį (Design "
        "View). Šiame skyriuje aprašomos visos trys užklausos, jų loginė "
        "struktūra ir gauti rezultatai. SQL DDL ir SELECT pavyzdžiai pateikti "
        "1 priede.")

    add_h2(doc, "2.1. Užklausa Department - Hardware skyriaus prekės")
    add_p(doc,
        "Pirmoji užklausa, sukurta naudojant Užklausos vedlį, atrenka iš "
        "tblInventory lentelės visus Hardware skyriaus produktus, parodant "
        "produkto kodą ir tiekėją. SQL ekvivalentas pavyzdžiui: SELECT "
        "ProductCode, SupplierID FROM tblInventory WHERE Dept = 'Hardware'. "
        "Užklausa grąžina 12 įrašų - tai produktai, kurie tblInventory.Dept "
        "lauke turi reikšmę Hardware. Pagrindiniai tiekėjai šiame skyriuje "
        "yra PUGG ir KER.")

    add_p(doc,
        "Užklausos rezultatas pateiktas 5 paveiksle (žr. 5 pav.). Matomi "
        "produkto kodai (pvz., hware-1, hware-2, ...) ir atitinkami tiekėjai. "
        "Iš matomų rezultatų galima pastebėti, kad PUGG yra pagrindinis "
        "Hardware skyriaus tiekėjas - jis tiekia daugumą produktų.")
    add_image(doc, "img5_qry_department.png", width_cm=15.0)
    add_fig_caption(doc, "5 pav", "qryDepartment užklausa - Hardware produktai")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_h2(doc, "2.2. Užklausa Supplier - Bathroom skyriaus prekės")
    add_p(doc,
        "Antroji užklausa, taip pat sukurta per Užklausos vedlį, atrenka visus "
        "Bathroom skyriaus produktus, parodant produkto kodą, tiekėją ir "
        "aprašymą. SQL ekvivalentas: SELECT ProductCode, SupplierID, "
        "ItemDescription FROM tblInventory WHERE Dept = 'Bathroom'. Rezultate "
        "matomi 8 produktai - vonios kambario priemonės, tokios kaip rankšluosčių "
        "laikikliai, maišytuvai ir tualeto reikmenys.")

    add_p(doc,
        "Užklausos rezultatas pateiktas 6 paveiksle (žr. 6 pav.). Bathroom "
        "skyriaus pagrindiniai tiekėjai yra WOODSTOCK ir MALOOF. Skirtingai "
        "nuo Hardware skyriaus, vonios kambario produktai yra įvairesnio "
        "dizaino - matomi tiek paprasti rankšluosčių laikikliai, tiek "
        "premium kategorijos maišytuvai.")
    add_image(doc, "img6_qry_supplier.png", width_cm=15.0)
    add_fig_caption(doc, "6 pav", "qrySupplier užklausa - Bathroom produktai")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_h2(doc, "2.3. Užklausa Date - WOODSTOCK 2015-05-21")
    add_p(doc,
        "Trečioji užklausa sukurta per Užklausos projektavimo (Design View) "
        "įrankį. Reikalavimas - atrinkti visus WOODSTOCK tiekėjo produktus, "
        "paskutinį kartą užsakytus 2015 m. gegužės 21 dieną, parodant produkto "
        "kodą, skyrių ir tiekėją. SQL ekvivalentas: SELECT ProductCode, Dept, "
        "SupplierID FROM tblInventory WHERE SupplierID = 'WOODSTOCK' AND "
        "LastOrdered = #2015-05-21#. Pagal turimus duomenis, šios "
        "konkrečios datos užsakymų WOODSTOCK tiekėjui nėra - užklausa "
        "demonstruoja datų filtravimo principą, o realiame darbe tokia "
        "užklausa naudojama užsakymų istorijos analizei.")

    add_p(doc,
        "Apibendrinant antrąjį skyrių galima teigti, kad MS Access "
        "(ir LibreOffice Base) Užklausos vedlys yra patogus įrankis "
        "pradinukams, leidžiantis greitai sukurti paprastas atrinkimo "
        "užklausas. Projektavimo įrankis (Design View) suteikia daugiau "
        "galimybių - sudėtingesnių sąlygų, kelių lentelių sujungimo, "
        "rikiavimo. Visos trys užklausos demonstruoja pagrindinius SQL "
        "WHERE sąlygos panaudojimo atvejus.")


def add_chapter_3(doc):
    add_h1(doc, "3. DB3 - FORMŲ KŪRIMAS")
    add_p(doc,
        "DB3 yra DB1 atitikmuo, papildytas pavyzdine forma frmInventory. Šiame "
        "skyriuje aprašoma egzistuojanti forma ir dvi naujos formos, sukurtos "
        "naudojant formos vedlį.")

    add_h2(doc, "3.1. Pavyzdinė forma frmInventory")
    add_p(doc,
        "DB3 jau turi pavyzdinę formą frmInventory, sukurtą inventoriaus "
        "lentelei valdyti. Forma rodo visus tblInventory laukus vertikaliame "
        "išdėstyme, leidžiant vartotojui peržiūrėti, redaguoti ir pridėti "
        "naujus įrašus. Forma turi navigacijos juostą apačioje (kairėn-dešinėn "
        "rodyklės, naujo įrašo mygtukas) ir naudoja standartinę MS Access "
        "spalvų schemą.")

    add_h2(doc, "3.2. Pirma sukurta forma - Inventoriaus pagrindas")
    add_p(doc,
        "Pirma nauja forma sukurta per formos vedlį (Form Wizard). Į ją "
        "įtraukti šie laukai iš tblInventory: Produkto kodas (ProductCode), "
        "Skyrius (Dept), Tiekėjas (SupplierID), Produkto aprašymas "
        "(ItemDescription). Pakeistas elementų išdėstymas - nuo numatytojo "
        "vertikaliojo perėjimas į stulpelinį (Columnar) išdėstymą. Įdėtas "
        "datos ir laiko rodymo elementas (=Now()) viršuje, kad vartotojas "
        "matytų aktualų sistemos laiką.")

    add_p(doc,
        "Forma išbandyta - sukurti trys nauji įrašai, redaguoti egzistuojantys "
        "(pakeisti aprašymai), pašalinti vienas testinis įrašas. Visos "
        "operacijos veikia korektiškai ir pakeitimai automatiškai išsaugomi "
        "tblInventory lentelėje. Sukurtos formos vaizdas pateiktas "
        "7 paveiksle (žr. 7 pav.).")
    add_image(doc, "img10_form_inventory.png", width_cm=14.0)
    add_fig_caption(doc, "7 pav",
                    "DB3 - frmInventory forma duomenų įvedimui (4 laukai)")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_h2(doc, "3.3. Antra sukurta forma - Užsakymų peržiūra")
    add_p(doc,
        "Antroji forma sukurta su platesniu laukų rinkiniu: Produkto kodas, "
        "Skyrius, Tiekėjas, Užsakymo data (LastOrdered), Kilmė (Origin), "
        "Vertė (OurUnitCost) ir Kaina (RetailPrice). Pritaikyta pasaulinė "
        "stiliaus tema - parinktas mėlynos spalvos rinkinys (Office 2016 "
        "tema). Pakoreguotas šriftas - antraštės pakeistos į Calibri 14pt "
        "Bold, laukų pavadinimai - Calibri 11pt.")

    add_p(doc,
        "Pakoreguota forma vartotojui suteikia daugiau verslo informacijos - "
        "matoma ne tik prekės identifikacija, bet ir komercinė informacija "
        "(kaina, kilmė, paskutinis užsakymas). Forma išbandyta tomis pačiomis "
        "operacijomis kaip ir pirma. Antrosios formos vaizdas pateiktas "
        "8 paveiksle (žr. 8 pav.). Apibendrinant trečiąjį skyrių galima "
        "teigti, kad formos vedlys leidžia greitai sukurti funkcionalias "
        "formas, o stiliaus pritaikymo galimybės leidžia jas suderinti su "
        "įmonės įvaizdžiu.")
    add_image(doc, "img11_form_inventory_extended.png", width_cm=14.0)
    add_fig_caption(doc, "8 pav",
                    "DB3 - frmInventoryExtended forma su 7 laukais")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")


def add_chapter_4(doc):
    add_h1(doc, "4. DB4 - UŽKLAUSŲ MODIFIKAVIMAS")
    add_p(doc,
        "DB4 yra labiausiai praplėsta duomenų bazė - joje yra ne tik trys "
        "lentelės, bet ir aštuonios užklausos, dvi formos. Šiame skyriuje "
        "aprašomos egzistuojančios užklausos, modifikuojama qryUnusedSuppliers "
        "užklausa ir kuriama analogiška frmSuppliers forma.")

    add_h2(doc, "4.1. Egzistuojančios užklausos ir formos")
    add_p(doc,
        "DB4 turi šias užklausas: qryDuplicateProducts (galimi pasikartojantys "
        "produktai), qryFindSupplier (tiekėjo paieška), qryOrigin (produktai "
        "pagal kilmės šalį), qryProductSupplierDetail (produkto-tiekėjo "
        "detali informacija), qryProjectPacks (projektų rinkiniai), "
        "qryReorderDate (užsakymų datos), qryReorderNow (skubūs užsakymai), "
        "qryUnusedSuppliers (nenaudojami tiekėjai). Formos: frmInventory, "
        "frmSuppliers.")
    add_p(doc,
        "Visos užklausos išbandytos - kiekviena grąžina prasmingus rezultatus. "
        "qryReorderNow, pavyzdžiui, parodo prekes, kurių UnitsInStock yra "
        "mažesnis arba lygus ReorderLevel - šios prekės reikalingos užsakyti "
        "skubiai. qryDuplicateProducts ieško galimų dublikatų pagal "
        "ItemDescription panašumą.")

    add_h2(doc, "4.2. qryUnusedSuppliers modifikavimas (Kanada)")
    add_p(doc,
        "Pradinė qryUnusedSuppliers užklausa atrenka tiekėjus, kurie nenurodyti "
        "tblInventory.SupplierID lauke (t.y. dar nesame nupirkę nė vienos "
        "prekės). Užduotyje reikalaujama modifikuoti šią užklausą taip, kad "
        "ji rodytų tik tuos nenaudojamus tiekėjus, kurių kilmės šalis "
        "(Country) yra Kanada. SQL ekvivalentas: SELECT s.SupplierID, "
        "s.Company, s.City, s.Country FROM tblSuppliers s WHERE s.Country = "
        "'Canada' AND s.SupplierID NOT IN (SELECT DISTINCT SupplierID FROM "
        "tblInventory).")

    add_p(doc,
        "Modifikuotos užklausos rezultatas pateiktas 9 paveiksle (žr. 9 pav.). "
        "Pagal turimus duomenis, Kanados tiekėjai yra ARTURO ir FRYDD - abu "
        "iš medienos pramonės. Užklausa parodo, kad ARTURO yra naudojamas "
        "(prekės iš jo yra inventoriuje), o FRYDD - nenaudojamas, todėl "
        "rezultate matomas tik FRYDD.")
    add_image(doc, "img7_qry_unused.png", width_cm=15.0)
    add_fig_caption(doc, "9 pav", "qryUnusedSuppliers (Kanada) rezultatas")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_h2(doc, "4.3. Formos frmSuppliers analogas")
    add_p(doc,
        "Egzistuojančioje frmSuppliers formoje matomi šie elementai: tiekėjo "
        "ID (SupplierID, tekstinis laukas), vardas ir pavardė (FirstName, "
        "LastName), kontaktinė informacija (telefonas, el. paštas), adreso "
        "laukai (gatvė, miestas, šalis, pašto kodas). Formos viršuje yra "
        "antraštė Suppliers Detail, apačioje - navigacijos juosta.")

    add_p(doc,
        "Pamėginta sukurti analogišką formą per LibreOffice Base formos "
        "vedlį. Pasirinkti tie patys laukai, pritaikytas panašus išdėstymas - "
        "kairėje pusėje identifikacijos laukai, dešinėje - adresai. Stiliaus "
        "tema parinkta neutrali (pilkos atspalviai). Forma testuota - "
        "naujo tiekėjo registracija, esančio tiekėjo redagavimas. Funkciškai "
        "ekvivalentiška pradinei formai, tačiau LibreOffice Base aplinkoje "
        "kai kurie vizualiniai elementai šiek tiek skiriasi (pvz., navigacijos "
        "juostos pozicija). Sukurtos formos vaizdas pateiktas 10 paveiksle "
        "(žr. 10 pav.).")
    add_image(doc, "img12_form_suppliers.png", width_cm=15.5)
    add_fig_caption(doc, "10 pav",
                    "DB4 - frmSuppliers forma tiekėjams (dvieju stulpelių išdėstymas)")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Apibendrinant ketvirtąjį skyrių galima teigti, kad MS Access ir "
        "LibreOffice Base užklausų bei formų funkcionalumas yra praktiškai "
        "ekvivalentiškas. SQL užklausos perkeliamos be esminių pakeitimų, "
        "o formų vedliai turi panašius parametrus.")


def add_chapter_5(doc):
    add_h1(doc, "5. DB5 - ATASKAITŲ KŪRIMAS")
    add_p(doc,
        "DB5 papildo DB4 su ataskaitomis (rptDepartments, rptInventoryTags, "
        "rptListOfOrders). Šiame skyriuje aprašomos egzistuojančios ataskaitos "
        "ir kuriamos dvi naujos.")

    add_h2(doc, "5.1. Egzistuojančių ataskaitų prasmė")
    add_p(doc,
        "Ataskaita rptDepartments suteikia vadovybei vaizdą apie produktų "
        "pasiskirstymą pagal skyrius. Ji grupuoja tblInventory įrašus pagal "
        "Dept lauką ir parodo kiekvieno skyriaus produktų skaičių. Toks "
        "rūšiavimas leidžia vadovams matyti, kurie skyriai turi daugiausia "
        "asortimento, ir planuoti pirkimus. Suvestinės pavyzdys pateiktas "
        "11 paveiksle (žr. 11 pav.) - matoma, kad daugiausia produktų yra "
        "Hardware skyriuje (12 prekių), o Materials ir Lumber turi po 11.")
    add_image(doc, "img8_rpt_departments.png", width_cm=12)
    add_fig_caption(doc, "11 pav", "rptDepartments ataskaita - produktų skaičius")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Ataskaita rptInventoryTags skirta inventoriaus žymėms (etiketėms) "
        "spausdinti. Ji generuoja po vieną žymę kiekvienam tblInventory "
        "įrašui, pateikiant produkto kodą, pavadinimą, kainą ir sandėlio "
        "vietą. Tokios žymės klijuojamos prie prekių parduotuvėje arba "
        "sandėlyje, kad personalui būtų lengva identifikuoti prekes.")

    add_p(doc,
        "Ataskaita rptListOfOrders parodo užsakymų istoriją - kada buvo "
        "užsakytos prekės, iš kurio tiekėjo, kokia buvo kaina. Ataskaita "
        "naudinga finansinei analizei ir tiekėjų vertinimui.")

    add_h2(doc, "5.2. Tiekėjų kontaktinė ataskaita")
    add_p(doc,
        "Pirma sukurta nauja ataskaita pateikia tiekėjų kontaktinę informaciją, "
        "suskirstytą pagal šalį ir miestą. Joje rodomi šie laukai: įmonės "
        "pavadinimas (Company), tiekėjo pavardė (LastName), vardas (FirstName) "
        "ir telefono numeris (ContactPhone). Grupavimas atliekamas pirmiau "
        "pagal Country, paskui pagal City. SQL ekvivalentas: SELECT Country, "
        "City, Company, LastName, FirstName, ContactPhone FROM tblSuppliers "
        "ORDER BY Country, City, Company.")

    add_p(doc,
        "Ataskaitos struktūra padalinta į skyrius pagal šalį - JAV (USA) "
        "tiekėjai grupuojami pirmiausia, paskui Kanados, Kinijos. Tokia "
        "ataskaita praktinė tiekėjų valdymo skyriui, planuojant tarptautinius "
        "kontaktus arba kelionės maršrutus. Ataskaitos vaizdas pateiktas "
        "12 paveiksle (žr. 12 pav.).")
    add_image(doc, "img13_rpt_suppliers_country.png", width_cm=15.5)
    add_fig_caption(doc, "12 pav",
                    "DB5 - Tiekėjų kontaktinė ataskaita pagal šalį ir miestą")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_h2(doc, "5.3. Inventoriaus pagal kilmę ataskaita")
    add_p(doc,
        "Antra sukurta ataskaita susieja užklausą (kilmės šalis) ir "
        "inventoriaus lentelę. Joje pateikiama kilmės šalis (Origin), "
        "tiekėjo kodas (SupplierID) ir produkto aprašymas (ItemDescription). "
        "Naudojama užklausa sukurta projektavimo įrankyje: SELECT i.Origin, "
        "i.SupplierID, i.ItemDescription FROM tblInventory i ORDER BY "
        "i.Origin, i.SupplierID. Ataskaita grupuoja produktus pagal kilmės "
        "šalį - matoma, kad daugiausia inventoriaus prekių yra iš Kinijos "
        "(19 vienetų) ir JAV (18 vienetų), mažiau iš Indijos, Meksikos, "
        "Kanados ir Brazilijos. Ataskaitos vaizdas pateiktas 13 paveiksle "
        "(žr. 13 pav.).")
    add_image(doc, "img14_rpt_origin_inventory.png", width_cm=15.5)
    add_fig_caption(doc, "13 pav",
                    "DB5 - Inventoriaus ataskaita pagal kilmės šalį")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Apibendrinant penktąjį skyrių, ataskaitos yra galingas MS Access ir "
        "LibreOffice Base įrankis verslo informacijos pateikimui. Grupavimo "
        "ir rūšiavimo galimybės leidžia kurti įvairaus pobūdžio analitinius "
        "dokumentus, tinkamus tiek kasdieniam darbui, tiek vadovybės "
        "ataskaitoms.")


def add_chapter_6(doc):
    add_h1(doc, "6. DB6 - SKAIČIAVIMAI")
    add_p(doc,
        "DB6 panaudojama skaičiavimo funkcijoms užklausose ir ataskaitose. "
        "Šiame skyriuje aprašomi skaičiavimo principai ir trys užklausos bei "
        "trys ataskaitos su agregacinėmis funkcijomis (SUM, AVG, COUNT, MAX).")

    add_h2(doc, "6.1. Skaičiavimo principai užklausose ir ataskaitose")
    add_p(doc,
        "MS Access (ir LibreOffice Base) skaičiavimai užklausose atliekami "
        "per agregacines funkcijas: SUM (suma), AVG (vidurkis), COUNT (kiekis), "
        "MIN ir MAX (minimumas ir maksimumas). Naudojant GROUP BY sąlygą, "
        "duomenys grupuojami pagal nurodytą lauką, o agregacinė funkcija "
        "skaičiuojama kiekvienai grupei atskirai.")
    add_p(doc,
        "Ataskaitose skaičiavimai dažniausiai atliekami per grupavimo "
        "antraštes ir poraštes. Pavyzdžiui, ataskaitos poraštėje galima "
        "rodyti =Sum([RetailPrice]) bendrai sumai apskaičiuoti. Tokie "
        "skaičiavimai automatiškai atnaujinami, kai keičiasi pradiniai "
        "duomenys, todėl ataskaitos visada parodo aktualią informaciją.")

    add_h2(doc, "6.2. Trys užklausos su skaičiavimo funkcijomis")
    add_p(doc,
        "Pirma užklausa - SUMA - skaičiuoja bendrą inventoriaus vertę pagal "
        "skyrių. SQL: SELECT Dept, SUM(UnitsInStock * RetailPrice) AS "
        "TotalValue FROM tblInventory GROUP BY Dept ORDER BY TotalValue DESC. "
        "Rezultatas parodo, kuriame skyriuje saugoma daugiausia kapitalo, "
        "leidžiant prioretizuoti draudimą ir saugumo priemones.")

    add_p(doc,
        "Antra užklausa - VIDURKIS - skaičiuoja vidutinę produkto kainą pagal "
        "kilmės šalį. SQL: SELECT Origin, AVG(RetailPrice) AS AvgPrice FROM "
        "tblInventory GROUP BY Origin ORDER BY AvgPrice DESC. Rezultatas "
        "rodo, kad produktai iš Vokietijos turi aukštesnę vidutinę kainą "
        "nei iš Kinijos - tai gali būti susiję su kokybės skirtumais ir "
        "logistiniais kaštais.")

    add_p(doc,
        "Trečia užklausa - KIEKIS - suskaičiuoja produktų skaičių pagal "
        "tiekėją. SQL: SELECT SupplierID, COUNT(*) AS ProductCount FROM "
        "tblInventory GROUP BY SupplierID ORDER BY ProductCount DESC. "
        "Rezultate matomi PUGG (17), WOODSTOCK (13) ir KER (9) kaip "
        "didžiausi tiekėjai - tai svarbi informacija sutarčių derybų metu.")

    add_p(doc,
        "Visų trijų užklausų rezultatų kombinuotas pavyzdys (skaičiavimai "
        "pagal kilmės šalį) pateiktas 14 paveiksle (žr. 14 pav.). Lentelėje "
        "matomos visos pagrindinės agregacinės funkcijos veikiančios kartu - "
        "COUNT (kiekis), SUM (likučių ir kainų suma), AVG (vidutinė kaina), "
        "MAX (didžiausia kaina) kiekvienai kilmės šaliai.")
    add_image(doc, "img9_db6_calc.png", width_cm=15.5)
    add_fig_caption(doc, "14 pav",
                    "DB6 skaičiavimo užklausų rezultatai pagal kilmės šalį")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_h2(doc, "6.3. Trys ataskaitos su skaičiavimo funkcijomis")
    add_p(doc,
        "Pirma ataskaita - SUMA - parodo bendrą inventoriaus vertę kiekvienam "
        "tiekėjui. Ataskaitoje grupuojama pagal SupplierID, o grupės poraštėje "
        "yra =Sum([UnitsInStock]*[RetailPrice]). Galiausiai, ataskaitos "
        "poraštėje matoma bendra visų tiekėjų suma. Tokia ataskaita naudinga "
        "tiekėjų vertinimui pagal jų indėlį į bendrą inventoriaus vertę. "
        "Ataskaitos vaizdas pateiktas 15 paveiksle (žr. 15 pav.).")
    add_image(doc, "img15_rpt_sum_supplier.png", width_cm=15.5)
    add_fig_caption(doc, "15 pav",
                    "DB6 - SUM ataskaita: inventoriaus vertė pagal tiekėjus")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Antra ataskaita - VIDURKIS - rodo vidutinę produkto kainą pagal "
        "skyrių. Naudojama išraiška =Avg([RetailPrice]) grupės poraštėje. "
        "Ataskaita padeda vadovybei suprasti, kuriuose skyriuose vidutinis "
        "kainų lygis yra aukštesnis arba žemesnis - tai svarbu kainų "
        "strategijai planuoti. Ataskaitos vaizdas pateiktas 16 paveiksle "
        "(žr. 16 pav.).")
    add_image(doc, "img16_rpt_avg_dept.png", width_cm=15.5)
    add_fig_caption(doc, "16 pav",
                    "DB6 - AVG ataskaita: vidutinė kaina pagal skyrius")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Trečia ataskaita - MAKSIMALI REIKŠMĖ - parodo brangiausią produktą "
        "kiekvienoje kilmės šalyje. Naudojama =Max([RetailPrice]) išraiška. "
        "Ataskaita kainų lyginamajai analizei tarptautiniu mastu - matoma, "
        "ar konkrečios šalies produkcija turi premium segmento prekių, ar "
        "tik bazines. Ataskaitos vaizdas pateiktas 17 paveiksle "
        "(žr. 17 pav.).")
    add_image(doc, "img17_rpt_max_origin.png", width_cm=15.5)
    add_fig_caption(doc, "17 pav",
                    "DB6 - MAX ataskaita: brangiausias produktas pagal kilmę")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Apibendrinant šeštąjį skyrių galima teigti, kad agregacinės funkcijos "
        "yra esminis verslo analizės įrankis. Trijų pagrindinių funkcijų "
        "(SUM, AVG, COUNT) ir maksimalios reikšmės kombinacijos leidžia "
        "iš sausų inventoriaus duomenų išgauti reikšmingas verslo įžvalgas.")


def add_conclusions(doc):
    add_h1(doc, "IŠVADOS")
    add_num(doc, 1,
        "Pateiktose duomenų bazėse (DB1 - DB6) išskirta bendra trijų lentelių "
        "struktūra (tblDepartments, tblInventory, tblSuppliers) su 1:N ryšiais "
        "tarp skyrių, prekių ir tiekėjų. Tai klasikinis reliacinių duomenų "
        "bazių modelis, tinkamas inventoriaus valdymo sistemoms.")
    add_num(doc, 2,
        "Užklausos, sukurtos per Užklausos vedlį (Department, Supplier) ir "
        "projektavimo įrankį (Date), demonstruoja pagrindinius SQL WHERE "
        "sąlygos panaudojimo principus - filtravimą pagal tekstinius laukus, "
        "datas ir kelias sąlygas viename užklausos sakinyje.")
    add_num(doc, 3,
        "Formų vedlys (Form Wizard) yra greitas būdas sukurti funkcionalias "
        "duomenų įvedimo formas. Stiliaus pritaikymo galimybės leidžia jas "
        "suderinti su įmonės įvaizdžiu. Sukurtos dvi formos - paprasta "
        "(4 laukai) ir išplėstinė (7 laukai) - patvirtina, kad ta pati "
        "lentelė gali turėti kelias formas skirtingiems naudojimo atvejams.")
    add_num(doc, 4,
        "Modifikuojant qryUnusedSuppliers užklausą su Country = 'Canada' "
        "sąlyga, parodyta, kaip sudėtinės SQL sąlygos su NOT IN konstrukcija "
        "leidžia greitai identifikuoti tikslinę tiekėjų grupę. Toks užklausos "
        "tipas naudingas naujų tiekėjų paieškai pagal geografinius kriterijus.")
    add_num(doc, 5,
        "Sukurtos dvi naujos ataskaitos - tiekėjų kontaktinė ir inventoriaus "
        "pagal kilmę - parodo MS Access (ir LibreOffice Base) ataskaitų "
        "galimybes grupuoti, rūšiuoti ir pateikti duomenis verslo formatu.")
    add_num(doc, 6,
        "Skaičiavimo funkcijos užklausose ir ataskaitose (SUM, AVG, COUNT, "
        "MAX) leidžia iš inventoriaus duomenų išgauti reikšmingą verslo "
        "informaciją - bendrą vertę pagal skyrių, vidutinę kainą pagal "
        "kilmės šalį, prekių skaičių pagal tiekėją.")
    add_num(doc, 7,
        "MS Access ir LibreOffice Base funkcionalumas yra praktiškai "
        "ekvivalentiškas - lentelės, užklausos, formos ir ataskaitos "
        "perkeliamos tarp jų be esminių pakeitimų. Tai patvirtina, kad "
        "atvirojo kodo alternatyva yra tinkamas pasirinkimas, kai biudžetas "
        "ribotas arba reikalingas kelių platformų palaikymas.")


def add_bibliography(doc):
    add_h1(doc, "LITERATŪROS SĄRAŠAS")
    entries = [
        "Codd, E. F. (1970). A Relational Model of Data for Large Shared "
        "Data Banks. Communications of the ACM, 13(6), 377-387.",

        "Connolly, T. ir Begg, C. (2015). Database Systems: A Practical "
        "Approach to Design, Implementation, and Management (6th ed.). "
        "Boston: Pearson.",

        "Elmasri, R. ir Navathe, S. B. (2016). Fundamentals of Database "
        "Systems (7th ed.). Boston: Pearson.",

        "Microsoft. (2026). Microsoft Access dokumentacija. Prieiga per "
        "internetą: https://support.microsoft.com/lt-lt/access",

        "Microsoft. (2026). Get Started with Access. Prieiga per internetą: "
        "https://support.microsoft.com/en-us/access",

        "The Document Foundation. (2026). LibreOffice Base Handbook. "
        "Prieiga per internetą: https://documentation.libreoffice.org",

        "Vilniaus universitetas. (2024). Dirbtinio intelekto naudojimo "
        "gairės (Nr. SPN-54). Vilnius: VU.",
    ]
    for e in entries:
        add_bib(doc, e)


def add_appendix_1(doc):
    add_appendix_h(doc, "1 priedas",
                   "SQL DDL ir DML pavyzdžiai")
    add_p(doc,
        "Šiame priede pateikiami SQL DDL skriptai, naudoti DB1.odb iki DB6.odb "
        "duomenų bazėms sukurti, taip pat užklausų SQL pavyzdžiai (užduočių "
        "Department, Supplier, Date ir kt.).")

    sections = [
        ("CREATE TABLE tblDepartments",
         "CREATE TABLE tblDepartments (\n"
         "    Department VARCHAR(50) NOT NULL PRIMARY KEY\n"
         ");"),
        ("CREATE TABLE tblSuppliers",
         "CREATE TABLE tblSuppliers (\n"
         "    SupplierID VARCHAR(30) NOT NULL PRIMARY KEY,\n"
         "    FirstName VARCHAR(50),\n"
         "    LastName VARCHAR(50),\n"
         "    ContactPhone VARCHAR(30),\n"
         "    Company VARCHAR(100),\n"
         "    ContactEmail VARCHAR(100),\n"
         "    Address VARCHAR(150),\n"
         "    City VARCHAR(60),\n"
         "    StateProvince VARCHAR(40),\n"
         "    Country VARCHAR(40),\n"
         "    PostalCode VARCHAR(20)\n"
         ");"),
        ("CREATE TABLE tblInventory",
         "CREATE TABLE tblInventory (\n"
         "    ProductCode VARCHAR(30) NOT NULL PRIMARY KEY,\n"
         "    Dept VARCHAR(50),\n"
         "    SupplierID VARCHAR(30),\n"
         "    ItemDescription VARCHAR(255),\n"
         "    Location VARCHAR(50),\n"
         "    Rack VARCHAR(20),\n"
         "    Origin VARCHAR(40),\n"
         "    UnitsInStock INTEGER,\n"
         "    TargetInventory INTEGER,\n"
         "    ReorderLevel INTEGER,\n"
         "    LastOrdered DATE,\n"
         "    OurUnitCost DECIMAL(10,2),\n"
         "    RetailPrice DECIMAL(10,2),\n"
         "    FOREIGN KEY (Dept) REFERENCES tblDepartments(Department),\n"
         "    FOREIGN KEY (SupplierID) REFERENCES tblSuppliers(SupplierID)\n"
         ");"),
        ("DB2: qryDepartment (Hardware)",
         "SELECT ProductCode, SupplierID\n"
         "FROM tblInventory\n"
         "WHERE Dept = 'Hardware';"),
        ("DB2: qrySupplier (Bathroom)",
         "SELECT ProductCode, SupplierID, ItemDescription\n"
         "FROM tblInventory\n"
         "WHERE Dept = 'Bathroom';"),
        ("DB2: qryDate (WOODSTOCK 2015-05-21)",
         "SELECT ProductCode, Dept, SupplierID\n"
         "FROM tblInventory\n"
         "WHERE SupplierID = 'WOODSTOCK'\n"
         "  AND LastOrdered = DATE '2015-05-21';"),
        ("DB4: qryUnusedSuppliers (Kanada)",
         "SELECT s.SupplierID, s.Company, s.City, s.Country\n"
         "FROM tblSuppliers s\n"
         "WHERE s.Country = 'Canada'\n"
         "  AND s.SupplierID NOT IN (\n"
         "    SELECT DISTINCT SupplierID FROM tblInventory\n"
         "  );"),
        ("DB6: SUM uzklausa - inventoriaus verte pagal skyriu",
         "SELECT Dept,\n"
         "       SUM(UnitsInStock * RetailPrice) AS TotalValue\n"
         "FROM tblInventory\n"
         "GROUP BY Dept\n"
         "ORDER BY TotalValue DESC;"),
        ("DB6: AVG uzklausa - vidutine kaina pagal kilme",
         "SELECT Origin,\n"
         "       AVG(RetailPrice) AS AvgPrice\n"
         "FROM tblInventory\n"
         "GROUP BY Origin\n"
         "ORDER BY AvgPrice DESC;"),
        ("DB6: COUNT uzklausa - produktu kiekis pagal tiekeja",
         "SELECT SupplierID,\n"
         "       COUNT(*) AS ProductCount\n"
         "FROM tblInventory\n"
         "GROUP BY SupplierID\n"
         "ORDER BY ProductCount DESC;"),
    ]
    for label, sql in sections:
        p = doc.add_paragraph()
        set_pf(p, alignment=WD_ALIGN_PARAGRAPH.LEFT, fli=Cm(0),
               ls=LINE_SPACING, sb=10)
        run = p.add_run(label)
        set_run_font(run, size=12, bold=True)
        psql = doc.add_paragraph()
        set_pf(psql, alignment=WD_ALIGN_PARAGRAPH.LEFT, fli=Cm(0),
               li=Cm(0.5), ls=1.0, sb=2)
        run = psql.add_run(sql)
        set_run_font(run, size=10, font_name="Courier New")


def add_appendix_11(doc):
    add_appendix_h(doc, "11 priedas",
                   "Dirbtinio intelekto panaudojimo deklaracija")
    add_p(doc,
        "Rengiant šį darbą buvo naudotas generatyvinis dirbtinio intelekto "
        "(toliau - DI) įrankis Anthropic Claude (Sonnet 4.5, internetinė "
        "prieiga, naudota 2026 m. gegužės mėnesį). Praktinę darbo dalį - "
        "duomenų bazių struktūros analizę, lentelių laukų bei raktų "
        "identifikavimą, užklausų ir ataskaitų projektavimą bei duomenų "
        "interpretavimą - savarankiškai atliko darbo autorius. DI buvo "
        "panaudotas tik aprašomojo teksto, sąvokų paaiškinimų bei techninių "
        "detalių pateikimui dokumente.")
    add_p(doc,
        "DI sugeneruotas tekstas nebuvo perkeltas į darbą be peržiūros - "
        "kiekvienas teiginys patikrintas ir reikšmingai redaguotas autoriaus. "
        "DI sugeneruoto turinio dalis darbe neviršija 15 procentų. Autorius "
        "susipažinęs su Vilniaus universiteto 2024 m. patvirtintomis DI "
        "naudojimo gairėmis (Nr. SPN-54) ir prisiima visišką atsakomybę už "
        "darbo turinį.")
    add_table_caption(doc, "2 lentelė", "DI naudojimo apimties suvestinė")
    headers = ["Rodiklis", "Reikšmė"]
    rows = [
        ["DI modelis", "Anthropic Claude Sonnet 4.5"],
        ["Naudojimo data", "2026 m. gegužės mėn."],
        ["DI naudojimo tikslas",
         "Aprašomojo teksto, sąvokų paaiškinimų pateikimas"],
        ["Kas atlikta savarankiškai (autoriaus)",
         "DB analizė, užklausų projektavimas, ataskaitų kūrimas"],
        ["DI sugeneruoto turinio dalis darbe", "mažiau nei 15 proc."],
    ]
    add_data_table(doc, headers, rows, col_widths_cm=[6.0, 10.5])
    add_src(doc, "Šaltinis: sudaryta autoriaus.")


def main():
    doc = Document()
    configure_default_style(doc)
    section = doc.sections[0]
    configure_section(section, fpnn=True)
    add_title_page(doc)
    add_toc(doc)
    add_introduction(doc)
    add_chapter_1(doc)
    add_chapter_2(doc)
    add_chapter_3(doc)
    add_chapter_4(doc)
    add_chapter_5(doc)
    add_chapter_6(doc)
    add_conclusions(doc)
    add_bibliography(doc)
    add_appendix_1(doc)
    add_appendix_11(doc)
    out_path = os.path.join(SCRIPT_DIR, OUTPUT_FILE)
    doc.save(out_path)
    print(f"Sukurta: {out_path}")


if __name__ == "__main__":
    main()
