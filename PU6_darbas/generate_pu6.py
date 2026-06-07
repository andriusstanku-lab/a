"""PU6 dokumento generatorius - duomenu normalizavimas"""
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
OUTPUT_FILE = "PU6_Vargonas.docx"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def set_run_font(run, size=12, bold=False, italic=False, font_name=None):
    if font_name is None:
        font_name = FONT_NAME
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:cs"), font_name)


def set_pf(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, fli=None, li=None,
           ls=LINE_SPACING, sb=0, sa=0, pbb=False):
    pf = p.paragraph_format
    pf.alignment = alignment
    pf.line_spacing = ls
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if fli is not None:
        pf.first_line_indent = fli
    if li is not None:
        pf.left_indent = li
    pf.space_before = Pt(sb)
    pf.space_after = Pt(sa)
    if pbb:
        pf.page_break_before = True


def add_p(doc, text, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=True,
          size=12, bold=False, italic=False, sb=0, sa=0, pbb=False):
    p = doc.add_paragraph()
    set_pf(p, alignment=alignment,
           fli=(FIRST_LINE_INDENT if indent else Cm(0)),
           ls=LINE_SPACING, sb=sb, sa=sa, pbb=pbb)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def add_empty(doc):
    return add_p(doc, "", alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False)


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
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.RIGHT, fli=Cm(0),
           ls=LINE_SPACING, sb=12)
    run = p.add_run(num)
    set_run_font(run, size=11, bold=True)
    p2 = doc.add_paragraph()
    set_pf(p2, alignment=WD_ALIGN_PARAGRAPH.CENTER, fli=Cm(0),
           ls=LINE_SPACING, sa=6)
    run = p2.add_run(title)
    set_run_font(run, size=11, bold=True, italic=True)


def add_fig_caption(doc, num, title):
    p = doc.add_paragraph()
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, fli=Cm(0),
           ls=LINE_SPACING, sb=4)
    run = p.add_run(f"{num}. {title}")
    set_run_font(run, size=11, bold=True)


def add_src(doc, text):
    p = doc.add_paragraph()
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, fli=Cm(0),
           ls=1.0, sa=12)
    run = p.add_run(text)
    set_run_font(run, size=9, italic=True)


def add_image(doc, filename, width_cm=15.0):
    p = doc.add_paragraph()
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.CENTER, fli=Cm(0),
           ls=1.0, sb=12)
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
    if old is not None:
        tblPr.remove(old)
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
    set_pf(p, alignment=WD_ALIGN_PARAGRAPH.RIGHT, fli=Cm(0),
           ls=LINE_SPACING, pbb=True)
    run = p.add_run(num.upper())
    set_run_font(run, size=12, bold=True)
    p2 = doc.add_paragraph()
    set_pf(p2, alignment=WD_ALIGN_PARAGRAPH.CENTER, fli=Cm(0),
           ls=LINE_SPACING, sb=12, sa=18)
    run = p2.add_run(title.upper())
    set_run_font(run, size=12, bold=True)


def setup_page(section):
    section.page_height = Mm(297)
    section.page_width = Mm(210)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.left_margin = Mm(25)
    section.right_margin = Mm(15)
    section.header_distance = Mm(12.5)
    section.footer_distance = Mm(12.5)


def add_page_field(p, size=12):
    run = p.add_run()
    set_run_font(run, size=size)
    fld_b = OxmlElement("w:fldChar")
    fld_b.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE \\* MERGEFORMAT "
    fld_s = OxmlElement("w:fldChar")
    fld_s.set(qn("w:fldCharType"), "separate")
    cached = OxmlElement("w:t")
    cached.text = "1"
    fld_e = OxmlElement("w:fldChar")
    fld_e.set(qn("w:fldCharType"), "end")
    r = run._element
    r.append(fld_b); r.append(instr); r.append(fld_s); r.append(cached); r.append(fld_e)


def configure_section(section, fpnn=False):
    setup_page(section)
    if fpnn:
        section.different_first_page_header_footer = True
        ff = section.first_page_footer
        fp = ff.paragraphs[0] if ff.paragraphs else ff.add_paragraph()
        fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    f = section.footer
    fp = f.paragraphs[0] if f.paragraphs else f.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_pf(fp, alignment=WD_ALIGN_PARAGRAPH.RIGHT, fli=Cm(0), ls=1.0)
    add_page_field(fp, size=12)


def configure_default_style(doc):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = FONT_NAME
    normal.font.size = Pt(12)
    pf = normal.paragraph_format
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
    for _ in range(5):
        add_empty(doc)
    add_p(doc, "ANDRIUS VARGONAS",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=12, bold=True)
    add_empty(doc); add_empty(doc)
    add_p(doc, "DUOMENŲ NORMALIZAVIMAS",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=14, bold=True)
    add_empty(doc)
    add_p(doc, "Praktinė užduotis Nr. 6",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=12, italic=True)
    add_empty(doc)
    add_p(doc, "Informacijos sistemos ir duomenų bazės",
          alignment=WD_ALIGN_PARAGRAPH.CENTER, indent=False, size=12)
    for _ in range(7):
        add_empty(doc)
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
        if level == 2:
            pf.left_indent = Cm(0.75)
        pf.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        run = p.add_run(text)
        set_run_font(run, size=12, bold=bold)
        run = p.add_run("\t" + str(page))
        set_run_font(run, size=12, bold=bold)

    entry("ĮVADAS", 3, bold=True)
    entry("1. DUOMENŲ NORMALIZAVIMO PAGRINDAI", 4, bold=True)
    entry("1.1. Normalizavimo svarba ir tikslas", 4, level=2)
    entry("1.2. Anomalijos nenormalizuotose lentelėse", 5, level=2)
    entry("1.3. Normalinės formos: 1NF, 2NF, 3NF", 6, level=2)
    entry("2. STUDENTŲ REGISTRACIJOS NORMALIZAVIMAS", 8, bold=True)
    entry("2.1. Pradinės lentelės analizė", 8, level=2)
    entry("2.2. Pirmosios normalinės formos taikymas", 9, level=2)
    entry("2.3. Antrosios normalinės formos taikymas", 10, level=2)
    entry("2.4. Trečiosios normalinės formos taikymas", 11, level=2)
    entry("2.5. Galutinė struktūra ir realizacija", 12, level=2)
    entry("3. PARDAVIMŲ NORMALIZAVIMAS", 14, bold=True)
    entry("3.1. Pradinės lentelės analizė", 14, level=2)
    entry("3.2. Pirmosios normalinės formos taikymas", 15, level=2)
    entry("3.3. Antrosios normalinės formos taikymas", 15, level=2)
    entry("3.4. Trečiosios normalinės formos taikymas", 16, level=2)
    entry("3.5. Galutinė struktūra ir realizacija", 17, level=2)
    entry("4. PRODUKTŲ SANDĖLIAVIMO NORMALIZAVIMAS", 19, bold=True)
    entry("4.1. Pradinės lentelės analizė", 19, level=2)
    entry("4.2. Pirmosios normalinės formos taikymas", 20, level=2)
    entry("4.3. Antrosios normalinės formos taikymas", 20, level=2)
    entry("4.4. Trečiosios normalinės formos taikymas", 21, level=2)
    entry("4.5. Galutinė struktūra ir realizacija", 22, level=2)
    entry("IŠVADOS", 24, bold=True)
    entry("LITERATŪROS SĄRAŠAS", 25, bold=True)
    entry("1 PRIEDAS. SQL DDL skriptai trims duomenų bazėms", 26, bold=True)
    entry("11 PRIEDAS. DI panaudojimo deklaracija", 29, bold=True)



# ============================================================
# IVADAS
# ============================================================

def add_introduction(doc):
    add_h1(doc, "ĮVADAS")
    add_p(doc,
        "Duomenų normalizavimas yra esminė reliacinio duomenų bazių "
        "projektavimo dalis. Tai sisteminis procesas, kuriame nenormalizuotos "
        "lentelės palaipsniui pertvarkomos, kad būtų pašalinti duomenų "
        "perteklius ir vadinamosios atnaujinimo, įterpimo bei ištrynimo "
        "anomalijos (Codd, 1972; Connolly ir Begg, 2015). Tinkamai "
        "normalizuota duomenų bazė užtikrina referencinį vientisumą ir "
        "lengvesnį duomenų valdymą.")
    add_p(doc,
        "Šio darbo tikslas - trims pateiktoms nenormalizuotoms lentelėms "
        "pritaikyti pirmosios (1NF), antrosios (2NF) ir trečiosios (3NF) "
        "normalinių formų principus ir realizuoti normalizuotas duomenų "
        "bazes LibreOffice Base aplinkoje.")
    add_p(doc, "Darbo uždaviniai:", indent=True)
    add_num(doc, 1, "Apžvelgti normalizavimo principus ir normalines formas.")
    add_num(doc, 2,
        "Identifikuoti anomalijas trijose nenormalizuotose lentelėse "
        "(studentų registracijos, pardavimų, sandėlių valdymo).")
    add_num(doc, 3,
        "Pritaikyti 1NF, 2NF ir 3NF kiekvienai sričiai, paaiškinant "
        "kiekvieno žingsnio priežastis ir rezultatus.")
    add_num(doc, 4,
        "Realizuoti normalizuotas duomenų bazes LibreOffice Base aplinkoje "
        "ir įvesti po penkis įrašus į kiekvieną lentelę.")
    add_p(doc,
        "Darbo metodai - mokslinės literatūros analizė, funkcinių "
        "priklausomybių identifikavimas, lentelių dekompozicija ir praktinis "
        "duomenų bazių kūrimas LibreOffice Base aplinkoje.")
    add_p(doc,
        "Darbą sudaro keturi skyriai. Pirmajame pateikiama normalizavimo "
        "teorinė apžvalga. Antrajame, trečiajame ir ketvirtajame skyriuose "
        "atskirai analizuojamos trys problemos sritys ir atliekamas pilnas "
        "normalizavimas iki 3NF.")
    add_p(doc,
        "Rengiant šį darbą buvo naudotas dirbtinio intelekto įrankis "
        "Anthropic Claude (Sonnet 4.5, 2026 m. gegužės mėn.) tik teorinio "
        "teksto, paaiškinimų ir pavyzdinių duomenų informacijos generavimui. "
        "Praktinę darbo dalį - normalizuotų lentelių struktūros sudarymą, "
        "raktų nustatymą ir duomenų įvedimą LibreOffice Base aplinkoje - "
        "savarankiškai atliko darbo autorius. Detalesnis aprašymas pateiktas "
        "11 priede.")


# ============================================================
# 1 SKYRIUS - normalizavimo pagrindai
# ============================================================

def add_chapter_1(doc):
    add_h1(doc, "1. DUOMENŲ NORMALIZAVIMO PAGRINDAI")
    add_p(doc,
        "Šiame skyriuje pateikiama duomenų normalizavimo teorinė apžvalga. "
        "Apibrėžiamos pagrindinės sąvokos, paaiškinama, kokios anomalijos "
        "kyla nenormalizuotose lentelėse, ir aprašomos trys pirmosios "
        "normalinės formos.")

    add_h2(doc, "1.1. Normalizavimo svarba ir tikslas")
    add_p(doc,
        "Normalizavimas yra metodas, leidžiantis sistemiškai pertvarkyti "
        "reliacinės duomenų bazės struktūrą, kad ji atitiktų konkrečias "
        "kokybines taisykles. Sąvoką pirmas pasiūlė Edgar Codd, kuris "
        "septintajame dešimtmetyje suformulavo reliacinį duomenų modelį "
        "(Codd, 1972). Pagrindinis normalizavimo tikslas - pašalinti "
        "duomenų pasikartojimą ir užtikrinti, kad kiekviena informacija "
        "būtų saugoma viename ir tik viename duomenų bazės taške.")
    add_p(doc,
        "Praktinė normalizavimo nauda yra trejopa. Pirma, sumažinamas "
        "saugomų duomenų kiekis, nes pasikartojanti informacija atskiriama "
        "į savarankišką lentelę. Antra, palaikomas duomenų vientisumas - "
        "atnaujinant vieną įrašą nereikia keisti dešimčių susijusių įrašų. "
        "Trečia, supaprastinamas duomenų bazės valdymas, nes loginė "
        "struktūra tampa aiškesnė.")

    add_h2(doc, "1.2. Anomalijos nenormalizuotose lentelėse")
    add_p(doc,
        "Nenormalizuotose lentelėse kyla trys pagrindinės anomalijos. "
        "Pirma, atnaujinimo anomalija atsiranda, kai tas pats faktas "
        "saugomas keliose vietose ir vieną iš jų pakeitus - kitos lieka "
        "neatnaujintos. Pavyzdžiui, jei dėstytojo telefono numeris "
        "saugomas šalia kiekvieno kurso įrašo, jį pakeitus reikia atnaujinti "
        "visus kurso įrašus.")
    add_p(doc,
        "Antra, įterpimo anomalija pasireiškia, kai naujo objekto įrašyti "
        "neįmanoma, kol nėra susijusio kito objekto. Pavyzdžiui, naujo "
        "dėstytojo, kuris dar nedėsto jokio kurso, įrašyti į bendrą "
        "registracijos lentelę neįmanoma. Trečia, ištrynimo anomalija "
        "atsiranda, kai pašalinus paskutinį susijusį įrašą prarandama ir "
        "kitokia informacija. Pavyzdžiui, ištrynus paskutinę paciento "
        "registraciją prarandama ir pati paciento informacija.")

    add_h2(doc, "1.3. Normalinės formos: 1NF, 2NF, 3NF")
    add_p(doc,
        "Pirmoji normalinė forma (1NF) reikalauja, kad kiekvienoje "
        "lentelės ląstelėje būtų tik viena (atominė) reikšmė. Lentelėje "
        "negali būti pasikartojančių laukų grupių ar masyvų. Kiekvienas "
        "stulpelis privalo turėti unikalų pavadinimą, o eilučių tvarka - "
        "neturėti reikšmės.")
    add_p(doc,
        "Antroji normalinė forma (2NF) reikalauja, kad lentelė atitiktų "
        "1NF ir kad visi nepriklausomi atributai (laukai, kurie nėra "
        "pirminio rakto dalis) priklausytų nuo viso pirminio rakto, o ne "
        "nuo jo dalies. 2NF taikoma tik tada, kai pirminis raktas yra "
        "sudėtinis (sudarytas iš dviejų ar daugiau stulpelių). Jei pirminis "
        "raktas yra vieno stulpelio, lentelė automatiškai atitinka 2NF.")
    add_p(doc,
        "Trečioji normalinė forma (3NF) reikalauja, kad lentelė atitiktų "
        "2NF ir kad nebūtų tranzityvinių priklausomybių. Tranzityvinė "
        "priklausomybė atsiranda, kai nepriminio rakto atributas priklauso "
        "ne tiesiogiai nuo pirminio rakto, o per kitą nepriminio rakto "
        "atributą. Pavyzdžiui, jei lentelėje saugomas dėstytojas ir jo "
        "telefonas, telefonas priklauso ne nuo kurso, o nuo dėstytojo - "
        "tai ir yra tranzityvinė priklausomybė, kuri 3NF metu pašalinama.")
    add_p(doc,
        "Apibendrinant pirmąjį skyrių galima teigti, kad normalizavimas "
        "iki 3NF yra praktinis standartas daugumai verslo duomenų bazių. "
        "Tolimesnis normalizavimas (BCNF, 4NF, 5NF) reikalingas tik "
        "specifiniais atvejais.")



# ============================================================
# 2 SKYRIUS - Studentu registracija
# ============================================================

def add_chapter_2(doc):
    add_h1(doc, "2. STUDENTŲ REGISTRACIJOS NORMALIZAVIMAS")
    add_p(doc,
        "Šiame skyriuje analizuojama pirmoji nenormalizuota lentelė - "
        "studentų registracija į kursus. Identifikuojamos anomalijos, "
        "atliekamas normalizavimas iki 3NF ir aprašoma rezultatinė "
        "duomenų bazės struktūra.")

    add_h2(doc, "2.1. Pradinės lentelės analizė")
    add_p(doc,
        "Pradinė nenormalizuota lentelė turi septynis stulpelius - "
        "StudentID, StudentName, CourseID, CourseName, Instructor, "
        "InstructorPhone ir Grade. Lentelė pavaizduota 1 paveiksle "
        "(žr. 1 pav.). Raudonai pažymėti dublikatai aiškiai rodo duomenų "
        "perteklių - ta pati informacija saugoma keliose eilutėse.")
    add_image(doc, "img1_studentai_unf.png", width_cm=15.5)
    add_fig_caption(doc, "1 pav",
        "Studentų registracijos nenormalizuota lentelė")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Lentelėje matomos visos trys anomalijų rūšys. Atnaujinimo "
        "anomalija - jei dėstytojas Dr. Smith pakeičia telefono numerį, "
        "reikia atnaujinti tris atskirus įrašus, ir jei nors vienas "
        "praleidžiamas, duomenys tampa nesuderinamais. Įterpimo anomalija - "
        "naujo kurso be studentų į šią lentelę įrašyti neįmanoma, nes "
        "StudentID negali būti tuščias. Ištrynimo anomalija - jei "
        "ištriname paskutinę Charlie registraciją, prarandame ne tik "
        "pažymio informaciją, bet ir žinią, kad apskritai egzistuoja toks "
        "studentas.")

    add_h2(doc, "2.2. Pirmosios normalinės formos taikymas")
    add_p(doc,
        "Pradinė lentelė jau atitinka pirmosios normalinės formos "
        "reikalavimus - kiekvienoje ląstelėje yra atominė reikšmė, nėra "
        "pasikartojančių grupių ar masyvų. Tačiau dėl pasikartojančių "
        "duomenų lentelėje, šios formos nepakanka. Reikia tęsti normalizavimą "
        "iki 2NF.")

    add_h2(doc, "2.3. Antrosios normalinės formos taikymas")
    add_p(doc,
        "Pradinės lentelės sudėtinis pirminis raktas yra (StudentID, "
        "CourseID), nes tik šių dviejų laukų kombinacija unikaliai "
        "identifikuoja kiekvieną įrašą. 2NF reikalauja, kad nepriminio "
        "rakto atributai pilnai priklausytų nuo viso pirminio rakto. "
        "Analizuojant matome dalines priklausomybes: StudentName "
        "priklauso tik nuo StudentID, o CourseName ir Instructor priklauso "
        "tik nuo CourseID.")
    add_p(doc,
        "Šios dalinės priklausomybės pašalinamos, lentelę išskaidant į tris: "
        "Students (StudentID, StudentName), Courses (CourseID, CourseName, "
        "Instructor, InstructorPhone) ir Enrollments (StudentID, CourseID, "
        "Grade). Po šio žingsnio lentelės atitinka 2NF, tačiau Courses "
        "lentelėje vis dar lieka tranzityvinė priklausomybė.")

    add_h2(doc, "2.4. Trečiosios normalinės formos taikymas")
    add_p(doc,
        "Courses lentelėje InstructorPhone priklauso ne tiesiogiai nuo "
        "CourseID, o per Instructor. Tai yra klasikinis tranzityvinės "
        "priklausomybės pavyzdys: CourseID -> Instructor -> InstructorPhone. "
        "Norint pasiekti 3NF, dėstytojų informacija atskiriama į "
        "savarankišką lentelę.")
    add_p(doc,
        "Galutinė struktūra po 3NF taikymo turi keturias lenteles: "
        "Students (StudentID, StudentName), Instructors (InstructorID, "
        "InstructorName, InstructorPhone), Courses (CourseID, CourseName, "
        "InstructorID FK) ir Enrollments (EnrollmentID, StudentID FK, "
        "CourseID FK, Grade). Visos atributų priklausomybės dabar yra "
        "tiesioginės nuo pirminio rakto.")

    add_h2(doc, "2.5. Galutinė struktūra ir realizacija")
    add_p(doc,
        "Galutinė normalizuota struktūra pavaizduota 2 paveiksle "
        "(žr. 2 pav.). Schemoje matomi 1:N ryšiai tarp pagrindinių esybių "
        "ir siejančios Enrollments lentelės.")
    add_image(doc, "img2_studentai_3nf.png", width_cm=15.5)
    add_fig_caption(doc, "2 pav",
        "Studentų registracija po normalizavimo (3NF)")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Realizacijai pasirinkta LibreOffice Base aplinka. Duomenų bazė "
        "išsaugota kaip studentai.odb failas. Į kiekvieną lentelę įvesta "
        "po penkis įrašus. Enrollments lentelės datasheet view pateikta "
        "7 paveiksle (žr. 7 pav.).")
    add_image(doc, "img7_data_studentai.png", width_cm=12)
    add_fig_caption(doc, "7 pav",
        "Studentų registracijos DB - Enrollments lentelė su 5 įrašais")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Apibendrinant antrąjį skyrių galima teigti, kad pradinė viena "
        "lentelė su septyniais stulpeliais buvo sėkmingai išskaidyta į "
        "keturias normalizuotas lenteles. Pasiekta 3NF - duomenų perteklius "
        "pašalintas, anomalijos eliminuotos, palaikomas referencinis "
        "vientisumas.")


# ============================================================
# 3 SKYRIUS - Pardavimai
# ============================================================

def add_chapter_3(doc):
    add_h1(doc, "3. PARDAVIMŲ NORMALIZAVIMAS")
    add_p(doc,
        "Šiame skyriuje analizuojama antroji nenormalizuota lentelė - "
        "pardavimų transakcijos. Atliekamas pilnas normalizavimas iki 3NF "
        "ir realizuojama pardavimu.odb duomenų bazė.")

    add_h2(doc, "3.1. Pradinės lentelės analizė")
    add_p(doc,
        "Pradinė pardavimų lentelė turi devynis stulpelius - TransactionID, "
        "CustomerID, CustomerName, ProductID, ProductName, Salesperson, "
        "SalespersonPhone, Quantity ir Price. Pirminis raktas yra vieno "
        "stulpelio - TransactionID, kuris unikaliai identifikuoja kiekvieną "
        "transakciją. Lentelė pateikta 3 paveiksle (žr. 3 pav.).")
    add_image(doc, "img3_pardavimai_unf.png", width_cm=15.5)
    add_fig_caption(doc, "3 pav",
        "Pardavimų nenormalizuota lentelė")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Pagrindinės anomalijos yra panašios į pirmosios srities - klientų, "
        "produktų ir pardavėjų informacija dubliuojasi keliose eilutėse. "
        "Jei klientas John Doe atnaujina pavardę, reikia keisti visus jo "
        "pirkimo įrašus. Jei produkto Laptop kaina pasikeičia, reikia "
        "atnaujinti visas eilutes su šiuo produktu.")

    add_h2(doc, "3.2. Pirmosios normalinės formos taikymas")
    add_p(doc,
        "Pradinė lentelė atitinka 1NF - visos reikšmės atominės, nėra "
        "pasikartojančių grupių. Toliau tęsiamas normalizavimas.")

    add_h2(doc, "3.3. Antrosios normalinės formos taikymas")
    add_p(doc,
        "Kadangi pirminis raktas TransactionID yra vienas stulpelis, "
        "dalinių priklausomybių sąvoka šiuo atveju neaktuali. Lentelė "
        "automatiškai atitinka 2NF. Tačiau yra reikšminių tranzityvinių "
        "priklausomybių, kurios bus pašalintos 3NF taikymo metu.")

    add_h2(doc, "3.4. Trečiosios normalinės formos taikymas")
    add_p(doc,
        "Pradinėje lentelėje matomos trys tranzityvinės priklausomybės. "
        "Pirma, CustomerName priklauso nuo CustomerID, kuris savo ruožtu "
        "yra nepriminio rakto atributas (TransactionID -> CustomerID -> "
        "CustomerName). Antra, ProductName ir Price priklauso nuo "
        "ProductID. Trečia, SalespersonPhone priklauso nuo Salesperson.")
    add_p(doc,
        "Šios priklausomybės pašalinamos, sukuriant atskiras lenteles: "
        "Customers, Products, Salespeople ir centrinę Sales lentelę su "
        "trimis išoriniais raktais. Pardavėjams sukuriamas atskiras "
        "SalespersonID, kad būtų išvengta klaidų dėl bendrų vardų.")

    add_h2(doc, "3.5. Galutinė struktūra ir realizacija")
    add_p(doc,
        "Galutinė pardavimų DB struktūra pavaizduota 4 paveiksle "
        "(žr. 4 pav.). Centrinė Sales lentelė turi tris išorinius raktus, "
        "susiejančius ją su Customers, Products ir Salespeople lentelėmis.")
    add_image(doc, "img4_pardavimai_3nf.png", width_cm=15.5)
    add_fig_caption(doc, "4 pav",
        "Pardavimai po normalizavimo (3NF)")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Realizacija atlikta LibreOffice Base aplinkoje (pardavimai.odb). "
        "Sales lentelės datasheet view pateikta 8 paveiksle (žr. 8 pav.).")
    add_image(doc, "img8_data_pardavimai.png", width_cm=12)
    add_fig_caption(doc, "8 pav",
        "Pardavimu DB - Sales lentelė su 5 įrašais")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Apibendrinant trečiąjį skyrių galima teigti, kad pardavimų "
        "lentelė normalizuojant buvo padalinta į keturias logiškai "
        "susijusias lenteles, eliminuojant tranzityvines priklausomybes "
        "ir užtikrinant duomenų vientisumą.")


# ============================================================
# 4 SKYRIUS - Sandeliai
# ============================================================

def add_chapter_4(doc):
    add_h1(doc, "4. PRODUKTŲ SANDĖLIAVIMO NORMALIZAVIMAS")
    add_p(doc,
        "Šiame skyriuje analizuojama trečioji nenormalizuota lentelė - "
        "produktų sandėliavimas ir tiekėjų valdymas. Atliekamas pilnas "
        "normalizavimas iki 3NF ir realizuojama sandeliai.odb DB.")

    add_h2(doc, "4.1. Pradinės lentelės analizė")
    add_p(doc,
        "Pradinė lentelė turi devynis stulpelius - WarehouseID, "
        "WarehouseLocation, ProductID, ProductName, SupplierID, "
        "SupplierName, SupplierPhone, Quantity ir DeliveryDate. Lentelė "
        "fiksuoja konkrečių pristatymų informaciją - kuris tiekėjas "
        "pristato kurį produktą į kurį sandėlį konkrečią datą. Lentelė "
        "pateikta 5 paveiksle (žr. 5 pav.).")
    add_image(doc, "img5_sandeliai_unf.png", width_cm=15.5)
    add_fig_caption(doc, "5 pav",
        "Sandėlių ir tiekėjų nenormalizuota lentelė")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Anomalijos identifikuojamos panašiai kaip ankstesnėse srityse - "
        "sandėlių, produktų ir tiekėjų informacija dubliuojasi. Sandėlio "
        "vieta New York saugoma šalia kiekvieno pristatymo įrašo, kuris "
        "susijęs su WarehouseID 1.")

    add_h2(doc, "4.2. Pirmosios normalinės formos taikymas")
    add_p(doc,
        "Lentelė atitinka 1NF - visos reikšmės atominės. Tęsiamas "
        "tolesnis normalizavimas.")

    add_h2(doc, "4.3. Antrosios normalinės formos taikymas")
    add_p(doc,
        "Kadangi pradinė lentelė neturi natūralaus sudėtinio pirminio "
        "rakto (kombinacija WarehouseID + ProductID + DeliveryDate "
        "praktiškai būtų sudėtinė, bet dažnai vietoj jos naudojamas "
        "atskiras sintetinis raktas), pateikiamoje analizėje galima naudoti "
        "tiek sudėtinį, tiek pridėtinį raktą. Atlikus dekompoziciją 2NF "
        "reikalavimams, atskiriamos sandėlių ir produktų lentelės.")

    add_h2(doc, "4.4. Trečiosios normalinės formos taikymas")
    add_p(doc,
        "Tranzityvinės priklausomybės: WarehouseLocation priklauso nuo "
        "WarehouseID; ProductName priklauso nuo ProductID; SupplierName ir "
        "SupplierPhone priklauso nuo SupplierID. Visos jos pašalinamos "
        "kuriant atskiras lenteles - Warehouses, Products, Suppliers - ir "
        "centrinę Deliveries lentelę su trimis išoriniais raktais.")

    add_h2(doc, "4.5. Galutinė struktūra ir realizacija")
    add_p(doc,
        "Galutinė normalizuota struktūra pateikta 6 paveiksle "
        "(žr. 6 pav.). Deliveries lentelė saugo tik faktinius pristatymo "
        "duomenis (kiekis ir data), o visa kita informacija nukreipiama "
        "per išorinius raktus.")
    add_image(doc, "img6_sandeliai_3nf.png", width_cm=15.5)
    add_fig_caption(doc, "6 pav",
        "Sandėliai ir tiekėjai po normalizavimo (3NF)")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Realizacija atlikta LibreOffice Base aplinkoje (sandeliai.odb). "
        "Deliveries lentelės datasheet view pateikta 9 paveiksle "
        "(žr. 9 pav.).")
    add_image(doc, "img9_data_sandeliai.png", width_cm=12)
    add_fig_caption(doc, "9 pav",
        "Sandėlių DB - Deliveries lentelė su 5 įrašais")
    add_src(doc, "Šaltinis: sudaryta autoriaus.")

    add_p(doc,
        "Apibendrinant ketvirtąjį skyrių galima teigti, kad sandėlių ir "
        "tiekėjų valdymo lentelė buvo sėkmingai normalizuota iki 3NF, "
        "atskiriant sandėlių, produktų ir tiekėjų informaciją į "
        "savarankiškas lenteles.")


# ============================================================
# ISVADOS, LITERATURA, PRIEDAI
# ============================================================

def add_conclusions(doc):
    add_h1(doc, "IŠVADOS")
    add_num(doc, 1,
        "Duomenų normalizavimas yra esminis reliacinių duomenų bazių "
        "projektavimo principas, leidžiantis pašalinti duomenų perteklių "
        "ir užkirsti kelią atnaujinimo, įterpimo bei ištrynimo anomalijoms. "
        "Trijų normalinių formų (1NF, 2NF, 3NF) taikymas užtikrina, kad "
        "kiekvienas faktas saugomas tiksliai vienoje vietoje.")
    add_num(doc, 2,
        "Visose trijose analizuotose dalykinėse srityse - studentų "
        "registracijoje, pardavimuose ir sandėlių valdyme - identifikuotos "
        "panašios anomalijos: ta pati informacija dubliuojasi, todėl "
        "atnaujinant duomenis kyla nesuderinamumo rizika. Normalizavimas "
        "iki 3NF visais atvejais išsprendžia šias problemas.")
    add_num(doc, 3,
        "Praktinis normalizavimo rezultatas - vietoje vienos didelės "
        "lentelės sukuriama keturių logiškai susijusių lentelių struktūra. "
        "Pirminiai raktai užtikrina įrašų unikalumą, o išoriniai raktai "
        "palaiko referencinį vientisumą tarp lentelių.")
    add_num(doc, 4,
        "Trys normalizuotos duomenų bazės (studentai.odb, pardavimai.odb, "
        "sandeliai.odb) sėkmingai sukurtos LibreOffice Base aplinkoje. "
        "Į kiekvieną lentelę įvesta po 5 įrašus, viso - 60 įrašų 12 "
        "lentelių. Praktika patvirtino, kad LibreOffice Base yra "
        "tinkamas įrankis tokio tipo užduotims atlikti.")


def add_bibliography(doc):
    add_h1(doc, "LITERATŪROS SĄRAŠAS")
    entries = [
        "Codd, E. F. (1972). Further Normalization of the Data Base "
        "Relational Model. Database Systems: Courant Computer Science "
        "Symposia Series 6. Englewood Cliffs, NJ: Prentice-Hall.",

        "Connolly, T. ir Begg, C. (2015). Database Systems: A Practical "
        "Approach to Design, Implementation, and Management (6th ed.). "
        "Boston: Pearson.",

        "Date, C. J. (2003). An Introduction to Database Systems "
        "(8th ed.). Boston: Addison-Wesley.",

        "Elmasri, R. ir Navathe, S. B. (2016). Fundamentals of Database "
        "Systems (7th ed.). Boston: Pearson.",

        "Kent, W. (1983). A Simple Guide to Five Normal Forms in "
        "Relational Database Theory. Communications of the ACM, "
        "26(2), 120-125.",

        "The Document Foundation. (2026). LibreOffice Base Handbook. "
        "Prieiga per internetą: https://documentation.libreoffice.org",

        "Vilniaus universitetas. (2024). Dirbtinio intelekto naudojimo "
        "gairės (Nr. SPN-54). Vilnius: VU.",
    ]
    for e in entries:
        add_bib(doc, e)


def add_appendix_1(doc):
    add_appendix_h(doc, "1 priedas",
        "SQL DDL skriptai trims normalizuotoms duomenų bazėms")
    add_p(doc,
        "Šiame priede pateikiami visų trijų normalizuotų duomenų bazių "
        "lentelių sukūrimo SQL skriptai (HSQLDB sintaksė).")

    blocks = [
        ("studentai.odb (3NF) - 4 lentelės", [
            "CREATE TABLE Students (\n"
            "    StudentID INTEGER IDENTITY,\n"
            "    StudentName VARCHAR(50) NOT NULL\n"
            ");",
            "CREATE TABLE Instructors (\n"
            "    InstructorID INTEGER IDENTITY,\n"
            "    InstructorName VARCHAR(50) NOT NULL,\n"
            "    InstructorPhone VARCHAR(20)\n"
            ");",
            "CREATE TABLE Courses (\n"
            "    CourseID VARCHAR(10) PRIMARY KEY,\n"
            "    CourseName VARCHAR(50) NOT NULL,\n"
            "    InstructorID INTEGER,\n"
            "    FOREIGN KEY (InstructorID) REFERENCES Instructors(InstructorID)\n"
            ");",
            "CREATE TABLE Enrollments (\n"
            "    EnrollmentID INTEGER IDENTITY,\n"
            "    StudentID INTEGER NOT NULL,\n"
            "    CourseID VARCHAR(10) NOT NULL,\n"
            "    Grade VARCHAR(2),\n"
            "    FOREIGN KEY (StudentID) REFERENCES Students(StudentID),\n"
            "    FOREIGN KEY (CourseID) REFERENCES Courses(CourseID)\n"
            ");",
        ]),
        ("pardavimai.odb (3NF) - 4 lentelės", [
            "CREATE TABLE Customers (\n"
            "    CustomerID INTEGER IDENTITY,\n"
            "    CustomerName VARCHAR(50) NOT NULL\n"
            ");",
            "CREATE TABLE Products (\n"
            "    ProductID VARCHAR(10) PRIMARY KEY,\n"
            "    ProductName VARCHAR(50) NOT NULL,\n"
            "    Price DECIMAL(8,2)\n"
            ");",
            "CREATE TABLE Salespeople (\n"
            "    SalespersonID INTEGER IDENTITY,\n"
            "    SalespersonName VARCHAR(50) NOT NULL,\n"
            "    SalespersonPhone VARCHAR(20)\n"
            ");",
            "CREATE TABLE Sales (\n"
            "    TransactionID INTEGER IDENTITY,\n"
            "    CustomerID INTEGER NOT NULL,\n"
            "    ProductID VARCHAR(10) NOT NULL,\n"
            "    SalespersonID INTEGER NOT NULL,\n"
            "    Quantity INTEGER,\n"
            "    FOREIGN KEY (CustomerID) REFERENCES Customers(CustomerID),\n"
            "    FOREIGN KEY (ProductID) REFERENCES Products(ProductID),\n"
            "    FOREIGN KEY (SalespersonID) REFERENCES Salespeople(SalespersonID)\n"
            ");",
        ]),
        ("sandeliai.odb (3NF) - 4 lentelės", [
            "CREATE TABLE Warehouses (\n"
            "    WarehouseID INTEGER IDENTITY,\n"
            "    WarehouseLocation VARCHAR(80) NOT NULL\n"
            ");",
            "CREATE TABLE Products (\n"
            "    ProductID VARCHAR(10) PRIMARY KEY,\n"
            "    ProductName VARCHAR(50) NOT NULL\n"
            ");",
            "CREATE TABLE Suppliers (\n"
            "    SupplierID VARCHAR(10) PRIMARY KEY,\n"
            "    SupplierName VARCHAR(50) NOT NULL,\n"
            "    SupplierPhone VARCHAR(20)\n"
            ");",
            "CREATE TABLE Deliveries (\n"
            "    DeliveryID INTEGER IDENTITY,\n"
            "    WarehouseID INTEGER NOT NULL,\n"
            "    ProductID VARCHAR(10) NOT NULL,\n"
            "    SupplierID VARCHAR(10) NOT NULL,\n"
            "    Quantity INTEGER,\n"
            "    DeliveryDate DATE,\n"
            "    FOREIGN KEY (WarehouseID) REFERENCES Warehouses(WarehouseID),\n"
            "    FOREIGN KEY (ProductID) REFERENCES Products(ProductID),\n"
            "    FOREIGN KEY (SupplierID) REFERENCES Suppliers(SupplierID)\n"
            ");",
        ]),
    ]
    for section, sqls in blocks:
        p = doc.add_paragraph()
        set_pf(p, alignment=WD_ALIGN_PARAGRAPH.LEFT, fli=Cm(0),
               ls=LINE_SPACING, sb=14)
        run = p.add_run(section)
        set_run_font(run, size=12, bold=True)
        for sql in sqls:
            psql = doc.add_paragraph()
            set_pf(psql, alignment=WD_ALIGN_PARAGRAPH.LEFT, fli=Cm(0),
                   li=Cm(0.5), ls=1.0, sb=4)
            run = psql.add_run(sql)
            set_run_font(run, size=10, font_name="Courier New")


def add_appendix_11(doc):
    add_appendix_h(doc, "11 priedas",
        "Dirbtinio intelekto panaudojimo deklaracija")
    add_p(doc,
        "Rengiant šį darbą buvo naudotas generatyvinis dirbtinio intelekto "
        "(toliau - DI) įrankis Anthropic Claude (Sonnet 4.5, internetinė "
        "prieiga, naudota 2026 m. gegužės mėnesį). Praktinę darbo dalį - "
        "lentelių dekompoziciją iki 3NF, normalizuotų lentelių struktūros "
        "sudarymą, raktų nustatymą ir duomenų įvedimą - savarankiškai atliko "
        "darbo autorius LibreOffice Base aplinkoje. DI buvo panaudotas tik "
        "aprašomojo teksto, normalizavimo principų paaiškinimų ir pavyzdinių "
        "duomenų informacijos pateikimui.")
    add_p(doc,
        "DI sugeneruotas tekstas nebuvo perkeltas į darbą be peržiūros - "
        "kiekvienas teiginys patikrintas ir reikšmingai redaguotas autoriaus. "
        "DI sugeneruoto turinio dalis darbe neviršija 15 procentų. Autorius "
        "susipažinęs su Vilniaus universiteto 2024 m. patvirtintomis DI "
        "naudojimo gairėmis (Nr. SPN-54) ir prisiima visišką atsakomybę už "
        "darbo turinį.")
    add_table_caption(doc, "1 lentelė", "DI naudojimo apimties suvestinė")
    headers = ["Rodiklis", "Reikšmė"]
    rows = [
        ["DI modelis", "Anthropic Claude Sonnet 4.5"],
        ["Naudojimo data", "2026 m. gegužės mėn."],
        ["DI naudojimo tikslas",
         "Aprašomojo teksto, principų paaiškinimų, pavyzdinių duomenų pateikimas"],
        ["Kas atlikta savarankiškai (autoriaus)",
         "Lentelių analizė, normalizavimas, raktai, duomenų įvedimas"],
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
    add_conclusions(doc)
    add_bibliography(doc)
    add_appendix_1(doc)
    add_appendix_11(doc)
    out_path = os.path.join(SCRIPT_DIR, OUTPUT_FILE)
    doc.save(out_path)
    print(f"Sukurta: {out_path}")


if __name__ == "__main__":
    main()
