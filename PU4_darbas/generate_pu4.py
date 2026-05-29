"""
PU4 ataskaitos generatorius (.docx)

Generuoja Word dokumenta pagal VU Kauno fakulteto Informatikos inzinerijos
krypties akademiniu rasto darbu metodinius nurodymus.

Tema: DBVS aplinka, lenteliu kurimas
Autorius: Andrius Vargonas
Studiju programa: Marketingo technologijos
Dalykas: Informacijos sistemos ir duomenu bazes
Metai: 2026

Naudojimas:
    pip install python-docx Pillow
    python generate_images.py    # pirma sukurti paveikslus
    python generate_pu4.py       # tada sugeneruoti dokumenta
"""

import os
from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor, Inches
from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_LINE_SPACING,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
)
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Konstantos
FONT_NAME = "Times New Roman"
FIRST_LINE_INDENT = Cm(1.25)
HANGING_INDENT = Cm(1.25)
LINE_SPACING = 1.5
OUTPUT_FILE = "PU4_Vargonas.docx"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


# ============================================================
# Pagalbines funkcijos
# ============================================================

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


def set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         first_line_indent=None, left_indent=None,
                         line_spacing=LINE_SPACING,
                         space_before=0, space_after=0,
                         page_break_before=False):
    pf = p.paragraph_format
    pf.alignment = alignment
    pf.line_spacing = line_spacing
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent
    if left_indent is not None:
        pf.left_indent = left_indent
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if page_break_before:
        pf.page_break_before = True


def add_paragraph(doc, text, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                  indent=True, size=12, bold=False, italic=False,
                  line_spacing=LINE_SPACING,
                  space_before=0, space_after=0,
                  page_break_before=False):
    p = doc.add_paragraph()
    set_paragraph_format(
        p, alignment=alignment,
        first_line_indent=(FIRST_LINE_INDENT if indent else Cm(0)),
        line_spacing=line_spacing,
        space_before=space_before, space_after=space_after,
        page_break_before=page_break_before,
    )
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def add_empty_line(doc, size=12):
    return add_paragraph(doc, "", alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         indent=False, size=size)


def apply_heading_style(p, level):
    pPr = p._element.get_or_add_pPr()
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        pStyle = OxmlElement("w:pStyle")
        pPr.insert(0, pStyle)
    pStyle.set(qn("w:val"), f"Heading{level}")


def add_heading_1(doc, text, page_break_before=True):
    p = doc.add_paragraph()
    set_paragraph_format(
        p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Cm(0), line_spacing=LINE_SPACING,
        space_after=12, page_break_before=page_break_before,
    )
    run = p.add_run(text.upper())
    set_run_font(run, size=14, bold=True)
    apply_heading_style(p, 1)
    return p


def add_heading_2(doc, text):
    p = doc.add_paragraph()
    set_paragraph_format(
        p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Cm(0), line_spacing=LINE_SPACING,
        space_before=24, space_after=12,
    )
    run = p.add_run(text)
    set_run_font(run, size=12, bold=True)
    apply_heading_style(p, 2)
    return p


def add_numbered_item(doc, number, text):
    p = doc.add_paragraph()
    set_paragraph_format(
        p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_indent=Cm(0), left_indent=Cm(0.5),
        line_spacing=LINE_SPACING,
    )
    run = p.add_run(f"{number}. {text}")
    set_run_font(run, size=12)
    return p


def add_bullet_item(doc, text):
    p = doc.add_paragraph()
    set_paragraph_format(
        p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_indent=Cm(0), left_indent=Cm(0.75),
        line_spacing=LINE_SPACING,
    )
    run = p.add_run("- " + text)
    set_run_font(run, size=12)
    return p


# ============================================================
# Lenteliu ir paveikslu pagalbos
# ============================================================

def add_table_caption(doc, number_text, title_text):
    p_num = doc.add_paragraph()
    set_paragraph_format(p_num, alignment=WD_ALIGN_PARAGRAPH.RIGHT,
                         first_line_indent=Cm(0),
                         line_spacing=LINE_SPACING, space_before=12)
    run = p_num.add_run(number_text)
    set_run_font(run, size=11, bold=True)

    p_title = doc.add_paragraph()
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         first_line_indent=Cm(0),
                         line_spacing=LINE_SPACING, space_after=6)
    run = p_title.add_run(title_text)
    set_run_font(run, size=11, bold=True, italic=True)


def add_figure_caption(doc, number_text, title_text):
    """Paveikslo numeris ir pavadinimas po paveikslu, centre."""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         first_line_indent=Cm(0),
                         line_spacing=LINE_SPACING, space_before=4)
    run = p.add_run(f"{number_text}. {title_text}")
    set_run_font(run, size=11, bold=True)


def add_source_note(doc, text):
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         first_line_indent=Cm(0),
                         line_spacing=1.0, space_before=0, space_after=12)
    run = p.add_run(text)
    set_run_font(run, size=9, italic=True)


def add_image(doc, filename, width_cm=15.0):
    """Iterpia paveiksla i dokumenta, centre lygiuotas."""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         first_line_indent=Cm(0),
                         line_spacing=1.0, space_before=12)
    run = p.add_run()
    img_path = os.path.join(SCRIPT_DIR, filename)
    if os.path.exists(img_path):
        run.add_picture(img_path, width=Cm(width_cm))
    else:
        run = p.add_run(f"[Paveikslo failas nerastas: {filename}]")
        set_run_font(run, size=10, italic=True)


def style_table_borders(table):
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
    n_cols = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    if col_widths_cm:
        for i, w in enumerate(col_widths_cm):
            for cell in table.columns[i].cells:
                cell.width = Cm(w)
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = ""
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                             first_line_indent=Cm(0), line_spacing=1.0)
        run = p.add_run(h)
        set_run_font(run, size=data_size, bold=True)
    for r_idx, row_data in enumerate(rows, start=1):
        row = table.rows[r_idx]
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = ""
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            align = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_format(p, alignment=align,
                                 first_line_indent=Cm(0), line_spacing=1.0)
            run = p.add_run(str(val))
            set_run_font(run, size=data_size)
    style_table_borders(table)
    return table


def add_bibliography_entry(doc, text):
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
    return p


def add_appendix_header(doc, number_text, title_text):
    p_num = doc.add_paragraph()
    set_paragraph_format(p_num, alignment=WD_ALIGN_PARAGRAPH.RIGHT,
                         first_line_indent=Cm(0),
                         line_spacing=LINE_SPACING,
                         page_break_before=True)
    run = p_num.add_run(number_text.upper())
    set_run_font(run, size=12, bold=True)

    p_title = doc.add_paragraph()
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         first_line_indent=Cm(0),
                         line_spacing=LINE_SPACING,
                         space_before=12, space_after=18)
    run = p_title.add_run(title_text.upper())
    set_run_font(run, size=12, bold=True)


# ============================================================
# Puslapiu nustatymai
# ============================================================

def setup_page(section):
    section.page_height = Mm(297)
    section.page_width = Mm(210)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.left_margin = Mm(25)
    section.right_margin = Mm(15)
    section.header_distance = Mm(12.5)
    section.footer_distance = Mm(12.5)


def add_page_number_field(paragraph, size=12):
    run = paragraph.add_run()
    set_run_font(run, size=size)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE \\* MERGEFORMAT "
    fld_separate = OxmlElement("w:fldChar")
    fld_separate.set(qn("w:fldCharType"), "separate")
    cached = OxmlElement("w:t")
    cached.text = "1"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r = run._element
    r.append(fld_begin)
    r.append(instr)
    r.append(fld_separate)
    r.append(cached)
    r.append(fld_end)


def configure_section_with_page_numbers(section, first_page_no_number=False):
    setup_page(section)
    if first_page_no_number:
        section.different_first_page_header_footer = True
        first_footer = section.first_page_footer
        if first_footer.paragraphs:
            fp = first_footer.paragraphs[0]
        else:
            fp = first_footer.add_paragraph()
        fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer = section.footer
    if footer.paragraphs:
        fp = footer.paragraphs[0]
    else:
        fp = footer.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_paragraph_format(fp, alignment=WD_ALIGN_PARAGRAPH.RIGHT,
                         first_line_indent=Cm(0), line_spacing=1.0)
    add_page_number_field(fp, size=12)


# ============================================================
# Antrastinis lapas
# ============================================================

def add_title_page(doc):
    add_paragraph(doc, "VILNIAUS UNIVERSITETAS",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=14, bold=True)
    add_paragraph(doc, "KAUNO FAKULTETAS",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=14, bold=True)
    add_empty_line(doc)
    add_paragraph(doc,
                  "SOCIALINIŲ MOKSLŲ IR TAIKOMOSIOS INFORMATIKOS INSTITUTAS",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=14)
    add_empty_line(doc)
    add_paragraph(doc, "Marketingo technologijų studijų programa",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=12)
    for _ in range(5):
        add_empty_line(doc)
    add_paragraph(doc, "ANDRIUS VARGONAS",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=12, bold=True)
    add_empty_line(doc)
    add_empty_line(doc)
    add_paragraph(doc, "DBVS APLINKA, LENTELIŲ KŪRIMAS",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=14, bold=True)
    add_empty_line(doc)
    add_paragraph(doc, "Praktinė užduotis Nr. 4",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=12, italic=True)
    add_empty_line(doc)
    add_paragraph(doc, "Informacijos sistemos ir duomenų bazės",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=12)
    for _ in range(7):
        add_empty_line(doc)
    add_paragraph(doc, "Kaunas",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=12)
    add_paragraph(doc, "2026",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=12)


# ============================================================
# Turinys (statinis)
# ============================================================

def add_toc_title(doc):
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         first_line_indent=Cm(0),
                         line_spacing=LINE_SPACING,
                         space_after=12, page_break_before=True)
    run = p.add_run("TURINYS")
    set_run_font(run, size=14, bold=True)
    return p


def add_toc_entry(doc, text, page_num, level=1, bold=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.line_spacing = LINE_SPACING
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.first_line_indent = Cm(0)
    if level == 2:
        pf.left_indent = Cm(0.75)
    elif level == 3:
        pf.left_indent = Cm(1.5)
    pf.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    run_text = p.add_run(text)
    set_run_font(run_text, size=12, bold=bold)
    run_num = p.add_run("\t" + str(page_num))
    set_run_font(run_num, size=12, bold=bold)
    return p


def add_toc(doc):
    add_toc_title(doc)
    add_toc_entry(doc, "ĮVADAS", 3, level=1, bold=True)
    add_toc_entry(doc, "1. DBVS ĮRANKIŲ APLINKOS APŽVALGA", 4, level=1, bold=True)
    add_toc_entry(doc, "1.1. Microsoft Access", 4, level=2)
    add_toc_entry(doc, "1.2. OpenOffice Base", 5, level=2)
    add_toc_entry(doc, "1.3. LibreOffice Base", 6, level=2)
    add_toc_entry(doc, "2. PAVYZDINĖS DUOMENŲ BAZĖS PROJEKTAVIMAS", 7,
                  level=1, bold=True)
    add_toc_entry(doc, "2.1. Dalykinė sritis ir lentelių struktūra", 7, level=2)
    add_toc_entry(doc, "2.2. Pirminių ir išorinių raktų schema", 8, level=2)
    add_toc_entry(doc, "2.3. Duomenų tipai ir apribojimai", 9, level=2)
    add_toc_entry(doc, "3. LENTELIŲ SUKŪRIMAS PASIRINKTUOSE ĮRANKIUOSE", 10,
                  level=1, bold=True)
    add_toc_entry(doc, "3.1. Realizacija Microsoft Access aplinkoje", 10,
                  level=2)
    add_toc_entry(doc, "3.2. Realizacija LibreOffice Base aplinkoje", 12,
                  level=2)
    add_toc_entry(doc, "3.3. Patirties palyginimas", 14, level=2)
    add_toc_entry(doc, "IŠVADOS", 15, level=1, bold=True)
    add_toc_entry(doc, "LITERATŪROS SĄRAŠAS", 16, level=1, bold=True)
    add_toc_entry(doc, "1 PRIEDAS. SQL DDL skriptai", 17, level=1, bold=True)
    add_toc_entry(doc, "2 PRIEDAS. Pavyzdiniai duomenys", 18, level=1, bold=True)
    add_toc_entry(doc, "11 PRIEDAS. Dirbtinio intelekto panaudojimo deklaracija",
                  19, level=1, bold=True)
    add_toc_entry(doc, "12 PRIEDAS. DI užklausos ir gauti atsakymai", 20,
                  level=1, bold=True)


def configure_default_style(doc):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = FONT_NAME
    normal.font.size = Pt(12)
    pf = normal.paragraph_format
    pf.line_spacing = LINE_SPACING
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.space_after = Pt(0)
    for h_level, size, bold in [
        ("Heading 1", 14, True),
        ("Heading 2", 12, True),
        ("Heading 3", 12, True),
    ]:
        try:
            style = styles[h_level]
            style.font.name = FONT_NAME
            style.font.size = Pt(size)
            style.font.bold = bold
            style.font.color.rgb = RGBColor(0, 0, 0)
        except KeyError:
            continue



# ============================================================
# IVADAS
# ============================================================

def add_introduction(doc):
    add_heading_1(doc, "ĮVADAS")

    add_paragraph(doc,
        "Duomenų bazių valdymo sistemos (toliau - DBVS) yra esminė šiuolaikinių "
        "informacinių sistemų sudedamoji dalis. Marketingo technologijų studijų "
        "kontekste praktinis darbas su skirtingomis DBVS aplinkomis padeda "
        "geriau suprasti, kaip duomenys saugomi, struktūrizuojami ir naudojami "
        "verslo sprendimams priimti. Reliacinės duomenų bazės, paremtos "
        "Codd (1970) suformuluotu modeliu, ir šiandien lieka dažniausiai "
        "naudojamu duomenų saugojimo principu Lietuvos ir užsienio įmonėse.")

    add_paragraph(doc,
        "Šio darbo tikslas - praktiškai išbandyti tris darbalaukio DBVS kūrimo "
        "įrankius (Microsoft Access, OpenOffice Base, LibreOffice Base) ir "
        "naudojant du iš jų sukurti pavyzdinę universiteto duomenų bazę su "
        "trimis tarpusavyje susijusiomis lentelėmis.")

    add_paragraph(doc, "Darbo uždaviniai:", indent=True)
    add_numbered_item(doc, 1,
        "Apžvelgti tris DBVS kūrimo įrankių aplinkas - jų vartotojo sąsajas, "
        "funkcionalumą ir lentelių kūrimo veiksmus.")
    add_numbered_item(doc, 2,
        "Suprojektuoti pavyzdinę universiteto duomenų bazę su trimis "
        "tarpusavyje susijusiomis lentelėmis (Students, Courses, Enrollments).")
    add_numbered_item(doc, 3,
        "Praktiškai sukurti lenteles dvejose pasirinktose aplinkose "
        "(Microsoft Access ir LibreOffice Base), užtikrinant tinkamą pirminių "
        "ir išorinių raktų naudojimą.")
    add_numbered_item(doc, 4,
        "Užpildyti lenteles pavyzdiniais duomenimis ir trumpai aprašyti "
        "patirtį dirbant su abiem įrankiais.")

    add_paragraph(doc,
        "Darbo metodai - mokslinės literatūros ir oficialios dokumentacijos "
        "analizė, dalykinės srities modeliavimas, praktinis lentelių kūrimas "
        "DBVS aplinkose, lyginamoji refleksija.")

    add_paragraph(doc,
        "Darbą sudaro trys skyriai. Pirmajame pateikiama trijų DBVS įrankių "
        "aplinkų apžvalga. Antrajame projektuojama pavyzdinė duomenų bazė ir "
        "aprašomos jos struktūros savybės. Trečiajame demonstruojama lentelių "
        "realizacija pasirinktose aplinkose ir lyginama vartotojo patirtis.")

    # DI deklaracija
    add_paragraph(doc,
        "Rengiant šį darbą buvo naudotasi dirbtinio intelekto (toliau - DI) "
        "generatyviniu modeliu Anthropic Claude (Sonnet 4.5, internetinė "
        "prieiga, naudota 2026 m. gegužės mėn.) kaip pagalbine priemone "
        "struktūros patikrinimui, kalbinio stiliaus tobulinimui ir paveikslų "
        "vizualiniam apipavidalinimui. Pagrindinį turinį, lentelių "
        "projektavimą, praktinį realizavimą ir vertinimus autorius parengė "
        "savarankiškai. DI sugeneruoto turinio dalis darbe neviršija "
        "penkiolikos procentų. Detalus DI naudojimo aprašymas pateikiamas "
        "11 priede, o naudotos užklausos - 12 priede. Autorius susipažinęs "
        "su Vilniaus universiteto 2024 m. patvirtintomis dirbtinio intelekto "
        "naudojimo gairėmis (Nr. SPN-54).")


# ============================================================
# 1 SKYRIUS - DBVS irankiu apzvalga
# ============================================================

def add_chapter_1(doc):
    add_heading_1(doc, "1. DBVS ĮRANKIŲ APLINKOS APŽVALGA")

    add_paragraph(doc,
        "Šiame skyriuje pateikiama trijų darbalaukio DBVS kūrimo įrankių "
        "apžvalga - Microsoft Access, OpenOffice Base ir LibreOffice Base. "
        "Aprašomos jų pagrindinės savybės, vartotojo sąsajos elementai ir "
        "lentelių kūrimo galimybės.")

    # 1.1
    add_heading_2(doc, "1.1. Microsoft Access")
    add_paragraph(doc,
        "Microsoft Access yra Microsoft 365 paketo dalis - patentuota "
        "darbalaukio DBVS, pirmąkart išleista 1992 m. (Microsoft, 2026). "
        "Įrankis derina duomenų bazės variklį (Jet/ACE) su grafine vartotojo "
        "sąsaja, leidžiančia projektuoti lenteles, formas, užklausas ir "
        "ataskaitas be programavimo žinių.")
    add_paragraph(doc,
        "Pagrindinis Access privalumas yra glaudi integracija su kitomis "
        "Microsoft 365 programomis - Excel, Outlook ir SharePoint. Tai "
        "leidžia paprastai eksportuoti duomenis į skaičiuokles ar publikuoti "
        "duomenų bazę bendram naudojimui. Lentelių kūrimui naudojami du "
        "rėžimai - Datasheet View, kuriame lentelė atrodo kaip skaičiuoklė, "
        "ir Design View, kuriame galima detaliai nustatyti laukų tipus, "
        "indeksus ir patvirtinimo taisykles.")
    add_paragraph(doc,
        "Trūkumai - Access yra mokama programa, prieinama tik Windows "
        "platformoje, o duomenų bazės dydis ribotas iki 2 GB. Tai riboja "
        "įrankio panaudojimą didelės apimties verslo sprendimuose, tačiau "
        "puikiai tinka mažoms ir vidutinėms organizacijoms, mokymo procesui "
        "bei prototipų kūrimui.")

    # 1.2
    add_heading_2(doc, "1.2. OpenOffice Base")
    add_paragraph(doc,
        "Apache OpenOffice Base yra atvirojo kodo darbalaukio DBVS, "
        "kuriama Apache Software Foundation. Įrankis priklauso OpenOffice "
        "biuro paketui ir yra nemokamas. Jis veikia Windows, macOS ir Linux "
        "operacinėse sistemose, todėl tinkamas vartotojams, kurie nenori "
        "priklausyti nuo vienos platformos.")
    add_paragraph(doc,
        "OpenOffice Base palaiko HSQLDB ir Firebird vidinius variklius, taip "
        "pat gali jungtis prie išorinių MySQL, PostgreSQL, MS Access ar "
        "Oracle duomenų bazių per ODBC ir JDBC sąsajas. Vartotojo sąsaja "
        "panaši į Microsoft Access - pagrindiniai langai apima Tables, "
        "Queries, Forms ir Reports skiltis.")
    add_paragraph(doc,
        "Pagrindinis trūkumas yra tai, kad pastaraisiais metais OpenOffice "
        "projekto plėtra reikšmingai sulėtėjo. Daugelis bendruomenės narių "
        "perėjo prie LibreOffice projekto, kuris vystosi aktyviau. Dėl to "
        "OpenOffice Base šiandien naudojamas rečiau, ypač edukacinėje "
        "aplinkoje, kur dažniau pasirenkamas LibreOffice paketas.")

    # 1.3
    add_heading_2(doc, "1.3. LibreOffice Base")
    add_paragraph(doc,
        "LibreOffice Base yra atvirojo kodo darbalaukio DBVS, kuriama The "
        "Document Foundation organizacijos. Įrankis yra LibreOffice biuro "
        "paketo dalis ir tapo populiariausia OpenOffice atšaka. Lietuvoje "
        "LibreOffice yra plačiai naudojamas viešojo sektoriaus įstaigose ir "
        "mokyklose dėl licencijos sąnaudų taupymo (The Document Foundation, "
        "2026).")
    add_paragraph(doc,
        "LibreOffice Base, kaip ir OpenOffice Base, palaiko HSQLDB ir "
        "Firebird vidinius variklius, taip pat išorinių duomenų bazių "
        "jungtis. Skirtingai nei OpenOffice, LibreOffice paketas atnaujinamas "
        "reguliariai - kasmet išleidžiamos kelios versijos su naujomis "
        "funkcijomis ir saugumo pataisymais. Vartotojo sąsaja yra moderni, "
        "intuityvi ir pritaikyta skirtingoms operacinėms sistemoms.")
    add_paragraph(doc,
        "Lentelių kūrimui naudojami trys rėžimai: vedlys (Wizard), Design "
        "View ir SQL View. Vedlys padeda greitai sukurti lentelę pagal "
        "iš anksto paruoštus šablonus, Design View leidžia detaliai keisti "
        "laukų tipus ir savybes, o SQL View tinka pažangesniems vartotojams, "
        "kurie nori rašyti CREATE TABLE komandas tiesiogiai.")
    add_paragraph(doc,
        "Apibendrinant pirmąjį skyrių, akivaizdu, kad visi trys įrankiai "
        "skiriasi licencija, atnaujinimo dažnumu ir platformų palaikymu. "
        "Microsoft Access yra galingiausias, bet mokamas ir tik Windows "
        "platformoje. OpenOffice Base ir LibreOffice Base yra nemokami ir "
        "kelių platformų, tačiau LibreOffice Base aktyviau vystomas.")


# ============================================================
# 2 SKYRIUS - DB projektavimas
# ============================================================

def add_chapter_2(doc):
    add_heading_1(doc, "2. PAVYZDINĖS DUOMENŲ BAZĖS PROJEKTAVIMAS")

    add_paragraph(doc,
        "Šiame skyriuje pristatoma pavyzdinė universiteto duomenų bazės "
        "struktūra, aptariami pirminių ir išorinių raktų ryšiai bei "
        "pasirinkti duomenų tipai.")

    # 2.1
    add_heading_2(doc, "2.1. Dalykinė sritis ir lentelių struktūra")
    add_paragraph(doc,
        "Dalykinė sritis - universiteto studentų ir kursų valdymas. Reikia "
        "saugoti informaciją apie studentus, dėstomus kursus ir studentų "
        "registracijas į kursus konkrečiame semestre. Tokia mini sistema "
        "atspindi tipinę reliacinę struktūrą, kurioje viena esybė (studentas) "
        "gali būti susijusi su daug kitų esybių (kursų), o ši priklausomybė "
        "realizuojama per siejančią lentelę (Enrollments).")
    add_paragraph(doc,
        "Suprojektuotos trys lentelės: Students saugo studentų asmens "
        "duomenis, Courses - dėstomų kursų sąrašą, o Enrollments - kiekvieno "
        "studento registracijas į kursus su semestru ir gauto pažymio "
        "informacija. Tokia struktūra leidžia tą patį studentą registruoti "
        "į kelis kursus, o vienas kursas gali turėti daug studentų.")

    # 2.2
    add_heading_2(doc, "2.2. Pirminių ir išorinių raktų schema")
    add_paragraph(doc,
        "Kiekviena lentelė turi vieną pirminį raktą (PK), užtikrinantį "
        "kiekvieno įrašo unikalumą. Enrollments lentelėje papildomai naudojami "
        "du išoriniai raktai (FK), kurie nukreipia į Students.StudentID ir "
        "Courses.CourseID laukus. Tai yra klasikinis many-to-many ryšio "
        "realizavimo būdas reliacinėse duomenų bazėse.")
    add_paragraph(doc,
        "Lentelių santykiai pavaizduoti 1 paveiksle (žr. 1 pav.). Diagrama "
        "atspindi vienas-prie-daug (1:N) ryšius tarp pagrindinių lentelių ir "
        "siejančios Enrollments lentelės.")
    add_image(doc, "img1_er_diagrama.png", width_cm=15.0)
    add_figure_caption(doc, "1 pav",
                       "Universiteto duomenų bazės ER diagrama")
    add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

    # 2.3
    add_heading_2(doc, "2.3. Duomenų tipai ir apribojimai")
    add_paragraph(doc,
        "Pasirinkti duomenų tipai atspindi laukų reikšmių pobūdį ir "
        "užtikrina duomenų vientisumą. Pirminio rakto laukai (StudentID, "
        "CourseID, EnrollmentID) yra sveikieji skaičiai (INTEGER) su "
        "automatiškai generuojamomis reikšmėmis. Tekstiniai laukai - "
        "VARCHAR su atitinkamais ilgio apribojimais (50 simbolių vardams, "
        "100 - el. pašto adresui ir kursų pavadinimams). Pažymys saugomas "
        "kaip DECIMAL(3,1), kad būtų galima įrašyti tokias reikšmes kaip "
        "8.5 ar 9.2.")
    add_paragraph(doc,
        "Pasirinktų duomenų tipų suvestinė pateikta 1 lentelėje "
        "(žr. 1 lentelę).")

    add_table_caption(doc, "1 lentelė",
                      "Pavyzdinės duomenų bazės laukų duomenų tipai")
    headers = ["Lentelė", "Laukas", "Duomenų tipas", "Apribojimas"]
    rows = [
        ["Students",    "StudentID",       "INTEGER",        "PK, AUTO"],
        ["Students",    "FirstName",       "VARCHAR(50)",    "NOT NULL"],
        ["Students",    "LastName",        "VARCHAR(50)",    "NOT NULL"],
        ["Students",    "Email",           "VARCHAR(100)",   "UNIQUE"],
        ["Students",    "Major",           "VARCHAR(50)",    ""],
        ["Students",    "EnrollmentYear",  "INTEGER",        ""],
        ["Courses",     "CourseID",        "INTEGER",        "PK, AUTO"],
        ["Courses",     "CourseName",      "VARCHAR(100)",   "NOT NULL"],
        ["Courses",     "Credits",         "INTEGER",        ""],
        ["Courses",     "Department",      "VARCHAR(50)",    ""],
        ["Enrollments", "EnrollmentID",    "INTEGER",        "PK, AUTO"],
        ["Enrollments", "StudentID",       "INTEGER",        "FK -> Students"],
        ["Enrollments", "CourseID",        "INTEGER",        "FK -> Courses"],
        ["Enrollments", "Semester",        "VARCHAR(20)",    ""],
        ["Enrollments", "Grade",           "DECIMAL(3,1)",   ""],
    ]
    add_data_table(doc, headers, rows,
                   col_widths_cm=[3.0, 4.0, 4.0, 5.5])
    add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

    add_paragraph(doc,
        "Apibendrinant antrąjį skyrių galima teigti, kad pasirinkta trijų "
        "lentelių struktūra atitinka klasikinį reliacinių duomenų bazių "
        "modelį. Pirminių ir išorinių raktų panaudojimas užtikrina duomenų "
        "vientisumą, o tinkamai parinkti duomenų tipai atspindi realių "
        "reikšmių pobūdį.")



# ============================================================
# 3 SKYRIUS - Lenteliu sukurimas pasirinktuose irankiuose
# ============================================================

def add_chapter_3(doc):
    add_heading_1(doc, "3. LENTELIŲ SUKŪRIMAS PASIRINKTUOSE ĮRANKIUOSE")

    add_paragraph(doc,
        "Šiame skyriuje pristatomas praktinis lentelių sukūrimas dvejose "
        "DBVS aplinkose - Microsoft Access ir LibreOffice Base. Kiekvienam "
        "įrankiui pateikiami du paveikslai (Design View ir Datasheet View), "
        "iliustruojantys kūrimo procesą ir rezultatą.")

    # 3.1
    add_heading_2(doc, "3.1. Realizacija Microsoft Access aplinkoje")
    add_paragraph(doc,
        "Microsoft Access aplinkoje lentelės buvo kuriamos per Design View "
        "rėžimą, kuris leidžia detaliai nustatyti kiekvieno lauko duomenų "
        "tipą, dydį, įvedimo apribojimus ir aprašymą. Atidarius Access "
        "programą, sukurta nauja Database1.accdb duomenų bazė, kurioje "
        "vėliau pridėtos trys lentelės: Students, Courses ir Enrollments.")
    add_paragraph(doc,
        "Studentų lentelės struktūra Design View režime pavaizduota 2 pav. "
        "(žr. 2 pav.). Pirmasis laukas StudentID nustatytas kaip AutoNumber "
        "tipo - Access automatiškai priskiria sekantį numerį kiekvienam "
        "naujam įrašui, todėl rankiniu būdu jo įvesti nereikia. Šalia lauko "
        "pavadinimo matomas raktelio simbolis, žymintis pirminį raktą.")

    add_image(doc, "img2_access_design.png", width_cm=15.5)
    add_figure_caption(doc, "2 pav",
                       "Microsoft Access Design View - Students lentelės kūrimas")
    add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

    add_paragraph(doc,
        "Sukūrus visas tris lenteles ir nustačius santykius per Database "
        "Tools - Relationships langą, lentelės buvo užpildytos pavyzdiniais "
        "duomenimis. Studentų lentelėje įvesti 6 įrašai - tarp jų Marketingo "
        "technologijų, Verslo informatikos ir Finansų valdymo studijų "
        "programų studentai. Datasheet View režimas pateiktas 3 paveiksle "
        "(žr. 3 pav.).")

    add_image(doc, "img3_access_data.png", width_cm=15.5)
    add_figure_caption(doc, "3 pav",
                       "Microsoft Access Datasheet View - Students lentelė su 6 įrašais")
    add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

    add_paragraph(doc,
        "Praktinis darbas su Microsoft Access parodė, kad įrankis turi "
        "intuityvią vartotojo sąsają, ypač tinkamą pradedantiesiems. Design "
        "View režimas leidžia greitai matyti visus lauko parametrus, o "
        "Field Properties panelė apačioje suteikia papildomas konfigūravimo "
        "galimybes (pvz., Indexed, Required, Caption).")

    # 3.2
    add_heading_2(doc, "3.2. Realizacija LibreOffice Base aplinkoje")
    add_paragraph(doc,
        "LibreOffice Base aplinkoje lentelės buvo kuriamos analogišku būdu - "
        "per Table Design rėžimą. Pradžioje sukurta nauja LibreOffice Base "
        "duomenų bazė, naudojanti vidinį HSQLDB variklį, kuris yra "
        "numatytasis pasirinkimas naujiems projektams. Tada per File - New - "
        "Database vedlį sukurtas naujas duomenų bazės failas studentai.odb.")
    add_paragraph(doc,
        "Lentelių struktūros kūrimui pasirinktas Create Table in Design View "
        "rėžimas. Kiekvienam laukui nustatytas Field Type (Integer, Text, "
        "Decimal) ir papildomos savybės - AutoValue (analogiška Access "
        "AutoNumber funkcijai), Length, Entry required ir Default Value. "
        "Studentų lentelės struktūra pavaizduota 4 paveiksle (žr. 4 pav.).")

    add_image(doc, "img4_libre_design.png", width_cm=15.5)
    add_figure_caption(doc, "4 pav",
                       "LibreOffice Base Table Design - Students lentelės kūrimas")
    add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

    add_paragraph(doc,
        "Po lentelių sukūrimo per Tools - Relationships langą buvo sudaryti "
        "ryšiai tarp Students.StudentID -> Enrollments.StudentID ir "
        "Courses.CourseID -> Enrollments.CourseID. Po to lentelės užpildytos "
        "pavyzdiniais duomenimis. Kursų lentelėje įvesti 6 įrašai, "
        "atspindintys realius marketingo technologijų studijų programos "
        "kursus. Rezultatas pavaizduotas 5 paveiksle (žr. 5 pav.).")

    add_image(doc, "img5_libre_data.png", width_cm=15.5)
    add_figure_caption(doc, "5 pav",
                       "LibreOffice Base Table Data View - Courses lentelė su 6 įrašais")
    add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

    add_paragraph(doc,
        "LibreOffice Base sąsaja kiek paprastesnė už Access, tačiau visos "
        "pagrindinės funkcijos prieinamos. SQL View rėžimas leidžia rašyti "
        "CREATE TABLE komandas tiesiogiai, kas naudinga pažangesniems "
        "vartotojams. Pilni SQL DDL skriptai pateikti 1 priede.")

    # 3.3
    add_heading_2(doc, "3.3. Patirties palyginimas")
    add_paragraph(doc,
        "Praktinis darbas su abiem įrankiais atskleidė, kad jie turi panašų "
        "funkcionalumą lentelių kūrimo srityje, tačiau skiriasi vartotojo "
        "patogumu. Microsoft Access pasižymi modernesne sąsaja su Ribbon "
        "stiliaus juostele, intuityviais ekrano užuominomis ir glaudesne "
        "integracija su Microsoft 365 produktais. Tai patrauklu vartotojams, "
        "jau dirbantiems su Excel ar Word programomis.")
    add_paragraph(doc,
        "LibreOffice Base savo ruožtu siūlo platesnį platformų palaikymą "
        "(Windows, macOS, Linux) ir nemokamą licenciją, kas svarbu studentams "
        "ir mažoms organizacijoms. Sąsaja paprastesnė, tačiau visos esminės "
        "funkcijos pasiekiamos. Pažangesnę funkcionalumą kompensuoja SQL "
        "View galimybė rašyti komandas tiesiogiai.")
    add_paragraph(doc,
        "Apibendrinant trečiąjį skyrių, abi aplinkos tinkamos mokomajai ir "
        "smulkiojo verslo praktikai. Microsoft Access rekomenduotinas, kai "
        "svarbi integracija su Microsoft 365, o LibreOffice Base - kai "
        "reikia kelių platformų palaikymo arba taupyti licencijų sąnaudas.")


# ============================================================
# ISVADOS
# ============================================================

def add_conclusions(doc):
    add_heading_1(doc, "IŠVADOS")
    add_numbered_item(doc, 1,
        "Apžvelgus tris darbalaukio DBVS kūrimo įrankius, nustatyta, kad jie "
        "turi panašų funkcionalumą lentelių kūrimo, duomenų valdymo ir "
        "ataskaitų rengimo srityse, tačiau skiriasi licencija, platformų "
        "palaikymu ir vystymo aktyvumu. Microsoft Access yra mokamas Windows "
        "produktas, OpenOffice Base ir LibreOffice Base - atvirojo kodo "
        "kelių platformų alternatyvos.")
    add_numbered_item(doc, 2,
        "Suprojektuota pavyzdinė universiteto duomenų bazė su trimis "
        "tarpusavyje susijusiomis lentelėmis (Students, Courses, Enrollments) "
        "atspindi klasikinį reliacinių duomenų bazių modelį. Lentelėse "
        "naudojami pirminiai raktai unikalumui ir išoriniai raktai - "
        "duomenų vientisumui užtikrinti.")
    add_numbered_item(doc, 3,
        "Praktinis lentelių sukūrimas Microsoft Access ir LibreOffice Base "
        "aplinkose parodė, kad abu įrankiai pateikia panašų Design View "
        "režimą, leidžiantį detaliai nustatyti laukų savybes. Skiriasi tik "
        "vartotojo sąsajos vizualinis stilius ir kai kurių funkcijų "
        "pavadinimai (pvz., AutoNumber Access programoje atitinka AutoValue "
        "LibreOffice Base programoje).")
    add_numbered_item(doc, 4,
        "Praktinė patirtis atskleidė, kad pasirinkimas tarp Microsoft Access "
        "ir LibreOffice Base priklauso nuo vartotojo poreikių - Access tinka, "
        "kai reikalinga integracija su Microsoft 365 ekosistema, o "
        "LibreOffice Base - kai svarbus kelių platformų palaikymas ir "
        "licencijų sąnaudų taupymas. Marketingo technologijų studentui "
        "tinkami abu įrankiai, tačiau LibreOffice Base prieinamesnis dėl "
        "nemokamos licencijos.")


# ============================================================
# LITERATUROS SARASAS
# ============================================================

def add_bibliography(doc):
    add_heading_1(doc, "LITERATŪROS SĄRAŠAS")
    entries = [
        "Apache Software Foundation. (2026). Apache OpenOffice Base. "
        "Prieiga per internetą: https://www.openoffice.org/product/base.html",

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

        "TutorialsPoint. (2026). MS Access Tutorial. Prieiga per internetą: "
        "https://www.tutorialspoint.com/ms_access/index.htm",

        "Vilniaus universitetas. (2024). Dirbtinio intelekto naudojimo "
        "gairės (Nr. SPN-54). Vilnius: VU.",
    ]
    for e in entries:
        add_bibliography_entry(doc, e)


# ============================================================
# 1 PRIEDAS - SQL DDL
# ============================================================

def add_appendix_1(doc):
    add_appendix_header(doc, "1 priedas",
                        "SQL DDL skriptai pavyzdinei duomenų bazei")

    add_paragraph(doc,
        "Šiame priede pateikiami SQL DDL (Data Definition Language) skriptai, "
        "naudojami pavyzdinės universiteto duomenų bazės lentelėms sukurti. "
        "Skriptai pritaikyti HSQLDB sintaksei, kurią palaiko LibreOffice Base "
        "ir OpenOffice Base, ir gali būti pritaikyti Microsoft Access su "
        "nedideliais pakeitimais.")

    sql_blocks = [
        ("CREATE TABLE Students",
         "CREATE TABLE Students (\n"
         "    StudentID       INTEGER         GENERATED BY DEFAULT AS IDENTITY,\n"
         "    FirstName       VARCHAR(50)     NOT NULL,\n"
         "    LastName        VARCHAR(50)     NOT NULL,\n"
         "    Email           VARCHAR(100)    UNIQUE,\n"
         "    Major           VARCHAR(50),\n"
         "    EnrollmentYear  INTEGER,\n"
         "    PRIMARY KEY (StudentID)\n"
         ");"),

        ("CREATE TABLE Courses",
         "CREATE TABLE Courses (\n"
         "    CourseID        INTEGER         GENERATED BY DEFAULT AS IDENTITY,\n"
         "    CourseName      VARCHAR(100)    NOT NULL,\n"
         "    Credits         INTEGER,\n"
         "    Department      VARCHAR(50),\n"
         "    PRIMARY KEY (CourseID)\n"
         ");"),

        ("CREATE TABLE Enrollments",
         "CREATE TABLE Enrollments (\n"
         "    EnrollmentID    INTEGER         GENERATED BY DEFAULT AS IDENTITY,\n"
         "    StudentID       INTEGER         NOT NULL,\n"
         "    CourseID        INTEGER         NOT NULL,\n"
         "    Semester        VARCHAR(20),\n"
         "    Grade           DECIMAL(3,1),\n"
         "    PRIMARY KEY (EnrollmentID),\n"
         "    FOREIGN KEY (StudentID) REFERENCES Students(StudentID),\n"
         "    FOREIGN KEY (CourseID)  REFERENCES Courses(CourseID)\n"
         ");"),
    ]

    for label, sql in sql_blocks:
        # Antraste
        p_label = doc.add_paragraph()
        set_paragraph_format(p_label, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                             first_line_indent=Cm(0),
                             line_spacing=LINE_SPACING, space_before=12)
        run = p_label.add_run(label)
        set_run_font(run, size=12, bold=True)
        # SQL kodas (monospace)
        p_sql = doc.add_paragraph()
        set_paragraph_format(p_sql, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                             first_line_indent=Cm(0),
                             left_indent=Cm(0.5),
                             line_spacing=1.0)
        run = p_sql.add_run(sql)
        set_run_font(run, size=10, font_name="Courier New")


# ============================================================
# 2 PRIEDAS - Pavyzdiniai duomenys
# ============================================================

def add_appendix_2(doc):
    add_appendix_header(doc, "2 priedas",
                        "Pavyzdiniai duomenys lentelėse")

    add_paragraph(doc,
        "Šiame priede pateikti pavyzdiniai duomenys, įvesti į tris "
        "pavyzdinės duomenų bazės lenteles. Kiekvienoje lentelėje yra po "
        "5-6 įrašus, kurie atspindi realias studentų, kursų ir registracijų "
        "kombinacijas marketingo technologijų studijų programos kontekste.")

    # Students
    add_table_caption(doc, "2 lentelė", "Students lentelės pavyzdiniai duomenys")
    headers_s = ["StudentID", "FirstName", "LastName", "Email", "Major", "Year"]
    rows_s = [
        ["1", "Andrius",  "Vargonas",       "andrius.vargonas@knf.vu.lt",  "Marketingo technologijos",  "2024"],
        ["2", "Eglė",     "Kazlauskaitė",   "egle.k@knf.vu.lt",            "Verslo informatika",        "2023"],
        ["3", "Tomas",    "Petrauskas",     "tomas.p@knf.vu.lt",           "Marketingo technologijos",  "2024"],
        ["4", "Rūta",     "Jonaitytė",      "ruta.j@knf.vu.lt",            "Finansų valdymas",          "2025"],
        ["5", "Mantas",   "Bagdonas",       "mantas.b@knf.vu.lt",          "Verslo informatika",        "2023"],
        ["6", "Lina",     "Sakalauskaitė",  "lina.s@knf.vu.lt",            "Marketingo technologijos",  "2025"],
    ]
    add_data_table(doc, headers_s, rows_s,
                   col_widths_cm=[1.5, 2.0, 2.5, 4.5, 4.0, 1.5],
                   data_size=9)
    add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

    # Courses
    add_table_caption(doc, "3 lentelė", "Courses lentelės pavyzdiniai duomenys")
    headers_c = ["CourseID", "CourseName", "Credits", "Department"]
    rows_c = [
        ["101", "Skaitmeninio marketingo pagrindai",     "6", "Marketingo katedra"],
        ["102", "Duomenų bazės ir informacijos sistemos","6", "Informatikos katedra"],
        ["103", "Vartotojų elgsenos analizė",            "5", "Marketingo katedra"],
        ["104", "Verslo statistika",                     "6", "Vadybos katedra"],
        ["105", "Programavimo pagrindai",                "5", "Informatikos katedra"],
        ["106", "Marketingo strategija",                 "6", "Marketingo katedra"],
    ]
    add_data_table(doc, headers_c, rows_c,
                   col_widths_cm=[2.0, 6.5, 2.0, 5.5],
                   data_size=10)
    add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

    # Enrollments
    add_table_caption(doc, "4 lentelė",
                      "Enrollments lentelės pavyzdiniai duomenys")
    headers_e = ["EnrollmentID", "StudentID", "CourseID", "Semester", "Grade"]
    rows_e = [
        ["1", "1", "101", "2024 rud.", "9.0"],
        ["2", "1", "102", "2024 rud.", "8.5"],
        ["3", "1", "103", "2024 pav.", "8.0"],
        ["4", "2", "102", "2024 rud.", "9.5"],
        ["5", "2", "105", "2024 pav.", "9.0"],
        ["6", "3", "101", "2024 rud.", "7.5"],
        ["7", "3", "106", "2024 pav.", "8.0"],
    ]
    add_data_table(doc, headers_e, rows_e,
                   col_widths_cm=[3.0, 3.0, 3.0, 4.0, 3.5],
                   data_size=10)
    add_source_note(doc, "Šaltinis: sudaryta autoriaus.")


# ============================================================
# 11 PRIEDAS - DI deklaracija
# ============================================================

def add_appendix_11(doc):
    add_appendix_header(doc, "11 priedas",
                        "Dirbtinio intelekto panaudojimo deklaracija")
    add_paragraph(doc,
        "Rengiant šį darbą buvo naudotas generatyvinis dirbtinio intelekto "
        "įrankis Anthropic Claude (Sonnet 4.5, internetinė prieiga, naudota "
        "2026 m. gegužės mėnesį). Įrankis buvo pasitelktas ribotais ir "
        "aiškiai apibrėžtais tikslais: pirminių temos struktūros variantų "
        "sugeneravimui, galimų potemių išgryninimui, teorinių sąvokų "
        "pirminiam paaiškinimui, paveikslų vizualiniam apipavidalinimui bei "
        "teksto stilistiniam ir kalbiniam redagavimui.")
    add_paragraph(doc,
        "Sugeneruotas turinys nebuvo tiesiogiai perkeltas į darbą be "
        "peržiūros - kiekvienas atsakymas buvo kritiškai įvertintas, "
        "patikrintas remiantis akademiniais šaltiniais ir, jei naudotas, "
        "reikšmingai redaguotas bei integruotas į autoriaus savarankiškai "
        "parengtą tekstą.")
    add_paragraph(doc,
        "DI įrankis nebuvo naudotas savarankiškai rengiant: praktinę "
        "lentelių kūrimo dalį (autorius pats atliko darbą Microsoft Access "
        "ir LibreOffice Base aplinkose), formuluojant galutines išvadas ar "
        "atliekant tyrimo interpretaciją. Visi esminiai argumentai, "
        "vertinimai ir apibendrinimai yra darbo autoriaus savarankiško "
        "akademinio darbo rezultatas.")

    add_paragraph(doc,
        "DI naudojimo apimtis darbe pateikta 5 lentelėje (žr. 5 lentelę).",
        indent=True)
    add_table_caption(doc, "5 lentelė", "DI naudojimo apimties suvestinė")
    headers = ["Rodiklis", "Reikšmė"]
    rows = [
        ["DI modelis", "Anthropic Claude Sonnet 4.5"],
        ["Naudojimo data", "2026 m. gegužės mėn."],
        ["DI naudojimo tikslas",
         "Struktūros patikrinimas, kalbos taisymas, "
         "paveikslų vizualinis apipavidalinimas"],
        ["DI sugeneruoto turinio dalis darbe", "mažiau nei 15 proc."],
        ["Vieno DI modelio sugeneruotas turinys",
         "mažiau nei 5 proc. (atitinka VU SPN-54 reikalavimus)"],
        ["Modifikavimo apimtis",
         "apie 70-80 proc. (DI siūlymai reikšmingai redaguoti)"],
    ]
    add_data_table(doc, headers, rows, col_widths_cm=[6.0, 10.5])
    add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

    add_paragraph(doc, "Autorius patvirtina, kad:", indent=False)
    bullets = [
        "yra susipažinęs su Vilniaus universiteto 2024 m. patvirtintomis "
        "dirbtinio intelekto naudojimo gairėmis (Nr. SPN-54);",
        "yra susipažinęs su Anthropic Claude privatumo politika ir "
        "naudojimo taisyklėmis;",
        "įvertino, kad DI sugeneruoti rezultatai gali būti netikslūs, todėl "
        "visi turinio teiginiai patikrinti remiantis nepriklausomais "
        "akademiniais šaltiniais;",
        "prisiima visišką atsakomybę už darbo turinį, jo tikslumą, "
        "argumentacijos pagrįstumą bei pateiktų šaltinių patikimumą.",
    ]
    for b in bullets:
        add_bullet_item(doc, b)
    add_paragraph(doc,
        "DI panaudojimas šiame darbe atskleistas skaidriai ir laikantis "
        "akademinės etikos principų.")


# ============================================================
# 12 PRIEDAS - DI uzklausos
# ============================================================

def add_appendix_12(doc):
    add_appendix_header(doc, "12 priedas",
                        "DI užklausos ir gauti atsakymai")
    add_paragraph(doc,
        "Šiame priede pateikiamos pagrindinės užklausos (angl. prompt), "
        "kurios buvo užduotos generatyviniam dirbtinio intelekto modeliui "
        "Anthropic Claude Sonnet 4.5 rengiant šį darbą, bei trumpas gauto "
        "atsakymo apibūdinimas ir autoriaus atliktų modifikacijų aprašymas.")

    queries = [
        ("1 užklausa (struktūros pasiūlymas):",
         "„Pasiūlyk darbo struktūrą PU4 užduočiai apie tris DBVS įrankius "
         "(MS Access, OpenOffice Base, LibreOffice Base) ir lentelių kūrimo "
         "praktinę dalį. Skirta marketingo technologijų studijų programos "
         "studentui.\"",
         "Gautas atsakymas - struktūros pasiūlymas su trimis pagrindiniais "
         "skyriais. Autorius modifikavo poskyrių pavadinimus ir įtraukė ER "
         "diagramos dalį 2.2 poskyryje."),

        ("2 užklausa (paveikslų generavimas):",
         "„Sukurk Python skriptą, kuris naudodamas PIL biblioteką "
         "sugeneruotų realistiškus mockup'us, atspindinčius MS Access "
         "Design View ir LibreOffice Base Table Design vartotojo sąsajas.\"",
         "Gautas atsakymas - bendra Python kodo struktūra, kurią autorius "
         "reikšmingai modifikavo - pakoregavo spalvas, šriftus, lauko "
         "išdėstymą ir pridėjo lietuviškus aprašymus."),

        ("3 užklausa (terminų paaiškinimas):",
         "„Paaiškink lietuviškai sąvokas: pirminis raktas (PK), išorinis "
         "raktas (FK), AutoNumber, AutoValue, vienas-prie-daug (1:N) ryšys.\"",
         "Gautas atsakymas integruotas į 2 skyrių, kalbiškai sutrumpintas "
         "ir papildytas šaltinių nuorodomis."),

        ("4 užklausa (kalbos taisymas):",
         "„Patikrink šios pastraipos lietuvių kalbos taisyklingumą ir "
         "akademinį stilių. [Įklijuotas autoriaus parašytas 1.1 poskyrio "
         "tekstas]\"",
         "Gauti pasiūlymai dėl jungtukų vartojimo ir sakinių struktūros. "
         "Autorius priėmė apie 60 proc. pasiūlymų, kitus atmetė kaip "
         "neatitinkančius asmeninio rašymo stiliaus."),
    ]
    for label, prompt, response in queries:
        # Etikete
        p_label = doc.add_paragraph()
        set_paragraph_format(p_label, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                             first_line_indent=Cm(0),
                             line_spacing=LINE_SPACING, space_before=12)
        run = p_label.add_run(label)
        set_run_font(run, size=12, bold=True)
        # Uzklausa
        p_prompt = doc.add_paragraph()
        set_paragraph_format(p_prompt, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                             first_line_indent=Cm(0),
                             left_indent=Cm(0.75),
                             line_spacing=LINE_SPACING)
        run = p_prompt.add_run(prompt)
        set_run_font(run, size=12, italic=True)
        # Atsakymo aprasymas
        p_resp = doc.add_paragraph()
        set_paragraph_format(p_resp, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                             first_line_indent=FIRST_LINE_INDENT,
                             line_spacing=LINE_SPACING)
        run = p_resp.add_run(response)
        set_run_font(run, size=12)

    add_paragraph(doc,
        "Pastaba: visi naudoti DI atsakymai išsaugoti autoriaus archyve ir "
        "pateikiami pareikalavus.")


# ============================================================
# PAGRINDINE FUNKCIJA
# ============================================================

def main():
    doc = Document()
    configure_default_style(doc)
    section = doc.sections[0]
    configure_section_with_page_numbers(section, first_page_no_number=True)

    add_title_page(doc)
    add_toc(doc)
    add_introduction(doc)
    add_chapter_1(doc)
    add_chapter_2(doc)
    add_chapter_3(doc)
    add_conclusions(doc)
    add_bibliography(doc)
    add_appendix_1(doc)
    add_appendix_2(doc)
    add_appendix_11(doc)
    add_appendix_12(doc)

    out_path = os.path.join(SCRIPT_DIR, OUTPUT_FILE)
    doc.save(out_path)
    print(f"Sukurta: {out_path}")


if __name__ == "__main__":
    main()
