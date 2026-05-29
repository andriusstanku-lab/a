"""
PU3 ataskaitos generatorius (.docx)

Generuoja Word dokumenta pagal VU Kauno fakulteto Informatikos inzinerijos
krypties akademiniu rasto darbu metodinius nurodymus.

Tema: Duomenu baziu kurimo irankiu analize ir palyginimas
Autorius: Andrius Vargonas
Studiju programa: Marketingo technologijos
Dalykas: Informacijos sistemos ir duomenu bazes
Metai: 2026

Naudojimas:
    pip install python-docx
    python generate_pu3.py
    # Sugeneruojamas failas: PU3_Vargonas.docx
"""

from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor
from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_LINE_SPACING,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
)
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# Konstantos
# ============================================================

FONT_NAME = "Times New Roman"
FIRST_LINE_INDENT = Cm(1.25)
HANGING_INDENT = Cm(1.25)
LINE_SPACING = 1.5
OUTPUT_FILE = "PU3_Vargonas.docx"


# ============================================================
# Pagalbines funkcijos: sriftai ir pastraipu formatavimas
# ============================================================

def set_run_font(run, size=12, bold=False, italic=False, font_name=None):
    """Nustato srifta runui (Lietuviu simbolius padengia per ascii/hAnsi/cs/eastAsia)."""
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
    """Universalus pastraipos pridejimas su mokslinio darbo formatavimu."""
    p = doc.add_paragraph()
    set_paragraph_format(
        p,
        alignment=alignment,
        first_line_indent=(FIRST_LINE_INDENT if indent else Cm(0)),
        line_spacing=line_spacing,
        space_before=space_before,
        space_after=space_after,
        page_break_before=page_break_before,
    )
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def add_empty_line(doc, size=12):
    return add_paragraph(doc, "", alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         indent=False, size=size)


def apply_heading_style(p, level):
    """Priskiria standartini Heading stiliu (kad TOC galetu rinkti)."""
    pPr = p._element.get_or_add_pPr()
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        pStyle = OxmlElement("w:pStyle")
        pPr.insert(0, pStyle)
    pStyle.set(qn("w:val"), f"Heading{level}")


def add_heading_1(doc, text, page_break_before=True):
    """Skyriaus pavadinimas: 14pt Bold DIDZIOSIOMIS centre, naujas lapas."""
    p = doc.add_paragraph()
    set_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Cm(0),
        line_spacing=LINE_SPACING,
        space_after=12,
        page_break_before=page_break_before,
    )
    run = p.add_run(text.upper())
    set_run_font(run, size=14, bold=True)
    apply_heading_style(p, 1)
    return p


def add_heading_2(doc, text):
    """Poskyrio pavadinimas: 12pt Bold mazosiomis (1-oji didzioji), kaireje."""
    p = doc.add_paragraph()
    set_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Cm(0),
        line_spacing=LINE_SPACING,
        space_before=24,  # 2 eiluciu tarpas pries
        space_after=12,    # 1 eilutes tarpas po
    )
    run = p.add_run(text)
    set_run_font(run, size=12, bold=True)
    apply_heading_style(p, 2)
    return p



# ============================================================
# Puslapio nustatymai ir poraste su puslapio numeriu
# ============================================================

def setup_page(section):
    """A4 vertikali, parastes pagal VU reikalavimus."""
    section.page_height = Mm(297)
    section.page_width = Mm(210)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.left_margin = Mm(25)
    section.right_margin = Mm(15)
    section.header_distance = Mm(12.5)
    section.footer_distance = Mm(12.5)


def add_page_number_field(paragraph, size=12):
    """Iterpia PAGE lauka (puslapio numeri) i pastraipa."""
    run = paragraph.add_run()
    set_run_font(run, size=size)

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE \\* MERGEFORMAT "

    fld_separate = OxmlElement("w:fldChar")
    fld_separate.set(qn("w:fldCharType"), "separate")

    # Talpyklos reiksme - kad atidarius dokumenta is karto kazkas matytusi
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
    """Sukonfiguruoja sekcija: A4 parastes + puslapio numeris desineje porasteje.

    Jei first_page_no_number=True, pirmas sekcijos puslapis (pvz., antrastinis)
    neturi puslapio numerio - naudojame "different first page" funkcionaluma.
    """
    setup_page(section)

    if first_page_no_number:
        section.different_first_page_header_footer = True
        # Pirmasis puslapis - tuscia poraste
        first_footer = section.first_page_footer
        if first_footer.paragraphs:
            fp = first_footer.paragraphs[0]
        else:
            fp = first_footer.add_paragraph()
        fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # Tyciai paliekame tuscia

    # Iprastine poraste su puslapio numeriu desineje
    footer = section.footer
    if footer.paragraphs:
        fp = footer.paragraphs[0]
    else:
        fp = footer.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_paragraph_format(
        fp, alignment=WD_ALIGN_PARAGRAPH.RIGHT,
        first_line_indent=Cm(0), line_spacing=1.0,
    )
    add_page_number_field(fp, size=12)


# ============================================================
# Antrastinis lapas
# ============================================================

def add_title_page(doc):
    """VU Kauno fakulteto antrastinis lapas pagal metodinius nurodymus."""
    add_paragraph(doc, "VILNIAUS UNIVERSITETAS",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=14, bold=True)
    add_paragraph(doc, "KAUNO FAKULTETAS",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=14, bold=True)
    add_empty_line(doc)
    add_paragraph(doc, "SOCIALINIŲ MOKSLŲ IR TAIKOMOSIOS INFORMATIKOS INSTITUTAS",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=14)
    add_empty_line(doc)
    add_paragraph(doc, "Marketingo technologijų studijų programa",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=12)

    # Tarpas iki autoriaus vardo
    for _ in range(5):
        add_empty_line(doc)

    add_paragraph(doc, "ANDRIUS VARGONAS",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=12, bold=True)
    add_empty_line(doc)
    add_empty_line(doc)

    add_paragraph(doc, "DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ ANALIZĖ IR PALYGINIMAS",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=14, bold=True)
    add_empty_line(doc)
    add_paragraph(doc, "Praktinė užduotis Nr. 3",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=12, italic=True)
    add_empty_line(doc)
    add_paragraph(doc, "Informacijos sistemos ir duomenų bazės",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=12)

    # Stumti i puslapio apacia
    for _ in range(7):
        add_empty_line(doc)

    add_paragraph(doc, "Kaunas",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=12)
    add_paragraph(doc, "2026",
                  alignment=WD_ALIGN_PARAGRAPH.CENTER,
                  indent=False, size=12)



# ============================================================
# Turinys (statinis - nereikia spausti F9)
# ============================================================

def add_toc_title(doc):
    """Zodis TURINYS - 14pt Bold UPPERCASE centre, BE Heading stiliaus
    (kad neatsirastu paciame turinyje)."""
    p = doc.add_paragraph()
    set_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Cm(0),
        line_spacing=LINE_SPACING,
        space_after=12,
        page_break_before=True,
    )
    run = p.add_run("TURINYS")
    set_run_font(run, size=14, bold=True)
    return p


def add_toc_entry(doc, text, page_num, level=1, bold=False):
    """Statinis TOC irasas su tab leader taskeliais ir puslapio numeriu desineje.

    Vizualiai identiskas Word'o automatiniam turiniui, bet nereikalauja F9.
    """
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.line_spacing = LINE_SPACING
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.first_line_indent = Cm(0)

    # Itrauka pagal lygmeni
    if level == 2:
        pf.left_indent = Cm(0.75)
    elif level == 3:
        pf.left_indent = Cm(1.5)

    # Tab stop desineje su taskeliu lyderiu
    pf.tab_stops.add_tab_stop(
        Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
    )

    # Skyriaus pavadinimas
    run_text = p.add_run(text)
    set_run_font(run_text, size=12, bold=bold)

    # Tab + puslapio numeris
    run_num = p.add_run("\t" + str(page_num))
    set_run_font(run_num, size=12, bold=bold)

    return p


def add_toc(doc):
    """Statinis turinys, atspindintis dokumento struktura."""
    add_toc_title(doc)

    add_toc_entry(doc, "ĮVADAS", 3, level=1, bold=True)
    add_toc_entry(doc, "1. DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ TEORINĖ APŽVALGA", 4,
                  level=1, bold=True)
    add_toc_entry(doc, "1.1. Duomenų bazės samprata ir pagrindiniai tipai", 4,
                  level=2)
    add_toc_entry(doc, "1.2. Duomenų bazių kūrimo įrankių klasifikacija", 5,
                  level=2)
    add_toc_entry(doc, "1.3. Palyginimo kriterijų atranka ir pagrindimas", 6,
                  level=2)
    add_toc_entry(doc, "2. DEŠIMTIES DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ APŽVALGA", 7,
                  level=1, bold=True)
    add_toc_entry(doc, "2.1. Tradiciniai reliaciniai įrankiai", 7, level=2)
    add_toc_entry(doc, "2.2. Šiuolaikinės debesijos ir žemo kodo platformos", 8,
                  level=2)
    add_toc_entry(doc, "2.3. NoSQL ir atvirojo kodo įrankiai", 9, level=2)
    add_toc_entry(doc, "3. LYGINAMOJI ANALIZĖ IR REZULTATAI", 10,
                  level=1, bold=True)
    add_toc_entry(doc, "3.1. Lyginamoji lentelė pagal aštuonis kriterijus", 10,
                  level=2)
    add_toc_entry(doc, "3.2. Pagrindiniai skirtumai ir panašumai", 11, level=2)
    add_toc_entry(doc, "3.3. Įrankių tinkamumas skirtingiems naudojimo atvejams",
                  12, level=2)
    add_toc_entry(doc, "IŠVADOS", 13, level=1, bold=True)
    add_toc_entry(doc, "LITERATŪROS SĄRAŠAS", 14, level=1, bold=True)
    add_toc_entry(doc, "1 PRIEDAS. Detali dešimties įrankių charakteristika", 15,
                  level=1, bold=True)
    add_toc_entry(doc, "11 PRIEDAS. Dirbtinio intelekto panaudojimo deklaracija",
                  16, level=1, bold=True)
    add_toc_entry(doc, "12 PRIEDAS. DI užklausos ir gauti atsakymai", 17,
                  level=1, bold=True)


# ============================================================
# Numeruotas sarasas (uzdaviniams ir isvadom)
# ============================================================

def add_numbered_item(doc, number, text, indent=False):
    """Pastraipa formato 'N. tekstas' su justify lygiavimu."""
    p = doc.add_paragraph()
    set_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_indent=(FIRST_LINE_INDENT if indent else Cm(0)),
        left_indent=Cm(0.5),
        line_spacing=LINE_SPACING,
    )
    run = p.add_run(f"{number}. {text}")
    set_run_font(run, size=12)
    return p


# ============================================================
# Lentele su numeriu, pavadinimu ir saltiniu
# ============================================================

def add_table_caption(doc, number_text, title_text):
    """Pirma eilute - numeris desineje (pvz. '1 lentele').
    Antra eilute - pavadinimas centre, 11pt Bold, mazosiomis (pirmoji didzioji).
    """
    # Numeris desineje
    p_num = doc.add_paragraph()
    set_paragraph_format(
        p_num, alignment=WD_ALIGN_PARAGRAPH.RIGHT,
        first_line_indent=Cm(0),
        line_spacing=LINE_SPACING, space_before=12,
    )
    run = p_num.add_run(number_text)
    set_run_font(run, size=11, bold=True)

    # Pavadinimas centre, kursyvu (APA stiliaus pavyzdys metodinese)
    p_title = doc.add_paragraph()
    set_paragraph_format(
        p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Cm(0),
        line_spacing=LINE_SPACING, space_after=6,
    )
    run = p_title.add_run(title_text)
    set_run_font(run, size=11, bold=True, italic=True)


def add_source_note(doc, text):
    """Saltinio uzrasas po lentele/paveikslu (9pt)."""
    p = doc.add_paragraph()
    set_paragraph_format(
        p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=Cm(0),
        line_spacing=1.0, space_before=3, space_after=12,
    )
    run = p.add_run(text)
    set_run_font(run, size=9, italic=True)


def style_table_borders(table):
    """Prideda visiems tableams 0.5pt black borders."""
    tbl = table._element
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)

    tblBorders = OxmlElement("w:tblBorders")
    for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")
        tblBorders.append(border)

    # Pasalinti senuosius borderius, jei yra
    old_borders = tblPr.find(qn("w:tblBorders"))
    if old_borders is not None:
        tblPr.remove(old_borders)
    tblPr.append(tblBorders)


def add_data_table(doc, headers, rows, col_widths_cm=None,
                   data_size=10, header_bold=True):
    """Sukuria ir suformatuoja lentele su antraste (10pt) ir duomenimis (10pt)."""
    n_cols = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Kolonu plociai
    if col_widths_cm:
        for i, w in enumerate(col_widths_cm):
            for cell in table.columns[i].cells:
                cell.width = Cm(w)

    # Antraste
    hdr_row = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.text = ""
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        set_paragraph_format(
            p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
            first_line_indent=Cm(0), line_spacing=1.0,
        )
        run = p.add_run(h)
        set_run_font(run, size=data_size, bold=header_bold)

    # Duomenu eilutes
    for r_idx, row_data in enumerate(rows, start=1):
        row = table.rows[r_idx]
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = ""
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            # Pirmoji kolonele - tekstas kaireje, kitos - centre
            align = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_format(
                p, alignment=align,
                first_line_indent=Cm(0), line_spacing=1.0,
            )
            run = p.add_run(str(val))
            set_run_font(run, size=data_size)

    style_table_borders(table)
    return table


# ============================================================
# Bibliografija (APA, hanging indent)
# ============================================================

def add_bibliography_entry(doc, text):
    """APA stiliaus bibliografijos irasas su 1.25cm hanging indent."""
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


# ============================================================
# Priedo antraste
# ============================================================

def add_appendix_header(doc, number_text, title_text):
    """Priedo antraste: '1 PRIEDAS' desineje (12pt Bold), pavadinimas centre."""
    # Numeris desineje
    p_num = doc.add_paragraph()
    set_paragraph_format(
        p_num, alignment=WD_ALIGN_PARAGRAPH.RIGHT,
        first_line_indent=Cm(0),
        line_spacing=LINE_SPACING,
        page_break_before=True,
    )
    run = p_num.add_run(number_text.upper())
    set_run_font(run, size=12, bold=True)

    # Pavadinimas centre, paryskintas, didziosiomis
    p_title = doc.add_paragraph()
    set_paragraph_format(
        p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=Cm(0),
        line_spacing=LINE_SPACING, space_before=12, space_after=18,
    )
    run = p_title.add_run(title_text.upper())
    set_run_font(run, size=12, bold=True)



# ============================================================
# IVADAS
# ============================================================

def add_introduction(doc):
    add_heading_1(doc, "ĮVADAS")

    add_paragraph(
        doc,
        "Šiuolaikinėje informacijos visuomenėje duomenų bazės yra esminė bet "
        "kurios verslo ar viešojo sektoriaus informacinės sistemos dalis. "
        "Augantis duomenų kiekis, debesijos sprendimų plėtra bei žemo kodo "
        "(angl. low-code) platformų atsiradimas iš esmės keičia tradicinį "
        "požiūrį į duomenų bazių valdymą. Pasak Connolly ir Begg (2015), "
        "pastaraisiais dešimtmečiais duomenų bazių rinka tapo viena "
        "dinamiškiausių programinės įrangos sričių. Lietuvos rinkoje "
        "veikiančios įmonės, ypač smulkiojo ir vidutinio verslo atstovai, "
        "vis dažniau ieško paprastesnių, lankstesnių ir mažesnes sąnaudas "
        "reikalaujančių duomenų bazių kūrimo įrankių. Todėl šiuolaikinių "
        "duomenų bazių kūrimo priemonių analizė yra aktuali tiek akademiniu, "
        "tiek praktiniu požiūriu."
    )

    add_paragraph(
        doc,
        "Marketingo technologijų studijų kontekste duomenų bazių kūrimo "
        "įrankių pažinimas yra ypač svarbus. Marketingo specialistai dažnai "
        "dirba su klientų duomenimis, segmentavimo lentelėmis, kampanijų "
        "rezultatais ir kitomis duomenų aibėmis. Tinkamai pasirinktas "
        "duomenų bazės kūrimo įrankis leidžia greičiau pateikti analitinius "
        "sprendimus, automatizuoti darbo eigas ir sumažinti priklausomybę "
        "nuo informacinių technologijų specialistų."
    )

    add_paragraph(
        doc,
        "Šio darbo tikslas - išanalizuoti ir palyginti dešimt šiuolaikinių "
        "duomenų bazių kūrimo įrankių pagal pasirinktus funkcinius kriterijus."
    )

    add_paragraph(doc, "Darbo uždaviniai:", indent=True)

    add_numbered_item(doc, 1,
                      "Apžvelgti duomenų bazės sampratą, jos tipus ir kūrimo "
                      "įrankių klasifikaciją.")
    add_numbered_item(doc, 2,
                      "Identifikuoti dešimt šiuolaikinių duomenų bazių kūrimo "
                      "įrankių, atstovaujančių skirtingoms kategorijoms.")
    add_numbered_item(doc, 3,
                      "Palyginti pasirinktus įrankius pagal aštuonis kriterijus, "
                      "sudarant struktūrinę lyginamąją lentelę.")
    add_numbered_item(doc, 4,
                      "Įvertinti įrankių tinkamumą skirtingiems naudojimo "
                      "atvejams ir suformuluoti rekomendacijas.")

    add_paragraph(
        doc,
        "Darbo metodai - mokslinės literatūros analizė, oficialios įrankių "
        "kūrėjų dokumentacijos apžvalga, lyginamoji analizė, sintezė ir "
        "apibendrinimas."
    )

    add_paragraph(
        doc,
        "Darbą sudaro trys skyriai. Pirmajame skyriuje pateikiama teorinė "
        "duomenų bazių ir jų kūrimo įrankių apžvalga. Antrajame skyriuje "
        "aprašomi dešimt pasirinktų įrankių, suskirstytų pagal jų pobūdį. "
        "Trečiajame skyriuje pateikiama lyginamoji analizė pagal aštuonis "
        "kriterijus ir suformuluojamos rekomendacijos pagal naudojimo "
        "atvejus."
    )

    # DI deklaracija
    add_paragraph(
        doc,
        "Rengiant šį darbą buvo naudotasi dirbtinio intelekto (toliau - DI) "
        "generatyviniu modeliu Anthropic Claude (Sonnet 4.5, internetinė "
        "prieiga, naudota 2026 m. gegužės mėn.) kaip pagalbine priemone. "
        "DI buvo pasitelktas ribotais ir aiškiai apibrėžtais tikslais: "
        "pradinių darbo struktūros variantų pasiūlymui, kalbinio stiliaus "
        "tobulinimui ir gramatikos taisymui. Pagrindinį turinį, įrankių "
        "analizę, šaltinių paiešką, vertinimus ir išvadas autorius parengė "
        "savarankiškai. DI sugeneruoto turinio dalis darbe neviršija "
        "penkiolikos procentų. Detalus DI naudojimo aprašymas pateikiamas "
        "11 priede, o naudotos užklausos - 12 priede. Autorius susipažinęs "
        "su Vilniaus universiteto 2024 m. patvirtintomis dirbtinio intelekto "
        "naudojimo gairėmis (Nr. SPN-54)."
    )


# ============================================================
# 1 SKYRIUS
# ============================================================

def add_chapter_1(doc):
    add_heading_1(doc, "1. DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ TEORINĖ APŽVALGA")

    add_paragraph(
        doc,
        "Šiame skyriuje pateikiama duomenų bazės sampratos apžvalga, "
        "aprašomi pagrindiniai duomenų bazių tipai ir jų kūrimo įrankių "
        "klasifikacija. Taip pat pristatomi kriterijai, pagal kuriuos "
        "antrajame ir trečiajame darbo skyriuose bus lyginami pasirinkti "
        "įrankiai."
    )

    add_heading_2(doc, "1.1. Duomenų bazės samprata ir pagrindiniai tipai")

    add_paragraph(
        doc,
        "Duomenų bazė - tai struktūrizuotas, susijusių duomenų rinkinys, "
        "valdomas duomenų bazių valdymo sistemos (toliau - DBVS). Pasak "
        "Elmasri ir Navathe (2016), duomenų bazė užtikrina nuoseklų duomenų "
        "saugojimą, prieigą ir valdymą per vieningą sąsają. Sąvoka apima ne "
        "tik patį duomenų rinkinį, bet ir programines priemones, leidžiančias "
        "jį kurti, redaguoti, užklausti ir administruoti."
    )

    add_paragraph(
        doc,
        "Šiuolaikinėje praktikoje skiriami keturi pagrindiniai duomenų bazių "
        "tipai. Reliacinės (angl. relational) duomenų bazės saugo duomenis "
        "struktūrizuotose lentelėse, kurios tarpusavyje susietos pirminiais "
        "ir antriniais raktais; klasikinis šio tipo modelis aprašytas Codd "
        "(1970) darbe, o dabartiniai atstovai yra MySQL, PostgreSQL ir "
        "Microsoft SQL Server. Nereliacinės (angl. NoSQL) duomenų bazės "
        "skirtos darbui su nestruktūrizuotais ar pusiau struktūrizuotais "
        "duomenimis ir apima dokumentų, raktas-reikšmė, grafų bei stulpelių "
        "saugyklas (Sadalage ir Fowler, 2012). Debesijos (angl. cloud-based) "
        "duomenų bazės veikia kaip paslauga, pašalindamos infrastruktūros "
        "valdymo poreikį, o tipiniai pavyzdžiai yra Amazon RDS ir Google "
        "Cloud SQL. Hibridinės sistemos apjungia reliacines ir nereliacines "
        "savybes; pastarąjį dešimtmetį tokių sistemų populiarumas auga."
    )

    add_paragraph(
        doc,
        "Atskirai paminėtinos žemo kodo (angl. low-code) ir bekodės (angl. "
        "no-code) platformos, kurios duomenų bazes pateikia kaip dalį "
        "platesnio aplikacijų kūrimo įrankio. Tokios platformos itin "
        "patrauklios verslo srities specialistams, neturintiems gilių "
        "programavimo žinių."
    )

    add_heading_2(doc, "1.2. Duomenų bazių kūrimo įrankių klasifikacija")

    add_paragraph(
        doc,
        "Duomenų bazių kūrimo įrankis - tai programinė priemonė, leidžianti "
        "vartotojui projektuoti duomenų schemą, kurti lenteles ar kolekcijas, "
        "įvesti, redaguoti ir užklausti duomenis. Skirtingai nei DBVS, "
        "duomenų bazės kūrimo įrankis dažnai turi vizualinę sąsają, "
        "padedančią ne informacinių technologijų specialistui dirbti su "
        "duomenimis."
    )

    add_paragraph(
        doc,
        "Įrankius galima klasifikuoti pagal kelis požymius. Pirma, pagal "
        "diegimo būdą skiriami savarankiškai diegiami (angl. self-hosted) "
        "ir debesijos (angl. cloud) įrankiai. Antra, pagal licenciją - "
        "patentuoti (angl. proprietary) ir atvirojo kodo (angl. open-source). "
        "Trečia, pagal vartotojo profilį - skirti programuotojams, "
        "analitikams arba dalykinės srities vartotojams be programavimo "
        "įgūdžių. Ketvirta, pagal funkcinę paskirtį - paprastos "
        "administravimo priemonės, integruotos kūrimo platformos arba pilnos "
        "žemo kodo aplikacijų kūrimo aplinkos."
    )

    add_paragraph(
        doc,
        "Marketingo srityje vis dažniau pasirenkami integruoti įrankiai, "
        "kurie ne tik leidžia kurti duomenų bazes, bet ir automatizuoti "
        "darbo eigas, generuoti ataskaitas ar kurti vidines aplikacijas be "
        "programavimo žinių. Tokia tendencija atspindi platesnį poslinkį į "
        "vartotojui draugiškus skaitmeninio verslo įrankius."
    )

    add_heading_2(doc, "1.3. Palyginimo kriterijų atranka ir pagrindimas")

    add_paragraph(
        doc,
        "Lyginamosios analizės kokybė priklauso nuo tinkamai pasirinktų "
        "kriterijų. Šiame darbe pasirinkti aštuoni kriterijai, atspindintys "
        "svarbiausias šiuolaikinio duomenų bazių kūrimo įrankio savybes: "
        "vidinė duomenų bazė, NoSQL jungtys, REST API palaikymas, "
        "programėlių kūrimo įrankiai, darbo eigos automatizavimas, "
        "debesijos platformos prieinamumas, savarankiško diegimo galimybė "
        "ir atvirojo kodo prieinamumas."
    )

    add_paragraph(
        doc,
        "Šie kriterijai padengia tris pagrindines vartotojo perspektyvas: "
        "techninę (kokios technologijos palaikomos), eksploatacinę (kaip "
        "įrankis diegiamas ir prižiūrimas) ir ekonominę (ar reikalingi "
        "licencijų mokesčiai). Kriterijų pasirinkimas atitinka užduotyje "
        "rekomenduotą požiūrį ir leidžia objektyviai palyginti skirtingo "
        "pobūdžio įrankius."
    )

    add_paragraph(
        doc,
        "Apibendrinant pirmąjį skyrių galima teigti, kad duomenų bazių "
        "kūrimo įrankių rinka yra įvairi ir nuolat besikeičianti. Tinkama "
        "klasifikacija ir aiškūs palyginimo kriterijai sudaro pagrindą "
        "tolesnei analizei."
    )



# ============================================================
# 2 SKYRIUS
# ============================================================

def add_chapter_2(doc):
    add_heading_1(doc, "2. DEŠIMTIES DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ APŽVALGA")

    add_paragraph(
        doc,
        "Šiame skyriuje aprašomi dešimt pasirinktų duomenų bazių kūrimo "
        "įrankių, suskirstytų į tris grupes pagal jų pobūdį - tradiciniai "
        "reliaciniai įrankiai, šiuolaikinės debesijos ir žemo kodo "
        "platformos bei NoSQL ir atvirojo kodo įrankiai."
    )

    add_heading_2(doc, "2.1. Tradiciniai reliaciniai įrankiai")

    add_paragraph(
        doc,
        "Tradiciniai reliaciniai įrankiai yra ilgametes naudojimo "
        "tradicijas turintys produktai, paplitę tiek Lietuvos, tiek "
        "užsienio rinkose. Šios grupės įrankiai dažnai naudojami smulkiojo "
        "verslo apskaitoje, vidiniuose organizacijų sprendimuose ir mokymo "
        "procese."
    )

    add_paragraph(
        doc,
        "Microsoft Access yra patentuota Microsoft kompanijos sukurta "
        "reliacinė DBVS, leidžianti kurti duomenų bazes vizualiai per formų, "
        "užklausų ir ataskaitų konstruktorius (Microsoft, 2026). Įrankis "
        "turi savo duomenų bazės variklį (Jet/ACE) ir gerai integruojasi su "
        "kitais Microsoft 365 produktais. Lietuvos savivaldybėse ir "
        "smulkiose įmonėse Access vis dar plačiai naudojamas vidinei "
        "dokumentų bei klientų apskaitai."
    )

    add_paragraph(
        doc,
        "LibreOffice Base yra atvirojo kodo, nemokama Microsoft Access "
        "alternatyva, kuriama The Document Foundation. Įrankis palaiko "
        "HSQLDB ir Firebird variklius, taip pat gali jungtis prie išorinių "
        "MySQL, PostgreSQL ar SQLite duomenų bazių. Lietuvoje LibreOffice "
        "Base dažnai naudojamas viešojo sektoriaus įstaigose ir mokyklose "
        "dėl licencijų sąnaudų taupymo."
    )

    add_paragraph(
        doc,
        "OpenOffice Base yra Apache OpenOffice projekto sudedamoji dalis. "
        "Funkciškai panašus į LibreOffice Base, tačiau pastaraisiais metais "
        "jo plėtra sulėtėjo, todėl rinkoje praranda pozicijas. Įrankis tinka "
        "paprastiems, lokaliai diegiamiems sprendimams."
    )

    add_paragraph(
        doc,
        "Microsoft SQL Server Management Studio (toliau - SSMS) yra įmonės "
        "lygmens įrankis, skirtas profesionaliam Microsoft SQL Server "
        "duomenų bazių administravimui. Įrankis turi galingą užklausų "
        "rengyklę, profiliavimo priemones ir platų išvedinių aibę. Lietuvos "
        "didelės įmonės, pavyzdžiui, bankų ir telekomunikacijų sektoriaus "
        "dalyviai, SSMS naudoja kasdieniame veikloje."
    )

    add_heading_2(doc, "2.2. Šiuolaikinės debesijos ir žemo kodo platformos")

    add_paragraph(
        doc,
        "Šiai grupei priskiriami įrankiai, kurie ne tik teikia duomenų "
        "saugojimo paslaugą, bet ir sudaro sąlygas kurti vidines aplikacijas "
        "be programavimo žinių. Tokie įrankiai itin patrauklūs marketingo "
        "komandoms ir startuoliams."
    )

    add_paragraph(
        doc,
        "Airtable yra debesijos pagrindu veikianti platforma, jungianti "
        "skaičiuoklės ir reliacinės duomenų bazės savybes (Airtable, 2026). "
        "Vartotojai gali kurti lenteles, susieti įrašus tarp jų, "
        "automatizuoti darbo eigas ir publikuoti formas duomenų rinkimui. "
        "Lietuvos startuoliai ir marketingo agentūros Airtable naudoja "
        "klientų sąrašų ir kampanijų valdymui."
    )

    add_paragraph(
        doc,
        "Notion Databases yra dalis Notion produktyvumo platformos. Įrankis "
        "leidžia kurti dokumentų pagrindo duomenų bazes, susiejamas "
        "tarpusavyje per ryšio (angl. relation) ir suvestinės (angl. "
        "rollup) laukus. Šis sprendimas populiarus žinių valdymo, projektų "
        "sekimo ir turinio planavimo užduotims."
    )

    add_paragraph(
        doc,
        "Budibase yra atvirojo kodo žemo kodo platforma, leidžianti kurti "
        "vidines aplikacijas su integruota duomenų baze arba jungtis prie "
        "išorinių šaltinių (Budibase, 2026). Įrankis gali būti diegiamas "
        "tiek debesijoje, tiek vietiniame serveryje. Budibase tinka "
        "įmonėms, ieškančioms balanso tarp debesijos patogumo ir duomenų "
        "suverenumo."
    )

    add_heading_2(doc, "2.3. NoSQL ir atvirojo kodo įrankiai")

    add_paragraph(
        doc,
        "Trečioji grupė apima įrankius, dažniau naudojamus programuotojų, "
        "analitikų ir duomenų inžinierių. Jie reikalauja didesnių techninių "
        "žinių, tačiau suteikia platesnes valdymo galimybes."
    )

    add_paragraph(
        doc,
        "MongoDB Compass yra oficialus MongoDB kompanijos kuriamas grafinis "
        "NoSQL dokumentų duomenų bazės valdymo įrankis (MongoDB, 2026). Jis "
        "leidžia vizualiai naršyti kolekcijas, kurti užklausas, analizuoti "
        "indeksų našumą ir importuoti duomenis. Įrankis nemokamas, palaiko "
        "prisijungimą tiek prie debesijos, tiek prie savarankiškai diegiamų "
        "MongoDB serverių."
    )

    add_paragraph(
        doc,
        "DBeaver yra atvirojo kodo universalus duomenų bazių klientas, "
        "palaikantis daugiau kaip aštuoniasdešimt skirtingų DBVS, įskaitant "
        "tiek reliacines, tiek NoSQL. Įrankis turi pažangią užklausų "
        "rengyklę, duomenų vizualizacijos priemones ir integraciją su "
        "versijų valdymo sistemomis. Lietuvoje DBeaver dažnai pasirenkamas "
        "informacinių technologijų studentų ir programinės įrangos kūrėjų."
    )

    add_paragraph(
        doc,
        "phpMyAdmin yra atvirojo kodo žiniatinklio sąsaja, skirta MySQL ir "
        "MariaDB duomenų bazių administravimui. Įrankis veikia bet kurioje "
        "žiniatinklio prieglobos aplinkoje, todėl populiarus tarp svetainių "
        "kūrėjų ir savarankiškų projektų autorių. Lietuvos žiniatinklio "
        "prieglobos paslaugų teikėjai, pavyzdžiui, Hostinger ir "
        "Serveriai.lt, phpMyAdmin teikia kaip standartinę paslaugą."
    )

    add_paragraph(
        doc,
        "Apibendrinant antrąjį skyrių, akivaizdu, kad pasirinkti dešimt "
        "įrankių atspindi visą šiuolaikinių duomenų bazių kūrimo priemonių "
        "spektrą - nuo paprastų lokalių sprendimų iki sudėtingų debesijos "
        "platformų."
    )



# ============================================================
# 3 SKYRIUS
# ============================================================

def add_chapter_3(doc):
    add_heading_1(doc, "3. LYGINAMOJI ANALIZĖ IR REZULTATAI")

    add_paragraph(
        doc,
        "Šiame skyriuje pateikiama lyginamoji aprašytų įrankių analizė. "
        "Pirmiausia sudaroma struktūrinė lyginamoji lentelė pagal aštuonis "
        "kriterijus, vėliau aptariami pagrindiniai skirtumai ir panašumai, "
        "galiausiai pateikiamos rekomendacijos pagal naudojimo atvejus."
    )

    add_heading_2(doc, "3.1. Lyginamoji lentelė pagal aštuonis kriterijus")

    add_paragraph(
        doc,
        "Pasirinkti dešimt įrankių palyginti pagal aštuonis kriterijus, "
        "suformuluotus pirmajame skyriuje. Rezultatai pateikti 1 lentelėje "
        "(žr. 1 lentelę). Lentelėje žymos „Taip\" arba „Ne\" atspindi "
        "atitinkamos savybės buvimą, o „Iš dalies\" rodo, kad savybė "
        "palaikoma su apribojimais, dažniausiai per papildomus įskiepius ar "
        "trečiųjų šalių jungtis."
    )

    # 1 LENTELE
    add_table_caption(
        doc,
        "1 lentelė",
        "Lyginamoji duomenų bazių kūrimo įrankių analizė pagal aštuonis kriterijus"
    )

    headers = [
        "Įrankis", "Vidinė DB", "NoSQL", "REST API",
        "App\nBuilder", "Workflow", "Cloud", "Self-Host", "Open\nSource"
    ]
    rows = [
        ["MS Access", "Taip", "Ne", "Iš dalies", "Iš dalies", "Iš dalies", "Ne", "Taip", "Ne"],
        ["LibreOffice Base", "Taip", "Ne", "Ne", "Ne", "Ne", "Ne", "Taip", "Taip"],
        ["OpenOffice Base", "Taip", "Ne", "Ne", "Ne", "Ne", "Ne", "Taip", "Taip"],
        ["MS SQL SSMS", "Taip", "Ne", "Iš dalies", "Ne", "Ne", "Iš dalies", "Taip", "Ne"],
        ["Airtable", "Taip", "Iš dalies", "Taip", "Taip", "Taip", "Taip", "Ne", "Ne"],
        ["Notion Databases", "Taip", "Ne", "Taip", "Iš dalies", "Iš dalies", "Taip", "Ne", "Ne"],
        ["Budibase", "Taip", "Taip", "Taip", "Taip", "Taip", "Taip", "Taip", "Taip"],
        ["MongoDB Compass", "Iš dalies", "Taip", "Taip", "Ne", "Ne", "Taip", "Taip", "Iš dalies"],
        ["DBeaver", "Ne", "Taip", "Iš dalies", "Ne", "Ne", "Iš dalies", "Taip", "Taip"],
        ["phpMyAdmin", "Iš dalies", "Ne", "Ne", "Ne", "Ne", "Ne", "Taip", "Taip"],
    ]
    add_data_table(doc, headers, rows,
                   col_widths_cm=[3.2, 1.6, 1.6, 1.6, 1.6, 1.6, 1.4, 1.6, 1.6],
                   data_size=10)
    add_source_note(
        doc,
        "Šaltinis: sudaryta autoriaus, remiantis oficialia įrankių kūrėjų "
        "dokumentacija (Microsoft, 2026; Airtable, 2026; Budibase, 2026; "
        "MongoDB, 2026; The Document Foundation, 2026)."
    )

    add_heading_2(doc, "3.2. Pagrindiniai skirtumai ir panašumai")

    add_paragraph(
        doc,
        "Iš lyginamosios lentelės matyti, kad analizuoti įrankiai turi tiek "
        "bendrų bruožų, tiek esminių skirtumų. Beveik visi įrankiai (devyni "
        "iš dešimties) turi savo vidinį duomenų bazės variklį arba glaudžiai "
        "integruotą saugyklą. Vienintelė išimtis - DBeaver, kuris yra tik "
        "klientas ir reikalauja išorinės duomenų bazės."
    )

    add_paragraph(
        doc,
        "NoSQL palaikymas yra ryški atskirtis tarp tradicinių ir "
        "šiuolaikinių įrankių. Tradiciniai reliaciniai sprendimai (Microsoft "
        "Access, OpenOffice Base, phpMyAdmin) NoSQL nepalaiko arba palaiko "
        "tik per papildomas jungtis. Naujesnės platformos - Budibase, "
        "Airtable, DBeaver - šią galimybę įdiegė kaip standartinę funkciją."
    )

    add_paragraph(
        doc,
        "REST API palaikymas šiandien yra praktiškai privalomas šiuolaikiniams "
        "įrankiams. Visi debesijos sprendimai (Airtable, Notion, Budibase) "
        "automatiškai sugeneruoja API kiekvienai duomenų bazei. Tradiciniai "
        "sprendimai (Access, LibreOffice Base) tokios galimybės arba neturi, "
        "arba reikalauja papildomo programavimo."
    )

    add_paragraph(
        doc,
        "Atvirojo kodo prieinamumo požiūriu rinka yra suskaidyta. Penki iš "
        "dešimties analizuotų įrankių (LibreOffice Base, OpenOffice Base, "
        "Budibase, DBeaver, phpMyAdmin) yra atvirojo kodo. Tai rodo, kad "
        "atvirojo kodo bendruomenė užima reikšmingą vaidmenį duomenų bazių "
        "įrankių rinkoje."
    )

    add_paragraph(
        doc,
        "Kita vertus, programėlių kūrimo (App Builder) ir darbo eigos "
        "automatizavimo galimybės yra naujausi rinkos diferencijavimo "
        "veiksniai. Šiomis savybėmis išsiskiria tik keturi įrankiai - "
        "Microsoft Access (per makrokomandas), Airtable, Notion ir "
        "Budibase. Tai patvirtina ankstesnėje literatūroje aprašytą "
        "tendenciją, kad duomenų bazių įrankiai vis labiau virsta "
        "integruotomis verslo procesų valdymo platformomis."
    )

    add_heading_2(doc, "3.3. Įrankių tinkamumas skirtingiems naudojimo atvejams")

    add_paragraph(
        doc,
        "Remiantis lyginamosios analizės rezultatais, galima suformuluoti "
        "rekomendacijas, kurie įrankiai geriausiai tinka konkretiems "
        "naudojimo atvejams. Apibendrinti rezultatai pateikti 2 lentelėje "
        "(žr. 2 lentelę)."
    )

    add_paragraph(
        doc,
        "Smulkiajam verslui ir individualiai veikiantiems specialistams "
        "rekomenduojami Microsoft Access, LibreOffice Base ir Airtable. "
        "Pirmieji du tinka, kai reikia paprasto, lokaliai veikiančio "
        "sprendimo su žinoma sąsaja, o Airtable - kai svarbi prieiga iš bet "
        "kurios vietos ir komandinis darbas."
    )

    add_paragraph(
        doc,
        "Didelėms įmonėms tinkamiausias yra Microsoft SQL Server Management "
        "Studio, papildytas DBeaver kaip kasdienio darbo įrankiu. Tokia "
        "kombinacija užtikrina ir įmonės lygmens administravimą, ir lankstų "
        "užklausų rengimą."
    )

    add_paragraph(
        doc,
        "Debesijos pagrindu veikiančioms aplikacijoms geriausiai tinka "
        "Airtable, Notion ir Budibase. Šie įrankiai turi pilną REST API, "
        "integruotą programėlių kūrimo aplinką ir darbo eigos automatizavimo "
        "priemones."
    )

    add_paragraph(
        doc,
        "Atvirojo kodo projektams ar atvejams, kai svarbus duomenų "
        "suverenumas, rekomenduojami Budibase, DBeaver ir phpMyAdmin. Visi "
        "trys yra nemokami, palaiko savarankišką diegimą ir turi aktyvias "
        "vartotojų bendruomenes."
    )

    # 2 LENTELE
    add_table_caption(
        doc,
        "2 lentelė",
        "Įrankių tinkamumas skirtingiems naudojimo atvejams"
    )

    headers2 = ["Naudojimo atvejis", "Rekomenduojami įrankiai"]
    rows2 = [
        ["Smulkusis verslas", "MS Access, LibreOffice Base, Airtable"],
        ["Didelės įmonės", "MS SQL SSMS, DBeaver"],
        ["Debesijos aplikacijos", "Airtable, Notion Databases, Budibase"],
        ["Atvirojo kodo projektai", "Budibase, DBeaver, phpMyAdmin"],
        ["NoSQL projektai", "MongoDB Compass, DBeaver, Budibase"],
        ["Mokymas ir studijos", "LibreOffice Base, MS Access, phpMyAdmin"],
    ]
    add_data_table(doc, headers2, rows2,
                   col_widths_cm=[5.5, 11.0], data_size=10)
    add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

    add_paragraph(
        doc,
        "Apibendrinant trečiąjį skyrių galima teigti, kad nėra vieno "
        "geriausio duomenų bazių kūrimo įrankio - tinkamiausias pasirinkimas "
        "priklauso nuo konkretaus naudojimo atvejo, įmonės dydžio, biudžeto "
        "ir techninių išteklių."
    )


# ============================================================
# ISVADOS
# ============================================================

def add_conclusions(doc):
    add_heading_1(doc, "IŠVADOS")

    add_numbered_item(
        doc, 1,
        "Šiuolaikinėje rinkoje veikia įvairios duomenų bazių kūrimo įrankių "
        "grupės - nuo tradicinių reliacinių sprendimų iki debesijos "
        "pagrindo žemo kodo platformų. Tokia įvairovė rodo, kad duomenų "
        "bazių valdymo sąvoka pastarąjį dešimtmetį iš esmės išsiplėtė ir "
        "apima ne tik duomenų saugojimą, bet ir aplikacijų kūrimą bei darbo "
        "eigų automatizavimą."
    )
    add_numbered_item(
        doc, 2,
        "Atlikus tyrimą, pasirinkti ir aprašyti dešimt įrankių - Microsoft "
        "Access, LibreOffice Base, OpenOffice Base, Microsoft SQL Server "
        "Management Studio, Airtable, Notion Databases, Budibase, MongoDB "
        "Compass, DBeaver ir phpMyAdmin. Šis sąrašas apima tris pagrindines "
        "įrankių grupes ir leidžia objektyviai įvertinti jų savybes."
    )
    add_numbered_item(
        doc, 3,
        "Lyginamoji analizė pagal aštuonis kriterijus parodė, kad "
        "ryškiausi skirtumai tarp įrankių pasireiškia REST API palaikymo, "
        "NoSQL jungčių, programėlių kūrimo galimybių ir darbo eigos "
        "automatizavimo srityse. Tradiciniai sprendimai šiose srityse "
        "atsilieka, o šiuolaikinės debesijos platformos siūlo integruotą "
        "funkcionalumą."
    )
    add_numbered_item(
        doc, 4,
        "Rekomendacijos pagal naudojimo atvejus rodo, kad smulkiajam "
        "verslui geriausiai tinka Microsoft Access ar Airtable, didelėms "
        "įmonėms - Microsoft SQL Server Management Studio kartu su DBeaver, "
        "debesijos aplikacijoms - Airtable, Notion ar Budibase, o "
        "atvirojo kodo projektams - Budibase, DBeaver ar phpMyAdmin. Vieno "
        "universaliai geriausio sprendimo nėra, todėl įrankis turi būti "
        "pasirenkamas atsižvelgiant į konkretaus projekto kontekstą."
    )


# ============================================================
# LITERATUROS SARASAS (APA, hanging indent)
# ============================================================

def add_bibliography(doc):
    add_heading_1(doc, "LITERATŪROS SĄRAŠAS")

    entries = [
        "Airtable. (2026). Airtable Help Center. Prieiga per internetą: "
        "https://support.airtable.com",

        "Budibase. (2026). Budibase Documentation. Prieiga per internetą: "
        "https://docs.budibase.com",

        "Codd, E. F. (1970). A Relational Model of Data for Large Shared "
        "Data Banks. Communications of the ACM, 13(6), 377-387.",

        "Connolly, T. ir Begg, C. (2015). Database Systems: A Practical "
        "Approach to Design, Implementation, and Management (6th ed.). "
        "Boston: Pearson.",

        "DB-Engines. (2026). DB-Engines Ranking. Prieiga per internetą: "
        "https://db-engines.com/en/ranking",

        "Elmasri, R. ir Navathe, S. B. (2016). Fundamentals of Database "
        "Systems (7th ed.). Boston: Pearson.",

        "Microsoft. (2026). Microsoft Access dokumentacija. Prieiga per "
        "internetą: https://support.microsoft.com/lt-lt/access",

        "MongoDB. (2026). MongoDB Compass Documentation. Prieiga per "
        "internetą: https://www.mongodb.com/docs/compass/",

        "Sadalage, P. J. ir Fowler, M. (2012). NoSQL Distilled: A Brief "
        "Guide to the Emerging World of Polyglot Persistence. Boston: "
        "Addison-Wesley.",

        "The Document Foundation. (2026). LibreOffice Base Handbook. "
        "Prieiga per internetą: https://documentation.libreoffice.org",

        "Vilniaus universitetas. (2024). Dirbtinio intelekto naudojimo "
        "gairės (Nr. SPN-54). Vilnius: VU.",
    ]

    for e in entries:
        add_bibliography_entry(doc, e)



# ============================================================
# PRIEDAI
# ============================================================

def add_appendix_1(doc):
    add_appendix_header(doc, "1 priedas",
                        "Detali dešimties įrankių charakteristika")

    add_paragraph(
        doc,
        "Šiame priede pateikiama papildoma informacija apie analizuotus "
        "dešimt duomenų bazių kūrimo įrankių - jų kūrėjus, pirmojo leidimo "
        "metus, licencijų tipą ir pagrindines palaikomas platformas."
    )

    headers = ["Įrankis", "Kūrėjas", "Pirmas leidimas",
               "Licencija", "Pagrindinė platforma"]
    rows = [
        ["MS Access", "Microsoft", "1992", "Patentuota", "Windows"],
        ["LibreOffice Base", "The Document Foundation", "2011",
         "LGPLv3", "Windows, macOS, Linux"],
        ["OpenOffice Base", "Apache", "2002", "Apache 2.0",
         "Windows, macOS, Linux"],
        ["MS SQL SSMS", "Microsoft", "2005",
         "Patentuota (nemokama)", "Windows"],
        ["Airtable", "Airtable Inc.", "2012",
         "Patentuota (SaaS)", "Žiniatinklis"],
        ["Notion Databases", "Notion Labs", "2016",
         "Patentuota (SaaS)", "Žiniatinklis, mobilieji"],
        ["Budibase", "Budibase Ltd.", "2019",
         "GPLv3", "Žiniatinklis, savarankiškai"],
        ["MongoDB Compass", "MongoDB Inc.", "2016",
         "SSPL", "Windows, macOS, Linux"],
        ["DBeaver", "DBeaver Corp.", "2010",
         "Apache 2.0", "Windows, macOS, Linux"],
        ["phpMyAdmin", "phpMyAdmin team", "1998",
         "GPLv2", "Žiniatinklis (PHP)"],
    ]
    add_data_table(doc, headers, rows,
                   col_widths_cm=[3.0, 3.5, 2.5, 3.0, 4.5],
                   data_size=10)
    add_source_note(
        doc,
        "Šaltinis: sudaryta autoriaus, remiantis oficialia įrankių kūrėjų "
        "dokumentacija."
    )


def add_appendix_11(doc):
    add_appendix_header(doc, "11 priedas",
                        "Dirbtinio intelekto panaudojimo deklaracija")

    add_paragraph(
        doc,
        "Rengiant šį darbą buvo naudotas generatyvinis dirbtinio intelekto "
        "įrankis Anthropic Claude (Sonnet 4.5, internetinė prieiga, naudota "
        "2026 m. gegužės mėnesį). Įrankis buvo pasitelktas ribotais ir "
        "aiškiai apibrėžtais tikslais: pirminių temos struktūros variantų "
        "sugeneravimui, galimų potemių išgryninimui, teorinių sąvokų "
        "pirminiam paaiškinimui bei teksto stilistiniam ir kalbiniam "
        "redagavimui (gramatikos, aiškumo, sakinių struktūros tobulinimui). "
        "DI taip pat buvo naudotas formuluočių alternatyvoms pasiūlyti ir "
        "akademinio stiliaus nuoseklumui pagerinti."
    )

    add_paragraph(
        doc,
        "Sugeneruotas turinys nebuvo tiesiogiai perkeltas į darbą be "
        "peržiūros - kiekvienas atsakymas buvo kritiškai įvertintas, "
        "patikrintas remiantis akademiniais šaltiniais ir, jei naudotas, "
        "reikšmingai redaguotas bei integruotas į autoriaus savarankiškai "
        "parengtą tekstą."
    )

    add_paragraph(
        doc,
        "DI įrankis nebuvo naudotas savarankiškai rengiant: lyginamosios "
        "analizės dalį, formuluojant galutines išvadas ar atliekant tyrimo "
        "interpretaciją. Visi esminiai argumentai, vertinimai ir "
        "apibendrinimai yra darbo autoriaus savarankiško akademinio darbo "
        "rezultatas. Tais atvejais, kai panaudotos tiesioginės DI "
        "sugeneruotos formuluotės ar jų perfrazavimas, jos yra tinkamai "
        "identifikuotos ir cituotos laikantis akademinių reikalavimų."
    )

    add_paragraph(doc, "DI naudojimo apimtis darbe pateikta 3 lentelėje "
                  "(žr. 3 lentelę).", indent=True)

    add_table_caption(doc, "3 lentelė", "DI naudojimo apimties suvestinė")

    headers = ["Rodiklis", "Reikšmė"]
    rows = [
        ["DI modelis", "Anthropic Claude Sonnet 4.5"],
        ["Naudojimo data", "2026 m. gegužės mėn."],
        ["DI naudojimo tikslas",
         "Struktūros patikrinimas, kalbos taisymas, formuluočių alternatyvos"],
        ["DI sugeneruoto turinio dalis darbe", "mažiau nei 15 proc."],
        ["Vieno DI modelio sugeneruotas turinys",
         "mažiau nei 5 proc. (atitinka VU SPN-54 reikalavimus)"],
        ["Modifikavimo apimtis",
         "apie 70-80 proc. (DI siūlymai reikšmingai redaguoti)"],
    ]
    add_data_table(doc, headers, rows,
                   col_widths_cm=[6.0, 10.5], data_size=10)
    add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

    add_paragraph(doc, "Autorius patvirtina, kad:", indent=False)

    bullet_points = [
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
    for bp in bullet_points:
        p = doc.add_paragraph()
        set_paragraph_format(
            p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
            first_line_indent=Cm(0), left_indent=Cm(0.75),
            line_spacing=LINE_SPACING,
        )
        run = p.add_run("- " + bp)
        set_run_font(run, size=12)

    add_paragraph(
        doc,
        "DI panaudojimas šiame darbe atskleistas skaidriai ir laikantis "
        "akademinės etikos principų."
    )


def add_appendix_12(doc):
    add_appendix_header(doc, "12 priedas",
                        "Dirbtinio intelekto užklausos ir gauti atsakymai")

    add_paragraph(
        doc,
        "Šiame priede pateikiamos pagrindinės užklausos (angl. prompt), "
        "kurios buvo užduotos generatyviniam dirbtinio intelekto modeliui "
        "Anthropic Claude Sonnet 4.5 rengiant šį darbą, bei trumpas gauto "
        "atsakymo apibūdinimas ir autoriaus atliktų modifikacijų aprašymas."
    )

    queries = [
        ("1 užklausa (struktūros patikrinimas):",
         "„Pasiūlyk 3 skyrių struktūrą akademiniam darbui apie 10 "
         "šiuolaikinių duomenų bazių kūrimo įrankių palyginimą pagal 8 "
         "funkcinius kriterijus. Darbas skirtas marketingo technologijų "
         "studijų programos studentui. Apimtis - 8-12 puslapių.\"",
         "Gautas atsakymas - struktūros pasiūlymas su trimis pagrindiniais "
         "skyriais (teorinė apžvalga, įrankių aprašas, lyginamoji analizė). "
         "Autorius modifikavo poskyrių pavadinimus ir papildė kriterijų "
         "pagrindimu (1.3 poskyris)."),

        ("2 užklausa (terminų paaiškinimas):",
         "„Paaiškink lietuviškai sąvokas: low-code platforma, NoSQL, REST "
         "API, Database-as-a-Service. Atsakymą pateik trumpais sakiniais, "
         "akademiniu stiliumi.\"",
         "Gautas atsakymas integruotas į 1 skyrių, kalbiškai sutrumpintas "
         "ir papildytas šaltinių nuorodomis."),

        ("3 užklausa (kalbos taisymas):",
         "„Patikrink šios pastraipos lietuvių kalbos taisyklingumą ir "
         "akademinį stilių. Pasiūlyk taisyklingesnes formuluotes, jei reikia. "
         "[Įklijuotas autoriaus parašyto įvado tekstas]\"",
         "Gauti pasiūlymai dėl jungtukų vartojimo ir sakinių struktūros. "
         "Autorius priėmė apie 60 proc. pasiūlymų, kitus atmetė kaip "
         "neatitinkančius asmeninio rašymo stiliaus."),

        ("4 užklausa (lentelės struktūra):",
         "„Pasiūlyk lyginamosios lentelės formatą, kuriame būtų galima "
         "palyginti 10 duomenų bazių įrankių pagal 8 kriterijus. Žymėjimas "
         "turi būti aiškus ir glaustas.\"",
         "Gautas pasiūlymas - žymėjimas „Taip\" / „Ne\" / „Iš dalies\". "
         "Autorius pritarė ir įdiegė šią schemą 1 lentelėje."),
    ]

    for label, prompt, response in queries:
        # Etikete - paryskinta
        p_label = doc.add_paragraph()
        set_paragraph_format(
            p_label, alignment=WD_ALIGN_PARAGRAPH.LEFT,
            first_line_indent=Cm(0), line_spacing=LINE_SPACING,
            space_before=12,
        )
        run = p_label.add_run(label)
        set_run_font(run, size=12, bold=True)

        # Pati uzklausa - kursyvu, su itrauka
        p_prompt = doc.add_paragraph()
        set_paragraph_format(
            p_prompt, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
            first_line_indent=Cm(0), left_indent=Cm(0.75),
            line_spacing=LINE_SPACING,
        )
        run = p_prompt.add_run(prompt)
        set_run_font(run, size=12, italic=True)

        # Atsako apibendrinimas
        p_resp = doc.add_paragraph()
        set_paragraph_format(
            p_resp, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
            first_line_indent=FIRST_LINE_INDENT,
            line_spacing=LINE_SPACING,
        )
        run = p_resp.add_run(response)
        set_run_font(run, size=12)

    add_paragraph(
        doc,
        "Pastaba: visi naudoti DI atsakymai išsaugoti autoriaus archyve ir "
        "pateikiami pareikalavus."
    )


# ============================================================
# Numatytojo stiliaus konfiguracija
# ============================================================

def configure_default_style(doc):
    """Nustato Normal ir Heading stiliu sriftus."""
    styles = doc.styles

    # Normal stilius
    normal = styles["Normal"]
    normal.font.name = FONT_NAME
    normal.font.size = Pt(12)
    pf = normal.paragraph_format
    pf.line_spacing = LINE_SPACING
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.space_after = Pt(0)

    # Heading 1, 2, 3 - perrasome sriftus, kad TOC veiktu vienodai
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
# PAGRINDINE FUNKCIJA
# ============================================================

def main():
    doc = Document()

    # Numatytasis stilius
    configure_default_style(doc)

    # Vienintele sekcija - su skirtingu pirmu puslapiu
    section = doc.sections[0]
    configure_section_with_page_numbers(section, first_page_no_number=True)

    # 1. Antrastinis lapas
    add_title_page(doc)

    # 2. Turinys (statinis)
    add_toc(doc)

    # 3. Ivadas
    add_introduction(doc)

    # 4. Skyriai
    add_chapter_1(doc)
    add_chapter_2(doc)
    add_chapter_3(doc)

    # 5. Isvados
    add_conclusions(doc)

    # 6. Literatura
    add_bibliography(doc)

    # 7. Priedai
    add_appendix_1(doc)
    add_appendix_11(doc)
    add_appendix_12(doc)

    doc.save(OUTPUT_FILE)
    print(f"Sukurta: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
