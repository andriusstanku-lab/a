"""
PU1 darbo Word dokumento generavimas — paprastas ir kompaktiškas variantas.
Atitinka Akademinių rašto darbų metodinius nurodymus (VU Kauno fakultetas).
"""

from docx import Document
from docx.shared import Pt, Cm, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ---------- Konfigūracija ----------
FONT_NAME = "Times New Roman"
FONT_SIZE_BODY = 12
FONT_SIZE_H1 = 14
FONT_SIZE_H2 = 12
FONT_SIZE_TABLE_TITLE = 11
FONT_SIZE_TABLE_DATA = 10
FONT_SIZE_SOURCE = 9


# ---------- Pagalbinės funkcijos ----------
def set_run_font(run, name=FONT_NAME, size=FONT_SIZE_BODY, bold=False, italic=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    rFonts.set(qn('w:cs'), name)
    rFonts.set(qn('w:eastAsia'), name)


def set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         line_spacing=1.5, first_line_indent=Cm(1.25),
                         space_before=0, space_after=0):
    pf = p.paragraph_format
    pf.alignment = alignment
    pf.line_spacing = line_spacing
    if first_line_indent is not None:
        pf.first_line_indent = first_line_indent
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)


def add_body(doc, text, indent=True):
    p = doc.add_paragraph()
    set_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        line_spacing=1.5,
        first_line_indent=Cm(1.25) if indent else None,
    )
    r = p.add_run(text)
    set_run_font(r, size=FONT_SIZE_BODY)
    return p


def add_h1(doc, text, page_break_before=True):
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 1']
    set_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        line_spacing=1.5,
        first_line_indent=None,
        space_before=0, space_after=12,
    )
    if page_break_before:
        p.paragraph_format.page_break_before = True
    r = p.add_run(text.upper())
    set_run_font(r, size=FONT_SIZE_H1, bold=True)
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p


def add_h2(doc, text):
    p = doc.add_paragraph()
    p.style = doc.styles['Heading 2']
    set_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        line_spacing=1.5,
        first_line_indent=None,
        space_before=12, space_after=6,
    )
    r = p.add_run(text)
    set_run_font(r, size=FONT_SIZE_H2, bold=True)
    r.font.color.rgb = RGBColor(0, 0, 0)
    return p


def add_table_caption(doc, number_text, title_text):
    p_num = doc.add_paragraph()
    set_paragraph_format(p_num, alignment=WD_ALIGN_PARAGRAPH.RIGHT,
                         line_spacing=1.0, first_line_indent=None,
                         space_before=6, space_after=0)
    r1 = p_num.add_run(number_text)
    set_run_font(r1, size=FONT_SIZE_TABLE_TITLE, bold=True)

    p_title = doc.add_paragraph()
    set_paragraph_format(p_title, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         line_spacing=1.0, first_line_indent=None,
                         space_before=0, space_after=6)
    r2 = p_title.add_run(title_text)
    set_run_font(r2, size=FONT_SIZE_TABLE_TITLE, bold=True)


def add_source_note(doc, text):
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                         line_spacing=1.0, first_line_indent=None,
                         space_before=2, space_after=12)
    r = p.add_run(text)
    set_run_font(r, size=FONT_SIZE_SOURCE, italic=True)


def add_picture(doc, image_path, width_cm=15.0):
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         line_spacing=1.0, first_line_indent=None,
                         space_before=8, space_after=0)
    run = p.add_run()
    run.add_picture(image_path, width=Cm(width_cm))


def add_picture_caption(doc, number_text, title_text):
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         line_spacing=1.0, first_line_indent=None,
                         space_before=4, space_after=2)
    r = p.add_run(f"{number_text}. {title_text}")
    set_run_font(r, size=FONT_SIZE_TABLE_TITLE, bold=True)


def add_list_item(doc, text, style='List Number'):
    p = doc.add_paragraph(style=style)
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         line_spacing=1.5, first_line_indent=None,
                         space_before=0, space_after=0)
    for r in p.runs:
        r.text = ''
    r = p.add_run(text)
    set_run_font(r, size=FONT_SIZE_BODY)


def add_field(paragraph, field_code):
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = field_code
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run = paragraph.add_run()
    run._element.append(fldChar1)
    run._element.append(instrText)
    run._element.append(fldChar2)
    run._element.append(fldChar3)
    return run


def add_centered_line(doc, text, size, bold=False, space_after_pt=0):
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         line_spacing=1.0, first_line_indent=None,
                         space_before=0, space_after=space_after_pt)
    r = p.add_run(text)
    set_run_font(r, size=size, bold=bold)
    return p


def add_toc_line(doc, left_text, right_text):
    """Statinė turinio eilutė: tekstas kairėje, puslapio numeris dešinėje."""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                         line_spacing=1.5, first_line_indent=None,
                         space_before=0, space_after=0)
    # Tab stop dešinėje pusėje
    pf = p.paragraph_format
    tab_stops = pf.tab_stops
    tab_stops.add_tab_stop(Cm(16.5), alignment=WD_ALIGN_PARAGRAPH.RIGHT)
    r1 = p.add_run(left_text)
    set_run_font(r1, size=FONT_SIZE_BODY)
    r2 = p.add_run(f"\t{right_text}")
    set_run_font(r2, size=FONT_SIZE_BODY)
    return p


# ---------- Dokumento kūrimas ----------
doc = Document()

# Numatytasis stilius
style = doc.styles['Normal']
style.font.name = FONT_NAME
style.font.size = Pt(FONT_SIZE_BODY)
rPr = style.element.get_or_add_rPr()
rFonts = rPr.find(qn('w:rFonts'))
if rFonts is None:
    rFonts = OxmlElement('w:rFonts')
    rPr.append(rFonts)
rFonts.set(qn('w:ascii'), FONT_NAME)
rFonts.set(qn('w:hAnsi'), FONT_NAME)
rFonts.set(qn('w:cs'), FONT_NAME)
rFonts.set(qn('w:eastAsia'), FONT_NAME)

# Paraštės: viršus 20mm, apačia 20mm, kairė 25mm, dešinė 15mm
section = doc.sections[0]
section.page_height = Mm(297)
section.page_width = Mm(210)
section.top_margin = Mm(20)
section.bottom_margin = Mm(20)
section.left_margin = Mm(25)
section.right_margin = Mm(15)
section.header_distance = Mm(12.5)
section.footer_distance = Mm(12.5)

# Heading stiliai
for style_name, sz, bold in [('Heading 1', FONT_SIZE_H1, True), ('Heading 2', FONT_SIZE_H2, True)]:
    s = doc.styles[style_name]
    s.font.name = FONT_NAME
    s.font.size = Pt(sz)
    s.font.bold = bold
    s.font.color.rgb = RGBColor(0, 0, 0)
    rPr = s.element.get_or_add_rPr()
    rFonts2 = rPr.find(qn('w:rFonts'))
    if rFonts2 is None:
        rFonts2 = OxmlElement('w:rFonts')
        rPr.append(rFonts2)
    rFonts2.set(qn('w:ascii'), FONT_NAME)
    rFonts2.set(qn('w:hAnsi'), FONT_NAME)
    rFonts2.set(qn('w:cs'), FONT_NAME)
    rFonts2.set(qn('w:eastAsia'), FONT_NAME)


# ============================================
# 1. ANTRAŠTINIS LAPAS
# ============================================
add_centered_line(doc, "VILNIAUS UNIVERSITETO", FONT_SIZE_H1, bold=True)
add_centered_line(doc, "KAUNO FAKULTETAS", FONT_SIZE_H1, bold=True, space_after_pt=30)
add_centered_line(doc, "SOCIALINIŲ MOKSLŲ IR TAIKOMOSIOS INFORMATIKOS INSTITUTAS",
                  FONT_SIZE_H1, bold=False, space_after_pt=154)
add_centered_line(doc, "Marketingo technologijų studijų programa",
                  FONT_SIZE_BODY, bold=False, space_after_pt=90)
add_centered_line(doc, "ANDRIUS VARGONAS", FONT_SIZE_BODY, bold=True, space_after_pt=90)
add_centered_line(doc, "INFORMACINIŲ SISTEMŲ IR DUOMENŲ BAZIŲ KŪRIMO REIKALAVIMŲ PALYGINIMAS",
                  FONT_SIZE_H1, bold=True, space_after_pt=30)
add_centered_line(doc, "Informacijos sistemos ir duomenų bazės",
                  FONT_SIZE_BODY, bold=False, space_after_pt=240)
add_centered_line(doc, "Kaunas 2026", FONT_SIZE_BODY, bold=False)

# ----- Sekcijos lūžis: titulinis lapas atskiras nuo likusio dokumento -----
new_section = doc.add_section(WD_SECTION.NEW_PAGE)
new_section.page_height = Mm(297)
new_section.page_width = Mm(210)
new_section.top_margin = Mm(20)
new_section.bottom_margin = Mm(20)
new_section.left_margin = Mm(25)
new_section.right_margin = Mm(15)
new_section.header_distance = Mm(12.5)
new_section.footer_distance = Mm(12.5)
new_section.footer.is_linked_to_previous = False
doc.sections[0].footer.is_linked_to_previous = False

# Puslapio numeris dešinėje pusėje footeryje
footer_p = new_section.footer.paragraphs[0]
footer_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
add_field(footer_p, "PAGE \\* MERGEFORMAT")
for r in footer_p.runs:
    set_run_font(r, size=FONT_SIZE_BODY)

# Pradėti numeravimą nuo 2 (titulinis = 1, bet jame nematomas)
sectPr = new_section._sectPr
pgNumType = OxmlElement('w:pgNumType')
pgNumType.set(qn('w:start'), '2')
sectPr.append(pgNumType)


# ============================================
# 2. TURINYS (statinis, ranka surašytas, be Word laukų)
# ============================================
p_t = doc.add_paragraph()
set_paragraph_format(p_t, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                     line_spacing=1.5, first_line_indent=None,
                     space_before=0, space_after=12)
r_t = p_t.add_run("TURINYS")
set_run_font(r_t, size=FONT_SIZE_H1, bold=True)

# Statinė turinio struktūra (puslapių numeriai numatomi pagal turinio kiekį)
add_toc_line(doc, "ĮVADAS", "3")
add_toc_line(doc, "1. INFORMACINIŲ SISTEMŲ IR DUOMENŲ BAZIŲ SAMPRATA", "4")
add_toc_line(doc, "    1.1. Informacinės sistemos apibrėžimas", "4")
add_toc_line(doc, "    1.2. Duomenų bazės apibrėžimas", "4")
add_toc_line(doc, "2. IS IR DB REIKALAVIMŲ PALYGINIMAS", "5")
add_toc_line(doc, "    2.1. Palyginamosios analizės lentelė", "5")
add_toc_line(doc, "    2.2. Pigu.lt ir Omniva Lietuva pavyzdžių analizė", "6")
add_toc_line(doc, "    2.3. Rašytinė palyginamoji analizė", "7")
add_toc_line(doc, "3. IS IR DB REIKALAVIMŲ RINKIMO ŠABLONAI", "8")
add_toc_line(doc, "    3.1. IS reikalavimų specifikacijos šablonai", "8")
add_toc_line(doc, "    3.2. DB reikalavimų rinkimo šablonai", "8")
add_toc_line(doc, "IŠVADOS", "10")
add_toc_line(doc, "LITERATŪROS SĄRAŠAS", "11")


# ============================================
# 3. ĮVADAS
# ============================================
add_h1(doc, "Įvadas")

add_body(doc,
    "Šių dienų skaitmeninėje aplinkoje informacinės sistemos (toliau – IS) ir "
    "duomenų bazės (toliau – DB) yra neatsiejami komponentai, užtikrinantys "
    "efektyvų organizacijų darbą ir teikiamų paslaugų kokybę. Nors abi "
    "technologijos yra glaudžiai susijusios, jų kūrimo reikalavimai skiriasi, "
    "atspindėdami skirtingas paskirtis ir naudotojų grupes."
)

add_body(doc,
    "Šio darbo tikslas – pasinaudojant interneto pagalba palyginti informacinių "
    "sistemų ir duomenų bazių kūrimo reikalavimus pagal pagrindinius aspektus "
    "ir pateikti realius pavyzdžius iš Lietuvos rinkos."
)

add_body(doc, "Darbo uždaviniai:")

add_list_item(doc, "Apžvelgti IS ir DB sampratą.")
add_list_item(doc, "Atlikti reikalavimų palyginamąją analizę pateikiant lentelę.")
add_list_item(doc, "Išanalizuoti realų pavyzdį – Pigu.lt ir Omniva Lietuva.")
add_list_item(doc, "Aprašyti pagrindinius reikalavimų rinkimo šablonus.")

add_body(doc,
    "Darbo metodai: literatūros ir interneto šaltinių analizė, palyginamoji "
    "analizė, atvejo analizė. Darbą sudaro trys skyriai: pirmajame pristatoma "
    "IS ir DB samprata, antrajame atliekamas palyginimas su realiu pavyzdžiu, "
    "trečiajame aptariami reikalavimų rinkimo šablonai."
)


# ============================================
# 4. 1 SKYRIUS — SAMPRATA
# ============================================
add_h1(doc, "1. Informacinių sistemų ir duomenų bazių samprata")

add_body(doc,
    "Šiame skyriuje aptariami pagrindiniai informacinių sistemų ir duomenų "
    "bazių apibrėžimai bei jų paskirtis."
)

# 1.1
add_h2(doc, "1.1. Informacinės sistemos apibrėžimas")

add_body(doc,
    "Informacinė sistema – tai organizuotas žmonių, procesų, programinės bei "
    "techninės įrangos rinkinys, skirtas duomenims rinkti, apdoroti, saugoti "
    "ir paskirstyti, siekiant palaikyti organizacijos veiklą ir sprendimų "
    "priėmimą (Laudon ir Laudon, 2020). IS aprėpia ne tik technologinį "
    "komponentą, bet ir verslo logiką bei vartotojo sąsajas. Tipiniai pavyzdžiai – "
    "elektroninės prekybos platformos, klientų valdymo sistemos."
)

# 1.2
add_h2(doc, "1.2. Duomenų bazės apibrėžimas")

add_body(doc,
    "Duomenų bazė – tai struktūruotas duomenų rinkinys, skirtas efektyviam "
    "duomenų saugojimui, paieškai ir valdymui (Connolly ir Begg, 2015). DB "
    "veikia per duomenų bazių valdymo sistemą (DBVS), kuri užtikrina duomenų "
    "vientisumą ir saugumą. Populiariausios DBVS yra MySQL, PostgreSQL, "
    "Microsoft SQL Server ir Oracle."
)

add_body(doc,
    "Apibendrinant galima teigti, kad IS naudoja DB kaip duomenų saugyklą, o "
    "DB struktūra priklauso nuo IS keliamų reikalavimų."
)


# ============================================
# 5. 2 SKYRIUS — PALYGINIMAS
# ============================================
add_h1(doc, "2. IS ir DB reikalavimų palyginimas")

add_body(doc,
    "Šiame skyriuje pateikiama struktūrinė palyginamoji lentelė, Pigu.lt ir "
    "Omniva Lietuva atvejo analizė bei rašytinė palyginamoji analizė."
)

# 2.1
add_h2(doc, "2.1. Palyginamosios analizės lentelė")

add_body(doc,
    "Pirmoje lentelėje (žr. 1 lentelę) pateikiamas struktūrinis IS ir DB "
    "kūrimo reikalavimų palyginimas pagal aspektus, nurodytus užduoties "
    "metodiniuose nurodymuose."
)

add_table_caption(doc, "1 lentelė", "IS ir DB kūrimo reikalavimų palyginimas")

table_data = [
    ["Aspektas", "Informacinės sistemos (IS)", "Duomenų bazės (DB)"],
    ["Taikymo sritis",
     "Verslo procesų valdymas, paslaugų teikimas, vartotojų sąveika.",
     "Struktūruotas duomenų saugojimas ir greita paieška."],
    ["Funkcionalumas",
     "Funkcijos verslo procesams, grafinė sąsaja, integracijos.",
     "Duomenų įvedimas, atnaujinimas, šalinimas, paieška, ataskaitos."],
    ["Duomenų valdymas",
     "Duomenų srautai tarp modulių, verslo logikos taikymas.",
     "Duomenų schema, lentelės, ryšiai, normalizacija."],
    ["Saugumas",
     "Vartotojų autentifikacija, prieigos lygiai, BDAR atitiktis.",
     "Šifravimas, atsarginės kopijos, prieigos teisės."],
    ["Vartotojai",
     "Klientai, darbuotojai, administratoriai – per grafinę sąsają.",
     "DB administratoriai ir programos – per SQL ar API."],
    ["Kūrimo procesas",
     "Reikalavimų rinkimas, projektavimas, kūrimas, testavimas, diegimas.",
     "ER diagrama → loginė schema → fizinė schema DBVS."],
    ["Realus pavyzdys",
     "Pigu.lt – elektroninės prekybos platforma.",
     "Omniva Lietuva siuntų sekimo duomenų bazė."],
]

table = doc.add_table(rows=len(table_data), cols=3)
table.style = 'Table Grid'
table.alignment = WD_ALIGN_PARAGRAPH.CENTER
table.autofit = False

col_widths = [Cm(3.8), Cm(6.3), Cm(6.3)]
for row in table.rows:
    for idx, cell in enumerate(row.cells):
        cell.width = col_widths[idx]

for row_idx, row_data in enumerate(table_data):
    for col_idx, cell_text in enumerate(row_data):
        cell = table.rows[row_idx].cells[col_idx]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.text = ""
        p = cell.paragraphs[0]
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                             line_spacing=1.0, first_line_indent=None,
                             space_before=2, space_after=2)
        r = p.add_run(cell_text)
        is_header = (row_idx == 0)
        set_run_font(r, size=FONT_SIZE_TABLE_DATA, bold=is_header)

add_source_note(doc, "Šaltinis: sudaryta autoriaus.")


# 2.2
add_h2(doc, "2.2. Pigu.lt ir Omniva Lietuva pavyzdžių analizė")

add_body(doc,
    "Pigu.lt yra viena didžiausių Lietuvos elektroninės prekybos platformų. "
    "Šios informacinės sistemos kūrimo reikalavimai apima produktų katalogą, "
    "krepšelio funkcionalumą, mokėjimų apdorojimą ir pristatymo logistiką. "
    "Sistema turi būti patogi tiek kompiuteryje, tiek mobiliajame telefone."
)

add_body(doc,
    "Omniva Lietuva, kaip pašto ir logistikos paslaugų teikėjas, valdo plačią "
    "paštomatų tinklo ir siuntų sekimo duomenų bazę. DB reikalavimai apima "
    "didelių duomenų kiekių apdorojimą realiu laiku, siuntų būsenų sekimą "
    "ir asmens duomenų saugojimą pagal BDAR."
)

add_body(doc,
    "Pigu.lt ir Omniva sąveika yra geras pavyzdys, kaip IS ir DB papildo "
    "viena kitą. Žemiau pateikta schema (žr. 1 pav.) iliustruoja šią sąveiką."
)

# 1 paveikslas
add_picture(doc, "/projects/sandbox/PU1_darbas/img1_seku_diagrama.png", width_cm=15.0)
add_picture_caption(doc, "1 pav", "Pigu.lt ir Omniva sąveikos schema")
add_source_note(doc, "Šaltinis: sudaryta autoriaus.")


# 2.3
add_h2(doc, "2.3. Rašytinė palyginamoji analizė")

add_body(doc,
    "Informacinės sistemos ir duomenų bazės yra glaudžiai susiję, tačiau jų "
    "kūrimo reikalavimai skiriasi. IS reikalavimai apima visą organizacinį "
    "kontekstą: vartotojų vaidmenis, verslo procesus ir vartotojo sąveiką. "
    "DB reikalavimai sutelkti ties duomenų struktūra, vientisumu ir saugumu."
)

add_body(doc,
    "Šios reikalavimų grupės papildo viena kitą. DB sukuriama remiantis IS "
    "funkciniais reikalavimais: nustačius, kokie verslo procesai bus "
    "automatizuoti, suprojektuojama, kokie duomenys turi būti saugomi. "
    "Pavyzdžiui, Pigu.lt sistema reikalauja saugoti produktus, klientus ir "
    "užsakymus – šie reikalavimai virsta DB schema su lentelėmis ir ryšiais."
)

add_body(doc,
    "Realus pavyzdys yra Pigu.lt ir Omniva sąveika: kai klientas užsako prekę "
    "Pigu.lt platformoje, IS perduoda siuntimo duomenis Omniva sistemai, o "
    "ši saugo juos savo DB ir grąžina sekimo numerį atgal į IS. Toks ryšys "
    "parodo, kad IS reikalavimai apibrėžia, kokia informacija perduodama, o "
    "DB reikalavimai – kaip ji saugoma ir apsaugota. Abi reikalavimų grupės "
    "yra vienodai svarbios kuriant patikimą skaitmeninę paslaugą."
)


# ============================================
# 6. 3 SKYRIUS — ŠABLONAI
# ============================================
add_h1(doc, "3. IS ir DB reikalavimų rinkimo šablonai")

add_body(doc,
    "Šiame skyriuje aprašomi dažniausiai naudojami šablonai informacinių "
    "sistemų ir duomenų bazių reikalavimams rinkti."
)

# 3.1
add_h2(doc, "3.1. IS reikalavimų specifikacijos šablonai")

add_body(doc,
    "Informacinių sistemų reikalavimams rinkti dažniausiai naudojami šie šablonai:"
)

add_list_item(doc, "Reikalavimų specifikacijos dokumentas (SRS) – aprašo funkcinius ir nefunkcinius sistemos reikalavimus.", style='List Bullet')
add_list_item(doc, "Panaudojimo atvejų šablonas – aprašo vartotojų sąveiką su sistema (UML diagrama).", style='List Bullet')
add_list_item(doc, "Vartotojo istorijos – Agile metodikoje naudojamas formatas: „Kaip vartotojas, noriu funkcionalumo, kad pasiekčiau tikslą\".", style='List Bullet')

# 3.2
add_h2(doc, "3.2. Duomenų bazių reikalavimų rinkimo šablonai")

add_body(doc,
    "Duomenų bazių projektavimui naudojami specializuoti šablonai:"
)

add_list_item(doc, "Esybių–ryšių diagrama (ERD) – grafinis modelis, vaizduojantis lenteles, laukus ir ryšius tarp jų.", style='List Bullet')
add_list_item(doc, "Duomenų žodynas – lentelinis aprašymas: lentelės, laukai, duomenų tipai, raktai.", style='List Bullet')
add_list_item(doc, "Loginė ir fizinė schema – etapinio DB projektavimo dokumentai.", style='List Bullet')

add_body(doc,
    "Antrajame paveiksle (žr. 2 pav.) pateikiamas supaprastintas e. prekybos "
    "ir logistikos duomenų bazės ER modelis."
)

# 2 paveikslas
add_picture(doc, "/projects/sandbox/PU1_darbas/img2_er_diagrama.png", width_cm=15.0)
add_picture_caption(doc, "2 pav", "Supaprastinta e. prekybos ir logistikos DB ER diagrama")
add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

add_body(doc,
    "Šablonai padeda projektų komandoms tiksliai dokumentuoti reikalavimus, "
    "mažina nesusipratimų riziką tarp užsakovų ir kūrėjų bei palengvina "
    "sistemos priežiūrą."
)


# ============================================
# 7. IŠVADOS
# ============================================
add_h1(doc, "Išvados")

add_list_item(doc,
    "IS ir DB yra glaudžiai susijusios, tačiau jų kūrimo reikalavimai "
    "skiriasi: IS aprėpia organizacijos veiklą, o DB sutelkta ties duomenų "
    "struktūra ir saugumu."
)
add_list_item(doc,
    "IS reikalavimai apima funkcinius (verslo procesus) ir nefunkcinius "
    "(saugumą, našumą) aspektus, o DB reikalavimai – schemą, normalizaciją "
    "ir prieigos kontrolę."
)
add_list_item(doc,
    "Pigu.lt ir Omniva Lietuva pavyzdys parodo, kad IS ir DB reikalavimai "
    "papildo vienas kitą: e. prekybos platforma per API perduoda duomenis "
    "logistikos paslaugų teikėjui, kuris saugo juos savo DB."
)
add_list_item(doc,
    "Tinkamai parinkti reikalavimų rinkimo šablonai (SRS, panaudojimo atvejai, "
    "ERD, duomenų žodynas) padeda sumažinti projektų klaidų riziką ir "
    "užtikrina sistemos kokybę."
)


# ============================================
# 8. LITERATŪROS SĄRAŠAS (APA)
# ============================================
add_h1(doc, "Literatūros sąrašas")

literature = [
    "Connolly, T. ir Begg, C. (2015). Database Systems: A Practical Approach to "
    "Design, Implementation, and Management (6th ed.). Boston: Pearson Education.",

    "Laudon, K. C. ir Laudon, J. P. (2020). Management Information Systems: "
    "Managing the Digital Firm (16th ed.). Harlow: Pearson Education.",

    "Sommerville, I. (2016). Software Engineering (10th ed.). Boston: Pearson.",

    "Pigu.lt. (n.d.). Apie mus. Prieiga per internetą: https://pigu.lt/lt/apie-mus",

    "Omniva. (n.d.). Apie Omniva Lietuvoje. Prieiga per internetą: "
    "https://www.omniva.lt/privatus/apie-omniva",
]

for src in literature:
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         line_spacing=1.5, first_line_indent=None,
                         space_before=0, space_after=6)
    p.paragraph_format.left_indent = Cm(1.25)
    p.paragraph_format.first_line_indent = Cm(-1.25)
    r = p.add_run(src)
    set_run_font(r, size=FONT_SIZE_BODY)


# ============================================
# IŠSAUGOJIMAS
# ============================================
output_path = "/projects/sandbox/PU1_darbas/PU1_Andrius_Vargonas.docx"
doc.save(output_path)
print(f"Sukurta: {output_path}")
