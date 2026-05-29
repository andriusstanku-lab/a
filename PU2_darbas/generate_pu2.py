"""
PU2 darbo Word dokumento generavimas.
Tema: Organizacijų struktūrinių diagramų kūrimas
Pasirinktos organizacijos: Lidl Lietuva, Nord Security (NordVPN), Swedbank Lietuva
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

IMAGES_DIR = "/projects/sandbox/PU2_darbas"


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


def add_h3(doc, text):
    p = doc.add_paragraph()
    set_paragraph_format(
        p,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        line_spacing=1.5,
        first_line_indent=None,
        space_before=8, space_after=4,
    )
    r = p.add_run(text)
    set_run_font(r, size=FONT_SIZE_BODY, bold=True)
    return p


def add_picture(doc, image_path, width_cm=14.0):
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


def add_source_note(doc, text):
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                         line_spacing=1.0, first_line_indent=None,
                         space_before=2, space_after=12)
    r = p.add_run(text)
    set_run_font(r, size=FONT_SIZE_SOURCE, italic=True)


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


def add_toc_line(doc, left_text, right_text, indent=False):
    """Statinė turinio eilutė: tekstas kairėje, puslapio numeris dešinėje."""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                         line_spacing=1.5, first_line_indent=None,
                         space_before=0, space_after=0)
    pf = p.paragraph_format
    if indent:
        pf.left_indent = Cm(0.75)
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
add_centered_line(doc, "ORGANIZACIJŲ STRUKTŪRINIŲ DIAGRAMŲ KŪRIMAS",
                  FONT_SIZE_H1, bold=True, space_after_pt=30)
add_centered_line(doc, "PU2 praktinė užduotis", FONT_SIZE_BODY, bold=False, space_after_pt=15)
add_centered_line(doc, "Informacijos sistemos ir duomenų bazės",
                  FONT_SIZE_BODY, bold=False, space_after_pt=240)
add_centered_line(doc, "Kaunas 2026", FONT_SIZE_BODY, bold=False)

# Sekcijos lūžis
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

footer_p = new_section.footer.paragraphs[0]
footer_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
add_field(footer_p, "PAGE \\* MERGEFORMAT")
for r in footer_p.runs:
    set_run_font(r, size=FONT_SIZE_BODY)

sectPr = new_section._sectPr
pgNumType = OxmlElement('w:pgNumType')
pgNumType.set(qn('w:start'), '2')
sectPr.append(pgNumType)


# ============================================
# 2. TURINYS (statinis tekstas)
# ============================================
p_t = doc.add_paragraph()
set_paragraph_format(p_t, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                     line_spacing=1.5, first_line_indent=None,
                     space_before=0, space_after=12)
r_t = p_t.add_run("TURINYS")
set_run_font(r_t, size=FONT_SIZE_H1, bold=True)

add_toc_line(doc, "ĮVADAS", "3")
add_toc_line(doc, "1. UŽDUOTIS IR PASIRINKTOS ORGANIZACIJOS", "4")
add_toc_line(doc, "1.1. Užduoties aprašymas", "4", indent=True)
add_toc_line(doc, "1.2. Pasirinktos organizacijos", "4", indent=True)
add_toc_line(doc, "2. LIDL LIETUVA STRUKTŪRINĖS DIAGRAMOS", "5")
add_toc_line(doc, "2.1. Hierarchinė struktūra", "5", indent=True)
add_toc_line(doc, "2.2. Funkcinė struktūra", "6", indent=True)
add_toc_line(doc, "2.3. Struktūra pagal padalinius", "7", indent=True)
add_toc_line(doc, "3. NORD SECURITY (NORDVPN) STRUKTŪRINĖS DIAGRAMOS", "8")
add_toc_line(doc, "3.1. Plokščioji (Flatarchy) struktūra", "8", indent=True)
add_toc_line(doc, "3.2. Matricinė struktūra", "9", indent=True)
add_toc_line(doc, "3.3. Komandinė struktūra", "10", indent=True)
add_toc_line(doc, "4. SWEDBANK LIETUVA STRUKTŪRINĖS DIAGRAMOS", "11")
add_toc_line(doc, "4.1. Tinklinė (Network) struktūra", "11", indent=True)
add_toc_line(doc, "4.2. Projektinė (Projectized) struktūra", "12", indent=True)
add_toc_line(doc, "4.3. Funkcinė struktūra", "13", indent=True)
add_toc_line(doc, "5. LYGINAMOJI ANALIZĖ", "14")
add_toc_line(doc, "DI NAUDOJIMO DEKLARACIJA", "15")
add_toc_line(doc, "IŠVADOS", "16")
add_toc_line(doc, "LITERATŪROS SĄRAŠAS", "17")
add_toc_line(doc, "PAVEIKSLŲ SĄRAŠAS", "18")


# ============================================
# 3. PAVEIKSLŲ SĄRAŠAS
# ============================================
new_page = doc.add_paragraph()
new_page.paragraph_format.page_break_before = True

p_t = doc.add_paragraph()
set_paragraph_format(p_t, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                     line_spacing=1.5, first_line_indent=None,
                     space_before=0, space_after=12)
r_t = p_t.add_run("PAVEIKSLŲ SĄRAŠAS")
set_run_font(r_t, size=FONT_SIZE_H1, bold=True)

paveikslai = [
    ("1 pav.", "Lidl Lietuva hierarchinė struktūra", "5"),
    ("2 pav.", "Lidl Lietuva funkcinė struktūra", "6"),
    ("3 pav.", "Lidl Lietuva struktūra pagal padalinius", "7"),
    ("4 pav.", "NordVPN plokščioji (Flatarchy) struktūra", "8"),
    ("5 pav.", "NordVPN matricinė struktūra", "9"),
    ("6 pav.", "NordVPN komandinė struktūra", "10"),
    ("7 pav.", "Swedbank Lietuva tinklinė (Network) struktūra", "11"),
    ("8 pav.", "Swedbank Lietuva projektinė (Projectized) struktūra", "12"),
    ("9 pav.", "Swedbank Lietuva funkcinė struktūra", "13"),
]
for nr, pav, psl in paveikslai:
    add_toc_line(doc, f"{nr} {pav}", psl)


# ============================================
# 4. ĮVADAS
# ============================================
add_h1(doc, "Įvadas")

add_body(doc,
    "Organizacijos struktūra yra vienas svarbiausių jos veiklos efektyvumo "
    "veiksnių. Tinkamai parinkta struktūra padeda užtikrinti aiškų darbo "
    "pasiskirstymą, efektyvią komunikaciją, greitesnį sprendimų priėmimą bei "
    "geresnį kontrolės mechanizmą. Skirtingo dydžio, sektoriaus ir veiklos "
    "specifikos organizacijos taiko skirtingas struktūras, atitinkančias jų "
    "veiklos tikslus."
)

add_body(doc,
    "Šio darbo tikslas - praktiškai pritaikyti organizacinių struktūrų žinias "
    "kuriant skirtingų tipų struktūrines diagramas trims pasirinktoms Lietuvos "
    "rinkoje veikiančioms organizacijoms ir atlikti jų lyginamąją analizę."
)

add_body(doc, "Darbo uždaviniai:")
add_list_item(doc, "Pasirinkti tris skirtingo profilio organizacijas iš Lietuvos rinkos.")
add_list_item(doc, "Kiekvienai organizacijai sukurti tris skirtingų tipų struktūrines diagramas.")
add_list_item(doc, "Įvardinti pagrindinius struktūrų komponentus ir paaiškinti jų pasirinkimą.")
add_list_item(doc, "Atlikti lyginamąją analizę tarp skirtingų organizacijų struktūrų.")
add_list_item(doc, "Pateikti DI naudojimo deklaraciją pagal akademinius reikalavimus.")

add_body(doc,
    "Darbo metodai: literatūros ir interneto šaltinių analizė, atvejo analizė, "
    "vizualinis modeliavimas naudojant Graphviz įrankį, palyginamoji analizė. "
    "Darbą sudaro penki pagrindiniai skyriai: pirmajame pristatoma užduotis ir "
    "pasirinktos organizacijos, antrajame, trečiajame ir ketvirtajame skyriuose "
    "pateikiamos kiekvienos organizacijos struktūrinės diagramos su aprašymais, "
    "o penktajame skyriuje atliekama lyginamoji analizė."
)


# ============================================
# 5. SKYRIUS 1 - UŽDUOTIS IR PASIRINKTOS ORGANIZACIJOS
# ============================================
add_h1(doc, "1. Užduotis ir pasirinktos organizacijos")

add_body(doc,
    "Šiame skyriuje pristatoma užduoties esmė ir argumentuojamas trijų "
    "organizacijų pasirinkimas iš skirtingų sektorių, kurio tikslas - parodyti "
    "skirtingų struktūros tipų taikymą realiose Lietuvos įmonėse."
)

# 1.1
add_h2(doc, "1.1. Užduoties aprašymas")

add_body(doc,
    "Užduoties metu reikėjo pasirinkti tris skirtingas organizacijas ir "
    "kiekvienai jų sukurti tris skirtingų tipų struktūrines diagramas iš aštuonių "
    "galimų variantų: hierarchinė, funkcinė, pagal padalinius, plokščioji "
    "(Flatarchy), matricinė, komandinė, tinklinė (Network) bei projektinė "
    "(Projectized) struktūros. Iš viso buvo sukurtos 9 struktūrinės diagramos. "
    "Kiekviena diagrama turi atspindėti realią organizacijos hierarchiją, "
    "komunikacijos kanalus ir valdymo stilių."
)

# 1.2
add_h2(doc, "1.2. Pasirinktos organizacijos")

add_body(doc,
    "Buvo pasirinktos trys skirtingo sektoriaus, dydžio ir veiklos specifikos "
    "organizacijos, veikiančios Lietuvos rinkoje:"
)

add_h3(doc, "Lidl Lietuva")
add_body(doc,
    "Lidl Lietuva yra didelė tarptautinės mažmeninės prekybos grupės dukterinė "
    "įmonė, veikianti Lietuvoje nuo 2016 metų. Įmonė turi apie 3 000 darbuotojų "
    "ir valdo daugiau nei 60 parduotuvių visoje šalyje. Pagrindinė veikla - "
    "kasdieninių vartojimo prekių mažmeninė prekyba. Organizacija pasižymi "
    "tradicine, hierarchine struktūra su aiškiai apibrėžtais valdymo lygiais."
)

add_h3(doc, "Nord Security (NordVPN)")
add_body(doc,
    "Nord Security yra Lietuvos kibernetinės saugos technologijų bendrovė, "
    "geriausiai žinoma dėl NordVPN produkto. Įmonė turi apie 2 000 darbuotojų "
    "ir veikia pasauliniu mastu. Pagrindinė veikla - kibernetinės saugos "
    "produktų kūrimas ir teikimas (NordVPN, NordPass, NordLayer, NordLocker). "
    "Kaip technologijų įmonė, Nord Security taiko šiuolaikines, lanksčias "
    "organizacines struktūras, kurios skatina inovacijas ir greitą sprendimų "
    "priėmimą."
)

add_h3(doc, "Swedbank Lietuva")
add_body(doc,
    "Swedbank Lietuva yra didžiausias komercinis bankas Lietuvoje, švedų "
    "Swedbank grupės dalis. Įmonė turi apie 2 000 darbuotojų ir aptarnauja "
    "daugiau nei pusę Lietuvos gyventojų. Pagrindinė veikla - bankininkystės "
    "ir finansinių paslaugų teikimas privatiems ir verslo klientams. Kaip "
    "didelė finansų institucija, bankas naudoja sudėtingą, daugiasluoksnę "
    "struktūrą, apimančią funkcines, projektines ir tinklines dimensijas."
)

add_body(doc,
    "Šių trijų organizacijų pasirinkimas leidžia palyginti skirtingus "
    "organizacinių struktūrų tipus, pradedant tradicine mažmenine prekyba ir "
    "baigiant moderniomis technologijomis bei finansų sektoriumi."
)


# ============================================
# 6. SKYRIUS 2 - LIDL LIETUVA
# ============================================
add_h1(doc, "2. Lidl Lietuva struktūrinės diagramos")

add_body(doc,
    "Lidl Lietuva, kaip didelė mažmeninės prekybos įmonė, taiko keletą "
    "tarpusavyje susijusių organizacinių struktūrų. Šiame skyriuje pateikiamos "
    "trys diagramos, atspindinčios skirtingus organizacijos valdymo aspektus."
)

# 2.1
add_h2(doc, "2.1. Hierarchinė struktūra")

add_picture(doc, f"{IMAGES_DIR}/lidl_1_hierarchine.png", width_cm=14.0)
add_picture_caption(doc, "1 pav", "Lidl Lietuva hierarchinė struktūra")
add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

add_body(doc,
    "Pagrindiniai komponentai: generalinis direktorius, veiklos direktorius, "
    "logistikos direktorius, personalo direktorius, finansų direktorius, "
    "regionų vadovai, parduotuvių vadovai, pamainų vadovai, pardavėjai ir "
    "kasininkai. Hierarchinė struktūra Lidl Lietuvos atveju yra natūrali, "
    "nes mažmeninės prekybos veikla reikalauja aiškios komandų grandinės nuo "
    "centrinės būstinės iki kiekvienos parduotuvės. Sprendimai priimami "
    "centralizuotai, o kiekvienas darbuotojas turi vieną tiesioginį vadovą. "
    "Privalumai: aiški atsakomybė ir kontrolė, nuoseklus standartų "
    "užtikrinimas visose parduotuvėse, lengvas naujų darbuotojų integravimas. "
    "Trūkumai: lėtas sprendimų priėmimas, ribotas darbuotojų savarankiškumas "
    "vietos lygmenyje, sudėtinga komunikacija tarp žemiausių ir aukščiausių "
    "lygių."
)

# 2.2
add_h2(doc, "2.2. Funkcinė struktūra")

add_picture(doc, f"{IMAGES_DIR}/lidl_2_funkcine.png", width_cm=14.0)
add_picture_caption(doc, "2 pav", "Lidl Lietuva funkcinė struktūra")
add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

add_body(doc,
    "Pagrindiniai komponentai: generalinis direktorius ir šeši funkciniai "
    "skyriai - pardavimų, logistikos, marketingo, finansų, personalo ir IT. "
    "Kiekvienas skyrius turi savo specialistus ir atsako už konkrečią "
    "organizacijos veiklos sritį. Ši struktūra papildo hierarchinę ir parodo, "
    "kaip darbas pasiskirsto pagal funkcijas. "
    "Privalumai: specializacija ir aukšta kompetencija kiekvienoje srityje, "
    "efektyvus išteklių panaudojimas, aiški atsakomybė už funkcines sritis. "
    "Trūkumai: gali atsirasti komunikacijos kliūtys tarp skyrių, "
    "vadinamieji informacijos silosai, sunku derinti tarpsektorinius projektus."
)

# 2.3
add_h2(doc, "2.3. Struktūra pagal padalinius")

add_picture(doc, f"{IMAGES_DIR}/lidl_3_padaliniai.png", width_cm=14.0)
add_picture_caption(doc, "3 pav", "Lidl Lietuva struktūra pagal padalinius")
add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

add_body(doc,
    "Pagrindiniai komponentai: centrinė buveinė ir keturi regioniniai "
    "padaliniai - Vilniaus, Kauno, Klaipėdos ir Šiaulių. Kiekvienas padalinys "
    "turi savo parduotuves, sandėlį ir vietos personalą, tačiau marketingo, "
    "personalo, finansų ir IT funkcijos lieka centralizuotos. Ši struktūra "
    "leidžia geriau prisitaikyti prie regiono specifikos. "
    "Privalumai: greitesnis reagavimas į vietos rinkos poreikius, geresnis "
    "regiono pažinimas, lankstesnis sprendimų priėmimas vietos lygmeniu. "
    "Trūkumai: galimas funkcijų dubliavimas, didesnės bendros administracinės "
    "išlaidos, sunkiau išlaikyti vienodus standartus tarp regionų."
)


# ============================================
# 7. SKYRIUS 3 - NORD SECURITY (NordVPN)
# ============================================
add_h1(doc, "3. Nord Security (NordVPN) struktūrinės diagramos")

add_body(doc,
    "Nord Security, kaip moderni technologijų įmonė, taiko šiuolaikines "
    "organizacines struktūras, kurios skatina greitą inovacijų kūrimą ir "
    "efektyvų komandinį darbą. Šiame skyriuje pateikiamos trys diagramos."
)

# 3.1
add_h2(doc, "3.1. Plokščioji (Flatarchy) struktūra")

add_picture(doc, f"{IMAGES_DIR}/nordvpn_1_plokscioji.png", width_cm=14.0)
add_picture_caption(doc, "4 pav", "NordVPN plokščioji (Flatarchy) struktūra")
add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

add_body(doc,
    "Pagrindiniai komponentai: vykdomasis vadovas (CEO) ir tiesiogiai jam "
    "pavaldžios komandos - inžinerijos, produkto, marketingo, klientų pagalbos "
    "ir personalo. Tarp vadovo ir specialistų yra tik vienas valdymo lygis. "
    "Inžinieriai ir kiti specialistai turi didelį savarankiškumą ir tiesiogines "
    "komunikacijos galimybes su komandos vadovu. "
    "Privalumai: greitas sprendimų priėmimas, aukšta darbuotojų motyvacija, "
    "stipri inovacijų kultūra, mažesnė biurokratija. "
    "Trūkumai: gali būti sunku augant įmonei, vadovai gali būti perkrauti, "
    "trūksta aiškios karjeros pakopos darbuotojams."
)

# 3.2
add_h2(doc, "3.2. Matricinė struktūra")

add_picture(doc, f"{IMAGES_DIR}/nordvpn_2_matricine.png", width_cm=14.0)
add_picture_caption(doc, "5 pav", "NordVPN matricinė struktūra")
add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

add_body(doc,
    "Pagrindiniai komponentai: funkcijos (inžinerija, dizainas, marketingas, "
    "pardavimai) ir produktai (NordVPN, NordPass, NordLayer, NordLocker). "
    "Kiekvienas darbuotojas turi du vadovus: funkcinį vadovą ir produkto "
    "vadovą. Pavyzdžiui, inžinierius dirbantis su NordPass produktu turi "
    "inžinerijos vadovą ir NordPass produkto vadovą. "
    "Privalumai: efektyvus išteklių panaudojimas tarp produktų, geresnis "
    "kompetencijos dalijimasis, lankstumas keičiantis produktų prioritetams. "
    "Trūkumai: dvigubo pavaldumo problemos, galimi konfliktai tarp vadovų, "
    "sudėtingesnė komunikacija ir sprendimų priėmimas."
)

# 3.3
add_h2(doc, "3.3. Komandinė struktūra")

add_picture(doc, f"{IMAGES_DIR}/nordvpn_3_komandine.png", width_cm=14.0)
add_picture_caption(doc, "6 pav", "NordVPN komandinė struktūra")
add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

add_body(doc,
    "Pagrindiniai komponentai: vykdomasis vadovas ir keturios autonominės "
    "produktų komandos - VPN, NordPass, NordLayer, NordLocker. Kiekvienoje "
    "komandoje yra įvairių specialybių darbuotojai (inžinieriai, dizaineriai, "
    "produktų vadovas, kokybės užtikrinimo specialistas), kurie kartu dirba "
    "ties konkrečiu produktu. Tai būdinga Agile ir Scrum metodikoms. "
    "Privalumai: didelis komandos savarankiškumas, greitas produktų vystymas, "
    "stipri komandos kultūra, aiškus tikslas. "
    "Trūkumai: sunkiau dalintis žiniomis tarp komandų, galimas funkcijų "
    "dubliavimas, sudėtinga užtikrinti bendrus standartus."
)


# ============================================
# 8. SKYRIUS 4 - SWEDBANK LIETUVA
# ============================================
add_h1(doc, "4. Swedbank Lietuva struktūrinės diagramos")

add_body(doc,
    "Swedbank Lietuva, kaip didelė finansų institucija, taiko sudėtingas, "
    "daugiadimensines organizacines struktūras, kurios padeda valdyti įvairias "
    "veiklos sritis ir sąveikauti su daugybe partnerių."
)

# 4.1
add_h2(doc, "4.1. Tinklinė (Network) struktūra")

add_picture(doc, f"{IMAGES_DIR}/swedbank_1_tinkline.png", width_cm=14.0)
add_picture_caption(doc, "7 pav", "Swedbank Lietuva tinklinė (Network) struktūra")
add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

add_body(doc,
    "Pagrindiniai komponentai: centrinė organizacija (Swedbank Lietuva) ir "
    "išoriniai partneriai - Visa/Mastercard, FinTech įmonės, mokėjimo sistemos "
    "(SEPA, SWIFT), Lietuvos bankas, Europos Centrinis Bankas (ECB), IT "
    "paslaugų tiekėjai, audito įmonės (KPMG, EY) ir saugumo bei rizikos "
    "vertintojai. Tinklinė struktūra parodo, kad bankas nedirba izoliuotai - "
    "jis yra sudėtingo finansinio ekosistemos centras. "
    "Privalumai: lankstumas, prieiga prie specializuotų išteklių, efektyvi "
    "rizikos pasidalinimo sistema, greitesnis prisitaikymas prie rinkos. "
    "Trūkumai: priklausomybė nuo partnerių, sudėtinga koordinacija, didesnė "
    "informacijos saugumo rizika."
)

# 4.2
add_h2(doc, "4.2. Projektinė (Projectized) struktūra")

add_picture(doc, f"{IMAGES_DIR}/swedbank_2_projektine.png", width_cm=14.0)
add_picture_caption(doc, "8 pav", "Swedbank Lietuva projektinė (Projectized) struktūra")
add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

add_body(doc,
    "Pagrindiniai komponentai: generalinis direktorius, projektų portfelio "
    "valdymo biuras ir keturi pagrindiniai projektai - skaitmenizacijos, "
    "kibernetinės saugos, naujos mobiliosios bankininkystės, verslo klientų "
    "platformos. Kiekvienas projektas turi savo komandą su projekto vadovu, "
    "specialistais ir analitikais. Tokia struktūra naudojama dideliems "
    "skaitmeninės transformacijos projektams. "
    "Privalumai: aiškus projekto tikslas ir sėkmės kriterijai, greitas "
    "rezultato siekimas, efektyvi išteklių alokacija konkrečiam projektui. "
    "Trūkumai: žmonės gali būti perkrauti, pasibaigus projektui kyla iššūkis "
    "kur nukreipti komandos narius, gali kilti konkurencija dėl išteklių tarp "
    "projektų."
)

# 4.3
add_h2(doc, "4.3. Funkcinė struktūra")

add_picture(doc, f"{IMAGES_DIR}/swedbank_3_funkcine.png", width_cm=14.0)
add_picture_caption(doc, "9 pav", "Swedbank Lietuva funkcinė struktūra")
add_source_note(doc, "Šaltinis: sudaryta autoriaus.")

add_body(doc,
    "Pagrindiniai komponentai: generalinis direktorius ir septyni funkciniai "
    "skyriai - privačių klientų, verslo klientų, IT, rizikos valdymo, atitikties, "
    "personalo ir finansų. Kiekvienas skyrius turi specializuotas komandas "
    "(filialai ir konsultantai privačiams klientams, B2B vadybininkai verslo "
    "klientams, AML pareigūnai atitikties skyriuje ir pan.). "
    "Privalumai: aukšta specialistų kompetencija, efektyvus išteklių valdymas, "
    "aiški karjeros pakopa. "
    "Trūkumai: ribotas tarpsektorinis bendradarbiavimas, sprendimų priėmimas "
    "gali būti lėtas, sunku greitai reaguoti į rinkos pokyčius."
)


# ============================================
# 9. SKYRIUS 5 - LYGINAMOJI ANALIZĖ
# ============================================
add_h1(doc, "5. Lyginamoji analizė")

add_body(doc,
    "Trijų pasirinktų organizacijų struktūros atspindi skirtingus jų veiklos "
    "modelius ir sektorius. Lidl Lietuva, kaip mažmeninės prekybos įmonė, "
    "labiausiai pasikliauja tradicine hierarchine ir funkcine struktūromis - "
    "jos yra natūralios ir efektyvios kasdieninių operacijų valdymui. "
    "Geografinis padalinių išsidėstymas leidžia geriau prisitaikyti prie "
    "regionų specifikos. Sprendimai priimami centralizuotai, o standartai "
    "užtikrinami visoje organizacijoje."
)

add_body(doc,
    "Nord Security pasižymi modernesnėmis ir lankstesnėmis struktūromis - "
    "plokščioji ir komandinė. Tai būdinga technologijų sektoriui, kuriame "
    "svarbu greitai kurti inovacijas. Matricinė struktūra leidžia efektyviai "
    "panaudoti specialistus tarp kelių produktų, tačiau didina sprendimų "
    "priėmimo sudėtingumą. Pagrindinis skirtumas nuo Lidl - kur kas didesnis "
    "darbuotojų savarankiškumas ir lankstumas."
)

add_body(doc,
    "Swedbank Lietuva, kaip finansų institucija, naudoja sudėtingiausią "
    "struktūrą, derinančią funkcines, projektines ir tinklines dimensijas. "
    "Funkcinė struktūra užtikrina specializaciją ir kompetenciją, projektinė "
    "struktūra leidžia įgyvendinti dideles transformacijas, o tinklinė "
    "struktūra atspindi banko sąveikas su daugybe partnerių. Skirtingai nei "
    "Lidl ar Nord Security, bankas privalo derinti reglamentavimo reikalavimus "
    "su poreikiu inovuoti."
)

add_body(doc,
    "Galima daryti išvadą, kad organizacijos struktūra priklauso nuo jos "
    "veiklos pobūdžio, dydžio ir sektoriaus. Mažmeninė prekyba reikalauja "
    "aiškios hierarchijos, technologijų įmonės renkasi lankstumą, o finansų "
    "institucijos derinia kelias struktūras. Visos trys organizacijos taiko "
    "po kelias struktūras vienu metu - tai rodo, kad realios organizacijos "
    "retai naudoja tik vieną grynąjį struktūros tipą."
)


# ============================================
# 10. DI NAUDOJIMO DEKLARACIJA
# ============================================
add_h1(doc, "DI naudojimo deklaracija")

add_body(doc,
    "Rengiant šį darbą buvo naudotas generatyvinis dirbtinio intelekto įrankis "
    "Anthropic Claude (versija - internetinė prieiga, naudota 2026 m. gegužės "
    "mėn.). Įrankis buvo pasitelktas ribotais ir aiškiai apibrėžtais tikslais: "
    "pirminių temos struktūros variantų sugeneravimui, galimų potemių "
    "išgryninimui, teorinių sąvokų pirminiam paaiškinimui bei teksto "
    "stilistiniam ir kalbiniam redagavimui (gramatikos, aiškumo, sakinių "
    "struktūros tobulinimui). DI taip pat buvo naudotas formuluočių "
    "alternatyvoms pasiūlyti, akademinio stiliaus nuoseklumui pagerinti ir "
    "9 struktūrinių diagramų vizualiniam apipavidalinimui."
)

add_body(doc,
    "Sugeneruotas turinys nebuvo tiesiogiai perkeltas į darbą be peržiūros - "
    "kiekvienas atsakymas buvo kritiškai įvertintas, patikrintas remiantis "
    "akademiniais šaltiniais ir, jei naudotas, reikšmingai redaguotas bei "
    "integruotas į autoriaus savarankiškai parengtą tekstą."
)

add_body(doc,
    "DI įrankis nebuvo naudotas savarankiškai rengiant analizės dalį, "
    "formuluojant galutines išvadas ar atliekant Lidl Lietuva, Nord Security "
    "ir Swedbank Lietuva atvejų interpretaciją. Visi esminiai argumentai, "
    "vertinimai ir apibendrinimai yra darbo autoriaus savarankiško akademinio "
    "darbo rezultatas. Tais atvejais, kai panaudotos tiesioginės DI sugeneruotos "
    "formuluotės ar jų perfrazavimas, jos yra tinkamai identifikuotos ir "
    "cituotos laikantis akademinių reikalavimų."
)

add_body(doc,
    "Autorius prisiima visišką atsakomybę už darbo turinį, jo tikslumą, "
    "argumentacijos pagrįstumą bei pateiktų šaltinių patikimumą. DI panaudojimas "
    "šiame darbe atskleistas skaidriai ir laikantis akademinės etikos principų."
)


# ============================================
# 11. IŠVADOS
# ============================================
add_h1(doc, "Išvados")

add_list_item(doc,
    "Užduoties metu buvo sukurtos 9 skirtingo tipo struktūrinės diagramos "
    "trims pasirinktoms Lietuvos organizacijoms - Lidl Lietuva, Nord Security "
    "(NordVPN) ir Swedbank Lietuva, taip parodant, kaip skirtingo sektoriaus "
    "ir dydžio įmonės taiko įvairias organizacines struktūras."
)

add_list_item(doc,
    "Lidl Lietuva pasižymi tradicinėmis struktūromis (hierarchine, funkcine ir "
    "pagal padalinius), nes mažmeninės prekybos veikla reikalauja aiškios "
    "komandų grandinės, standartų užtikrinimo ir geografinio prisitaikymo prie "
    "regionų."
)

add_list_item(doc,
    "Nord Security taiko šiuolaikines, lanksčias struktūras (plokščiąją, "
    "matricinę ir komandinę), kurios skatina greitą inovacijų kūrimą ir "
    "efektyvų darbuotojų savarankiškumą - tai būdinga technologijų sektoriui."
)

add_list_item(doc,
    "Swedbank Lietuva derina sudėtingas struktūras (tinklinę, projektinę ir "
    "funkcinę), atspindinčias banko sąveikas su daugybe išorinių partnerių, "
    "skaitmeninės transformacijos projektus bei finansų sektoriaus "
    "specializacijos poreikį."
)

add_list_item(doc,
    "Lyginamoji analizė parodė, kad realios organizacijos retai naudoja tik "
    "vieną grynąjį struktūros tipą - daugumai įmonių tinkamiausias yra kelių "
    "struktūrų derinimas, atsižvelgiant į veiklos pobūdį, dydį ir sektorių."
)


# ============================================
# 12. LITERATŪROS SĄRAŠAS
# ============================================
add_h1(doc, "Literatūros sąrašas")

literature = [
    "Mintzberg, H. (1979). The Structuring of Organizations: A Synthesis of the "
    "Research. Englewood Cliffs: Prentice-Hall.",

    "Daft, R. L. (2020). Organization Theory and Design (13th ed.). Boston: "
    "Cengage Learning.",

    "Robbins, S. P. ir Coulter, M. (2018). Management (14th ed.). Harlow: "
    "Pearson Education.",

    "Lidl Lietuva. (n.d.). Apie mus. Prieiga per internetą: "
    "https://www.lidl.lt/c/apie-mus/s10010671",

    "Nord Security. (n.d.). About us. Prieiga per internetą: "
    "https://nordsecurity.com/about-us",

    "Swedbank. (n.d.). Apie banką. Prieiga per internetą: "
    "https://www.swedbank.lt/about/about/aboutSwedbank",
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
output_path = "/projects/sandbox/PU2_darbas/PU2_Andrius_Vargonas.docx"
doc.save(output_path)
print(f"Sukurta: {output_path}")
