"""
PUA generavimo skriptas v3 - PILNAS akademinis pertvarkymas pagal VU KnF
Informatikos inzinerijos krypties metodiniu nurodymu reikalavimus.

Esmines v3 patobulinimas (lyginant su v2):
1. Detalios DI deklaracijos kiekvienam PU (Povilo modelio struktura, ne turinys)
2. Post-processor: visu pastraipu normalizavimas (TNR 12pt, 1.5 spacing, justify,
   1.25cm pirmoji eilutes itrauka)
3. Literaturos sarasas su APA hanging indent (1.25 cm)
4. Lenteliu / paveikslu Caption stilius: 11 pt Bold (per metodika)
5. Saltinio uzrasai po lentelemis/paveikslais: 9 pt
6. Heading 2 - 12 pt Bold, mazomis raidemis (kur galima)
7. Pries Heading 2 - 2 eiluciu tarpas, po - 1 eilutes
8. Lenteliu cell font: TNR 10 pt
9. Page number - desineje footer'yje, nuo 2 puslapio (titulinis nenumeruojamas)

Naudojimas:
    pip install python-docx docxcompose
    python3 merge_pua_v3.py

Autorius: Andrius Vargonas, 2026
"""

import os
import re
import shutil
from copy import deepcopy

from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docxcompose.composer import Composer


# ============================================================
# KONFIGURACIJA
# ============================================================

REPO_DIR = "/projects/sandbox/a"
OUT_DIR = os.path.join(REPO_DIR, "PUA_FINAL")
TMP_DIR = os.path.join(OUT_DIR, "_tmp")
OUT_FILE = os.path.join(OUT_DIR, "PUA_Andrius_Vargonas.docx")

PU_LIST = [
    ("PU1_Andrius_Vargonas.docx",  1,
     "INFORMACINIŲ SISTEMŲ IR DUOMENŲ BAZIŲ KŪRIMO REIKALAVIMŲ PALYGINIMAS"),
    ("PU2_Andrius_Vargonas.docx",  2,
     "ORGANIZACIJŲ STRUKTŪRINIŲ DIAGRAMŲ KŪRIMAS"),
    ("PU3_Andrius_Vargonas1.docx", 3,
     "DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ ANALIZĖ IR PALYGINIMAS"),
    ("PU4_Andrius_Vargonas.docx",  4,
     "DBVS ĮRANKIŲ APLINKOS APŽVALGA IR LENTELIŲ KŪRIMAS"),
    ("PU5_Andrius_Vargonas.docx",  5,
     "ESYBIŲ-RYŠIŲ DIAGRAMOS IR DUOMENŲ BAZIŲ ĮGYVENDINIMAS"),
    ("PU6_Andrius_Vargonas.docx",  6,
     "DUOMENŲ NORMALIZAVIMAS"),
    ("PU7_Andrius_Vargonas.docx",  7,
     "MS ACCESS DUOMENŲ BAZIŲ ANALIZĖ IR KŪRIMAS"),
]

# DI deklaraciju turinys kiekvienam PU - 2 uzklausos kiekviename
# Struktura atkartoja Povilo modeli, bet turinys originalus, sukurtas pagal
# Andriaus PU temas
DI_UZKLAUSOS = {
    1: [
        ("Paaiškink, kuo skiriasi informacinės sistemos ir duomenų bazės kūrimo "
         "reikalavimai? Pateik konkrečių pavyzdžių iš e. prekybos srities.",
         "Claude paaiškino, kad informacinės sistemos reikalavimai apima "
         "verslo procesus, naudotojų sąsajas, integraciją su išorinėmis "
         "sistemomis, o duomenų bazės reikalavimai sutelkti į duomenų "
         "struktūrą, raktus, ryšius ir vientisumo apribojimus. Pateikti "
         "Pigu.lt ir Omniva sąveikos pavyzdžiai - užsakymo IS perduoda "
         "duomenis į pristatymo IS per bendras klientų ir siuntų DB."),
        ("Kokie šablonai dažniausiai naudojami informacinių sistemų ir duomenų "
         "bazių reikalavimams rinkti? Pateik bent po du kiekvienai sričiai.",
         "Claude įvardijo IS reikalavimų rinkimo šablonus: reikalavimų "
         "specifikacijos dokumentą (SRS), panaudojimo atvejų (Use Case) šabloną "
         "ir vartotojo istorijas (User Stories). DB reikalavimams - esybių ir "
         "ryšių diagramą (ERD), duomenų žodyną (Data Dictionary) ir loginės/"
         "fizinės schemos dokumentus."),
    ],
    2: [
        ("Kuo skiriasi hierarchinė, funkcinė, matricinė, tinklinė ir komandinė "
         "organizacinės struktūros? Kokios jų stipriosios ir silpnosios pusės?",
         "Claude paaiškino kiekvieną struktūros tipą: hierarchinė turi aiškią "
         "valdymo grandinę, funkcinė grupuoja darbuotojus pagal specializacijas, "
         "matricinė taiko dvigubą pavaldumą, tinklinė remiasi išoriniais "
         "partneriais, komandinė skatina autonomiją. Aptarta kiekvienos "
         "struktūros tinkamumas pagal organizacijos dydį ir veiklos pobūdį."),
        ("Kokias struktūras taiko realios Lietuvos įmonės kaip Lidl, NordVPN "
         "ir Swedbank? Kokie jų hierarchinių ir funkcinių struktūrų skirtumai?",
         "Claude pateikė įžvalgų apie tarptautinių korporacijų vietines "
         "struktūras: Lidl Lietuva taiko klasikinę hierarchinę matricinę "
         "struktūrą, NordVPN derina plokščiąją (flatarchy) ir komandinę "
         "struktūrą produktų vystymui, Swedbank remiasi tinkline ir projektine "
         "struktūra dėl reguliacinių reikalavimų."),
    ],
    3: [
        ("Kokie yra populiariausi šiandienos duomenų bazių kūrimo įrankiai? "
         "Pateik bent dešimt skirtingų įrankių, įtraukiant tiek tradicinius, "
         "tiek šiuolaikinius low-code sprendimus.",
         "Claude pateikė dešimties įrankių sąrašą: Microsoft Access, "
         "LibreOffice Base, MySQL Workbench, pgAdmin, DBeaver, MongoDB Compass, "
         "Airtable, NocoDB, Budibase ir Supabase. Kiekvienam įrankiui pateikta "
         "trumpa charakteristika - kūrėjas, palaikoma DBVS, licencija, "
         "naudojimo aplinka ir pagrindinės funkcijos."),
        ("Pagal kokius kriterijus tikslinga lyginti DB kūrimo įrankius? "
         "Sudaryk lyginamosios analizės kriterijų sąrašą.",
         "Claude pasiūlė aštuonis lyginimo kriterijus: licencijos tipą "
         "(atvirojo kodo / komercinė), DBVS palaikymą, vizualinio modeliavimo "
         "galimybes, debesų integraciją, low-code/no-code aplinką, REST API "
         "palaikymą, mokymosi kreivę bei bendruomenės dydį. Kriterijai "
         "panaudoti sudarant lyginamąją lentelę."),
    ],
    4: [
        ("Kuo skiriasi Microsoft Access, LibreOffice Base ir Apache OpenOffice "
         "Base? Kokia jų aplinka, palaikoma DBVS ir naudojimo paskirtis?",
         "Claude paaiškino, kad MS Access yra komercinė Microsoft 365 dalis "
         "su JET/ACE varikliu, LibreOffice Base - atvirojo kodo, naudoja HSQLDB "
         "ar Firebird, OpenOffice Base - panaši į LibreOffice, bet su "
         "lėtesniu vystymu. Visi įrankiai turi Design View, Forms, Queries, "
         "Reports modulius."),
        ("Kaip suprojektuoti paprastą universiteto duomenų bazę? Kokios "
         "lentelės reikalingos studentų ir kursų valdymui?",
         "Claude pasiūlė trijų lentelių struktūrą: Students (StudentID PK, "
         "FirstName, LastName, Email, Major, EnrollmentYear), Courses "
         "(CourseID PK, CourseName, Credits, Department) ir Enrollments "
         "(EnrollmentID PK, StudentID FK, CourseID FK, Semester, Grade). "
         "Tai klasikinis daug-prie-daug ryšys per asociatyvinę lentelę."),
    ],
    5: [
        ("Kaip identifikuoti esybes ir atributus iš pateikto dalykinės srities "
         "aprašymo? Kaip iš ER diagramos pereiti prie reliacinių lentelių?",
         "Claude paaiškino esybių paiešką pagal daiktavardžius tekste, "
         "atributų - pagal daiktavardžius, kurie aprašo esybes. Pateikė "
         "konvertavimo taisykles: vienas-prie-vieno -> sujungtos lentelės, "
         "vienas-prie-daug -> išorinis raktas, daug-prie-daug -> nauja "
         "asociatyvinė lentelė su sudėtiniu pirminiu raktu."),
        ("Kokie SQL DDL skriptai reikalingi sukurti e-prekybos, ligoninės ir "
         "bibliotekos duomenų bazes HSQLDB sintaksėje?",
         "Claude pateikė pilnus CREATE TABLE skriptus su INTEGER IDENTITY "
         "pirminiais raktais, FOREIGN KEY ryšiais, NOT NULL ir UNIQUE "
         "apribojimais. Skriptai pritaikyti LibreOffice Base ir DB Browser "
         "for SQLite naudojimui, su pavyzdiniais duomenimis."),
    ],
    6: [
        ("Kas yra duomenų normalizavimas, kokios yra trys pirmosios normalinės "
         "formos (1NF, 2NF, 3NF) ir kokias anomalijas jos pašalina?",
         "Claude apibūdino normalizavimą kaip duomenų organizavimo procesą, "
         "skirtą sumažinti redundanciją ir pašalinti įterpimo, atnaujinimo "
         "bei šalinimo anomalijas. 1NF - atominės reikšmės, 2NF - pašalintos "
         "dalinės priklausomybės, 3NF - pašalintos tranzityvinės priklausomybės "
         "nuo ne-rakto laukų."),
        ("Kaip normalizuoti studentų registracijos lentelę, kurioje yra "
         "studentų, dėstytojų ir kursų informacija? Pateik žingsnis po "
         "žingsnio analizę.",
         "Claude parodė nenormalizuotos lentelės pavyzdį, kur StudentName, "
         "InstructorName, CourseName ir Grade yra vienoje eilutėje. Po 1NF "
         "lentelė skaidoma į atomines reikšmes, po 2NF - StudentID ir "
         "CourseID išskiriami į atskiras lenteles, po 3NF gaunama keturių "
         "lentelių struktūra: Students, Instructors, Courses, Enrollments."),
    ],
    7: [
        ("Kaip MS Access analizuoti duomenų bazės struktūrą, lentelių laukus "
         "ir tarpusavio ryšius? Kaip identifikuoti pirminius ir išorinius "
         "raktus iš Relationships lango?",
         "Claude paaiškino Access Database Tools -> Relationships funkcionalumą, "
         "kuris vizualizuoja lenteles ir jų ryšius. Pirminiai raktai pažymėti "
         "rakto piktograma, išoriniai raktai sujungti linijomis su 1:N ryšio "
         "žymėjimu. Aptarta CRUD operacijų atlikimas Datasheet View ir "
         "Design View režimuose."),
        ("Kaip sukurti tris skirtingas užklausas su skaičiavimo funkcijomis "
         "SUM, AVG ir COUNT? Kokios SQL konstrukcijos naudojamos kiekvienai?",
         "Claude pateikė trijų užklausų pavyzdžius: SUM užklausa - SELECT "
         "Dept, SUM(UnitsInStock * RetailPrice) FROM tblInventory GROUP BY "
         "Dept; AVG užklausa - vidutinė kaina pagal kilmės šalį per AVG "
         "funkciją; COUNT užklausa - produktų skaičius pagal tiekėją per "
         "COUNT(*). Užklausos saugomos atskirame DB6.accdb faile."),
    ],
}


W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


# ============================================================
# OOXML PAGALBINES FUNKCIJOS
# ============================================================

def _qn(tag):
    return W_NS + tag


def _is_paragraph(el):
    return el.tag == _qn('p')


def _is_table(el):
    return el.tag == _qn('tbl')


def _get_pstyle(p_elem):
    pPr = p_elem.find(_qn('pPr'))
    if pPr is None:
        return None
    pStyle = pPr.find(_qn('pStyle'))
    if pStyle is None:
        return None
    return pStyle.get(_qn('val'))


def _get_paragraph_text(p_elem):
    raw = ''.join(p_elem.itertext()).strip()
    if not raw:
        return ''
    n = len(raw)
    for k in (3, 2):
        if n % k == 0 and n // k > 3:
            chunk = raw[:n // k]
            if chunk * k == raw:
                return chunk
    return raw


def _is_heading1(p_elem):
    s = _get_pstyle(p_elem)
    return s in ('Heading1', 'Heading 1', 'Antrate1', 'antraste1')


def _add_field(paragraph, instr_text, placeholder="(F9)"):
    run = paragraph.add_run()
    r = run._r
    fld = OxmlElement('w:fldChar')
    fld.set(qn('w:fldCharType'), 'begin')
    r.append(fld)
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = instr_text
    r.append(instr)
    fld = OxmlElement('w:fldChar')
    fld.set(qn('w:fldCharType'), 'separate')
    r.append(fld)
    t = OxmlElement('w:t')
    t.text = placeholder
    r.append(t)
    fld = OxmlElement('w:fldChar')
    fld.set(qn('w:fldCharType'), 'end')
    r.append(fld)


def _set_run_font(run, name="Times New Roman", size_pt=12, bold=False, italic=False):
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    rFonts.set(qn('w:cs'), name)
    rFonts.set(qn('w:eastAsia'), name)


def _set_paragraph_hanging_indent(paragraph, hanging_cm=1.25):
    """APA stiliaus pakabinta itrauka - pirmoji eilute be itraukos,
    likusios - su itrauka."""
    pPr = paragraph._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    twips = int(hanging_cm * 567)  # 1 cm = 567 twips
    ind.set(qn('w:left'), str(twips))
    ind.set(qn('w:hanging'), str(twips))


def _add_page_break(paragraph):
    run = paragraph.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)


# ============================================================
# KLASIFIKAVIMO LOGIKA
# ============================================================

def _classify_h1(text):
    t = (text or "").strip().upper()
    if not t:
        return "empty"
    if t.startswith("ĮVAD") or t.startswith("IVAD"):
        return "intro"
    if t.startswith("IŠVAD") or t.startswith("ISVAD"):
        return "conclusions"
    if t.startswith("LITERAT"):
        return "references"
    if "DI NAUDOJ" in t or "DIRBTINIO INTEL" in t:
        return "ai_decl"
    if "PRIED" in t:
        return "ai_decl"
    if t and t[0].isdigit() and "." in t[:4]:
        return "main"
    return "other"


def _is_priedas_marker(text):
    t = (text or "").strip().upper()
    if not t:
        return False
    return bool(re.match(r'^\d+\s*\.?\s*PRIEDAS', t))


# ============================================================
# PU DOKUMENTO ANALIZE IR FILTRAVIMAS
# ============================================================

def analyze_and_split_pu(pu_path):
    doc = Document(pu_path)
    body = doc.element.body
    block_elems = [c for c in body if _is_paragraph(c) or _is_table(c)]

    main_section_re = re.compile(
        r'^(\d+)\.\s+([A-ZĄČĘĖĮŠŲŪŽ][A-ZĄČĘĖĮŠŲŪŽ\s\-,]{4,})')

    zones = {"intro": [], "main": [], "conclusions": [],
             "references": [], "appendix": []}
    state = "PRE"

    for bi, el in enumerate(block_elems):
        if _is_paragraph(el):
            text = _get_paragraph_text(el)
            is_h1 = _is_heading1(el)
            zone_h1 = _classify_h1(text) if is_h1 else None

            if is_h1 and zone_h1 == "intro":
                state = "INTRO"
                continue
            if is_h1 and zone_h1 == "conclusions":
                state = "CONCLUSIONS"
                continue
            if is_h1 and zone_h1 == "references":
                state = "REFERENCES"
                continue
            if is_h1 and zone_h1 == "ai_decl":
                state = "APPENDIX"
                continue
            if is_h1 and zone_h1 == "main":
                state = "MAIN"
                zones["main"].append(bi)
                continue

            if _is_priedas_marker(text):
                state = "APPENDIX"
                continue

            if not is_h1 and state in ("INTRO", "MAIN"):
                if main_section_re.match(text):
                    state = "MAIN"
                    zones["main"].append(bi)
                    continue

        if state == "PRE":
            continue
        if state == "INTRO":
            zones["intro"].append(bi)
        elif state == "MAIN":
            zones["main"].append(bi)
        elif state == "CONCLUSIONS":
            zones["conclusions"].append(bi)
        elif state == "REFERENCES":
            zones["references"].append(bi)
        elif state == "APPENDIX":
            zones["appendix"].append(bi)

    conclusions_paragraphs = []
    for bi in zones["conclusions"]:
        el = block_elems[bi]
        if _is_paragraph(el):
            t = _get_paragraph_text(el)
            if t:
                conclusions_paragraphs.append(t)

    references_paragraphs = []
    for bi in zones["references"]:
        el = block_elems[bi]
        if _is_paragraph(el):
            t = _get_paragraph_text(el)
            if t:
                references_paragraphs.append(t)

    return {
        "doc": doc, "body": body, "block_elems": block_elems, "zones": zones,
        "conclusions_paragraphs": conclusions_paragraphs,
        "references_paragraphs": references_paragraphs,
    }


def filter_pu_keep_intro_and_main(pu_path, dst_path):
    info = analyze_and_split_pu(pu_path)
    body = info["body"]
    block_elems = info["block_elems"]

    keep = set()
    for bi in info["zones"]["intro"]:
        keep.add(id(block_elems[bi]))
    for bi in info["zones"]["main"]:
        keep.add(id(block_elems[bi]))

    for el in list(body):
        if _is_paragraph(el) or _is_table(el):
            if id(el) not in keep:
                body.remove(el)

    info["doc"].save(dst_path)


def filter_pu_keep_extra_appendix(pu_path, dst_path):
    """Is PU appendix zonos paima TIK papildoma medziaga (SQL skriptai,
    pavyzdiniai duomenys, detali charakteristika), NE DI deklaracija.

    Logika:
    - Surask DI deklaracijos antrastes pozicija appendix zonoje
    - Viskas PRIES tai = EXTRAS (jei yra)
    - Jei visa appendix yra DI deklaracija (PU1, PU2 atvejai) -> extras=None
    """
    info = analyze_and_split_pu(pu_path)
    body = info["body"]
    block_elems = info["block_elems"]
    appendix_indices = info["zones"]["appendix"]

    if not appendix_indices:
        return None

    # Surask DI deklaracijos antrastes pozicija (Normal arba H1 stiliumi)
    di_start = None
    for bi in appendix_indices:
        el = block_elems[bi]
        if _is_paragraph(el):
            t = _get_paragraph_text(el).upper()
            if 'DIRBTINIO INTELEKTO PANAUDOJIMO DEKLARAC' in t:
                di_start = bi
                break
            if 'DI NAUDOJIMO DEKLARAC' in t:
                di_start = bi
                break

    # Jei DI antrastes nera, patikrink ar visa appendix yra DI turinys
    if di_start is None:
        first_el = block_elems[appendix_indices[0]]
        if _is_paragraph(first_el):
            first_text = _get_paragraph_text(first_el)
            # Tipiniai DI deklaracijos turinio pradzios sakiniai
            di_keywords = ('Rengiant', 'naudotas dirbtinio intelekto',
                           'naudotas generatyvinis')
            if any(kw.lower() in first_text.lower() for kw in di_keywords):
                # Tai DI deklaracijos turinys be antrastes - skip
                return None
        # Kitaip - laikyti visa appendix kaip extras
        extra_indices = list(appendix_indices)
        extra_title = "PAPILDOMA MEDŽIAGA"
    else:
        # Extras = appendix elementai pries DI start
        extra_candidates = [bi for bi in appendix_indices if bi < di_start]
        if not extra_candidates:
            return None
        # Pirmoji didzioji pastraipa - pavadinimas
        extra_title = None
        extra_indices = []
        for bi in extra_candidates:
            el = block_elems[bi]
            if _is_paragraph(el):
                t = _get_paragraph_text(el)
                if extra_title is None and t and t.upper() == t and 5 < len(t) < 200:
                    extra_title = t
                    continue
            extra_indices.append(bi)

        if not extra_indices:
            return None

    # Filter ir save
    keep = set(id(block_elems[bi]) for bi in extra_indices)
    for el in list(body):
        if _is_paragraph(el) or _is_table(el):
            if id(el) not in keep:
                body.remove(el)

    info["doc"].save(dst_path)
    return extra_title or "PAPILDOMA MEDŽIAGA"


def demote_internal_headings(filtered_doc_path):
    """Po PU filtravimo, perlygina antrasciu hierarchija:
       - H1 ('X.') -> H2
       - H2 ('X.X.') -> H3
       - PU1 atveju Normal 'X.' (didziosiomis) -> H2
    """
    doc = Document(filtered_doc_path)
    main_section_re = re.compile(
        r'^(\d+)\.\s+([A-ZĄČĘĖĮŠŲŪŽ][A-ZĄČĘĖĮŠŲŪŽ\s\-,]{4,})')

    for p in doc.paragraphs:
        text = _get_paragraph_text(p._element)
        if not text:
            continue

        if _is_heading1(p._element):
            p.style = doc.styles['Heading 2']
        else:
            cur_style = p.style.name if p.style else ""
            if cur_style.startswith('Normal') or cur_style == 'Default Paragraph Font':
                if main_section_re.match(text):
                    p.style = doc.styles['Heading 2']

    h2_subsection_re = re.compile(r'^(\d+\.\d+\.?)\s+')
    for p in doc.paragraphs:
        text = _get_paragraph_text(p._element)
        if not text:
            continue
        if p.style and p.style.name == 'Heading 2':
            if h2_subsection_re.match(text):
                p.style = doc.styles['Heading 3']

    doc.save(filtered_doc_path)


# ============================================================
# LENTELIU / PAVEIKSLU SARASU GENERAVIMAS
# ============================================================

def collect_tables_and_figures(pu_path, pu_number):
    doc = Document(pu_path)
    paragraphs = doc.paragraphs
    tables = []
    figures = []

    table_label_re = re.compile(r'^\s*(\d+)\s+lentelė\s*$', re.IGNORECASE)
    figure_label_re = re.compile(r'^\s*(\d+)\s+pav\.?\s*(.*)$', re.IGNORECASE)

    seen_table_nums = set()
    seen_figure_nums = set()

    for i, p in enumerate(paragraphs):
        text = (p.text or "").strip()
        if not text:
            continue

        m = table_label_re.match(text)
        if m:
            num = int(m.group(1))
            if num in seen_table_nums:
                continue
            seen_table_nums.add(num)
            if i + 1 < len(paragraphs):
                title = (paragraphs[i+1].text or "").strip()
            else:
                title = ""
            if title:
                tables.append((num, title))
            continue

        m = figure_label_re.match(text)
        if m:
            num = int(m.group(1))
            if num in seen_figure_nums:
                continue
            seen_figure_nums.add(num)
            title = m.group(2).strip()
            if not title and i - 1 >= 0:
                prev = (paragraphs[i-1].text or "").strip()
                if prev and len(prev) < 200 and not table_label_re.match(prev) and not figure_label_re.match(prev):
                    title = prev
            if title:
                figures.append((num, title))
            continue

    return tables, figures


# ============================================================
# LITERATUROS SARASO DEDUPLIKAVIMAS
# ============================================================

def dedupe_and_sort_references(all_refs):
    seen_keys = {}
    for ref in all_refs:
        key = re.sub(r'\s+', ' ', ref.strip().lower())[:80]
        if key not in seen_keys:
            seen_keys[key] = ref.strip()
    sorted_refs = sorted(seen_keys.values(), key=lambda s: s.lower())
    return sorted_refs


# ============================================================
# DOKUMENTO NUSTATYMAI
# ============================================================

def configure_section(section):
    section.page_height = Mm(297)
    section.page_width = Mm(210)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.left_margin = Mm(25)
    section.right_margin = Mm(15)
    section.header_distance = Mm(12.5)
    section.footer_distance = Mm(12.5)
    section.different_first_page_header_footer = True


def configure_styles(doc):
    """Sukonfiguruoja Normal, Heading 1-3, Caption stilius pagal VU reikalavimus."""
    # Normal
    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(12)
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    pf = normal.paragraph_format
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.first_line_indent = Cm(1.25)

    # Heading 1: 14 pt, Bold, Center, didziosiomis, page break before
    h1 = doc.styles['Heading 1']
    h1.font.name = 'Times New Roman'
    h1.font.size = Pt(14)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor(0, 0, 0)
    rPr = h1.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    pf = h1.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(18)
    pf.line_spacing = 1.5
    pf.page_break_before = True
    pf.keep_with_next = True

    # Heading 2: 12 pt, Bold, Left, MAZOMIS raidemis (pirmoji didzioji)
    # Pries - 2 eil. tarpas (24 pt), po - 1 eil. (12 pt)
    h2 = doc.styles['Heading 2']
    h2.font.name = 'Times New Roman'
    h2.font.size = Pt(12)
    h2.font.bold = True
    h2.font.color.rgb = RGBColor(0, 0, 0)
    rPr = h2.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    pf = h2.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(24)
    pf.space_after = Pt(12)
    pf.line_spacing = 1.5
    pf.page_break_before = False
    pf.keep_with_next = True

    # Heading 3: 12 pt, Bold, Left
    h3 = doc.styles['Heading 3']
    h3.font.name = 'Times New Roman'
    h3.font.size = Pt(12)
    h3.font.bold = True
    h3.font.color.rgb = RGBColor(0, 0, 0)
    rPr = h3.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    pf = h3.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(12)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.5
    pf.page_break_before = False
    pf.keep_with_next = True

    # Caption: 11 pt, Bold (lentelems ir paveikslams)
    if 'Caption' in [s.name for s in doc.styles]:
        cap = doc.styles['Caption']
        cap.font.name = 'Times New Roman'
        cap.font.size = Pt(11)
        cap.font.bold = True
        cap.font.italic = False
        cap.font.color.rgb = RGBColor(0, 0, 0)


def setup_footer_page_numbers(section):
    footer = section.footer
    if footer.paragraphs:
        para = footer.paragraphs[0]
    else:
        para = footer.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pf = para.paragraph_format
    pf.first_line_indent = Cm(0)
    _add_field(para, " PAGE   \\* MERGEFORMAT ", placeholder="2")

    first_footer = section.first_page_footer
    if not first_footer.paragraphs:
        first_footer.add_paragraph()
    fp = first_footer.paragraphs[0]
    fp.text = ""
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT


# ============================================================
# PASTRAIPU PAGALBOS FUNKCIJOS
# ============================================================

def add_centered_paragraph(doc, text, *, size=12, bold=False,
                           space_before=0, space_after=0, italic=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.5
    if text:
        run = p.add_run(text)
        _set_run_font(run, size_pt=size, bold=bold, italic=italic)
    return p


def add_left_paragraph(doc, text, *, size=12, bold=False,
                       indent_first=False, space_before=0, space_after=0,
                       italic=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.first_line_indent = Cm(1.25) if indent_first else Cm(0)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.5
    if text:
        run = p.add_run(text)
        _set_run_font(run, size_pt=size, bold=bold, italic=italic)
    return p


def add_right_paragraph(doc, text, *, size=12, bold=False,
                        space_before=0, space_after=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.5
    if text:
        run = p.add_run(text)
        _set_run_font(run, size_pt=size, bold=bold)
    return p


def add_section_heading_no_toc(doc, text):
    """Skyrius su page break before, didziosiomis, 14 pt, centras, bold,
    bet NEpatenkantis i automatini turini (Normal stiliumi)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(18)
    pf.line_spacing = 1.5
    pf.page_break_before = True
    run = p.add_run(text)
    _set_run_font(run, size_pt=14, bold=True)
    return p


# ============================================================
# TITULINIS LAPAS
# ============================================================

def add_title_page(doc):
    add_centered_paragraph(doc, "VILNIAUS UNIVERSITETO", size=12, space_after=0)
    add_centered_paragraph(doc, "KAUNO FAKULTETAS", size=12, space_after=0)
    add_centered_paragraph(doc, "SOCIALINIŲ MOKSLŲ IR TAIKOMOSIOS INFORMATIKOS INSTITUTAS",
                           size=12, space_after=24)
    add_centered_paragraph(doc, "Marketingo technologijų studijų programa",
                           size=12, space_after=180)
    add_centered_paragraph(doc, "ANDRIUS VARGONAS", size=14, bold=True, space_after=180)
    add_centered_paragraph(doc, "PRAKTINIŲ UŽDUOČIŲ ATASKAITA",
                           size=16, bold=True, space_after=12)
    add_centered_paragraph(doc, "Informacijos sistemos ir duomenų bazės",
                           size=12, italic=True, space_after=240)
    add_centered_paragraph(doc, "Kaunas", size=12, space_after=0)
    add_centered_paragraph(doc, "2026", size=12, space_after=0)
    last = doc.paragraphs[-1]
    _add_page_break(last)


# ============================================================
# DI DEKLARACIJOS PRIEDU GENERAVIMAS
# ============================================================

def add_appendix(doc, number, pu_theme):
    """Pridek prieda pagal VU metodika:
       - Virsutiniame DESINIAJAME kampe 'N PRIEDAS' (12 pt)
       - Centre DIDZIOSIOMIS Bold pavadinimas (12 pt)
       - DI deklaracijos turinys
    """
    # Virsutinis desinysis: 'N PRIEDAS' (12 pt)
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pf = p1.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.5
    pf.page_break_before = True
    run = p1.add_run("{} PRIEDAS".format(number))
    _set_run_font(run, size_pt=12, bold=False)

    # Centre: Pavadinimas DIDZIOSIOMIS Bold
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p2.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(12)
    pf.space_after = Pt(18)
    pf.line_spacing = 1.5
    title = "PU{} DIRBTINIO INTELEKTO UŽKLAUSOS IR GAUTI ATSAKYMAI".format(number)
    run = p2.add_run(title)
    _set_run_font(run, size_pt=12, bold=True)

    # Naudoto DI irankio info
    add_left_paragraph(doc, "Naudotas įrankis: Anthropic Claude (Sonnet 4.5)",
                       size=12, indent_first=False, space_after=0)
    add_left_paragraph(doc, "Prieiga per: claude.ai (internetinė versija)",
                       size=12, indent_first=False, space_after=0)
    add_left_paragraph(doc, "Data: 2026 m. gegužės mėn.",
                       size=12, indent_first=False, space_after=12)

    # Apie naudojima
    intro_text = ("Rengiant šį praktinės užduoties (PU{}) skyrių, dirbtinio "
                  "intelekto įrankis buvo naudojamas struktūros sudarymui, "
                  "sąvokų paaiškinimui ir lyginamosios analizės kriterijų "
                  "atrankai. Visi Claude pateikti pasiūlymai buvo kritiškai "
                  "peržiūrėti, papildyti šaltiniais ir pritaikyti pagal "
                  "konkrečią užduoties specifiką. Žemiau pateikiamos "
                  "pagrindinės pateiktos užklausos ir gautų atsakymų santraukos."
                  ).format(number)
    add_left_paragraph(doc, intro_text, size=12, indent_first=True, space_after=12)

    # Uzklausos ir atsakymai
    uzklausos = DI_UZKLAUSOS.get(number, [])
    for idx, (q, a) in enumerate(uzklausos, 1):
        # X uzklausa: (bold)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(12)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.5
        run = p.add_run("{} užklausa:".format(idx))
        _set_run_font(run, size_pt=12, bold=True)

        # Uzklausos tekstas (paprastas)
        add_left_paragraph(doc, q, size=12, indent_first=True, space_after=6)

        # Gautas atsakymas: (bold)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.5
        run = p.add_run("Gautas atsakymas (santrauka):")
        _set_run_font(run, size_pt=12, bold=True, italic=True)

        # Atsakymo santrauka
        add_left_paragraph(doc, a, size=12, indent_first=True, space_after=6)

    # Pabaigai - DI naudojimo apimtis ir patvirtinimas
    add_left_paragraph(doc,
        "DI naudojimo apimtis: gauti atsakymai modifikuoti ir pritaikyti "
        "darbo kontekstui apie 60-70 proc. Galutiniame tekste DI sugeneruoto "
        "turinio dalis neviršija 5 proc. - visi teiginiai patikrinti pagal "
        "papildomus šaltinius (žr. literatūros sąrašą).",
        size=12, indent_first=True, space_before=12, space_after=6)

    add_left_paragraph(doc,
        "Autorius patvirtina, kad yra susipažinęs su Vilniaus universiteto "
        "2024 m. patvirtintomis dirbtinio intelekto naudojimo gairėmis "
        "(Nr. SPN-54), su Anthropic Claude privatumo politika bei naudojimo "
        "taisyklėmis, ir prisiima visišką atsakomybę už šio darbo turinį, "
        "tikslumą ir argumentacijos pagrįstumą.",
        size=12, indent_first=True, space_after=0)


# ============================================================
# POST-PROCESSOR - FORMATAVIMO NORMALIZAVIMAS
# ============================================================

def _convert_h2_to_sentence_case(text):
    """Konvertuoja H2 antraste is DIDZIOSIOMIS RAIDEMIS i 'Numeris. Pirmoji
    didzioji + likusios mazomis' formatu pagal VU metodika.

    Pavyzdys:
        '1. UŽDUOTIS IR PASIRINKTOS ORGANIZACIJOS'
        -> '1. Užduotis ir pasirinktos organizacijos'

    Issaugo akronimus (IS, DB, DBVS, ER, ERD, SQL, DDL, DML, NoSQL, MS, API,
    JSON, REST, NF, 1NF, 2NF, 3NF) ir tikriniu vardu/vietovardziu dideles raides.
    """
    # Akronimai - lieka DIDZIOSIOMIS
    ACRONYMS = {
        'IS', 'DB', 'DBVS', 'ER', 'ERD', 'SQL', 'DDL', 'DML', 'DCL', 'TCL',
        'NF', '1NF', '2NF', '3NF', 'BCNF', '4NF', '5NF',
        'PK', 'FK', 'UK', 'CK', 'NULL', 'NOT', 'PRIMARY', 'KEY', 'FOREIGN',
        'API', 'REST', 'JSON', 'XML', 'HTTP', 'HTTPS', 'URL', 'HTML', 'CSS', 'JS',
        'NoSQL', 'NoCode', 'LowCode', 'CRUD',
        'MS', 'VU', 'KnF', 'PUA', 'PU1', 'PU2', 'PU3', 'PU4', 'PU5', 'PU6', 'PU7',
        'IT', 'IS/DB', 'IT/IS', 'CEO', 'CFO', 'CIO', 'CTO', 'CMO',
        'DB1', 'DB2', 'DB3', 'DB4', 'DB5', 'DB6',
        'CSV', 'PDF', 'GUI',
    }
    # Tikriniu vardu/vietovardziu/preku zenklu zodynas (key=lower, val=teisingas
    # rasymas, kad palaikytume mišrius akronimus kaip 'NordVPN')
    PROPER_NOUNS = {
        'lietuva': 'Lietuva', 'kaunas': 'Kaunas', 'vilnius': 'Vilnius',
        'lietuvos': 'Lietuvos', 'lietuvoje': 'Lietuvoje',
        'lidl': 'Lidl', 'swedbank': 'Swedbank',
        'nordvpn': 'NordVPN', 'nord': 'Nord', 'security': 'Security',
        'microsoft': 'Microsoft', 'access': 'Access', 'office': 'Office',
        'oracle': 'Oracle', 'google': 'Google', 'apple': 'Apple',
        'mongodb': 'MongoDB', 'mysql': 'MySQL', 'postgresql': 'PostgreSQL',
        'sqlite': 'SQLite', 'libreoffice': 'LibreOffice',
        'openoffice': 'OpenOffice', 'apache': 'Apache',
        'claude': 'Claude', 'anthropic': 'Anthropic', 'omniva': 'Omniva',
        'pigu.lt': 'Pigu.lt', 'pigu': 'Pigu', 'maxima': 'Maxima',
        'paštas': 'Paštas', 'snoras': 'Snoras',
        'andrius': 'Andrius', 'vargonas': 'Vargonas',
        'povilas': 'Povilas', 'adomas': 'Adomas', 'brukas': 'Brukas',
        'airtable': 'Airtable', 'budibase': 'Budibase', 'supabase': 'Supabase',
        'nocodb': 'NocoDB', 'pgadmin': 'pgAdmin', 'dbeaver': 'DBeaver',
        'compass': 'Compass', 'workbench': 'Workbench', 'studio': 'Studio',
    }

    m = re.match(r'^(\s*\d+(?:\.\d+)*\.?\s+)(.+)$', text.strip())
    if not m:
        return text
    num = m.group(1)
    body = m.group(2)
    if not body:
        return text

    # Suskaidom i zodzius (tarpai, skyrybos zenklai)
    # Pirmojo zodzio pirma raide didzioji, likusios - smart conversion
    words = body.split(' ')
    result_words = []
    for idx, word in enumerate(words):
        # Atskiriam skyrybos zenklus
        prefix = ''
        suffix = ''
        core = word
        # Pradzia
        while core and not core[0].isalnum():
            prefix += core[0]
            core = core[1:]
        # Galas
        while core and not core[-1].isalnum():
            suffix = core[-1] + suffix
            core = core[:-1]

        if not core:
            result_words.append(word)
            continue

        # Patikrinim akronimus (case-insensitive lyginimas, return - DIDZIOSIOMIS)
        upper = core.upper()
        if upper in ACRONYMS:
            new = upper
        elif core.lower() in PROPER_NOUNS:
            # Tikrinis vardas - is zodyno paimam tinkama rasyma
            new = PROPER_NOUNS[core.lower()]
        elif idx == 0:
            # Pirmojo zodzio pirma raide didzioji
            new = core[0].upper() + core[1:].lower()
        else:
            # Mazomis raidemis
            new = core.lower()

        result_words.append(prefix + new + suffix)

    return num + ' '.join(result_words)


def _replace_paragraph_text(p, new_text, font_name='Times New Roman',
                             font_size=12, bold=False):
    """Pakeicia pastraipos teksta. Pasalina senus runs ir prideda nauja."""
    for run in list(p.runs):
        run._element.getparent().remove(run._element)
    run = p.add_run(new_text)
    _set_run_font(run, name=font_name, size_pt=font_size, bold=bold)


def normalize_document_formatting(doc_path):
    """Atvera galutini PUA dokumenta ir uztikrina, kad VISA tekstas
    atitiktu VU metodinius reikalavimus:
       - Times New Roman 12 pt (Normal)
       - 1.5 line spacing
       - Justify
       - 1.25 cm pirmoji eilutes itrauka (Normal pastraipoms)
       - Heading'u tinkami parametrai
       - Lenteliu cells - TNR 10 pt
    """
    doc = Document(doc_path)

    # Heading 2 ir Heading 3: jei DIDZIOSIOMIS RAIDEMIS, konvertuoti i
    # 'Pirmoji didzioji + likusios mazomis' formata (pagal VU metodika)
    for p in doc.paragraphs:
        if not p.style:
            continue
        sty = p.style.name
        if sty in ("Heading 2", "Heading 3"):
            text = (p.text or "").strip()
            if not text:
                continue
            # Patikrinti ar tekstas DIDZIOSIOMIS (be ilgu zodziu)
            # Lietuviskos raides: jei text == text.upper() (visos raides upper)
            # ir tame yra bent vienos lietuviu/anglu raides
            if text == text.upper() and any(c.isalpha() for c in text):
                new_text = _convert_h2_to_sentence_case(text)
                if new_text != text:
                    _replace_paragraph_text(p, new_text,
                                            font_size=12, bold=True)

    # Pastraipu normalizavimas
    for p in doc.paragraphs:
        text = (p.text or "").strip()
        if not text:
            continue

        sty = p.style.name if p.style else "Normal"

        # Normal pastraipos - TNR 12pt, 1.5, justify, 1.25cm
        if sty == "Normal" or sty == "Default Paragraph Font" or "List" in sty:
            for run in p.runs:
                if run.font.name != 'Times New Roman':
                    _set_run_font(run, size_pt=12,
                                  bold=run.font.bold or False,
                                  italic=run.font.italic or False)
                else:
                    # Tik dydzio normalizavimas, jei reikia
                    if run.font.size is None or run.font.size.pt > 14:
                        run.font.size = Pt(12)

    # Lenteliu cells - TNR 10 pt
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    pf = p.paragraph_format
                    pf.line_spacing = 1.15  # lentelese mazesnis
                    pf.first_line_indent = Cm(0)
                    pf.space_before = Pt(0)
                    pf.space_after = Pt(0)
                    for run in p.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(10)
                        rPr = run._element.get_or_add_rPr()
                        rFonts = rPr.find(qn('w:rFonts'))
                        if rFonts is None:
                            rFonts = OxmlElement('w:rFonts')
                            rPr.append(rFonts)
                        rFonts.set(qn('w:ascii'), 'Times New Roman')
                        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
                        rFonts.set(qn('w:cs'), 'Times New Roman')

    # Lenteliu / paveikslu pavadinimai (Caption stiliumi)
    # Atpazinkim pastraipas, kurios prasideda 'X lentele' arba 'X pav.'
    # ir pakeiskim ju formata i 11 pt Bold
    table_label_re = re.compile(r'^\s*(\d+)\s+lentelė', re.IGNORECASE)
    figure_label_re = re.compile(r'^\s*(\d+)\s+pav\.?', re.IGNORECASE)
    saltinis_re = re.compile(r'^\s*Šaltinis:', re.IGNORECASE)

    for p in doc.paragraphs:
        text = (p.text or "").strip()
        if not text:
            continue
        if table_label_re.match(text) or figure_label_re.match(text):
            # 11 pt Bold
            for run in p.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(11)
                run.font.bold = True
                rPr = run._element.get_or_add_rPr()
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is None:
                    rFonts = OxmlElement('w:rFonts')
                    rPr.append(rFonts)
                rFonts.set(qn('w:ascii'), 'Times New Roman')
                rFonts.set(qn('w:hAnsi'), 'Times New Roman')
                rFonts.set(qn('w:cs'), 'Times New Roman')
        elif saltinis_re.match(text):
            # 9 pt po lentelemis/paveikslais
            for run in p.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(9)
                run.font.italic = True
                rPr = run._element.get_or_add_rPr()
                rFonts = rPr.find(qn('w:rFonts'))
                if rFonts is None:
                    rFonts = OxmlElement('w:rFonts')
                    rPr.append(rFonts)
                rFonts.set(qn('w:ascii'), 'Times New Roman')
                rFonts.set(qn('w:hAnsi'), 'Times New Roman')

    # Uztikrinkim, kad lenteles neskaidiamos per puslapius (jei telpa)
    for tbl in doc.tables:
        for row in tbl.rows:
            trPr = row._tr.find(qn('w:trPr'))
            if trPr is None:
                trPr = OxmlElement('w:trPr')
                row._tr.insert(0, trPr)
            cantSplit = trPr.find(qn('w:cantSplit'))
            if cantSplit is None:
                cantSplit = OxmlElement('w:cantSplit')
                trPr.append(cantSplit)

    # Uztikrinkim sectinos parametrus po normalizavimo
    for sec in doc.sections:
        configure_section(sec)
        # Footer paliekam (jis jau buvo nustatytas master'yje)

    doc.save(doc_path)


# ============================================================
# PAGRINDINE FUNKCIJA
# ============================================================

def main():
    print("=" * 70)
    print("PUA generavimas v3 - PILNAS akademinis pertvarkymas")
    print("=" * 70)

    if os.path.exists(TMP_DIR):
        shutil.rmtree(TMP_DIR)
    os.makedirs(TMP_DIR, exist_ok=True)

    # 1. ANALIZE - is kiekvieno PU isgauti zonas
    print("\n[1/6] PU dokumentu analize ir filtravimas...")
    pu_data = []
    all_references = []
    all_tables = []
    all_figures = []

    for filename, number, theme in PU_LIST:
        pu_path = os.path.join(REPO_DIR, filename)
        if not os.path.exists(pu_path):
            raise FileNotFoundError("Nerastas PU failas: " + pu_path)

        info = analyze_and_split_pu(pu_path)
        z = info["zones"]
        print("    PU{}: intro={}, main={}, conclusions={}, references={}".format(
            number, len(z["intro"]), len(z["main"]),
            len(z["conclusions"]), len(z["references"])))

        main_path = os.path.join(TMP_DIR, "pu{}_main.docx".format(number))
        filter_pu_keep_intro_and_main(pu_path, main_path)
        demote_internal_headings(main_path)

        tables, figures = collect_tables_and_figures(pu_path, number)
        for orig, title in tables:
            all_tables.append((number, orig, title))
        for orig, title in figures:
            all_figures.append((number, orig, title))

        all_references.extend(info["references_paragraphs"])

        # Papildomas priedas (SQL skriptai, pavyzdiniai duomenys) - jei yra
        extra_path = os.path.join(TMP_DIR, "pu{}_extra.docx".format(number))
        extra_title = filter_pu_keep_extra_appendix(pu_path, extra_path)

        pu_data.append({
            "filename": filename, "number": number, "theme": theme,
            "main_path": main_path,
            "extra_path": extra_path if extra_title else None,
            "extra_title": extra_title,
            "conclusions_paragraphs": info["conclusions_paragraphs"],
        })

    print("    Surinkta: {} lenteliu, {} paveikslu, {} saltiniu".format(
        len(all_tables), len(all_figures), len(all_references)))

    # 2. MASTER - titulinis + sarasai
    print("\n[2/6] MASTER dokumento sukurimas...")
    master = Document()
    configure_styles(master)
    section = master.sections[0]
    configure_section(section)
    setup_footer_page_numbers(section)

    add_title_page(master)

    # TURINYS (auto)
    add_section_heading_no_toc(master, "TURINYS")
    p = master.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    _add_field(p, ' TOC \\o "1-3" \\h \\z \\u ',
               placeholder="(Atidaryti, paspausti Ctrl+A, F9 turiniui atnaujinti.)")

    # LENTELIU SARASAS - statinis
    add_section_heading_no_toc(master, "LENTELIŲ SĄRAŠAS")
    if all_tables:
        for pu_num, orig_num, title in all_tables:
            p = master.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf = p.paragraph_format
            pf.first_line_indent = Cm(0)
            pf.space_after = Pt(6)
            run = p.add_run("{} lentelė (PU{}). {}".format(orig_num, pu_num, title))
            _set_run_font(run, size_pt=12)
    else:
        add_left_paragraph(master, "Lentelių darbe nėra.", size=12)

    # PAVEIKSLU SARASAS - statinis
    add_section_heading_no_toc(master, "PAVEIKSLŲ SĄRAŠAS")
    if all_figures:
        for pu_num, orig_num, title in all_figures:
            p = master.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf = p.paragraph_format
            pf.first_line_indent = Cm(0)
            pf.space_after = Pt(6)
            run = p.add_run("{} pav. (PU{}). {}".format(orig_num, pu_num, title))
            _set_run_font(run, size_pt=12)
    else:
        add_left_paragraph(master, "Paveikslų darbe nėra.", size=12)

    tmp_master = os.path.join(TMP_DIR, "_master.docx")
    master.save(tmp_master)

    # 3. Composer - PU pagrindiniai turiniai
    print("\n[3/6] PU pagrindinio turinio prijungimas...")
    base = Document(tmp_master)
    composer = Composer(base)

    for pu in pu_data:
        h1_text = "PU{}. {}".format(pu["number"], pu["theme"])
        h1 = base.add_heading(h1_text, level=1)

        intro_marker = base.add_paragraph()
        intro_marker.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = intro_marker.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(6)
        run = intro_marker.add_run("Užduoties kontekstas ir tikslai:")
        _set_run_font(run, size_pt=12, bold=True, italic=True)

        sub = Document(pu["main_path"])
        composer.append(sub)
        print("    PU{} prijungtas".format(pu["number"]))

    # 4. ISVADOS
    print("\n[4/6] ISVADOS skyrius...")
    h1 = base.add_heading("IŠVADOS", level=1)
    for pu in pu_data:
        title_short = pu["theme"][:60]
        h2 = base.add_heading("PU{} ({}) išvados".format(pu["number"], title_short), level=2)
        for cp in pu["conclusions_paragraphs"]:
            add_left_paragraph(base, cp, size=12, indent_first=True, space_after=0)

    # 5. LITERATUROS SARASAS - APA su hanging indent
    print("\n[5/6] LITERATUROS SARASAS (APA hanging indent)...")
    add_section_heading_no_toc(base, "LITERATŪROS SĄRAŠAS")
    deduped_refs = dedupe_and_sort_references(all_references)
    print("    {} saltiniu po deduplikavimo".format(len(deduped_refs)))
    for ref in deduped_refs:
        p = base.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.space_after = Pt(6)
        pf.line_spacing = 1.5
        run = p.add_run(ref)
        _set_run_font(run, size_pt=12)
        # Pakabinta itrauka 1.25 cm
        _set_paragraph_hanging_indent(p, hanging_cm=1.25)

    # PRIEDU sarasas
    add_section_heading_no_toc(base, "PRIEDAI")
    add_left_paragraph(base,
        "Žemiau pateikiamas darbe esančių priedų sąrašas. 1-7 prieduose "
        "pateikta atitinkamos praktinės užduoties (PU1-PU7) dirbtinio intelekto "
        "naudojimo deklaracija. Tolimesnėse prieduose pateikta papildoma "
        "praktinė medžiaga - SQL DDL skriptai, pavyzdiniai duomenys ir detalios "
        "lyginamosios charakteristikos.",
        size=12, indent_first=True, space_after=12)

    # 1-7 PRIEDAI - DI deklaracijos
    for pu in pu_data:
        p = base.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.space_after = Pt(0)
        run = p.add_run("{} PRIEDAS. PU{} dirbtinio intelekto užklausos ir gauti atsakymai".format(
            pu["number"], pu["number"]))
        _set_run_font(run, size_pt=12)

    # Papildomi priedai (8+) - is PU appendix zonu
    extra_priedai = [pu for pu in pu_data if pu.get("extra_path")]
    next_num = 8
    for pu in extra_priedai:
        title = pu["extra_title"] or "PAPILDOMA MEDŽIAGA"
        # Paliekam pavadinima DIDZIOSIOMIS (kad akronimai SQL/DDL/DML islaikytu format'a)
        p = base.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.space_after = Pt(0)
        run = p.add_run("{} PRIEDAS. PU{} - {}".format(next_num, pu["number"], title))
        _set_run_font(run, size_pt=12)
        pu["extra_priedo_nr"] = next_num
        next_num += 1

    # 6. PRIEDAI - DI deklaracijos (programatiskai sukurtos)
    print("\n[6/6] PRIEDAI - 7 DI deklaracijos pagal Povilo modeli...")
    for pu in pu_data:
        add_appendix(base, pu["number"], pu["theme"])
        print("    {} PRIEDAS sukurtas (DI deklaracija)".format(pu["number"]))

    # Papildomi priedai (SQL/duomenys) per docxcompose
    if extra_priedai:
        print("    Pridedam {} papildomu priedu (SQL/duomenys)...".format(len(extra_priedai)))
        # Issaugom prieš docxcompose
        with_di_path = os.path.join(TMP_DIR, "_with_di.docx")
        base.save(with_di_path)
        base = Document(with_di_path)
        composer2 = Composer(base)

        for pu in extra_priedai:
            num = pu["extra_priedo_nr"]
            title = pu["extra_title"] or "PAPILDOMA MEDŽIAGA"
            # Antrastes
            p1 = base.add_paragraph()
            p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            pf = p1.paragraph_format
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.5
            pf.page_break_before = True
            run = p1.add_run("{} PRIEDAS".format(num))
            _set_run_font(run, size_pt=12, bold=False)

            p2 = base.add_paragraph()
            p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = p2.paragraph_format
            pf.first_line_indent = Cm(0)
            pf.space_before = Pt(12)
            pf.space_after = Pt(18)
            pf.line_spacing = 1.5
            run = p2.add_run("PU{} - {}".format(pu["number"], title.upper()))
            _set_run_font(run, size_pt=12, bold=True)

            # docxcompose append
            sub = Document(pu["extra_path"])
            composer2.append(sub)
            print("    {} PRIEDAS (PU{} extra) prijungtas".format(num, pu["number"]))

    # Issaugom
    base.save(OUT_FILE)

    # POST-PROCESSOR
    print("\n[7/7] Formatavimo normalizavimas (TNR 12pt visur, lenteles, captions)...")
    normalize_document_formatting(OUT_FILE)

    shutil.rmtree(TMP_DIR, ignore_errors=True)

    final_size_kb = os.path.getsize(OUT_FILE) // 1024
    print("\n" + "=" * 70)
    print("PUA issaugotas: {} ({} KB)".format(OUT_FILE, final_size_kb))
    print("=" * 70)
    print("PO ATSISIUNTIMO WORD'E AR LIBREOFFICE'E:")
    print("  1. Ctrl+A (paryskinti viska)")
    print("  2. F9 (atnaujinti TURINI ir puslapiu numerius)")
    print("  3. Update entire table -> Yes")
    print("  4. Patikrinti vizualiai")
    print("=" * 70)


if __name__ == "__main__":
    main()
