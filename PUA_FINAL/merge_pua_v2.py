"""
PUA (Praktiniu uzduociu ataskaita) sujungimo skriptas v2.

Pagal "Povilo modeli" - vientisas akademinis dokumentas pagal VU KnF reikalavimus:
    1. Vienas titulinis lapas
    2. TURINYS (auto TOC field)
    3. LENTELIU SARASAS (statinis - sugeneruotas is PU lenteliu)
    4. PAVEIKSLU SARASAS (statinis - sugeneruotas is PU paveikslu)
    5. PU1.-PU7. (kiekvienas Heading 1 su tema; vidiniai TOC'ai, tituliniai
       lapai pasalinami; ivadas integruojamas po H1 antrastes)
    6. ISVADOS (vienas bendras skyrius su PU1-PU7 isvadu poskyriais)
    7. LITERATUROS SARASAS (sujungtas, deduplikuotas, abeceles tvarka)
    8. PRIEDAI (1 PRIEDAS - 7 PRIEDAS: DI deklaracijos ir papildoma medziaga)

Naudojimas:
    pip install python-docx docxcompose
    python3 merge_pua_v2.py

Autorius: Andrius Vargonas, 2026
"""

import os
import re
import shutil
from copy import deepcopy

from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor
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

# PU duomenys: (failas, numeris, oficiali tema antrastei)
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

# OOXML namespace
W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


# ============================================================
# OOXML pagalbines funkcijos
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
    # Pataisom dubliavima: kartais Word TOC/SDT blokai dubliuoja teksta 2-3 kartus
    # PU1 ir PU2 atvejais 'ĮVADAS' tampa 'ĮVADASĮVADASĮVADAS'
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


def _add_page_break(paragraph):
    run = paragraph.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)


# ============================================================
# Klasifikavimo logika
# ============================================================

def _classify_h1(text):
    """Grazina zonos tipa pagal Heading 1 teksta."""
    t = (text or "").strip().upper()
    if not t:
        return "empty"
    # Ivadas
    if t.startswith("ĮVAD") or t.startswith("IVAD"):
        return "intro"
    # Isvados
    if t.startswith("IŠVAD") or t.startswith("ISVAD"):
        return "conclusions"
    # Literaturos sarasas
    if t.startswith("LITERAT"):
        return "references"
    # DI deklaracija (PU2 atveju ji yra Heading 1 PRIES isvadas)
    if "DI NAUDOJ" in t or "DIRBTINIO INTEL" in t:
        return "ai_decl"
    # Priedas (jei butu Heading 1)
    if "PRIED" in t:
        return "ai_decl"
    # Pagrindinis turinys: prasideda "1.", "2." ir t.t.
    if t and t[0].isdigit() and "." in t[:4]:
        return "main"
    return "other"


def _is_priedas_marker(text):
    """Patikrina ar pastraipa zymi 'X PRIEDAS' (paprasta, ne Heading).
    Tolerantiska del Word TOC/SDT dubliavimo, pvz. '1 PRIEDAS11 PRIEDAS PRIEDAS'.
    """
    t = (text or "").strip().upper()
    if not t:
        return False
    # Pradzia: skaicius + space/dot + PRIEDAS (be $ - tolerantiska dubliavimui)
    return bool(re.match(r'^\d+\s*\.?\s*PRIEDAS', t))


# ============================================================
# PU dokumento analize ir filtravimas
# ============================================================

def analyze_and_split_pu(pu_path):
    """State machine - klasifikuoja kiekviena block elementa i zona:
       PRE | INTRO | MAIN | CONCLUSIONS | REFERENCES | APPENDIX

    Pereinama riba pagal Heading 1 markerius (ĮVADAS, IŠVADOS, LITERATŪROS,
    DI NAUDOJIMO DEKLARACIJA) ARBA pagal Normal stiliaus 'X. PAVADINIMAS'
    pastraipas (PU1 atveju), arba pagal 'X PRIEDAS' markerius.

    Grazina dict su block'u indeksais kiekvienai zonai.
    """
    doc = Document(pu_path)
    body = doc.element.body
    block_elems = [c for c in body if _is_paragraph(c) or _is_table(c)]

    # Regex pagrindiniam turinio skyriui (PU1 atveju Normal stiliumi)
    main_section_re = re.compile(r'^(\d+)\.\s+([A-ZĄČĘĖĮŠŲŪŽ][A-ZĄČĘĖĮŠŲŪŽ\s\-,IRTŲ]{4,})$')

    zones = {
        "intro": [],
        "main": [],
        "conclusions": [],
        "references": [],
        "appendix": [],
    }
    state = "PRE"

    for bi, el in enumerate(block_elems):
        if _is_paragraph(el):
            text = _get_paragraph_text(el)
            is_h1 = _is_heading1(el)
            zone_h1 = _classify_h1(text) if is_h1 else None

            # ===== H1 markeriai =====
            if is_h1 and zone_h1 == "intro":
                state = "INTRO"
                continue  # H1 ĮVADAS pacios netraukiam
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
                # H1 "X." Heading 1 - traukiam i main
                state = "MAIN"
                zones["main"].append(bi)
                continue

            # ===== Normal pastraipos atpazinimas =====
            # 'X PRIEDAS' markeris (ne Heading) - PRIEDU pradzia
            if _is_priedas_marker(text):
                state = "APPENDIX"
                continue

            # Normal "X. PAVADINIMAS" - pagrindinio skyriaus pradzia
            # (PU1 atvejis, kur skyriai ne Heading 1)
            # Bet TIK po ivado (state == "INTRO") arba kai jau MAIN
            if not is_h1 and state in ("INTRO", "MAIN"):
                if main_section_re.match(text):
                    state = "MAIN"
                    zones["main"].append(bi)
                    continue

        # Lentele arba kt. - traukiam pagal state
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

    # Issprausk teksta isvadu ir literaturos pastraipoms
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
        "doc": doc,
        "body": body,
        "block_elems": block_elems,
        "zones": zones,
        "conclusions_paragraphs": conclusions_paragraphs,
        "references_paragraphs": references_paragraphs,
    }


def filter_pu_keep_intro_and_main(pu_path, dst_path):
    """Sukuria filtruota PU dokumento kopija, paliekant TIK ivado pastraipas
    (po H1 ĮVADAS) + pagrindini turini (po pirmos H1 'X.' iki pries IŠVADAS).
    Pasalinami: titulinis, vidinis TURINYS, H1 ĮVADAS pati, IŠVADAS, LITERATŪRA,
    DI deklaracija, priedai."""
    info = analyze_and_split_pu(pu_path)
    body = info["body"]
    block_elems = info["block_elems"]

    keep = set()
    for bi in info["zones"]["intro"]:
        keep.add(id(block_elems[bi]))
    for bi in info["zones"]["main"]:
        keep.add(id(block_elems[bi]))

    # Pasalink visus kt. block elementus is body'o
    for el in list(body):
        if _is_paragraph(el) or _is_table(el):
            if id(el) not in keep:
                body.remove(el)

    info["doc"].save(dst_path)


def demote_internal_headings(filtered_doc_path):
    """Po PU filtravimo, perlygina antrasciu hierarchija:
       - Originalus Heading 1 (PU vidiniai 'X.' skyriai) -> Heading 2
       - Originalus Heading 2 ('X.X.' poskyriai) -> Heading 3
       - PU1 atveju Normal '1.' '2.' '3.' (didziosios raides) -> Heading 2

    Tai uztikrina, kad PUA Heading 1 lygmenyje liks tik 'PUx. TEMA' antrastes.
    """
    doc = Document(filtered_doc_path)

    # Surinkim visus paragraphus is body (ne tik doc.paragraphs - tos pacios)
    main_section_re = re.compile(r'^(\d+)\.\s+([A-ZĄČĘĖĮŠŲŪŽ][A-ZĄČĘĖĮŠŲŪŽ\s\-,]{4,})')

    # Pirmas etapas - pakeisti stiliu
    for p in doc.paragraphs:
        text = _get_paragraph_text(p._element)
        if not text:
            continue

        if _is_heading1(p._element):
            # Heading 1 -> Heading 2
            p.style = doc.styles['Heading 2']
        else:
            # Patikrink ar tai Normal pastraipa su 'X. PAVADINIMAS' (PU1 atvejis)
            cur_style = p.style.name if p.style else ""
            if cur_style in ('Normal', 'Default Paragraph Font') or cur_style.startswith('Normal'):
                if main_section_re.match(text):
                    p.style = doc.styles['Heading 2']

    # Antras etapas - originalus Heading 2 -> Heading 3
    # Bet musu pirmas etapas jau pakeite kai kuriuos H1 i H2,
    # taigi po pirmo etapo turime "uzslepta" originalu H2 pavadinima.
    # Reikia atskirti: po pirmo etapo H2 yra arba (a) buves H1, arba (b) tikras originalus H2
    # Laimei, mes galim atpazinti is teksto formato:
    # - "X. PAVADINIMAS" (didziosiomis) - (a) buves H1
    # - "X.X. Pavadinimas" (mazomis) - (b) tikras originalus H2 (turi buti H3)
    h2_subsection_re = re.compile(r'^(\d+\.\d+\.?)\s+')

    for p in doc.paragraphs:
        text = _get_paragraph_text(p._element)
        if not text:
            continue

        if p.style and p.style.name == 'Heading 2':
            if h2_subsection_re.match(text):
                # Tai originalus PU H2 (X.X.) -> demote i H3
                p.style = doc.styles['Heading 3']

    doc.save(filtered_doc_path)


def filter_pu_keep_appendix(pu_path, dst_path):
    """Filtruoja PU dokumenta, paliekant TIK 'X PRIEDAS' / DI deklaracijos zona."""
    info = analyze_and_split_pu(pu_path)
    body = info["body"]
    block_elems = info["block_elems"]

    keep = set()
    for bi in info["zones"]["appendix"]:
        keep.add(id(block_elems[bi]))

    for el in list(body):
        if _is_paragraph(el) or _is_table(el):
            if id(el) not in keep:
                body.remove(el)

    info["doc"].save(dst_path)


# ============================================================
# Lenteliu / paveikslu sarasu generavimas
# ============================================================

def collect_tables_and_figures(pu_path, pu_number):
    """Iteruoja per PU dokumenta. Surask pastraipas, kurios atitinka:
       - 'X lentele' (tada SEKANCIA pastraipa = pavadinimas)
       - 'X pav. ...' arba 'X pav...'  (tame paciame teksto bloke = pavadinimas)
    
    Grazina:
        tables: list of (originalus_nr, pavadinimas)
        figures: list of (originalus_nr, pavadinimas)
    """
    doc = Document(pu_path)
    paragraphs = doc.paragraphs
    tables = []
    figures = []

    # Re patternai
    table_label_re = re.compile(r'^\s*(\d+)\s+lentelė\s*$', re.IGNORECASE)
    figure_label_re = re.compile(r'^\s*(\d+)\s+pav\.?\s*(.*)$', re.IGNORECASE)

    for i, p in enumerate(paragraphs):
        text = (p.text or "").strip()
        if not text:
            continue

        # Lentele - vien tik "X lentele" be teksto
        m = table_label_re.match(text)
        if m:
            num = int(m.group(1))
            # Pavadinimas - sekanti pastraipa
            if i + 1 < len(paragraphs):
                title = (paragraphs[i+1].text or "").strip()
            else:
                title = ""
            if title:
                tables.append((num, title))
            continue

        # Paveikslas - "X pav. Pavadinimas" arba "X pav Pavadinimas"
        m = figure_label_re.match(text)
        if m:
            num = int(m.group(1))
            title = m.group(2).strip()
            if not title:
                # Gal pavadinimas pries (paveikslo title virsuje?)
                if i - 1 >= 0:
                    prev = (paragraphs[i-1].text or "").strip()
                    if prev and len(prev) < 200 and not table_label_re.match(prev) and not figure_label_re.match(prev):
                        title = prev
            # Cia title gali buti ir tuscias - kartais paveikslas yra ant titulinio
            if title:
                figures.append((num, title))
            continue

    return tables, figures


# ============================================================
# Literatūros sąrašo deduplikavimas
# ============================================================

def dedupe_and_sort_references(all_refs):
    """Suima literaturos saltinius is visu PU, deduplikuoja
    (pagal pirmus 60 simbolius case-insensitive), surusiuoja abeceles tvarka."""
    seen_keys = {}
    for ref in all_refs:
        key = re.sub(r'\s+', ' ', ref.strip().lower())[:80]
        if key not in seen_keys:
            seen_keys[key] = ref.strip()
    sorted_refs = sorted(seen_keys.values(), key=lambda s: s.lower())
    return sorted_refs


# ============================================================
# MASTER dokumento sukurimas
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

    # Heading 1 - PU skyriai (didziosiomis 14 pt center bold, page break before)
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
    pf.space_after = Pt(12)
    pf.line_spacing = 1.5
    pf.page_break_before = True

    # Heading 2 - poskyriai (12 pt left bold, mazosiomis raidemis)
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

    # Heading 3
    h3 = doc.styles['Heading 3']
    h3.font.name = 'Times New Roman'
    h3.font.size = Pt(12)
    h3.font.bold = True
    h3.font.color.rgb = RGBColor(0, 0, 0)
    pf = h3.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(12)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.5
    pf.page_break_before = False


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
                       indent_first=False, space_before=0, space_after=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.first_line_indent = Cm(1.25) if indent_first else Cm(0)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.5
    if text:
        run = p.add_run(text)
        _set_run_font(run, size_pt=size, bold=bold)
    return p


def add_section_heading_no_toc(doc, text):
    """TURINYS, PAVEIKSLU SARASAS, LENTELIU SARASAS, ISVADOS, LITERATUROS SARASAS,
    PRIEDAI - bus suskirstyti i naujus puslapius (page break before per stiliu),
    bet NEPATEKS i automatini turini.
    Naudosim tiesiog Normal stiliu su dideliu sriftu."""
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


def add_appendix_heading(doc, number, title):
    """N PRIEDAS antraste:
       Virsuje desineje 'N PRIEDAS' (12 pt didziosiomis)
       Centras 'PAVADINIMAS' (didziosiomis pusjuodziu)"""
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pf = p1.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(12)
    pf.line_spacing = 1.5
    pf.page_break_before = True
    run = p1.add_run("{} PRIEDAS".format(number))
    _set_run_font(run, size_pt=12, bold=False)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p2.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(18)
    pf.line_spacing = 1.5
    run = p2.add_run(title)
    _set_run_font(run, size_pt=12, bold=True)


# ============================================================
# Titulinis lapas
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

    # Page break i kita lapa
    last = doc.paragraphs[-1]
    _add_page_break(last)


# ============================================================
# PAGRINDINE FUNKCIJA
# ============================================================

def main():
    print("=" * 70)
    print("PUA generavimas v2 - Povilo modelis")
    print("=" * 70)

    # Pasiruosk darbinius katalogus
    if os.path.exists(TMP_DIR):
        shutil.rmtree(TMP_DIR)
    os.makedirs(TMP_DIR, exist_ok=True)

    # 1. ANALIZE - is kiekvieno PU isgauti zonas
    print("\n[1/5] Kiekvieno PU dokumento analize ir filtravimas...")
    pu_data = []
    all_references = []
    all_tables = []   # list of (pu_number, original_num, title)
    all_figures = []  # list of (pu_number, original_num, title)

    for filename, number, theme in PU_LIST:
        pu_path = os.path.join(REPO_DIR, filename)
        if not os.path.exists(pu_path):
            raise FileNotFoundError("Nerastas PU failas: " + pu_path)

        info = analyze_and_split_pu(pu_path)
        z = info["zones"]
        print("    PU{}: intro={}, main={}, conclusions={}, references={}, appendix={}".format(
            number, len(z["intro"]), len(z["main"]),
            len(z["conclusions"]), len(z["references"]), len(z["appendix"])))

        # Filtravimas - main turinys
        main_path = os.path.join(TMP_DIR, "pu{}_main.docx".format(number))
        filter_pu_keep_intro_and_main(pu_path, main_path)
        # Demotinkim vidinius skyrius (H1 -> H2, H2 -> H3)
        demote_internal_headings(main_path)

        # Filtravimas - priedai (DI deklaracija + SQL/duomenys)
        appendix_path = os.path.join(TMP_DIR, "pu{}_appendix.docx".format(number))
        filter_pu_keep_appendix(pu_path, appendix_path)

        # Surink lenteliu/paveikslu pavadinimus (is originalo, kad neprasarsutume)
        tables, figures = collect_tables_and_figures(pu_path, number)
        for orig, title in tables:
            all_tables.append((number, orig, title))
        for orig, title in figures:
            all_figures.append((number, orig, title))

        # Surink literaturos saltinius
        all_references.extend(info["references_paragraphs"])

        pu_data.append({
            "filename": filename,
            "number": number,
            "theme": theme,
            "main_path": main_path,
            "appendix_path": appendix_path,
            "conclusions_paragraphs": info["conclusions_paragraphs"],
        })

    print("\n  Surinkta:")
    print("    Lenteliu pavadinimu: {}".format(len(all_tables)))
    print("    Paveikslu pavadinimu: {}".format(len(all_figures)))
    print("    Literaturos saltiniu (su dublikatais): {}".format(len(all_references)))

    # 2. MASTER dokumentas
    print("\n[2/5] MASTER dokumento sukurimas (titulinis + sarasai)...")
    master = Document()
    configure_styles(master)
    section = master.sections[0]
    configure_section(section)
    setup_footer_page_numbers(section)

    # Titulinis lapas
    add_title_page(master)

    # TURINYS (Word TOC field)
    add_section_heading_no_toc(master, "TURINYS")
    p = master.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    _add_field(p, ' TOC \\o "1-3" \\h \\z \\u ',
               placeholder="(Po atidarymo paspauskite Ctrl+A, F9 turiniui atnaujinti.)")

    # LENTELIU SARASAS - statinis sarasas su PU prefiksu
    add_section_heading_no_toc(master, "LENTELIŲ SĄRAŠAS")
    if all_tables:
        for pu_num, orig_num, title in all_tables:
            p = master.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf = p.paragraph_format
            pf.first_line_indent = Cm(0)
            pf.space_after = Pt(0)
            run = p.add_run("{} lentelė (PU{}). {}".format(orig_num, pu_num, title))
            _set_run_font(run, size_pt=12)
    else:
        add_left_paragraph(master, "Lentelių darbe nėra.", size=12)

    # PAVEIKSLU SARASAS - statinis sarasas su PU prefiksu
    add_section_heading_no_toc(master, "PAVEIKSLŲ SĄRAŠAS")
    if all_figures:
        for pu_num, orig_num, title in all_figures:
            p = master.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf = p.paragraph_format
            pf.first_line_indent = Cm(0)
            pf.space_after = Pt(0)
            run = p.add_run("{} pav. (PU{}). {}".format(orig_num, pu_num, title))
            _set_run_font(run, size_pt=12)
    else:
        add_left_paragraph(master, "Paveikslų darbe nėra.", size=12)

    # Issaugok master pries Composer
    tmp_master = os.path.join(TMP_DIR, "_master.docx")
    master.save(tmp_master)

    # 3. Composer - prijungiam PU pagrindini turini
    print("\n[3/5] PU dokumentu pagrindinio turinio prijungimas...")
    base = Document(tmp_master)
    composer = Composer(base)

    for pu in pu_data:
        # Pridek H1 antraste pries PU
        heading_text = "PU{}. {}".format(pu["number"], pu["theme"])
        h1 = base.add_heading(heading_text, level=1)
        # Padek "Užduoties tekstas:" prefiksa pries pagrindini turini
        # (jis turi buti integruotas i pagrindini PU dokumenta)
        intro_marker = base.add_paragraph()
        intro_marker.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = intro_marker.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(6)
        run = intro_marker.add_run("Užduoties kontekstas ir tikslai:")
        _set_run_font(run, size_pt=12, bold=True, italic=True)

        # Composer prijungia filtruota PU (intro + main)
        sub = Document(pu["main_path"])
        composer.append(sub)

        size_kb = os.path.getsize(pu["main_path"]) // 1024
        print("    PU{} prijungtas (filtruotas dydis: {} KB)".format(pu["number"], size_kb))

    # 4. ISVADOS - vienas bendras skyrius
    print("\n[4/5] Bendras ISVADOS skyrius...")
    h1 = base.add_heading("IŠVADOS", level=1)

    for pu in pu_data:
        # Heading 2 "PU{} isvados"
        h2 = base.add_heading("PU{} ({}) išvados".format(pu["number"], pu["theme"][:50]), level=2)
        # Pridek isvadu pastraipas
        for cp in pu["conclusions_paragraphs"]:
            add_left_paragraph(base, cp, size=12, indent_first=True, space_after=0)

    # 5. LITERATUROS SARASAS - sujungtas, deduplikuotas, abeceles tvarka
    print("\n[5/5] Bendras LITERATUROS SARASAS ir PRIEDAI...")
    add_section_heading_no_toc(base, "LITERATŪROS SĄRAŠAS")
    deduped_refs = dedupe_and_sort_references(all_references)
    print("    Sujungtu literaturos saltiniu (po deduplikavimo): {}".format(len(deduped_refs)))
    for ref in deduped_refs:
        p = base.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        # APA stilius: pakabinta itrauka 1.25 cm
        pf.left_indent = Cm(1.25)
        pf.space_after = Pt(6)
        try:
            from docx.shared import Twips
            # nustatom hanging indent
            from docx.oxml.ns import qn as _qn2
            pPr = p._p.get_or_add_pPr()
            ind = pPr.find(_qn2('w:ind'))
            if ind is None:
                ind = OxmlElement('w:ind')
                pPr.append(ind)
            ind.set(_qn2('w:hanging'), str(int(1.25 * 567)))  # 1.25 cm = 708 twips
            ind.set(_qn2('w:left'), str(int(1.25 * 567)))
        except Exception:
            pass
        run = p.add_run(ref)
        _set_run_font(run, size_pt=12)

    # PRIEDAI - sąrašas + 7 priedai
    add_section_heading_no_toc(base, "PRIEDAI")
    add_left_paragraph(base,
        "Žemiau pateikiamas darbe esančių priedų sąrašas. Kiekviename priede pateikiama "
        "atitinkamos praktinės užduoties dirbtinio intelekto naudojimo deklaracija ir "
        "papildoma medžiaga (SQL skriptai, pavyzdiniai duomenys ir kt.).",
        size=12, indent_first=True, space_after=12)

    for pu in pu_data:
        p = base.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.space_after = Pt(0)
        run = p.add_run("{} PRIEDAS. PU{} dirbtinio intelekto naudojimo deklaracija ir papildoma medžiaga".format(
            pu["number"], pu["number"]))
        _set_run_font(run, size_pt=12)

    # Issaugok base pries pridedant priedus per Composer
    base_with_lit = os.path.join(TMP_DIR, "_base_with_lit.docx")
    base.save(base_with_lit)

    # Atidaryk is naujo per Composer'i ir prijungti priedus
    final_base = Document(base_with_lit)
    composer2 = Composer(final_base)

    for pu in pu_data:
        # Antraste "X PRIEDAS / pavadinimas"
        appendix_title = "PU{} DIRBTINIO INTELEKTO NAUDOJIMO DEKLARACIJA IR PAPILDOMA MEDZIAGA".format(pu["number"])
        add_appendix_heading(final_base, pu["number"], appendix_title)
        # Composer prijungia priedu turini
        if os.path.exists(pu["appendix_path"]):
            sub = Document(pu["appendix_path"])
            composer2.append(sub)

    # Issaugom galutini
    composer2.save(OUT_FILE)

    # Pasvalink laikina
    shutil.rmtree(TMP_DIR, ignore_errors=True)

    final_size_kb = os.path.getsize(OUT_FILE) // 1024
    print("\n" + "=" * 70)
    print("PUA issaugotas: {} ({} KB)".format(OUT_FILE, final_size_kb))
    print("=" * 70)
    print("ATIDARE WORD'E AR LIBREOFFICE'E:")
    print("  1. Ctrl+A (paryskinti viska)")
    print("  2. F9 (atnaujinti TURINYS)")
    print("  3. Patikrinti vizualiai")
    print("=" * 70)


if __name__ == "__main__":
    main()
