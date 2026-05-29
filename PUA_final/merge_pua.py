"""
PUA - Praktinių užduočių ataskaita - sujungimo skriptas.

Sujungia PU1, PU3, PU4, PU5, PU6, PU7 dokumentus į vieną PUA dokumentą:
- Vienas titulinis lapas
- Automatinis turinys (Word TOC field)
- Automatinis paveikslų sąrašas (TOC \\h \\c "pav")
- Automatinis lentelių sąrašas (TOC \\h \\c "lentelė")
- 7 skyriai (PU2 - placeholder, nes užduotis nepateikta)
- Sklandus paveikslų ir lentelių persinumeravimas (1..N)
- Skyrių poskyrių persinumeravimas pagal naują skyriaus numerį

Autorius: Andrius Vargonas, VU Kauno fakultetas
"""

from __future__ import annotations

import copy
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsmap
from docx.shared import Cm, Mm, Pt, RGBColor
from docxcompose.composer import Composer

# ----------------------------------------------------------------------------
# Konfigūracija
# ----------------------------------------------------------------------------

PUA_DIR = Path("/projects/sandbox/PUA")
OUT_DIR = Path("/projects/sandbox/a/PUA_final")
OUT_DOCX = OUT_DIR / "PUA_Andrius_Vargonas.docx"

# (chapter_no, pu_id, title, source_relpath_or_None)
CHAPTERS = [
    (1, "PU1", "INFORMACINIŲ SISTEMŲ IR DUOMENŲ BAZIŲ SAMPRATA", "PU1/PU1_Andrius_Vargonas.docx"),
    (2, "PU2", "UŽDUOTIS NEPATEIKTA", None),
    (3, "PU3", "DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ ANALIZĖ IR PALYGINIMAS", "PU3/PU3_Andrius_Vargonas.docx"),
    (4, "PU4", "DBVS APLINKA, LENTELIŲ KŪRIMAS", "PU4/PU4_Andrius_Vargonas.docx"),
    (5, "PU5", "DUOMENŲ BAZĖS LENTELIŲ KŪRIMAS PAGAL ERD", "PU5/PU5_Andrius_Vargonas.docx"),
    (6, "PU6", "DUOMENŲ NORMALIZAVIMAS", "PU6/PU6_Andrius_Vargonas.docx"),
    (7, "PU7", "MS ACCESS DUOMENŲ BAZIŲ ANALIZĖ IR KŪRIMAS", "PU7/PU7_Andrius_Vargonas.docx"),
]

# Paragrafų tekstai, žymintys turinio ribas šaltiniuose (didžiosios raidės)
END_HEADINGS = ("IŠVADOS", "LITERATŪROS SĄRAŠAS", "PRIEDAI", "PRIEDAS")

# Kapitono prefiksai antraščių aptikimui
HEADING_NUM_RE = re.compile(r"^\s*(\d+(?:\.\d+)*)\.?\s+(.*)$")
PAV_RE = re.compile(r"^(\s*)(\d+)(\s*pav\b.*)$", re.IGNORECASE)
LENT_RE = re.compile(r"^(\s*)(\d+)(\s*lentel.*)$", re.IGNORECASE)


# ----------------------------------------------------------------------------
# Pagalbinės XML funkcijos
# ----------------------------------------------------------------------------

def _add_field(paragraph, instr_text: str, placeholder: str = ""):
    """Įdeda Word lauką (field) į paragrafą - veikia kaip TOC, SEQ ir pan."""
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run._r.append(fld_begin)

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instr_text
    run._r.append(instr)

    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run._r.append(fld_sep)

    if placeholder:
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = placeholder
        run._r.append(t)

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_end)


def _add_page_break(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


# ----------------------------------------------------------------------------
# Bazinis dokumento formatavimas (TNR 12pt, paraštės VU Kauno fakulteto)
# ----------------------------------------------------------------------------

def _setup_document_styles(doc):
    """Nustato dokumento stilius: TNR 12pt, paraštės, pradinio puslapio orientacija."""
    # Paraštės: 20mm kairė, 20mm dešinė, 25mm viršus, 15mm apačia (KFAK reikalavimai)
    for section in doc.sections:
        section.left_margin = Mm(25)   # KFAK reikalauja 25mm kairėje (rišimui)
        section.right_margin = Mm(15)  # 15mm dešinė
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.header_distance = Mm(12.5)
        section.footer_distance = Mm(12.5)

    # Numatytasis stilius - Times New Roman 12pt
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), "Times New Roman")
    rfonts.set(qn("w:hAnsi"), "Times New Roman")
    rfonts.set(qn("w:cs"), "Times New Roman")
    rfonts.set(qn("w:eastAsia"), "Times New Roman")

    # Heading 1, 2, 3 stiliai - TNR, paryškintas, atitinkamo dydžio
    for h_name, size in (("Heading 1", 14), ("Heading 2", 13), ("Heading 3", 12)):
        if h_name in doc.styles:
            hs = doc.styles[h_name]
            hs.font.name = "Times New Roman"
            hs.font.size = Pt(size)
            hs.font.bold = True
            hs.font.color.rgb = RGBColor(0, 0, 0)


# ----------------------------------------------------------------------------
# Titulinis lapas
# ----------------------------------------------------------------------------

def _add_title_page(doc):
    """Sukuria vieną tipinį VU Kauno fakulteto titulinį lapą."""
    # Viršutinė universiteto info
    p1 = doc.add_paragraph()
    p1.alignment = 1  # CENTER
    r1 = p1.add_run("VILNIAUS UNIVERSITETAS\nKAUNO FAKULTETAS")
    r1.bold = True
    r1.font.size = Pt(14)

    p2 = doc.add_paragraph()
    p2.alignment = 1
    r2 = p2.add_run("Programų sistemų katedra")
    r2.font.size = Pt(12)

    # Tarpas
    for _ in range(6):
        doc.add_paragraph()

    # Autorius
    p3 = doc.add_paragraph()
    p3.alignment = 1
    r3 = p3.add_run("Andrius Vargonas")
    r3.bold = True
    r3.font.size = Pt(13)

    p4 = doc.add_paragraph()
    p4.alignment = 1
    r4 = p4.add_run("Programų sistemos, II kursas")
    r4.font.size = Pt(12)

    # Tarpas
    for _ in range(2):
        doc.add_paragraph()

    # Pavadinimas
    p5 = doc.add_paragraph()
    p5.alignment = 1
    r5 = p5.add_run("PRAKTINIŲ UŽDUOČIŲ ATASKAITA (PUA)")
    r5.bold = True
    r5.font.size = Pt(18)

    p6 = doc.add_paragraph()
    p6.alignment = 1
    r6 = p6.add_run("Dalykas: Duomenų bazės")
    r6.font.size = Pt(13)

    # Tarpas
    for _ in range(8):
        doc.add_paragraph()

    # Vieta ir data
    p7 = doc.add_paragraph()
    p7.alignment = 1
    r7 = p7.add_run("Kaunas, 2026")
    r7.font.size = Pt(12)

    _add_page_break(doc)


# ----------------------------------------------------------------------------
# Turinys, paveikslų ir lentelių sąrašai
# ----------------------------------------------------------------------------

def _add_toc(doc):
    h = doc.add_paragraph("TURINYS")
    h.style = doc.styles["Heading 1"]
    h.alignment = 1
    p = doc.add_paragraph()
    _add_field(p, 'TOC \\o "1-3" \\h \\z \\u', placeholder="Spustelėkite dešiniu pelės klavišu ir pasirinkite \"Atnaujinti lauką\" turiniui sugeneruoti.")
    _add_page_break(doc)


def _add_figures_list(doc):
    h = doc.add_paragraph("PAVEIKSLŲ SĄRAŠAS")
    h.style = doc.styles["Heading 1"]
    h.alignment = 1
    p = doc.add_paragraph()
    _add_field(p, 'TOC \\h \\z \\c "pav"', placeholder="Spustelėkite dešiniu klavišu ir pasirinkite \"Atnaujinti lauką\" paveikslų sąrašui.")
    _add_page_break(doc)


def _add_tables_list(doc):
    h = doc.add_paragraph("LENTELIŲ SĄRAŠAS")
    h.style = doc.styles["Heading 1"]
    h.alignment = 1
    p = doc.add_paragraph()
    _add_field(p, 'TOC \\h \\z \\c "lentelė"', placeholder="Spustelėkite dešiniu klavišu ir pasirinkite \"Atnaujinti lauką\" lentelių sąrašui.")
    _add_page_break(doc)


# ----------------------------------------------------------------------------
# Šaltinio dokumento pjaustymas ir transformavimas
# ----------------------------------------------------------------------------

def _is_heading(paragraph) -> int | None:
    """Grąžina antraštės lygį (1, 2, 3) jei tai antraštė, kitu atveju None."""
    style = paragraph.style.name if paragraph.style else ""
    m = re.match(r"^Heading (\d+)$", style)
    if m:
        return int(m.group(1))
    if style == "Title":
        return 0
    return None


def _set_heading_style(paragraph, level: int):
    paragraph.style = paragraph.part.document.styles[f"Heading {level}"]


def _replace_paragraph_text(paragraph, new_text: str):
    """Pakeičia paragrafo tekstą, išsaugant pirmąjį run formatavimą."""
    # Išsaugome pirmąjį run kaip šabloną
    runs = paragraph.runs
    if not runs:
        paragraph.add_run(new_text)
        return
    # Pirmasis run gauna naują tekstą, kiti pašalinami
    runs[0].text = new_text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def _renumber_heading_text(text: str, chapter_no: int) -> str:
    """Į priekį pridedamas naujo skyriaus numeris.

    "1. TITLE"      -> "{chapter}.1. TITLE"
    "1.1. Subtitle" -> "{chapter}.1.1. Subtitle"
    """
    m = HEADING_NUM_RE.match(text)
    if not m:
        return text
    nums = m.group(1)  # "1" arba "1.1" arba "1.1.1"
    rest = m.group(2)
    return f"{chapter_no}.{nums}. {rest}"


def _renumber_caption(text: str, kind: str, new_num: int) -> str:
    """Persinumeruoja paveikslo arba lentelės parašą."""
    if kind == "pav":
        return PAV_RE.sub(lambda m: f"{m.group(1)}{new_num}{m.group(3)}", text)
    return LENT_RE.sub(lambda m: f"{m.group(1)}{new_num}{m.group(3)}", text)


def _para_text_clean(text: str) -> str:
    return text.strip().upper()


def _build_section_docx(
    chapter_no: int,
    pu_id: str,
    title: str,
    source_path: Path,
    fig_counter: list[int],
    tab_counter: list[int],
    out_path: Path,
):
    """Pjausto šaltinio dokumentą ir sukuria sekcijos dokumentą su transformuotu turiniu.

    - Įdeda viršuje 1 lygio antraštę "{chapter_no}. {pu_id} - {title}"
    - Įjungia tik turinį tarp pirmos numeruotos antraštės ir IŠVADOS / LITERATŪROS SĄRAŠAS
    - Demotina šaltinio antraštes (Heading 1->2, Heading 2->3, Heading 3->4)
    - Persinumeruoja antraštes (pridėdamas chapter_no. priekyje)
    - Persinumeruoja paveikslų ir lentelių parašus
    """
    src_doc = Document(str(source_path))

    # 1. Identifikuoti turinio ribas (paragrafų indeksuose)
    paragraphs = src_doc.paragraphs
    start_idx = None
    end_idx = None

    for i, p in enumerate(paragraphs):
        if _is_heading(p) is None:
            continue
        text = p.text.strip()
        text_clean = _para_text_clean(text)
        # Pradžia: pirmas Heading 1, kuris prasideda numeriu
        if start_idx is None and HEADING_NUM_RE.match(text) and _is_heading(p) == 1:
            start_idx = i
            continue
        # Pabaiga: IŠVADOS / LITERATŪROS SĄRAŠAS / PRIEDAI
        if start_idx is not None and any(text_clean.startswith(end_kw) for end_kw in END_HEADINGS):
            end_idx = i
            break

    if start_idx is None:
        raise RuntimeError(f"Negalima rasti pradžios skyriaus: {source_path}")
    if end_idx is None:
        end_idx = len(paragraphs)

    # 2. Sudaryti rinkinį leidžiamų XML elementų (paragrafų)
    keep_para_elements = set()
    for p in paragraphs[start_idx:end_idx]:
        keep_para_elements.add(p._p)

    # 3. Pakeisti šaltinio dokumento body - palikti tik leidžiamus elementus
    body = src_doc.element.body
    children = list(body)
    for child in children:
        tag = child.tag.split("}")[-1]
        if tag == "p":
            if child not in keep_para_elements:
                body.remove(child)
        elif tag == "tbl":
            # Lentelę paliekam jei ji yra tarp paragrafų ribų (artumo principu)
            # Patikriname, ar prieš lentelę esantis paragrafas yra mūsų ribose
            prev = child.getprevious()
            keep_table = False
            while prev is not None:
                if prev.tag == qn("w:p"):
                    if prev in keep_para_elements:
                        keep_table = True
                    break
                prev = prev.getprevious()
            # Arba po lentelės:
            if not keep_table:
                nxt = child.getnext()
                while nxt is not None:
                    if nxt.tag == qn("w:p"):
                        if nxt in keep_para_elements:
                            keep_table = True
                        break
                    nxt = nxt.getnext()
            if not keep_table:
                body.remove(child)
        elif tag == "sectPr":
            # Paliekam paskutinę sectPr, bet pašaliname kitas
            pass
        else:
            # Kitos sekcijų savybės etc - paliekame ramybėje
            pass

    # 4. Persinumeruoti antraštes ir parašus, demotinti antraštes
    for p in src_doc.paragraphs:
        lvl = _is_heading(p)
        if lvl is not None and lvl >= 1:
            text = p.text.strip()
            if HEADING_NUM_RE.match(text):
                new_text = _renumber_heading_text(text, chapter_no)
                _replace_paragraph_text(p, new_text)
            # Demotiname (Heading 1 -> Heading 2)
            new_lvl = min(lvl + 1, 9)
            try:
                _set_heading_style(p, new_lvl)
            except KeyError:
                pass
            continue

        # Parašai
        text = p.text
        m_pav = PAV_RE.match(text)
        m_lent = LENT_RE.match(text)
        if m_pav:
            fig_counter[0] += 1
            new_text = _renumber_caption(text, "pav", fig_counter[0])
            _replace_paragraph_text(p, new_text)
        elif m_lent:
            tab_counter[0] += 1
            new_text = _renumber_caption(text, "lentelė", tab_counter[0])
            _replace_paragraph_text(p, new_text)

    # 5. Įdėti į pradžią naują skyriaus antraštę
    chapter_heading_text = f"{chapter_no}. {pu_id} - {title}"

    # Sukuriame naują paragraph elementą su Heading 1
    new_h = src_doc.paragraphs[0] if src_doc.paragraphs else None
    h_p = OxmlElement("w:p")
    h_pPr = OxmlElement("w:pPr")
    h_pStyle = OxmlElement("w:pStyle")
    h_pStyle.set(qn("w:val"), "Heading1")
    h_pPr.append(h_pStyle)
    h_p.append(h_pPr)
    h_r = OxmlElement("w:r")
    h_t = OxmlElement("w:t")
    h_t.text = chapter_heading_text
    h_r.append(h_t)
    h_p.append(h_r)

    # Įstatyti naują antraštę kaip pirmą turinio elementą body'yje
    body = src_doc.element.body
    first_content = None
    for child in body:
        tag = child.tag.split("}")[-1]
        if tag in ("p", "tbl"):
            first_content = child
            break
    if first_content is not None:
        first_content.addprevious(h_p)
    else:
        # Įdedame prieš sectPr
        for child in body:
            if child.tag.split("}")[-1] == "sectPr":
                child.addprevious(h_p)
                break
        else:
            body.append(h_p)

    src_doc.save(str(out_path))


def _build_placeholder_section(chapter_no: int, pu_id: str, out_path: Path):
    """Sukuria PU2 placeholder sekcijos dokumentą."""
    doc = Document()
    _setup_document_styles(doc)

    h = doc.add_paragraph(f"{chapter_no}. {pu_id} - UŽDUOTIS NEPATEIKTA")
    h.style = doc.styles["Heading 1"]

    p = doc.add_paragraph()
    p.add_run(
        "Pastaba: dėstytoja nepateikė šios užduoties teksto. "
        "Užduotis bus įgyvendinta, kai bus gautas oficialus reikalavimų aprašymas."
    )

    p2 = doc.add_paragraph()
    p2.add_run(
        "Šis skyrius palieka rezervuotą vietą galutiniam turiniui ir yra "
        "įtrauktas į turinį dėl numeracijos vientisumo."
    )

    doc.save(str(out_path))


# ----------------------------------------------------------------------------
# Pagrindinis sujungimas
# ----------------------------------------------------------------------------

def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    fig_counter = [0]
    tab_counter = [0]

    tmpdir = Path(tempfile.mkdtemp(prefix="pua_merge_"))
    try:
        # 1) Sukurti pagrindinį dokumentą su tituliu, TOC, paveikslų ir lentelių sąrašais
        master = Document()
        _setup_document_styles(master)
        _add_title_page(master)
        _add_toc(master)
        _add_figures_list(master)
        _add_tables_list(master)

        master_path = tmpdir / "_master.docx"
        master.save(str(master_path))

        # 2) Sukurti kiekvienos sekcijos dokumentą atskirame faile
        section_paths = []
        for chapter_no, pu_id, title, src_rel in CHAPTERS:
            sec_path = tmpdir / f"sec_{chapter_no:02d}_{pu_id}.docx"
            if src_rel is None:
                _build_placeholder_section(chapter_no, pu_id, sec_path)
            else:
                src_path = PUA_DIR / src_rel
                _build_section_docx(
                    chapter_no, pu_id, title, src_path,
                    fig_counter, tab_counter, sec_path,
                )
            section_paths.append(sec_path)
            print(f"  Sukurtas sekcijos failas: {sec_path.name}")

        # 3) Sujungti su docxcompose
        print(f"\n  Sujungiama {len(section_paths)} sekcijų į pagrindinį dokumentą...")
        master = Document(str(master_path))
        composer = Composer(master)
        for sp in section_paths:
            composer.append(Document(str(sp)))

        # 4) Išsaugome galutinį dokumentą
        composer.save(str(OUT_DOCX))
        print(f"\n  Išsaugota: {OUT_DOCX}")
        print(f"  Iš viso paveikslų: {fig_counter[0]}")
        print(f"  Iš viso lentelių: {tab_counter[0]}")

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    main()
