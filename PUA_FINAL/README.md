# PUA - Praktinių užduočių ataskaita (v2 - Povilo modelis)

Vilniaus universiteto Kauno fakultetas
Socialinių mokslų ir taikomosios informatikos institutas
Studijų programa: Marketingo technologijos
Dalykas: Informacijos sistemos ir duomenų bazės
Autorius: Andrius Vargonas
Metai: 2026, Kaunas

## Failai

- `PUA_Andrius_Vargonas.docx` - galutinė PUA ataskaita (vientisas akademinis dokumentas)
- `merge_pua_v2.py` - naujasis sujungimo skriptas (Povilo modelis)
- `merge_pua.py` - senasis sujungimo skriptas (V1 - paprastas suklijavimas)
- `README.md` - šis failas

## Esmė - kas pakeista nuo V1

V1 modelis (`merge_pua.py`) tiesiog suklijavo 7 PU dokumentus vieną po kito,
todėl liko 7 atskiri tituliniai lapai, 7 atskiri TURINYS'ai, 7 atskiri įvadai,
7 atskiros išvados, 7 atskiri literatūros sąrašai ir 7 dubliuojantys priedai.

V2 modelis (`merge_pua_v2.py`) **automatiškai pertvarko** 7 PU dokumentus į
**vientisą akademinį** dokumentą pagal VU KnF reikalavimus ir Povilo Adomo
Bruko PUA šabloną:

| Dalis | V1 | V2 (dabar) |
|---|---|---|
| Tituliniai lapai | 8 (1 PUA + 7 PU vidiniai) | **1** |
| Vidiniai PU TOC'ai | 7 | **0** |
| Vidiniai PU įvadai | 7 atskiri | integruoti į PU skyrius |
| IŠVADOS | 7 atskiri Heading 1 | **1** bendras skyrius su PU1-PU7 poskyriais |
| LITERATŪROS SĄRAŠAS | 7 atskiri | **1** sujungtas, deduplikuotas, abėcėlės tvarka |
| PRIEDAI | "11 PRIEDAS" pasikartoja 7 kartus | **7 nuoseklūs** (1 PRIEDAS - 7 PRIEDAS) |
| Heading 1 | 32 (chaosas TOC'e) | **8** (PU1-PU7 + IŠVADOS) |
| Heading 2 | 0 (nesutvarkyta) | **34** (PU vidiniai skyriai) |
| Heading 3 | 0 | **82** (PU poskyriai) |

## PUA struktūra (V2)

1. **Titulinis lapas**
2. **TURINYS** - Word TOC field (auto, atnaujinamas F9)
3. **LENTELIŲ SĄRAŠAS** - statinis sąrašas (15 lentelių su PU prefiksu)
4. **PAVEIKSLŲ SĄRAŠAS** - statinis sąrašas (50 paveikslų su PU prefiksu)
5. **PU1.-PU7. skyriai** - kiekvienas su:
   - Heading 1 antrašte "PU{X}. {TEMA}"
   - "Užduoties kontekstas ir tikslai:" intro tekstu
   - Heading 2 sub-skyriai (1., 2., 3...)
   - Heading 3 poskyriai (1.1., 1.2., ...)
   - Lentelės, paveikslai, SQL skriptai - palikti originalūs
6. **IŠVADOS** - vienas bendras skyrius
   - Po jo Heading 2 poskyriai "PU1 išvados", "PU2 išvados", ...
7. **LITERATŪROS SĄRAŠAS** - vienas bendras (29 šaltiniai po deduplikavimo)
8. **PRIEDAI** - sąrašas + 7 priedai
   - 1 PRIEDAS - PU1 DI deklaracija + papildoma medžiaga
   - 2 PRIEDAS - PU2 DI deklaracija + papildoma medžiaga
   - ... 7 PRIEDAS - PU7 ...

## Techniniai parametrai (VU KnF)

- Lapas: A4 (210x297 mm), vertikali orientacija
- Paraštės: viršus 20 mm, apačia 20 mm, kairė 25 mm, dešinė 15 mm
- Header / Footer: po 12.5 mm
- Šriftas: Times New Roman, 12 pt
- Tarpas tarp eilučių: 1.5
- Tekstas: Justify
- Pirmosios eilutės įtrauka: 1.25 cm
- Heading 1: 14 pt, Bold, Center, didžiosiomis raidėmis, page break before
- Heading 2: 12 pt, Bold, Left, mažosiomis raidėmis (išskyrus pirmąją)
- Caption (lentelės/paveikslai): 11 pt, Bold
- Puslapių numeracija: dešinėje footer'yje, arabiški skaitmenys be taško
- Titulinis lapas nenumeruojamas (skaičiuojamas)

## SVARBU - po atsisiuntimo

Atidarius dokumentą Microsoft Word arba LibreOffice Writer:

1. Paspauskite `Ctrl+A` (paryškinti viską)
2. Paspauskite `F9` (atnaujinti visus laukus)
3. Jei Word'as klausia, ar atnaujinti TOC - rinkitės "Update entire table"
4. Patikrinkite, kad TURINYS užsipildė skyrių pavadinimais ir puslapiais
5. LENTELIŲ ir PAVEIKSLŲ sąrašai jau yra statiniai (nereikia F9)
6. Patikrinkite vizualiai, kad puslapių numeracija eina teisingai
7. Įkelkite į VMA iki **2026-05-30**

## Kaip pakartoti generavimą

```bash
pip install python-docx docxcompose
python3 merge_pua_v2.py
```

Skriptas tikisi, kad 7 PU .docx failai yra `/projects/sandbox/a/` šaknyje.
Galite redaguoti `REPO_DIR` ir `OUT_DIR` konstantas skripto viršuje.

## Pastabos

- Originalūs jūsų darbo turiniai (lentelės, paveikslai, SQL skriptai, ERD diagramos)
  **NEPAKEISTI** - tik perskirstyta dokumento struktūra.
- Vidiniai PU titulinai lapai, TOC'ai pašalinti.
- Vidiniai PU skyriai (1., 2., 3.) demotinti į Heading 2; poskyriai (1.1., 1.2.)
  į Heading 3, kad atsirastų aiški hierarchija TURINYJE.
- Visi 7 PU literatūros sąrašai sujungti į vieną; dublikatai pašalinti pagal
  pirmus 80 simbolių (case-insensitive); abėcėlės tvarka.
- Visos 7 PU DI deklaracijos perkeltos į PRIEDUS (1-7 PRIEDAS) su tęstine
  numeracija.
