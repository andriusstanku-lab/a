# PUA - Praktinių užduočių ataskaita

Vilniaus universiteto Kauno fakultetas
Socialinių mokslų ir taikomosios informatikos institutas
Studijų programa: Marketingo technologijos
Dalykas: Informacijos sistemos ir duomenų bazės
Autorius: Andrius Vargonas
Metai: 2026, Kaunas

## Failai

- `PUA_Andrius_Vargonas.docx` - galutinė PUA ataskaita (sujungti 7 PU darbai)
- `merge_pua.py` - sujungimo skriptas (python-docx + docxcompose)
- `README.md` - šis failas

## PUA struktūra

1. Titulinis lapas (vienas, ne 7 atskiri)
2. TURINYS (Word TOC field, automatinis)
3. PAVEIKSLŲ SĄRAŠAS (TOC \c "Pav")
4. LENTELIŲ SĄRAŠAS (TOC \c "Lentelė")
5. 1. PRAKTINĖ UŽDUOTIS PU1 - originalus PU1 dokumentas
6. 2. PRAKTINĖ UŽDUOTIS PU2 - originalus PU2 dokumentas
7. 3. PRAKTINĖ UŽDUOTIS PU3 - originalus PU3 dokumentas
8. 4. PRAKTINĖ UŽDUOTIS PU4 - originalus PU4 dokumentas
9. 5. PRAKTINĖ UŽDUOTIS PU5 - originalus PU5 dokumentas
10. 6. PRAKTINĖ UŽDUOTIS PU6 - originalus PU6 dokumentas
11. 7. PRAKTINĖ UŽDUOTIS PU7 - originalus PU7 dokumentas

## Techniniai parametrai

- Lapas: A4 (210x297 mm), vertikali orientacija
- Paraštės: viršus 20 mm, apačia 20 mm, kairė 25 mm, dešinė 15 mm
- Header / Footer atstumas: po 12.5 mm
- Šriftas: Times New Roman, 12 pt
- Tarpas tarp eilučių: 1.5
- Tekstas: Justify
- Pirmosios eilutės įtrauka: 1.25 cm
- Puslapių numeracija: dešinėje footer'yje, arabiški skaitmenys be taško
- Titulinis lapas nenumeruojamas (skaičiuojamas), numeracija prasideda nuo 2 puslapio

## SVARBU - po atsisiuntimo

Atidarius dokumentą Microsoft Word arba LibreOffice Writer:

1. Paspauskite `Ctrl+A` (paryškinti viską)
2. Paspauskite `F9` (atnaujinti visus laukus)
3. Jei Word'as klausia, ar atnaujinti TOC - rinkitės "Update entire table"
4. Patikrinkite, kad TURINYS, PAVEIKSLŲ SĄRAŠAS, LENTELIŲ SĄRAŠAS užsipildė
5. Patikrinkite vizualiai, kad puslapių numeracija eina teisingai
6. Įkelkite į VMA iki **2026-05-30**

## Kaip pakartoti generavimą

```bash
pip install python-docx docxcompose
python3 merge_pua.py
```

Skriptas tikisi, kad 7 PU .docx failai yra `/projects/sandbox/a/` šaknyje.
Galite redaguoti `REPO_DIR` ir `OUT_DIR` konstantas skripto viršuje.

## Pastabos

- 7 originalūs PU dokumentai prijungti per `docxcompose.Composer` - jų vidinė
  struktūra (antraštiniai lapai, vidiniai TOC'ai, paveikslai, lentelės) palikta
  nepakeista, kaip ir buvo studento įkelta.
- Prieš kiekvieną PU pridėtas vienas Heading 1 antraštės puslapis, kad būtų
  galima automatiškai surinkti pagrindinį TURINYS.
- Caption stilius (paveikslams ir lentelėms) sukonfigūruotas 11 pt, bold,
  Times New Roman.
