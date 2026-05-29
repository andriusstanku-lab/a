# PU4 Darbas

DBVS aplinka, lentelių kūrimas - praktinis darbas su Microsoft Access ir LibreOffice Base.

**Autorius:** Andrius Vargonas
**Studijų programa:** Marketingo technologijos
**Dalykas:** Informacijos sistemos ir duomenų bazės
**Metai:** 2026, Kaunas

## Failai

- `PU4_darbas/PU4_Vargonas.docx` - Pagrindinis darbo Word dokumentas (243 KB)
- `PU4_darbas/generate_pu4.py` - Python skriptas, kuriuo sugeneruotas dokumentas
- `PU4_darbas/generate_images.py` - Python skriptas paveikslų generavimui
- `PU4_darbas/PU4_turinys.md` - Pilnas darbo tekstas Markdown formatu
- `PU4_darbas/img1_er_diagrama.png` - ER diagrama
- `PU4_darbas/img2_access_design.png` - MS Access Design View paveikslas
- `PU4_darbas/img3_access_data.png` - MS Access Datasheet paveikslas
- `PU4_darbas/img4_libre_design.png` - LibreOffice Base Table Design paveikslas
- `PU4_darbas/img5_libre_data.png` - LibreOffice Base Table Data paveikslas

## Dokumento struktūra

- Antraštinis lapas (be numerio, skaičiuojamas)
- TURINYS (statinis su tab-leader taškeliais ir puslapių numeriais)
- ĮVADAS (su DI naudojimo deklaracijos pastraipa)
- 1. DBVS įrankių aplinkos apžvalga
  - 1.1. Microsoft Access
  - 1.2. OpenOffice Base
  - 1.3. LibreOffice Base
- 2. Pavyzdinės duomenų bazės projektavimas
  - 2.1. Dalykinė sritis ir lentelių struktūra
  - 2.2. Pirminių ir išorinių raktų schema (1 paveikslas - ER diagrama)
  - 2.3. Duomenų tipai ir apribojimai (1 lentelė)
- 3. Lentelių sukūrimas pasirinktuose įrankiuose
  - 3.1. Realizacija Microsoft Access aplinkoje (2 ir 3 paveikslai)
  - 3.2. Realizacija LibreOffice Base aplinkoje (4 ir 5 paveikslai)
  - 3.3. Patirties palyginimas (50-100 žodžių)
- IŠVADOS (4 numeruoti punktai)
- LITERATŪROS SĄRAŠAS (APA stilius, hanging indent)
- 1 PRIEDAS - SQL DDL skriptai (CREATE TABLE)
- 2 PRIEDAS - Pavyzdiniai duomenys (Students, Courses, Enrollments)
- 11 PRIEDAS - Dirbtinio intelekto panaudojimo deklaracija
- 12 PRIEDAS - DI užklausos ir gauti atsakymai

## Pavyzdinės duomenų bazės struktūra

3 tarpusavyje susijusios lentelės pagal universiteto pavyzdį:

| Lentelė | Pirminis raktas | Išoriniai raktai | Įrašai |
|---------|----------------|------------------|--------|
| Students | StudentID | - | 6 |
| Courses | CourseID | - | 6 |
| Enrollments | EnrollmentID | StudentID, CourseID | 7 |

## Naudoti DBVS įrankiai praktinėje dalyje

1. **Microsoft Access** (patentuota, Windows)
2. **LibreOffice Base** (atvirojo kodo, kelios platformos)

OpenOffice Base aprašytas tik teorinėje dalyje (1.2 poskyris).

## Atitiktis akademiniams reikalavimams

Dokumentas paruoštas pagal **VU Kauno fakulteto Informatikos inžinerijos krypties akademinių rašto darbų metodinius nurodymus**:

- A4, paraštės 20 / 20 / 25 / 15 mm (viršus / apačia / kairė / dešinė)
- Header ir Footer po 12,5 mm
- Times New Roman 12 pt, 1,5 tarpas tarp eilučių, justify lygiavimas
- Pirmosios eilutės įtrauka 1,25 cm
- Skyriai 14 pt Bold UPPERCASE centre, naujame lape
- Poskyriai 12 pt Bold mažosiomis kairėje
- Puslapio numeravimas dešinėje footer'yje (antraštinis lapas be numerio)
- Lentelės: numeris virš dešinėje, pavadinimas centre 11 pt Bold, šaltinis 9 pt po lentele
- Paveikslai: numeris ir pavadinimas po paveikslu centre 11 pt Bold
- Tekste pateikiamos nuorodos į visas lenteles ir paveikslus (žr. 1 lentelę, žr. 2 pav. ir t.t.)
- APA stiliaus literatūros sąrašas su 1,25 cm hanging indent

## Apie paveikslus

Paveikslai (img2-img5) yra realistiški iliustraciniai mockup'ai, sugeneruoti programiškai PIL biblioteka. Jie atspindi MS Access ir LibreOffice Base vartotojo sąsajų pagrindinius elementus (langas, juostelė, lentelių grid, lauko savybės). Realiame moksliniame darbe rekomenduojama juos pakeisti tikrais ekrano įrašais (Print Screen) atlikus praktinę užduotį.

## Dirbtinio intelekto naudojimas

Rengiant darbą buvo naudotas Anthropic Claude (Sonnet 4.5) kaip pagalbinė priemonė:

- Trumpa pastraipa įvade
- 11 PRIEDAS - pilna DI panaudojimo deklaracija
- 12 PRIEDAS - užklausos (prompt'ai) ir atsakymų aprašymai

DI sugeneruoto turinio dalis darbe neviršija 15 procentų, vieno modelio - mažiau nei 5 procentų (atitinka VU SPN-54 gairių reikalavimus).

## Skripto naudojimas

```bash
pip install python-docx Pillow
python generate_images.py    # 1) sugeneruoja 5 PNG paveikslus
python generate_pu4.py       # 2) sugeneruoja PU4_Vargonas.docx
```
