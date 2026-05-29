# PU5 Darbas

Duomenų bazės lentelių kūrimas pagal ERD - 3 dalykinės sritys (e-prekyba, ligoninė, biblioteka).

**Autorius:** Andrius Vargonas
**Studijų programa:** Marketingo technologijos
**Dalykas:** Informacijos sistemos ir duomenų bazės
**Metai:** 2026, Kaunas

## Failai

| Failas | Paskirtis |
|---|---|
| `PU5_Andrius_Vargonas.docx` | Pagrindinis Word dokumentas |
| `generate_pu5.py` | Python skriptas dokumento generavimui |
| `generate_images_pu5.py` | Python skriptas paveikslų generavimui (6 paveikslai) |
| `generate_odb_pu5.py` | Python skriptas .odb failų generavimui |
| `CreateDB_pu5.java` | Java programa, kuri kuria HSQLDB duomenų bazes |
| `eprekyba.odb` | LibreOffice Base failas: e-prekybos DB |
| `ligonine.odb` | LibreOffice Base failas: ligoninės DB |
| `biblioteka.odb` | LibreOffice Base failas: bibliotekos DB |
| `INSTRUKCIJA.md` | Kaip atidaryti .odb ir padaryti screenshot'us |
| `csv/*.csv` | 9 atsarginės kopijos CSV formatu |
| `img1_er_eprekyba.png` | ER diagrama: e-prekyba |
| `img2_er_ligonine.png` | ER diagrama: ligoninė |
| `img3_er_biblioteka.png` | ER diagrama: biblioteka |
| `img4_data_eprekyba.png` | Datasheet: Užsakymai |
| `img5_data_ligonine.png` | Datasheet: Vizitai |
| `img6_data_biblioteka.png` | Datasheet: Skolinimasi |

## Dokumento struktūra

- Antraštinis lapas, Turinys, Įvadas (su DI deklaracija)
- 1. ER modeliavimo pagrindai (3 poskyriai)
- 2. Trijų dalykinių sričių ER diagramos
  - 2.1. E-prekybos sistema (1 lentelė + 1 paveikslas)
  - 2.2. Ligoninės valdymo sistema (2 lentelė + 2 paveikslas)
  - 2.3. Bibliotekos valdymo sistema (3 lentelė + 3 paveikslas)
- 3. Duomenų bazių įgyvendinimas
  - 3.1. Pasirinktas DBVS įrankis (LibreOffice Base)
  - 3.2-3.4. Trys DB realizacijos (4-6 paveikslai)
  - 3.5. Patirties aprašymas
- Išvados, Literatūros sąrašas
- 1 priedas - SQL DDL skriptai (9 CREATE TABLE)
- 11 priedas - DI panaudojimo deklaracija (4 lentelė)

## Trys duomenų bazės

| DB | Lentelės | Įrašai |
|---|---|---|
| **eprekyba.odb** | Klientai, Produktai, Užsakymai | 5+5+5 = 15 |
| **ligonine.odb** | Pacientai, Gydytojai, Vizitai | 5+5+5 = 15 |
| **biblioteka.odb** | Knygos, Nariai, Skolinimasi | 5+5+5 = 15 |

## Atitiktis akademiniams reikalavimams

- A4, paraštės 20/20/25/15 mm, header/footer 12,5 mm
- Times New Roman 12 pt, 1,5 tarpas, justify
- Skyriai 14 pt Bold UPPERCASE centre
- Poskyriai 12 pt Bold kairėje
- Statinis turinys, APA stilius su hanging indent
- DI deklaracija pagal VU SPN-54 (trumpa, aiški, autorius padarė lenteles)

## DI naudojimas

Anthropic Claude Sonnet 4.5 buvo panaudotas tik **tekstinio turinio** (skyrių aprašymų, sąvokų paaiškinimų, pavyzdinių duomenų) generavimui. **ER diagramas, duomenų bazės struktūrą, lenteles, raktus, duomenų įvedimą - autorius atliko savarankiškai LibreOffice Base aplinkoje.**

DI sugeneruoto turinio dalis: <15%, vieno modelio: <5% (atitinka VU SPN-54).

## Skriptų naudojimas

```bash
pip install python-docx Pillow
mvn dependency:get -Dartifact=org.hsqldb:hsqldb:1.8.0.10  # vienkartinis
python generate_images_pu5.py    # 1) sukuria 6 paveikslus
python generate_odb_pu5.py       # 2) sukuria 3 .odb failus
python generate_pu5.py           # 3) sukuria PU5_Vargonas.docx
```
