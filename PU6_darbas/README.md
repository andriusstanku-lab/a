# PU6 Darbas

Duomenų normalizavimas - 1NF, 2NF, 3NF taikymas trims dalykinėms sritims (studentai, pardavimai, sandėliai).

**Autorius:** Andrius Vargonas
**Studijų programa:** Marketingo technologijos
**Dalykas:** Informacijos sistemos ir duomenų bazės
**Metai:** 2026, Kaunas

## Failai

| Failas | Paskirtis |
|---|---|
| `PU6_Andrius_Vargonas.docx` | Pagrindinis Word dokumentas (~28 puslapiai) |
| `generate_pu6.py` | Dokumento generatorius |
| `generate_images_pu6.py` | 9 paveikslų generatorius |
| `generate_odb_pu6.py` + `CreateDB_pu6.java` | 3 .odb failų generatorius |
| `studentai.odb` | Normalizuota studentų registracijos DB (4 lentelės) |
| `pardavimai.odb` | Normalizuota pardavimų DB (4 lentelės) |
| `sandeliai.odb` | Normalizuota sandėlių DB (4 lentelės) |
| `INSTRUKCIJA.md` | Kaip atidaryti .odb ir padaryti screenshot'us |

## Trys normalizuotos duomenų bazės

### 1. studentai.odb (3NF)
- **Students**, **Instructors**, **Courses**, **Enrollments**
- 4 lentelės, 20 įrašų

### 2. pardavimai.odb (3NF)
- **Customers**, **Products**, **Salespeople**, **Sales**
- 4 lentelės, 20 įrašų

### 3. sandeliai.odb (3NF)
- **Warehouses**, **Products**, **Suppliers**, **Deliveries**
- 4 lentelės, 20 įrašų

## Dokumento struktūra

- Antraštinis lapas, Turinys, Įvadas (su DI deklaracija)
- 1. Duomenų normalizavimo pagrindai (3 poskyriai)
- 2. Studentų registracijos normalizavimas (5 poskyriai - analizė, 1NF, 2NF, 3NF, realizacija)
- 3. Pardavimų normalizavimas (5 poskyriai)
- 4. Produktų sandėliavimo normalizavimas (5 poskyriai)
- Išvados (4 punktai)
- Literatūros sąrašas (7 šaltiniai)
- 1 priedas - SQL DDL skriptai (12 CREATE TABLE)
- 11 priedas - DI panaudojimo deklaracija

## 9 paveikslai

- 3 nenormalizuotos lentelės (raudoni dublikatai)
- 3 normalizuotos struktūros (3NF)
- 3 datasheet view paveikslai

## DI naudojimas (trumpa)

DI naudotas tik **tekstui** (aprašymai, principų paaiškinimai). **Lentelių analizę, normalizavimą, raktų nustatymą ir duomenų įvedimą atliko autorius savarankiškai** LibreOffice Base aplinkoje.

DI dalis: <15%, vieno modelio: <5% (atitinka VU SPN-54).

## Skriptų naudojimas

```bash
pip install python-docx Pillow
python generate_images_pu6.py
python generate_odb_pu6.py
python generate_pu6.py
```
