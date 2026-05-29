# PU7 Darbas

MS Access duomenų bazių analizė ir kūrimas - 6 duomenų bazių (DB1-DB6) išsami analizė.

**Autorius:** Andrius Vargonas
**Studijų programa:** Marketingo technologijos
**Dalykas:** Informacijos sistemos ir duomenų bazės
**Metai:** 2026, Kaunas

## Failai

| Failas | Paskirtis |
|---|---|
| `PU7_Andrius_Vargonas.docx` | Pagrindinis Word dokumentas (~26 puslapiai) |
| `generate_pu7.py` | Dokumento generatorius |
| `generate_images_pu7.py` | 9 paveikslų generatorius |
| `generate_odb_pu7.py` + `CreateDB_pu7.java` | LibreOffice Base failų generatorius |
| `DB1.odb` - `DB6.odb` | LibreOffice Base ekvivalentai (51 inventoriaus įrašas) |
| `extracted/*.csv` | Duomenys, ištraukti iš .accdb failų |
| `INSTRUKCIJA.md` | Kaip atidaryti .odb failus |

## Pastaba dėl įrankio

Užduotyje pateikti `.accdb` (Microsoft Access) failai. Kadangi neturim MS Access, praktinė darbo dalis atlikta **LibreOffice Base** aplinkoje. `.accdb` failų turinys (lentelės ir duomenys) buvo perkeltas į ekvivalenčius `.odb` failus per Python `access-parser` biblioteką ir HSQLDB duomenų bazės variklį.

## Dokumento struktūra (6 skyriai)

- 1. DB1 - Lentelių analizė (struktūra, raktai, ryšiai, operacijos)
- 2. DB2 - Užklausų kūrimas (Department, Supplier, Date)
- 3. DB3 - Formų kūrimas (frmInventory + 2 naujos)
- 4. DB4 - Užklausų modifikavimas (qryUnusedSuppliers Kanada, frmSuppliers)
- 5. DB5 - Ataskaitų kūrimas (rptDepartments, 2 naujos)
- 6. DB6 - Skaičiavimai (SUM, AVG, COUNT, MAX)

Plus: Įvadas, Išvados (7 punktai), Literatūra, 1 priedas (SQL DDL/DML), 11 priedas (DI deklaracija).

## DB struktūra (visose 5 .accdb)

| Lentelė | Įrašai | Raktas |
|---|---|---|
| tblDepartments | 18 skyrių | Department (PK) |
| tblSuppliers | 18 tiekėjų | SupplierID (PK) |
| tblInventory | 51 produktas | ProductCode (PK), Dept FK, SupplierID FK |

Plus: 8 užklausos, 2 formos, 3 ataskaitos (DB4-DB6).

## Skripto naudojimas

```bash
pip install python-docx Pillow access-parser
mvn dependency:get -Dartifact=org.hsqldb:hsqldb:1.8.0.10
javac -cp /root/.m2/repository/org/hsqldb/hsqldb/1.8.0.10/hsqldb-1.8.0.10.jar CreateDB_pu7.java
java -cp .:/root/.m2/repository/org/hsqldb/hsqldb/1.8.0.10/hsqldb-1.8.0.10.jar CreateDB_pu7 . extracted
python3 generate_odb_pu7.py
python3 generate_images_pu7.py
python3 generate_pu7.py
```
