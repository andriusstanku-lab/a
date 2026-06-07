# PU4 PRAKTINĖS UŽDUOTIES TURINYS

**Darbo pavadinimas:** DBVS APLINKA, LENTELIŲ KŪRIMAS
**Autorius:** Andrius Vargonas
**Studijų programa:** Marketingo technologijų studijų programa
**Dalykas:** Informacijos sistemos ir duomenų bazės
**Fakultetas:** VU Kauno fakultetas
**Institutas:** Socialinių mokslų ir taikomosios informatikos institutas
**Metai:** 2026
**Miestas:** Kaunas

---

## TURINYS

```
ĮVADAS .......................................................... 3
1. DBVS ĮRANKIŲ APLINKOS APŽVALGA ............................... 4
   1.1. Microsoft Access ........................................ 4
   1.2. OpenOffice Base ......................................... 5
   1.3. LibreOffice Base ........................................ 6
2. PAVYZDINĖS DUOMENŲ BAZĖS PROJEKTAVIMAS ........................ 7
   2.1. Dalykinė sritis ir lentelių struktūra ................... 7
   2.2. Pirminių ir išorinių raktų schema ....................... 8
   2.3. Duomenų tipai ir apribojimai ............................ 9
3. LENTELIŲ SUKŪRIMAS PASIRINKTUOSE ĮRANKIUOSE .................. 10
   3.1. Realizacija Microsoft Access aplinkoje ................. 10
   3.2. Realizacija LibreOffice Base aplinkoje ................. 12
   3.3. Patirties palyginimas .................................. 14
IŠVADOS ........................................................ 15
LITERATŪROS SĄRAŠAS ............................................ 16
1 PRIEDAS. SQL DDL skriptai .................................... 17
2 PRIEDAS. Pavyzdiniai duomenys ................................ 18
11 PRIEDAS. DI panaudojimo deklaracija ......................... 19
12 PRIEDAS. DI užklausos ....................................... 20
```

---

## ĮVADAS

Duomenų bazių valdymo sistemos (toliau - DBVS) yra esminė šiuolaikinių informacinių sistemų sudedamoji dalis. Marketingo technologijų studijų kontekste praktinis darbas su skirtingomis DBVS aplinkomis padeda geriau suprasti, kaip duomenys saugomi, struktūrizuojami ir naudojami verslo sprendimams priimti. Reliacinės duomenų bazės, paremtos Codd (1970) suformuluotu modeliu, ir šiandien lieka dažniausiai naudojamu duomenų saugojimo principu Lietuvos ir užsienio įmonėse.

Šio darbo tikslas - praktiškai išbandyti tris darbalaukio DBVS kūrimo įrankius (Microsoft Access, OpenOffice Base, LibreOffice Base) ir naudojant du iš jų sukurti pavyzdinę universiteto duomenų bazę su trimis tarpusavyje susijusiomis lentelėmis.

**Darbo uždaviniai:**

1. Apžvelgti tris DBVS kūrimo įrankių aplinkas - jų vartotojo sąsajas, funkcionalumą ir lentelių kūrimo veiksmus.
2. Suprojektuoti pavyzdinę universiteto duomenų bazę su trimis tarpusavyje susijusiomis lentelėmis (Students, Courses, Enrollments).
3. Praktiškai sukurti lenteles dvejose pasirinktose aplinkose (Microsoft Access ir LibreOffice Base), užtikrinant tinkamą pirminių ir išorinių raktų naudojimą.
4. Užpildyti lenteles pavyzdiniais duomenimis ir trumpai aprašyti patirtį dirbant su abiem įrankiais.

Darbo metodai - mokslinės literatūros ir oficialios dokumentacijos analizė, dalykinės srities modeliavimas, praktinis lentelių kūrimas DBVS aplinkose, lyginamoji refleksija.

Darbą sudaro trys skyriai. Pirmajame pateikiama trijų DBVS įrankių aplinkų apžvalga. Antrajame projektuojama pavyzdinė duomenų bazė ir aprašomos jos struktūros savybės. Trečiajame demonstruojama lentelių realizacija pasirinktose aplinkose ir lyginama vartotojo patirtis.

**DI deklaracija:** Rengiant šį darbą buvo naudotasi Anthropic Claude (Sonnet 4.5) modeliu kaip pagalbine priemone struktūros patikrinimui, kalbinio stiliaus tobulinimui ir paveikslų vizualiniam apipavidalinimui. Detalus naudojimo aprašymas - 11 priede.

---

## 1. DBVS ĮRANKIŲ APLINKOS APŽVALGA

### 1.1. Microsoft Access

Microsoft Access yra Microsoft 365 paketo dalis - patentuota darbalaukio DBVS, pirmąkart išleista 1992 m. Įrankis derina duomenų bazės variklį (Jet/ACE) su grafine vartotojo sąsaja. Pagrindinis privalumas - integracija su Microsoft 365 produktais. Trūkumai - mokama, tik Windows, riba 2 GB.

### 1.2. OpenOffice Base

Apache OpenOffice Base - atvirojo kodo nemokama DBVS, veikianti Windows, macOS ir Linux. Palaiko HSQLDB, Firebird ir išorinių DB jungtis. Trūkumas - lėtesnė plėtra, daugelis vartotojų perėjo prie LibreOffice.

### 1.3. LibreOffice Base

LibreOffice Base - aktyviausiai vystoma OpenOffice atšaka. Lentelių kūrimui - vedlys, Design View ir SQL View režimai. Lietuvoje plačiai naudojamas viešojo sektoriaus įstaigose ir mokyklose.

---

## 2. PAVYZDINĖS DUOMENŲ BAZĖS PROJEKTAVIMAS

### 2.1. Dalykinė sritis ir lentelių struktūra

Universiteto studentų ir kursų valdymo sistema. Trys lentelės:
- **Students** - studentų duomenys
- **Courses** - dėstomų kursų sąrašas
- **Enrollments** - registracijos (siejanti lentelė)

### 2.2. Pirminių ir išorinių raktų schema

```
Students                Enrollments               Courses
--------                -----------               -------
PK StudentID    1 ---> N FK StudentID  N <--- 1 PK CourseID
   FirstName             PK EnrollmentID            CourseName
   LastName              FK CourseID                Credits
   Email                 Semester                   Department
   Major                 Grade
   EnrollmentYear
```

### 2.3. Duomenų tipai ir apribojimai

| Lentelė | Laukas | Tipas | Apribojimas |
|---------|--------|-------|-------------|
| Students | StudentID | INTEGER | PK, AUTO |
| Students | FirstName | VARCHAR(50) | NOT NULL |
| Students | LastName | VARCHAR(50) | NOT NULL |
| Students | Email | VARCHAR(100) | UNIQUE |
| Courses | CourseID | INTEGER | PK, AUTO |
| Courses | CourseName | VARCHAR(100) | NOT NULL |
| Enrollments | EnrollmentID | INTEGER | PK, AUTO |
| Enrollments | StudentID | INTEGER | FK -> Students |
| Enrollments | CourseID | INTEGER | FK -> Courses |
| Enrollments | Grade | DECIMAL(3,1) | |

---

## 3. LENTELIŲ SUKŪRIMAS PASIRINKTUOSE ĮRANKIUOSE

### 3.1. Realizacija Microsoft Access aplinkoje

Lentelės kuriamos per Design View režimą. StudentID nustatytas kaip AutoNumber - automatiškai numeruojamas pirminis raktas. Field Properties panelėje - Indexed, Required, Caption ir kt. savybės. Po to per Database Tools - Relationships nustatomi išoriniai raktai. Rezultatas - 6 įrašai Students lentelėje.

### 3.2. Realizacija LibreOffice Base aplinkoje

Sukurta studentai.odb duomenų bazė su HSQLDB varikliu. Lentelės kuriamos per Create Table in Design View. AutoValue savybė atitinka Access AutoNumber. Per Tools - Relationships nustatomi ryšiai. Rezultatas - 6 įrašai Courses lentelėje su realiais marketingo studijų programos kursais.

### 3.3. Patirties palyginimas (50-100 žodžių)

Praktinis darbas su abiem įrankiais atskleidė, kad jie turi panašų funkcionalumą lentelių kūrimo srityje, tačiau skiriasi vartotojo patogumu. Microsoft Access pasižymi modernesne sąsaja su Ribbon stiliaus juostele ir glaudesne integracija su Microsoft 365. LibreOffice Base savo ruožtu siūlo platesnį platformų palaikymą ir nemokamą licenciją. Sąsaja paprastesnė, tačiau visos esminės funkcijos pasiekiamos. SQL View galimybė LibreOffice Base privalumas pažangesniems vartotojams. Abi aplinkos tinkamos mokomajai ir smulkiojo verslo praktikai.

---

## IŠVADOS

1. Apžvelgus tris DBVS įrankius, nustatyta, kad Microsoft Access yra mokamas Windows produktas, OpenOffice Base ir LibreOffice Base - atvirojo kodo kelių platformų alternatyvos. LibreOffice Base aktyviau vystoma.

2. Suprojektuota pavyzdinė universiteto DB su trimis tarpusavyje susijusiomis lentelėmis atspindi klasikinį reliacinių duomenų bazių modelį. Pirminiai ir išoriniai raktai užtikrina duomenų vientisumą.

3. Praktinis lentelių sukūrimas Access ir LibreOffice Base parodė, kad abu įrankiai pateikia panašų Design View režimą. Skiriasi tik vartotojo sąsajos stilius (Access AutoNumber = LibreOffice AutoValue).

4. Pasirinkimas tarp Access ir LibreOffice Base priklauso nuo poreikių: Access - kai svarbi M365 integracija, LibreOffice Base - kai reikia kelių platformų ir nemokamos licencijos.

---

## LITERATŪROS SĄRAŠAS

Apache Software Foundation. (2026). *Apache OpenOffice Base*. Prieiga per internetą: https://www.openoffice.org/product/base.html

Codd, E. F. (1970). A Relational Model of Data for Large Shared Data Banks. *Communications of the ACM*, 13(6), 377-387.

Connolly, T. ir Begg, C. (2015). *Database Systems: A Practical Approach to Design, Implementation, and Management* (6th ed.). Boston: Pearson.

Elmasri, R. ir Navathe, S. B. (2016). *Fundamentals of Database Systems* (7th ed.). Boston: Pearson.

Microsoft. (2026). *Microsoft Access dokumentacija*. Prieiga per internetą: https://support.microsoft.com/lt-lt/access

The Document Foundation. (2026). *LibreOffice Base Handbook*. Prieiga per internetą: https://documentation.libreoffice.org

TutorialsPoint. (2026). *MS Access Tutorial*. Prieiga per internetą: https://www.tutorialspoint.com/ms_access/index.htm

Vilniaus universitetas. (2024). *Dirbtinio intelekto naudojimo gairės* (Nr. SPN-54). Vilnius: VU.

---

## 1 PRIEDAS. SQL DDL skriptai

```sql
CREATE TABLE Students (
    StudentID       INTEGER         GENERATED BY DEFAULT AS IDENTITY,
    FirstName       VARCHAR(50)     NOT NULL,
    LastName        VARCHAR(50)     NOT NULL,
    Email           VARCHAR(100)    UNIQUE,
    Major           VARCHAR(50),
    EnrollmentYear  INTEGER,
    PRIMARY KEY (StudentID)
);

CREATE TABLE Courses (
    CourseID        INTEGER         GENERATED BY DEFAULT AS IDENTITY,
    CourseName      VARCHAR(100)    NOT NULL,
    Credits         INTEGER,
    Department      VARCHAR(50),
    PRIMARY KEY (CourseID)
);

CREATE TABLE Enrollments (
    EnrollmentID    INTEGER         GENERATED BY DEFAULT AS IDENTITY,
    StudentID       INTEGER         NOT NULL,
    CourseID        INTEGER         NOT NULL,
    Semester        VARCHAR(20),
    Grade           DECIMAL(3,1),
    PRIMARY KEY (EnrollmentID),
    FOREIGN KEY (StudentID) REFERENCES Students(StudentID),
    FOREIGN KEY (CourseID)  REFERENCES Courses(CourseID)
);
```

---

## 11 PRIEDAS. DI panaudojimo deklaracija

DI modelis: Anthropic Claude Sonnet 4.5
Naudojimo data: 2026 m. gegužės mėn.
Naudojimo tikslas: struktūra, kalba, paveikslai
DI sugeneruoto turinio dalis: <15%
Vieno modelio: <5% (atitinka VU SPN-54)
Modifikavimo apimtis: ~70-80%
