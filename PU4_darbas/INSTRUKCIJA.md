# Instrukcija: kaip naudoti `studentai.odb` LibreOffice Base aplinkoje

Ši instrukcija paaiškina, kaip atidaryti sukurtą duomenų bazės failą ir
padaryti tikras ekrano nuotraukas, kurios galės pakeisti `img4_libre_design.png`
ir `img5_libre_data.png` mockup'us PU4 darbe.

## 1. LibreOffice įdiegimas

Jei dar neturite LibreOffice savo kompiuteryje, atsisiųskite nemokamą versiją:

- **Windows / Mac / Linux:** https://www.libreoffice.org/download/

Įdiegimo metu užtikrinkite, kad būtų pažymėta **„LibreOffice Base"** komponentas
(pagal numatytuosius nustatymus jis įjungtas).

## 2. Failo atidarymas

1. Atsisiųskite `studentai.odb` iš GitHub:
   https://github.com/andriusstanku-lab/a/raw/pu4-darbas/PU4_darbas/studentai.odb
2. Dukart spustelėkite ant atsisiusto failo.
3. LibreOffice Base atsidarys su paruošta duomenų baze.
4. Jei pasirodo prisijungimo langas - vartotojo vardas: `SA`, slaptažodis tuščias.

## 3. Lentelių apžvalga

Kairėje pusėje yra **Database** panele - spustelėkite **Tables** mygtuką.
Pamatysite tris lenteles:

- `STUDENTS` - 6 įrašai
- `COURSES` - 6 įrašai
- `ENROLLMENTS` - 7 įrašai

## 4. Ekrano nuotraukos PU4 darbui

### a) Lentelės struktūros nuotrauka (paveikslo Nr. 4 pakaitalui)

1. Dešiniu klavišu spustelėkite ant `STUDENTS` lentelės.
2. Pasirinkite **„Edit..."** (lietuviškai: „Redaguoti...").
3. Atsidarys **Table Design** langas su lauko pavadinimais ir tipais.
4. Padarykite ekrano nuotrauką:
   - **Windows:** `Win + Shift + S`, pažymėkite langą, įklijuokite Paint
   - **macOS:** `Cmd + Shift + 4`, pažymėkite langą
   - **Linux:** `Print Screen` arba `gnome-screenshot`
5. Išsaugokite kaip `img4_libre_design.png`.

### b) Lentelės duomenų nuotrauka (paveikslo Nr. 5 pakaitalui)

1. Dukart spustelėkite ant `COURSES` lentelės.
2. Atsidarys **Table Data View** su 6 įrašais.
3. Padarykite ekrano nuotrauką (žr. a) punktą).
4. Išsaugokite kaip `img5_libre_data.png`.

### c) Papildomai - ryšių diagrama

1. Atidarykite **Tools - Relationships...**.
2. Pamatysite tris lenteles su nubrežtais 1:N ryšiais (Students-Enrollments, Courses-Enrollments).
3. Galite padaryti papildomą nuotrauką ir pakeisti `img1_er_diagrama.png` arba
   palikti egzistuojančią diagramą.

## 5. Failų pakeitimas dokumente

Kai turėsite tikras nuotraukas:

1. Pakeiskite mockup'o failus PU4_darbas/ aplanke:
   - `img4_libre_design.png`
   - `img5_libre_data.png`
2. Iš naujo paleiskite generavimo skriptą:
   ```
   python generate_pu4.py
   ```
3. Naujai sugeneruotame `PU4_Vargonas.docx` bus tikros nuotraukos.

## 6. CSV failų alternatyva

Jeigu `.odb` failas nepavyksta atidaryti su jūsų LibreOffice versija, naudokite
CSV failus iš `csv/` aplanko (`students.csv`, `courses.csv`, `enrollments.csv`).
LibreOffice Base juos gali importuoti per **File - New - Database - Connect to
existing database - Text**.

## 7. MS Access alternatyva

`.odb` failas neveikia Microsoft Access programoje (skirtingi formatai).
MS Access ekrano nuotraukoms (`img2`, `img3`) reikia:

- Importuoti CSV failus į naują `.accdb` duomenų bazę,
- Arba paleisti SQL DDL skriptus iš `1 PRIEDO`,
- Arba palikti esamus mockup'us (deklaruota 11 priede kaip DI sugeneruotos
  iliustracijos).

## Pastabos

- HSQLDB versija: 1.8.0.10 (būtent ši versija veikia kartu su LibreOffice Base
  embedded mode).
- Visi tekstiniai duomenys yra be lietuviškų diakritinių ženklų, kad nebūtų
  problemų dėl koduočių (UTF-8/ISO).
- Po nuotraukų darymo failų pavadinimuose nenaudokite tarpų ar lietuviškų raidžių.
