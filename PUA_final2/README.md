# PUA - Praktinių užduočių ataskaita (galutinė versija)

**Autorius:** Andrius Vargonas
**Fakultetas:** Vilniaus universitetas, Kauno fakultetas
**Dalykas:** Duomenų bazės
**Data:** 2026-05

## Apie šį aplanką

Šis aplankas turi sujungtą galutinę PUA, sukurtą iš vartotojo įkeltų 7 originalių .docx
failų (`pu1-darbas` šaka, repo šaknyje):

- `PU1_Andrius_Vargonas.docx`
- `PU2_Andrius_Vargonas.docx`
- `PU3_Andrius_Vargonas1.docx`
- `PU4_Andrius_Vargonas.docx`
- `PU5_Andrius_Vargonas.docx`
- `PU6_Andrius_Vargonas.docx`
- `PU7_Andrius_Vargonas.docx`

## Failai

- **`PUA_Andrius_Vargonas.docx`** - sujungtas pristatymui rytoj (2026-05-30)
  skirtas dokumentas (~1.35 MB).
- `merge_pua_user.py` - Python skriptas, kuris atliko sujungimą
  (`python-docx` + `docxcompose`).

## Sujungimo struktūra

1. **Antraštinis lapas** (VU Kauno fakultetas, "PRAKTINIŲ UŽDUOČIŲ ATASKAITA (PUA)",
   autorius, Kaunas 2026)
2. **TURINYS** - automatinis Word `TOC` laukas (atnaujinti su F9)
3. **PAVEIKSLŲ SĄRAŠAS** - automatinis `TOC \c "Pav"` laukas
4. **LENTELIŲ SĄRAŠAS** - automatinis `TOC \c "Lentelė"` laukas
5. Toliau seka 7 skyriai (po Heading 1 skirtuką su pavadinimu),
   kiekviename - originalus PU dokumento turinys:
   - 1. PU1 - Informacinės sistemos ir duomenų bazės
   - 2. PU2 - Reliacinis duomenų modelis
   - 3. PU3 - Microsoft Access aplinka
   - 4. PU4 - Duomenų bazės kūrimas Java priemonėmis
   - 5. PU5 - Duomenų bazių pavyzdžiai
   - 6. PU6 - Normalizavimas iki 3NF
   - 7. PU7 - Vaizdai (Views) ir užklausos

## SVARBU - prieš spausdinant

Atidarykite Word, **paspauskite Ctrl + A**, tada **F9**, kad atsinaujintų:
- TURINYS
- PAVEIKSLŲ SĄRAŠAS
- LENTELIŲ SĄRAŠAS

Originaliuose PU dokumentuose paliktos jų pačių titulinės/turinio dalys -
tai daro skyrių pradžias aiškiomis dėstytojai. Jei norėsite, galite jas
pašalinti rankiniu būdu.

## Pastaba dėl DB2

Vartotojas pranešė, kad **DB2.accdb į `pu1-darbas` šaką nebuvo įkeltas**.
Repo šaknyje yra: DB1, DB3, DB4, DB5, DB6 (.accdb). PUA dokumentas remiasi
tik įkeltais .docx failais, todėl tai PUA struktūros nepaveikė.
