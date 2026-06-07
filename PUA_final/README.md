# PUA - Praktinių užduočių ataskaita

**Autorius:** Andrius Vargonas
**Fakultetas:** Vilniaus universitetas, Kauno fakultetas
**Studijų programa:** Programų sistemos
**Dalykas:** Duomenų bazės
**Pristatymo data:** 2026-05-30

## Aprašymas

Šis aplankas talpina visas šešių Praktinių užduočių (PU1, PU3-PU7) sujungtą
ataskaitą `PUA_Andrius_Vargonas.docx`, kuri buvo sugeneruota automatiškai
iš atskirų PU šakų `pu1-darbas`, `pu3-darbas` ... `pu7-darbas` repozitorijos
`andriusstanku-lab/a`.

## Aplanko struktūra

```
PUA_final/
├── PUA_Andrius_Vargonas.docx   - Galutinė sujungta ataskaita
├── README.md                   - Šis failas
├── merge_pua.py                - Sujungimo Python skriptas
├── attachments/                - Visų DBVS bazių .odb failai
│   ├── pu4_studentai.odb
│   ├── pu5_eprekyba.odb, pu5_ligonine.odb, pu5_biblioteka.odb
│   ├── pu6_studentai.odb, pu6_pardavimai.odb, pu6_sandeliai.odb
│   ├── pu7_DB1.odb ... pu7_DB6.odb
│   └── pu7_DB2_self_generated.odb (DB2 originalo dėstytoja nepateikė)
├── images/                     - Visi paveikslai sunumeruoti nuosekliai
│   └── 01_pav.png ... 40_pav.png
├── scripts/                    - Java skriptai .odb failams kurti
└── csv/                        - Žaliavinis duomenų rinkinys lentelėms
```

## Sujungimo loginis modelis

Sujungtos ataskaitos struktūra:

1. Vienas titulinis lapas (KFAK formato)
2. Automatinis turinys (Word `TOC \o "1-3"` field)
3. Automatinis paveikslų sąrašas (Word `TOC \c "pav"` field)
4. Automatinis lentelių sąrašas (Word `TOC \c "lentelė"` field)
5. 7 skyriai:
   - **1. PU1 - INFORMACINIŲ SISTEMŲ IR DUOMENŲ BAZIŲ SAMPRATA**
   - **2. PU2 - UŽDUOTIS NEPATEIKTA** (placeholder, nes dėstytoja nepateikė užduoties teksto)
   - **3. PU3 - DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ ANALIZĖ IR PALYGINIMAS**
   - **4. PU4 - DBVS APLINKA, LENTELIŲ KŪRIMAS**
   - **5. PU5 - DUOMENŲ BAZĖS LENTELIŲ KŪRIMAS PAGAL ERD**
   - **6. PU6 - DUOMENŲ NORMALIZAVIMAS**
   - **7. PU7 - MS ACCESS DUOMENŲ BAZIŲ ANALIZĖ IR KŪRIMAS**

## Persinumeravimas

- Kiekvieno PU vidiniai poskyriai persinumeruoti pagal naują skyriaus numerį.
  Pvz., PU3 turėjo "1.1. Duomenų bazės samprata" - po sujungimo tampa
  "3.1.1. Duomenų bazės samprata".
- Paveikslai persinumeruoti nuosekliai per visą dokumentą: 1 pav. - 32 pav.
- Lentelės persinumeruotos nuosekliai: 1 lentelė - 8 lentelė.

## TOC laukų atnaujinimas

Atidarius dokumentą Microsoft Word'e:

1. `Ctrl+A` (pažymėti viską)
2. `F9` (atnaujinti visus laukus)
3. Pasirinkti "Update entire table" jei klausia

LibreOffice Writer'yje: `Tools > Update > Update All`.

## Bendros statistikos

- Iš viso paveikslų: 32
- Iš viso lentelių: 8
- Skyrių: 7 (iš jų 1 placeholder)
- Poskyrių: 23+ (Heading 2)
- Dokumento dydis: ~1.05 MB

## DI panaudojimo deklaracija

Šio darbo kūrime buvo naudoti dirbtinio intelekto įrankiai (Claude AI) -
turinio struktūrizavimui, lyginamosioms analizėms, kalbos formulavimui ir
techninių iliustracijų generavimui. Visi rezultatai buvo autoriaus peržiūrėti
ir patikslinti.
