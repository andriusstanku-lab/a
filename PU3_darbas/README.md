# PU3 Darbas

Duomenų bazių kūrimo įrankių analizė ir palyginimas.

**Autorius:** Andrius Vargonas
**Studijų programa:** Marketingo technologijos
**Dalykas:** Informacijos sistemos ir duomenų bazės
**Metai:** 2026, Kaunas

## Failai

- `PU3_darbas/PU3_Andrius_Vargonas.docx` - Pagrindinis darbo Word dokumentas
- `PU3_darbas/generate_pu3.py` - Python skriptas, kuriuo sugeneruotas dokumentas (galima paleisti iš naujo: `pip install python-docx && python generate_pu3.py`)
- `PU3_darbas/PU3_turinys.md` - Pilnas darbo tekstas Markdown formatu (atsarginė kopija ir referencinis tekstas)

## Dokumento struktūra

- Antraštinis lapas (be numerio, skaičiuojamas)
- TURINYS (statinis su tab-leader taškeliais ir puslapių numeriais)
- ĮVADAS (su DI naudojimo deklaracijos pastraipa)
- 1. Duomenų bazių kūrimo įrankių teorinė apžvalga
  - 1.1. Duomenų bazės samprata ir pagrindiniai tipai
  - 1.2. Duomenų bazių kūrimo įrankių klasifikacija
  - 1.3. Palyginimo kriterijų atranka ir pagrindimas
- 2. Dešimties duomenų bazių kūrimo įrankių apžvalga
  - 2.1. Tradiciniai reliaciniai įrankiai (MS Access, LibreOffice Base, OpenOffice Base, MS SQL SSMS)
  - 2.2. Šiuolaikinės debesijos ir žemo kodo platformos (Airtable, Notion, Budibase)
  - 2.3. NoSQL ir atvirojo kodo įrankiai (MongoDB Compass, DBeaver, phpMyAdmin)
- 3. Lyginamoji analizė ir rezultatai
  - 3.1. Lyginamoji lentelė pagal aštuonis kriterijus (1 lentelė)
  - 3.2. Pagrindiniai skirtumai ir panašumai
  - 3.3. Įrankių tinkamumas skirtingiems naudojimo atvejams (2 lentelė)
- IŠVADOS (4 numeruoti punktai)
- LITERATŪROS SĄRAŠAS (APA stilius, hanging indent)
- 1 PRIEDAS - Detali dešimties įrankių charakteristika
- 11 PRIEDAS - Dirbtinio intelekto panaudojimo deklaracija
- 12 PRIEDAS - DI užklausos ir gauti atsakymai

## Analizuoti įrankiai (10 vnt.)

| # | Įrankis | Tipas | Licencija |
|---|---|---|---|
| 1 | MS Access | Reliacinis | Patentuota |
| 2 | LibreOffice Base | Reliacinis | LGPLv3 (atvirojo kodo) |
| 3 | OpenOffice Base | Reliacinis | Apache 2.0 (atvirojo kodo) |
| 4 | MS SQL SSMS | Reliacinis (imones) | Patentuota (nemokama) |
| 5 | Airtable | Debesijos / Hibridinis | Patentuota (SaaS) |
| 6 | Notion Databases | Dokumentu / Ziniu valdymas | Patentuota (SaaS) |
| 7 | Budibase | Zemo kodo platforma | GPLv3 (atvirojo kodo) |
| 8 | MongoDB Compass | NoSQL | SSPL |
| 9 | DBeaver | Universalus klientas | Apache 2.0 (atvirojo kodo) |
| 10 | phpMyAdmin | Ziniatinklio | GPLv2 (atvirojo kodo) |

## Palyginimo kriterijai (8 vnt.)

1. Vidinė duomenų bazė (Internal Database Support)
2. NoSQL jungtys (NoSQL Connectors)
3. REST API palaikymas
4. Programėlių kūrimo įrankiai (App Builder)
5. Darbo eigos automatizavimas (Workflow Automations)
6. Debesijos platformos prieinamumas (Cloud Platform)
7. Savarankiškas diegimas (Self-Hosting)
8. Atvirojo kodo prieinamumas (Open-Source)

## Atitiktis akademiniams reikalavimams

Dokumentas paruoštas pagal **VU Kauno fakulteto Informatikos inžinerijos krypties akademinių rašto darbų metodinius nurodymus**:

- A4, paraštės 20 / 20 / 25 / 15 mm (viršus / apačia / kairė / dešinė)
- Header ir Footer po 12,5 mm
- Times New Roman 12 pt, 1,5 tarpas tarp eilučių, justify lygiavimas
- Pirmosios eilutės įtrauka 1,25 cm
- Skyriai 14 pt Bold UPPERCASE centre, naujame lape
- Poskyriai 12 pt Bold mažosiomis kairėje
- Puslapio numeravimas dešinėje footer'yje (antraštinis lapas be numerio, kiti puslapiai numeruoti nuo 2)
- Lentelės pagal akademinius reikalavimus (numeris virš dešinėje, pavadinimas centre 11 pt Bold, šaltinis po lentele 9 pt)
- Tekste pateikiamos nuorodos į visas lenteles
- APA stiliaus literatūros sąrašas su 1,25 cm hanging indent

## Dirbtinio intelekto naudojimas

Rengiant darbą buvo naudotas Anthropic Claude (Sonnet 4.5) kaip pagalbinė priemonė struktūros patikrinimui, kalbos taisymui ir formuluočių alternatyvoms. Visi reikalingi DI naudojimo aprašai pateikti darbe:

- Trumpa pastraipa įvade
- 11 PRIEDAS - pilna DI panaudojimo deklaracija su naudojimo apimties suvestinės lentele
- 12 PRIEDAS - užklausos (prompt'ai) ir atsakymų aprašymai

DI sugeneruoto turinio dalis darbe neviršija 15 procentų, vieno modelio - mažiau nei 5 procentų (atitinka VU SPN-54 gairių reikalavimus).

## Skripto naudojimas

```bash
pip install python-docx
python generate_pu3.py
# Sugeneruojamas: PU3_Vargonas.docx
```
