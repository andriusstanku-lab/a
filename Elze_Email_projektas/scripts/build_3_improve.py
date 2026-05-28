"""3 etapas — IMPROVE: Konkretūs pasiūlymai (5 psl. + starter pack).

Klausimas: Konkrečiai — ką Elzė turėtų daryti? Kokia yra implementacijos metodika?
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _docx_helper import (
    new_document, add_title, add_h1, add_h2, add_para, add_bullet,
    add_numbered, add_static_toc, add_table_simple, add_quote_box,
    add_pagebreak, add_header_block,
)

doc = new_document()

add_header_block(
    doc,
    university='VILNIAUS UNIVERSITETAS',
    faculty='Ekonomikos ir verslo administravimo fakultetas',
    course='E. pardavimų grupinis projektas',
    project_title='3 etapas — IMPROVE\nKonkretūs pasiūlymai ir starter pack',
    channel='Email / community building',
    group='[Grupės narių vardai, pavardės]',
    date='2026 m. gegužė',
)

add_static_toc(doc, [
    ('1. Trijų laiko horizontų logika', 2),
    ('2. Pasiūlymas Nr. 1 — Trumpas (1 mėn.): „Kanalo įjungimas"', 2),
    ('3. Pasiūlymas Nr. 2 — Vidutinis (3 mėn.): „Reguliaraus ritmo įsitvirtinimas"', 4),
    ('4. Pasiūlymas Nr. 3 — Ilgas (12 mėn.): „Karjeros bazė"', 5),
    ('5. Pasiūlymų sumavimas — kaštai, metrikos, prioritetas', 6),
    ('6. Starter pack — kas paruošta pirmadieniui', 7),
])

add_pagebreak(doc)

# ============================================================
# 1. LOGIKA
# ============================================================
add_h1(doc, '1. Trijų laiko horizontų logika')

add_para(doc,
    'Vietoje vieno didelio sprendimo siūlome tris atskirus, viena nuo kitos '
    'priklausančius pasiūlymus, paskirstytus per skirtingus laiko horizontus. Logika '
    'yra paprasta: 1 mėn. pasiūlymas turi pasirodyti bent kažkas matomo (kanalas '
    'fiziškai egzistuoja); 3 mėn. pasiūlymas — kanalas ima dirbti reguliariai; '
    '12 mėn. pasiūlymas — kanalas tampa Elzės karjeros baze, ne tik šios konkrečios '
    'knygos rinkodara.')

add_para(doc,
    'Visi trys pasiūlymai paremti viena prielaida: Elzė yra viena, ji nėra rinkodaros '
    'komanda, ir bet koks darbas, kurio savaitėje ji negali aprėpti, yra nereikalingas. '
    'Todėl kiekvieną pasiūlymą įvertiname laiku per savaitę, ne svajonėmis.')

# ============================================================
# 2. PASIŪLYMAS Nr. 1 — TRUMPAS
# ============================================================
add_h1(doc, '2. Pasiūlymas Nr. 1 — Trumpas (1 mėn.): „Kanalo įjungimas"')

add_h2(doc, '2.1. Tikslas')

add_para(doc,
    'Kad per 4 savaites Elzės svetainėje atsirastų veikianti subscribe forma, kad '
    'pasirašiusieji gautų savitvarkę 3 laiškų welcome seką su nemokama knygos '
    'ištrauka, ir kad iki mėnesio pabaigos prenumeratorių skaičius pasiektų '
    '40–70 (tikėtinas scenarijus iš ANALYZE).')

add_h2(doc, '2.2. Konkrečios užduotys (4 savaitės)')

add_table_simple(
    doc,
    headers=['Sav.', 'Užduotis', 'Rezultatas',
             'Laikas'],
    rows=[
        ['1',
         'Susikurti MailerLite paskyrą; sukurti pirmą prenumeratorių sąrašą; '
         'patvirtinti SPF/DKIM įrašus per esamą hostingą.',
         'Veikianti, oficiali siuntinė pristatomumui užtikrinti.',
         '~1,5 val.'],
        ['1',
         'Paruošti vieną PDF — pirmą skyrių iš „Turtingas(is)". Pridėti minimalų '
         'titulinį puslapį su dedikacija skaitytojui („Ačiū, kad atvėrei").',
         '„Lead magnet" — knygos pavyzdys, kurį galima atiduoti už el. paštą.',
         '~2 val.'],
        ['2',
         'Įdėti subscribe formą į elzėsknygos.lt — apačioje pagrindinio puslapio, '
         'autorės puslapio pabaigoje ir atskiros sekcijos „Newsletter" pavidalu.',
         'Trys svetainės taškai, kuriuose lankytojas gali pasirašyti.',
         '~1 val.'],
        ['2',
         'Sukurti 3 laiškų welcome seriją MailerLite automation builder\'iuose '
         '(žr. starter pack — laiškai paruošti).',
         'Pasirašiusieji savaitės bėgyje gauna 3 laiškus automatiškai.',
         '~1 val. (kopijuojant iš starter pack)'],
        ['3',
         'Pakeisti placeholder kontaktinį el. paštą į realų '
         '`autore@elzesknygos.lt` (per esamą hostingą — be papildomų išlaidų).',
         'Profesionalumo signalas; veikiantis kanalas atsiliepimams.',
         '~30 min.'],
        ['3',
         'Atnaujinti Instagram bio: pridėti aiškų CTA „Newsletter ↓" su nuoroda į '
         'subscribe puslapį (ne į pagrindinį, o tiesiai į prenumeratos langą).',
         'Bio konvertuoja sekėjus į prenumeratorius.',
         '~15 min.'],
        ['4',
         'Paskelbti 3 Instagram įrašus (postą, story seriją, vieną Reels), kuriuose '
         'minimas newsletter kaip „literatūrinis dienoraštis".',
         'Pirmoji organinio augimo banga.',
         '~2 val.'],
        ['4',
         'Asmeniškai pakviesti pirmus 30 žmonių (draugai, šeima, mokyklos draugai) '
         'pasirašyti — tai nesąžininga tik tada, kai apsimetama, jog jie „atsitiktiniai".',
         'Branduolio bazės sukūrimas.',
         '~1,5 val.'],
    ],
    col_widths_cm=[1.2, 7.5, 5.5, 2.3],
)

add_para(doc,
    'Pirmojo mėnesio sėkmės kriterijus, kurį Elzė matuos pati: bent 30 prenumeratorių '
    '(konservatyvus scenarijus iš ANALYZE) ir bent 50 % open rate pirmajam welcome '
    'laiškui (industrijos benchmark — welcome rodikliai natūraliai aukšti).')

add_h2(doc, '2.3. Kodėl pirmasis mėnuo svarbus')

add_para(doc,
    'Šiame pasiūlyme nėra rinkodarinio žygio — yra infrastruktūros sutvarkymas. Net '
    'jei Elzė nieko daugiau nedarytų, ši infrastruktūra dirbs ją 24 valandas per parą: '
    'bet kuris naujas svetainės lankytojas dabar turi būdą pasilikti. Tai yra svarbiausias '
    'darbas — be jo visi tolesni pasiūlymai negalimi.')

# ============================================================
# 3. PASIŪLYMAS Nr. 2 — VIDUTINIS
# ============================================================
add_h1(doc, '3. Pasiūlymas Nr. 2 — Vidutinis (3 mėn.): „Reguliaraus ritmo įsitvirtinimas"')

add_h2(doc, '3.1. Tikslas')

add_para(doc,
    'Kad per ateinančius 2–3 mėnesius newsletter taptų reguliarus, atpažįstamas '
    'literatūrinis kanalas, ne vienkartinis pranešimas. Iki trečio mėnesio pabaigos — '
    'bent 100 prenumeratorių (tikėtinas scenarijus) ir bent 4 paskelbti laiškai.')

add_h2(doc, '3.2. Newsletter formato pasiūlymas — „Iš Elzės dirbtuvių"')

add_para(doc,
    'Vietoje bendro „author newsletter" siūlome konkretų formatą, kuris remiasi tuo, '
    'kas Elzės autorės puslapyje jau yra: literatūra, muzika, poezija, asmeniniai '
    'pasvarstymai. Kiekvienas laiškas — keturių blokų struktūra, kurią Elzė užpildo '
    'per 30–40 min.')

add_table_simple(
    doc,
    headers=['Bloko pavadinimas', 'Ką jis daro'],
    rows=[
        ['1. Vienas sakinys',
         'Asmeninis sveikinimas. Du–trys eilutės iš Elzės savaitės — be reklamos.'],
        ['2. „Šią savaitę skaičiau / klausiausi"',
         'Vienas autorius, eilėraštis arba albumas, kurį Elzė vartojo. 3–5 sakiniai, kodėl.'],
        ['3. „Iš dirbtuvių"',
         'Trumpas užkulisinis žvilgsnis: rašymo procesas, koregavimai, abejonės. '
         'Tai turinys, kurio Instagram\'e nėra.'],
        ['4. P. S.',
         'Vienas paprastas kvietimas: signavimas Marijampolėje, mokyklos lankymas, '
         'klausimas skaitytojui. Vieną kartą — pakvietimas pirkti naują knygą '
         '(kai bus). Bet tik vieną kartą per laišką.'],
    ],
    col_widths_cm=[5.0, 11.0],
)

add_para(doc,
    'Toks laiško formatas turi du privalumus. Pirma, jis paruoštas pakartoti — '
    'kiekvieną kartą Elzė turi tik užpildyti tuos pačius keturis blokus, ne sugalvoti '
    'naują struktūrą. Antra, jis tinka skaitytojui, kuris pasirašė ne dėl '
    'pardavimo, o dėl Elzės balso — todėl jis lieka, ne atsisako prenumeratos.')

add_h2(doc, '3.3. Ritmas ir disciplinos bei pakartotinumo principas')

add_bullet(doc, 'Pirmus 3 mėn. — vienas laiškas kas 2 savaites (8 laiškai per 16 sav.).',
           bold_prefix='Periodas — ')
add_bullet(doc, 'Sekmadienio 19:00 (Lietuvos laiku) — pasiekia žmones rytdienos rytui '
           'arba pirmadieniu kelionėje į mokslus/darbą.', bold_prefix='Diena ir laikas — ')
add_bullet(doc, 'Visada tas pats. „Iš Elzės dirbtuvių #N — [vienas potemis]". Pavyzdžiui: '
           '„Iš Elzės dirbtuvių #3 — apie tai, kodėl Mačernio eilėraštį atsimenu '
           'kavinėje".', bold_prefix='Subjekto eilutės šablonas — ')
add_bullet(doc, 'Niekada neperkeliamas. Jei savaitę negali rašyti, geriau nesiunčiamas '
           'iš viso, nei „skubantis" laiškas.', bold_prefix='Pataisos mechanikos — ')

add_h2(doc, '3.4. Outreach knygų bloggeriams (papildoma veikla)')

add_para(doc,
    'Per pirmąją trijų mėnesių pusę — paruoštas asmeninis outreach laiškas (žr. starter '
    'pack laišką E5) ir išsiųsta 5–7 BookTok / Goodreads / Instagram knygų '
    'bloggeriams Lietuvoje su pasiūlymu nemokamai gauti egzempliorių mainais į '
    'sąžiningą atsiliepimą. Tikėtinas grįžimas — 1–2 sutiks. Net vienas atsakas '
    'duos 50–150 papildomų prenumeratorių per jų kanalus.')

# ============================================================
# 4. PASIŪLYMAS Nr. 3 — ILGAS
# ============================================================
add_h1(doc, '4. Pasiūlymas Nr. 3 — Ilgas (12 mėn.): „Karjeros bazė"')

add_h2(doc, '4.1. Tikslas')

add_para(doc,
    'Kad per metus newsletter taptų ne tik knygos „Turtingas(is)" rinkodaros įrankiu, '
    'bet ir karjeros baze — vieta, kurioje yra ~250–500 žmonių, jau pažįstančių Elzės '
    'literatūrinį balsą iki tos akimirkos, kai pasirodys antra knyga, eilėraščių rinkinys '
    'ar bet koks naujas kūrinys. Tai yra esminis ilgalaikis svertas: antra knyga '
    'paleidžiama jau ne į tuštumą, o į sutelktus 250–500 tikrų skaitytojų.')

add_h2(doc, '4.2. Strateginiai pjūviai per metus')

add_table_simple(
    doc,
    headers=['Mėn.', 'Pjūvis', 'Sėkmės indikatorius'],
    rows=[
        ['1–3',
         'Kanalo įjungimas + pirmas reguliarumas (Pasiūlymas Nr. 1 ir 2)',
         '100 prenumeratorių, 4 laiškai paskelbti'],
        ['4–6',
         'Mokyklų lankymas (1–2 vizitai). Newsletter signup forma kiekviename '
         'vizite — popierinė lapelis su QR kodu.',
         '+50–100 prenumeratorių per vizitus'],
        ['7–9',
         'Pirmasis „signavimo turas" (Marijampolė, Vilnius, Kaunas — bibliotekos '
         'arba kavinės). Newsletter — pagrindinis kvietimo kanalas.',
         '~250 prenumeratorių, atviras santykis su pirmais skaitytojais'],
        ['10–12',
         'Antrojo kūrinio teaserio paskelbimas TIK newsletter\'e (3 dienos pirmenybės).',
         '~500 prenumeratorių, pirma reali „pirmenybės" galios demonstracija'],
    ],
    col_widths_cm=[1.5, 8.5, 6.0],
)

add_h2(doc, '4.3. Po 12 mėn. — kaip atrodo „sėkmė"')

add_para(doc,
    'Po metų Elzė turi: (a) ~250–500 prenumeratorių, kurių 35–45 % atveria kiekvieną '
    'laišką; (b) realų pasitikėjimo signalą leidėjams arba partneriams — „turiu '
    '500 žmonių, kuriems galiu pranešti tiesiogiai"; (c) pirmąją bazę galimybei '
    'monetizuoti antrą knygą be tarpininko. Tai ne pažadas — tai pasikartojantis '
    'pavyzdys iš dešimčių indie autorių atvejų (žr. ANALYZE skyrių 3.2).')

# ============================================================
# 5. PASIŪLYMŲ SUMAVIMAS
# ============================================================
add_h1(doc, '5. Pasiūlymų sumavimas — kaštai, metrikos, prioritetas')

add_table_simple(
    doc,
    headers=['Pasiūlymas', 'Laiko kaštai', 'Pinigų kaštai', 'Pagrindinė metrika'],
    rows=[
        ['Nr. 1 (1 mėn.)',
         '~7 val. vienkartiniai + 30 min./sav.',
         '0 € (galimai 10–20 €/m. domeno paštui)',
         'Prenumeratorių sk. (≥30), 1 welcome open rate (≥50 %)'],
        ['Nr. 2 (3 mėn.)',
         '~45 min. / sav.',
         '0 €',
         'Newsletter open rate (≥35 %), prenum. sk. (≥100), 0 % perduotų laiškų'],
        ['Nr. 3 (12 mėn.)',
         'Tas pats 45–60 min. / sav. + 2–3 vizitai',
         '~30–60 € (kelionės, atspausdinti QR lapeliai)',
         'Prenum. sk. (≥250), reguliarumas neprarastas'],
    ],
    col_widths_cm=[3.0, 4.5, 3.5, 5.0],
)

add_h2(doc, '5.1. Sudėtingumo įvertinimas (realistinė skalė)')

add_bullet(doc,
    'techniškai paprastas (plugin\'ai egzistuoja), bet reikalauja vieno savaitgalio '
    'koncentracijos. Sudėtingumas: 4/10.',
    bold_prefix='Pasiūlymas Nr. 1 — ')
add_bullet(doc,
    'pagrindinis iššūkis — ne rašyti, o nepamiršti rašyti. Disciplinos klausimas. '
    'Sudėtingumas: 6/10.',
    bold_prefix='Pasiūlymas Nr. 2 — ')
add_bullet(doc,
    'reikalauja energijos ir socialinio drąsumo (signavimai, mokyklų lankymas). '
    'Tai jau autorinė karjera, o ne vien rinkodara. Sudėtingumas: 7/10.',
    bold_prefix='Pasiūlymas Nr. 3 — ')

add_h2(doc, '5.2. Kuriam pasiūlymui paruoštas „starter pack"')

add_para(doc,
    'Pagal projekto reikalavimus paruoštas pilnas starter pack Pasiūlymui Nr. 1 — t. y. '
    'kanalo įjungimui per pirmąjį mėnesį. Tai yra svarbiausias atskaitos taškas — '
    'visi tolimesni pasiūlymai be jo neveikia. Žr. atskirą priedą '
    '„STARTER_PACK_Email_kanalo_ijungimas".')

# ============================================================
# 6. STARTER PACK — APIBENDRINIMAS
# ============================================================
add_h1(doc, '6. Starter pack — kas paruošta pirmadieniui')

add_para(doc,
    'Starter pack yra fizinis dokumentas (priedas), kurį Elzė gauna kartu su pristatymu '
    'ir kuriame yra:')

add_bullet(doc,
    'paruošti tiesioginiam kopijavimui į MailerLite Welcome serijos automation. '
    'Lietuvių kalba, su Elzės balsu (parašyti remiantis jos autorės puslapio tonu).',
    bold_prefix='5 paruošti email draft\'ai — ')
add_bullet(doc,
    'su konkrečiais asmenimis (knygų bloggeriai, BookTok kūrėjos, mokyklų lietuvių '
    'kalbos mokytojai). Studentų grupė įrašo 5 konkrečius asmenis pristatymo metu, '
    'remiantis pasitarimu su Elze.',
    bold_prefix='1 outreach laiškas — ')
add_bullet(doc,
    'tekstas, kurį galima dėti į svetainę bet kurioje vietoje (1 sakinys, 3 sakiniai, '
    '1 pastraipa).',
    bold_prefix='3 subscribe formos copy variantai — ')
add_bullet(doc,
    'paprastas .txt failas su pirmojo skyriaus išskirtu tekstu, su nurodymais maketuotojui '
    'B. Strolytei (kuri jau dirbo su knyga).',
    bold_prefix='Lead magnet PDF turinys — ')
add_bullet(doc,
    'instagram bio nauja versija + 3 Instagram įrašų tekstai (postui, stories, Reels), '
    'kviečiantys į newsletter.',
    bold_prefix='Instagram CTA paketas — ')

add_para(doc,
    'Visa starter pack apimtis — apie 6–8 puslapius, kuriuose yra paruoštų laiškų '
    'tekstų, ne abstrakčių rekomendacijų. Tikslas — kad Elzė pirmadienio rytą '
    'galėtų pasakyti: „Aš galiu šitą padaryti šiandien."')

# ============================================================
out = '/projects/sandbox/dokumentai/3_IMPROVE_Pasiulymai.docx'
doc.save(out)
print(f'Sukurta: {out}')

import os
print(f'Dydis: {os.path.getsize(out)} baitų')
