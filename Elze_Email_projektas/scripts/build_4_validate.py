"""4 etapas — VALIDATE: Master Document (8–10 psl.).

Klausimas: Kaip patikrinti, kad pasiūlymas veikia, ir kada jį atsisakyti, jei
neveikia? Pilnas konsoliduotas dokumentas, paruoštas perduoti Elzei.
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
    project_title='MASTER DOCUMENT\nElzės Zdancevičiūtės knygos „Turtingas(is)" '
                  'pardavimo augimo planas\npagal Email / community building kanalą',
    channel='Email / community building',
    group='[Grupės narių vardai, pavardės]',
    date='2026 m. gegužė',
)

add_static_toc(doc, [
    ('Santrauka — vienu žvilgsniu', 2),
    ('1. Esama būsena (OBSERVE)', 3),
    ('2. Diagnostika ir potencialas (ANALYZE)', 4),
    ('3. Trys pasiūlymai (IMPROVE)', 6),
    ('4. Matavimo planas (VALIDATE)', 8),
    ('5. „Pull the plug" — kada atsisakyti', 9),
    ('6. Pirmadienio veiksmų sąrašas Elzei', 10),
    ('7. Priedai', 10),
])

add_pagebreak(doc)

# ============================================================
# SANTRAUKA
# ============================================================
add_h1(doc, 'Santrauka — vienu žvilgsniu')

add_quote_box(doc,
    '„Elzė šiandien neturi būdo paimti skaitytojo el. pašto. Per vieną savaitgalį tai '
    'pakeičiama, ir per metus iš to gali atsirasti 250–500 nuoširdžiai įsitraukusių '
    'skaitytojų bazė antrai knygai — kainuojant maždaug 60 min. per savaitę ir 0 €."')

add_para(doc,
    'Šis dokumentas yra pilnas konsoliduotas Email / community building kanalo planas '
    'Elzės Zdancevičiūtės knygai „Turtingas(is)" (Piko Valanda, 2025; tiražas 100 egz.). '
    'Jį sudaro keturi tarpusavyje sujungti etapai — esamos būsenos auditas, '
    'diagnostika ir potencialo įvertinimas, trys konkretūs pasiūlymai (1 mėn., 3 mėn., '
    '12 mėn.) bei matavimo ir „atsisakymo" planas.')

add_para(doc,
    'Pagrindinė išvada — Email kanalas Elzei šiandien faktiškai neegzistuoja, todėl '
    'projektas pradedamas nuo nulio. Tai paradoksaliai gera žinia: nereikia ardyti '
    'klaidų — tik statyti tinkamai. Visas planas yra realistinis vienam asmeniui '
    '(neprivaloma rinkodaros komanda) ir telpa į 60–90 min. per savaitę.')

# ============================================================
# 1. ESAMA BŪSENA
# ============================================================
add_h1(doc, '1. Esama būsena (OBSERVE)')

add_h2(doc, '1.1. Tiesioginio kliento turto auditas')

add_para(doc,
    'Atlikome elzėsknygos.lt ir viešų socialinių tinklų profilių auditą. Pagrindiniai '
    'objektyvūs radiniai (visi patvirtinti pagal HTML kodą, ne pagal įspūdį):')

add_bullet(doc, 'visoje trijų puslapių svetainėje yra 0 (nulis) HTML formų ir '
           '0 įvedimo laukų, kurių tipas yra „email".', bold_prefix='Email įėjimas — ')
add_bullet(doc, 'jokios užuominos apie naujienlaiškį, prenumeratą, MailerLite, '
           'Mailchimp, Substack ar kitą email-rinkimo įrankį HTML kode nėra.',
           bold_prefix='Įrankio signalai — ')
add_bullet(doc, 'pagrindinio puslapio sekcijoje „Susiekite su Elze" matomas el. '
           'paštas yra `elze.zdanceviciute@email.com`, o HTML kode `mailto:` veda '
           'į `example@email.com` — placeholder.', bold_prefix='Kontaktas — ')
add_bullet(doc, '„ĮSIGYTI KNYGĄ" mygtukas svetainės kode neturi WooCommerce, Stripe '
           'ar Paysera signalų — t. y. pati svetainė knygos neparduoda.',
           bold_prefix='Pirkimas — ')
add_bullet(doc, 'pagrindinio puslapio tekste pavadinime parašyta „EIzę" '
           '(vietoje „Elzę") — smulki, bet matoma rašybos klaida.',
           bold_prefix='Smulkmena — ')

add_h2(doc, '1.2. Tiražo kontekstas')

add_para(doc,
    'Knygos „Turtingas(is)" tiražas yra 100 egzempliorių (kolofonas svetainėje). Šis '
    'skaičius radikaliai keičia projekto tikslą: Email kanalas Elzės atveju nėra šios '
    'konkrečios knygos rinkodara, — jis yra autorinės karjeros bazė, kuri toliau '
    'naudojama antrai knygai, viešiems skaitymams ir bet kokiam ateities autorystės '
    'žingsniui.')

add_h2(doc, '1.3. Skaitytojo kelionės žemėlapis — kur ji nutrūksta')

add_para(doc,
    'Atstatėme tipinę kelionę žmogaus, kuris nepažįsta Elzės, bet GALĖTŲ tapti jos '
    'skaitytoju. Iš šešių žingsnių (atradimas, svetainė, sprendimas, pirkimas, po '
    'pirkimo, po skaitymo) du yra fatalūs — sprendimo momentas ir „po skaitymo". '
    'Abu jie nutrūksta ne dėl turinio kokybės, o dėl mechanizmo nebuvimo: žmogus, '
    'kuris jau pajuto susidomėjimą, neturi būdo „pasižymėti" ateičiai.')

# ============================================================
# 2. DIAGNOSTIKA IR POTENCIALAS
# ============================================================
add_h1(doc, '2. Diagnostika ir potencialas (ANALYZE)')

add_h2(doc, '2.1. Pagrindinė diagnozė')

add_para(doc,
    'Elzė turi labai stiprų turinio pagrindą newsletter\'iui — autentišką literatūrinį '
    'balsą (Camus, Dostojevskis, Škėma, Mačernis, Kernagis) ir asmeninę istoriją '
    '„kaip mergina, neapykantą lietuvių kalbai pavertusi knyga". Tai yra Lietuvos '
    'kontekste retas turinio rinkinys. Bet jokio mechanizmo, kuris šį balsą paverstų '
    'tęstiniu santykiu, šiuo metu nėra. Tai yra konversijos — ne pardavimo — problema.')

add_h2(doc, '2.2. Industrijos benchmark\'ai (2025–2026 m. duomenys)')

add_table_simple(
    doc,
    headers=['Rodiklis', 'Vidurkis', 'Šaltinis'],
    rows=[
        ['Newsletter open rate', '31–43 % (medijai/leidybai ~34 %, top 40+ %)',
         'MailerLite, Brevo, Thunderbit, AdMailr'],
        ['„Welcome" open rate', '4× įprasto = ~55–65 %',
         'EmailToolTester, Stackmatix'],
        ['CTR kampaniniam laiškui', '2–4 %', 'Klaviyo, Brevo'],
        ['Unsubscribe rate (laiškui)', '< 0,4 % — norma; > 1 % — perspėjimas',
         'Brevo, Campaign Monitor'],
        ['Email subscriber > IG follower', 'Ženkliai didesnė pirkimo konversija',
         'Chanticleer, Vappingo'],
    ],
    col_widths_cm=[5.0, 6.5, 4.5],
)

add_h2(doc, '2.3. Realistinė augimo prognozė')

add_table_simple(
    doc,
    headers=['Laiko tarpas', 'Konservatyvi', 'Tikėtina', 'Optimisti­nė'],
    rows=[
        ['Po 1 mėn.', '20', '40', '70'],
        ['Po 3 mėn.', '50', '100', '180'],
        ['Po 6 mėn.', '120', '250', '400'],
        ['Po 12 mėn.', '250', '500', '900'],
    ],
    col_widths_cm=[3.5, 3.5, 3.5, 3.5],
)

add_para(doc,
    'Lyginimas su realiu atveju: indie autoriaus Richie Billing case study (StoryOrigin) '
    'rodo identišką augimo trajektoriją — nuo 0 iki 100+ per pirmus 6–8 mėn. ir 400–500 '
    'per metus, naudojant tik reader magnet ir newsletter swap\'us. „Tikėtinas" '
    'scenarijus nėra optimistinis — jis pasikartojantis.')

add_h2(doc, '2.4. Kaštai (laiko + pinigų)')

add_para(doc,
    'Pradinis nustatymas — apie 6,5–7 val. (vienas savaitgalis). Tęstinė rutina — apie '
    '55–60 min. per savaitę. Piniginiai kaštai pirmus 6–9 mėn. — 0 € (MailerLite '
    'nemokama versija iki 500 prenum.; Substack visiškai nemokama). Mokamas planas '
    '(~10 €/mėn.) reikalingas tik tada, kai prenumeratorių skaičius artėja prie 500.')

add_h2(doc, '2.5. Trys didžiausios kliūtys')

add_numbered(doc, 1,
    'visoje svetainėje neegzistuoja jokia subscribe forma. Vienos eilutės techninis '
    'sprendimas (WordPress plugin\'as), bet šiandien jis kainuoja kiekvieną svetainės '
    'lankytoją.',
    bold_prefix='Kliūtis A — kanalo įėjimo nėra. ')
add_numbered(doc, 2,
    'Elzė turi stiprią kūrybinę istoriją, bet neturi „skaitymo priežasties grįžti". '
    'Newsletter yra reguliaraus santykio kanalas; jam reikia paprasto pažado.',
    bold_prefix='Kliūtis B — nėra aiškaus skaitytojo pažado. ')
add_numbered(doc, 3,
    'placeholder kontaktinis el. paštas mažina pasitikėjimą. Real kontaktas — '
    'paprastas, bet būtinas profesionalumo signalas.',
    bold_prefix='Kliūtis C — pirminio pasitikėjimo signalų trūkumas. ')

# ============================================================
# 3. TRYS PASIŪLYMAI
# ============================================================
add_h1(doc, '3. Trys pasiūlymai (IMPROVE)')

add_h2(doc, '3.1. Pasiūlymas Nr. 1 — Trumpas (1 mėn.): „Kanalo įjungimas"')

add_para(doc,
    'Tikslas: per 4 savaites Elzės svetainėje atsiranda veikianti subscribe forma; '
    'pasirašiusieji gauna 3 laiškų welcome seką su nemokama knygos ištrauka; iki '
    'mėnesio pabaigos prenumeratorių skaičius — 30–70.')

add_table_simple(
    doc,
    headers=['Sav.', 'Užduotis', 'Laikas'],
    rows=[
        ['1', 'MailerLite paskyra + sąrašas + DKIM/SPF', '~1,5 val.'],
        ['1', 'Lead magnet PDF (1-as skyrius)', '~2 val.'],
        ['2', 'Subscribe forma 3 vietose svetainėje', '~1 val.'],
        ['2', '3 welcome laiškų automation (iš starter pack)', '~1 val.'],
        ['3', 'Realus el. paštas autore@elzesknygos.lt', '~30 min.'],
        ['3', 'Instagram bio CTA atnaujinimas', '~15 min.'],
        ['4', '3 IG įrašai apie newsletter', '~2 val.'],
        ['4', 'Asmeniškas pakvietimas 30 žmonėms', '~1,5 val.'],
    ],
    col_widths_cm=[1.5, 11.0, 3.5],
)

add_para(doc,
    'Sėkmės kriterijus 1 mėn.: ≥30 prenumeratorių; ≥50 % open rate pirmojo welcome '
    'laiško. Kaštai: ~7 val. vienkartiniai + 30 min./sav.; 0 € (galimai 10–20 €/m. domeno paštui).')

add_h2(doc, '3.2. Pasiūlymas Nr. 2 — Vidutinis (3 mėn.): „Reguliaraus ritmo įsitvirtinimas"')

add_para(doc,
    'Tikslas: per 2–3 mėnesius newsletter tampa reguliarus, atpažįstamas literatūrinis '
    'kanalas. Iki trečio mėn. pabaigos — ≥100 prenumeratorių, ≥4 paskelbti laiškai.')

add_para(doc,
    'Newsletter formatas — „Iš Elzės dirbtuvių" — keturių blokų struktūra: (1) vienas '
    'sakinys iš savaitės; (2) „šią savaitę skaičiau / klausiausi" — autorius, eilėraštis '
    'ar albumas; (3) „iš dirbtuvių" — užkulisinis žvilgsnis į rašymo procesą; '
    '(4) P. S. — vienas paprastas kvietimas (signavimas, klausimas skaitytojui). '
    'Ritmas — vienas laiškas kas 2 savaites, sekmadienio 19:00.')

add_para(doc,
    'Papildoma veikla: outreach 5–7 BookTok / Goodreads / Instagram knygų bloggeriams '
    'Lietuvoje (paruoštas laiškas — starter pack laiškas E5). Tikėtina — 1–2 sutiks; '
    '+50–150 prenumeratorių per jų kanalus.')

add_para(doc,
    'Sėkmės kriterijus 3 mėn.: open rate ≥35 %; prenum. sk. ≥100; jokio laiško '
    'praleidimo (reguliarumas — svarbiausias). Kaštai: ~45 min./sav.; 0 €.')

add_h2(doc, '3.3. Pasiūlymas Nr. 3 — Ilgas (12 mėn.): „Karjeros bazė"')

add_para(doc,
    'Tikslas: per metus newsletter tampa karjeros baze — vieta su 250–500 žmonių, '
    'pažįstančių Elzės balsą iki tos akimirkos, kai pasirodys antra knyga ar bet koks '
    'naujas kūrinys.')

add_table_simple(
    doc,
    headers=['Mėn.', 'Pjūvis', 'Indikatorius'],
    rows=[
        ['1–3', 'Įjungimas + pirmas reguliarumas', '~100 prenum., 4 laiškai'],
        ['4–6', 'Mokyklų lankymas (1–2 vizitai)', '+50–100 prenum.'],
        ['7–9', 'Signavimo turas (Marijampolė, Vilnius, Kaunas)', '~250 prenum.'],
        ['10–12', 'Antrojo kūrinio teaserio paskelbimas TIK newsletter\'e',
         '~500 prenum.'],
    ],
    col_widths_cm=[1.5, 8.5, 6.0],
)

add_para(doc,
    'Sėkmės kriterijus 12 mėn.: prenum. sk. ≥250 (konservatyvus) iki ≥500 (tikėtinas); '
    'reguliarumas neprarastas. Kaštai: tas pats 45–60 min./sav. + 30–60 €/m. (kelionės, '
    'QR lapeliai).')

# ============================================================
# 4. MATAVIMO PLANAS
# ============================================================
add_h1(doc, '4. Matavimo planas (VALIDATE)')

add_h2(doc, '4.1. Ką matuoti — ir ko nematuoti')

add_para(doc,
    'Sąmoningai siūlome trumpą metrikų sąrašą. Daugiau metrikų reiškia daugiau dėmesio '
    'metrikoms ir mažiau — laiškams. Vienam autoriui per savaitę galima sąžiningai '
    'sekti 4–5 skaičius, ne daugiau.')

add_table_simple(
    doc,
    headers=['Metrika', 'Tikslas (1 mėn.)', 'Tikslas (3 mėn.)', 'Tikslas (12 mėn.)'],
    rows=[
        ['Bendras prenum. skaičius', '≥30', '≥100', '≥250'],
        ['Welcome serijos open rate', '≥50 %', '≥50 %', '≥50 %'],
        ['Įprasto laiško open rate', '— (per anksti)', '≥35 %', '≥35 %'],
        ['CTR (paspaudimo dažnis)', '— (per anksti)', '≥3 %', '≥3 %'],
        ['Unsubscribe rate (laiškui)', '<1 %', '<0,5 %', '<0,5 %'],
    ],
    col_widths_cm=[5.0, 3.5, 3.5, 4.0],
)

add_h2(doc, '4.2. Kaip matuoti')

add_bullet(doc, 'Visus skaičius rodo MailerLite/Substack panelis automatiškai. '
           'Nereikia jokio papildomo įrankio.', bold_prefix='Įrankis — ')
add_bullet(doc, 'Kartą per mėnesį (paskutinį sekmadienį) Elzė užrašo 5 metrikas '
           'į vieną Google Sheet eilutę. Užtrunka <10 min.',
           bold_prefix='Periodas — ')
add_bullet(doc, 'Po 3 mėn. Elzė palygina prognozę su realybe. Po 6 mėn. — pakartojama. '
           'Po 12 mėn. — sprendžiama, ar tęsti taip, ar keisti formatą.',
           bold_prefix='Patikrinimas — ')

add_h2(doc, '4.3. Sėkmės ir nesėkmės kriterijai')

add_table_simple(
    doc,
    headers=['Etapas', 'Sėkmė =', 'Nesėkmė ='],
    rows=[
        ['Po 1 mėn.', '≥30 prenum., welcome open ≥50 %',
         '<15 prenum. ARBA welcome open <30 %'],
        ['Po 3 mėn.', '≥100 prenum., open rate ≥30 %, 0 praleistų laiškų',
         '<50 prenum. ARBA open rate <20 % ARBA praleisti 2+ laiškai iš eilės'],
        ['Po 12 mėn.', '≥250 prenum., reguliarumas neprarastas',
         '<150 prenum. arba reguliarumas prarastas >2 mėn.'],
    ],
    col_widths_cm=[3.0, 6.5, 6.5],
)

# ============================================================
# 5. PULL THE PLUG
# ============================================================
add_h1(doc, '5. „Pull the plug" — kada atsisakyti')

add_para(doc,
    'Šio etapo tikslas — pagal rubriką — pasakyti, KADA Elzė turėtų sąžiningai pasakyti '
    '„šitas kanalas man neveikia, einu kitur". Mūsų kaip konsultantų pareiga — '
    'pateikti aiškią ribą, ne saugoti save nuo nepalankaus rezultato.')

add_h2(doc, '5.1. Trys konkretūs „pull the plug" indikatoriai')

add_numbered(doc, 1,
    'jei po 6 sav. nuo welcome serijos paleidimo Elzė turi <20 prenumeratorių IR '
    'open rate <30 %, vadinasi, arba subscribe forma blogai matoma, arba lead magnet '
    'nepatraukia. Tai keičiama greitai (per dieną), bet jei po dar 4 sav. niekas '
    'nepasikeitė — Elzė turėtų pereiti prie kito kanalo (pvz., TikTok / BookTok), '
    'nes Email auditorijai jos reikia kažkaip pirma pasiekti.',
    bold_prefix='Indikatorius #1 (audiencija). ')

add_numbered(doc, 2,
    'jei po 3 mėn. Elzė nesugebėjo paskelbti bent 4 laiškų, vadinasi, formatas '
    'jai per sunkus arba ritmas per intensyvus. Tai NĖRA Elzės kaltė — tai signalas, '
    'kad reikia perskaityti formatą (gal vieną laišką per mėnesį, ne kas dvi savaites; '
    'gal trumpesnį formatą; gal Substack vietoje MailerLite). Bet jei po 6 mėn. '
    'situacija nesikeičia — kanalas Elzei netinka šiame karjeros etape, ir tai '
    'sąžiningas atsakymas.',
    bold_prefix='Indikatorius #2 (kūrybinis krūvis). ')

add_numbered(doc, 3,
    'jei po 6 mėn. open rate <20 %, o unsubscribe rate >2 % per laišką, vadinasi, '
    'prenumeratoriams nepasakomas tas pats turinys, kuriam jie pasirašė. Reikia '
    'arba keisti formatą (po vienos atviros diskusijos su skaitytojais — '
    'newsletter\'yje galima paklausti), arba pripažinti, kad konkretus literatūrinis '
    'kampas nėra tas, kuris šią auditoriją laiko.',
    bold_prefix='Indikatorius #3 (santykio kokybė). ')

add_h2(doc, '5.2. Kas nėra „pull the plug" signalas')

add_para(doc,
    'Svarbu paminėti, ko NEIŠLAIKYTI kaip nesėkmės. Pirma — lėtas augimas. Pirmus '
    '3 mėn. augimas yra natūraliai lėtas; kanalo augimo kreivė yra eksponentinė, ne '
    'linijinė, todėl 50 prenumeratorių po 3 mėn. nereiškia, kad po 12 mėn. bus 200. '
    'Antra — vienas blogai pavykęs laiškas. Kiekvienas autorius turi laiškų, kurie '
    'neatsiveria; svarbu serija, ne vienetas. Trečia — palyginimas su didesniais '
    'autoriais. Mūsų prognozė remiasi indie atvejais (Richie Billing tipas), ne '
    'Lietuvos žvaigždėmis su komandomis ir biudžetais.')

# ============================================================
# 6. PIRMADIENIO RYTAS
# ============================================================
add_h1(doc, '6. Pirmadienio veiksmų sąrašas Elzei')

add_para(doc,
    'Tai yra šio dokumento svarbiausia dalis pagal rubriką: ką Elzė turėtų padaryti '
    'pirmadienio rytą po pristatymo, kad projektas neliktų teorijoje.')

add_numbered(doc, 1,
    'atidaryti mailerlite.com, susikurti nemokamą paskyrą su `autore@elzesknygos.lt` '
    'adresu (jei dar neturi šio adreso — pirma jį susikurti per esamą hostingą). '
    'Užtrunka ~30 min.',
    bold_prefix='07:30 — ')

add_numbered(doc, 2,
    'atsidaryti starter pack failą „STARTER_PACK_Email_kanalo_ijungimas" ir '
    'pasiruošti kopijuoti laiškų tekstus tiesiai į MailerLite Welcome automation. '
    '5 paruošti laiškai — viskas paruošta, lieka tik 5 minutės kiekvienam '
    'patikrinti ir paspausti „Save".',
    bold_prefix='08:00 — ')

add_numbered(doc, 3,
    'paimti knygos „Turtingas(is)" pirmąjį skyrių (jau yra rankraštyje), išsaugoti '
    'kaip PDF su titulinio puslapiu „Pirmieji puslapiai. Ačiū, kad atvėrei. — Elzė". '
    'Įkelti į MailerLite kaip welcome laiško priedą.',
    bold_prefix='09:00 — ')

add_numbered(doc, 4,
    'įdėti subscribe formą į svetainę. Pati forma sugeneruojama MailerLite\'e (HTML '
    'kodo blokas) ir įklijuojama trijose vietose: pagrindinio puslapio apačioje, '
    'autorės puslapio pabaigoje, naujoje atskiroje sekcijoje „Newsletter".',
    bold_prefix='10:30 — ')

add_numbered(doc, 5,
    'pakeisti Instagram bio: pridėti eilutę „Newsletter ↓ link in bio" ir paskutiniame '
    'linktree linke (arba bio nuoroda) — naujasis prenumeratos puslapis.',
    bold_prefix='11:30 — ')

add_numbered(doc, 6,
    'asmeniškai parašyti 10 artimiems žmonėms (Marijampolėje ir už jos ribų), prašant '
    'pasirašyti pirmaisiais. Tai sąžininga branduolio bazė, ne reklama.',
    bold_prefix='13:00 — ')

add_para(doc,
    'Iš viso pirmadienis — apie 5–6 val. Kanalas šiandien iš nulinės būsenos pereina '
    'į veikiantį.')

# ============================================================
# 7. PRIEDAI
# ============================================================
add_h1(doc, '7. Priedai')

add_bullet(doc, 'Esamos būsenos screenshot\'ai elzėsknygos.lt (A1–A6).',
           bold_prefix='Priedas A — ')
add_bullet(doc, '5 paruošti email draft\'ai welcome serijai + 1 outreach laiškas '
           'knygų bloggeriams + 3 subscribe formos copy variantai + Instagram CTA '
           'paketas.',
           bold_prefix='STARTER_PACK_Email_kanalo_ijungimas.docx — ')
add_bullet(doc, 'OBSERVE etapas, ANALYZE etapas, IMPROVE etapas — atskiri pilni '
           'dokumentai, jei norima skaityti detaliau.',
           bold_prefix='Atskiri etapų dokumentai — ')

# ============================================================
out = '/projects/sandbox/dokumentai/4_VALIDATE_Master_Document.docx'
doc.save(out)
print(f'Sukurta: {out}')

import os
print(f'Dydis: {os.path.getsize(out)} baitų')
