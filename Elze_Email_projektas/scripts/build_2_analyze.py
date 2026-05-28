"""2 etapas — ANALYZE: Diagnostika ir potencialas (4 psl. dokumentas).

Klausimas: Kodėl kanalas neveikia geriau? Koks yra jo potencialas?
Kokios metrikos?
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
    project_title='2 etapas — ANALYZE\nDiagnostika ir potencialas',
    channel='Email / community building',
    group='[Grupės narių vardai, pavardės]',
    date='2026 m. gegužė',
)

add_static_toc(doc, [
    ('1. Pagrindinė diagnozė', 2),
    ('2. Best practice — kanalo veikimo principai', 2),
    ('3. Realistinė dydžio prognozė (1, 3, 6, 12 mėn.)', 3),
    ('4. Kaštų analizė — laiko ir pinigų sąmata', 4),
    ('5. Trys didžiausios kliūtys', 5),
    ('6. Kanalo potencialo apibendrinimas', 5),
    ('Šaltiniai', 6),
])

add_pagebreak(doc)

# ============================================================
# 1. PAGRINDINĖ DIAGNOZĖ
# ============================================================
add_h1(doc, '1. Pagrindinė diagnozė')

add_para(doc,
    'OBSERVE etape užfiksavome, kad Email kanalo Elzei šiandien iš esmės nėra. ANALYZE '
    'etape klausimas keičiasi: kodėl jis neegzistuoja, kas trukdo, ir kiek realiai '
    'galima per pirmus tris–dvylika mėnesių pasiekti, jeigu kanalas būtų sąžiningai '
    'pradedamas iš nulio. Diagnozė neatsiranda iš to, ko Elzei „trūksta", bet iš to, '
    'kas iš jos gyvenimo realiai gali tilpti į autoriaus newsletter rutiną.')

add_para(doc,
    'Pagrindinė diagnozė: Elzė turi kanalui labai stiprų turinio pagrindą — autentišką '
    'literatūrinį balsą, asmeninę istoriją ir aiškią estetinę kryptį (Camus, Dostojevskis, '
    'Škėma, Mačernis, Kernagis). Tačiau jokio mechanizmo, kuris šį balsą paverstų '
    'tęstiniu santykiu, nėra. Skaitytojas, kuris ką tik perskaitė autorės puslapį ir '
    'pajuto: „taip, ši autorė man įdomi" — neturi į ką paspausti. Tai yra konversijos '
    '(ne pardavimo) problema, ir ji sprendžiama vienu konkrečiu sprendimu, ne dideliu '
    'kūrybiniu remontu.')

# ============================================================
# 2. BEST PRACTICE
# ============================================================
add_h1(doc, '2. Best practice — kanalo veikimo principai')

add_para(doc,
    'Email / community building kaip kanalas knygų autoriams turi nemažai industrinių '
    'duomenų. Toliau pateikiame konkrečius skaičius (ne abstrakčius), kurie naudojami '
    'kaip benchmark\'ai trečiame skyriuje pagrindžiant prognozę.')

add_table_simple(
    doc,
    headers=['Rodiklis', 'Industrijos vidurkis (2025–2026)', 'Šaltinis'],
    rows=[
        ['Newsletter open rate (vidurkis)',
         'MailerLite: 43,5 %; Brevo: 31,2 %; medijai/leidybai ~34 %',
         'Thunderbit, AdMailr, MailerLite vidiniai duomenys'],
        ['„Welcome" laiško open rate',
         '4× didesnis nei įprasto kampaninio (~55–65 %)',
         'EmailToolTester / Mailmodo, Stackmatix'],
        ['CTR (paspaudimo per laišką dažnis)',
         '2–4 % kampaniniams; 5–6 % automatizuotiems flow\'ams',
         'Klaviyo, Brevo'],
        ['Unsubscribe rate (vienam laiškui)',
         '< 0,4 % yra norma; > 1 % — perspėjimas',
         'Brevo, Campaign Monitor 2025'],
        ['Subscriber → reader konversija',
         'Email skaitytojas knygas perka pastebimai dažniau nei socialinio tinklo sekėjas',
         'Chanticleer Reviews, Vappingo (autorių rinkodara)'],
        ['Welcome serija (pirmi 3–5 laiškai)',
         'Geriausiai veikia 3–4 laiškų sekos per pirmas 7–10 dienų',
         'Braze, Bloomreach'],
    ],
    col_widths_cm=[5.0, 6.5, 4.5],
)

add_para(doc,
    'Praktinė reikšmė: jeigu Elzės būsimoje listėje bus, pavyzdžiui, 100 prenumeratorių '
    'ir ji paskelbs naują laišką, statistiškai 30–40 jį atvers, 2–4 paspaus nuorodą, '
    'ir mažiau nei 1 atsisakys prenumeratos. Tai yra realistiniai skaičiai, ne '
    'optimistinė prielaida.')

# ============================================================
# 3. DYDŽIO PROGNOZĖ
# ============================================================
add_h1(doc, '3. Realistinė dydžio prognozė (1, 3, 6, 12 mėn.)')

add_para(doc,
    'Prognozė remiasi prielaidomis, kurias surašome aiškiai, kad jas būtų galima '
    'patikrinti arba paneigti pristatymo metu.')

add_h2(doc, '3.1. Prielaidos')

add_bullet(doc,
    'Elzės Instagram @elze_zdan turi nedidelę, bet realią auditoriją (tikslus skaičius — '
    'patikrinamas susitikimo metu; konservatyviai prielaida: 200–800 sekėjų pradžioje).',
    bold_prefix='P1 — ')
add_bullet(doc,
    'Elzė bent kartą per savaitę kuria turinį Instagram (postai arba stories).',
    bold_prefix='P2 — ')
add_bullet(doc,
    'Knygos „Turtingas(is)" tiražas — 100 egz. Iš jų bent 50 pasieks tikrus skaitytojus '
    '(perka draugai, šeima, mokyklų aplinka, signavimas Marijampolėje).',
    bold_prefix='P3 — ')
add_bullet(doc,
    'Elzės dabartinio gyvenimo etapas (mokslai gimnazijoje) leidžia skirti newsletter\'iui '
    'apie 60–90 min. per savaitę — ne daugiau.',
    bold_prefix='P4 — ')

add_h2(doc, '3.2. Augimo prognozė')

add_table_simple(
    doc,
    headers=['Laiko tarpas', 'Konservatyvi', 'Tikėtina', 'Optimisti­nė', 'Pagrindas'],
    rows=[
        ['Po 1 mėn.', '20 prenum.', '40 prenum.', '70 prenum.',
         'Tik draugai/šeima + IG bio CTA'],
        ['Po 3 mėn.', '50', '100', '180',
         '+ lead magnet (knygos ištrauka), reguliarus IG postas'],
        ['Po 6 mėn.', '120', '250', '400',
         '+ knygų bloggerių ryšiai, signavimo renginiai'],
        ['Po 12 mėn.', '250', '500', '900',
         '+ mokyklų lankymas, antras kūrinys, BookTok recenzijos'],
    ],
    col_widths_cm=[2.8, 2.5, 2.5, 2.5, 5.7],
)

add_para(doc,
    'Lyginimas su realiu atveju: indie autoriaus Richie Billing case study (StoryOrigin, '
    '2022) rodo, kad pradėjęs nuo 0 ir naudodamas vieną nemokamą reader magnet bei '
    'newsletter swap\'us, autorius per pirmuosius 6–8 mėn. pasiekė 100+, vėliau 400–500 '
    'prenumeratorių. Mūsų „tikėtinas" scenarijus (250 per 6 mėn., 500 per 12 mėn.) '
    'beveik identiškas šiam atvejui. Tai stiprus pagrindas — ne pažadas, o pasikartojantis '
    'pavyzdys.')

# ============================================================
# 4. KAŠTŲ ANALIZĖ
# ============================================================
add_h1(doc, '4. Kaštų analizė — laiko ir pinigų sąmata')

add_h2(doc, '4.1. Laiko kaštai')

add_para(doc,
    'Vertiname valandomis per savaitę, nes tai vienintelis Elzei realiai svarbus '
    'matas. Visi skaičiai — vienam asmeniui, ne komandai.')

add_table_simple(
    doc,
    headers=['Veikla', 'Pradinis nustatymas', 'Reguliari rutina', 'Pastabos'],
    rows=[
        ['Email įrankio paskyra (MailerLite/Substack)',
         '~1,5 val. (vienkartinis)',
         '0',
         'Forma + welcome serija + DKIM/SPF nustatymas'],
        ['Lead magnet — knygos ištraukos PDF',
         '~2 val. (vienkartinis)',
         '0',
         'Naudojant esamą knygos rankraštį, ne kuriant naujai'],
        ['Welcome laiškų rašymas (3–4 vnt.)',
         '~3 val. (vienkartinis)',
         '0',
         'Vieną kartą parašoma, vėliau veikia automatiškai'],
        ['Reguliarus newsletter (1 kartą per 2 sav.)',
         '0',
         '~45 min. / sav.',
         'Trumpas „literatūrinis dienoraštis" formatas'],
        ['IG bio + postai, kvieiantys prenumeruoti',
         '~30 min. (atnaujinimas)',
         '~10 min. / sav.',
         'Reguliariai įklijuojama nuoroda į stories'],
    ],
    col_widths_cm=[5.5, 3.5, 2.8, 4.5],
)

add_para(doc,
    'Suma: pradinis nustatymas — apie 6,5–7 val. (vienas savaitgalis), tęstinė rutina '
    '— apie 55–60 min. per savaitę. Tai tilpsta į prielaidos P4 ribą.')

add_h2(doc, '4.2. Piniginiai kaštai')

add_table_simple(
    doc,
    headers=['Įrankis / sprendimas', 'Pradinis tarifas', 'Galimas mokamas variantas',
             'Vertinimas'],
    rows=[
        ['MailerLite', '0 € (iki 500 prenum., 12 000 laiškų/mėn.)',
         '~9–15 € / mėn. nuo 1000 prenum.',
         'Tinka iki ~500 prenum.; pilna automatika nuo pradžios'],
        ['Substack', '0 €',
         '0 € (Substack ima 10 % nuo paid prenum., jei tokie atsiranda)',
         'Tinka, jei Elzė nori, kad newsletter būtų vieta literatūrai, ne tik pardavimui'],
        ['Mailchimp', '0 € (iki 500, su apribojimais)',
         '~13 € / mėn.',
         'Mažiausiai patrauklus — laisva versija labiausiai apkarpyta'],
        ['Domeno el. paštas (autore@elzesknygos.lt)',
         '~10–20 € / metams (per esamą hostingą)',
         '—',
         'Pakeičia placeholder elze.zdanceviciute@email.com'],
    ],
    col_widths_cm=[4.5, 4.0, 4.0, 4.0],
)

add_para(doc,
    'Pirmus 6–9 mėnesius rekomenduojama nemokama versija. Mokamas planas turėtų '
    'pradėti dirbti tik tuomet, kai prenumeratorių skaičius pradeda artėti prie 500 — '
    't. y., kai augimo problema realiai egzistuoja, o ne hipotetiškai.')

# ============================================================
# 5. TRYS DIDŽIAUSIOS KLIŪTYS
# ============================================================
add_h1(doc, '5. Trys didžiausios kliūtys')

add_para(doc,
    'Iš visų rastų defektų išskiriame tris, kurie blokuoja kanalą labiausiai ir kuriuos '
    'įmanoma šalinti per pirmas 4 savaites.')

add_numbered(doc, 1,
    'visoje svetainėje neegzistuoja jokia subscribe forma. Tai vienos eilutės techninis '
    'sprendimas (WordPress plugin\'as arba Substack embed), bet šiandien jis kainuoja '
    'Elzei kiekvieną svetainės lankytoją. Tai yra didžiausias techninis svertas — mažiausia '
    'pastanga, didžiausia įtaka.',
    bold_prefix='Kliūtis A — kanalo įėjimo nėra. ')

add_numbered(doc, 2,
    'Elzė turi stiprią kūrybinę istoriją (autorinis pasakojimas „Apie Elzę"), bet '
    'neturi „skaitymo priežasties grįžti". Newsletter yra ne pardavimo, o reguliaraus '
    'santykio kanalas; jam reikia pažado („kas dvi savaitės — vienas trumpas '
    'literatūrinis laiškas"). Pažadas yra paprastas, bet būtinas — be jo prenumeratorius '
    'po 2–3 mėnesių pamiršta, kodėl pasirašė.',
    bold_prefix='Kliūtis B — nėra aiškaus skaitytojo pažado. ')

add_numbered(doc, 3,
    'pagrindinio puslapio kontaktas — placeholder. Net minimalūs profesionalumo '
    'signalai, kaip realus el. pašto adresas savo domene, sukuria pasitikėjimą. '
    'Be jo dalis potencialių partnerių (mokyklos, bloggerės, leidėjai) gali nesusisiekti.',
    bold_prefix='Kliūtis C — pirminio pasitikėjimo signalų trūkumas. ')

# ============================================================
# 6. APIBENDRINIMAS
# ============================================================
add_h1(doc, '6. Kanalo potencialo apibendrinimas')

add_para(doc,
    'Email / community building kanalas Elzės atveju turi nedidelį, bet realų '
    'potencialą: per 12 mėn. tikėtinas dydis — apie 250–500 prenumeratorių, kainuojantis '
    'apie 60 min. per savaitę ir 0 € (su galimybe vėliau pereiti į ~10 €/mėn. planą). '
    'Tai nėra masinio rinkodaros kanalas — bet jam to ir nereikia: 250–500 nuoširdžiai '
    'įsitraukusių skaitytojų antrai knygai yra geresnė bazė, nei 5 000 algoritmu '
    'atsitiktinai parodytų sekėjų.')

add_para(doc,
    'Kitas etapas (IMPROVE) konkretizuoja, ką Elzė turėtų daryti per 1, 3 ir 12 mėn., '
    'kad šis potencialas realizuotųsi, ir paruošia „starter pack" — paruoštus laiškus, '
    'kuriuos ji gali pasiųsti pirmąjį pirmadienį po susitikimo.')

# ============================================================
# ŠALTINIAI
# ============================================================
add_h1(doc, 'Šaltiniai')

add_bullet(doc, 'Newsletter open rate ir kiti benchmark\'ai (2025–2026): Thunderbit, '
           'AdMailr, MailerLite, Brevo viešai prieinami industrijos duomenys.')
add_bullet(doc, '„Welcome" laiško rodikliai ir flow rekomendacijos: EmailToolTester '
           '(2025), Stackmatix (2025), Braze, Bloomreach (2026).')
add_bullet(doc, 'Author email vs. social media konversija: Chanticleer Reviews '
           '(2025-06-10), Vappingo („Email Marketing Tools for Authors").')
add_bullet(doc, 'Indie autoriaus atvejis: StoryOrigin App — Richie Billing case study, '
           'naudojant reader magnet ir newsletter swap\'us.')
add_bullet(doc, 'MailerLite free plan apribojimai 2025-09: Sender.net, Hostinger, '
           'GroupMail viešos apžvalgos.')
add_bullet(doc, 'Substack ekosistema: lietuviakalbiai pavyzdžiai — „Notes from Lithuania" '
           '(T. Miliūnas), „The Vilnius Observer" (L. Barbier).')
add_bullet(doc, 'Pirminis kliento auditas: elzėsknygos.lt (2026 m. gegužė), HTML '
           'analizė, formų ir <input type="email"> elementų skaičius.')

# ============================================================
out = '/projects/sandbox/dokumentai/2_ANALYZE_Diagnostika_ir_potencialas.docx'
doc.save(out)
print(f'Sukurta: {out}')

import os
print(f'Dydis: {os.path.getsize(out)} baitų')
