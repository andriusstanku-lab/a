"""6 dokumentas — 10 min. pristatymo struktūra ir kalbėjimo planas.

Pagal rubriką: 1 sakinys, esama būsena, diagnostika, pasiūlymai, matavimas,
„pirmadienio rytas".
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
    project_title='10 MIN. PRISTATYMO STRUKTŪRA\nElzės Zdancevičiūtės knygos „Turtingas(is)" '
                  'Email kanalo planas',
    channel='Email / community building',
    group='[Grupės narių vardai, pavardės]',
    date='2026 m. gegužė',
)

add_static_toc(doc, [
    ('1. Bendros taisyklės pristatymui', 2),
    ('2. Skaidres ir kalbėjimo planas (laiko žymėmis)', 2),
    ('3. Tikėtini Elzės klausimai ir mūsų atsakymai', 5),
    ('4. Vaidmenų pasiskirstymas grupėje', 6),
    ('5. „Cheat sheet" — kalbant be skaidrių', 6),
])

add_pagebreak(doc)

# ============================================================
# 1. BENDROS TAISYKLĖS
# ============================================================
add_h1(doc, '1. Bendros taisyklės pristatymui')

add_bullet(doc, 'Elzė turi išklausyti, ne mes — pasiteisinti. Visada kreipiamės į ją '
           '(„Elze, mums atrodo…", „Tau pravartu žinoti…"), o ne į auditoriją.',
           bold_prefix='Auditorija — ')
add_bullet(doc, '10 min. = 600 sek. Vidutinis kalbėjimo greitis 130 ž./min. = ~1300 žodžių. '
           'Tekstas, kurį surašome — apie 1100–1200 žodžių, kad turėtume rezervo pauzėms.',
           bold_prefix='Tempas — ')
add_bullet(doc, 'Mažiau slide\'ų — geriau. 6–7 skaidrės maks. Po vieną didelę mintį per skaidrę.',
           bold_prefix='Skaidres — ')
add_bullet(doc, 'Ne „aš jaučiu, kad…" — „Mes patikrinome HTML kodą, ir radome 0 formų". '
           'Visi teiginiai paremiami konkretumu.', bold_prefix='Faktiškumas — ')
add_bullet(doc, 'Į kiekvieną kritinę galimybę galimas atsakymas yra paruoštas '
           '(žr. 3 skyrių). Niekada „nežinau, paklausiu vėliau" — visada bandymas '
           'atsakyti dabar arba aiški riba: „Šito nepatikrinome — patikrinsime per dieną."',
           bold_prefix='Klausimai — ')

# ============================================================
# 2. SKAIDRĖS IR KALBĖJIMO PLANAS
# ============================================================
add_h1(doc, '2. Skaidres ir kalbėjimo planas (laiko žymėmis)')

# ---- 0:00 - 1:00 ----
add_h2(doc, '0:00–1:00 — Vienas sakinys')

add_quote_box(doc,
    'SKAIDRĖ 1 — vienas sakinys per visą ekraną:\n'
    '„Elzė šiandien neturi būdo paimti skaitytojo el. pašto. Per vieną savaitgalį tai '
    'pakeičiama, ir per metus iš to gali atsirasti 250–500 nuoširdžiai įsitraukusių '
    'skaitytojų bazė antrai knygai."')

add_para(doc,
    'KALBĖJIMAS (apie 80 ž., 35 sek., palieka 25 sek. tylai ir akių kontaktui):',
    bold=True)

add_para(doc,
    '„Elze, prieš pradedant — mūsų visa istorija sutelpa į vieną sakinį, kurį '
    'matai ekrane. Šiandien tu turi puikią svetainę, gerą knygą ir realią Instagram '
    'auditoriją — bet tu neturi nė vieno būdo paimti skaitytojo el. pašto. Mūsų '
    'pasiūlymas yra: per vieną savaitgalį tai sutvarkyti, o per metus iš to gauti '
    '250–500 žmonių, kurie tave skaitys ir tada, kai pasirodys antroji knyga."')

# ---- 1:00 - 3:00 ----
add_h2(doc, '1:00–3:00 — Esama būsena (OBSERVE)')

add_quote_box(doc,
    'SKAIDRĖ 2 — screenshot\'ai ir trys skaičiai:\n'
    '• 0 (nulis) HTML formų visoje svetainėje\n'
    '• Kontaktinis paštas — placeholder „example@email.com"\n'
    '• Tiražas — 100 egz.\n'
    '[Apatinė juosta: skaitytojo kelionės žemėlapis, žingsniai 3 ir 6 paryškinti '
    'raudonai — „kelionė nutrūksta čia"]')

add_para(doc,
    'KALBĖJIMAS (apie 220 ž., 1 min. 40 sek.):',
    bold=True)

add_para(doc,
    '„Pradėjome nuo tiesioginio audito — kaip vartotojai. Atsidarėme tavo svetainę '
    'inkognito lange, perėjome ją taip, kaip eitų bet kuris žmogus, kuris '
    'nepažįsta tavęs. Trys dalykai:')

add_para(doc,
    '„Pirma — visoje svetainėje yra nulis HTML formų ir nulis input laukų, kurių '
    'tipas yra „email". Tai reiškia, kad jei tavo puslapį šiandien aplankys tūkstantis '
    'žmonių, nė vieno iš jų niekas neištruks po vizito.')

add_para(doc,
    '„Antra — kontaktinis adresas, kurį matome — „elze.zdanceviciute@email.com" — '
    'svetainės kode rodo placeholder „example@email.com". Net tau norėdamas '
    'parašyti — žmogus negali.')

add_para(doc,
    '„Trečia — knygos tiražas yra 100 egz. Tai mums pakeitė viską: ši užduotis '
    'nėra apie šią knygą. Tai apie tavo karjeros bazę. 250–500 prenumeratorių — '
    'tai geresnė bazė antrai knygai, nei 5 000 atsitiktinių IG sekėjų."')

add_para(doc,
    '„Apačioje matai, kur tavo skaitytojo kelionė nutrūksta — žingsnis 3, kai jis '
    'nori „pasižymėti, neperka šiandien", ir žingsnis 6, kai jis perskaitė knygą '
    'ir nori sužinoti, kas toliau. Abu — neturi tiesioginio kanalo."')

# ---- 3:00 - 5:00 ----
add_h2(doc, '3:00–5:00 — Diagnostika (ANALYZE)')

add_quote_box(doc,
    'SKAIDRĖ 3 — trys kliūtys:\n'
    'A — Kanalo įėjimo nėra (subscribe forma neegzistuoja)\n'
    'B — Nėra aiškaus skaitytojo pažado\n'
    'C — Pirminio pasitikėjimo signalų trūkumas (placeholder paštas)\n\n'
    'SKAIDRĖ 4 — augimo prognozė:\n'
    '[Lentelė iš ANALYZE: Po 1/3/6/12 mėn., konservatyvi/tikėtina/optimistinė]\n'
    '+ Realus pavyzdys: Richie Billing → 100 → 400–500 per metus.')

add_para(doc,
    'KALBĖJIMAS (apie 220 ž.):', bold=True)

add_para(doc,
    '„Iš visų rastų defektų išskyrėme tris pagrindinius. A — nėra kanalo įėjimo. '
    'B — nėra skaitytojo pažado, kodėl jis turėtų grįžti. C — placeholder kontaktas '
    'mažina pasitikėjimą. Visi trys atsitvarko per pirmąjį mėnesį."')

add_para(doc,
    '„Dabar — kiek tai realiai gali pasiekti? Naudojome industrijos benchmark\'us — '
    'MailerLite, Brevo, AdMailr 2025–2026 m. duomenys, ne mūsų prielaidos. Vidutinis '
    'newsletter open rate šiandien — 31 iki 43 %. Welcome laiškų open rate yra 4× '
    'didesnis. Email skaitytojai knygas perka pastebimai dažniau nei socialinių '
    'tinklų sekėjai."')

add_para(doc,
    '„Mūsų prognozė — konservatyvi, tikėtina, optimistinė — pateikta lentelėje. '
    'Tikėtinas scenarijus po metų — 500 prenumeratorių. Tai nėra optimistinis '
    'paveikslas — tai tas pats skaičius, kurį pasiekė indie autorius Richie Billing '
    '(StoryOrigin atvejis): pradėjo nuo nulio, baigė metus su 400–500. Mūsų '
    'prielaidos yra tokios pačios kaip jo — nemokama lead magnet plius reguliarus '
    'ritmas."')

# ---- 5:00 - 8:00 ----
add_h2(doc, '5:00–8:00 — Pasiūlymai (IMPROVE)')

add_quote_box(doc,
    'SKAIDRĖ 5 — trys laiko horizontai:\n'
    '1 mėn. → „Kanalo įjungimas" — 7 val. vienkartiniai, 0 €\n'
    '3 mėn. → „Reguliarus ritmas" — 45 min./sav., 0 €\n'
    '12 mėn. → „Karjeros bazė" — tas pats + 30–60 €/m.\n\n'
    'SKAIDRĖ 6 — STARTER PACK:\n'
    '[Atvira fizinė knygelė / planšetė] '
    '5 paruošti welcome laiškai + 1 outreach šablonas + 3 subscribe copy variantai\n'
    '+ Instagram CTA paketas + lead magnet PDF struktūra')

add_para(doc,
    'KALBĖJIMAS (apie 320 ž., 2 min. 25 sek.):', bold=True)

add_para(doc,
    '„Pasiūlymai — trys, paskirstyti per laiką."')

add_para(doc,
    '„Pirmas, vieno mėnesio: „Kanalo įjungimas". Tai vieno savaitgalio darbas — '
    'maždaug 7 valandos. Sukuri MailerLite paskyrą — nemokamą, iki 500 prenumeratorių. '
    'Įdedi subscribe formą į 3 svetainės vietas. Paruoši 3 welcome laiškus — bet '
    'jų rašyti nereikia, mes juos parašėme, tu tik kopijuoji ir patikrini, ar tau '
    'skamba savai. Lead magnet — pirmas knygos skyrius. Pakeiti placeholder paštą '
    'į autore@elzesknygos.lt. Pirmojo mėnesio tikslas — 30–70 prenumeratorių.""')

add_para(doc,
    '„Antras, trijų mėnesių: „Reguliarus ritmas". Vienas laiškas kas dvi savaites, '
    'sekmadienio 19:00. Formatas — „Iš Elzės dirbtuvių": vienas sakinys iš tavo '
    'savaitės; ką šią savaitę skaitei ar klausei (pvz., Mačernio septintąją viziją); '
    'iš dirbtuvių — kažkas užkulisinio iš antros knygos rašymo; P. S. Kiekvienas '
    'laiškas — 30–40 minučių. Po 3 mėnesių — apie 100 prenumeratorių."')

add_para(doc,
    '„Trečias, dvylikos mėnesių: „Karjeros bazė". Tas pats ritmas plius 1–2 mokyklų '
    'vizitai, 1 signavimo turas (Marijampolė, Vilnius, Kaunas). Ir paskutinis akcentas — '
    'kai būsi pasirengus paskelbti antrąją knygą, pirmieji 3 dienos pirmumo skelbiamos '
    'TIK newsletter\'yje. Tai ne reklamos triukas — tai tikras pripažinimas tiems, '
    'kurie buvo su tavimi nuo pradžios. Po metų — apie 250 iki 500 prenumeratorių."')

add_para(doc,
    '„Tau perduodame ne tik šitą planą — perduodame STARTER PACK. Tai yra atskiras '
    'dokumentas, kuriame yra 5 paruošti welcome laiškai, vienas outreach šablonas '
    'knygų bloggeriams, trys subscribe formos copy variantai, Instagram bio ir '
    'postų tekstai, ir lead magnet PDF struktūra. Visa tai parašyta tavo balsu — '
    'rėmėmės tavo pačio autorės puslapio tonu (Camus, Mačernis, Kernagis). Jei kuri '
    'nors eilutė skamba ne tavoji — keiti."')

# ---- 8:00 - 9:00 ----
add_h2(doc, '8:00–9:00 — Matavimas ir „pull the plug" (VALIDATE)')

add_quote_box(doc,
    'SKAIDRĖ 7 — matavimo lentelė:\n'
    '5 metrikos, 3 laiko taškai (1, 3, 12 mėn.)\n'
    '+ Trys aiškūs „pull the plug" indikatoriai')

add_para(doc,
    'KALBĖJIMAS (apie 130 ž., 1 min.):', bold=True)

add_para(doc,
    '„Tau svarbu žinoti ne tik kada kanalas veikia, bet ir kada — sąžiningai — '
    'pasakyti „neveikia, einu kitur". Mūsų pareiga konsultantų — nupiešti šitą '
    'ribą iš anksto."')

add_para(doc,
    '„Matuoji penkias metrikas, 10 minučių per mėnesį, vieną Google Sheet eilutę. '
    'Trys „pull the plug" indikatoriai: jei po 6 sav. mažiau nei 20 prenumeratorių '
    'IR welcome open rate <30 % — kanalas nepraleidžia, eik prie BookTok. Jei per '
    '3 mėn. nesugebi paskelbti 4 laiškų — formatas tau per sunkus, mažiname iki '
    '1 per mėnesį arba pereiname į Substack. Jei per 6 mėn. open rate <20 % — '
    'klausiame skaitytojų, kas neveikia, ir keičiame turinį arba pripažįstame, '
    'kad balsui auditorija dar nesutelkta."')

# ---- 9:00 - 10:00 ----
add_h2(doc, '9:00–10:00 — „Pirmadienio rytas"')

add_quote_box(doc,
    'SKAIDRĖ 8 — pirmadienio veiksmų sąrašas (6 punktai, laiko žymėmis):\n'
    '07:30 — MailerLite paskyra\n'
    '08:00 — Welcome laiškai iš starter pack (kopijuojama)\n'
    '09:00 — Lead magnet PDF (1-as skyrius)\n'
    '10:30 — Subscribe forma 3 vietose svetainėje\n'
    '11:30 — Instagram bio CTA\n'
    '13:00 — 10 asmeninių pakvietimų\n'
    '~5–6 val. iš viso')

add_para(doc,
    'KALBĖJIMAS (apie 130 ž., 1 min.):', bold=True)

add_para(doc,
    '„Paskutinis dalykas, ir pats svarbiausias. Mes čia — du / trys studentai. Tu — '
    'viena. Visa, ką mes paruošėme, neturi vertės, jei pirmadienio rytą tu sėdėsi prie '
    'stalo ir nebus aišku, nuo ko pradėti. Todėl pažiūrėk į ekraną."')

add_para(doc,
    '„Šeši žingsniai, su laiko žymėmis. 07:30 — atidarai MailerLite, susikuri paskyrą. '
    '08:00 — atsidaromi mūsų starter pack, kopijuoji 3 welcome laiškus į automation. '
    '09:00 — paimi pirmąjį knygos skyrių, išsaugoji kaip PDF, įkėlei kaip priedą. '
    '10:30 — įdedi subscribe formą į 3 svetainės vietas. 11:30 — atnaujini Instagram '
    'bio. 13:00 — parašai 10 artimiems žmonėms su kvietimu pasirašyti. Iš viso — '
    'apie penkios–šešios valandos. Iki vakarienės kanalas, kurio šiandien nėra, '
    'jau veikia."')

add_para(doc,
    '„Mūsų prašymas tau, Elze: pasakyk dabar, ar šitas pirmadienis tau atrodo '
    'įmanomas. Jei kažkas atrodo per daug — keičiame čia pat."')

# ============================================================
# 3. TIKĖTINI ELZĖS KLAUSIMAI
# ============================================================
add_h1(doc, '3. Tikėtini Elzės klausimai ir mūsų atsakymai')

add_table_simple(
    doc,
    headers=['Galimas Elzės klausimas', 'Mūsų atsakymas'],
    rows=[
        ['„Aš negaliu rašyti laiško kas dvi savaites — gimnazija labai sunki."',
         '„Tada keičiame ritmą iš karto į vieną kartą per mėnesį. Open rate '
         'nesumažės smarkiai, nes prenumeratoriai vis tiek lauks. Geriau retai ir '
         'kokybiškai, nei dažnai ir paviršutiniškai. Starter pack vis tiek tinka — '
         'tik welcome serija."'],
        ['„Aš nemoku rinkodaros — nesijaučiu jaukiai prašyti žmonių pasirašyti."',
         '„Mes ir nemanome, kad tu turi mokėti rinkodaros. Tu moki rašyti — todėl '
         'paketo visi tekstai parašyti tavo, ne reklamos balsu. Pirmieji 30 '
         'prenumeratorių bus draugai ir šeima — niekam nereikia įtikinėti."'],
        ['„O jeigu po 3 mėn. niekas neprenumeruos?"',
         '„Tada turime aiškų „pull the plug" indikatorių (skyrius 5 dokumente). '
         'Jei po 6 sav. <20 prenum. ir welcome <30 % — pereiname į kitą kanalą '
         '(BookTok arba renginiai). Tai yra įsipareigojimas mums kaip konsultantams: '
         'mes nesitenkinsime „dar pamėginkim", o sąžiningai pripažinsime."'],
        ['„Ar negalėčiau tiesiog naudoti Substack?"',
         '„Galite — yra svarbus argumentas. Substack laisvas, turi „Notes" social '
         'discovery sluoksnį, nereikia įdėti formų į svetainę. Trūkumas — mažesnė '
         'integracija su tavo elzėsknygos.lt domenu. Mūsų rekomendacija — MailerLite '
         'pirmus 6 mėn., bet jei tau atrodo per sudėtinga, Substack yra normalus '
         'B planas."'],
        ['„Kaip užtikrinti, kad žmonės iš tikro skaitytų laiškus?"',
         '„Tu negali — gali tik užtikrinti, kad kiekvienas laiškas būtų vertas '
         'skaitymo. Industrijos vidurkis — 35 % atveria, 3 % paspaudžia. Jei tu '
         'pasieksi tuos skaičius, esi gerai. Jei ne — keisi formatą po 3 mėn."'],
        ['„Iš kur jūs žinote, kad mano tiražas — 100 egz.?"',
         '„Iš tavo svetainės — tai yra puslapyje „Knyga" pateiktas kolofonas, '
         'kuriame yra ISBN, leidėjas „Piko Valanda", 213 psl. ir tiražas. Mes '
         'tiesiog skaitėme."'],
    ],
    col_widths_cm=[6.5, 9.5],
)

# ============================================================
# 4. VAIDMENŲ PASISKIRSTYMAS
# ============================================================
add_h1(doc, '4. Vaidmenų pasiskirstymas grupėje')

add_para(doc,
    'Pristatymo metu kalba visi grupės nariai — taip rubrika vertina komandinį '
    'darbą. Pasiūlymas:')

add_table_simple(
    doc,
    headers=['Sek.', 'Laiko intervalas', 'Atsakingas', 'Tema'],
    rows=[
        ['1', '0:00–3:00', 'Narys A',
         'Vienas sakinys + esama būsena (OBSERVE)'],
        ['2', '3:00–5:00', 'Narys B',
         'Diagnostika ir prognozė (ANALYZE)'],
        ['3', '5:00–8:00', 'Narys A arba C',
         'Pasiūlymai + starter pack pristatymas'],
        ['4', '8:00–9:00', 'Narys B',
         'Matavimas ir „pull the plug"'],
        ['5', '9:00–10:00', 'Narys C arba A',
         'Pirmadienio rytas + atviras klausimas Elzei'],
    ],
    col_widths_cm=[1.5, 3.5, 3.5, 7.5],
)

add_para(doc,
    'Jei grupė — 2 asmenys (ne 3), Narys A imasi 1 ir 3 dalį (5 min.), Narys B — '
    '2, 4, 5 (5 min.).')

# ============================================================
# 5. CHEAT SHEET
# ============================================================
add_h1(doc, '5. „Cheat sheet" — kalbant be skaidrių')

add_para(doc,
    'Jeigu pristatymo metu nutrūktų projektorius arba reikėtų kalbėti be skaidrių, '
    'kiekvienas grupės narys turi atminti šešis pagrindinius skaičius:')

add_bullet(doc, 'svetainėje šiandien.', bold_prefix='0 formų — ')
add_bullet(doc, 'knygos tiražas. Šis projektas nėra apie ją — apie karjerą.',
           bold_prefix='100 egz. — ')
add_bullet(doc, 'tikėtinas prenumeratorių sk. po 3 mėn., 500 — po 12 mėn.',
           bold_prefix='100 ir 500 — ')
add_bullet(doc, 'newsletter open rate vidurkis. Welcome — 4× daugiau (~55–65 %).',
           bold_prefix='35 % — ')
add_bullet(doc, 'reali pirmo mėnesio darbo apimtis. 60 min./sav. — toliau.',
           bold_prefix='7 val. — ')
add_bullet(doc, 'kainuoja kanalas pirmus 6–9 mėn. (MailerLite free).',
           bold_prefix='0 € — ')

add_para(doc,
    'Visi šeši skaičiai — paremti faktais iš dokumentų. Jeigu Elzė klausia „iš kur '
    'tas skaičius?" — atsakymas yra šaltinių sąraše ANALYZE arba originaliame '
    'svetainės audite.')

# ============================================================
out = '/projects/sandbox/dokumentai/5_PRISTATYMAS_10_min_struktura.docx'
doc.save(out)
print(f'Sukurta: {out}')

import os
print(f'Dydis: {os.path.getsize(out)} baitų')
