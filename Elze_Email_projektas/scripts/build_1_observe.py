"""1 etapas — OBSERVE: Esama būsena (3 psl. dokumentas).

Klausimas: Kaip Email / community building kanalas ŠIANDIEN veikia Elzei
(jei iš viso veikia)?
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _docx_helper import (
    new_document, add_title, add_h1, add_h2, add_para, add_bullet,
    add_numbered, add_static_toc, add_table_simple, add_quote_box,
    add_pagebreak, add_header_block,
)

doc = new_document()

# ---- Tituliniai duomenys ----
add_header_block(
    doc,
    university='VILNIAUS UNIVERSITETAS',
    faculty='Ekonomikos ir verslo administravimo fakultetas',
    course='E. pardavimų grupinis projektas',
    project_title='1 etapas — OBSERVE\nElzės Zdancevičiūtės knygos „Turtingas(is)" Kaizen auditas',
    channel='Email / community building',
    group='[Grupės narių vardai, pavardės]',
    date='2026 m. gegužė',
)

# ---- TURINYS (statinis) ----
add_static_toc(doc, [
    ('1. Tikslas ir metodas', 2),
    ('2. Esama Elzės pozicija Email kanale', 2),
    ('3. Lyginamųjų autorių analizė tame pačiame kanale', 3),
    ('4. Vartotojo („skaitytojo") kelionės žemėlapis', 3),
    ('5. Friction points — kur kelionė nutrūksta', 4),
    ('6. Apibendrinimas', 4),
    ('Priedas A. Screenshot\'ų sąrašas', 4),
])

add_pagebreak(doc)

# ============================================================
# 1. TIKSLAS IR METODAS
# ============================================================
add_h1(doc, '1. Tikslas ir metodas')

add_para(doc,
    'Šio etapo tikslas — be vertinimo ir be prielaidų užfiksuoti, kaip Email ir community '
    'building kanalas Elzės Zdancevičiūtės atveju veikia šiandien. Klausimas yra kuklus, bet '
    'esminis: ar potencialus skaitytojas, susidomėjęs Elze ar jos romanu „Turtingas(is)", '
    'turi būdą palaikyti su autore tęstinį ryšį, neperleisdamas savo dėmesio Instagram, '
    'TikTok ar Facebook algoritmui? Email yra vienintelis kanalas, kurio pati platforma '
    'negali iš autoriaus „atimti" — todėl jis ir yra atskiro audito vertas.')

add_para(doc,
    'Metodas — tiesioginis kliento turto patikrinimas „kaip vartotojas". Tikrinome '
    'oficialią svetainę elzėsknygos.lt, viešus socialinių tinklų profilius, ėjome per '
    'tipinius skaitytojo žingsnius (atradimas → svetainė → bandymas pirkti → po pirkimo) '
    'ir fiksavome, kur Elzė turi galimybę paimti skaitytojo el. paštą, o kur jį '
    'praleidžia. Visi pastebėjimai — iš realios svetainės, ne iš prielaidų.')

# ============================================================
# 2. ESAMA ELZĖS POZICIJA EMAIL KANALE
# ============================================================
add_h1(doc, '2. Esama Elzės pozicija Email kanale')

add_h2(doc, '2.1. Svetainė elzėsknygos.lt')

add_para(doc,
    'Svetainė sukurta WordPress pagrindu (LiteSpeed cache infrastruktūra), turi tris '
    'puslapius: pagrindinį, „Knyga" ir „Autorė". Visa naršymo struktūra paprasta ir '
    'aiški. Tačiau viso viešo svetainės HTML kodo automatinė analizė rodo:')

add_bullet(doc, '0 (nulis) HTML formų visoje svetainėje;', bold_prefix='')
add_bullet(doc, '0 įvedimo laukų, kurių tipas yra „email";', bold_prefix='')
add_bullet(doc,
    'jokios užuominos apie „naujienlaiškį", „prenumeratą", „naujienoms", '
    'MailerLite, Mailchimp, Substack ar kitą email-rinkimo įrankį (nei kode, nei tekste);',
    bold_prefix='')
add_bullet(doc,
    'nėra jokio pop-up, exit-intent, sidebar widget\'o ar inline CTA, kuris kviestų '
    'palikti el. pašto adresą.',
    bold_prefix='')

add_para(doc,
    'Tai reiškia, kad Email kanalo, kaip turto, šiandien neegzistuoja: net jeigu '
    'svetainę aplankys 1 000 žmonių, Elzei neliks jokios žinios, kad jie buvo, ir nebus '
    'būdo prie jų sugrįžti. Skaitytojas yra anoniminis vienkartinis lankytojas.')

add_h2(doc, '2.2. Kontaktas „Susiekite su Elze" puslapyje')

add_para(doc,
    'Pagrindinio puslapio sekcijoje „Susiekite su Elze" nurodyti trys kontaktai. Du iš jų '
    '— Facebook ir Instagram — yra socialinių tinklų sąsajos. Trečiasis, el. paštas, '
    'svetainės kode rodo placeholder reikšmę („mailto:example@email.com"), o '
    'matomame tekste — adresą su domenu „@email.com", kuris nėra realus Elzės '
    'pašto domenas. Tai antrasis kritinis radinys: net tiems, kurie nori parašyti '
    'autorei tiesiogiai, kelias yra užkliuvęs.')

add_h2(doc, '2.3. Pagrindinis CTA „ĮSIGYTI KNYGĄ"')

add_para(doc,
    'Pagrindinis svetainės kvietimas — mygtukas „ĮSIGYTI KNYGĄ". Svetainės kode nematyti '
    'WooCommerce, Stripe, Paysera ar kitos pirkimo infrastruktūros. Tai reiškia, kad '
    'pati svetainė knygos neparduoda; mygtukas greičiausiai veda į išorinę parduotuvę '
    '(pvz., knygos.lt, Patogu Pirkti) arba į autorės žinutę. Šio žingsnio realią '
    'eigą reikia patikrinti kliento susitikimo metu (žr. 4 skyrių).')

add_h2(doc, '2.4. Tiražas ir konteksto svarba')

add_para(doc,
    'Knygos „Turtingas(is)" tiražas yra 100 egzempliorių (šaltinis — kolofonas svetainės '
    'puslapyje „Knyga"). Tai svarbus skaičius, kuris keičia projekto tikslą: Email '
    'kanalas šiame etape nėra tik priemonė parduoti dabartinę knygą, — jis yra ilgalaikis '
    'autorinės karjeros turtas. Net jei visa 100 egz. partija būtų išparduota, autorinis '
    'augimas priklausys nuo to, ar Elzė ateityje turės bazę žmonių, į kuriuos gali '
    'tiesiogiai kreiptis dėl antros knygos, signavimų, viešų skaitymų.')

# ============================================================
# 3. LYGINAMIEJI AUTORIAI
# ============================================================
add_h1(doc, '3. Lyginamųjų autorių analizė tame pačiame kanale')

add_para(doc,
    'Pasirinkome tris autorius, kurių situacija yra panaši (jaunas balsas, savilaida arba '
    'maža nepriklausoma leidykla, debiutinė knyga arba ankstyva karjera) ir kurie aktyviai '
    'naudoja Email / community building kanalą. Tikslas — ne kopijuoti, o suprasti, kaip '
    'kanalas atrodo, kai jis veikia.')

add_table_simple(
    doc,
    headers=['Autorius / projektas', 'Kanalo forma', 'Ką galima pasimokyti'],
    rows=[
        ['Tautvydas Miliūnas — „Notes from Lithuania" (Substack)',
         'Substack newsletter, lietuvio rašytojo balsas anglų k.',
         'Asmeninis tonas + reguliarus ritmas; subscribe forma matoma kiekviename įraše.'],
        ['Lukas Barbier — „The Vilnius Observer" (Substack)',
         'Reguliarus newsletter apie Lietuvą / Baltiją',
         'Aiškus pažadas skaitytojui („kas savaitę gausi…"); be lead magnet — vien '
         'reguliarumas ir tema.'],
        ['Indie autoriaus atvejis (Richie Billing, StoryOrigin case study)',
         'Reader magnet (nemokama trumpa istorija) + newsletter swaps',
         'Iš 0 → 100+ prenumeratorių per pirmą kampaniją; vėliau augo iki 400–500. '
         'Rodo, kad debiutiniam autoriui įmanoma sutelkti pirmą branduolį.'],
    ],
    col_widths_cm=[4.5, 4.5, 7.0],
)

add_para(doc,
    'Bendras dėsningumas: visi trys turi (a) vieną aiškų subscribe pažadą, (b) viešai '
    'matomą pasirašymo formą savo pagrindiniame kanale, (c) reguliarų — net jei retą — '
    'kontakto ritmą. Nė vienas iš trijų nestato strategijos ant viralinio momento; '
    'visi stato ant ramaus, prognozuojamo grįžimo.')

# ============================================================
# 4. VARTOTOJO KELIONĖS ŽEMĖLAPIS
# ============================================================
add_h1(doc, '4. Vartotojo („skaitytojo") kelionės žemėlapis')

add_para(doc,
    'Atstatėme tipinę kelionę žmogaus, kuris ŠIANDIEN nepažįsta Elzės, bet GALĖTŲ '
    'tapti jos skaitytoju per Email kanalą. Kelionė turi šešis žingsnius; prie '
    'kiekvieno fiksuojama, kas vyksta dabar.')

add_table_simple(
    doc,
    headers=['Žingsnis', 'Ką žmogus daro', 'Ką randa šiandien'],
    rows=[
        ['1. Atradimas',
         'Pamato Elzę Instagram įraše, draugo poste arba straipsnyje',
         'Mato pavadinimą „Turtingas(is)" ir nuorodą į elzėsknygos.lt'],
        ['2. Pirmas svetainės apsilankymas',
         'Perskaito apie knygą, pažiūri „Autorė" puslapį',
         'Gauna autentišką, gerai parašytą Elzės pasakojimą'],
        ['3. Sprendimo momentas',
         'Galvoja: „Įdomu, bet šiandien neperku"',
         'NĖRA jokios alternatyvos „pirkti dabar" — žmogus išeina be pėdsako'],
        ['4. Bandymas pirkti',
         'Spaudžia „ĮSIGYTI KNYGĄ"',
         'Patikrintinas kliento susitikimo metu (kur veda mygtukas?)'],
        ['5. Po pirkimo',
         'Knyga atvyksta paštu / pasiima',
         'Ar gauna padėkos laišką? Ar prašoma palikti atsiliepimą? Patikrinti.'],
        ['6. Po skaitymo',
         'Nori sužinoti, kas toliau',
         'Vienintelis būdas — sekti Instagram. Algoritmas sprendžia, ar parodys.'],
    ],
    col_widths_cm=[2.5, 5.0, 8.5],
)

add_para(doc,
    'Esminis kelionės defektas — 3 ir 6 žingsniai. Trečiajame žingsnyje žmogus, kuris '
    'jau paliko savo dėmesį Elzės puslapyje, neturi būdo „pasižymėti" — todėl, kad ir '
    'koks geras būtų svetainės turinys, lankytojas dingsta. Šeštajame žingsnyje '
    'jau pirkęs ir perskaitęs žmogus — t. y. tikras Elzės skaitytojas — taip pat '
    'neturi tiesioginio kanalo gauti naujienas apie ateities kūrybą, todėl ryšys '
    'visiškai priklauso nuo socialinių tinklų algoritmų.')

# ============================================================
# 5. FRICTION POINTS
# ============================================================
add_h1(doc, '5. Friction points — kur kelionė nutrūksta')

add_para(doc,
    'Apibendrinome konkrečius momentus, kuriais skaitytojo kelionė nutrūksta arba '
    'atsiranda nereikalinga trintis. Pateikiame juos rūšiuotus pagal įtaką, ne pagal '
    'sudėtingumą.')

add_numbered(doc, 1,
    'visoje svetainėje nėra nė vienos prenumeratos formos, lauko ar nuorodos. '
    'Skaitytojas, kuris jau yra svetainėje, nepalieka jokio pėdsako.',
    bold_prefix='„Tylioji" išėjimo problema — ')

add_numbered(doc, 2,
    'matoma reikšmė rodo `elze.zdanceviciute@email.com`, o HTML kode — '
    '`mailto:example@email.com`. Net norintis parašyti — negali.',
    bold_prefix='Sugedęs kontaktinis el. paštas — ')

add_numbered(doc, 3,
    'pagrindinis CTA neturi alternatyvos „dabar dar neperku, bet noriu sekti". '
    'Realybėje didžioji dauguma pirmųjų lankytojų nepiks per pirmą apsilankymą — '
    'tai industrijos standartas, ne Elzės problema.',
    bold_prefix='Vienakelis CTA — ')

add_numbered(doc, 4,
    'pagrindinio puslapio tekste pavadinime parašyta „EIzę" (vietoje „Elzę") — '
    'tai labai smulki, bet realiai matomos klaidos pavyzdys. Verta paminėti, nes '
    'tai paprastas, nemokamas patobulinimas.',
    bold_prefix='Smulki, bet matoma rašybos klaida — ')

add_numbered(doc, 5,
    'autorės puslapyje yra labai stiprus, asmeninis tekstas (Camus, Dostojevskis, '
    'Kernagis, Mačernis), bet šis turinys nedirba kanalui — niekur nesiūloma '
    'gauti panašaus formato pasakojimų ateityje.',
    bold_prefix='Ne dirbantis turinys — ')

# ============================================================
# 6. APIBENDRINIMAS
# ============================================================
add_h1(doc, '6. Apibendrinimas')

add_para(doc,
    'Vienas sakinys: Email / community building kanalo Elzės atveju šiandien faktiškai '
    'nėra — yra svetainė, kuri pristato knygą, bet neleidžia paimti skaitytojo, ir yra '
    'socialiniai tinklai, kurių lankytojas Elzei nepriklauso. Tai reiškia, kad bet '
    'koks šio kanalo veiksmas pradedamas nuo nulio, o tai yra ir gera žinia: nėra '
    'paveldėtų klaidų, kurios trukdytų statyti tinkamai.')

add_para(doc,
    'Kitame etape (ANALYZE) kelsime klausimą — koks yra realistinis šio kanalo dydis '
    'per 3, 6 ir 12 mėn., kiek tai kainuoja Elzei laiku ir pinigais, ir kurios trys '
    'kliūtys yra svarbiausios, kad jas atblokuotume.')

# ============================================================
# PRIEDAS A — SCREENSHOT'AI
# ============================================================
add_h1(doc, 'Priedas A. Screenshot\'ų sąrašas')

add_para(doc,
    'Žemiau išvardyti screenshot\'ai, kuriuos pridedame prie šio dokumento atskiruose '
    'failuose (priedas A1–A6). Visi padaryti 2026 m. gegužės mėn., naudojant inkognito '
    'naršyklės langą, kad rezultatas nebūtų asmeninių algoritmų paveiktas.')

add_bullet(doc, 'Pagrindinio puslapio elzėsknygos.lt viršus su CTA „ĮSIGYTI KNYGĄ".',
           bold_prefix='A1 — ')
add_bullet(doc, 'Pagrindinio puslapio sekcija „Susiekite su Elze" su matoma '
           'placeholder el. pašto reikšme.', bold_prefix='A2 — ')
add_bullet(doc, 'Visas „Knyga" puslapis su kolofonu (tiražas 100 egz., ISBN, leidėjas).',
           bold_prefix='A3 — ')
add_bullet(doc, 'Visas „Autorė" puslapis (kaip turinio pavyzdys ateities newsletter\'iui).',
           bold_prefix='A4 — ')
add_bullet(doc, 'Naršyklės „View Source" rodantis HTML kodą be jokios <form> arba '
           '<input type="email"> elemento.', bold_prefix='A5 — ')
add_bullet(doc, 'Instagram profilio @elze_zdan ekrano nuotrauka su sekėjų skaičiumi '
           '(taikymo data — pirmas susitikimas su Elze).',
           bold_prefix='A6 — ')

# ============================================================
# Saugojimas
# ============================================================
out = '/projects/sandbox/dokumentai/1_OBSERVE_Esama_busena.docx'
doc.save(out)
print(f'Sukurta: {out}')

# Patikrinam dydį
import os
print(f'Dydis: {os.path.getsize(out)} baitų')
