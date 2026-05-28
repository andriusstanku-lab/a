"""STARTER PACK — paruošti laiškai, kuriuos Elzė kopijuoja į MailerLite tiesiogiai.

Visi tekstai parašyti remiantis Elzės pačios autorės puslapio tonu (Camus, Dostojevskis,
Mačernis, Kernagis), kad balsas išliktų jos, ne mūsų.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _docx_helper import (
    new_document, add_title, add_h1, add_h2, add_para, add_bullet,
    add_numbered, add_static_toc, add_table_simple, add_quote_box,
    add_pagebreak, add_header_block,
)
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH


doc = new_document()

add_header_block(
    doc,
    university='VILNIAUS UNIVERSITETAS',
    faculty='Ekonomikos ir verslo administravimo fakultetas',
    course='E. pardavimų grupinis projektas',
    project_title='STARTER PACK\nParuošti tekstai Elzės Email kanalo įjungimui',
    channel='Email / community building',
    group='[Grupės narių vardai, pavardės]',
    date='2026 m. gegužė',
)

add_static_toc(doc, [
    ('Įvadas — kaip naudoti šį paketą', 2),
    ('Dalis I. Welcome serija (5 laiškai)', 2),
    ('  E1. Welcome — „Sveika, atvėrei duris" (siunčiama iš karto)', 3),
    ('  E2. Knygos kontekstas — „Kodėl Camus epigrafa" (po 2 dienų)', 4),
    ('  E3. Iš dirbtuvių — „Kaip atrodo rašymo procesas" (po 5 dienų)', 5),
    ('  E4. Klausimas tau — „Kuri knyga tave pakeitė" (po 8 dienų)', 6),
    ('  E5. Pirmas reguliarus laiškas — „Iš Elzės dirbtuvių #1" (po 14 dienų)', 7),
    ('Dalis II. Outreach laiškas knygų bloggeriams', 8),
    ('Dalis III. Subscribe formos copy (3 variantai)', 9),
    ('Dalis IV. Instagram CTA paketas', 10),
    ('Dalis V. Lead magnet PDF — turinio struktūra', 11),
])

add_pagebreak(doc)

# ============================================================
# ĮVADAS
# ============================================================
add_h1(doc, 'Įvadas — kaip naudoti šį paketą')

add_para(doc,
    'Šis paketas yra ne planas. Tai paruošti tekstai, kuriuos Elzė gali kopijuoti '
    'tiesiogiai į MailerLite, į svetainę arba į Instagram. Visi laiškai parašyti su '
    'Elzės balsu — remiantis jos autorės puslapyje minėtais autoriais, muzika ir asmenine '
    'istorija. Jei kuris nors tekstas Elzei skamba „ne jos balsu", jį pakeičia ji '
    'pati — paketas yra atspirties taškas, ne galutinis variantas.')

add_para(doc,
    'Welcome serija paleidžiama automatiškai pagal MailerLite trigger\'į „naujas '
    'prenumeratorius". Po penkių laiškų pirmasis ciklas baigiasi; toliau Elzė siunčia '
    'įprastus „Iš Elzės dirbtuvių #N" laiškus pagal pasirinktą ritmą (rekomenduojama — '
    'kas 2 sav., sekmadienio 19:00).')

add_para(doc,
    'Visi tekstai turi vienodą struktūrą — antraštė (subjektas), pagrindinis tekstas, '
    'aiški P. S. eilutė (jei reikia). Subjekto eilutės parinktos taip, kad jos atrodytų '
    'kaip iš tikro žmogaus laiškas, ne iš rinkodaros sistemos.')

# ============================================================
# DALIS I — WELCOME SERIJA
# ============================================================
add_h1(doc, 'Dalis I. Welcome serija (5 laiškai)')

add_para(doc,
    'Visi welcome laiškai siunčiami iš `autore@elzesknygos.lt` su parašu „— Elzė". '
    'Vidutinis ilgis — 150–250 žodžių. Tonas — autorės puslapio tonas: paprastas, '
    'asmeniškas, nesistengia pamokyti.')

# ----- E1 -----
add_h2(doc, 'E1. Welcome laiškas (siunčiamas iš karto po pasirašymo)')

add_quote_box(doc, 'SUBJEKTAS: Sveika, atvėrei duris')

add_para(doc,
    'Sveika,')

add_para(doc,
    'rašau iš Marijampolės, ir pirmiausia — ačiū. Ne dėl etiketo, o iš tikrųjų: '
    'tai, kad palikai man savo el. paštą, reiškia, jog skiri minutę man tarp viso, '
    'ką šiandien gali skaityti. To nepriimu kaip savaime suprantamo dalyko.')

add_para(doc,
    'Šis laiškų ciklas yra mažas — penki laiškai per dvi savaites, o paskui kas dvi '
    'savaites — vienas trumpas „literatūrinis laiškas iš dirbtuvių". Niekada nesiūlysiu '
    'tau nieko pirkti dažniau nei kartą per kelis mėnesius. Tai yra mano savaitės '
    'kambario garsas, ne reklamos kanalas.')

add_para(doc,
    'Pridedu pirmuosius „Turtingas(is)" puslapius — kaip dovaną, kad galėtum susipažinti '
    'su Gabrieliumi, prieš nusprendžiant, ar nori knygą skaityti pilnai.')

add_para(doc, 'Iki kito laiško,')
add_para(doc, '— Elzė')

add_para(doc,
    'P. S. Jei kada nors panorėsi atsisakyti — viename laiško apačios mygtuke. Be aiškinimų. '
    'Tai sąžininga.')

# ----- E2 -----
add_h2(doc, 'E2. „Kodėl Camus epigrafa" (siunčiama po 2 dienų)')

add_quote_box(doc, 'SUBJEKTAS: Vienas sakinys, kurį parašiau ant pirmojo puslapio')

add_para(doc, 'Sveika dar kartą,')

add_para(doc,
    '„Au milieu de l\'hiver, j\'apprenais enfin qu\'il y avait en moi un été invincible." '
    'Tai Albert Camus sakinys, kurį pridėjau prieš pirmąjį „Turtingas(is)" puslapį. '
    'Lietuviškai: „Žiemos viduryje pagaliau supratau, kad manyje yra nenugalima vasara."')

add_para(doc,
    'Kai rašiau knygą, tas sakinys man buvo apie Gabrielių — pagrindinį herojų. '
    'Septyniolikmetis, kuris bando suprasti, kas iš tikrųjų svarbu, kai aplink atrodo '
    'beprasmiška. Bet rašydama supratau, kad tas sakinys yra ir apie mane — apie tą '
    'mergaitę, kuri pirmus ketverius metus mokykloje skiemenuodavo, o paskui pradėjo '
    'skaityti Dostojevskį, Kafką, Škėmą.')

add_para(doc,
    'Tai yra antras dalykas, kurį norėjau tau pasakyti — kad nė viena knyga, kurią '
    'parašysiu, nebus apie tobulą herojų. Visi mano žmonės bus žiemos viduryje. Ir, '
    'kaip Camus sako, juose vis tiek bus vasara.')

add_para(doc, '— Elzė')

# ----- E3 -----
add_h2(doc, 'E3. „Iš dirbtuvių — kaip atrodo rašymo procesas" (po 5 dienų)')

add_quote_box(doc, 'SUBJEKTAS: Pirmas knygos puslapis, kurį parašiau (ir kuris dingo)')

add_para(doc, 'Sveika,')

add_para(doc,
    'Šiandien — vienas užkulisinis dalykas, kurio Instagrame neminėjau. Pirmą „Turtingas(is)" '
    'puslapį parašiau būdama penkiolikos. Jis nepateko į knygą — buvo per daug '
    'patetiškas, per daug pamokomas, per daug bandantis būti svarbus. Trinau jį, '
    'paskui rašiau iš naujo, tada vėl trinau. Kažkur dvyliktame variante atsirado '
    'Gabrielius — toks, koks galiausiai gyvena knygoje.')

add_para(doc,
    'Tai bandau pasakyti, kad nė vienas tekstas, kurį skaitai — nei mano knygos, '
    'nei šio laiško — neatsiranda iš pirmo karto. Net šis sakinys, kurį dabar skaitai, '
    'yra trečiasis variantas. Pirmajame buvo „daug perrašiau"; antrajame — „rašyti '
    'man buvo sunku"; tik dabar atsirado tas, kuris nori tau ką nors pasakyti.')

add_para(doc,
    'Jeigu pati kažką rašai — atsakymas, ar verta toliau, dažnai yra ne pirmajame, '
    'bet trečiajame variante.')

add_para(doc, '— Elzė')

add_para(doc,
    'P. S. Šią savaitę klausausi V. Kernagio akustinio albumo. Rašydama dažnai '
    'klausiausi būtent jo — tai mano fonas, kuriame Gabrielius pradėjo kalbėti.')

# ----- E4 -----
add_h2(doc, 'E4. „Klausimas tau" (po 8 dienų)')

add_quote_box(doc, 'SUBJEKTAS: Kuri knyga tave pakeitė?')

add_para(doc, 'Sveika,')

add_para(doc,
    'Apsigalvojau ir nusprendžiau šiandien rašyti trumpą laišką — ne apie save. '
    'Apie tave.')

add_para(doc,
    'Mane pakeitė „Balta drobulė" — Antano Škėmos knyga. Pirmą kartą skaičiau ne dėl '
    'mokyklos, o dėl to, kad mama davė ir pasakė, kad iš jos pradėjo suprasti '
    'tylą. Aš tada to nesupratau — bet dabar suprantu.')

add_para(doc,
    'Norėčiau išgirsti tave: kuri knyga tau buvo tokia? Atsakyk tiesiog į šį '
    'laišką — eis tiesiogiai į mano dėžutę, niekas kitas nematys. Atsakymą '
    'galiu (jei sutinki) panaudoti viename iš ateities laiškų — '
    'anonimiškai, be tavo vardo. Bet tik tada, jei aiškiai sutiksi.')

add_para(doc,
    'Šis laiškas — paskutinis iš welcome serijos. Toliau gausi vieną laišką kas dvi '
    'savaites — sekmadienio vakare. Jeigu kuriuo nors metu ritmas tau bus per dažnas '
    'arba per retas, parašyk man tiesiogiai.')

add_para(doc, '— Elzė')

# ----- E5 (PIRMAS REGULIARUS) -----
add_h2(doc, 'E5. Pirmas reguliarus laiškas „Iš Elzės dirbtuvių #1" (po 14 dienų)')

add_quote_box(doc, 'SUBJEKTAS: Iš Elzės dirbtuvių #1 — apie tylą Mačernio eilėraštyje')

add_para(doc, 'Sveika,')

add_para(doc,
    '<b>Šią savaitę.</b> Pirmadienį Marijampolės bibliotekoje signavau dvi knygas. Du '
    'žmonės — daugiau, nei tikėjausi, ir tuo pačiu mažiau, nei kažkas iš šono '
    'pasakytų. Bet abu liko pasikalbėti ilgiau nei numatyta. Tas yra tikra dovana.')

add_para(doc,
    '<b>Šią savaitę skaičiau.</b> Vytauto Mačernio „Vizijų" rinkinį, septintąją viziją. '
    'Eilutė, kurioje jis sako: „Aš čia gyvas — taip aš čia esu" — ji kartojasi mintyse '
    'jau trečią dieną. Yra eilėraščių, kurie kalba apie tylą; šis pats yra tyla.')

add_para(doc,
    '<b>Iš dirbtuvių.</b> Pradėjau dirbti su antra knyga. Be detalių — bet personažas '
    'jau turi vardą, ir aš jau žinau, kuriame mieste jis gyvena. Tai daugiau, nei '
    'turėjau prieš mėnesį.')

add_para(doc,
    'P. S. Šią savaitę nieko nesiūlau. Sekmadienio vakaras — paprasčiausiai '
    'pažymėti.')

add_para(doc, '— Elzė')

add_para(doc,
    '[Pastaba grupei: <b>...</b> Word\'e galima paryškinti, kad tekstas atrodytų '
    'su pajuodintomis poskyrių antraštėmis. MailerLite editor\'iuje tai daroma '
    'tiesiog Bold mygtuku.]', italic=True)

# ============================================================
# DALIS II — OUTREACH
# ============================================================
add_pagebreak(doc)
add_h1(doc, 'Dalis II. Outreach laiškas knygų bloggeriams')

add_para(doc,
    'Šis laiškas siunčiamas iš `autore@elzesknygos.lt` 5–7 konkretiems Lietuvos '
    'knygų bloggeriams / BookTok kūrėjoms, kurie jau pristato YA arba lietuvių '
    'autorių kūrinius. Konkretus sąrašas (5–7 vardai) susidaromas pristatymo metu '
    'kartu su Elze — ji geriausiai žino, kurie kūrėjai jai svarbūs.')

add_h2(doc, 'O1. Outreach šablonas (su tinkamomis pakeitimo vietomis)')

add_quote_box(doc, 'SUBJEKTAS: [Vardas], dovana — pirma mano knyga')

add_para(doc, 'Sveika, [Vardas],')

add_para(doc,
    'Esu Elzė Zdancevičiūtė, šešiolikmetė iš Marijampolės. 2025 m. išleidau pirmąjį '
    'romaną „Turtingas(is)" (leidyklos „Piko Valanda", 213 p.) — istoriją apie '
    'septyniolikmetį Gabrielių, kuris ieško prasmės net tada, kai viskas atrodo '
    'beprasmiška.')

add_para(doc,
    'Sekiu tavo [kanalą / Instagram / TikTok / Goodreads — pasirinkti] jau kurį '
    'laiką ir ypač įsidėmėjau [konkretus pavyzdys: tavo apžvalgą apie X knygą / '
    'pokalbį apie Y autorę]. Tu kalbi apie knygas tokia kalba, kuri man pačiai '
    'svarbi.')

add_para(doc,
    'Norėčiau atsiųsti tau nemokamą egzempliorių — tiesiog dovaną, jokio mainomo '
    'įsipareigojimo. Jei knyga tau pasirodys verta paminėti — labai apsidžiaugčiau. '
    'Jei ne — tai irgi sąžiningas atsakymas.')

add_para(doc,
    'Jei sutinki, parašyk man savo paštomato arba namų adresą atgaliniu laišku, '
    'ir per kelias dienas knyga bus pas tave.')

add_para(doc, 'Iš anksto ačiū už tavo laiką,')
add_para(doc, '— Elzė')
add_para(doc, 'autore@elzesknygos.lt | https://elzėsknygos.lt | @elze_zdan')

add_h2(doc, 'O2. Konkrečių asmenų sąrašas — užpildoma kartu su Elze')

add_para(doc,
    'Pristatymo metu studentų grupė kartu su Elze sudaro sąrašą. Pavyzdinis formatas:')

add_table_simple(
    doc,
    headers=['#', 'Vardas / kanalas', 'Platforma', 'Kodėl tinka',
             'Susisiek. el. paštas / DM'],
    rows=[
        ['1', '[užpildoma]', '[IG / TikTok / Goodreads]', '[konkreti priežastis]', '[kontaktas]'],
        ['2', '', '', '', ''],
        ['3', '', '', '', ''],
        ['4', '', '', '', ''],
        ['5', '', '', '', ''],
        ['6', '', '', '', ''],
        ['7', '', '', '', ''],
    ],
    col_widths_cm=[0.8, 3.5, 2.2, 5.0, 4.5],
)

add_para(doc,
    'Studentai į tuščius langelius įveda 5–7 vardus po pasitarimo su Elze. Tai '
    'sąmoningai paliktas tuščias laukas — sąrašas turi būti realus, ne mūsų išgalvotas.')

# ============================================================
# DALIS III — SUBSCRIBE COPY
# ============================================================
add_pagebreak(doc)
add_h1(doc, 'Dalis III. Subscribe formos copy (3 variantai)')

add_para(doc,
    'Trys versijos to paties pakvietimo, skirtingoms svetainės vietoms.')

add_h2(doc, 'V1. Trumpa (1 sakinys) — apatinė puslapio juostą')

add_quote_box(doc,
    '„Vienas trumpas laiškas iš mano dirbtuvių, kas dvi savaites. Be reklamos."  '
    '[laukas: tavo el. paštas]   [mygtukas: Pasirašyti]')

add_h2(doc, 'V2. Vidutinė (3 sakiniai) — autorės puslapio pabaigoje')

add_quote_box(doc,
    '„Jei norėtum sekti, kas dirbtuvėse ir kas mano lentynoje — pasirašyk čia. '
    'Vienas trumpas laiškas kas dvi savaites: ką šiomis dienomis skaitau, klausau, '
    'rašau. Pirmąjį „Turtingas(is)" skyrių dovanoju iškart kaip ačiū."  '
    '[laukas]   [mygtukas: Pasirašyti]')

add_h2(doc, 'V3. Ilga (1 pastraipa) — atskiras „Newsletter" puslapis')

add_quote_box(doc,
    '„Nuo 2025 m. rašau trumpus literatūrinius laiškus — ne reklamą, o asmeninius '
    'pasvarstymus apie tai, kas mane šiomis dienomis kalbina: vienas eilėraštis, '
    'vienas albumas, viena pastraipa iš dirbtuvių. Pasiunčiu kas dvi savaites, '
    'sekmadienio vakare. Jei norėtum gauti — palik el. paštą, ir iškart atsiųsiu '
    'pirmąjį „Turtingas(is)" skyrių kaip ačiū. Atsisakyti gali bet kada — vienu '
    'paspaudimu apačioje. — Elzė"  [laukas]   [mygtukas: Pasirašyti]')

# ============================================================
# DALIS IV — INSTAGRAM CTA
# ============================================================
add_pagebreak(doc)
add_h1(doc, 'Dalis IV. Instagram CTA paketas')

add_h2(doc, 'I1. Bio versija')

add_quote_box(doc,
    'Elzė Zdancevičiūtė\n'
    'Knyga „Turtingas(is)" — Piko Valanda, 2025\n'
    'Marijampolė\n'
    '↓ Newsletter (kas 2 sav.)\n'
    '[link: https://elzėsknygos.lt/newsletter]')

add_h2(doc, 'I2. Statinis postas (kvadratas)')

add_quote_box(doc,
    '[Vaizdas: Elzė rašo prie stalo, knyga šalia, šviesa pro langą]\n\n'
    'CAPTION:\n'
    '„Pradedu rašyti vieną trumpą laišką kas dvi savaites — apie tai, kas dirbtuvėse, '
    'ką skaitau, ką klausau. Be reklamos.\n\n'
    'Jeigu nori gauti, prenumeruoja per nuorodą bio (↑). Pirmąjį „Turtingas(is)" '
    'skyrių dovanoju iškart.\n\n'
    'Sekmadienio 19:00. Tas pats laikas, kiekvieną kartą."')

add_h2(doc, 'I3. Stories serija (3 ekranai)')

add_quote_box(doc,
    'EKRANAS 1:\n'
    '[Tekstas baltame fone] „Pradedu newsletter\'į. Kuklus."\n\n'
    'EKRANAS 2:\n'
    '[Tekstas + nuoroda swipe up] „Vienas laiškas kas dvi savaites. Skaitau, '
    'klausausi, rašau. Be reklamos."\n\n'
    'EKRANAS 3:\n'
    '[Tekstas + sticker su klausimu] „Ar pasirašytum? ↓ link bio"')

add_h2(doc, 'I4. Reels scenarijus (15–20 sek.)')

add_quote_box(doc,
    'KADRAS 1 (3 sek.): Elzė kameroje, tyliai sako: „Niekada nemokėjau rinkodaros."\n\n'
    'KADRAS 2 (5 sek.): „Bet aš mokėsiu rašyti laiškus. Po vieną. Kas dvi savaites."\n\n'
    'KADRAS 3 (5 sek.): [Ekrane atsiranda subscribe forma] „Jei nori gauti, '
    'pasirašyk per nuorodą bio."\n\n'
    'KADRAS 4 (3 sek.): [Atgal į veidą] „Sekmadienio vakaras. Tas pats laikas."')

# ============================================================
# DALIS V — LEAD MAGNET PDF
# ============================================================
add_pagebreak(doc)
add_h1(doc, 'Dalis V. Lead magnet PDF — turinio struktūra')

add_para(doc,
    'Lead magnet — nemokama dovana mainais į el. paštą — Elzės atveju yra pirmasis '
    'knygos „Turtingas(is)" skyrius. Konkrečiai šis paketas reikalauja iš Elzės '
    'paimti rankraštį (Microsoft Word arba InDesign formatu, kuriame jau dirbo '
    'maketuotoja B. Strolytė) ir išskirti pirmuosius ~15 puslapių.')

add_h2(doc, 'Lead magnet PDF struktūra')

add_table_simple(
    doc,
    headers=['Puslapis', 'Turinys'],
    rows=[
        ['1 (titulinis)',
         '„TURTINGAS(IS) — pirmieji puslapiai. Ačiū, kad atvėrei. — Elzė"'],
        ['2',
         'Camus epigrafa (lietuviškai ir prancūziškai), kaip yra knygoje'],
        ['3–14',
         'Pirmas knygos skyrius (originalas iš leidyklos „Piko Valanda" maketo)'],
        ['15 (paskutinis)',
         '„Tikiuosi, Gabrielius tave palydėjo iki čia. Jeigu nori toliau — knygą '
         'gali įsigyti pas mane tiesiogiai (autore@elzesknygos.lt) arba '
         '[knygos.lt / Patogu Pirkti — užpildoma]. Iki kito laiško sekmadienio '
         'vakare. — Elzė"'],
    ],
    col_widths_cm=[2.5, 13.5],
)

add_para(doc,
    'PDF generavimas — paprastas: Microsoft Word „Save As → PDF". Elzė šį PDF įkelia '
    'į MailerLite kaip welcome laiško E1 priedą; per ten pirmas laiškas automatiškai '
    'pristato lead magnet kiekvienam naujam prenumeratoriui.')

add_h2(doc, 'Pastaba apie autorines teises')

add_para(doc,
    'Pirmas knygos skyrius yra Elzės nuosavybė kaip autorės. Leidykla „Piko Valanda", '
    'pagal įprastą Lietuvos leidybos praktiką, neturi išimtinių teisių į ištraukas '
    'rinkodaros tikslams. Vis dėlto rekomenduojame Elzei trumpai (vienu el. laišku) '
    'informuoti leidyklą, kad ištrauka bus naudojama lead magnet pavidalu — tai '
    'profesionalu ir uždaro bet kokius klausimus iš anksto.')

# ============================================================
out = '/projects/sandbox/dokumentai/STARTER_PACK_Email_kanalo_ijungimas.docx'
doc.save(out)
print(f'Sukurta: {out}')

import os
print(f'Dydis: {os.path.getsize(out)} baitų')
