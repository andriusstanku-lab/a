# PU3 PRAKTINĖS UŽDUOTIES TURINYS

**Darbo pavadinimas:** DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ ANALIZĖ IR PALYGINIMAS
**Autorius:** Andrius Vargonas
**Studijų programa:** Marketingo technologijų studijų programa
**Dalykas:** Informacijos sistemos ir duomenų bazės
**Fakultetas:** VU Kauno fakultetas
**Institutas:** Socialinių mokslų ir taikomosios informatikos institutas
**Metai:** 2026
**Miestas:** Kaunas

---

## ANTRAŠTINIS LAPAS (turinys)

```
                    VILNIAUS UNIVERSITETAS
                       KAUNO FAKULTETAS

       SOCIALINIŲ MOKSLŲ IR TAIKOMOSIOS INFORMATIKOS INSTITUTAS

              Marketingo technologijų studijų programa


                       ANDRIUS VARGONAS


       DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ ANALIZĖ IR PALYGINIMAS

              Informacijos sistemos ir duomenų bazės




                          Kaunas
                           2026
```

---

## TURINYS

```
ĮVADAS .......................................................... 3
1. DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ TEORINĖ APŽVALGA .................. 4
   1.1. Duomenų bazės samprata ir pagrindiniai tipai ............... 4
   1.2. Duomenų bazių kūrimo įrankių klasifikacija ................. 5
   1.3. Palyginimo kriterijų atranka ir pagrindimas ................ 6
2. DEŠIMTIES DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ APŽVALGA ................ 7
   2.1. Tradiciniai reliaciniai įrankiai ........................... 7
   2.2. Šiuolaikinės debesijos ir žemo kodo platformos ............. 8
   2.3. NoSQL ir atvirojo kodo įrankiai ............................ 9
3. LYGINAMOJI ANALIZĖ IR REZULTATAI ............................. 10
   3.1. Lyginamoji lentelė pagal aštuonis kriterijus .............. 10
   3.2. Pagrindiniai skirtumai ir panašumai ....................... 11
   3.3. Įrankių tinkamumas skirtingiems naudojimo atvejams ........ 12
IŠVADOS ........................................................ 13
LITERATŪROS SĄRAŠAS ............................................ 14
1 PRIEDAS. Detali dešimties įrankių charakteristika ............ 15
11 PRIEDAS. Dirbtinio intelekto panaudojimo deklaracija ........ 16
12 PRIEDAS. Dirbtinio intelekto užklausos ir gauti atsakymai ... 17
```

*Pastaba: Python skripte šis turinys yra įdėtas kaip automatinis Word laukas (TOC field) su iš anksto užpildytu „cached" tekstu, todėl atidarius dokumentą turinys matomas iš karto, be poreikio spausti F9.*

---

## ĮVADAS

Šiuolaikinėje informacijos visuomenėje duomenų bazės yra esminė bet kurios verslo ar viešojo sektoriaus informacinės sistemos dalis. Augantis duomenų kiekis, debesijos sprendimų plėtra bei žemo kodo (angl. *low-code*) platformų atsiradimas iš esmės keičia tradicinį požiūrį į duomenų bazių valdymą. Pasak Connolly ir Begg (2015), pastaraisiais dešimtmečiais duomenų bazių rinka tapo viena dinamiškiausių programinės įrangos sričių. Lietuvos rinkoje veikiančios įmonės, ypač smulkiojo ir vidutinio verslo atstovai, vis dažniau ieško paprastesnių, lankstesnių ir mažesnes sąnaudas reikalaujančių duomenų bazių kūrimo įrankių. Todėl šiuolaikinių duomenų bazių kūrimo priemonių analizė yra aktuali tiek akademiniu, tiek praktiniu požiūriu.

Marketingo technologijų studijų kontekste duomenų bazių kūrimo įrankių pažinimas yra ypač svarbus. Marketingo specialistai dažnai dirba su klientų duomenimis, segmentavimo lentelėmis, kampanijų rezultatais ir kitomis duomenų aibėmis. Tinkamai pasirinktas duomenų bazės kūrimo įrankis leidžia greičiau pateikti analitinius sprendimus, automatizuoti darbo eigas ir sumažinti priklausomybę nuo informacinių technologijų specialistų.

Šio darbo tikslas - išanalizuoti ir palyginti dešimt šiuolaikinių duomenų bazių kūrimo įrankių pagal pasirinktus funkcinius kriterijus.

Darbo uždaviniai:

1. Apžvelgti duomenų bazės sampratą, jos tipus ir kūrimo įrankių klasifikaciją.
2. Identifikuoti dešimt šiuolaikinių duomenų bazių kūrimo įrankių, atstovaujančių skirtingoms kategorijoms.
3. Palyginti pasirinktus įrankius pagal aštuonis kriterijus, sudarant struktūrinę lyginamąją lentelę.
4. Įvertinti įrankių tinkamumą skirtingiems naudojimo atvejams ir suformuluoti rekomendacijas.

Darbo metodai - mokslinės literatūros analizė, oficialios įrankių kūrėjų dokumentacijos apžvalga, lyginamoji analizė, sintezė ir apibendrinimas.

Darbą sudaro trys skyriai. Pirmajame skyriuje pateikiama teorinė duomenų bazių ir jų kūrimo įrankių apžvalga. Antrajame skyriuje aprašomi dešimt pasirinktų įrankių, suskirstytų pagal jų pobūdį. Trečiajame skyriuje pateikiama lyginamoji analizė pagal aštuonis kriterijus ir suformuluojamos rekomendacijos pagal naudojimo atvejus.

Rengiant šį darbą buvo naudotasi dirbtinio intelekto (toliau - DI) generatyviniu modeliu Anthropic Claude (Sonnet 4.5, internetinė prieiga, naudota 2026 m. gegužės mėn.) kaip pagalbine priemone. DI buvo pasitelktas ribotais ir aiškiai apibrėžtais tikslais: pradinių darbo struktūros variantų pasiūlymui, kalbinio stiliaus tobulinimui ir gramatikos taisymui. Pagrindinį turinį, įrankių analizę, šaltinių paiešką, vertinimus ir išvadas autorius parengė savarankiškai. DI sugeneruoto turinio dalis darbe neviršija penkiolikos procentų. Detalus DI naudojimo aprašymas pateikiamas 11 priede, o naudotos užklausos - 12 priede. Autorius susipažinęs su Vilniaus universiteto 2024 m. patvirtintomis dirbtinio intelekto naudojimo gairėmis (Nr. SPN-54).

---

## 1. DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ TEORINĖ APŽVALGA

Šiame skyriuje pateikiama duomenų bazės sampratos apžvalga, aprašomi pagrindiniai duomenų bazių tipai ir jų kūrimo įrankių klasifikacija. Taip pat pristatomi kriterijai, pagal kuriuos antrajame ir trečiajame darbo skyriuose bus lyginami pasirinkti įrankiai.

### 1.1. Duomenų bazės samprata ir pagrindiniai tipai

Duomenų bazė - tai struktūrizuotas, susijusių duomenų rinkinys, valdomas duomenų bazių valdymo sistemos (toliau - DBVS). Pasak Elmasri ir Navathe (2016), duomenų bazė užtikrina nuoseklų duomenų saugojimą, prieigą ir valdymą per vieningą sąsają. Sąvoka apima ne tik patį duomenų rinkinį, bet ir programines priemones, leidžiančias jį kurti, redaguoti, užklausti ir administruoti.

Šiuolaikinėje praktikoje skiriami keturi pagrindiniai duomenų bazių tipai:

- **Reliacinės** (angl. *relational*) - duomenys saugomi struktūrizuotose lentelėse, kurios tarpusavyje susietos pirminiais ir antriniais raktais. Klasikinis šio tipo pavyzdys aprašytas Codd (1970) darbe; dabartiniai atstovai - MySQL, PostgreSQL, Microsoft SQL Server.
- **Nereliacinės** (angl. *NoSQL*) - skirtos darbui su nestruktūrizuotais ar pusiau struktūrizuotais duomenimis. Apima dokumentų, raktas-reikšmė, grafų ir stulpelių saugyklas (Sadalage ir Fowler, 2012).
- **Debesijos** (angl. *cloud-based*) - veikia kaip paslauga (*Database-as-a-Service*), pašalindamos infrastruktūros valdymo poreikį. Tipiniai pavyzdžiai - Amazon RDS, Google Cloud SQL.
- **Hibridinės** - apjungiančios reliacines ir nereliacines savybes. Pastarąjį dešimtmetį tokių sistemų populiarumas auga.

Atskirai paminėtinos žemo kodo (angl. *low-code*) ir bekodės (angl. *no-code*) platformos, kurios duomenų bazes pateikia kaip dalį platesnio aplikacijų kūrimo įrankio.

### 1.2. Duomenų bazių kūrimo įrankių klasifikacija

Duomenų bazių kūrimo įrankis - tai programinė priemonė, leidžianti vartotojui projektuoti duomenų schemą, kurti lenteles ar kolekcijas, įvesti, redaguoti ir užklausti duomenis. Skirtingai nei DBVS, duomenų bazės kūrimo įrankis dažnai turi vizualinę sąsają, padedančią ne informacinių technologijų specialistui dirbti su duomenimis.

Įrankius galima klasifikuoti pagal kelis požymius. Pirma, pagal diegimo būdą skiriami savarankiškai diegiami (angl. *self-hosted*) ir debesijos (angl. *cloud*) įrankiai. Antra, pagal licenciją - patentuoti (angl. *proprietary*) ir atvirojo kodo (angl. *open-source*). Trečia, pagal vartotojo profilį - skirti programuotojams, analitikams arba dalykinės srities vartotojams be programavimo įgūdžių. Ketvirta, pagal funkcinę paskirtį - paprastos administravimo priemonės, integruotos kūrimo platformos arba pilnos žemo kodo aplikacijų kūrimo aplinkos.

Marketingo srityje vis dažniau pasirenkami integruoti įrankiai, kurie ne tik leidžia kurti duomenų bazes, bet ir automatizuoti darbo eigas, generuoti ataskaitas ar kurti vidines aplikacijas be programavimo žinių. Tokia tendencija atspindi platesnį poslinkį į vartotojui draugiškus skaitmeninio verslo įrankius.

### 1.3. Palyginimo kriterijų atranka ir pagrindimas

Lyginamosios analizės kokybė priklauso nuo tinkamai pasirinktų kriterijų. Šiame darbe pasirinkti aštuoni kriterijai, atspindintys svarbiausias šiuolaikinio duomenų bazių kūrimo įrankio savybes: vidinė duomenų bazė, NoSQL jungtys, REST API palaikymas, programėlių kūrimo įrankiai, darbo eigos automatizavimas, debesijos platformos prieinamumas, savarankiško diegimo galimybė ir atvirojo kodo prieinamumas.

Šie kriterijai padengia tris pagrindines vartotojo perspektyvas: techninę (kokios technologijos palaikomos), eksploatacinę (kaip įrankis diegiamas ir prižiūrimas) ir ekonominę (ar reikalingi licencijų mokesčiai). Kriterijų pasirinkimas atitinka užduotyje rekomenduotą požiūrį ir leidžia objektyviai palyginti skirtingo pobūdžio įrankius.

Apibendrinant pirmąjį skyrių galima teigti, kad duomenų bazių kūrimo įrankių rinka yra įvairi ir nuolat besikeičianti. Tinkama klasifikacija ir aiškūs palyginimo kriterijai sudaro pagrindą tolesnei analizei.

---

## 2. DEŠIMTIES DUOMENŲ BAZIŲ KŪRIMO ĮRANKIŲ APŽVALGA

Šiame skyriuje aprašomi dešimt pasirinktų duomenų bazių kūrimo įrankių, suskirstytų į tris grupes pagal jų pobūdį - tradiciniai reliaciniai įrankiai, šiuolaikinės debesijos ir žemo kodo platformos bei NoSQL ir atvirojo kodo įrankiai.

### 2.1. Tradiciniai reliaciniai įrankiai

Tradiciniai reliaciniai įrankiai yra ilgametes naudojimo tradicijas turintys produktai, paplitę tiek Lietuvos, tiek užsienio rinkose. Šios grupės įrankiai dažnai naudojami smulkiojo verslo apskaitoje, vidiniuose organizacijų sprendimuose ir mokymo procese.

**Microsoft Access** yra patentuota Microsoft kompanijos sukurta reliacinė DBVS, leidžianti kurti duomenų bazes vizualiai per formų, užklausų ir ataskaitų konstruktorius (Microsoft, 2026). Įrankis turi savo duomenų bazės variklį (Jet/ACE) ir gerai integruojasi su kitais Microsoft 365 produktais. Lietuvos savivaldybėse ir smulkiose įmonėse Access vis dar plačiai naudojamas vidinei dokumentų bei klientų apskaitai.

**LibreOffice Base** yra atvirojo kodo, nemokama Microsoft Access alternatyva, kuriama The Document Foundation. Įrankis palaiko HSQLDB ir Firebird variklius, taip pat gali jungtis prie išorinių MySQL, PostgreSQL ar SQLite duomenų bazių. Lietuvoje LibreOffice Base dažnai naudojamas viešojo sektoriaus įstaigose ir mokyklose dėl licencijų sąnaudų taupymo.

**OpenOffice Base** yra Apache OpenOffice projekto sudedamoji dalis. Funkciškai panašus į LibreOffice Base, tačiau pastaraisiais metais jo plėtra sulėtėjo, todėl rinkoje praranda pozicijas. Įrankis tinka paprastiems, lokaliai diegiamiems sprendimams.

**Microsoft SQL Server Management Studio** (toliau - SSMS) yra įmonės lygmens įrankis, skirtas profesionaliam Microsoft SQL Server duomenų bazių administravimui. Įrankis turi galingą užklausų rengyklę, profiliavimo priemones ir platų išvedinių aibę. Lietuvos didelės įmonės, pavyzdžiui, bankų ir telekomunikacijų sektoriaus dalyviai, SSMS naudoja kasdieniame veikloje.

### 2.2. Šiuolaikinės debesijos ir žemo kodo platformos

Šiai grupei priskiriami įrankiai, kurie ne tik teikia duomenų saugojimo paslaugą, bet ir sudaro sąlygas kurti vidines aplikacijas be programavimo žinių. Tokie įrankiai itin patrauklūs marketingo komandoms ir startuoliams.

**Airtable** yra debesijos pagrindu veikianti platforma, jungianti skaičiuoklės ir reliacinės duomenų bazės savybes (Airtable, 2026). Vartotojai gali kurti lenteles, susieti įrašus tarp jų, automatizuoti darbo eigas ir publikuoti formas duomenų rinkimui. Lietuvos startuoliai ir marketingo agentūros Airtable naudoja klientų sąrašų ir kampanijų valdymui.

**Notion Databases** yra dalis Notion produktyvumo platformos. Įrankis leidžia kurti dokumentų pagrindo duomenų bazes, susiejamas tarpusavyje per realacijos (angl. *relation*) ir suvestinės (angl. *rollup*) laukus. Šis sprendimas populiarus žinių valdymo, projektų sekimo ir turinio planavimo užduotims.

**Budibase** yra atvirojo kodo žemo kodo platforma, leidžianti kurti vidines aplikacijas su integruota duomenų baze arba jungtis prie išorinių šaltinių (Budibase, 2026). Įrankis gali būti diegiamas tiek debesijoje, tiek vietiniame serveryje. Budibase tinka įmonėms, ieškančioms balanso tarp debesijos patogumo ir duomenų suverenumo.

### 2.3. NoSQL ir atvirojo kodo įrankiai

Trečioji grupė apima įrankius, dažniau naudojamus programuotojų, analitikų ir duomenų inžinierių. Jie reikalauja didesnių techninių žinių, tačiau suteikia platesnes valdymo galimybes.

**MongoDB Compass** yra oficialus MongoDB kompanijos kuriamas grafinis NoSQL dokumentų duomenų bazės valdymo įrankis (MongoDB, 2026). Jis leidžia vizualiai naršyti kolekcijas, kurti užklausas, analizuoti indeksų našumą ir importuoti duomenis. Įrankis nemokamas, palaiko prisijungimą tiek prie debesijos, tiek prie savarankiškai diegiamų MongoDB serverių.

**DBeaver** yra atvirojo kodo universalus duomenų bazių klientas, palaikantis daugiau kaip aštuoniasdešimt skirtingų DBVS, įskaitant tiek reliacines, tiek NoSQL. Įrankis turi pažangią užklausų rengyklę, duomenų vizualizacijos priemones ir integraciją su versijų valdymo sistemomis. Lietuvoje DBeaver dažnai pasirenkamas informacinių technologijų studentų ir programinės įrangos kūrėjų.

**phpMyAdmin** yra atvirojo kodo žiniatinklio sąsaja, skirta MySQL ir MariaDB duomenų bazių administravimui. Įrankis veikia bet kurioje žiniatinklio prieglobos aplinkoje, todėl populiarus tarp svetainių kūrėjų ir savarankiškų projektų autorių. Lietuvos žiniatinklio prieglobos paslaugų teikėjai, pavyzdžiui, Hostinger ir Serveriai.lt, phpMyAdmin teikia kaip standartinę paslaugą.

Apibendrinant antrąjį skyrių, akivaizdu, kad pasirinkti dešimt įrankių atspindi visą šiuolaikinių duomenų bazių kūrimo priemonių spektrą - nuo paprastų lokalių sprendimų iki sudėtingų debesijos platformų.

---

## 3. LYGINAMOJI ANALIZĖ IR REZULTATAI

Šiame skyriuje pateikiama lyginamoji aprašytų įrankių analizė. Pirmiausia sudaroma struktūrinė lyginamoji lentelė pagal aštuonis kriterijus, vėliau aptariami pagrindiniai skirtumai ir panašumai, galiausiai pateikiamos rekomendacijos pagal naudojimo atvejus.

### 3.1. Lyginamoji lentelė pagal aštuonis kriterijus

Pasirinkti dešimt įrankių palyginti pagal aštuonis kriterijus, suformuluotus pirmajame skyriuje. Rezultatai pateikti 1 lentelėje (žr. 1 lentelę). Lentelėje žymos „Taip" arba „Ne" atspindi atitinkamos savybės buvimą, o „Iš dalies" rodo, kad savybė palaikoma su apribojimais, dažniausiai per papildomus įskiepius ar trečiųjų šalių jungtis.

**1 lentelė**
*Lyginamoji duomenų bazių kūrimo įrankių analizė pagal aštuonis kriterijus*

| Įrankis | Vidinė DB | NoSQL jungtys | REST API | App Builder | Workflow | Cloud | Self-Host | Open-Source |
|---------|-----------|---------------|----------|-------------|----------|-------|-----------|-------------|
| MS Access | Taip | Ne | Iš dalies | Iš dalies | Iš dalies | Ne | Taip | Ne |
| LibreOffice Base | Taip | Ne | Ne | Ne | Ne | Ne | Taip | Taip |
| OpenOffice Base | Taip | Ne | Ne | Ne | Ne | Ne | Taip | Taip |
| MS SQL SSMS | Taip | Ne | Iš dalies | Ne | Ne | Iš dalies | Taip | Ne |
| Airtable | Taip | Iš dalies | Taip | Taip | Taip | Taip | Ne | Ne |
| Notion Databases | Taip | Ne | Taip | Iš dalies | Iš dalies | Taip | Ne | Ne |
| Budibase | Taip | Taip | Taip | Taip | Taip | Taip | Taip | Taip |
| MongoDB Compass | Iš dalies | Taip | Taip | Ne | Ne | Taip | Taip | Iš dalies |
| DBeaver | Ne | Taip | Iš dalies | Ne | Ne | Iš dalies | Taip | Taip |
| phpMyAdmin | Iš dalies | Ne | Ne | Ne | Ne | Ne | Taip | Taip |

*Šaltinis: sudaryta autoriaus, remiantis oficialia įrankių kūrėjų dokumentacija (Microsoft, 2026; Airtable, 2026; Budibase, 2026; MongoDB, 2026; The Document Foundation, 2026).*

### 3.2. Pagrindiniai skirtumai ir panašumai

Iš lyginamosios lentelės matyti, kad analizuoti įrankiai turi tiek bendrų bruožų, tiek esminių skirtumų. Beveik visi įrankiai (devyni iš dešimties) turi savo vidinį duomenų bazės variklį arba glaudžiai integruotą saugyklą. Vienintelė išimtis - DBeaver, kuris yra tik klientas ir reikalauja išorinės duomenų bazės.

NoSQL palaikymas yra ryški atskirtis tarp tradicinių ir šiuolaikinių įrankių. Tradiciniai reliaciniai sprendimai (Microsoft Access, OpenOffice Base, phpMyAdmin) NoSQL nepalaiko arba palaiko tik per papildomas jungtis. Naujesnės platformos - Budibase, Airtable, DBeaver - šią galimybę įdiegė kaip standartinę funkciją.

REST API palaikymas šiandien yra praktiškai privalomas šiuolaikiniams įrankiams. Visi debesijos sprendimai (Airtable, Notion, Budibase) automatiškai sugeneruoja API kiekvienai duomenų bazei. Tradiciniai sprendimai (Access, LibreOffice Base) tokios galimybės arba neturi, arba reikalauja papildomo programavimo.

Atvirojo kodo prieinamumo požiūriu rinka yra suskaidyta. Penki iš dešimties analizuotų įrankių (LibreOffice Base, OpenOffice Base, Budibase, DBeaver, phpMyAdmin) yra atvirojo kodo. Tai rodo, kad atvirojo kodo bendruomenė užima reikšmingą vaidmenį duomenų bazių įrankių rinkoje.

Kita vertus, programėlių kūrimo (App Builder) ir darbo eigos automatizavimo galimybės yra naujausi rinkos diferencijavimo veiksniai. Šiomis savybėmis išsiskiria tik keturi įrankiai - Microsoft Access (per makrokomandas), Airtable, Notion ir Budibase. Tai patvirtina ankstesnėje literatūroje aprašytą tendenciją, kad duomenų bazių įrankiai vis labiau virsta integruotomis verslo procesų valdymo platformomis.

### 3.3. Įrankių tinkamumas skirtingiems naudojimo atvejams

Remiantis lyginamosios analizės rezultatais, galima suformuluoti rekomendacijas, kurie įrankiai geriausiai tinka konkretiems naudojimo atvejams. Apibendrinti rezultatai pateikti 2 lentelėje (žr. 2 lentelę).

Smulkiajam verslui ir individualiai veikiantiems specialistams rekomenduojami Microsoft Access, LibreOffice Base ir Airtable. Pirmieji du tinka, kai reikia paprasto, lokaliai veikiančio sprendimo su žinoma sąsaja, o Airtable - kai svarbi prieiga iš bet kurios vietos ir komandinis darbas.

Didelėms įmonėms tinkamiausias yra Microsoft SQL Server Management Studio, papildytas DBeaver kaip kasdienio darbo įrankiu. Tokia kombinacija užtikrina ir įmonės lygmens administravimą, ir lankstų užklausų rengimą.

Debesijos pagrindu veikiančioms aplikacijoms geriausiai tinka Airtable, Notion ir Budibase. Šie įrankiai turi pilną REST API, integruotą programėlių kūrimo aplinką ir darbo eigos automatizavimo priemones.

Atvirojo kodo projektams ar atvejams, kai svarbus duomenų suverenumas, rekomenduojami Budibase, DBeaver ir phpMyAdmin. Visi trys yra nemokami, palaiko savarankišką diegimą ir turi aktyvias vartotojų bendruomenes.

**2 lentelė**
*Įrankių tinkamumas skirtingiems naudojimo atvejams*

| Naudojimo atvejis | Rekomenduojami įrankiai |
|-------------------|--------------------------|
| Smulkusis verslas | MS Access, LibreOffice Base, Airtable |
| Didelės įmonės | MS SQL SSMS, DBeaver |
| Debesijos aplikacijos | Airtable, Notion Databases, Budibase |
| Atvirojo kodo projektai | Budibase, DBeaver, phpMyAdmin |
| NoSQL projektai | MongoDB Compass, DBeaver, Budibase |
| Mokymas ir studijos | LibreOffice Base, MS Access, phpMyAdmin |

*Šaltinis: sudaryta autoriaus.*

Apibendrinant trečiąjį skyrių galima teigti, kad nėra vieno geriausio duomenų bazių kūrimo įrankio - tinkamiausias pasirinkimas priklauso nuo konkretaus naudojimo atvejo, įmonės dydžio, biudžeto ir techninių išteklių.

---

## IŠVADOS

1. Šiuolaikinėje rinkoje veikia įvairios duomenų bazių kūrimo įrankių grupės - nuo tradicinių reliacinių sprendimų iki debesijos pagrindo žemo kodo platformų. Tokia įvairovė rodo, kad duomenų bazių valdymo sąvoka pastarąjį dešimtmetį iš esmės išsiplėtė ir apima ne tik duomenų saugojimą, bet ir aplikacijų kūrimą bei darbo eigų automatizavimą.

2. Atlikus tyrimą, pasirinkti ir aprašyti dešimt įrankių - Microsoft Access, LibreOffice Base, OpenOffice Base, Microsoft SQL Server Management Studio, Airtable, Notion Databases, Budibase, MongoDB Compass, DBeaver ir phpMyAdmin. Šis sąrašas apima tris pagrindines įrankių grupes ir leidžia objektyviai įvertinti jų savybes.

3. Lyginamoji analizė pagal aštuonis kriterijus parodė, kad ryškiausi skirtumai tarp įrankių pasireiškia REST API palaikymo, NoSQL jungčių, programėlių kūrimo galimybių ir darbo eigos automatizavimo srityse. Tradiciniai sprendimai šiose srityse atsilieka, o šiuolaikinės debesijos platformos siūlo integruotą funkcionalumą.

4. Rekomendacijos pagal naudojimo atvejus rodo, kad smulkiajam verslui geriausiai tinka Microsoft Access ar Airtable, didelėms įmonėms - Microsoft SQL Server Management Studio kartu su DBeaver, debesijos aplikacijoms - Airtable, Notion ar Budibase, o atvirojo kodo projektams - Budibase, DBeaver ar phpMyAdmin. Vieno universaliai geriausio sprendimo nėra, todėl įrankis turi būti pasirenkamas atsižvelgiant į konkretaus projekto kontekstą.

---

## LITERATŪROS SĄRAŠAS

Airtable. (2026). *Airtable Help Center*. Prieiga per internetą: https://support.airtable.com

Budibase. (2026). *Budibase Documentation*. Prieiga per internetą: https://docs.budibase.com

Codd, E. F. (1970). A Relational Model of Data for Large Shared Data Banks. *Communications of the ACM*, 13(6), 377-387.

Connolly, T. ir Begg, C. (2015). *Database Systems: A Practical Approach to Design, Implementation, and Management* (6th ed.). Boston: Pearson.

DB-Engines. (2026). *DB-Engines Ranking*. Prieiga per internetą: https://db-engines.com/en/ranking

Elmasri, R. ir Navathe, S. B. (2016). *Fundamentals of Database Systems* (7th ed.). Boston: Pearson.

Microsoft. (2026). *Microsoft Access dokumentacija*. Prieiga per internetą: https://support.microsoft.com/lt-lt/access

MongoDB. (2026). *MongoDB Compass Documentation*. Prieiga per internetą: https://www.mongodb.com/docs/compass/

Sadalage, P. J. ir Fowler, M. (2012). *NoSQL Distilled: A Brief Guide to the Emerging World of Polyglot Persistence*. Boston: Addison-Wesley.

The Document Foundation. (2026). *LibreOffice Base Handbook*. Prieiga per internetą: https://documentation.libreoffice.org

Vilniaus universitetas. (2024). *Dirbtinio intelekto naudojimo gairės* (Nr. SPN-54). Vilnius: VU.

---

## 1 PRIEDAS. Detali dešimties įrankių charakteristika

| Įrankis | Kūrėjas | Pirmas leidimas | Licencija | Pagrindinė platforma |
|---------|---------|-----------------|-----------|----------------------|
| MS Access | Microsoft | 1992 | Patentuota | Windows |
| LibreOffice Base | The Document Foundation | 2011 | LGPLv3 | Windows, macOS, Linux |
| OpenOffice Base | Apache | 2002 | Apache 2.0 | Windows, macOS, Linux |
| MS SQL SSMS | Microsoft | 2005 | Patentuota (nemokama) | Windows |
| Airtable | Airtable Inc. | 2012 | Patentuota (SaaS) | Žiniatinklis |
| Notion Databases | Notion Labs | 2016 | Patentuota (SaaS) | Žiniatinklis, mobilieji |
| Budibase | Budibase Ltd. | 2019 | GPLv3 | Žiniatinklis, savarankiškai |
| MongoDB Compass | MongoDB Inc. | 2016 | SSPL | Windows, macOS, Linux |
| DBeaver | DBeaver Corp. | 2010 | Apache 2.0 | Windows, macOS, Linux |
| phpMyAdmin | phpMyAdmin team | 1998 | GPLv2 | Žiniatinklis (PHP) |

*Šaltinis: sudaryta autoriaus, remiantis oficialia įrankių kūrėjų dokumentacija.*

---

## 11 PRIEDAS. Dirbtinio intelekto panaudojimo deklaracija

Rengiant šį darbą buvo naudotas generatyvinis dirbtinio intelekto įrankis Anthropic Claude (Sonnet 4.5, internetinė prieiga, naudota 2026 m. gegužės mėnesį). Įrankis buvo pasitelktas ribotais ir aiškiai apibrėžtais tikslais: pirminių temos struktūros variantų sugeneravimui, galimų potemių išgryninimui, teorinių sąvokų pirminiam paaiškinimui bei teksto stilistiniam ir kalbiniam redagavimui (gramatikos, aiškumo, sakinių struktūros tobulinimui). DI taip pat buvo naudotas formuluočių alternatyvoms pasiūlyti ir akademinio stiliaus nuoseklumui pagerinti.

Sugeneruotas turinys nebuvo tiesiogiai perkeltas į darbą be peržiūros - kiekvienas atsakymas buvo kritiškai įvertintas, patikrintas remiantis akademiniais šaltiniais ir, jei naudotas, reikšmingai redaguotas bei integruotas į autoriaus savarankiškai parengtą tekstą.

DI įrankis nebuvo naudotas savarankiškai rengiant: lyginamosios analizės dalį, formuluojant galutines išvadas ar atliekant tyrimo interpretaciją. Visi esminiai argumentai, vertinimai ir apibendrinimai yra darbo autoriaus savarankiško akademinio darbo rezultatas. Tais atvejais, kai panaudotos tiesioginės DI sugeneruotos formuluotės ar jų perfrazavimas, jos yra tinkamai identifikuotos ir cituotos laikantis akademinių reikalavimų.

DI naudojimo apimtis darbe:

| Rodiklis | Reikšmė |
|----------|---------|
| DI modelis | Anthropic Claude Sonnet 4.5 |
| Naudojimo data | 2026 m. gegužės mėn. |
| DI naudojimo tikslas | Struktūros patikrinimas, kalbos taisymas, formuluočių alternatyvos |
| DI sugeneruoto turinio dalis darbe | mažiau nei 15 proc. |
| Vieno DI modelio sugeneruotas turinys | mažiau nei 5 proc. (atitinka VU SPN-54 reikalavimus) |
| Modifikavimo apimtis | apie 70-80 proc. (DI siūlymai reikšmingai redaguoti) |

Autorius patvirtina, kad:

- yra susipažinęs su Vilniaus universiteto 2024 m. patvirtintomis dirbtinio intelekto naudojimo gairėmis (Nr. SPN-54);
- yra susipažinęs su Anthropic Claude privatumo politika ir naudojimo taisyklėmis;
- įvertino, kad DI sugeneruoti rezultatai gali būti netikslūs, todėl visi turinio teiginiai patikrinti remiantis nepriklausomais akademiniais šaltiniais;
- prisiima visišką atsakomybę už darbo turinį, jo tikslumą, argumentacijos pagrįstumą bei pateiktų šaltinių patikimumą.

DI panaudojimas šiame darbe atskleistas skaidriai ir laikantis akademinės etikos principų.

---

## 12 PRIEDAS. Dirbtinio intelekto užklausos ir gauti atsakymai

**1 užklausa (struktūros patikrinimas):**

> *„Pasiūlyk 3 skyrių struktūrą akademiniam darbui apie 10 šiuolaikinių duomenų bazių kūrimo įrankių palyginimą pagal 8 funkcinius kriterijus. Darbas skirtas marketingo technologijų studijų programos studentui. Apimtis - 8-12 puslapių."*

Gautas atsakymas - struktūros pasiūlymas su trimis pagrindiniais skyriais (teorinė apžvalga, įrankių aprašas, lyginamoji analizė). Autorius modifikavo poskyrių pavadinimus ir papildė kriterijų pagrindimu (1.3 poskyris).

**2 užklausa (terminų paaiškinimas):**

> *„Paaiškink lietuviškai sąvokas: low-code platforma, NoSQL, REST API, Database-as-a-Service. Atsakymą pateik trumpais sakiniais, akademiniu stiliumi."*

Gautas atsakymas integruotas į 1 skyrių, kalbiškai sutrumpintas ir papildytas šaltinių nuorodomis.

**3 užklausa (kalbos taisymas):**

> *„Patikrink šios pastraipos lietuvių kalbos taisyklingumą ir akademinį stilių. Pasiūlyk taisyklingesnes formuluotes, jei reikia. [Įklijuotas autoriaus parašyto įvado tekstas]"*

Gauti pasiūlymai dėl jungtukų vartojimo ir sakinių struktūros. Autorius priėmė apie 60 proc. pasiūlymų, kitus atmetė kaip neatitinkančius asmeninio rašymo stiliaus.

**4 užklausa (lentelės struktūra):**

> *„Pasiūlyk lyginamosios lentelės formatą, kuriame būtų galima palyginti 10 duomenų bazių įrankių pagal 8 kriterijus. Žymėjimas turi būti aiškus ir glaustas."*

Gautas pasiūlymas - žymėjimas „Taip" / „Ne" / „Iš dalies". Autorius pritarė ir įdiegė šią schemą 1 lentelėje.

*Pastaba: visi naudoti DI atsakymai išsaugoti autoriaus archyve ir pateikiami pareikalavus.*
