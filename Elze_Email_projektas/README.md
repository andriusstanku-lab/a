# Elzės Zdancevičiūtės knygos „Turtingas(is)" pardavimo augimo planas

**E. pardavimų grupinis projektas** | Vilniaus universitetas | 2026 m. gegužė

Kanalas: **Email / community building**
Klientė: Elzė Zdancevičiūtė, [elzėsknygos.lt](https://xn--elzsknygos-tmb.lt/)
Knyga: „Turtingas(is)" (Piko Valanda, 2025; tiražas 100 egz.)

---

## Dokumentai (`dokumentai/`)

Visi paruošti perduoti dėstytojui ir Elzei. VU stilius: Times New Roman 12pt, 1,5 tarpai, justify, statinis TURINYS (be auto-laukų — atsidaro be jokio atnaujinimo).

| Failas | Apimtis | Etapas |
|---|---|---|
| `1_OBSERVE_Esama_busena.docx` | ~5 psl. + screenshot priedai | OBSERVE — esama būsena |
| `2_ANALYZE_Diagnostika_ir_potencialas.docx` | ~5 psl. | ANALYZE — diagnostika ir potencialas |
| `3_IMPROVE_Pasiulymai.docx` | ~6 psl. | IMPROVE — 3 pasiūlymai (1, 3, 12 mėn.) |
| `4_VALIDATE_Master_Document.docx` | ~8 psl. | VALIDATE — pilnas konsoliduotas Master Document |
| `STARTER_PACK_Email_kanalo_ijungimas.docx` | ~7 psl. | Starter pack — 5 paruošti laiškai + outreach šablonas |
| `5_PRISTATYMAS_10_min_struktura.docx` | ~7 psl. | 10 min. pristatymo struktūra ir kalbėjimo planas |

## Pagalbiniai aplankai

- `scripts/` — Python skriptai, kuriais sugeneruoti dokumentai. Reikia `python-docx`. Paleidžiama: `pip install python-docx && python build_1_observe.py` ir t. t.
- `research/` — saugomi pradiniai HTML šaltiniai (kliento svetainė, Instagram), kuriais remtasi audite.

---

## Pagrindiniai realūs radiniai (ne prielaidos)

Iš tiesioginio `elzėsknygos.lt` audito:

- **0 (nulis) HTML formų** visoje svetainėje — newsletter neegzistuoja.
- **Kontaktinis paštas yra placeholder** `mailto:example@email.com` (HTML kode), tekste rodo `elze.zdanceviciute@email.com`.
- **Tiražas — 100 egz.** (kolofonas svetainės puslapyje „Knyga"). Tai keičia projekto tikslą: kanalas yra ne šios knygos, o autorinės karjeros bazė.
- **Knyga**: 213 psl., ISBN 9786094222245, leidykla „Piko Valanda".
- **Tipo klaida** pagrindiniame puslapyje: „EIzę" vietoje „Elzę".
- **Literatūrinis identitetas autorės puslapyje**: Camus, Dostojevskis, Kafka, Škėma, Mačernis, Kernagis, Bob Dylan, The Doors.

---

## Industrijos benchmark'ai (su šaltiniais, 2025–2026)

- Newsletter open rate vidurkis: 31–43 % (MailerLite, Brevo, Thunderbit, AdMailr).
- Welcome laiško open rate: 4× įprasto = ~55–65 % (EmailToolTester, Stackmatix).
- Indie autoriaus pavyzdys: Richie Billing — pasikartojantis 0 → 100 → 400–500 prenumeratorių per metus, naudojant tik reader magnet ir newsletter swaps (StoryOrigin case study).
- MailerLite nemokamas planas 2025–2026: iki 500 prenumeratorių, 12 000 laiškų/mėn.

---

## Kas reikia padaryti grupei prieš pristatymą

1. Įrašyti grupės narių vardus visuose dokumentuose vietoje `[Grupės narių vardai, pavardės]`.
2. Padaryti screenshot'us (Priedas A1–A6 pagal sąrašą OBSERVE dokumente) — naudoti inkognito naršyklę.
3. Patikrinti Instagram `@elze_zdan` sekėjų skaičių susitikimo su Elze metu (mūsų prielaida: 200–800).
4. Patikrinti rankomis, kur veda „ĮSIGYTI KNYGĄ" mygtukas (svetainėje nėra WooCommerce signalų — tikriausiai išorinė parduotuvė).
5. Užpildyti starter pack outreach lentelę — 5–7 konkretūs Lietuvos knygų bloggeriai/BookTok kūrėjos kartu su Elze.

---

## Trijų pasiūlymų santrauka

| Pasiūlymas | Laiko kaštai | Pinigų kaštai | Pagrindinė metrika |
|---|---|---|---|
| **Nr. 1 (1 mėn.)** „Kanalo įjungimas" | ~7 val. vienkartiniai + 30 min./sav. | 0 € (galimai 10–20 €/m. domeno paštui) | ≥30 prenum., welcome open ≥50 % |
| **Nr. 2 (3 mėn.)** „Reguliarus ritmas" | ~45 min. / sav. | 0 € | open rate ≥35 %, ≥100 prenum. |
| **Nr. 3 (12 mėn.)** „Karjeros bazė" | tas pats + 2–3 vizitai | ~30–60 €/m. | ≥250 prenum., reguliarumas išlaikytas |

---

## Repository informacija

Šaka: `email-community-building-elze`
Šis projektas yra atskiras nuo `PU1_darbas/` — anksčiau įdėtas darbas tame pačiame repo.
