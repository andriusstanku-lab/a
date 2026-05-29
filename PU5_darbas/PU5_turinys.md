# PU5 PRAKTINĖS UŽDUOTIES TURINYS

**Darbo pavadinimas:** DUOMENŲ BAZĖS LENTELIŲ KŪRIMAS PAGAL ERD
**Autorius:** Andrius Vargonas
**Metai:** 2026, Kaunas

## TURINYS

```
ĮVADAS .............................................. 3
1. ESYBIŲ-RYŠIŲ MODELIAVIMO PAGRINDAI ................ 4
   1.1. ER diagramos samprata ir paskirtis ............ 4
   1.2. Esybės, atributai ir ryšiai .................. 5
   1.3. Konvertavimas į reliacines lenteles .......... 6
2. TRIJŲ DALYKINIŲ SRIČIŲ ER DIAGRAMOS ............... 7
   2.1. E-prekybos sistema ........................... 7
   2.2. Ligoninės valdymo sistema .................... 9
   2.3. Bibliotekos valdymo sistema .................. 11
3. DUOMENŲ BAZIŲ ĮGYVENDINIMAS ....................... 13
   3.1. Pasirinktas DBVS įrankis ..................... 13
   3.2. E-prekybos duomenų bazė ...................... 14
   3.3. Ligoninės duomenų bazė ....................... 15
   3.4. Bibliotekos duomenų bazė ..................... 16
   3.5. Patirties aprašymas .......................... 17
IŠVADOS ............................................. 18
LITERATŪROS SĄRAŠAS ................................. 19
1 PRIEDAS. SQL DDL skriptai ......................... 20
11 PRIEDAS. DI panaudojimo deklaracija .............. 22
```

## ĮVADAS (santrauka)

Tikslas: trims dalykinėms sritims sukurti ER diagramas ir, remiantis jomis, įgyvendinti tris reliacines duomenų bazes LibreOffice Base aplinkoje.

DI naudojimas: tik tekstui ir paaiškinimams. Diagramas, lenteles, raktus, duomenis - autorius padarė savarankiškai.

## 3 dalykinės sritys

### E-prekyba
| Esybė | PK | Atributai | FK |
|---|---|---|---|
| Klientai | KlientoID | Vardas, El. paštas | - |
| Produktai | ProduktoID | Pavadinimas, Kaina, Kategorija | - |
| Užsakymai | UžsakymoID | Data, Kiekis, Adresas | KlientoID, ProduktoID |

### Ligoninė
| Esybė | PK | Atributai | FK |
|---|---|---|---|
| Pacientai | PacientoID | Vardas | - |
| Gydytojai | GydytojoID | Vardas, Specializacija | - |
| Vizitai | VizitoID | Data, Diagnozė, Pastabos | PacientoID, GydytojoID |

### Biblioteka
| Esybė | PK | Atributai | FK |
|---|---|---|---|
| Knygos | KnygosID | Pavadinimas, Autorius, ISBN | - |
| Nariai | NarioID | Vardas, Narystės data | - |
| Skolinimasi | SkolinimoID | Skol. data, Grąž. data | KnygosID, NarioID |

## IŠVADOS (4 punktai)

1. ER diagrama yra esminis tarpinis žingsnis tarp dalykinės srities supratimo ir reliacinės DB.
2. Visoms 3 sritims taikytas šablonas: 2 pagrindinės esybės + 1 siejanti esybė (1:N ryšiai).
3. 3 .odb failai sėkmingai sukurti su PK/FK ir 5 įrašais kiekvienoje lentelėje.
4. LibreOffice Base patogus mokomajai praktikai - intuityvus Design View ir Relationships langas.
