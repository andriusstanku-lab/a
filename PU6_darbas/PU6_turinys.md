# PU6 PRAKTINĖS UŽDUOTIES TURINYS

**Tema:** Duomenų normalizavimas (1NF, 2NF, 3NF)
**Autorius:** Andrius Vargonas
**Metai:** 2026

## Trijų dalykinių sričių analizė

### I. Studentų registracija
| Žingsnis | Lentelės |
|---|---|
| **Pradžia** | StudentCourseRegistration (7 stulpeliai, 1 lentelė) |
| **Po 1NF** | Tas pats (jau atominės reikšmės) |
| **Po 2NF** | Students, Courses, Enrollments (3 lentelės) |
| **Po 3NF** | Students, Instructors, Courses, Enrollments (4 lentelės) |

### II. Pardavimai
| Žingsnis | Lentelės |
|---|---|
| **Pradžia** | Sales (9 stulpeliai, 1 lentelė) |
| **Po 1NF/2NF** | Tas pats (TransactionID - vienas PK, nėra dalinių) |
| **Po 3NF** | Customers, Products, Salespeople, Sales (4 lentelės) |

### III. Sandėliai
| Žingsnis | Lentelės |
|---|---|
| **Pradžia** | WarehouseProductSupplier (9 stulpeliai, 1 lentelė) |
| **Po 1NF** | Tas pats (jau atominės reikšmės) |
| **Po 3NF** | Warehouses, Products, Suppliers, Deliveries (4 lentelės) |

## Pagrindinės normalizavimo idėjos

**1NF:** Atominės reikšmės, nėra pasikartojančių grupių
**2NF:** 1NF + nepriminio rakto atributai pilnai priklauso nuo viso PK
**3NF:** 2NF + nėra tranzityvinių priklausomybių (atributų, priklausančių per kitą atributą)

## Trys realizuotos duomenų bazės

| DB failas | Lentelės | Įrašai | Apimtis |
|---|---|---|---|
| studentai.odb | 4 | 20 | 3 KB |
| pardavimai.odb | 4 | 20 | 3 KB |
| sandeliai.odb | 4 | 20 | 3 KB |

Iš viso: **12 lentelių, 60 įrašų, 9 KB**.
