# Instrukcija: kaip naudoti 3 PU5 .odb failus

## 1. Atsisiųskite LibreOffice (jei dar neturite)

https://www.libreoffice.org/download/

## 2. Atsisiųskite tris .odb failus iš GitHub

- https://github.com/andriusstanku-lab/a/raw/pu5-darbas/PU5_darbas/eprekyba.odb
- https://github.com/andriusstanku-lab/a/raw/pu5-darbas/PU5_darbas/ligonine.odb
- https://github.com/andriusstanku-lab/a/raw/pu5-darbas/PU5_darbas/biblioteka.odb

## 3. Atidarykite kiekvieną dukart spustelėję

Vartotojo vardas: `SA`, slaptažodis tuščias.

## 4. Padarykite ekrano nuotraukas

Kiekvienai duomenų bazei rekomenduoju padaryti 1-2 ekrano nuotraukas:

### eprekyba.odb
- Atidarykite **Tools - Relationships** (rodo 1:N ryšius)
- Atidarykite **Užsakymai** lentelę (Datasheet View) - 5 įrašai
- Išsaugokite kaip `img4_data_eprekyba.png`

### ligonine.odb
- Atidarykite **Vizitai** lentelę
- Išsaugokite kaip `img5_data_ligonine.png`

### biblioteka.odb
- Atidarykite **Skolinimasi** lentelę
- Išsaugokite kaip `img6_data_biblioteka.png`

## 5. Pakeiskite mockup paveikslus tikromis nuotraukomis

Įkelkite atnaujintus paveikslus į GitHub (drag-and-drop) ir pranešite man - perpaleisiu skriptą su tikromis nuotraukomis.

## CSV alternatyva

Jei .odb failai nepasiekiami, naudokite CSV failus iš `csv/` aplanko per **File - New - Database - Connect to existing database - Text**.

## Komandos LibreOffice Base srityje

| Komanda | Veiksmas |
|---|---|
| Tables - Edit | Lentelės struktūros redagavimas (Design View) |
| Tables - dukart spustelti | Lentelės duomenų peržiūra (Datasheet View) |
| Tools - Relationships | Lentelių santykių diagrama |
| Tools - SQL | SQL užklausų rengyklė |
