"""
Generuoja du PU1 darbo paveikslus naudojant Graphviz (vietoj matplotlib).
Graphviz automatiškai sutvarko išdėstymą, todėl rodyklės neis per dėžutes.

1 pav. - Pigu.lt ir Omniva sąveikos schema (paprastas srautas)
2 pav. - Supaprastinta e. prekybos ir logistikos ER diagrama
"""

from graphviz import Digraph


# =====================================================
# 1 pav. - Pigu.lt ir Omniva sąveikos schema
# =====================================================
g1 = Digraph("seku_diagrama", format="png")
g1.attr(rankdir="TB", bgcolor="white", pad="0.3", nodesep="0.5", ranksep="0.6")
g1.attr("node",
        shape="box",
        style="rounded,filled",
        fillcolor="white",
        color="black",
        fontname="DejaVu Sans",
        fontsize="11",
        margin="0.2,0.12")
g1.attr("edge",
        fontname="DejaVu Sans",
        fontsize="10",
        color="black")

# Mazgai
g1.node("klientas", "KLIENTAS\n(naršyklėje arba programėlėje)", fillcolor="#f5f5f5")
g1.node("pigu", "Pigu.lt informacinė sistema\n(svetainė, krepšelis, mokėjimas)", fillcolor="#e8f4fd")
g1.node("api", "Omniva API sąsaja\n(duomenų perdavimas tarp sistemų)", fillcolor="#fff8e1")
g1.node("db", "Omniva duomenų bazė\n(saugomi siuntų įrašai)", fillcolor="#e8f5e9")

# Rodyklės su trumpais aprašymais
g1.edge("klientas", "pigu", label="  1. Užsako prekę  \n  ir pasirenka Omniva paštomatą  ")
g1.edge("pigu", "api", label="  2. Perduoda siuntimo duomenis  ")
g1.edge("api", "db", label="  3. Užregistruoja siuntą  ")
g1.edge("db", "api", label="  4. Grąžina sekimo numerį  ", style="dashed")
g1.edge("api", "pigu", label="  5. Patvirtina užsakymą  ", style="dashed")
g1.edge("pigu", "klientas", label="  6. Parodo sekimo numerį  ", style="dashed")

g1.render("/projects/sandbox/PU1_darbas/img1_seku_diagrama", cleanup=True)
print("Sukurta: img1_seku_diagrama.png")


# =====================================================
# 2 pav. - ER diagrama (paprasta, su HTML lentelėmis)
# =====================================================
g2 = Digraph("er_diagrama", format="png")
g2.attr(rankdir="LR", bgcolor="white", pad="0.3", nodesep="0.6", ranksep="1.0")
g2.attr("node", shape="plaintext", fontname="DejaVu Sans")
g2.attr("edge", fontname="DejaVu Sans", fontsize="10", color="black", arrowsize="0.7")


def entity_label(title, fields):
    """Sukuria HTML lentelę esybės atributams."""
    rows = "".join(
        f'<TR><TD ALIGN="LEFT" PORT="{f[0]}">{f[1]}</TD></TR>'
        for f in fields
    )
    return f"""<
    <TABLE BORDER="1" CELLBORDER="0" CELLSPACING="0" CELLPADDING="4">
      <TR><TD BGCOLOR="#d9d9d9"><B>{title}</B></TD></TR>
      {rows}
    </TABLE>>"""


# Esybės su atributais (P – pirminis raktas, F – išorinis raktas)
g2.node("klientas", entity_label("KLIENTAS", [
    ("kid", "kliento numeris (P)"),
    ("vardas", "vardas"),
    ("pavarde", "pavardė"),
    ("epastas", "el. paštas"),
    ("tel", "telefonas"),
]))

g2.node("uzsakymas", entity_label("UŽSAKYMAS", [
    ("uid", "užsakymo numeris (P)"),
    ("kid", "kliento numeris (F)"),
    ("data", "užsakymo data"),
    ("suma", "bendra suma"),
    ("busena", "būsena"),
]))

g2.node("eilute", entity_label("UŽSAKYMO EILUTĖ", [
    ("eid", "eilutės numeris (P)"),
    ("uid", "užsakymo numeris (F)"),
    ("pid", "produkto numeris (F)"),
    ("kiekis", "kiekis"),
    ("kaina", "kaina už vienetą"),
]))

g2.node("produktas", entity_label("PRODUKTAS", [
    ("pid", "produkto numeris (P)"),
    ("pav", "pavadinimas"),
    ("kaina", "kaina"),
    ("likutis", "likutis sandėlyje"),
]))

g2.node("siunta", entity_label("SIUNTA", [
    ("sid", "siuntos numeris (P)"),
    ("uid", "užsakymo numeris (F)"),
    ("adresas", "paštomato adresas"),
    ("busena", "pristatymo būsena"),
]))

# Ryšiai (kardinalumas: 1 ir N)
g2.edge("klientas:kid", "uzsakymas:kid", label="1 : N", taillabel="pateikia ", labeldistance="2")
g2.edge("uzsakymas:uid", "eilute:uid", label="1 : N", taillabel="turi ", labeldistance="2")
g2.edge("produktas:pid", "eilute:pid", label="1 : N", taillabel="yra ", labeldistance="2")
g2.edge("uzsakymas:uid", "siunta:uid", label="1 : 1", taillabel="pristatoma ", labeldistance="2")

g2.render("/projects/sandbox/PU1_darbas/img2_er_diagrama", cleanup=True)
print("Sukurta: img2_er_diagrama.png")
