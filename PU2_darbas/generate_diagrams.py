"""
Generuoja 9 organizacijų struktūrines diagramas naudojant Graphviz.

Pasirinktos organizacijos:
- Lidl Lietuva: hierarchinė, funkcinė, struktūra pagal padalinius
- Nord Security (NordVPN): plokščioji, matricinė, komandinė
- Swedbank Lietuva: tinklinė, projektinė, funkcinė
"""

from graphviz import Digraph

OUTPUT_DIR = "/projects/sandbox/PU2_darbas"

# Bendri stiliaus nustatymai
COMMON_NODE_ATTR = {
    "shape": "box",
    "style": "rounded,filled",
    "fontname": "DejaVu Sans",
    "fontsize": "11",
    "margin": "0.15,0.08",
}

# =====================================================
# LIDL LIETUVA
# =====================================================

# 1. Lidl - HIERARCHINĖ struktūra
g = Digraph("lidl_hierarchine", format="png")
g.attr(rankdir="TB", bgcolor="white", pad="0.3", nodesep="0.4", ranksep="0.6")
g.attr("node", **COMMON_NODE_ATTR, fillcolor="#e3f2fd", color="#1565c0")
g.attr("edge", color="#1565c0", arrowsize="0.7")

g.node("ceo", "Generalinis direktorius\n(Lidl Lietuva)", fillcolor="#1565c0", fontcolor="white")
g.node("ops", "Veiklos direktorius")
g.node("log", "Logistikos direktorius")
g.node("hr", "Personalo direktorius")
g.node("fin", "Finansų direktorius")

g.edge("ceo", "ops")
g.edge("ceo", "log")
g.edge("ceo", "hr")
g.edge("ceo", "fin")

# Regionų vadovai
g.node("r_vln", "Vilniaus regiono\nvadovas")
g.node("r_kau", "Kauno regiono\nvadovas")
g.node("r_kla", "Klaipėdos regiono\nvadovas")

g.edge("ops", "r_vln")
g.edge("ops", "r_kau")
g.edge("ops", "r_kla")

# Parduotuvių vadovai
g.node("pard1", "Parduotuvės\nvadovas (1)")
g.node("pard2", "Parduotuvės\nvadovas (2)")
g.node("pard3", "Parduotuvės\nvadovas (3)")

g.edge("r_vln", "pard1")
g.edge("r_kau", "pard2")
g.edge("r_kla", "pard3")

# Pamainų vadovai
g.node("pam1", "Pamainos\nvadovas")
g.node("pam2", "Pamainos\nvadovas")
g.node("pam3", "Pamainos\nvadovas")

g.edge("pard1", "pam1")
g.edge("pard2", "pam2")
g.edge("pard3", "pam3")

# Darbuotojai
g.node("d1", "Pardavėjai,\nkasininkai")
g.node("d2", "Pardavėjai,\nkasininkai")
g.node("d3", "Pardavėjai,\nkasininkai")

g.edge("pam1", "d1")
g.edge("pam2", "d2")
g.edge("pam3", "d3")

g.render(f"{OUTPUT_DIR}/lidl_1_hierarchine", cleanup=True)
print("OK lidl_1_hierarchine.png")

# 2. Lidl - FUNKCINĖ struktūra
g = Digraph("lidl_funkcine", format="png")
g.attr(rankdir="TB", bgcolor="white", pad="0.3", nodesep="0.5", ranksep="0.7")
g.attr("node", **COMMON_NODE_ATTR, fillcolor="#e3f2fd", color="#1565c0")
g.attr("edge", color="#1565c0", arrowsize="0.7")

g.node("ceo", "Generalinis direktorius", fillcolor="#1565c0", fontcolor="white")

# 6 funkciniai skyriai
g.node("f_pard", "Pardavimų\nskyrius")
g.node("f_log", "Logistikos\nskyrius")
g.node("f_mark", "Marketingo\nskyrius")
g.node("f_fin", "Finansų\nskyrius")
g.node("f_hr", "Personalo\nskyrius")
g.node("f_it", "IT skyrius")

g.edge("ceo", "f_pard")
g.edge("ceo", "f_log")
g.edge("ceo", "f_mark")
g.edge("ceo", "f_fin")
g.edge("ceo", "f_hr")
g.edge("ceo", "f_it")

# Skyrių darbuotojai
g.node("d_pard", "Parduotuvių\nvadovai\nKasininkai\nPardavėjai")
g.node("d_log", "Sandėlio\nspecialistai\nVairuotojai")
g.node("d_mark", "Marketingo\nspecialistai\nDizaineriai")
g.node("d_fin", "Buhalteriai\nFinansų\nanalitikai")
g.node("d_hr", "Personalo\nspecialistai\nMokymo vad.")
g.node("d_it", "Sistemų\nadministrat.\nProgramuotojai")

g.edge("f_pard", "d_pard")
g.edge("f_log", "d_log")
g.edge("f_mark", "d_mark")
g.edge("f_fin", "d_fin")
g.edge("f_hr", "d_hr")
g.edge("f_it", "d_it")

g.render(f"{OUTPUT_DIR}/lidl_2_funkcine", cleanup=True)
print("OK lidl_2_funkcine.png")

# 3. Lidl - STRUKTŪRA PAGAL PADALINIUS (geografinė)
g = Digraph("lidl_padaliniai", format="png")
g.attr(rankdir="TB", bgcolor="white", pad="0.3", nodesep="0.4", ranksep="0.6")
g.attr("node", **COMMON_NODE_ATTR, fillcolor="#e3f2fd", color="#1565c0")
g.attr("edge", color="#1565c0", arrowsize="0.7")

g.node("ceo", "Generalinis direktorius\n(Lidl Lietuva)", fillcolor="#1565c0", fontcolor="white")

# Centrinė buveinė
g.node("centras", "Centrinė buveinė\n(Marketingo, HR, Finansai, IT)",
       fillcolor="#bbdefb")

# 4 regioniniai padaliniai
g.node("p_vln", "Vilniaus padalinys")
g.node("p_kau", "Kauno padalinys")
g.node("p_kla", "Klaipėdos padalinys")
g.node("p_sia", "Šiaulių padalinys")

g.edge("ceo", "centras")
g.edge("ceo", "p_vln")
g.edge("ceo", "p_kau")
g.edge("ceo", "p_kla")
g.edge("ceo", "p_sia")

# Kiekviename regione - parduotuvės, sandėlis, vietos personalas
g.node("vln_p", "Parduotuvės\nSandėlis\nVietos personalas")
g.node("kau_p", "Parduotuvės\nSandėlis\nVietos personalas")
g.node("kla_p", "Parduotuvės\nSandėlis\nVietos personalas")
g.node("sia_p", "Parduotuvės\nSandėlis\nVietos personalas")

g.edge("p_vln", "vln_p")
g.edge("p_kau", "kau_p")
g.edge("p_kla", "kla_p")
g.edge("p_sia", "sia_p")

g.render(f"{OUTPUT_DIR}/lidl_3_padaliniai", cleanup=True)
print("OK lidl_3_padaliniai.png")


# =====================================================
# NORD SECURITY (NordVPN)
# =====================================================

# 4. NordVPN - PLOKŠČIOJI (Flatarchy)
g = Digraph("nord_plokscioji", format="png")
g.attr(rankdir="TB", bgcolor="white", pad="0.3", nodesep="0.4", ranksep="0.5")
g.attr("node", **COMMON_NODE_ATTR, fillcolor="#e8f5e9", color="#2e7d32")
g.attr("edge", color="#2e7d32", arrowsize="0.7")

g.node("ceo", "Vykdomasis vadovas (CEO)", fillcolor="#2e7d32", fontcolor="white")

# Tiesiogiai pavaldūs - mažai lygių
g.node("eng", "Inžinerijos\nkomanda")
g.node("prod", "Produkto\nkomanda")
g.node("mark", "Marketingo\nkomanda")
g.node("supp", "Klientų\npagalbos\nkomanda")
g.node("hr", "Personalo\nkomanda")

g.edge("ceo", "eng")
g.edge("ceo", "prod")
g.edge("ceo", "mark")
g.edge("ceo", "supp")
g.edge("ceo", "hr")

# Komandų nariai - tiesiogiai
g.node("eng_t", "Inžinieriai\n(savarankiški)")
g.node("prod_t", "Produkto\nspecialistai")
g.node("mark_t", "Marketologai")
g.node("supp_t", "Klientų\nkonsultantai")
g.node("hr_t", "HR\nspecialistai")

g.edge("eng", "eng_t")
g.edge("prod", "prod_t")
g.edge("mark", "mark_t")
g.edge("supp", "supp_t")
g.edge("hr", "hr_t")

g.render(f"{OUTPUT_DIR}/nordvpn_1_plokscioji", cleanup=True)
print("OK nordvpn_1_plokscioji.png")

# 5. NordVPN - MATRICINĖ
g = Digraph("nord_matricine", format="png")
g.attr(rankdir="LR", bgcolor="white", pad="0.3", nodesep="0.4", ranksep="0.7")
g.attr("node", shape="plaintext", fontname="DejaVu Sans")

# Naudojam HTML lentelę matricinei struktūrai vaizduoti
table_html = """<
<TABLE BORDER="0" CELLBORDER="1" CELLSPACING="0" CELLPADDING="8">
  <TR>
    <TD BGCOLOR="#388e3c"><FONT COLOR="white"><B>FUNKCIJOS / PRODUKTAI</B></FONT></TD>
    <TD BGCOLOR="#c8e6c9"><B>NordVPN</B></TD>
    <TD BGCOLOR="#c8e6c9"><B>NordPass</B></TD>
    <TD BGCOLOR="#c8e6c9"><B>NordLayer</B></TD>
    <TD BGCOLOR="#c8e6c9"><B>NordLocker</B></TD>
  </TR>
  <TR>
    <TD BGCOLOR="#c8e6c9"><B>Inžinerija</B></TD>
    <TD>Inžinieriai</TD>
    <TD>Inžinieriai</TD>
    <TD>Inžinieriai</TD>
    <TD>Inžinieriai</TD>
  </TR>
  <TR>
    <TD BGCOLOR="#c8e6c9"><B>Dizainas</B></TD>
    <TD>UX/UI</TD>
    <TD>UX/UI</TD>
    <TD>UX/UI</TD>
    <TD>UX/UI</TD>
  </TR>
  <TR>
    <TD BGCOLOR="#c8e6c9"><B>Marketingas</B></TD>
    <TD>Marketologai</TD>
    <TD>Marketologai</TD>
    <TD>Marketologai</TD>
    <TD>Marketologai</TD>
  </TR>
  <TR>
    <TD BGCOLOR="#c8e6c9"><B>Pardavimai</B></TD>
    <TD>Pardavimų sp.</TD>
    <TD>Pardavimų sp.</TD>
    <TD>Pardavimų sp.</TD>
    <TD>Pardavimų sp.</TD>
  </TR>
</TABLE>>"""

g.node("matrix", table_html)

# Pridedam paaiškinimą
g.attr("node", shape="box", style="rounded,filled", fillcolor="#e8f5e9", color="#2e7d32",
       fontsize="10")
g.node("paaiskinimas",
       "Kiekvienas darbuotojas turi du vadovus:\n"
       "1) funkcinis vadovas (pvz., inžinerijos vadovas)\n"
       "2) produkto vadovas (pvz., NordVPN vadovas)",
       fillcolor="#f1f8e9")

g.render(f"{OUTPUT_DIR}/nordvpn_2_matricine", cleanup=True)
print("OK nordvpn_2_matricine.png")

# 6. NordVPN - KOMANDINĖ
g = Digraph("nord_komandine", format="png")
g.attr(rankdir="TB", bgcolor="white", pad="0.3", nodesep="0.5", ranksep="0.6")
g.attr("node", **COMMON_NODE_ATTR, fillcolor="#e8f5e9", color="#2e7d32")
g.attr("edge", color="#2e7d32", arrowsize="0.7")

g.node("ceo", "Vykdomasis vadovas", fillcolor="#2e7d32", fontcolor="white")

# Savarankiškos kross-funkcinės komandos
g.node("k_vpn", "VPN produkto komanda", fillcolor="#a5d6a7")
g.node("k_pass", "NordPass komanda", fillcolor="#a5d6a7")
g.node("k_layer", "NordLayer komanda", fillcolor="#a5d6a7")
g.node("k_locker", "NordLocker komanda", fillcolor="#a5d6a7")

g.edge("ceo", "k_vpn")
g.edge("ceo", "k_pass")
g.edge("ceo", "k_layer")
g.edge("ceo", "k_locker")

# Kiekvienoje komandoje skirtingi specialistai
g.node("t_vpn", "Inžinieriai\nDizaineriai\nProduktų vadovas\nQA specialistas")
g.node("t_pass", "Inžinieriai\nDizaineriai\nProduktų vadovas\nQA specialistas")
g.node("t_layer", "Inžinieriai\nDizaineriai\nProduktų vadovas\nB2B vadybin.")
g.node("t_locker", "Inžinieriai\nDizaineriai\nProduktų vadovas\nQA specialistas")

g.edge("k_vpn", "t_vpn")
g.edge("k_pass", "t_pass")
g.edge("k_layer", "t_layer")
g.edge("k_locker", "t_locker")

g.render(f"{OUTPUT_DIR}/nordvpn_3_komandine", cleanup=True)
print("OK nordvpn_3_komandine.png")


# =====================================================
# SWEDBANK LIETUVA
# =====================================================

# 7. Swedbank - TINKLINĖ (Network)
g = Digraph("swed_tinkline", format="png")
g.attr(layout="circo", bgcolor="white", pad="0.3", nodesep="0.6")
g.attr("node", **COMMON_NODE_ATTR, fillcolor="#fff3e0", color="#e65100")
g.attr("edge", color="#e65100", arrowsize="0.7", arrowhead="none")

# Centras
g.node("centras", "Swedbank Lietuva\n(centrinė organizacija)",
       fillcolor="#e65100", fontcolor="white", shape="ellipse")

# Išoriniai partneriai
g.node("visa", "Visa /\nMastercard")
g.node("fintech", "FinTech\nįmonės")
g.node("mok", "Mokėjimo\nsistemos\n(SEPA, Swift)")
g.node("lb", "Lietuvos\nbankas")
g.node("ecb", "Europos\nCentrinis\nbankas (ECB)")
g.node("it", "IT paslaugų\ntiekėjai")
g.node("audit", "Audito\nįmonės\n(KPMG, EY)")
g.node("svf", "Saugumo ir\nrizikos\nvertintojai")

# Visi sujungti su centru
for node in ["visa", "fintech", "mok", "lb", "ecb", "it", "audit", "svf"]:
    g.edge("centras", node)

g.render(f"{OUTPUT_DIR}/swedbank_1_tinkline", cleanup=True)
print("OK swedbank_1_tinkline.png")

# 8. Swedbank - PROJEKTINĖ (Projectized)
g = Digraph("swed_projektine", format="png")
g.attr(rankdir="TB", bgcolor="white", pad="0.3", nodesep="0.5", ranksep="0.6")
g.attr("node", **COMMON_NODE_ATTR, fillcolor="#fff3e0", color="#e65100")
g.attr("edge", color="#e65100", arrowsize="0.7")

g.node("ceo", "Generalinis direktorius", fillcolor="#e65100", fontcolor="white")
g.node("ppm", "Projektų portfelio\nvaldymo biuras", fillcolor="#ffe0b2")

g.edge("ceo", "ppm")

# 4 projektai
g.node("p1", "Skaitmenizacijos\nprojektas", fillcolor="#ffcc80")
g.node("p2", "Kibernetinės saugos\nprojektas", fillcolor="#ffcc80")
g.node("p3", "Naujo mobiliosios\nbankininkystės\nprojektas", fillcolor="#ffcc80")
g.node("p4", "Verslo klientų\nplatformos\nprojektas", fillcolor="#ffcc80")

g.edge("ppm", "p1")
g.edge("ppm", "p2")
g.edge("ppm", "p3")
g.edge("ppm", "p4")

# Kiekvienas projektas turi savo komandą
g.node("k1", "Projekto vadovas\nIT inžinieriai\nAnalitikai\nDizaineriai")
g.node("k2", "Projekto vadovas\nKibersaugos sp.\nIT inžinieriai\nRizikos analit.")
g.node("k3", "Projekto vadovas\nMobilūs kūrėjai\nUX dizaineriai\nQA specialistai")
g.node("k4", "Projekto vadovas\nB2B specialistai\nIT inžinieriai\nVerslo analit.")

g.edge("p1", "k1")
g.edge("p2", "k2")
g.edge("p3", "k3")
g.edge("p4", "k4")

g.render(f"{OUTPUT_DIR}/swedbank_2_projektine", cleanup=True)
print("OK swedbank_2_projektine.png")

# 9. Swedbank - FUNKCINĖ
g = Digraph("swed_funkcine", format="png")
g.attr(rankdir="TB", bgcolor="white", pad="0.3", nodesep="0.5", ranksep="0.7")
g.attr("node", **COMMON_NODE_ATTR, fillcolor="#fff3e0", color="#e65100")
g.attr("edge", color="#e65100", arrowsize="0.7")

g.node("ceo", "Generalinis direktorius", fillcolor="#e65100", fontcolor="white")

# 7 funkciniai skyriai
g.node("priv", "Privačių klientų\nskyrius")
g.node("ver", "Verslo klientų\nskyrius")
g.node("it", "IT skyrius")
g.node("riz", "Rizikos valdymo\nskyrius")
g.node("ati", "Atitikties\nskyrius")
g.node("hr", "Personalo\nskyrius")
g.node("fin", "Finansų\nskyrius")

g.edge("ceo", "priv")
g.edge("ceo", "ver")
g.edge("ceo", "it")
g.edge("ceo", "riz")
g.edge("ceo", "ati")
g.edge("ceo", "hr")
g.edge("ceo", "fin")

# Skyrių komandos
g.node("d_priv", "Filialai\nKlientų\nkonsultantai\nKredito sp.")
g.node("d_ver", "B2B vadybin.\nVerslo analit.\nKredito eksp.")
g.node("d_it", "Sistemų\nadministrat.\nProgramuotojai\nDevOps")
g.node("d_riz", "Rizikos\nanalitikai\nPiniginių srautų\nspecialistai")
g.node("d_ati", "Atitikties\nspecialistai\nTeisininkai\nAML pareigūnai")
g.node("d_hr", "Personalo\nspecialistai\nMokymai\nKarjera")
g.node("d_fin", "Buhalteriai\nFinansų\nplanavimas\nAtaskaitos")

g.edge("priv", "d_priv")
g.edge("ver", "d_ver")
g.edge("it", "d_it")
g.edge("riz", "d_riz")
g.edge("ati", "d_ati")
g.edge("hr", "d_hr")
g.edge("fin", "d_fin")

g.render(f"{OUTPUT_DIR}/swedbank_3_funkcine", cleanup=True)
print("OK swedbank_3_funkcine.png")

print("\nVisos 9 diagramos sukurtos sėkmingai!")
