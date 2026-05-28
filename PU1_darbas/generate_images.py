"""
Generuoja du paveikslus PU1 darbui:
1 pav. - Pigu.lt ir Omniva sąveikos sekų diagrama
2 pav. - Supaprastinta e. prekybos ir logistikos ER diagrama
"""

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Rectangle
from matplotlib.lines import Line2D


# Bendri formatavimo nustatymai - juoda/balta akademiniam stiliui
plt.rcParams['font.family'] = 'serif'
plt.rcParams['font.serif'] = ['DejaVu Serif', 'Liberation Serif', 'Times New Roman']
plt.rcParams['font.size'] = 10


# =====================================================
# 1 pav. - Sekų diagrama: Pigu.lt ↔ Omniva sąveika
# =====================================================
fig, ax = plt.subplots(figsize=(11, 7))
ax.set_xlim(0, 12)
ax.set_ylim(0, 11)
ax.axis('off')

# Aktorių pozicijos (x koordinatės)
actors = [
    (1.5, "Klientas\n(naršyklė)"),
    (4.5, "Pigu.lt\nInformacinė sistema"),
    (7.5, "Omniva\nAPI sąsaja"),
    (10.5, "Omniva\nDuomenų bazė"),
]

# Aktorių antraštės (dėžutės viršuje)
for x, name in actors:
    box = FancyBboxPatch((x - 1.1, 9.7), 2.2, 0.9,
                          boxstyle="round,pad=0.05",
                          edgecolor='black', facecolor='white', linewidth=1.2)
    ax.add_patch(box)
    ax.text(x, 10.15, name, ha='center', va='center', fontsize=9.5, fontweight='bold')

# Vertikalios "lifelines" (gyvenimo linijos)
for x, _ in actors:
    ax.plot([x, x], [9.7, 0.5], color='gray', linestyle=(0, (3, 3)), linewidth=0.8)


def draw_message(y, x_from, x_to, label, dashed=False, return_arrow=False):
    """Piešia žinutės rodyklę su tekstu virš jos."""
    style = '<|-' if return_arrow else '-|>'
    arrow = FancyArrowPatch((x_from, y), (x_to, y),
                             arrowstyle=style,
                             mutation_scale=15,
                             linewidth=1.3,
                             linestyle='--' if dashed else '-',
                             color='black')
    ax.add_patch(arrow)
    mid_x = (x_from + x_to) / 2
    ax.text(mid_x, y + 0.15, label, ha='center', va='bottom', fontsize=9)


# Žinutės (iš viršaus į apačią)
draw_message(8.7, 1.5, 4.5, "1. Užsako prekę, pasirenka Omniva paštomatą")
draw_message(7.7, 4.5, 7.5, "2. POST /siuntos (siuntimo duomenys)")
draw_message(6.7, 7.5, 10.5, "3. INSERT INTO siunta(...)")
draw_message(5.7, 10.5, 7.5, "4. Grąžina tracking_id", dashed=True, return_arrow=False)
draw_message(4.7, 7.5, 4.5, '5. {"tracking_id": "OMX..."}', dashed=True, return_arrow=False)
draw_message(3.7, 4.5, 1.5, "6. Užsakymo patvirtinimas + sekimo nuoroda", dashed=True, return_arrow=False)

# Aktyvinimo (activation) dėžutės ant lifelines
def activation(x, y_top, y_bot):
    rect = Rectangle((x - 0.12, y_bot), 0.24, y_top - y_bot,
                     facecolor='lightgray', edgecolor='black', linewidth=0.8)
    ax.add_patch(rect)


activation(4.5, 8.7, 3.7)   # Pigu.lt aktyvi
activation(7.5, 7.7, 4.7)   # Omniva API aktyvi
activation(10.5, 6.7, 5.7)  # Omniva DB aktyvi

# Antraštė
ax.text(6, 0.15, "Pastaba: ištisinė rodyklė – sinchroninis kvietimas; punktyrinė – atsakymas.",
        ha='center', va='center', fontsize=8, style='italic', color='dimgray')

plt.tight_layout()
plt.savefig('/projects/sandbox/PU1_darbas/img1_seku_diagrama.png',
            dpi=200, bbox_inches='tight', facecolor='white')
plt.close()
print("Sukurta: img1_seku_diagrama.png")


# =====================================================
# 2 pav. - ER diagrama (supaprastinta)
# =====================================================
fig, ax = plt.subplots(figsize=(11, 7.5))
ax.set_xlim(0, 12)
ax.set_ylim(0, 9)
ax.axis('off')


def entity_box(x, y, title, attrs, w=2.6, color='white'):
    """Piešia esybės dėžutę su antrašte ir atributais."""
    n_attrs = len(attrs)
    h = 0.6 + 0.35 * n_attrs

    # Antraštė
    head = Rectangle((x - w/2, y + h - 0.55), w, 0.55,
                     facecolor='#d9d9d9', edgecolor='black', linewidth=1.2)
    ax.add_patch(head)
    ax.text(x, y + h - 0.275, title, ha='center', va='center',
            fontsize=10, fontweight='bold')

    # Atributų skiltis
    body = Rectangle((x - w/2, y), w, h - 0.55,
                     facecolor=color, edgecolor='black', linewidth=1.2)
    ax.add_patch(body)
    for i, attr in enumerate(attrs):
        prefix = ""
        text = attr
        if "PK" in attr:
            text = attr.replace(" (PK)", "")
            prefix = "[PK] "
        elif "FK" in attr:
            text = attr.replace(" (FK)", "")
            prefix = "[FK] "
        ax.text(x - w/2 + 0.15, y + h - 0.55 - 0.25 - i * 0.35,
                f"{prefix}{text}",
                ha='left', va='center', fontsize=8.8,
                fontweight='bold' if 'PK' in attr else 'normal',
                style='italic' if 'FK' in attr else 'normal')
    return (x, y, w, h)


def draw_relation(e1, e2, card1="1", card2="N", label=""):
    """Piešia liniją tarp dviejų esybių su kardinalumais."""
    x1, y1, w1, h1 = e1
    x2, y2, w2, h2 = e2

    # Sujungimo taškai (centrų vidurys)
    cx1, cy1 = x1, y1 + h1 / 2
    cx2, cy2 = x2, y2 + h2 / 2

    # Linija
    ax.plot([cx1, cx2], [cy1, cy2], color='black', linewidth=1.2)

    # Kardinalumai
    # Pirmajame gale
    dx, dy = cx2 - cx1, cy2 - cy1
    length = (dx**2 + dy**2)**0.5
    ux, uy = dx / length, dy / length
    offset = 0.4 + max(w1, w2) / 4
    ax.text(cx1 + ux * offset, cy1 + uy * offset + 0.15,
            card1, fontsize=9, fontweight='bold', ha='center')
    ax.text(cx2 - ux * offset, cy2 - uy * offset + 0.15,
            card2, fontsize=9, fontweight='bold', ha='center')

    if label:
        mx, my = (cx1 + cx2) / 2, (cy1 + cy2) / 2
        ax.text(mx, my + 0.25, label, fontsize=8.5,
                ha='center', va='center', style='italic',
                bbox=dict(facecolor='white', edgecolor='none', pad=2))


# Esybės
e_klientas = entity_box(1.7, 5.5, "KLIENTAS", [
    "klientoID (PK)",
    "vardas",
    "pavarde",
    "email",
    "telefonas",
])

e_uzsakymas = entity_box(5.5, 5.5, "UZSAKYMAS", [
    "uzsakymoID (PK)",
    "klientoID (FK)",
    "data",
    "bendra_suma",
    "busena",
])

e_eilute = entity_box(9.5, 5.5, "UZSAKYMO_EILUTE", [
    "eilutesID (PK)",
    "uzsakymoID (FK)",
    "produktoID (FK)",
    "kiekis",
    "kaina_uz_vnt",
])

e_produktas = entity_box(9.5, 1.0, "PRODUKTAS", [
    "produktoID (PK)",
    "pavadinimas",
    "kaina",
    "kiekis_sandelyje",
])

e_siunta = entity_box(2.5, 1.0, "SIUNTA", [
    "siuntosID (PK)",
    "uzsakymoID (FK)",
    "tracking_nr",
    "pastomato_adresas",
    "busena",
])

# Ryšiai
draw_relation(e_klientas, e_uzsakymas, "1", "N", "pateikia")
draw_relation(e_uzsakymas, e_eilute, "1", "N", "turi")
draw_relation(e_eilute, e_produktas, "N", "1", "yra")
draw_relation(e_uzsakymas, e_siunta, "1", "1", "pristatoma")

# Antraštė
ax.text(6, 8.6, "E. prekybos ir logistikos duomenų bazės ER diagrama (supaprastinta)",
        ha='center', va='center', fontsize=11, fontweight='bold')
ax.text(6, 0.2,
        "Pastaba: [PK] - pirminis raktas, [FK] - isorinis raktas.",
        ha='center', va='center', fontsize=8, style='italic', color='dimgray')

plt.tight_layout()
plt.savefig('/projects/sandbox/PU1_darbas/img2_er_diagrama.png',
            dpi=200, bbox_inches='tight', facecolor='white')
plt.close()
print("Sukurta: img2_er_diagrama.png")
