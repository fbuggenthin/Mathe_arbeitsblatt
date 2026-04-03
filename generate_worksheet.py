#!/usr/bin/env python3
"""
Geometrie-Arbeitsblatt Generator – Klasse 10 (v2)
Verwendet echte mpl_toolkits.mplot3d-Achsen mit Poly3DCollection
fuer mathematisch korrekte 3D-Koerper.
"""

import io
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import Polygon as MplPolygon
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ---------------------------------------------------------------------------
# Farb-Palette
# ---------------------------------------------------------------------------
C2D_FACE  = "#AED6F1"
C2D_EDGE  = "#1A5276"
C3D_TOP   = "#FDEBD0"
C3D_FRONT = "#F0B27A"
C3D_SIDE  = "#CA6F1E"
C3D_EDGE  = "#784212"
CH_COLOR  = "#1E8449"
CDIM      = "#1A5276"
CARROW    = "#7D3C98"
CFORMULA  = "#C0392B"

DPI   = 150
ELEV  = 22
AZIM  = -52


# ===========================================================================
# 3D-Hilfsfunktionen
# ===========================================================================

def style_3d_ax(ax, title, formula):
    """Entfernt Gitter/Achsen, setzt Titel und Formel."""
    ax.set_title(title, fontsize=12, fontweight="bold", color=C3D_EDGE, pad=6)
    ax.text2D(0.5, 0.0, formula, transform=ax.transAxes,
              ha="center", va="bottom", fontsize=12,
              color=CFORMULA, fontweight="bold")
    for pane in (ax.xaxis.pane, ax.yaxis.pane, ax.zaxis.pane):
        pane.fill = False
        pane.set_edgecolor("white")
    ax.set_xticks([])
    ax.set_yticks([])
    ax.set_zticks([])
    ax.grid(False)
    ax.set_axis_off()
    ax.view_init(elev=ELEV, azim=AZIM)


def add_label_3d(ax, p1, p2, label, offset=(0.05, 0), color=CH_COLOR):
    """Pfeil mit Label zwischen zwei 3D-Punkten."""
    ax.plot3D(*zip(p1, p2), color=color, lw=1.8)
    mid = [(a + b) / 2 for a, b in zip(p1, p2)]
    ax.text(mid[0] + offset[0], mid[1] + offset[1], mid[2],
            label, fontsize=10, color=color, fontweight="bold")


def cuboid_poly3d(ax, a=1.0, b=1.0, h=1.0):
    """Zeichnet einen Quader mit korrekten sichtbaren Flaechen."""
    v = np.array([
        [0, 0, 0], [a, 0, 0], [a, b, 0], [0, b, 0],  # Boden 0-3
        [0, 0, h], [a, 0, h], [a, b, h], [0, b, h],  # Decke 4-7
    ])
    # Sichtbare Flaechen fuer Ansicht von vorne-rechts-oben
    faces  = [
        [v[0], v[1], v[5], v[4]],  # Vorne
        [v[1], v[2], v[6], v[5]],  # Rechts
        [v[4], v[5], v[6], v[7]],  # Oben
    ]
    colors = [C3D_FRONT, C3D_SIDE, C3D_TOP]
    poly = Poly3DCollection(faces, zsort="min", alpha=1.0)
    poly.set_facecolor(colors)
    poly.set_edgecolor(C3D_EDGE)
    poly.set_linewidth(1.5)
    ax.add_collection3d(poly)
    # Versteckte Kanten gestrichelt
    hidden = [(v[2], v[3]), (v[3], v[0]), (v[3], v[7])]
    for p1, p2 in hidden:
        ax.plot3D(*zip(p1, p2), color=C3D_EDGE, lw=0.9,
                  linestyle="--", alpha=0.45)
    # Sichtbare Kanten vorne
    visible = [
        (v[0], v[1]), (v[1], v[2]), (v[0], v[4]),
        (v[1], v[5]), (v[2], v[6]), (v[4], v[5]),
        (v[5], v[6]), (v[4], v[7]), (v[6], v[7]),
    ]
    for p1, p2 in visible:
        ax.plot3D(*zip(p1, p2), color=C3D_EDGE, lw=1.5)
    ax.set_xlim(0, a)
    ax.set_ylim(0, b)
    ax.set_zlim(0, h)
    return v


def cylinder_poly3d(ax, r=1.0, h=1.5, n=60):
    """Zeichnet einen Zylinder als Poly3DCollection."""
    theta = np.linspace(0, 2 * np.pi, n)
    # Boden- und Deckelkreis
    bx = r * np.cos(theta)
    by = r * np.sin(theta)
    bz = np.zeros(n)
    tx, ty, tz = bx.copy(), by.copy(), np.full(n, h)

    # Seitenmantelflächen (schmale Trapeze)
    side = []
    for i in range(n - 1):
        side.append([
            [bx[i], by[i], 0],
            [bx[i+1], by[i+1], 0],
            [tx[i+1], ty[i+1], h],
            [tx[i],   ty[i],   h],
        ])
    side_poly = Poly3DCollection(side, alpha=0.85, zsort="min")
    side_poly.set_facecolor(C3D_FRONT)
    side_poly.set_edgecolor("none")
    ax.add_collection3d(side_poly)

    # Deckelscheibe
    top_verts = [[[tx[i], ty[i], h] for i in range(n)]]
    top_poly = Poly3DCollection(top_verts, alpha=1.0)
    top_poly.set_facecolor(C3D_TOP)
    top_poly.set_edgecolor(C3D_EDGE)
    top_poly.set_linewidth(1.2)
    ax.add_collection3d(top_poly)

    # Bodenscheibe (halb sichtbar)
    bot_verts = [[[bx[i], by[i], 0] for i in range(n)]]
    bot_poly = Poly3DCollection(bot_verts, alpha=0.5)
    bot_poly.set_facecolor(C3D_SIDE)
    bot_poly.set_edgecolor(C3D_EDGE)
    bot_poly.set_linewidth(1.2)
    ax.add_collection3d(bot_poly)

    # Umriss-Linien fuer Kanten
    ax.plot3D(bx, by, bz, color=C3D_EDGE, lw=1.2)
    ax.plot3D(tx, ty, tz, color=C3D_EDGE, lw=1.5)

    ax.set_xlim(-r, r)
    ax.set_ylim(-r, r)
    ax.set_zlim(0, h)
    ax.set_box_aspect([2*r, 2*r, h])


def cone_poly3d(ax, r=1.0, h=1.6, n=60):
    """Zeichnet einen Kegel als Poly3DCollection."""
    theta = np.linspace(0, 2 * np.pi, n)
    bx = r * np.cos(theta)
    by = r * np.sin(theta)

    # Seitenflaechen (Dreiecke zur Spitze)
    apex = np.array([0, 0, h])
    side = []
    for i in range(n - 1):
        side.append([
            [bx[i], by[i], 0],
            [bx[i+1], by[i+1], 0],
            apex,
        ])
    side_poly = Poly3DCollection(side, alpha=0.85, zsort="min")
    side_poly.set_facecolor(C3D_FRONT)
    side_poly.set_edgecolor("none")
    ax.add_collection3d(side_poly)

    # Bodenscheibe
    bot_verts = [[[bx[i], by[i], 0] for i in range(n)]]
    bot_poly = Poly3DCollection(bot_verts, alpha=0.6)
    bot_poly.set_facecolor(C3D_SIDE)
    bot_poly.set_edgecolor(C3D_EDGE)
    bot_poly.set_linewidth(1.2)
    ax.add_collection3d(bot_poly)

    ax.plot3D(bx, by, np.zeros(n), color=C3D_EDGE, lw=1.2)

    ax.set_xlim(-r, r)
    ax.set_ylim(-r, r)
    ax.set_zlim(0, h)
    ax.set_box_aspect([2*r, 2*r, h])


def triangular_prism_poly3d(ax, g=1.5, h_a=1.0, depth=1.2):
    """Zeichnet ein Dreiecksprisma."""
    # Dreieck-Grundriss: gleichschenklig
    tri = np.array([
        [0,      0,    0],
        [g,      0,    0],
        [g/2, h_a,    0],
    ])
    tri_top = tri.copy()
    tri_top[:, 2] = depth

    # Flaechen
    front  = [tri[0], tri[1], tri[2]]           # vorderes Dreieck
    back   = [tri_top[0], tri_top[1], tri_top[2]]
    bot    = [tri[0], tri[1], tri_top[1], tri_top[0]]
    right  = [tri[1], tri[2], tri_top[2], tri_top[1]]
    left   = [tri[2], tri[0], tri_top[0], tri_top[2]]

    faces  = [bot, right, front]
    colors = [C3D_FRONT, C3D_SIDE, C3D_TOP]
    poly = Poly3DCollection(faces, alpha=1.0, zsort="min")
    poly.set_facecolor(colors)
    poly.set_edgecolor(C3D_EDGE)
    poly.set_linewidth(1.5)
    ax.add_collection3d(poly)

    # Hintere Kanten gestrichelt
    for p1, p2 in [(tri_top[0], tri_top[1]),
                   (tri_top[1], tri_top[2]),
                   (tri_top[2], tri_top[0]),
                   (tri[0], tri_top[0]),
                   (tri[1], tri_top[1]),
                   (tri[2], tri_top[2])]:
        ax.plot3D(*zip(p1, p2), color=C3D_EDGE, lw=0.9, linestyle="--", alpha=0.4)

    # Sichtbare Kanten vorne
    for p1, p2 in [(tri[0], tri[1]), (tri[1], tri[2]), (tri[2], tri[0])]:
        ax.plot3D(*zip(p1, p2), color=C3D_EDGE, lw=1.5)

    ax.set_xlim(0, g)
    ax.set_ylim(0, h_a)
    ax.set_zlim(0, depth)
    ax.set_box_aspect([g, h_a, depth])


# ===========================================================================
# 2D-Zeichenfunktionen
# ===========================================================================

def draw_2d_square(ax):
    ax.set_xlim(-0.25, 1.45)
    ax.set_ylim(-0.45, 1.45)
    ax.set_aspect("equal")
    ax.axis("off")
    sq = mpatches.Rectangle((0.1, 0.1), 1.0, 1.0,
                             lw=2, edgecolor=C2D_EDGE, facecolor=C2D_FACE)
    ax.add_patch(sq)
    # Bemaßung
    ax.annotate("", xy=(1.1, -0.05), xytext=(0.1, -0.05),
                arrowprops=dict(arrowstyle="<->", color=CDIM, lw=1.3))
    ax.text(0.6, -0.17, "$a$", ha="center", fontsize=12,
            color=CDIM, fontweight="bold")
    ax.annotate("", xy=(-0.05, 1.1), xytext=(-0.05, 0.1),
                arrowprops=dict(arrowstyle="<->", color=CDIM, lw=1.3))
    ax.text(-0.18, 0.6, "$a$", ha="center", fontsize=12,
            color=CDIM, fontweight="bold")
    ax.set_title("Quadrat", fontsize=12, fontweight="bold", color=C2D_EDGE)
    ax.text(0.6, -0.37, r"$A = a^2$", ha="center", fontsize=13,
            color=CFORMULA, fontweight="bold")


def draw_2d_rectangle(ax):
    ax.set_xlim(-0.3, 1.85)
    ax.set_ylim(-0.45, 1.25)
    ax.set_aspect("equal")
    ax.axis("off")
    rect = mpatches.Rectangle((0.1, 0.1), 1.5, 0.9,
                               lw=2, edgecolor=C2D_EDGE, facecolor=C2D_FACE)
    ax.add_patch(rect)
    ax.annotate("", xy=(1.6, -0.05), xytext=(0.1, -0.05),
                arrowprops=dict(arrowstyle="<->", color=CDIM, lw=1.3))
    ax.text(0.85, -0.17, "$a$", ha="center", fontsize=12,
            color=CDIM, fontweight="bold")
    ax.annotate("", xy=(-0.05, 1.0), xytext=(-0.05, 0.1),
                arrowprops=dict(arrowstyle="<->", color=CDIM, lw=1.3))
    ax.text(-0.18, 0.55, "$b$", ha="center", fontsize=12,
            color=CDIM, fontweight="bold")
    ax.set_title("Rechteck", fontsize=12, fontweight="bold", color=C2D_EDGE)
    ax.text(0.85, -0.37, r"$A = a \cdot b$", ha="center", fontsize=13,
            color=CFORMULA, fontweight="bold")


def draw_2d_circle(ax):
    ax.set_xlim(-1.5, 1.5)
    ax.set_ylim(-1.5, 1.5)
    ax.set_aspect("equal")
    ax.axis("off")
    circ = plt.Circle((0, 0), 1.0, lw=2,
                       edgecolor=C2D_EDGE, facecolor=C2D_FACE)
    ax.add_patch(circ)
    ax.plot([0, 1.0], [0, 0], color=CDIM, lw=1.8)
    ax.text(0.52, 0.1, "$r$", fontsize=12, color=CDIM, fontweight="bold")
    ax.set_title("Kreis", fontsize=12, fontweight="bold", color=C2D_EDGE)
    ax.text(0, -1.35, r"$A = \pi \cdot r^2$", ha="center", fontsize=13,
            color=CFORMULA, fontweight="bold")


def draw_2d_triangle(ax):
    ax.set_xlim(-0.25, 1.85)
    ax.set_ylim(-0.45, 1.35)
    ax.set_aspect("equal")
    ax.axis("off")
    pts = np.array([[0.1, 0.1], [1.6, 0.1], [0.85, 1.1]])
    tri = MplPolygon(pts, closed=True, lw=2,
                     edgecolor=C2D_EDGE, facecolor=C2D_FACE)
    ax.add_patch(tri)
    # Grundlinie
    ax.annotate("", xy=(1.6, -0.05), xytext=(0.1, -0.05),
                arrowprops=dict(arrowstyle="<->", color=CDIM, lw=1.3))
    ax.text(0.85, -0.17, "$g$", ha="center", fontsize=12,
            color=CDIM, fontweight="bold")
    # Hoehe gestrichelt
    ax.plot([0.85, 0.85], [0.1, 1.1], color=CH_COLOR, lw=1.5, ls="--")
    ax.annotate("", xy=(1.0, 1.1), xytext=(1.0, 0.1),
                arrowprops=dict(arrowstyle="<->", color=CH_COLOR, lw=1.3))
    ax.text(1.12, 0.6, "$h_A$", fontsize=11, color=CH_COLOR, fontweight="bold")
    ax.set_title("Dreieck", fontsize=12, fontweight="bold", color=C2D_EDGE)
    ax.text(0.85, -0.37, r"$A = \frac{1}{2} \cdot g \cdot h_A$",
            ha="center", fontsize=13, color=CFORMULA, fontweight="bold")


# ===========================================================================
# Figuren-Paare
# ===========================================================================

def make_pair_figure(draw_2d, draw_3d, figsize=(13, 5.2)):
    fig = plt.figure(figsize=figsize, facecolor="white")
    gs = fig.add_gridspec(1, 3, width_ratios=[2.2, 0.7, 2.2],
                          left=0.03, right=0.97, bottom=0.12, top=0.92,
                          wspace=0.05)
    ax2d  = fig.add_subplot(gs[0])
    ax_ar = fig.add_subplot(gs[1])
    ax3d  = fig.add_subplot(gs[2], projection="3d")

    draw_2d(ax2d)

    # Pfeil-Achse
    ax_ar.set_xlim(0, 1)
    ax_ar.set_ylim(0, 1)
    ax_ar.axis("off")
    ax_ar.annotate("", xy=(0.88, 0.5), xytext=(0.12, 0.5),
                   arrowprops=dict(arrowstyle="-|>", color=CARROW,
                                   lw=2.5, mutation_scale=20))
    ax_ar.text(0.5, 0.62, "+ Höhe $h$", ha="center", va="bottom",
               fontsize=10, color=CH_COLOR, fontweight="bold")

    draw_3d(ax3d)

    buf = io.BytesIO()
    fig.savefig(buf, format="jpeg", dpi=DPI, bbox_inches="tight",
                facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------- 1. Quadrat / Würfel ----------
def _3d_cube(ax):
    v = cuboid_poly3d(ax, a=1.0, b=1.0, h=1.0)
    # Hoehenbeschriftung
    ax.plot3D([1, 1], [0, 0], [0, 1], color=CH_COLOR, lw=1.8)
    ax.text(1.12, 0.0, 0.5, "$a$", fontsize=10, color=CH_COLOR, fontweight="bold")
    ax.set_box_aspect([1, 1, 1])
    style_3d_ax(ax, "Würfel", r"$V = a^2 \cdot a = a^3$")


# ---------- 2. Rechteck / Quader ----------
def _3d_cuboid(ax):
    v = cuboid_poly3d(ax, a=1.5, b=1.0, h=1.0)
    ax.plot3D([1.5, 1.5], [0, 0], [0, 1.0], color=CH_COLOR, lw=1.8)
    ax.text(1.62, 0.0, 0.5, "$h$", fontsize=10, color=CH_COLOR, fontweight="bold")
    ax.set_box_aspect([1.5, 1, 1])
    style_3d_ax(ax, "Quader", r"$V = a \cdot b \cdot h$")


# ---------- 3. Kreis / Zylinder ----------
def _3d_cylinder(ax):
    cylinder_poly3d(ax, r=1.0, h=1.5)
    # Radius auf Deckel
    ax.plot3D([0, 1], [0, 0], [1.5, 1.5], color=CDIM, lw=1.8)
    ax.text(0.5, 0.1, 1.62, "$r$", fontsize=10, color=CDIM, fontweight="bold")
    # Hoehe
    ax.plot3D([1.15, 1.15], [0, 0], [0, 1.5], color=CH_COLOR, lw=1.8)
    ax.text(1.28, 0.0, 0.7, "$h$", fontsize=10, color=CH_COLOR, fontweight="bold")
    style_3d_ax(ax, "Zylinder", r"$V = \pi \cdot r^2 \cdot h$")


# ---------- 4. Dreieck / Prisma ----------
def _3d_prism(ax):
    triangular_prism_poly3d(ax, g=1.5, h_a=1.0, depth=1.2)
    # Tiefenpfeil
    ax.plot3D([0.75, 0.75], [0, 0], [0, 1.2], color=CH_COLOR, lw=1.8)
    ax.text(0.85, 0.05, 0.6, "$h$", fontsize=10, color=CH_COLOR, fontweight="bold")
    style_3d_ax(ax, "Dreiecksprisma",
                r"$V = \frac{1}{2} \cdot g \cdot h_A \cdot h$")


# ---------- 5. Kreis / Kegel ----------
def _3d_cone(ax):
    cone_poly3d(ax, r=1.0, h=1.6)
    ax.plot3D([0, 1], [0, 0], [0, 0], color=CDIM, lw=1.8)
    ax.text(0.5, 0.1, 0.05, "$r$", fontsize=10, color=CDIM, fontweight="bold")
    ax.plot3D([1.15, 1.15], [0, 0], [0, 1.6], color=CH_COLOR, lw=1.8)
    ax.text(1.28, 0.0, 0.8, "$h$", fontsize=10, color=CH_COLOR, fontweight="bold")
    style_3d_ax(ax, "Kegel", r"$V = \frac{1}{3} \cdot \pi \cdot r^2 \cdot h$")


# ===========================================================================
# Inhaltsliste
# ===========================================================================

SHAPES = [
    (
        "Quadrat → Würfel",
        "Das Quadrat (A = a²) wird durch Strecken um die Höhe a zum Würfel: V = a³.",
        draw_2d_square, _3d_cube,
    ),
    (
        "Rechteck → Quader",
        "Das Rechteck (A = a·b) wird durch Hinzufügen der Höhe h zum Quader: V = a·b·h.",
        draw_2d_rectangle, _3d_cuboid,
    ),
    (
        "Kreis → Zylinder",
        "Der Kreis (A = π·r²) wird durch eine Höhe h zum Zylinder: V = π·r²·h.",
        draw_2d_circle, _3d_cylinder,
    ),
    (
        "Dreieck → Dreiecksprisma",
        "Das Dreieck (A = ½·g·h_A) wird durch die Tiefe h zum Dreiecksprisma: V = ½·g·h_A·h.",
        draw_2d_triangle, _3d_prism,
    ),
    (
        "Kreis → Kegel",
        "Beim Kegel verjüngt sich der Körper zur Spitze – daher der Faktor ⅓: V = ⅓·π·r²·h.",
        draw_2d_circle, _3d_cone,
    ),
]


# ===========================================================================
# Word-Dokument
# ===========================================================================

def add_rule(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "AAAAAA")
    pBdr.append(bottom)
    pPr.append(pBdr)


def build_document(output_path="geometrie_arbeitsblatt.docx"):
    doc = Document()

    # A4 Querformat
    section = doc.sections[0]
    section.page_width    = Cm(29.7)
    section.page_height   = Cm(21.0)
    section.left_margin   = Cm(1.8)
    section.right_margin  = Cm(1.8)
    section.top_margin    = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    # Titel
    title = doc.add_heading("Geometrie: Vom Flächeninhalt zum Volumen", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph("Klasse 10  ·  Wie entsteht aus einer 2D-Figur ein 3D-Körper?")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].font.size = Pt(13)
    sub.runs[0].font.color.rgb = RGBColor(0x1A, 0x52, 0x76)

    intro = doc.add_paragraph(
        "Prinzip: Jeder raeumliche Koerper entsteht, indem man eine ebene Grundfigur "
        "um eine Hoehe h in den Raum 'schiebt' (Prisma, Zylinder) oder zuspitzt "
        "(Pyramide, Kegel). Es gilt: V = Grundflaeche x Hoehe (Prismen/Zylinder) "
        "bzw. V = 1/3 x Grundflaeche x Hoehe (Pyramiden/Kegel)."
    )
    intro.runs[0].font.size = Pt(11)

    for shape_title, explanation, fn2d, fn3d in SHAPES:
        add_rule(doc)
        h2 = doc.add_heading(shape_title, level=2)
        h2.runs[0].font.color.rgb = RGBColor(0x1A, 0x52, 0x76)

        p = doc.add_paragraph(explanation)
        p.runs[0].font.size = Pt(11)

        buf = make_pair_figure(fn2d, fn3d)
        pic_p = doc.add_paragraph()
        pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic_p.add_run().add_picture(buf, width=Inches(9.8))

    # Formel-Tabelle
    add_rule(doc)
    doc.add_heading("Zusammenfassung der Formeln", level=2)

    rows = [
        ("Figur / Körper",           "Fläche A",              "Volumen V"),
        ("Quadrat / Würfel",          "a²",                    "a³"),
        ("Rechteck / Quader",         "a · b",                 "a · b · h"),
        ("Kreis / Zylinder",          "π · r²",                "π · r² · h"),
        ("Dreieck / Dreiecksprisma",  "½ · g · h_A",           "½ · g · h_A · h"),
        ("Kreis / Kegel",             "π · r²",                "⅓ · π · r² · h"),
    ]
    table = doc.add_table(rows=len(rows), cols=3)
    table.style = "Light Shading Accent 1"
    for r_i, row_data in enumerate(rows):
        row = table.rows[r_i]
        for c_i, txt in enumerate(row_data):
            cell = row.cells[c_i]
            cell.text = txt
            run = cell.paragraphs[0].runs[0]
            run.font.size = Pt(11)
            if r_i == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    doc.save(output_path)
    print(f"Dokument gespeichert: {output_path}")
    return output_path


if __name__ == "__main__":
    build_document()
