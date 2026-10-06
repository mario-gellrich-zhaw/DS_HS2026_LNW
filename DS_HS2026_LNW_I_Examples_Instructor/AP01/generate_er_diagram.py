"""
Generate a clean ER diagram for cinema.db using matplotlib.
Layout: 3-column grid — no overlaps, generous spacing.
"""
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

fig, ax = plt.subplots(figsize=(22, 10))
ax.set_xlim(0, 22)
ax.set_ylim(2.5, 13.2)
ax.axis("off")
fig.patch.set_facecolor("#F8F9FA")

# ── colour palette ────────────────────────────────────────────────────────────
C_HEADER = "#2E4057"
C_ROW    = "#EAF0FB"
C_ROW2   = "#FFFFFF"
C_BORDER = "#2E4057"
C_PK     = "#E63946"
C_FK     = "#457B9D"
C_TEXT   = "#1D3557"

ROW_H  = 0.40
HDR_H  = 0.52
COL_W  = 3.0

# ── table definitions (name, x, y, columns) ──────────────────────────────────
# columns: (name, type, flag)  flag: "PK" | "FK" | ""
#
# Layout (3 columns, top → bottom):
#   Col 1 (x=0.8):  movies
#   Col 2 (x=8.0):  cinemas  →  screenings  →  tickets
#   Col 3 (x=15.5): staff (top)  /  customers (bottom, aligned with tickets)
#
TABLES = {
    "movies": dict(x=0.8, y=10.8, cols=[
        ("movieid",      "TEXT",     "PK"),
        ("title",        "TEXT",     ""),
        ("genre",        "TEXT",     ""),
        ("duration_min", "INTEGER",  ""),
        ("release_year", "INTEGER",  ""),
        ("director",     "TEXT",     ""),
        ("age_rating",   "INTEGER",  ""),
    ]),
    "cinemas": dict(x=8.0, y=12.8, cols=[
        ("cinemaid",    "TEXT",    "PK"),
        ("name",        "TEXT",    ""),
        ("city",        "TEXT",    ""),
        ("total_seats", "INTEGER", ""),
    ]),
    "staff": dict(x=15.5, y=12.8, cols=[
        ("staffid",   "TEXT", "PK"),
        ("firstname", "TEXT", ""),
        ("lastname",  "TEXT", ""),
        ("cinemaid",  "TEXT", "FK"),
        ("role",      "TEXT", ""),
        ("hire_date", "DATE", ""),
    ]),
    "screenings": dict(x=8.0, y=9.6, cols=[
        ("screeningid",        "TEXT",     "PK"),
        ("movieid",            "TEXT",     "FK"),
        ("cinemaid",           "TEXT",     "FK"),
        ("screening_datetime", "DATETIME", ""),
        ("base_price",         "REAL",     ""),
    ]),
    "customers": dict(x=15.5, y=6.5, cols=[
        ("customerid",      "TEXT", "PK"),
        ("firstname",       "TEXT", ""),
        ("lastname",        "TEXT", ""),
        ("email",           "TEXT", ""),
        ("birthdate",       "DATE", ""),
        ("membership_type", "TEXT", ""),
    ]),
    "tickets": dict(x=8.0, y=6.2, cols=[
        ("ticketid",          "TEXT",     "PK"),
        ("screeningid",       "TEXT",     "FK"),
        ("customerid",        "TEXT",     "FK"),
        ("purchase_datetime", "DATETIME", ""),
        ("seat_number",       "TEXT",     ""),
        ("paid_price",        "REAL",     ""),
    ]),
}


def draw_table(ax, name, x, y, cols):
    """Draw an entity box; return dict of column anchor points."""
    n_rows = len(cols)
    height = HDR_H + n_rows * ROW_H

    # shadow
    shadow = mpatches.FancyBboxPatch(
        (x + 0.06, y - height - 0.06), COL_W, height,
        boxstyle="round,pad=0.02", linewidth=0,
        facecolor="#CCCCCC", zorder=1
    )
    ax.add_patch(shadow)

    # header
    hdr = mpatches.FancyBboxPatch(
        (x, y - HDR_H), COL_W, HDR_H,
        boxstyle="round,pad=0.02", linewidth=1.5,
        edgecolor=C_BORDER, facecolor=C_HEADER, zorder=2
    )
    ax.add_patch(hdr)
    ax.text(x + COL_W / 2, y - HDR_H / 2, name,
            ha="center", va="center", fontsize=10.5,
            fontweight="bold", color="white", zorder=3)

    anchors = {}
    for i, (col_name, col_type, flag) in enumerate(cols):
        ry   = y - HDR_H - i * ROW_H
        fill = C_ROW if i % 2 == 0 else C_ROW2
        rect = mpatches.FancyBboxPatch(
            (x, ry - ROW_H), COL_W, ROW_H,
            boxstyle="square,pad=0", linewidth=0.5,
            edgecolor="#BBBBBB", facecolor=fill, zorder=2
        )
        ax.add_patch(rect)

        if flag == "PK":
            badge_col, badge_txt = C_PK, "PK"
        elif flag == "FK":
            badge_col, badge_txt = C_FK, "FK"
        else:
            badge_col, badge_txt = None, ""

        if badge_col:
            badge = mpatches.FancyBboxPatch(
                (x + 0.07, ry - ROW_H + 0.06), 0.40, ROW_H - 0.12,
                boxstyle="round,pad=0.02", linewidth=0,
                facecolor=badge_col, zorder=3
            )
            ax.add_patch(badge)
            ax.text(x + 0.07 + 0.20, ry - ROW_H / 2, badge_txt,
                    ha="center", va="center", fontsize=6.5,
                    fontweight="bold", color="white", zorder=4)
            name_x = x + 0.54
        else:
            name_x = x + 0.14

        ax.text(name_x, ry - ROW_H / 2, col_name,
                ha="left", va="center", fontsize=8.5, color=C_TEXT, zorder=4)
        ax.text(x + COL_W - 0.10, ry - ROW_H / 2, col_type,
                ha="right", va="center", fontsize=7.5,
                color="#666666", style="italic", zorder=4)

        anchors[col_name] = {
            "right": (x + COL_W, ry - ROW_H / 2),
            "left":  (x,          ry - ROW_H / 2),
            "top":   (x + COL_W / 2, y),
            "bot":   (x + COL_W / 2, y - HDR_H - n_rows * ROW_H),
        }
    return anchors


all_anchors = {}
for tbl_name, info in TABLES.items():
    a = draw_table(ax, tbl_name, info["x"], info["y"], info["cols"])
    all_anchors[tbl_name] = a


# ── relationship arrows ───────────────────────────────────────────────────────
def arrow(ax, p1, p2, label="1 : n", rad=0.0):
    mx = (p1[0] + p2[0]) / 2
    my = (p1[1] + p2[1]) / 2
    ax.annotate("", xy=p2, xytext=p1,
                arrowprops=dict(
                    arrowstyle="-|>", color=C_FK,
                    lw=1.6,
                    connectionstyle=f"arc3,rad={rad}"
                ), zorder=5)
    # label offset: push perpendicular to the arrow direction
    dx = p2[0] - p1[0]
    dy = p2[1] - p1[1]
    length = (dx**2 + dy**2) ** 0.5
    if length > 0:
        perp_x = -dy / length * 0.22
        perp_y =  dx / length * 0.22
    else:
        perp_x, perp_y = 0, 0.22
    ax.text(mx + perp_x, my + perp_y, label,
            ha="center", va="center", fontsize=8,
            color=C_FK, fontweight="bold", zorder=6)


# 1. movies → screenings  (movieid)
arrow(ax,
      all_anchors["movies"]["movieid"]["right"],
      all_anchors["screenings"]["movieid"]["left"])

# 2. cinemas → screenings (cinemaid) — vertical
arrow(ax,
      all_anchors["cinemas"]["cinemaid"]["bot"],
      all_anchors["screenings"]["cinemaid"]["top"])

# 3. cinemas → staff (cinemaid) — horizontal
arrow(ax,
      all_anchors["cinemas"]["cinemaid"]["right"],
      all_anchors["staff"]["cinemaid"]["left"])

# 4. screenings → tickets (screeningid) — vertical
arrow(ax,
      all_anchors["screenings"]["screeningid"]["bot"],
      all_anchors["tickets"]["screeningid"]["top"])

# 5. customers → tickets (customerid) — horizontal
arrow(ax,
      all_anchors["customers"]["customerid"]["left"],
      all_anchors["tickets"]["customerid"]["right"])


# ── legend ────────────────────────────────────────────────────────────────────
legend_items = [
    mpatches.Patch(facecolor=C_PK, label="Primärschlüssel (PK)"),
    mpatches.Patch(facecolor=C_FK, label="Fremdschlüssel (FK)"),
]
ax.legend(handles=legend_items, loc="lower right",
          fontsize=9.5, framealpha=0.9, edgecolor="#AAAAAA")

ax.set_title("ER-Diagramm: cinema.db",
             fontsize=14, fontweight="bold", color=C_TEXT, pad=12)

plt.tight_layout()
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cinema_er_diagram.png")
plt.savefig(out, dpi=150, bbox_inches="tight", facecolor=fig.get_facecolor())
print(f"ER diagram → {out}")
