"""
Draws mock-ups of the EURO 2028 Power BI report pages from the SAMPLE returns
(synthetic data), using the same calculations as the Power BI model
(reference_calc.py). Output: PNG per page in this folder.
"""
import datetime as dt
import os
import sys
from collections import defaultdict

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.patches import FancyBboxPatch  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import reference_calc as RC  # noqa: E402

# palette (validated): categorical slots 1-3, blue sequential steps, chrome
CAT = ["#2a78d6", "#eb6834", "#1baf7a"]
SEQ = {"Low": "#b7d3f6", "Medium": "#5598e7", "High": "#256abf", "Super High": "#0d366b"}
RISKS = ["Low", "Medium", "High", "Super High"]
CANVAS, CARD, NAVY = "#f3f2f1", "#ffffff", "#1f3864"
INK, INK2, INK3, GRID = "#0b0b0b", "#52514e", "#8a8984", "#e6e5e1"
GOOD, WARN, CRIT = "#0ca30c", "#fab219", "#d03b3b"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": GRID,
                     "axes.labelcolor": INK2, "xtick.color": INK2, "ytick.color": INK2})
W, H = 16, 9


def money(v, short=True):
    if short and abs(v) >= 1e6:
        return f"£{v / 1e6:.1f}m"
    if short and abs(v) >= 1e3:
        return f"£{v / 1e3:.0f}k"
    return f"£{v:,.0f}"


def num(v):
    return f"{v:,.0f}"


def frame(title, active, slicers):
    fig = plt.figure(figsize=(W, H), facecolor=CANVAS)
    top = fig.add_axes([0, 0.925, 1, 0.075])
    top.set_axis_off()
    top.add_patch(plt.Rectangle((0, 0), 1, 1, color=NAVY, transform=top.transAxes))
    top.text(0.012, 0.5, "EURO 2028  |  Policing resource requirement", color="white", fontsize=15,
             fontweight="bold", va="center", transform=top.transAxes)
    top.text(0.988, 0.5, "OFFICIAL  -  SAMPLE DATA (synthetic)", color="#f8cbad", fontsize=10, fontweight="bold",
             va="center", ha="right", transform=top.transAxes)
    nav = fig.add_axes([0, 0.885, 1, 0.04])
    nav.set_axis_off()
    pages = ["Overview", "Match Day", "Fixture Calendar", "Wider Tournament", "Data Quality"]
    x = 0.012
    for p in pages:
        on = p == active
        nav.text(x, 0.5, p, fontsize=10.5, fontweight="bold" if on else "normal", color=NAVY if on else INK2,
                 va="center", transform=nav.transAxes)
        if on:
            nav.plot([x, x + 0.0068 * len(p)], [0.08, 0.08], color=NAVY, lw=3, transform=nav.transAxes)
        x += 0.0068 * len(p) + 0.03
    # slicers row
    x = 0.012
    for label, value in slicers:
        ax = fig.add_axes([x, 0.825, 0.135, 0.05])
        ax.set_axis_off()
        ax.add_patch(FancyBboxPatch((0.01, 0.05), 0.98, 0.9, boxstyle="round,pad=0,rounding_size=0.08",
                                    fc=CARD, ec="#c8c6c4", transform=ax.transAxes))
        ax.text(0.06, 0.68, label, fontsize=8, color=INK3, transform=ax.transAxes, va="center")
        ax.text(0.06, 0.3, value, fontsize=10, color=INK, transform=ax.transAxes, va="center")
        ax.text(0.92, 0.3, "v", fontsize=9, color=INK3, transform=ax.transAxes, va="center")
        x += 0.145
    fig.text(0.012, 0.795, title, fontsize=13, fontweight="bold", color=INK)
    return fig


def card(fig, rect, label, value, sub=None, status=None):
    ax = fig.add_axes(rect)
    ax.set_axis_off()
    ax.add_patch(FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0,rounding_size=0.04", fc=CARD, ec="#e1dfdd",
                                transform=ax.transAxes))
    ax.text(0.06, 0.74, label, fontsize=9.5, color=INK2, transform=ax.transAxes)
    ax.text(0.06, 0.36, value, fontsize=22, fontweight="bold", color=INK, transform=ax.transAxes)
    if sub:
        ax.text(0.06, 0.12, sub, fontsize=8.5, color=INK3, transform=ax.transAxes)
    if status:
        col, sym = status
        ax.text(0.9, 0.74, sym, fontsize=10, color="white", fontweight="bold", ha="center", va="center",
                transform=ax.transAxes, bbox=dict(boxstyle="circle,pad=0.35", fc=col, ec="none"))


def panel(fig, rect, title, left=0.06):
    bg = fig.add_axes(rect)
    bg.set_axis_off()
    bg.add_patch(FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0,rounding_size=0.02", fc=CARD, ec="#e1dfdd",
                                transform=bg.transAxes))
    x, y, w, h = rect
    fig.text(x + 0.01, y + h - 0.035, title, fontsize=11, fontweight="bold", color=INK)
    ax = fig.add_axes([x + left, y + 0.06, w - left - 0.02, h - 0.13])
    ax.set_facecolor(CARD)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color(GRID)
    ax.spines["bottom"].set_color(GRID)
    return ax


def legend(ax, items, loc="upper right"):
    hs = [plt.Rectangle((0, 0), 1, 1, color=c) for _, c in items]
    ax.legend(hs, [n for n, _ in items], loc=loc, frameon=False, fontsize=8.5, ncol=len(items),
              handlelength=1, handleheight=1, columnspacing=1.2)


def table(fig, rect, title, head, rows, shade_col=None, widths=None):
    bg = fig.add_axes(rect)
    bg.set_axis_off()
    bg.add_patch(FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0,rounding_size=0.02", fc=CARD, ec="#e1dfdd",
                                transform=bg.transAxes))
    x, y, w, h = rect
    fig.text(x + 0.01, y + h - 0.035, title, fontsize=11, fontweight="bold", color=INK)
    ax = fig.add_axes([x + 0.01, y + 0.02, w - 0.02, h - 0.12])
    ax.set_axis_off()
    n = len(rows) + 1
    widths = widths or [1 / len(head)] * len(head)
    xs = [sum(widths[:i]) for i in range(len(head))]
    rh = 1 / n
    vals = [v for r in rows for v in r[1:] if isinstance(v, (int, float))]
    vmax = max(vals) if vals else 1
    for j, hname in enumerate(head):
        ax.text(xs[j] + 0.01, 1 - rh / 2, hname, fontsize=8.5, fontweight="bold", color=INK2, va="center",
                ha="left" if j == 0 else "right" if False else "left")
    ax.plot([0, 1], [1 - rh, 1 - rh], color=GRID, lw=1)
    for i, r in enumerate(rows):
        yy = 1 - rh * (i + 1.5)
        for j, v in enumerate(r):
            if shade_col is not None and j >= shade_col and isinstance(v, (int, float)):
                t = v / vmax
                ax.add_patch(plt.Rectangle((xs[j] + 0.004, yy - rh * 0.42), widths[j] - 0.008, rh * 0.84,
                                           color=SEQ["Low"], alpha=0.25 + 0.6 * t, lw=0))
            txt = v if isinstance(v, str) else (money(v) if r and shade_col is not None else num(v))
            bold = r[0] == "Total"
            ax.text(xs[j] + 0.01, yy, txt, fontsize=9, color=INK, va="center", fontweight="bold" if bold else None)
        ax.plot([0, 1], [yy - rh / 2, yy - rh / 2], color="#f0efec", lw=0.8)
    return ax


# --------------------------------------------------------------------- data
md, wt, cal = RC.run(os.path.join(ROOT, "samples", "returns"),
                     os.path.join(ROOT, "samples", "EURO_2028_Reference_Data_SAMPLE.xlsx"))
FORCES = sorted({f["Force"] for f in md})
_, _, _, _, FIXTURES = RC.load_reference(os.path.join(ROOT, "samples", "EURO_2028_Reference_Data_SAMPLE.xlsx"))


def tot(rows, field, **flt):
    return sum(r[field] for r in rows if all(r[k] == v for k, v in flt.items()))


cost = lambda rows, **f: tot(rows, "Overtime", **f) + tot(rows, "Package", **f)  # noqa: E731
wt_cost = lambda **f: tot(wt, "TOvertime", **f) + tot(wt, "TPackage", **f)  # noqa: E731
OUT = []


def save(fig, name):
    p = os.path.join(HERE, name)
    fig.savefig(p, dpi=150, facecolor=CANVAS)
    plt.close(fig)
    OUT.append(p)


# --------------------------------------------------------------- 1 Overview
fig = frame("Overview - whole tournament", "Overview",
            [("Force", "All"), ("Supplied by", "All"), ("Cost basis", "Original basis")])
cal_total, wt_total = cost(cal), wt_cost()
cards = [("Forces returned", "3", "of 3 expected", (GOOD, "✓")),
         ("Tournament cost", money(cal_total + wt_total), "fixture days + wider tournament", None),
         ("Fixture-day cost", money(cal_total), f"{len({f[0] for f in FIXTURES})} fixtures on the calendar", None),
         ("Wider tournament cost", money(wt_total), "32 days (7 / 11 / 11 / 3)", None),
         ("Issues to chase", "0", "data quality checks", (GOOD, "✓"))]
for i, c in enumerate(cards):
    card(fig, [0.012 + i * 0.197, 0.62, 0.185, 0.155], *c)
ax = panel(fig, [0.012, 0.03, 0.54, 0.57], "Tournament cost by force", left=0.1)
y = range(len(FORCES))
a = [cost(cal, Force=f) for f in FORCES]
b = [wt_cost(Force=f) for f in FORCES]
ax.barh(y, a, color=CAT[0], height=0.5, edgecolor=CARD, linewidth=2)
ax.barh(y, b, left=a, color=CAT[1], height=0.5, edgecolor=CARD, linewidth=2)
for i in y:
    ax.text(a[i] + b[i] + max(a) * 0.03, i, money(a[i] + b[i]), va="center", fontsize=10, color=INK)
ax.set_yticks(list(y), FORCES)
ax.invert_yaxis()
ax.xaxis.set_major_formatter(lambda v, _: money(v))
ax.grid(axis="x", color=GRID, lw=0.8)
ax.set_axisbelow(True)
ax.set_xlim(0, max(i + j for i, j in zip(a, b)) * 1.2)
legend(ax, [("Fixture days", CAT[0]), ("Wider tournament", CAT[1])], loc="lower right")
ax = panel(fig, [0.565, 0.03, 0.423, 0.57], "Officer deployments for ONE fixture, by risk level (all forces)")
vals = [tot(md, "Officers", Risk=r) for r in RISKS]
ax.bar(RISKS, vals, color=[SEQ[r] for r in RISKS], width=0.55, edgecolor=CARD, linewidth=2)
for i, v in enumerate(vals):
    ax.text(i, v + max(vals) * 0.02, num(v), ha="center", fontsize=10, color=INK)
ax.grid(axis="y", color=GRID, lw=0.8)
ax.set_axisbelow(True)
ax.set_ylim(0, max(vals) * 1.15)
ax.set_ylabel("Officer deployments (officers x days, MD-2 to MD+1)")
save(fig, "1_Overview.png")

# --------------------------------------------------------------- 2 Match Day
fig = frame("Match Day - what ONE fixture needs (Risk = High)", "Match Day",
            [("Risk level", "High"), ("Force", "All"), ("Supplied by", "All"), ("Rank", "All"),
             ("Role / Skill", "All")])
hi = [f for f in md if f["Risk"] == "High"]
cards = [("Officer deployments", num(tot(hi, "Officers")), "officers x days, MD-2 to MD+1", None),
         ("Peak day officers", num(max(tot(hi, "Officers", Day=d) for d in ["MD-2", "MD-1", "MD", "MD+1"])),
          "busiest day (Match Day)", None),
         ("PSUs / Serials", num(tot(hi, "PSUs")), "force supplied + mutual aid", None),
         ("Cost (original basis)", money(cost(hi)), "overtime + mutual aid PSUs", None),
         ("Mutual aid share", f"{tot(hi, 'Officers', Supply='Mutual aid') / tot(hi, 'Officers'):.0%}",
          "of officer deployments", None)]
for i, c in enumerate(cards):
    card(fig, [0.012 + i * 0.197, 0.62, 0.185, 0.155], *c)
sec = defaultdict(float)
for f in hi:
    sec[f["Section"]] += f["Officers"]
top = sorted(sec.items(), key=lambda kv: -kv[1])[:10]
ax = panel(fig, [0.012, 0.03, 0.38, 0.57], "Officer deployments by section (top 10)", left=0.15)
ax.barh(range(len(top)), [v for _, v in top], color=CAT[0], height=0.6, edgecolor=CARD, linewidth=2)
for i, (_, v) in enumerate(top):
    ax.text(v + top[0][1] * 0.02, i, num(v), va="center", fontsize=9, color=INK)
ax.set_yticks(range(len(top)), [k for k, _ in top], fontsize=8.5)
ax.invert_yaxis()
ax.set_xlim(0, top[0][1] * 1.15)
ax.grid(axis="x", color=GRID, lw=0.8)
ax.set_axisbelow(True)
ax = panel(fig, [0.405, 0.03, 0.27, 0.57], "Officers by day")
days = ["MD-2", "MD-1", "MD", "MD+1"]
dv = [tot(hi, "Officers", Day=d) for d in days]
ax.bar(["Match Day -2", "Match Day -1", "Match Day", "Match Day +1"], dv, color=SEQ["High"], width=0.55,
       edgecolor=CARD, linewidth=2)
for i, v in enumerate(dv):
    ax.text(i, v + max(dv) * 0.02, num(v), ha="center", fontsize=9.5, color=INK)
ax.tick_params(axis="x", labelsize=8)
ax.set_ylim(0, max(dv) * 1.15)
ax.grid(axis="y", color=GRID, lw=0.8)
ax.set_axisbelow(True)
rows = [[f] + [cost(md, Force=f, Risk=r) for r in RISKS] for f in FORCES]
rows.append(["Total"] + [cost(md, Risk=r) for r in RISKS])
table(fig, [0.688, 0.03, 0.3, 0.57], "Cost of ONE fixture: force x risk level", ["Force"] + RISKS, rows,
      shade_col=1, widths=[0.28, 0.18, 0.18, 0.18, 0.18])
save(fig, "2_Match_Day.png")

# ------------------------------------------------------- 3 Fixture calendar
fig = frame("Fixture calendar - requirement by date", "Fixture Calendar",
            [("Force", "All"), ("Date range", "06/06 - 22/06/2028"), ("Supplied by", "All")])
d0, d1 = dt.date(2028, 6, 6), dt.date(2028, 6, 22)
dates = [d0 + dt.timedelta(i) for i in range((d1 - d0).days + 1)]
by = {f: [tot(cal, "Officers", Force=f, Date=d) for d in dates] for f in FORCES}
byc = {f: [cost(cal, Force=f, Date=d) for d in dates] for f in FORCES}
for k, (data, title, fmt) in enumerate(((by, "Officer deployments by date", num),
                                       (byc, "Cost by date (original basis)", money))):
    ax = panel(fig, [0.012, 0.415 - k * 0.385, 0.976, 0.37], title)
    ax.set_position([0.07, 0.415 - k * 0.385 + 0.085, 0.9, 0.37 - 0.15])
    base = [0] * len(dates)
    for i, f in enumerate(FORCES):
        ax.bar(range(len(dates)), data[f], bottom=base, color=CAT[i], width=0.7, edgecolor=CARD, linewidth=1.5)
        base = [a + b for a, b in zip(base, data[f])]
    for i, v in enumerate(base):
        if v:
            ax.text(i, v + max(base) * 0.02, fmt(v), ha="center", fontsize=7.5, color=INK2)
    fx_on = defaultdict(list)
    for fx, d, risk, force in FIXTURES:
        if f"{fx} {risk}" not in fx_on[d]:
            fx_on[d].append(f"{fx} {risk}")
    ax.set_xticks(range(len(dates)), [d.strftime("%a %d") + ("\n" + "\n".join(fx_on[d]) if d in fx_on else "")
                                      for d in dates], fontsize=8)
    ax.set_ylim(0, max(base) * 1.25)
    ax.yaxis.set_major_formatter(lambda v, _, f=fmt: f(v))
    ax.grid(axis="y", color=GRID, lw=0.8)
    ax.set_axisbelow(True)
    if k == 0:
        legend(ax, [(f, CAT[i]) for i, f in enumerate(FORCES)], loc="upper right")
save(fig, "3_Fixture_Calendar.png")

# ------------------------------------------------------- 4 Wider tournament
fig = frame("Wider tournament - typical day and 32-day total", "Wider Tournament",
            [("Force", "All"), ("Supplied by", "All"), ("Cost basis", "Original basis")])
cards = [("Typical-day officers (High)", num(tot(wt, "Officers", Risk="High")), "all forces", None),
         ("Officer days (tournament)", num(tot(wt, "TOfficers")), "typical day x days at each risk", None),
         ("Wider tournament cost", money(wt_total), "overtime + mutual aid PSUs", None),
         ("incl normal duty", money(wt_total + tot(wt, "TNormal")), "adds normal tour of duty hours", None)]
for i, c in enumerate(cards):
    card(fig, [0.012 + i * 0.247, 0.62, 0.235, 0.155], *c)
ax = panel(fig, [0.012, 0.03, 0.55, 0.57], "Wider tournament cost by force and risk level", left=0.1)
left = [0] * len(FORCES)
for r in RISKS:
    v = [wt_cost(Force=f, Risk=r) for f in FORCES]
    ax.barh(range(len(FORCES)), v, left=left, color=SEQ[r], height=0.5, edgecolor=CARD, linewidth=2)
    left = [a + b for a, b in zip(left, v)]
for i, v in enumerate(left):
    ax.text(v + max(left) * 0.02, i, money(v), va="center", fontsize=10, color=INK)
ax.set_yticks(range(len(FORCES)), FORCES)
ax.invert_yaxis()
ax.set_xlim(0, max(left) * 1.2)
ax.xaxis.set_major_formatter(lambda v, _: money(v))
ax.grid(axis="x", color=GRID, lw=0.8)
ax.set_axisbelow(True)
legend(ax, [(r, SEQ[r]) for r in RISKS], loc="lower right")
rows = [[f] + [tot(wt, "Officers", Force=f, Risk=r) for r in RISKS] for f in FORCES]
rows.append(["Total"] + [tot(wt, "Officers", Risk=r) for r in RISKS])
tb = table(fig, [0.575, 0.03, 0.413, 0.57], "Typical-day officers: force x risk level", ["Force"] + RISKS,
           rows, widths=[0.3, 0.175, 0.175, 0.175, 0.175])
save(fig, "4_Wider_Tournament.png")

# ------------------------------------------------------------ 5 Data quality
fig = frame("Data quality - what to chase with forces", "Data Quality",
            [("Force", "All"), ("Page", "All")])
cards = [("Forms received", "3", "3 forces", (GOOD, "✓")),
         ("Issues to chase", "4", "illustrative examples below", (WARN, "!")),
         ("Forces with issues", "2", "Sample Force B, Sample Force C", (WARN, "!")),
         ("Fixture forces missing", "0", "every fixture force has returned", (GOOD, "✓"))]
for i, c in enumerate(cards):
    card(fig, [0.012 + i * 0.247, 0.62, 0.235, 0.155], *c)
rows = [["Sample Force B", "MD Collection", "Row 141: Additional Armed Policing / Armed Response Overlay / Mutual aid",
         "Numbers entered but no Rank chosen"],
        ["Sample Force B", "WT Collection", "Row 52: Additional Public Order / Public Safety / Force supplied",
         "Numbers entered but no Resource given"],
        ["Sample Force C", "MD Collection", "Row 23: Public Order / Stadium Vicinity ... - Level 2 / Mutual aid",
         "Numbers entered but no Role / Skill (or unit type) chosen"],
        ["Reference data", "Fixtures", "-",
         "Force on the fixture list has no Match Day figures (check spelling)"]]
table(fig, [0.012, 0.03, 0.976, 0.57], "Issues (ILLUSTRATIVE examples - the sample returns themselves have none)",
      ["Force", "Page", "Line", "Issue"], rows, widths=[0.14, 0.11, 0.42, 0.33])
save(fig, "5_Data_Quality.png")

print("\n".join(OUT))
