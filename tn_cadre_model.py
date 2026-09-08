#!/usr/bin/env python3
"""
Tamil Nadu health workforce, cadres other than doctors.
Annual qualified output projected to 2050.

SCOPE AND LIMITS. This is an EDUCATION OUTPUT projection, not a workforce stock
projection. It answers how many people will qualify each year. It does not say
how many are practising, because no register comparable to the NMC register was
available for any cadre other than doctors. There is therefore no opening stock,
no attrition, and no need or demand benchmark for these cadres.

GNM and ANM nursing are absent entirely. They sit with the Tamil Nadu Nurses and
Midwives Council, not the university, so they appear nowhere in the source files.
Nursing figures here are degree level only and understate total nursing output.

Method follows the doctor model: seats, then fill rate, then completion rate,
then a course length lag. Growth is fitted in LEVELS, never as a compound rate.

Inputs:  new data/TN MGRMU/SEAT COUNT STAT 15052026.xlsx
         new data/TN MGRMU/passout COUNT STAT 15052026.xlsx
Outputs: tn_cadre_projection_2050.csv, cadre_out/*.png
"""
import openpyxl, re, csv, os, statistics
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

OUT = "cadre_out"
os.makedirs(OUT, exist_ok=True)
YEARS = np.arange(2021, 2051)
IDX = {y: i for i, y in enumerate(YEARS)}

plt.rcParams.update({
    "figure.dpi": 140, "savefig.dpi": 140, "font.size": 9,
    "font.family": "serif", "font.serif": ["Times New Roman", "DejaVu Serif"],
    "axes.edgecolor": "black", "axes.labelcolor": "black", "text.color": "black",
    "xtick.color": "black", "ytick.color": "black", "axes.titlecolor": "black",
    "axes.grid": True, "grid.color": "#d9d9d9", "grid.linewidth": 0.5,
    "axes.spines.top": False, "axes.spines.right": False,
    "figure.facecolor": "white", "axes.facecolor": "white", "legend.frameon": False,
})
thou = FuncFormatter(lambda v, p: f"{v:,.0f}")

SEAT_FILE = "new data/TN MGRMU/SEAT COUNT STAT 15052026.xlsx"
PASS_FILE = "new data/TN MGRMU/passout COUNT STAT 15052026.xlsx"

# Cadre, seat sheet, pass-out sheet, course length in years (to qualification).
# Course lengths include internship where the course carries one.
CADRES = [
    ("BSc Nursing",            "BSCN",                "BSCN 66 Pass out",        4, "Nursing"),
    ("MSc Nursing",            "M.Sc.Nursing",        "MSc.Nur  30",             2, "Nursing"),
    ("Post Basic BSc Nursing", "Post Basic Nursing ", "Post Basic Nursing - 68", 2, "Nursing"),
    ("BDS",                    "BDS",                 "BDS 54 PASSOUT",          5, "Dental"),
    ("MDS",                    "MDS",                 "MDS 24 PASS OUT",         3, "Dental"),
    ("B.Pharm",                "B.Pharm.",            "BPHARM",                  4, "Pharmacy"),
    ("M.Pharm",                "MPHARM",              "MPHARM",                  2, "Pharmacy"),
    ("Pharm.D",                "PHARMD",              "PHARMD",                  6, "Pharmacy"),
    ("BPT",                    "BPT",                 "BPT 74 PASS out",         5, "Rehabilitation"),
    ("MPT",                    "MPT",                 "MPT PASS out",            2, "Rehabilitation"),
    ("BOT",                    "BOT",                 "BOT Pass out",            5, "Rehabilitation"),
    ("MOT",                    "MOT",                 "MOT Pass out",            2, "Rehabilitation"),
    ("BAMS (Ayurveda)",        "BAMS",                "bams 64 passout",         6, "AYUSH"),
    ("BSMS (Siddha)",          "BSMS",                "bsms 60 passout",         6, "AYUSH"),
    ("BUMS (Unani)",           "BUMS",                "bums 62 passout",         6, "AYUSH"),
    ("BNYS (Naturopathy)",     "BNYS",                "bnys 82 passout",         6, "AYUSH"),
    ("BHMS (Homoeopathy)",     "BHMS",                "BHMS 58 PASSOUT",         6, "AYUSH"),
    ("MD Siddha",              "MD SIDDHA",           "MD SIDDHA 32 PASSOUT",    3, "AYUSH"),
    ("MD Homoeopathy",         "MD HOMOEOPATHY",      "md homoeo 44 passout",    3, "AYUSH"),
    ("AHS PG (allied)",        "AHS PG",              "AHS PG PASSOUT",          2, "Allied health"),
]

# Excluded, with reasons recorded rather than silently dropped.
EXCLUDED = [
    ("GNM nursing", "Not present in any source file. Sits with the Tamil Nadu Nurses "
                    "and Midwives Council, not the university."),
    ("ANM nursing", "Not present in any source file. Same reason as GNM."),
    ("AHS UG (allied health, degree)",
     "Present but internally inconsistent: the sheet reports 10,593 admissions against "
     "3,904 sanctioned seats for 2025, a fill rate of 271 per cent, and pass-outs swing "
     "from 80 to 10,716 to 508 across four years. Not usable without clarification from "
     "the university."),
    ("MD Naturopathy and Yoga, MD Unani",
     "Seat counts present but very small (65 and 11 seats) and pass-out sheets are "
     "incomplete. Folded into no projection."),
    ("Pharm.D Post Baccalaureate",
     "Only three pass-out years, two of them clearly partial."),
]

def _norm(s):
    return re.sub(r"[^a-z]", "", str(s).lower())

YCOLS = {"academicyear", "academicyr"}
ECOLS = {"examyear"}
SCOLS = {"sanctionedseats", "totalapprovedseats", "sanctionedseat"}
ACOLS = {"studentadmitted", "admittedcount", "seatsfilled", "studentsadmitted", "stud"}
PCOLS = {"studentspassedout", "studentpassedout"}

def read_sheets(path, ycols, valuesets):
    """Return {sheet: {year: [v1, v2, ...]}}. Handles the five header layouts
    used across the workbook by normalising column names."""
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for ws in wb.worksheets:
        hdr, agg = None, {}
        for row in ws.iter_rows(values_only=True):
            cells = [_norm(x) if x is not None else "" for x in row]
            if any(c in ycols for c in cells) and any(c in valuesets[0] for c in cells):
                hdr = cells
                continue
            if not hdr:
                continue
            d = dict(zip(hdr, row))
            y = next((d.get(c) for c in ycols if d.get(c) is not None), None)
            try:
                y = int(float(y))
            except (TypeError, ValueError):
                continue
            if not 2015 <= y <= 2030:
                continue
            rec = agg.setdefault(y, [0] * len(valuesets))
            for i, vs in enumerate(valuesets):
                v = next((d.get(c) for c in vs if d.get(c) is not None), None)
                try:
                    rec[i] += int(float(v))
                except (TypeError, ValueError):
                    pass
        if agg:
            out[ws.title] = agg
    wb.close()
    return out

SEATS = read_sheets(SEAT_FILE, YCOLS, [SCOLS, ACOLS])
PASSO = read_sheets(PASS_FILE, ECOLS, [PCOLS])

def usable_passouts(series):
    """Drop years below half the median. Across this workbook several cadres have
    a single year that is clearly a partial extract rather than a real collapse,
    for example BSc Nursing showing 710 in 2024 against 7,315 to 10,311 elsewhere."""
    if not series:
        return {}, []
    med = statistics.median(series.values())
    good = {y: v for y, v in series.items() if v >= 0.5 * med}
    return good, sorted(set(series) - set(good))

def build(name, s_sheet, p_sheet, dur):
    s = SEATS.get(s_sheet, {})
    p = {y: v[0] for y, v in PASSO.get(p_sheet, {}).items()}
    good, dropped = usable_passouts(p)
    yrs = sorted(s)
    sanc = {y: s[y][0] for y in yrs}
    adm = {y: s[y][1] for y in yrs}
    fill = (sum(adm.values()) / sum(sanc.values())) if sum(sanc.values()) else 0.0

    # Completion, where an admission cohort can actually be matched to its
    # own pass-out year inside the observed window. Otherwise flagged as assumed.
    ratios = []
    for py, pv in good.items():
        ay = py - dur
        if ay in adm and adm[ay] > 0:
            ratios.append(pv / adm[ay])
    if ratios:
        completion, comp_basis = float(np.median(ratios)), f"observed ({len(ratios)} cohort match{'es' if len(ratios) > 1 else ''})"
    else:
        completion, comp_basis = None, "assumed"
    return dict(name=name, dur=dur, sanc=sanc, adm=adm, fill=fill,
                passouts=p, usable=good, dropped=dropped,
                completion=completion, comp_basis=comp_basis)

BUILT = [build(n, ss, ps, d) for n, ss, ps, d, _ in CADRES]
GROUP = {n: g for n, _, _, _, g in CADRES}

# Seat data begins in 2021, so an output year is only fully covered once every
# course length has a source cohort. The longest course here is six years, which
# makes 2027 the first year in which all cadres report. Earlier output years are
# partial and totals across them would understate, so they are not summed.
FIRST_COMPLETE = 2021 + max(d for _, _, _, d, _ in CADRES)
REPORT_YEARS = (FIRST_COMPLETE, 2030, 2040, 2050)

# Cadres with no matchable cohort take the median of those that do have one.
_obs = [b["completion"] for b in BUILT if b["completion"] is not None]
DEFAULT_COMPLETION = float(np.median(_obs)) if _obs else 0.90
for b in BUILT:
    if b["completion"] is None:
        b["completion"] = DEFAULT_COMPLETION

def project_seats(b, path):
    """Two paths, because four or five annual observations cannot support a
    fitted ceiling. 'Frozen' holds the last observed sanctioned level. 'Trend'
    continues the fitted linear trend to 2035 and then holds, on the reasoning
    that capacity cannot expand indefinitely against a falling state population."""
    yrs = sorted(b["sanc"])
    last_y = max(yrs)
    last_v = b["sanc"][last_y]
    if path == "frozen" or len(yrs) < 3:
        return {y: last_v for y in YEARS}
    slope = float(np.polyfit(np.array(yrs, float), [b["sanc"][y] for y in yrs], 1)[0])
    out = {}
    for y in YEARS:
        if y <= last_y:
            out[y] = b["sanc"].get(y, last_v)
        else:
            out[y] = max(0.0, last_v + slope * (min(y, 2035) - last_y))
    return out

def qualified_output(b, path):
    seats = project_seats(b, path)
    out = {}
    for y in YEARS:
        sy = y - b["dur"]
        if sy in seats:
            out[y] = seats[sy] * b["fill"] * b["completion"]
        elif sy in b["sanc"]:
            out[y] = b["sanc"][sy] * b["fill"] * b["completion"]
        else:
            out[y] = float("nan")
    return out

# --------------------------------------------------------------------------
# Console summary
# --------------------------------------------------------------------------
lines = []
def emit(s=""):
    lines.append(s); print(s)

emit("=" * 96)
emit("TAMIL NADU HEALTH WORKFORCE, CADRES OTHER THAN DOCTORS")
emit("Annual qualified output projected to 2050")
emit("=" * 96)
emit()
emit("THIS IS AN EDUCATION OUTPUT PROJECTION, NOT A WORKFORCE STOCK PROJECTION.")
emit("No register comparable to the NMC register was available for these cadres, so")
emit("there is no opening stock, no attrition, and no need or demand benchmark.")
emit()
emit("PARAMETERS ESTIMATED FROM THE SOURCE FILES")
emit("-" * 96)
emit(f"{'Cadre':<24}{'Group':<16}{'Course':>7}{'2025 seats':>11}{'Fill':>7}"
     f"{'Completion':>12}  {'Completion basis':<28}")
for b in BUILT:
    ly = max(b["sanc"])
    emit(f"{b['name']:<24}{GROUP[b['name']]:<16}{b['dur']:>6}y{b['sanc'][ly]:>11,}"
         f"{b['fill']:>7.0%}{b['completion']:>12.0%}  {b['comp_basis']:<28}")
emit()
emit(f"Default completion where no cohort match was possible: {DEFAULT_COMPLETION:.0%} "
     f"(median of {len(_obs)} observed)")

emit()
emit("DATA QUALITY: PASS-OUT YEARS EXCLUDED AS PARTIAL EXTRACTS")
emit("-" * 96)
any_drop = False
for b in BUILT:
    if b["dropped"]:
        any_drop = True
        vals = ", ".join(f"{y}: {b['passouts'][y]:,}" for y in b["dropped"])
        kept = ", ".join(f"{v:,}" for v in sorted(b["usable"].values()))
        emit(f"  {b['name']:<24} excluded {vals:<24} kept: {kept}")
if not any_drop:
    emit("  none")

emit()
emit("NOT COVERED, AND WHY")
emit("-" * 96)
for name, why in EXCLUDED:
    emit(f"  {name}")
    for chunk in [why[i:i + 88] for i in range(0, len(why), 88)]:
        emit(f"      {chunk}")

emit()
emit("PROJECTED ANNUAL QUALIFIED OUTPUT (trend path)")
emit("-" * 96)
emit(f"First fully covered output year is {FIRST_COMPLETE}: seat data begins in 2021 and the")
emit("longest course here is six years. Earlier years are partial and are not totalled.")
emit()
emit(f"{'Cadre':<24}" + "".join(f"{y:>10}" for y in REPORT_YEARS))
tot = {y: 0.0 for y in REPORT_YEARS}
for b in BUILT:
    q = qualified_output(b, "trend")
    emit(f"{b['name']:<24}" + "".join(
        f"{q[y]:>10,.0f}" if q[y] == q[y] else f"{'n/a':>10}" for y in REPORT_YEARS))
    for y in tot:
        if q[y] == q[y]:
            tot[y] += q[y]
emit(f"{'TOTAL (these cadres)':<24}" + "".join(f"{tot[y]:>10,.0f}" for y in REPORT_YEARS))

emit()
emit("BY GROUP, TREND PATH")
emit("-" * 96)
groups = sorted(set(GROUP.values()))
emit(f"{'Group':<18}" + "".join(f"{y:>10}" for y in REPORT_YEARS))
GROUPTOT = {g: {} for g in groups}
for g in groups:
    for y in REPORT_YEARS:
        v = sum(qualified_output(b, "trend")[y] for b in BUILT
                if GROUP[b["name"]] == g and qualified_output(b, "trend")[y] == qualified_output(b, "trend")[y])
        GROUPTOT[g][y] = v
    emit(f"{g:<18}" + "".join(f"{GROUPTOT[g][y]:>10,.0f}" for y in REPORT_YEARS))

# --------------------------------------------------------------------------
# CSV
# --------------------------------------------------------------------------
with open("tn_cadre_projection_2050.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["year", "basis", "cadre", "group", "course_years",
                "sanctioned_seats_frozen", "sanctioned_seats_trend",
                "fill_rate", "completion_rate", "completion_basis",
                "qualified_output_frozen", "qualified_output_trend",
                "full_coverage_year"])
    for b in BUILT:
        sf, st = project_seats(b, "frozen"), project_seats(b, "trend")
        qf, qt = qualified_output(b, "frozen"), qualified_output(b, "trend")
        for y in YEARS:
            w.writerow([y, "Observed" if y <= 2025 else "Projected", b["name"],
                        GROUP[b["name"]], b["dur"],
                        f"{sf[y]:.0f}", f"{st[y]:.0f}",
                        f"{b['fill']:.4f}", f"{b['completion']:.4f}", b["comp_basis"],
                        "" if qf[y] != qf[y] else f"{qf[y]:.0f}",
                        "" if qt[y] != qt[y] else f"{qt[y]:.0f}",
                        "yes" if y >= FIRST_COMPLETE else "no, partial"])

with open("tn_cadre_summary.txt", "w") as f:
    f.write("\n".join(lines) + "\n")

# --------------------------------------------------------------------------
# Charts, monochrome
# --------------------------------------------------------------------------
GREYS = ["#000000", "#000000", "#595959", "#595959", "#8c8c8c", "#8c8c8c"]
DASH = ["-", "--", "-.", ":", "-", "--"]

def finish(ax, ylab):
    ax.set_ylabel(ylab, fontsize=9)
    ax.set_xlabel("Year", fontsize=9)
    ax.yaxis.set_major_formatter(thou)
    ax.set_axisbelow(True)
    ax.grid(axis="x", visible=False)

def save(fig, name):
    fig.tight_layout()
    fig.savefig(f"{OUT}/{name}.png", bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("  wrote", f"{OUT}/{name}.png")

# Chart 1: qualified output by group
fig, ax = plt.subplots(figsize=(8.6, 4.8))
PY_ = np.array([y for y in YEARS if y >= FIRST_COMPLETE])
for i, g in enumerate(groups):
    series = []
    for y in PY_:
        v = sum(qualified_output(b, "trend")[y] for b in BUILT
                if GROUP[b["name"]] == g and qualified_output(b, "trend")[y] == qualified_output(b, "trend")[y])
        series.append(v)
    ax.plot(PY_, series, color=GREYS[i % len(GREYS)], linestyle=DASH[i % len(DASH)],
            lw=1.8, label=g)
    ax.annotate(f"{series[-1]:,.0f}", (PY_[-1], series[-1]), xytext=(4, 0),
                textcoords="offset points", fontsize=8, va="center")
ax.set_xlim(FIRST_COMPLETE, 2055)
ax.legend(loc="upper left", fontsize=8.5)
finish(ax, "Qualified output per year")
save(fig, "cadre_fig1_output_by_group")

# Chart 2: nursing detail
fig, ax = plt.subplots(figsize=(8.6, 4.8))
for i, b in enumerate([x for x in BUILT if GROUP[x["name"]] == "Nursing"]):
    q = qualified_output(b, "trend")
    NY = [y for y in YEARS if y >= FIRST_COMPLETE]
    ax.plot(NY, [q[y] for y in NY], color=GREYS[i % len(GREYS)],
            linestyle=DASH[i % len(DASH)], lw=1.8, label=b["name"])
    ax.annotate(f"{q[NY[-1]]:,.0f}", (NY[-1], q[NY[-1]]), xytext=(4, 0),
                textcoords="offset points", fontsize=8, va="center")
ax.set_xlim(FIRST_COMPLETE, 2055)
ax.legend(loc="upper left", fontsize=8.5)
finish(ax, "Qualified output per year")
save(fig, "cadre_fig2_nursing")

# Chart 3: fill rates, the key finding outside medicine
fig, ax = plt.subplots(figsize=(8.6, 5.4))
order = sorted(BUILT, key=lambda b: b["fill"])
ax.barh([b["name"] for b in order], [b["fill"] * 100 for b in order],
        color="#808080", edgecolor="black", linewidth=0.6, height=0.68)
for i, b in enumerate(order):
    ax.text(b["fill"] * 100 + 1.2, i, f"{b['fill']:.0%}", va="center", fontsize=8)
ax.set_xlim(0, 112)
ax.set_xlabel("Seats actually filled, per cent", fontsize=9)
ax.grid(axis="y", visible=False)
ax.set_axisbelow(True)
save(fig, "cadre_fig3_fill_rates")

print("\nWrote tn_cadre_projection_2050.csv and tn_cadre_summary.txt")
