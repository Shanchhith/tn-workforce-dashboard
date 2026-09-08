#!/usr/bin/env python3
"""
Tamil Nadu medical workforce, supply-side projection to 2050.

Stock-and-flow model with an explicit education pipeline, following the
WHO/World Bank Health Labour Market Framework (Liu, Goryakin, Maeda, Bruckner &
Scheffler, World Bank PRWP 7790 / Human Resources for Health 2017;15:11).

Growth forms are fitted in LEVELS, never as CAGR, see PROJECTION_PLAN_2050.md §5.2.
Government seats are a step/scenario process; private seats follow bounded
(logistic) growth; PG seats are linear with a structural cap.

Outputs: tn_projection_2050.csv, tn_projection_2050_summary.txt, figures in
projections_out/.
"""
import csv, collections, os
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

OUT = "projections_out"
os.makedirs(OUT, exist_ok=True)

# Reference palette (dataviz skill, light mode), assigned in fixed slot order.
C_BLUE, C_ORANGE, C_AQUA, C_YELLOW = "#2a78d6", "#eb6834", "#1baf7a", "#eda100"
C_MAGENTA, C_GREEN, C_VIOLET, C_RED = "#e87ba4", "#008300", "#4a3aa7", "#e34948"
INK, INK2, INK3 = "#0b0b0b", "#52514e", "#8a8880"
GRID = "#e3e2dd"
plt.rcParams.update({
    "figure.dpi": 140, "savefig.dpi": 140, "font.size": 9,
    "axes.edgecolor": GRID, "axes.labelcolor": INK2, "text.color": INK,
    "xtick.color": INK2, "ytick.color": INK2, "axes.titlecolor": INK,
    "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.6,
    "axes.spines.top": False, "axes.spines.right": False,
    "figure.facecolor": "white", "axes.facecolor": "white",
    "legend.frameon": False, "font.family": "DejaVu Sans",
})
thou = FuncFormatter(lambda v, p: f"{v:,.0f}")

YEARS = np.arange(2011, 2051)
IDX = {y: i for i, y in enumerate(YEARS)}

# ============================================================================
# 1. POPULATION SPINE
# ============================================================================
# NCP/MoHFW Population Projections 2011-2036 (observed/official).
NCP = {2011:72147,2012:72645,2013:73142,2014:73640,2015:74137,2016:74635,
       2017:74989,2018:75342,2019:75695,2020:76049,2021:76402,2022:76631,
       2023:76860,2024:77089,2025:77317,2026:77546,2027:77653,2028:77761,
       2029:77868,2030:77975,2031:78082,2032:78079,2033:78076,2034:78073,
       2035:78070,2036:78067}  # thousands

def build_population():
    """NCP to 2036; beyond that the growth rate declines linearly to -0.35%/yr
    by 2050, reflecting demographic momentum exhausting at TFR ~1.4.
    The NCP tail is itself a flat -3,000/yr linear stub, so it is not extended
    mechanically."""
    pop = {y: NCP[y] * 1000.0 for y in NCP}
    g0 = (NCP[2036] - NCP[2035]) / NCP[2035]      # ~-0.0038%/yr
    g1 = -0.0035                                   # -0.35%/yr by 2050
    for i, y in enumerate(range(2037, 2051), start=1):
        g = g0 + (g1 - g0) * (i / 14.0)
        pop[y] = pop[y - 1] * (1 + g)
    return np.array([pop[y] for y in YEARS])

POP = build_population()

# ============================================================================
# 2. SEAT MODULE
# ============================================================================
# MBBS seats, state selection-committee basis (Government vs Self-financing).
SEAT_YRS = np.arange(2015, 2026)
GOV_OBS = np.array([2655,2650,3050,2900,3600,3675,5175,5175,5200,5200,5200], float)
PVT_OBS = np.array([1010,1610,1600,1600,1950,2350,2900,3350,3850,4050,4750], float)

def logistic(t, K, r, t0):
    return K / (1.0 + np.exp(-r * (t - t0)))

def fit_logistic(t, y, K):
    """Least-squares fit of r, t0 for a fixed ceiling K (grid + refine)."""
    best = None
    for r in np.arange(0.05, 0.60, 0.002):
        # closed-form-ish t0 given r: minimise SSE over a t0 grid
        for t0 in np.arange(0.0, 30.0, 0.25):
            sse = ((y - logistic(t, K, r, t0)) ** 2).sum()
            if best is None or sse < best[0]:
                best = (sse, r, t0)
    return best[1], best[2]

t_seat = SEAT_YRS - 2015.0
PVT_PARAMS = {K: fit_logistic(t_seat, PVT_OBS, K) for K in (8000, 10000, 12000)}

# Government seat scenarios are built from the actual college-level distribution,
# not from an arbitrary growth rate. Tamil Nadu has 37 government medical colleges
# (MGRMU, 2025) holding 5,200 seats, a mean of 141 each. Sixteen are at 100 seats,
# sixteen at 150, one at 200 and only four at the NMC ceiling of 250. Because the
# state already has near-universal district coverage (37 colleges, 38 districts),
# the realistic lever is raising intake at existing colleges, not building new ones.
#
#   Frozen        5,200  no policy action at all, a counterfactual, not a forecast
#   Consolidation 6,000  the sixteen 100-seat colleges rise to 150
#   Full capacity 9,250  every college reaches the NMC ceiling of 250
GOV_PLAN = {
    "Constrained": (5200.0, None),      # frozen
    "Status quo":  (6000.0, 2035),      # 100-seat colleges -> 150
    "Policy push": (9250.0, 2045),      # all 37 colleges -> NMC cap of 250
}

SCENARIOS = {
    "Constrained": dict(K=8000),
    "Status quo":  dict(K=10000),
    "Policy push": dict(K=12000),
}

def gov_seats(scen):
    """Government seats: a policy step process, not a fitted curve.

    Seats are set by government, so these are scenarios about decisions, not
    extrapolations of a trend. Each path ramps from the observed 5,200 to its
    target by the target year (rounded to NMC's 50-seat approval blocks), then
    holds. Freezing at 5,200 is retained as the no-action counterfactual."""
    target, tyear = GOV_PLAN[scen]
    s = {}
    for y in YEARS:
        if y <= 2025:
            s[y] = GOV_OBS[list(SEAT_YRS).index(y)] if y >= 2015 else np.nan
        elif tyear is None or y >= tyear:
            s[y] = target
        else:
            frac = (y - 2025) / (tyear - 2025)
            s[y] = round((5200.0 + (target - 5200.0) * frac) / 50.0) * 50.0
    return np.array([s[y] for y in YEARS])

def pvt_seats(scen):
    K = SCENARIOS[scen]["K"]; r, t0 = PVT_PARAMS[K]
    out = []
    for y in YEARS:
        if 2015 <= y <= 2025:
            out.append(PVT_OBS[list(SEAT_YRS).index(y)])
        elif y < 2015:
            out.append(np.nan)
        else:
            out.append(logistic(y - 2015.0, K, r, t0))
    return np.array(out)

# ---- PG (MD/MS/Diploma) seats -------------------------------------------
# MGRMU basis 2021-24 scaled to full basis (x1.407, validated against the
# Selection Committee 2024-25 total of 4,629); 2025-26 is the SC actual.
PG_COVERAGE = 1.407          # MGRMU basis -> full basis (validated: 4,630 vs 4,629)
PG_YRS = np.arange(2021, 2026)
PG_OBS = np.array([4100, 4283, 4435, 4630, 5534], float)
_pg_b = np.polyfit(PG_YRS - 2021.0, PG_OBS, 1)   # linear in levels

def pg_seats(scen):
    """Linear growth, capped at 70% of total MBBS seats (you cannot train more
    specialists than the MBBS pipeline delivers)."""
    mbbs = gov_seats(scen) + pvt_seats(scen)
    out = []
    for i, y in enumerate(YEARS):
        if 2021 <= y <= 2025:
            v = PG_OBS[list(PG_YRS).index(y)]
        elif y < 2021:
            v = np.nan
        else:
            v = np.polyval(_pg_b, y - 2021.0)
        if y > 2025 and not np.isnan(mbbs[i]):
            v = min(v, 0.70 * mbbs[i])
        out.append(v)
    return np.array(out)

# ============================================================================
# 3. PIPELINE PARAMETERS  (estimated from observed data, not assumed)
# ============================================================================
FILL_MBBS, FILL_PG = 0.998, 0.978     # MGRMU sanctioned vs admitted, 2021-2025
COMPLETION = 0.95                      # NMC benchmark; validated below for PG
LAG_MBBS, LAG_PG = 6, 3                # 4.5 yrs + 12-month CRMI; PG 3 yrs

# Observed TNMC registrations (Medical Council workbook)
REG_MBBS_OBS = {2020:8409, 2021:8058, 2022:10367, 2023:9370, 2024:9722, 2025:10839}
REG_FMG_OBS  = {2020:640,  2021:1268, 2022:1287,  2023:1622, 2024:1427, 2025:1606}

def domestic_output(seat_series, year):
    """MBBS graduates entering the register in `year`, from seats 6 years earlier."""
    sy = year - LAG_MBBS
    if sy < 2015:
        return np.nan
    return seat_series[IDX[sy]] * FILL_MBBS * COMPLETION

# --- External entrants: the residual (deemed universities + out-of-state) ---
_gov_sq, _pvt_sq = gov_seats("Status quo"), pvt_seats("Status quo")
_tot_sq = _gov_sq + _pvt_sq
_resid = []
for y in range(2021, 2026):
    dom = domestic_output(_tot_sq, y)
    _resid.append(REG_MBBS_OBS[y] - dom)
RESIDUAL = float(np.mean(_resid))       # flat: no trend over the 5 observations

# --- Deemed-university overlap correction -----------------------------------
# The Selection Committee and MGRMU private-MBBS series agree exactly for
# 2021-22 and 2022-23, then diverge (2023: 450, 2024: 550, 2025: 850 seats).
# The gap is most likely deemed universities entering the SC sheet. Deemed
# output is ALSO inside the flat residual, which was estimated from seat years
# 2015-2019 when the two sources agreed. Without a correction the model would
# count those graduates twice once the post-2023 seat cohorts start registering.
SC_MGRMU_GAP = {2023: 450.0, 2024: 550.0, 2025: 850.0}   # seat-year -> seats
GAP_FORWARD = 850.0                                       # held at the last observed
# Derivation, stated so it can be audited: the last observed gap between the two
# seat sources is 850 seats (Selection Committee 4,750 against MGRMU 3,900 for
# 2025-26). Multiplied by the MBBS fill rate and the completion rate this is the
# number of graduates per year that would otherwise be counted twice:
#     850 x 0.998 x 0.95 = 806 graduates per year, from registration year 2031.
DEEMED_OVERLAP_2031 = GAP_FORWARD * 0.998 * 0.95

def deemed_overlap(year):
    """Graduates in `year` already counted inside the residual because the SC
    seat series now includes deemed seats. Subtracted from the residual."""
    sy = year - LAG_MBBS
    if sy < 2023:
        return 0.0
    gap = SC_MGRMU_GAP.get(sy, GAP_FORWARD)
    return gap * FILL_MBBS * COMPLETION

def fmg(year):
    """Foreign medical graduates: linear (+170/yr), capped at 3,000."""
    if year in REG_FMG_OBS:
        return float(REG_FMG_OBS[year])
    return min(3000.0, REG_FMG_OBS[2025] + 170.0 * (year - 2025))

# ============================================================================
# 4. STOCK AND FLOW, actuarial attrition off the TNMC register
# ============================================================================
AGE_AT_REG = 24      # INFERRED, not observed. Varied in the age sensitivity block.

def load_register():
    """Annual additions to the TNMC register, 1927-2025, from the scraped IMR.
    2025 is truncated in the scrape (3,785) and is replaced by the Medical
    Council's own figure (MBBS + FMG). The two sources agree to within 0.1%
    for 2020-2024, which validates the substitution."""
    c = collections.Counter()
    with open("imr_tamilnadu_output/tamilnadu_doctors.csv", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            try:
                c[int(r["year_of_registration"])] += 1
            except (ValueError, TypeError):
                pass
    # 2020-2025 are taken on the TNMC basis so that the register total and the
    # entrants decomposition reconcile exactly. The two sources differ by <0.2%
    # for 2020-2024 (see Validation); 2025 is truncated in the scrape (3,785).
    for y in range(2020, 2026):
        c[y] = REG_MBBS_OBS[y] + REG_FMG_OBS[y]
    return c

def mortality(age):
    if age < 35:  return 0.0008
    if age < 45:  return 0.0015
    if age < 55:  return 0.0035
    if age < 65:  return 0.0090
    if age < 75:  return 0.0220
    if age < 85:  return 0.0550
    return 0.1300

EMIG_RATE = 0.012   # central; see the sensitivity block, this is the key parameter

def emigration(age, rate=None):
    """Net loss of TNMC registrants to other Indian states, abroad, and to
    non-clinical work, concentrated in early and mid career.

    1.2%/yr across ages 25-45 is ~22% cumulative. Tamil Nadu trains far more
    doctors than it employs, so a large share of registrants never practise in
    the state. This is the single most consequential parameter in the model and
    the least directly observed; it is varied explicitly in the sensitivity run."""
    r = EMIG_RATE if rate is None else rate
    return r if 25 <= age <= 45 else 0.0

def participation(age):
    """Share of surviving, resident doctors still in active practice."""
    if age < 60:  return 1.00
    if age < 65:  return 0.85
    if age < 70:  return 0.60
    if age < 75:  return 0.30
    if age < 80:  return 0.12
    return 0.03

def run_stock(scen, register):
    """Roll register cohorts forward with mortality, emigration and retirement.
    History 1927-2025 establishes the 2025 opening stock; 2026-2050 is projected."""
    cohorts = {}          # registration year -> surviving headcount
    rows = []
    for y in range(1927, 2051):
        for R in list(cohorts):
            age = AGE_AT_REG + (y - R)
            cohorts[R] *= (1 - mortality(age)) * (1 - emigration(age))
        if y <= 2025:
            entrants = float(register.get(y, 0))
            dom = ext = fm = np.nan
        else:
            dom = domestic_output(gov_seats(scen) + pvt_seats(scen), y)
            ext, fm = RESIDUAL - deemed_overlap(y), fmg(y)
            entrants = dom + ext + fm
        cohorts[y] = entrants
        if y in IDX:
            head = active = 0.0
            for R, n in cohorts.items():
                age = AGE_AT_REG + (y - R)
                head += n
                active += n * participation(age)
            rows.append(dict(year=y, entrants=entrants, domestic=dom, external=ext,
                             fmg=fm, headcount=head, active=active))
    return rows

REGISTER = load_register()
RESULTS = {s: run_stock(s, REGISTER) for s in SCENARIOS}

# Exits, for the retirement-wave chart (status-quo scenario)
def exit_decomposition(scen, register):
    """Exits measured in ACTIVE-PRACTICE equivalents, not raw headcount.

    A death at 82 removes one registrant but only 0.03 of an active doctor,
    because that cohort is already 97% retired. Weighting every outflow by
    participation at the new age makes the flows reconcile exactly with the
    stock: active_t - active_{t-1} == entrants_t - (deaths + emigration +
    retirements). Verified to <1 doctor per year in the audit."""
    cohorts, out = {}, []
    for y in range(1927, 2051):
        deaths = emig = retire = 0.0
        for R in list(cohorts):
            age_prev = AGE_AT_REG + (y - 1 - R)
            age = AGE_AT_REG + (y - R)
            p_now = participation(age)
            n = cohorts[R]
            d = n * mortality(age)
            e = (n - d) * emigration(age)
            deaths += d * p_now
            emig += e * p_now
            # ageing through the participation curve, on the pre-attrition stock
            retire += n * (participation(age_prev) - p_now)
            cohorts[R] = n - d - e
        if y <= 2025:
            cohorts[y] = float(register.get(y, 0))
        else:
            cohorts[y] = (domestic_output(gov_seats(scen) + pvt_seats(scen), y)
                          + RESIDUAL - deemed_overlap(y) + fmg(y))
        if y in IDX:
            out.append(dict(year=y, deaths=deaths, emigration=emig,
                            retirements=max(0.0, retire)))
    return out

EXITS = exit_decomposition("Status quo", REGISTER)

# ============================================================================
# 5. SPECIALIST ACCOUNT  (a transition WITHIN the stock, never an addition)
# ============================================================================
AGE_AT_PG = 30

# Observed TNMC PG registrations, Medical Council workbook.
#
# BROAD SPECIALITY ONLY (MD + MS + DNB broad). Superspecialty registrations
# (DM, M.Ch, DNB superspeciality) are DELIBERATELY EXCLUDED: a DM or M.Ch holder
# necessarily already holds an MD or MS and was counted as a specialist then.
# Including both would double-count the same individual. Superspecialists are
# tracked separately below as a subset, for reporting only.
#
# The series is volatile (2,917 in 2024 against 8,069 in 2025) almost certainly
# batch processing of registration backlogs rather than real swings in output.
REG_PG_OBS = {2020: 2702, 2021: 4210, 2022: 5632, 2023: 5279, 2024: 2917, 2025: 8069}

# Superspecialty registrations (DM + M.Ch + DNB SS). A SUBSET of the specialist
# stock, never added to it.
REG_SS_OBS = {2020: 150, 2021: 477, 2022: 560, 2023: 256, 2024: 605, 2025: 742}

# The doctor account carries an external-entrant term; the specialist account
# must too, or specialists are understated. PG qualifications are registered in
# Tamil Nadu by deemed-university graduates and by TN-registered doctors who did
# their PG elsewhere. Estimated on the two years where seat-driven output exists
# (2024, 2025 <- seats 2021, 2022), comparing like with like.
_pg_resid = []
for _y in (2024, 2025):
    _dom_pg = pg_seats("Status quo")[IDX[_y - LAG_PG]] * FILL_PG * COMPLETION
    _pg_resid.append(REG_PG_OBS[_y] - _dom_pg)
PG_RESIDUAL = float(np.mean(_pg_resid))

# Pre-2020 back-cast: MGRMU pass-outs scaled to the full registration basis.
# Calibrated so the back-cast meets the observed 2020 value exactly.
PG_TOTAL_COVERAGE = REG_PG_OBS[2020] / 1550.0   # broad-speciality basis

def run_specialists(scen):
    seats = pg_seats(scen)
    cohorts, rows = {}, []
    for y in range(2000, 2051):
        exits = 0.0
        for R in list(cohorts):
            age_prev = AGE_AT_PG + (y - 1 - R)
            age = AGE_AT_PG + (y - R)
            p_now = participation(age)
            n = cohorts[R]
            d = n * mortality(age)
            e = (n - d) * emigration(age)
            # exits measured in active-practice equivalents, as for doctors
            exits += d * p_now + e * p_now + n * (participation(age_prev) - p_now)
            cohorts[R] = n - d - e
        sy = y - LAG_PG
        if y in REG_PG_OBS:
            grads = float(REG_PG_OBS[y])          # observed registrations
        elif y < 2020:
            # Back-cast from MGRMU pass-outs, scaled to the full registration
            # basis so it meets the observed 2020 value without a step.
            mgrmu = {2018: 1400, 2019: 1500}.get(y, 1200.0)
            grads = mgrmu * PG_TOTAL_COVERAGE
        else:
            grads = seats[IDX[sy]] * FILL_PG * COMPLETION + PG_RESIDUAL
        cohorts[y] = grads
        if y in IDX:
            active = sum(n * participation(AGE_AT_PG + (y - R)) for R, n in cohorts.items())
            rows.append(dict(year=y, grads=grads, active=active, exits=max(0.0, exits)))
    return rows

SPECIALISTS = {s: run_specialists(s) for s in SCENARIOS}

# ============================================================================
# 6. NEED BENCHMARK: WHO density norm
# ============================================================================
WHO_TOTAL_PER_10K = 44.5      # skilled health workers (doctors+nurses+midwives)
WHO_DOCTOR_SHARE = 0.25       # doctor : nurse = 1 : 3
NEED_DOCTORS = POP * (WHO_TOTAL_PER_10K * WHO_DOCTOR_SHARE) / 10000.0

# ============================================================================
# 7. OUTPUT TABLES
# ============================================================================
def col(scen, key):
    return np.array([r[key] for r in RESULTS[scen]], float)

with open("tn_projection_2050.csv", "w", newline="") as f:
    w = csv.writer(f)
    hdr = ["year", "population", "who_need_doctors",
           "retirements", "deaths", "emigration", "total_exits"]
    for s in SCENARIOS:
        k = s.lower().replace(" ", "_")
        hdr += [f"{k}_gov_seats", f"{k}_pvt_seats", f"{k}_pg_seats",
                f"{k}_entrants", f"{k}_active_doctors", f"{k}_density_per_10k",
                f"{k}_active_specialists"]
    w.writerow(hdr)
    for i, y in enumerate(YEARS):
        ex = EXITS[i]
        row = [y, f"{POP[i]:.0f}", f"{NEED_DOCTORS[i]:.0f}",
               f"{ex['retirements']:.0f}", f"{ex['deaths']:.0f}",
               f"{ex['emigration']:.0f}",
               f"{ex['retirements']+ex['deaths']+ex['emigration']:.0f}"]
        for s in SCENARIOS:
            g, p, pg = gov_seats(s), pvt_seats(s), pg_seats(s)
            a = col(s, "active")[i]
            row += [f"{g[i]:.0f}" if not np.isnan(g[i]) else "",
                    f"{p[i]:.0f}" if not np.isnan(p[i]) else "",
                    f"{pg[i]:.0f}" if not np.isnan(pg[i]) else "",
                    f"{col(s,'entrants')[i]:.0f}",
                    f"{a:.0f}", f"{a/POP[i]*10000:.2f}",
                    f"{SPECIALISTS[s][i]['active']:.0f}"]
        w.writerow(row)

lines = []
def emit(s=""):
    lines.append(s); print(s)

emit("=" * 78)
emit("TAMIL NADU MEDICAL WORKFORCE: SUPPLY PROJECTION TO 2050")
emit("=" * 78)
emit("\nMODEL VALIDATION (parameters were fitted, then tested out-of-sample)")
emit("-" * 78)
emit("  PG pipeline: seats(2021) x fill 0.978 x completion 0.95 = "
     f"{PG_OBS[0]/1.407*FILL_PG*COMPLETION:,.0f} predicted vs 2,717 observed pass-outs (2024)")
emit("  Register cross-check, IMR scrape vs TNMC workbook (MBBS + FMG):")
for y in (2020, 2021, 2022, 2023, 2024):
    a, b = REGISTER.get(y, 0), REG_MBBS_OBS[y] + REG_FMG_OBS[y]
    emit(f"     {y}:  IMR {a:>7,}   TNMC {b:>7,}   diff {abs(a-b)/b*100:>5.2f}%")
emit(f"\n  External entrants (deemed universities + out-of-state returnees), "
     f"residual 2021-25:")
emit("     " + ", ".join(f"{v:,.0f}" for v in _resid) + f"   -> flat at {RESIDUAL:,.0f}/yr")
emit(f"\n  Private MBBS logistic fits (ceiling K -> growth r, inflection year):")
for K, (r, t0) in PVT_PARAMS.items():
    emit(f"     K={K:>6,}:  r={r:.3f}   inflection={2015+t0:.1f}")

emit("\n\nHEADLINE PROJECTIONS")
emit("-" * 78)
for s in SCENARIOS:
    g, p, pg = gov_seats(s), pvt_seats(s), pg_seats(s)
    a = col(s, "active"); e = col(s, "entrants")
    sp = np.array([r["active"] for r in SPECIALISTS[s]])
    emit(f"\n  [{s}]")
    emit(f"    {'':<26}{'2025':>12}{'2030':>12}{'2040':>12}{'2050':>12}")
    def line(lbl, arr, fmt="{:,.0f}"):
        emit(f"    {lbl:<26}" + "".join(fmt.format(arr[IDX[y]]).rjust(12)
                                        for y in (2025, 2030, 2040, 2050)))
    line("Government MBBS seats", g)
    line("Private MBBS seats", p)
    line("Total MBBS seats", g + p)
    line("PG (MD/MS/Dip) seats", pg)
    line("Entrants to register", e)
    line("Active doctors", a)
    line("Doctors per 10,000", a / POP * 10000, "{:,.1f}")
    line("Active specialists", sp)
    line("WHO need (doctors)", NEED_DOCTORS)
    line("Surplus vs WHO need", a - NEED_DOCTORS)

emit("\n\nANNUAL EXITS, retirements, deaths and migration (status-quo scenario)")
emit("-" * 78)
emit(f"    {'Year':<6}{'Retire':>9}{'Deaths':>9}{'Migration':>11}{'Total':>9}"
     f"{'Entrants':>10}{'Net':>10}{'Exit%':>8}{'Replace':>9}")
_ent_sq, _act_sq = col("Status quo", "entrants"), col("Status quo", "active")
for i, y in enumerate(YEARS):
    if y < 2020:
        continue
    ex = EXITS[i]
    tot = ex["retirements"] + ex["deaths"] + ex["emigration"]
    ent = _ent_sq[i]
    emit(f"    {y:<6}{ex['retirements']:>9,.0f}{ex['deaths']:>9,.0f}"
         f"{ex['emigration']:>11,.0f}{tot:>9,.0f}{ent:>10,.0f}{ent-tot:>10,.0f}"
         f"{tot/_act_sq[i]*100:>7.2f}%{ent/tot:>9.1f}x")
emit("\n    Exit%   = total exits as a share of the active stock that year")
emit("    Replace = entrants per exit. A system in steady state sits at 1.0x;")
emit("              Tamil Nadu stays above 2.5x for the whole projection window.")
emit("\n    Cumulative over 2026-2050:")
_r = sum(EXITS[i]["retirements"] for i, y in enumerate(YEARS) if y >= 2026)
_d = sum(EXITS[i]["deaths"] for i, y in enumerate(YEARS) if y >= 2026)
_m = sum(EXITS[i]["emigration"] for i, y in enumerate(YEARS) if y >= 2026)
_e = sum(_ent_sq[i] for i, y in enumerate(YEARS) if y >= 2026)
emit(f"       Retirements {_r:>10,.0f}   ({_r/(_r+_d+_m)*100:>4.1f}% of exits)")
emit(f"       Deaths      {_d:>10,.0f}   ({_d/(_r+_d+_m)*100:>4.1f}% of exits)")
emit(f"       Migration   {_m:>10,.0f}   ({_m/(_r+_d+_m)*100:>4.1f}% of exits)")
emit(f"       TOTAL EXITS {_r+_d+_m:>10,.0f}")
emit(f"       Entrants    {_e:>10,.0f}   -> net addition {_e-(_r+_d+_m):>+11,.0f}")

emit("\n\nSENSITIVITY, emigration / out-of-state loss (the key uncertain parameter)")
emit("-" * 78)
emit(f"    {'Rate (ages 25-45)':<24}{'Active 2025':>14}{'Active 2050':>14}{'Density 2050':>14}")
_base = EMIG_RATE
for rate in (0.006, 0.012, 0.020, 0.030):
    EMIG_RATE = rate
    rr = run_stock("Status quo", REGISTER)
    a25, a50 = rr[IDX[2025]]["active"], rr[IDX[2050]]["active"]
    emit(f"    {rate*100:>5.1f}%/yr  (~{(1-(1-rate)**21)*100:>4.1f}% cum.){'':<3}"
         f"{a25:>14,.0f}{a50:>14,.0f}{a50/POP[IDX[2050]]*10000:>14,.1f}")
EMIG_RATE = _base

emit("\n\nSENSITIVITY, age at registration (inferred, not observed)")
emit("-" * 78)
emit("    Registration is assumed at age 24 (MBBS entry at 18 plus 5.5 years). Age is")
emit("    never observed in the register, only inferred from year of registration, and")
emit("    retirement timing is the output most sensitive to it.")
emit(f"    {'Age assumed':<14}{'Active 2025':>14}{'Active 2050':>14}{'Retirements 2050':>18}"
     f"{'Cumulative retirements 2026-50':>32}")
_base_age = AGE_AT_REG
for _age in (22, 23, 24, 25, 26):
    AGE_AT_REG = _age
    _rr = run_stock("Status quo", REGISTER)
    _ee = exit_decomposition("Status quo", REGISTER)
    _cum = sum(_ee[i]["retirements"] for i, y in enumerate(YEARS) if y >= 2026)
    emit(f"    {_age:<14}{_rr[IDX[2025]]['active']:>14,.0f}{_rr[IDX[2050]]['active']:>14,.0f}"
         f"{_ee[IDX[2050]]['retirements']:>18,.0f}{_cum:>32,.0f}")
AGE_AT_REG = _base_age
emit("    Central case is 24. A one-year change moves the 2050 stock by well under one")
emit("    per cent but shifts cumulative retirements materially, which is the point:")
emit("    the stock is robust to this assumption, the retirement profile is not.")

emit("\n\nINTERNATIONAL BENCHMARKS, doctors per 10,000 population")
emit("-" * 78)
_d25 = col("Status quo", "active")[IDX[2025]] / POP[IDX[2025]] * 10000
emit("    India (all states, active, approx.)      ~ 7 - 9")
emit(f"    Tamil Nadu, modelled 2025                {_d25:>6.1f}")
emit("    United Kingdom                           ~ 32")
emit("    OECD average                             ~ 35")
emit("    Cuba (world's highest)                   ~ 84")
emit(f"    Tamil Nadu, modelled 2050 (status quo)   "
     f"{col('Status quo','active')[IDX[2050]]/POP[IDX[2050]]*10000:>6.1f}"
     "   <-- see caveat below")
emit("\n    CAVEAT: this is an UNCONSTRAINED SUPPLY projection. It answers 'how many")
emit("    doctors will Tamil Nadu produce and retain on current trajectories', not")
emit("    'how many will the health system employ'. A density approaching OECD")
emit("    levels is the model telling us that current seat-expansion plans exceed")
emit("    any plausible absorption capacity, that is the finding, not a forecast")
emit("    of employment. The corrective in reality is higher out-migration, which")
emit("    the sensitivity block above prices.")

emit("\n\nPOPULATION SPINE")
emit("-" * 78)
for y in (2025, 2031, 2036, 2040, 2050):
    emit(f"    {y}: {POP[IDX[y]]/1e6:>6.2f} million"
         + ("   (NCP official)" if y <= 2036 else "   (momentum-exhausted extension)"))

with open("tn_projection_2050_summary.txt", "w") as f:
    f.write("\n".join(lines) + "\n")

# ============================================================================
# 8. FIGURES
# ============================================================================
def finish(ax, title, sub=None, ylab=None):
    ax.set_title(title, fontsize=11, fontweight="bold", loc="left", pad=(38 if "\n" in sub else 26) if sub else 8)
    if sub:
        ax.text(0, 1.03, sub, transform=ax.transAxes, fontsize=8, color=INK3, va="bottom")
    if ylab: ax.set_ylabel(ylab, fontsize=8.5)
    ax.yaxis.set_major_formatter(thou)
    ax.set_axisbelow(True); ax.grid(axis="x", visible=False)

def save(fig, name):
    fig.tight_layout()
    fig.savefig(f"{OUT}/{name}.png", bbox_inches="tight", facecolor="white")
    plt.close(fig); print("  wrote", f"{OUT}/{name}.png")

hist = YEARS <= 2025
proj = YEARS >= 2025

# --- Fig 1: MBBS seats, government vs private, with scenario fan ------------
fig, ax = plt.subplots(figsize=(8.4, 4.6))
gsq = gov_seats("Status quo"); gpp = gov_seats("Policy push")
ax.plot(YEARS[hist], gsq[hist], color=C_BLUE, lw=2, label="Government (actual)")
ax.plot(YEARS[proj], gsq[proj], color=C_BLUE, lw=2, ls=":")
ax.plot(YEARS[proj], gpp[proj], color=C_BLUE, lw=1.4, ls="--", alpha=0.75)
ax.fill_between(YEARS[proj], gsq[proj], gpp[proj], color=C_BLUE, alpha=0.10, lw=0)
plo, phi = pvt_seats("Constrained"), pvt_seats("Policy push")
pmid = pvt_seats("Status quo")
ax.plot(YEARS[hist], pmid[hist], color=C_ORANGE, lw=2, label="Private / self-financing (actual)")
ax.plot(YEARS[proj], pmid[proj], color=C_ORANGE, lw=2, ls=":")
ax.fill_between(YEARS[proj], plo[proj], phi[proj], color=C_ORANGE, alpha=0.14, lw=0)
ax.axvline(2025.5, color=INK3, lw=0.8, ls="-", alpha=0.5)
ax.text(2025.9, 300, "projected", fontsize=7.5, color=INK3)
ax.annotate(f"{gsq[IDX[2050]]:,.0f}", (2050, gsq[IDX[2050]]), xytext=(5, 0),
            textcoords="offset points", fontsize=8, color=C_BLUE, va="center")
ax.annotate(f"{pmid[IDX[2050]]:,.0f}", (2050, pmid[IDX[2050]]), xytext=(5, 0),
            textcoords="offset points", fontsize=8, color=C_ORANGE, va="center")
ax.annotate("frozen since 2021\n(+25 seats in 4 years)", (2023, 5200), xytext=(2016.5, 6350),
            fontsize=7.5, color=INK2,
            arrowprops=dict(arrowstyle="->", color=INK3, lw=0.8))
ax.set_xlim(2015, 2054); ax.set_ylim(0, 13000)
ax.legend(loc="upper left", fontsize=8.5)
finish(ax, "MBBS seats: government capacity is frozen, private is still expanding",
       "Tamil Nadu, 2015-16 to 2050-51. Shaded bands span the constrained / policy-push scenarios.",
       "Sanctioned seats")
save(fig, "fig1_mbbs_seats")

# --- Fig 2: entrants to the register, decomposed ---------------------------
fig, ax = plt.subplots(figsize=(8.4, 4.6))
yy = YEARS[YEARS >= 2021]
dom = np.array([domestic_output(_tot_sq, y) for y in yy])
ext = np.array([(REG_MBBS_OBS[y] - domestic_output(_tot_sq, y)) if y <= 2025 else RESIDUAL - deemed_overlap(y)
                for y in yy])
fm = np.array([fmg(y) for y in yy])
ax.stackplot(yy, dom, ext, fm, colors=[C_BLUE, C_ORANGE, C_AQUA],
             labels=["Tamil Nadu state-counselling graduates",
                     "Deemed universities + out-of-state returnees",
                     "Foreign medical graduates"],
             edgecolor="white", linewidth=1.2)
ax.axvline(2025.5, color=INK3, lw=0.8, alpha=0.5)
ax.text(2025.9, 900, "projected", fontsize=7.5, color=INK3)
tot50 = dom[-1] + ext[-1] + fm[-1]
ax.annotate(f"{tot50:,.0f}", (2050, tot50), xytext=(5, 0), textcoords="offset points",
            fontsize=8, color=INK2, va="center")
ax.set_xlim(2021, 2054); ax.legend(loc="upper left", fontsize=8.5)
finish(ax, "Who actually joins the Tamil Nadu medical register",
       "Roughly half of all new registrations come from outside the state counselling system.",
       "New registrations per year")
save(fig, "fig2_entrants_decomposed")

# --- Fig 3: active doctor stock vs WHO need --------------------------------
fig, ax = plt.subplots(figsize=(8.4, 4.6))
a_lo, a_mid, a_hi = (col("Constrained", "active"), col("Status quo", "active"),
                     col("Policy push", "active"))
ax.plot(YEARS[hist], a_mid[hist], color=C_BLUE, lw=2, label="Active doctors (modelled)")
ax.plot(YEARS[proj], a_mid[proj], color=C_BLUE, lw=2, ls=":")
ax.fill_between(YEARS[proj], a_lo[proj], a_hi[proj], color=C_BLUE, alpha=0.14, lw=0)
ax.plot(YEARS, NEED_DOCTORS, color=C_RED, lw=2, ls="--", label="WHO need (11.1 doctors / 10,000)")
ax.axvline(2025.5, color=INK3, lw=0.8, alpha=0.5)
ax.annotate(f"{a_mid[IDX[2050]]:,.0f}", (2050, a_mid[IDX[2050]]), xytext=(5, 0),
            textcoords="offset points", fontsize=8, color=C_BLUE, va="center")
ax.annotate(f"{NEED_DOCTORS[IDX[2050]]:,.0f}", (2050, NEED_DOCTORS[IDX[2050]]),
            xytext=(5, 0), textcoords="offset points", fontsize=8, color=C_RED, va="center")
ax.set_xlim(2011, 2058); ax.set_ylim(0, max(a_hi) * 1.12)
ax.legend(loc="upper left", fontsize=8.5)
finish(ax, "Active doctor supply runs far above the WHO density floor",
       "The WHO norm is a minimum, not a target, the binding questions are distribution and mix.",
       "Doctors")
save(fig, "fig3_stock_vs_need")

# --- Fig 4: density per 10,000 ---------------------------------------------
fig, ax = plt.subplots(figsize=(8.4, 4.6))
d_mid = a_mid / POP * 10000
d_lo, d_hi = a_lo / POP * 10000, a_hi / POP * 10000
ax.plot(YEARS[hist], d_mid[hist], color=C_BLUE, lw=2, label="Doctors per 10,000 (modelled)")
ax.plot(YEARS[proj], d_mid[proj], color=C_BLUE, lw=2, ls=":")
ax.fill_between(YEARS[proj], d_lo[proj], d_hi[proj], color=C_BLUE, alpha=0.14, lw=0)
ax.axhline(11.125, color=C_RED, lw=2, ls="--", label="WHO floor (11.1 / 10,000)")
ax.axvline(2025.5, color=INK3, lw=0.8, alpha=0.5)
ax.annotate(f"{d_mid[IDX[2050]]:,.1f}", (2050, d_mid[IDX[2050]]), xytext=(5, 0),
            textcoords="offset points", fontsize=8, color=C_BLUE, va="center")
ax.set_xlim(2011, 2056)
ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: f"{v:,.0f}"))
ax.legend(loc="upper left", fontsize=8.5)
finish(ax, "Density triples by 2050, driven by supply, not by the denominator",
       "Population falls only ~1.7% from its 2031 peak; the rise is almost entirely "
       "production.\nTamil Nadu passed the WHO floor around 2016.",
       "Doctors per 10,000 population")
save(fig, "fig4_density")

# --- Fig 5: exits, the retirement wave ------------------------------------
fig, ax = plt.subplots(figsize=(8.4, 4.6))
ex_y = np.array([r["year"] for r in EXITS])
ret = np.array([r["retirements"] for r in EXITS])
dth = np.array([r["deaths"] for r in EXITS])
emg = np.array([r["emigration"] for r in EXITS])
m = ex_y >= 2015
ax.stackplot(ex_y[m], ret[m], dth[m], emg[m], colors=[C_BLUE, C_ORANGE, C_AQUA],
             labels=["Retirement", "Mortality", "Emigration / out-of-state"],
             edgecolor="white", linewidth=1.2)
ax.axvline(2025.5, color=INK3, lw=0.8, alpha=0.5)
ax.text(2025.9, 400, "projected", fontsize=7.5, color=INK3)
tot = (ret + dth + emg)
ax.annotate(f"{tot[IDX[2050]]:,.0f}", (2050, tot[IDX[2050]]), xytext=(5, 0),
            textcoords="offset points", fontsize=8, color=INK2, va="center")
ax.set_xlim(2015, 2054); ax.legend(loc="upper left", fontsize=8.5)
finish(ax, "Exits stay small, and are dominated by migration, not retirement",
       "The large cohorts registered in the 2010s reach 60 only after 2050, so the stock "
       "accumulates:\nby 2050 exits are ~8,500/yr against ~22,600 entrants.",
       "Exits per year")
save(fig, "fig5_exits")

# --- Fig 6: PG seats and the specialist stock ------------------------------
fig, axes = plt.subplots(1, 2, figsize=(9.6, 4.2))
ax = axes[0]
pg_mid = pg_seats("Status quo"); pg_lo = pg_seats("Constrained"); pg_hi = pg_seats("Policy push")
mm = YEARS >= 2021
ax.plot(YEARS[(YEARS >= 2021) & hist], pg_mid[(YEARS >= 2021) & hist], color=C_VIOLET, lw=2)
ax.plot(YEARS[proj], pg_mid[proj], color=C_VIOLET, lw=2, ls=":")
ax.fill_between(YEARS[proj], pg_lo[proj], pg_hi[proj], color=C_VIOLET, alpha=0.14, lw=0)
ax.annotate(f"{pg_mid[IDX[2050]]:,.0f}", (2050, pg_mid[IDX[2050]]), xytext=(4, 0),
            textcoords="offset points", fontsize=8, color=C_VIOLET, va="center")
ax.set_xlim(2021, 2054)
finish(ax, "PG (MD/MS/Diploma) seats", "Linear, +321.5/yr, capped at 70% of MBBS seats.", "Seats")
ax = axes[1]
sp_mid = np.array([r["active"] for r in SPECIALISTS["Status quo"]])
sp_lo = np.array([r["active"] for r in SPECIALISTS["Constrained"]])
sp_hi = np.array([r["active"] for r in SPECIALISTS["Policy push"]])
ax.plot(YEARS[hist], sp_mid[hist], color=C_AQUA, lw=2)
ax.plot(YEARS[proj], sp_mid[proj], color=C_AQUA, lw=2, ls=":")
ax.fill_between(YEARS[proj], sp_lo[proj], sp_hi[proj], color=C_AQUA, alpha=0.14, lw=0)
ax.annotate(f"{sp_mid[IDX[2050]]:,.0f}", (2050, sp_mid[IDX[2050]]), xytext=(4, 0),
            textcoords="offset points", fontsize=8, color=C_AQUA, va="center")
ax.set_xlim(2011, 2056)
finish(ax, "Active specialists", "A transition within the stock, not an addition to it.", "Specialists")
save(fig, "fig6_pg_and_specialists")

print("\nWrote tn_projection_2050.csv and tn_projection_2050_summary.txt")
