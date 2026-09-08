"""
Tamil Nadu health workforce projection engine.

A pure, parameterised reimplementation of tn_supply_model.py and
tn_cadre_model.py with no side effects: it prints nothing, writes nothing and
draws nothing. Every assumption is an argument with a default equal to the
validated model value, so the dashboard can vary any of them.

Reproduces the published model exactly at default settings. Verified by
verify_engine.py.
"""
from __future__ import annotations
from dataclasses import dataclass, field, asdict
import csv, os, re, statistics
import numpy as np

# Project root, one level up from this file.
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REGISTER_CSV = os.path.join(ROOT, "imr_tamilnadu_output", "tamilnadu_doctors.csv")
# Publishable aggregates. These carry no personal data: the register is reduced
# to a count per year, and the cadre workbooks to the three series the model
# actually uses. The engine prefers these and falls back to the raw sources.
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
REGISTER_AGG = os.path.join(DATA_DIR, "register_by_year.csv")
CADRE_AGG = os.path.join(DATA_DIR, "cadre_series.json")
SEAT_FILE = os.path.join(ROOT, "new data", "TN MGRMU", "SEAT COUNT STAT 15052026.xlsx")
PASS_FILE = os.path.join(ROOT, "new data", "TN MGRMU", "passout COUNT STAT 15052026.xlsx")

# ---------------------------------------------------------------------------
# Observed source data. Short series are held here so the dashboard can show
# and edit them; the large register and the cadre workbooks are read from disk.
# ---------------------------------------------------------------------------
NCP_POP = {2011: 72147, 2012: 72645, 2013: 73142, 2014: 73640, 2015: 74137,
           2016: 74635, 2017: 74989, 2018: 75342, 2019: 75695, 2020: 76049,
           2021: 76402, 2022: 76631, 2023: 76860, 2024: 77089, 2025: 77317,
           2026: 77546, 2027: 77653, 2028: 77761, 2029: 77868, 2030: 77975,
           2031: 78082, 2032: 78079, 2033: 78076, 2034: 78073, 2035: 78070,
           2036: 78067}                                        # thousands

SEAT_YEARS = list(range(2015, 2026))
GOV_SEATS_OBS = [2655, 2650, 3050, 2900, 3600, 3675, 5175, 5175, 5200, 5200, 5200]
PVT_SEATS_OBS = [1010, 1610, 1600, 1600, 1950, 2350, 2900, 3350, 3850, 4050, 4750]

PG_YEARS = list(range(2021, 2026))
PG_SEATS_OBS = [4100, 4283, 4435, 4630, 5534]                  # full basis

REG_MBBS_OBS = {2020: 8409, 2021: 8058, 2022: 10367, 2023: 9370,
                2024: 9722, 2025: 10839}
REG_FMG_OBS = {2020: 640, 2021: 1268, 2022: 1287, 2023: 1622,
               2024: 1427, 2025: 1606}
# Broad speciality only. Superspecialty is deliberately excluded: a DM or M.Ch
# holder already counted as a specialist at MD or MS.
REG_PG_OBS = {2020: 2702, 2021: 4210, 2022: 5632, 2023: 5279,
              2024: 2917, 2025: 8069}

SC_MGRMU_GAP = {2023: 450.0, 2024: 550.0, 2025: 850.0}

# Government seat scenarios, built from the 2025 college distribution:
# 37 colleges, 5,200 seats, sixteen at 100, sixteen at 150, one at 200,
# four at the NMC ceiling of 250.
GOV_SCENARIOS = {
    "Frozen at 5,200 (no action)":            (5200.0, None),
    "Consolidation, 6,000 by 2035":           (6000.0, 2035),
    "Full capacity, 9,250 by 2045":           (9250.0, 2045),
}
PVT_SCENARIOS = {
    "Constrained, ceiling 8,000":  (8000.0, 0.226, 8.5),
    "Status quo, ceiling 10,000":  (10000.0, 0.204, 10.5),
    "Expansion, ceiling 12,000":   (12000.0, 0.188, 12.25),
}

DEFAULT_MORTALITY = [(35, 0.0008), (45, 0.0015), (55, 0.0035), (65, 0.0090),
                     (75, 0.0220), (85, 0.0550), (999, 0.1300)]
DEFAULT_PARTICIPATION = [(60, 1.00), (65, 0.85), (70, 0.60), (75, 0.30),
                         (80, 0.12), (999, 0.03)]


@dataclass
class Params:
    """Every assumption in the doctor model. Defaults are the validated values."""
    end_year: int = 2050
    # Population
    pop_growth_2050: float = -0.0035
    # Government seats
    gov_2025: float = 5200.0
    gov_target: float = 6000.0
    gov_target_year: int | None = 2035
    # Private seats, logistic
    pvt_ceiling: float = 10000.0
    pvt_growth_r: float = 0.204
    pvt_midpoint: float = 10.5          # years after 2015
    # PG seats
    pg_intercept: float = 3953.4
    pg_slope: float = 321.5
    pg_cap_share: float = 0.70
    # Pipeline
    fill_mbbs: float = 0.998
    fill_pg: float = 0.978
    completion: float = 0.95
    lag_mbbs: int = 6
    lag_pg: int = 3
    # Entrants
    external_residual: float = 5381.0475
    gap_forward: float = 850.0
    fmg_base: float = 1606.0
    fmg_increment: float = 170.0
    fmg_ceiling: float = 3000.0
    pg_residual: float = 1598.67735
    # Stock
    age_at_registration: int = 24
    age_at_pg: int = 30
    emigration_rate: float = 0.012
    emigration_age_lo: int = 25
    emigration_age_hi: int = 45
    mortality_bands: list = field(default_factory=lambda: list(DEFAULT_MORTALITY))
    participation_bands: list = field(default_factory=lambda: list(DEFAULT_PARTICIPATION))
    # Need benchmark
    who_norm: float = 44.5
    who_doctor_share: float = 0.25

    def to_dict(self):
        return asdict(self)


# ---------------------------------------------------------------------------
# Source data readers, cached
# ---------------------------------------------------------------------------
_REGISTER_CACHE = None

def load_register(path: str = REGISTER_CSV) -> dict:
    """Annual additions to the TNMC register, 1927 to 2025.

    2020 to 2025 are taken on the TNMC basis so the register total and the
    entrants decomposition reconcile exactly. The scrape and the TNMC workbook
    differ by under 0.2 per cent for those years, and the 2025 scrape is
    truncated at 3,785 against a true 12,445."""
    global _REGISTER_CACHE
    if _REGISTER_CACHE is not None:
        return dict(_REGISTER_CACHE)
    # Preferred: the aggregated year counts, which contain no personal data.
    if os.path.exists(REGISTER_AGG) and os.path.getsize(REGISTER_AGG) > 0:
        counts = {}
        with open(REGISTER_AGG, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                counts[int(row["year"])] = int(row["registrations"])
        for y in range(2020, 2026):
            counts[y] = REG_MBBS_OBS[y] + REG_FMG_OBS[y]
        _REGISTER_CACHE = dict(counts)
        return dict(counts)
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        raise FileNotFoundError(
            f"Register not readable: {path}\n"
            "If this file shows 0 bytes, Dropbox has it set to online only. "
            "Right-click the project folder in Finder and choose Make Available Offline."
        )
    counts = {}
    with open(path, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            try:
                y = int(row["year_of_registration"])
            except (TypeError, ValueError, KeyError):
                continue
            counts[y] = counts.get(y, 0) + 1
    for y in range(2020, 2026):
        counts[y] = REG_MBBS_OBS[y] + REG_FMG_OBS[y]
    _REGISTER_CACHE = dict(counts)
    return dict(counts)


def _norm(s):
    return re.sub(r"[^a-z]", "", str(s).lower())

_YCOLS = {"academicyear", "academicyr"}
_ECOLS = {"examyear"}
_SCOLS = {"sanctionedseats", "totalapprovedseats", "sanctionedseat"}
_ACOLS = {"studentadmitted", "admittedcount", "seatsfilled", "studentsadmitted", "stud"}
_PCOLS = {"studentspassedout", "studentpassedout"}


def _read_sheets(path, ycols, valuesets):
    import openpyxl
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        raise FileNotFoundError(
            f"Source workbook not readable: {os.path.basename(path)}\n"
            "If it shows 0 bytes, Dropbox has it set to online only."
        )
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


# ---------------------------------------------------------------------------
# Doctor model
# ---------------------------------------------------------------------------
def _band(bands, age):
    for cut, val in bands:
        if age < cut:
            return val
    return bands[-1][1]


def population(p: Params) -> dict:
    pop = {y: NCP_POP[y] * 1000.0 for y in NCP_POP}
    g0 = (NCP_POP[2036] - NCP_POP[2035]) / NCP_POP[2035]
    span = max(1, p.end_year - 2036)
    for i, y in enumerate(range(2037, p.end_year + 1), start=1):
        g = g0 + (p.pop_growth_2050 - g0) * (i / span)
        pop[y] = pop[y - 1] * (1 + g)
    return pop


def gov_seats(p: Params) -> dict:
    s = {}
    for y in range(2011, p.end_year + 1):
        if y < 2015:
            s[y] = float("nan")
        elif y <= 2025:
            s[y] = float(GOV_SEATS_OBS[SEAT_YEARS.index(y)])
        elif p.gov_target_year is None or y >= p.gov_target_year:
            s[y] = p.gov_target
        else:
            frac = (y - 2025) / (p.gov_target_year - 2025)
            s[y] = round((p.gov_2025 + (p.gov_target - p.gov_2025) * frac) / 50.0) * 50.0
    return s


def pvt_seats(p: Params) -> dict:
    s = {}
    for y in range(2011, p.end_year + 1):
        if y < 2015:
            s[y] = float("nan")
        elif y <= 2025:
            s[y] = float(PVT_SEATS_OBS[SEAT_YEARS.index(y)])
        else:
            s[y] = p.pvt_ceiling / (1.0 + np.exp(-p.pvt_growth_r * (y - 2015 - p.pvt_midpoint)))
    return s


def pg_seats(p: Params) -> dict:
    g, v = gov_seats(p), pvt_seats(p)
    s = {}
    for y in range(2011, p.end_year + 1):
        if y < 2021:
            s[y] = float("nan")
            continue
        if y <= 2025:
            val = float(PG_SEATS_OBS[PG_YEARS.index(y)])
        else:
            val = p.pg_intercept + p.pg_slope * (y - 2021)
            mbbs = g[y] + v[y]
            if mbbs == mbbs:
                val = min(val, p.pg_cap_share * mbbs)
        s[y] = val
    return s


def deemed_overlap(p: Params, year: int) -> float:
    sy = year - p.lag_mbbs
    if sy < 2023:
        return 0.0
    gap = SC_MGRMU_GAP.get(sy, p.gap_forward)
    return gap * p.fill_mbbs * p.completion


def fmg(p: Params, year: int) -> float:
    if year in REG_FMG_OBS:
        return float(REG_FMG_OBS[year])
    return min(p.fmg_ceiling, p.fmg_base + p.fmg_increment * (year - 2025))


def run_doctors(p: Params, register: dict | None = None) -> dict:
    """Full doctor projection. Returns arrays keyed by year."""
    if register is None:
        register = load_register()
    years = list(range(2011, p.end_year + 1))
    pop, g, v, pg = population(p), gov_seats(p), pvt_seats(p), pg_seats(p)
    total_seats = {y: (g[y] + v[y]) for y in years}

    def domestic(y):
        sy = y - p.lag_mbbs
        if sy < 2015 or sy not in total_seats:
            return float("nan")
        return total_seats[sy] * p.fill_mbbs * p.completion

    dom, ext, fm, ent = {}, {}, {}, {}
    for y in years:
        dom[y] = domestic(y)
        if y <= 2025:
            ent[y] = float(register.get(y, 0))
            ext[y] = (REG_MBBS_OBS[y] - dom[y]) if y in REG_MBBS_OBS and dom[y] == dom[y] else float("nan")
            fm[y] = float(REG_FMG_OBS.get(y, float("nan")))
        else:
            ext[y] = p.external_residual - deemed_overlap(p, y)
            fm[y] = fmg(p, y)
            ent[y] = dom[y] + ext[y] + fm[y]

    mort = lambda a: _band(p.mortality_bands, a)
    part = lambda a: _band(p.participation_bands, a)
    emig = lambda a: p.emigration_rate if p.emigration_age_lo <= a <= p.emigration_age_hi else 0.0

    cohorts, active, retire, deaths, migr = {}, {}, {}, {}, {}
    for y in range(1927, p.end_year + 1):
        d_sum = e_sum = r_sum = 0.0
        for R in list(cohorts):
            a_prev = p.age_at_registration + (y - 1 - R)
            a = p.age_at_registration + (y - R)
            pn = part(a)
            n = cohorts[R]
            d = n * mort(a)
            e = (n - d) * emig(a)
            d_sum += d * pn
            e_sum += e * pn
            r_sum += n * (part(a_prev) - pn)
            cohorts[R] = n - d - e
        cohorts[y] = float(register.get(y, 0)) if y <= 2025 else ent.get(y, 0.0)
        if y in ent or y >= 2011:
            act = sum(n * part(p.age_at_registration + (y - R)) for R, n in cohorts.items())
            if y >= 2011:
                active[y] = act
                retire[y] = max(0.0, r_sum)
                deaths[y] = d_sum
                migr[y] = e_sum

    # Specialists, a transition within the stock and never an addition to it.
    sp_cohorts, spec, sp_grads, sp_exits = {}, {}, {}, {}
    pg_total_coverage = REG_PG_OBS[2020] / 1550.0
    for y in range(2000, p.end_year + 1):
        ex = 0.0
        for R in list(sp_cohorts):
            a_prev = p.age_at_pg + (y - 1 - R)
            a = p.age_at_pg + (y - R)
            pn = part(a)
            n = sp_cohorts[R]
            d = n * mort(a)
            e = (n - d) * emig(a)
            ex += d * pn + e * pn + n * (part(a_prev) - pn)
            sp_cohorts[R] = n - d - e
        if y in REG_PG_OBS:
            grads = float(REG_PG_OBS[y])
        elif y < 2020:
            grads = {2018: 1400, 2019: 1500}.get(y, 1200.0) * pg_total_coverage
        else:
            sy = y - p.lag_pg
            grads = (pg[sy] * p.fill_pg * p.completion + p.pg_residual) if sy in pg else 0.0
        sp_cohorts[y] = grads
        if y >= 2011:
            spec[y] = sum(n * part(p.age_at_pg + (y - R)) for R, n in sp_cohorts.items())
            sp_grads[y] = grads
            sp_exits[y] = max(0.0, ex)

    need = {y: pop[y] * p.who_norm * p.who_doctor_share / 10000.0 for y in years}
    return dict(
        years=years, population=pop, gov_seats=g, pvt_seats=v,
        total_seats=total_seats, pg_seats=pg, domestic=dom, external=ext, fmg=fm,
        entrants=ent, retirements=retire, deaths=deaths, migration=migr,
        exits={y: retire[y] + deaths[y] + migr[y] for y in years},
        active=active, density={y: active[y] / pop[y] * 10000 for y in years},
        specialists=spec, pg_output=sp_grads, specialist_exits=sp_exits,
        who_need=need, surplus={y: active[y] - need[y] for y in years},
    )


# ---------------------------------------------------------------------------
# Cadre model
# ---------------------------------------------------------------------------
CADRES = [
    ("BSc Nursing", "BSCN", "BSCN 66 Pass out", 4, "Nursing"),
    ("MSc Nursing", "M.Sc.Nursing", "MSc.Nur  30", 2, "Nursing"),
    ("Post Basic BSc Nursing", "Post Basic Nursing ", "Post Basic Nursing - 68", 2, "Nursing"),
    ("BDS", "BDS", "BDS 54 PASSOUT", 5, "Dental"),
    ("MDS", "MDS", "MDS 24 PASS OUT", 3, "Dental"),
    ("B.Pharm", "B.Pharm.", "BPHARM", 4, "Pharmacy"),
    ("M.Pharm", "MPHARM", "MPHARM", 2, "Pharmacy"),
    ("Pharm.D", "PHARMD", "PHARMD", 6, "Pharmacy"),
    ("BPT", "BPT", "BPT 74 PASS out", 5, "Rehabilitation"),
    ("MPT", "MPT", "MPT PASS out", 2, "Rehabilitation"),
    ("BOT", "BOT", "BOT Pass out", 5, "Rehabilitation"),
    ("MOT", "MOT", "MOT Pass out", 2, "Rehabilitation"),
    ("BAMS (Ayurveda)", "BAMS", "bams 64 passout", 6, "AYUSH"),
    ("BSMS (Siddha)", "BSMS", "bsms 60 passout", 6, "AYUSH"),
    ("BUMS (Unani)", "BUMS", "bums 62 passout", 6, "AYUSH"),
    ("BNYS (Naturopathy)", "BNYS", "bnys 82 passout", 6, "AYUSH"),
    ("BHMS (Homoeopathy)", "BHMS", "BHMS 58 PASSOUT", 6, "AYUSH"),
    ("MD Siddha", "MD SIDDHA", "MD SIDDHA 32 PASSOUT", 3, "AYUSH"),
    ("MD Homoeopathy", "MD HOMOEOPATHY", "md homoeo 44 passout", 3, "AYUSH"),
    ("AHS PG (allied)", "AHS PG", "AHS PG PASSOUT", 2, "Allied health"),
]

NOT_COVERED = [
    ("GNM nursing", "Absent from every source file. Sits with the Tamil Nadu Nurses "
     "and Midwives Council, not the university."),
    ("ANM nursing", "Absent from every source file. Same reason as GNM."),
    ("AHS UG (allied health, degree)",
     "Present but internally inconsistent: 10,593 admissions against 3,904 sanctioned "
     "seats for 2025, a fill rate of 271 per cent, and pass-outs swinging from 80 to "
     "10,716 to 508 across four years."),
    ("MD Naturopathy and Yoga, MD Unani",
     "Very small, 65 and 11 seats, with incomplete pass-out sheets."),
    ("Pharm.D Post Baccalaureate", "Only three pass-out years, two clearly partial."),
]

_CADRE_CACHE = None

def load_cadres(end_year: int = 2050, hold_year: int = 2035) -> dict:
    """Build every cadre from the MGRMU workbooks. Cached on first call."""
    global _CADRE_CACHE
    key = (end_year, hold_year)
    if _CADRE_CACHE is not None and _CADRE_CACHE[0] == key:
        return _CADRE_CACHE[1]
    agg = None
    if os.path.exists(CADRE_AGG) and os.path.getsize(CADRE_AGG) > 0:
        import json as _json
        agg = _json.load(open(CADRE_AGG))
    else:
        seats = _read_sheets(SEAT_FILE, _YCOLS, [_SCOLS, _ACOLS])
        passo = _read_sheets(PASS_FILE, _ECOLS, [_PCOLS])
    built = []
    for name, s_sheet, p_sheet, dur, group in CADRES:
        if agg is not None:
            a = agg[name]
            s = {int(y): [a["sanctioned"][y], a["admitted"][y]] for y in a["sanctioned"]}
            p_ = {int(y): v for y, v in a["passouts"].items()}
        else:
            s = seats.get(s_sheet, {})
            p_ = {y: v[0] for y, v in passo.get(p_sheet, {}).items()}
        med = statistics.median(p_.values()) if p_ else 0
        good = {y: v for y, v in p_.items() if med and v >= 0.5 * med}
        dropped = sorted(set(p_) - set(good))
        yrs = sorted(s)
        sanc = {y: s[y][0] for y in yrs}
        adm = {y: s[y][1] for y in yrs}
        fill = (sum(adm.values()) / sum(sanc.values())) if sum(sanc.values()) else 0.0
        ratios = [pv / adm[py - dur] for py, pv in good.items()
                  if (py - dur) in adm and adm[py - dur] > 0]
        completion = float(np.median(ratios)) if ratios else None
        basis = (f"observed ({len(ratios)} cohort match"
                 f"{'es' if len(ratios) != 1 else ''})") if ratios else "assumed"
        built.append(dict(name=name, group=group, dur=dur, sanc=sanc, adm=adm,
                          fill=fill, completion=completion, comp_basis=basis,
                          passouts=p_, usable=good, dropped=dropped))
    obs = [b["completion"] for b in built if b["completion"] is not None]
    default_completion = float(np.median(obs)) if obs else 0.90
    for b in built:
        if b["completion"] is None:
            b["completion"] = default_completion
    out = dict(built=built, default_completion=default_completion,
               n_observed=len(obs),
               first_complete=2021 + max(d for _, _, _, d, _ in CADRES))
    _CADRE_CACHE = (key, out)
    return out


def cadre_seats(b: dict, path: str, end_year: int, hold_year: int = 2035) -> dict:
    """Frozen holds the last observed level. Trend continues the fitted linear
    trend to hold_year then holds flat. Fitted in levels, never as a rate."""
    yrs = sorted(b["sanc"])
    last_y = max(yrs)
    last_v = b["sanc"][last_y]
    out = {}
    if path == "frozen" or len(yrs) < 3:
        slope = 0.0
    else:
        slope = float(np.polyfit(np.array(yrs, float),
                                 [b["sanc"][y] for y in yrs], 1)[0])
    for y in range(2021, end_year + 1):
        if y <= last_y:
            out[y] = float(b["sanc"].get(y, last_v))
        else:
            out[y] = max(0.0, last_v + slope * (min(y, hold_year) - last_y))
    return out


def run_cadres(end_year: int = 2050, path: str = "trend",
               hold_year: int = 2035, fill_override: dict | None = None,
               completion_override: dict | None = None) -> dict:
    """Annual qualified output per cadre. NOT a practising workforce."""
    data = load_cadres(end_year, hold_year)
    first = data["first_complete"]
    results = {}
    for b in data["built"]:
        seats = cadre_seats(b, path, end_year, hold_year)
        fill = (fill_override or {}).get(b["name"], b["fill"])
        comp = (completion_override or {}).get(b["name"], b["completion"])
        out = {}
        for y in range(first, end_year + 1):
            sy = y - b["dur"]
            out[y] = seats[sy] * fill * comp if sy in seats else float("nan")
        results[b["name"]] = dict(output=out, seats=seats, fill=fill,
                                  completion=comp, group=b["group"],
                                  dur=b["dur"], comp_basis=b["comp_basis"],
                                  dropped=b["dropped"], passouts=b["passouts"],
                                  sanc=b["sanc"])
    return dict(results=results, first_complete=first,
                default_completion=data["default_completion"],
                n_observed=data["n_observed"])
