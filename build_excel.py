#!/usr/bin/env python3
"""Build TN_Doctor_Projections_2050.xlsx.

House style: Times New Roman throughout, black only, no fills, no colour.
No titles or notes inside sheets, the sheet tab names them; the cells carry
data only. Historical vs projected is flagged by a 'Basis' column rather than
by shading, so the workbook stays monochrome and prints cleanly.
"""
import io, contextlib
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.chart import LineChart, AreaChart, BarChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.utils import get_column_letter

with contextlib.redirect_stdout(io.StringIO()):
    import tn_supply_model as m
    import tn_cadre_model as c

FONT = "Times New Roman"
BODY = Font(name=FONT, size=10, color="000000")
BOLD = Font(name=FONT, size=10, bold=True, color="000000")
WRAP = Alignment(wrap_text=True, vertical="top")
WRAPC = Alignment(wrap_text=True, vertical="center", horizontal="center")
BOT = Border(bottom=Side(style="thin", color="000000"))
TOP = Border(top=Side(style="thin", color="000000"))

# Monochrome chart styling: black, mid grey, light grey + dash variants.
GREYS = ["000000", "808080", "BFBFBF"]
DASHES = ["solid", "dash", "sysDot"]

wb = openpyxl.Workbook()
wb.remove(wb.active)

YEARS = list(m.YEARS)
N = len(YEARS)
SQ, CN, PP = "Status quo", "Constrained", "Policy push"


def new_sheet(name):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    return ws


def write_table(ws, headers, rows, widths, formats=None, start_row=1):
    """Header row in bold with a rule under it; data rows plain. No fills."""
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=start_row, column=j, value=h)
        c.font = BOLD; c.alignment = WRAPC; c.border = BOT
    ws.row_dimensions[start_row].height = 30
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    for i, row in enumerate(rows, start=start_row + 1):
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=i, column=j)
            c.value = None if (v is None or (isinstance(v, float) and v != v)) else v
            c.font = BODY
            if formats and formats.get(j):
                c.number_format = formats[j]
    return start_row + 1 + len(rows)


def style_chart(ch, ylab, w=19, h=10):
    ch.y_axis.title = ylab
    ch.x_axis.title = "Year"
    ch.width, ch.height = w, h
    ch.style = None
    for axis in (ch.x_axis, ch.y_axis):
        axis.txPr = None
    return ch


def mono_lines(ch):
    for k, s in enumerate(ch.series):
        s.graphicalProperties.line.solidFill = "000000"
        s.graphicalProperties.line.width = 20000
        s.graphicalProperties.line.dashStyle = DASHES[k % len(DASHES)]
        s.marker = Marker(symbol="none")
        s.smooth = False


def mono_areas(ch):
    for k, s in enumerate(ch.series):
        s.graphicalProperties.solidFill = GREYS[k % len(GREYS)]
        s.graphicalProperties.line.solidFill = "000000"
        s.graphicalProperties.line.width = 9000


def basis(y):
    return "Observed" if y <= 2025 else "Projected"


# ---------------------------------------------------------------- Contents
# ------------------------------------------------------------ citations
# Full references, written once and used wherever a source is named.
CITE_LIU = ("Liu JX, Goryakin Y, Maeda A, Bruckner T, Scheffler RM. Global Health "
            "Workforce Labor Market Projections for 2030. Human Resources for Health. "
            "2017;15:11. doi:10.1186/s12960-017-0187-2. First issued as World Bank Policy "
            "Research Working Paper 7790, Washington DC: World Bank; 2016.")
CITE_WHO_GS = ("World Health Organization. Global Strategy on Human Resources for Health: "
               "Workforce 2030. Geneva: WHO; 2016.")
CITE_WHO_HLMA = ("World Health Organization. Health Labour Market Analysis Guidebook. "
                 "Geneva: WHO; 2021.")
CITE_NCP = ("National Commission on Population, Ministry of Health and Family Welfare. "
            "Population Projections for India and States 2011-2036: Report of the "
            "Technical Group on Population Projections. New Delhi: Government of India; "
            "July 2019.")


ws = new_sheet("Contents")
write_table(ws,
    ["Sheet", "Contents"],
    [["HOW TO READ THIS WORKBOOK", ""],
     ["Inputs", "Every parameter in one cell. Change one and the workbook recalculates"],
     ["Assumptions", "Every judgment call, its status, and what needs State confirmation"],
     ["Variables", "Every column defined: formula, input data, source, code location"],
     ["DOCTORS", ""],
     ["Projection", "Main annual series 2011-2050, status quo. Live formulas throughout"],
     ["Scenarios", "The three scenarios compared at 2025, 2030, 2040, 2050"],
     ["Seats", "Government and private MBBS seats, and PG seats, by scenario"],
     ["Entrants", "Annual additions to the medical register, by origin"],
     ["Exits", "Retirements, deaths and migration, each year"],
     ["Specialists", "PG output and the active specialist stock"],
     ["Sensitivity", "Emigration rate; international density benchmarks"],
     ["Age sensitivity", "Age at registration varied from 22 to 26"],
     ["Validation", "Three checks, each labelled by how independent it actually is"],
     ["OTHER CADRES", ""],
     ["Cadre parameters", "Course length, fill rate, completion rate and its basis"],
     ["Cadre seats", "Sanctioned seats by cadre and year, the input to the formulas"],
     ["Cadre output", "Annual qualified output, twenty cadres, all live formulas"],
     ["Cadre by group", "The same totalled into six groups, with chart"],
     ["Nursing", "Nursing cadres in detail, with chart"],
     ["Data quality", "Pass-out years excluded as partial extracts"],
     ["Not covered", "Cadres absent or unusable, and the reason in each case"],
     ["ALL CADRES", ""],
     ["Total output", "Annual qualified output across every cadre including medical"],
     ["REFERENCE", ""],
     ["Parameters", "Every parameter, its value, and its source"],
     ["Methodology", "Method in full, module by module"],
     ["Provenance", "Every figure classified: taken directly, estimated, derived, or assumed"],
     ["Data Sources", "Every source used, with publisher, coverage and link"]],
    [18, 92])

# --------------------------------------------------------------- Variables
ws = new_sheet("Variables")
V = [
 ("Year", "Projection horizon, 2011 to 2050. 2011 is the base year of the official "
  "population series; 2025 is the last observed year.", "n/a", "n/a",
  "tn_supply_model.py line 42"),
 ("Basis", "Flag: Observed to 2025, Projected from 2026. Replaces shading so the "
  "workbook stays monochrome.", "Observed if Year <= 2025", "n/a", "build_excel.py, basis()"),
 ("Population", "Tamil Nadu total population. The denominator for every density "
  "figure and for WHO need.",
  "2011-2025: official NCP values used verbatim. 2026-2050: Pop(t) = Pop(t-1) x "
  "(1 + g), where g is the NCP series' own 2021 to 2025 average annual rate, "
  "0.298 per cent, held constant (Inputs B28).",
  "NCP published table, 15 annual values to 2025",
  CITE_NCP + " The NCP series itself plateaus near 78 million from 2031; "
  "holding its recent rate instead is our choice so that the WHO requirement keeps "
  "rising with the population. Code: build_population()"),
 ("Government MBBS seats", "Sanctioned MBBS seats in government colleges, state "
  "selection-committee basis.",
  "To 2025: observed, used as given. From 2026: a POLICY SCENARIO, not an "
  "extrapolation, because government sets these seats. Three paths ramp from 5,200 to "
  "a target and then hold: frozen at 5,200 (no action at all); 6,000 by 2035 (the "
  "sixteen 100-seat colleges rise to 150); 9,250 by 2045 (all 37 colleges reach the "
  "NMC ceiling of 250, requiring no new college). Ramps round to NMC's 50-seat blocks.",
  "11 annual observations 2015-16 to 2025-26, plus the 2025 college-level seat "
  "distribution",
  "TN Medical Selection Committee UG Data Sheet; MGRMU Seat Count Stat for the "
  "college-level distribution. Code: lines 95-135"),
 ("Private MBBS seats", "Sanctioned MBBS seats in self-financing colleges, same basis.",
  "To 2025: observed. From 2026: logistic S-curve, Seats(t) = K / (1 + exp(-r(t - t0))). "
  "K is the ceiling, set by scenario at 8,000 / 10,000 / 12,000; r and t0 are fitted "
  "by least squares in LEVELS to the 11 observed years. Fitted: K=10,000 gives r=0.204 "
  "and an inflection at 2025.5.",
  "11 annual observations, 2015-16 to 2025-26",
  "TN Medical Selection Committee, UG MBBS BDS Data Sheet. Code: lines 76-93, 120-130"),
 ("Total MBBS seats", "All sanctioned MBBS seats.",
  "Government MBBS seats + Private MBBS seats. No separate estimation.", "n/a",
  "Derived. Code: Projection sheet build, build_excel.py"),
 ("PG seats", "Sanctioned postgraduate medical seats (MD, MS and PG Diploma), full "
  "basis including deemed universities.",
  "2021-2024: MGRMU seats scaled to full basis by 1.407. 2025-26: Selection Committee "
  "actual (5,534). From 2026: linear in levels, +321.5 seats/yr, then capped at 70% of "
  "Total MBBS seats.",
  "5 annual observations, 2021 to 2025-26",
  "MGRMU Seat Count Stat; TN Selection Committee MD and MS seat files. The 1.407 "
  "factor is validated: it reproduces 4,630 against the Committee's independent "
  "4,629. Code: lines 135-155"),
 ("Entrants to register", "Doctors added to the Tamil Nadu medical register each year. "
  "The inflow to the stock.",
  "To 2025: observed registrations (TNMC MBBS + foreign medical graduates for "
  "2020-2025; scraped register counts before that). From 2026: Domestic + External + "
  "FMG, where Domestic = Total MBBS seats six years earlier x 0.998 fill x 0.95 "
  "completion; External = 5,381 minus the deemed-overlap correction (806/yr from "
  "2031); FMG = 1,606 + 170 per year, capped at 3,000.",
  "TNMC registrations 2020-2025; scraped register 1927-2019; seat series for the "
  "projected years",
  "TNMC Registrations workbook; NMC Indian Medical Register (scraped). Fill rate from "
  "MGRMU sanctioned-vs-admitted. Code: lines 160-210, 264-289"),
 ("Total exits", "Doctors leaving active practice each year, measured in "
  "active-practice equivalents, not raw headcount.",
  "Retirements + Deaths + Migration, each weighted by participation at the new age. "
  "Deaths = sum over cohorts of n x mortality(age) x participation(age). Migration = "
  "n x (1-mortality) x emigration(age) x participation(age). Retirements = n x "
  "[participation(age-1) - participation(age)]. The participation weighting is what "
  "makes exits reconcile exactly with the stock.",
  "Register cohorts by year of registration; age-banded mortality; calibrated "
  "emigration; a retirement participation curve",
  "Cohorts from the scraped register. Mortality from SRS-type age patterns adjusted "
  "for the professional class. Emigration is CALIBRATED at 1.2%/yr for ages 25-45, "
  "not observed. Code: lines 233-262, 295-328"),
 ("Net change", "Year-on-year change in the active doctor stock.",
  "Entrants to register - Total exits. This equals the actual change in Active "
  "doctors exactly, verified to under 1 doctor in all 39 year-transitions.", "n/a",
  "Derived. Identity checked in the counting audit"),
 ("Active doctors", "Doctors alive, resident in Tamil Nadu, and still in active "
  "practice. NOT the same as cumulative registrations.",
  "Register cohorts rolled forward from 1927. Each year every cohort is reduced by "
  "mortality and emigration, then multiplied by a participation factor set by its "
  "age. Active = sum over cohorts of surviving headcount x participation(age). Age is "
  "inferred as 24 at registration plus years elapsed.",
  "193,264 scraped registrations by year, 1927-2025, with 2020-2025 taken on the TNMC "
  "basis",
  "NMC Indian Medical Register (scraped) and TNMC Registrations workbook. Age is "
  "INFERRED, not observed. Code: lines 212-292"),
 ("Doctors per 10,000", "Doctor density, the standard workforce comparator.",
  "Active doctors / Population x 10,000.", "n/a", "Derived"),
 ("Active specialists", "Doctors holding a broad-speciality postgraduate "
  "qualification and still in active practice. A SUBSET of Active doctors, never "
  "added to it.",
  "Same cohort machinery as Active doctors, with age 30 at PG completion. Inflow: "
  "2020-2025 observed TNMC broad-speciality registrations; from 2026, PG seats three "
  "years earlier x 0.978 fill x 0.95 completion, plus an external PG term of 1,599/yr. "
  "Superspecialty (DM, M.Ch, DNB SS) is EXCLUDED because those doctors already counted "
  "as specialists at MD or MS.",
  "TNMC PG registrations 2020-2025; MGRMU pass-outs for the back-cast; PG seat series",
  "TNMC Registrations workbook; MGRMU Passout Count Stat. Code: lines 333-389"),
 ("WHO need (doctors)", "The number of doctors the WHO density norm implies for this "
  "population. A floor, not a target.",
  "Population x 44.5 x 0.25 / 10,000, i.e. Population x 11.125 per 10,000. The 44.5 "
  "covers doctors, nurses and midwives together; doctors are one quarter of it on the "
  "1:3 doctor-to-nurse split.",
  "Population series; two published norm constants",
  CITE_WHO_GS + " for 44.5; the same source for the 1:3 doctor to nurse split. "
  "Code: NEED_DOCTORS"),
 ("Surplus vs WHO need", "How far actual supply runs above or below the WHO floor.",
  "Active doctors - WHO need. Positive means supply exceeds the floor.", "n/a",
  "Derived"),
 ("Revised projected need (doctors)", "A need line revised for what the state's own "
  "economy will support: it starts from the doctors Tamil Nadu actually has in 2025 "
  "and rises with income, ageing and the shift away from out-of-pocket payment. The "
  "demand side of the health labour market framework, so a behavioural benchmark "
  "rather than a fixed per-head norm like the WHO floor.",
  "ln(physicians per 1,000) = -9.882 + 0.231 ln GDPpc(t-1) + 0.531 ln GDPpc(t-4) "
  "- 0.518 ln GDPpc(t-5) - 0.099 ln OOPpc(t-2) + 0.516 ln Pop65(t-3) + country effect "
  "(Table 1). The country effect pins the level, so the series is anchored on the "
  "modelled 2025 density and moved each year by the three drivers: density(t) = "
  "density(t-1) x (1+g)^0.244 x (1+g+adj)^-0.099 x (1+p65)^0.516, times population. "
  "Inputs B32 to B38.",
  "Published elasticities; three assumed growth paths (Inputs B35 to B38)",
  CITE_LIU + " Code: build_revised_need()"),
 ("Surplus vs revised need", "How far supply runs above the revised projected need.",
  "Active doctors - revised projected need.", "n/a", "Derived"),
]
end = write_table(ws,
    ["Variable", "What it is", "How it is computed", "Input data",
     "Source and code location"],
    [list(x) for x in V], [22, 40, 66, 34, 52])
for i in range(2, end):
    for j in range(1, 6):
        ws.cell(row=i, column=j).alignment = WRAP
    ws.row_dimensions[i].height = 92
ws.freeze_panes = "B2"

# ------------------------------------------------------------------ Inputs
# Every parameter the workbook computes from lives here in one cell, so a
# reviewer can change one number and watch the whole model move.
ws = new_sheet("Inputs")
_r, _t0 = m.PVT_PARAMS[10000]
_pg_slope, _pg_int = float(m._pg_b[0]), float(m._pg_b[1])
_g0 = (m.NCP[2036] - m.NCP[2035]) / m.NCP[2035]   # NCP tail, shown for reference
INP = [
 ("WHO skilled health worker density norm", 44.5, "per 10,000 population",
  "WHO, Global Strategy on HRH: Workforce 2030 (2016)"),
 ("WHO doctor share of that norm", 0.25, "share",
  "WHO (2016), doctor to nurse ratio 1 to 3"),
 ("MBBS fill rate", m.FILL_MBBS, "share of sanctioned seats",
  "OBSERVED. MGRMU sanctioned against admitted, 2021-2025"),
 ("PG fill rate", m.FILL_PG, "share of sanctioned seats",
  "OBSERVED. MGRMU PG Medical sheet, 2021-2024"),
 ("Completion rate", m.COMPLETION, "share of admissions",
  "NMC benchmark. Validated: PG predicted 2,707 against 2,717 observed"),
 ("MBBS course lag", m.LAG_MBBS, "years",
  "4.5 academic years plus 12-month internship, registration on completion"),
 ("PG course lag", m.LAG_PG, "years", "MD, MS, DM, M.Ch, DNB broad duration"),
 ("External entrants, flat", m.RESIDUAL, "per year",
  "OBSERVED RESIDUAL. Registrations minus modelled domestic output, 2021-2025 mean"),
 ("Seat-source gap, 2025-26 and later", m.GAP_FORWARD, "seats",
  "OBSERVED. Selection Committee 4,750 against MGRMU 3,900 for 2025-26"),
 ("Seat-source gap, seat year 2023", 450, "seats", "OBSERVED. SC 3,850 against MGRMU 3,400"),
 ("Seat-source gap, seat year 2024", 550, "seats", "OBSERVED. SC 4,050 against MGRMU 3,500"),
 ("Deemed overlap removed from residual", None, "graduates per year",
  "DERIVED, see formula in cell B13: gap x fill x completion"),
 ("Deemed overlap first full year", 2031, "year",
  "Seat year 2025 plus the 6-year MBBS lag"),
 ("Foreign medical graduates, 2025 base", m.REG_FMG_OBS[2025], "per year",
  "OBSERVED. TNMC registrations 2025"),
 ("Foreign medical graduates, annual increment", 170, "per year",
  "OBSERVED TREND. TNMC 2020-2025, fitted in levels"),
 ("Foreign medical graduates, ceiling", 3000, "per year",
  "ASSUMPTION. Roughly double the 2025 level, capped as domestic capacity expands"),
 ("Private MBBS ceiling K, status quo", 10000, "seats",
  "EXPERT-JUDGMENT SCENARIO BOUND, not a derived figure. See Assumptions sheet"),
 ("Private MBBS logistic growth r", _r, "per year",
  "FITTED by least squares in levels to 11 observed years, given K"),
 ("Private MBBS logistic midpoint t0", _t0, "years after 2015",
  "FITTED with r. Inflection year = 2015 + t0"),
 ("Government MBBS 2025 level", 5200, "seats", "OBSERVED. Selection Committee 2025-26"),
 ("Government MBBS target, status quo", 6000, "seats",
  "POLICY SCENARIO. The sixteen 100-seat colleges rise to 150"),
 ("Government MBBS target year, status quo", 2035, "year",
  "POLICY SCENARIO. Pending confirmation with the State"),
 ("PG seats linear intercept", _pg_int, "seats at 2021",
  "FITTED in levels to the 2021-2025 full-basis series"),
 ("PG seats linear slope", _pg_slope, "seats per year", "FITTED with the intercept"),
 ("PG seats cap, share of MBBS seats", 0.70, "share",
  "OWN ASSUMPTION, not sourced. See Assumptions sheet"),
 ("PG external entrants, flat", m.PG_RESIDUAL, "per year",
  "OBSERVED RESIDUAL. TNMC PG registrations minus seat-driven output, 2024-2025"),
 ("Population growth rate beyond 2025", None, "per year",
  "DERIVED, see formula in cell B28: NCP 2021 to 2025 average annual rate, held "
  "constant from 2026. The NCP series itself plateaus from 2031"),
 ("Population, last NCP year used verbatim", 2025, "year",
  "The model's last observed year. NCP tail growth for reference: "
  f"{_g0:+.4%} per year at 2036"),
 ("Emigration and out-of-state loss", m.EMIG_RATE, "per year, ages 25-45",
  "CALIBRATED, not observed. See Sensitivity sheet"),
 ("Age at registration", m.AGE_AT_REG, "years",
  "INFERRED, not observed. See Age sensitivity sheet"),
 ("Revised need: income elasticity", m.RN_E_GDP, "elasticity",
  "PUBLISHED. Sum of the three GDP per capita lags in Table 1, 0.231 + 0.531 - 0.518, "
  "of " + CITE_LIU),
 ("Revised need: out-of-pocket elasticity", m.RN_E_OOP, "elasticity",
  "PUBLISHED. Table 1 of Liu et al. 2017 (full reference in row 32)"),
 ("Revised need: population 65+ elasticity", m.RN_E_POP65, "elasticity",
  "PUBLISHED. Table 1 of Liu et al. 2017 (full reference in row 32)"),
 ("Revised need: real GSDP per capita growth, 2026", m.RN_GDP_G_2026, "per year",
  "ASSUMED, FOR STATE CONFIRMATION. Recent Tamil Nadu real GSDP growth less population "
  "growth. See Assumptions sheet"),
 ("Revised need: real GSDP per capita growth, 2050", m.RN_GDP_G_2050, "per year",
  "ASSUMED, FOR STATE CONFIRMATION. The 2026 rate eases linearly to this by 2050"),
 ("Revised need: OOP growth less GDP growth", m.RN_OOP_ADJ, "per year",
  "ASSUMED. Zero holds the out-of-pocket share of spending constant"),
 ("Revised need: growth of population aged 65+", m.RN_POP65_G, "per year",
  "ASSUMED, FOR STATE CONFIRMATION. To be replaced by the NCP 2019 age tables for "
  "Tamil Nadu"),
]
write_table(ws, ["Parameter", "Value", "Unit", "Basis and source"],
            [[a, b, c_, d] for a, b, c_, d in INP], [42, 14, 24, 62])
ws["B13"] = "=B10*B4*B6"         # deemed overlap, computed live
ws["B13"].number_format = "#,##0"
# population growth beyond 2025: the NCP series' own 2021 to 2025 rate, live
ws["B28"] = "=(Projection!C16/Projection!C12)^(1/4)-1"
for i in range(2, 2 + len(INP)):
    ws.cell(row=i, column=4).alignment = WRAP
    ws.row_dimensions[i].height = 26
for i, fmt in [(2, "0.0"), (3, "0.00"), (4, "0.000"), (5, "0.000"), (6, "0.00"),
               (7, "0"), (8, "0"), (9, "#,##0"), (10, "#,##0"), (11, "#,##0"),
               (12, "#,##0"), (13, "#,##0"), (14, "0"), (15, "#,##0"), (16, "#,##0"),
               (17, "#,##0"), (18, "#,##0"), (19, "0.000"), (20, "0.00"), (21, "#,##0"),
               (22, "#,##0"), (23, "0"), (24, "#,##0"), (25, "0.0"), (26, "0.00"),
               (27, "#,##0"), (28, "0.00000"), (29, "0"), (30, "0.000"), (31, "0"),
               (32, "0.000"), (33, "0.000"), (34, "0.000"), (35, "0.000"), (36, "0.000"),
               (37, "0.000"), (38, "0.000")]:
    ws.cell(row=i, column=2).number_format = fmt
IN = "Inputs!"

# -------------------------------------------------------------- Projection
# Live formulas throughout. Only three kinds of cell hold a typed number:
# published source data, observed registrations, and the three cohort-model
# flows (retirements, deaths, migration) which depend on the full age
# distribution and cannot be written as a sheet formula.
ws = new_sheet("Projection")
g, p, pg = m.gov_seats(SQ), m.pvt_seats(SQ), m.pg_seats(SQ)
act, ent = m.col(SQ, "active"), m.col(SQ, "entrants")
spec = [r["active"] for r in m.SPECIALISTS[SQ]]
spex = [r["exits"] for r in m.SPECIALISTS[SQ]]
spgr = [r["grads"] for r in m.SPECIALISTS[SQ]]
tot_sq = g + p

HEAD = ["Year", "Basis", "Population", "Government MBBS seats", "Private MBBS seats",
        "Total MBBS seats", "PG seats", "Domestic MBBS output", "External entrants",
        "Foreign medical graduates", "Entrants to register", "Retirements", "Deaths",
        "Migration", "Total exits", "Net change", "Active doctors",
        "Doctors per 10,000", "PG qualified output", "Specialist exits",
        "Active specialists", "WHO need (doctors)", "Surplus vs WHO need",
        "Revised projected need (doctors)", "Surplus vs revised need"]
write_table(ws, HEAD, [], [7, 10, 12, 13, 12, 12, 10, 12, 11, 13, 12, 11, 9, 11,
                           10, 10, 12, 11, 12, 11, 12, 12, 13, 14, 13])

def RW(y):
    return y - 2009            # 2011 -> row 2

for i, y in enumerate(YEARS):
    r = RW(y)
    ws.cell(row=r, column=1, value=y)
    ws.cell(row=r, column=2, value=basis(y))
    # C population
    if y <= 2025:
        ws.cell(row=r, column=3, value=float(m.POP[i]))
    else:
        ws.cell(row=r, column=3, value=f"=C{r-1}*(1+{IN}$B$28)")
    # D government seats
    if y <= 2025:
        ws.cell(row=r, column=4, value=None if g[i] != g[i] else round(g[i]))
    else:
        ws.cell(row=r, column=4,
                value=f"=IF(A{r}>={IN}$B$23,{IN}$B$22,"
                      f"ROUND(({IN}$B$21+({IN}$B$22-{IN}$B$21)*(A{r}-2025)"
                      f"/({IN}$B$23-2025))/50,0)*50)")
    # E private seats
    if y <= 2025:
        ws.cell(row=r, column=5, value=None if p[i] != p[i] else round(p[i]))
    else:
        ws.cell(row=r, column=5,
                value=f"={IN}$B$18/(1+EXP(-{IN}$B$19*(A{r}-2015-{IN}$B$20)))")
    # F total seats
    if g[i] == g[i]:
        ws.cell(row=r, column=6, value=f"=D{r}+E{r}")
    # G PG seats
    if 2021 <= y <= 2025:
        ws.cell(row=r, column=7, value=round(pg[i]))
    elif y > 2025:
        ws.cell(row=r, column=7,
                value=f"=MIN({IN}$B$24+{IN}$B$25*(A{r}-2021),{IN}$B$26*F{r})")
    # H domestic MBBS output
    if y - m.LAG_MBBS >= 2015:
        ws.cell(row=r, column=8, value=f"=F{RW(y - m.LAG_MBBS)}*{IN}$B$4*{IN}$B$6")
    # I external entrants
    if 2021 <= y <= 2025:
        ws.cell(row=r, column=9, value=float(m.REG_MBBS_OBS[y] - m.domestic_output(tot_sq, y)))
    elif y > 2025:
        ws.cell(row=r, column=9,
                value=f"={IN}$B$9-IF(A{r}<2029,0,IF(A{r}=2029,{IN}$B$11*{IN}$B$4*{IN}$B$6,"
                      f"IF(A{r}=2030,{IN}$B$12*{IN}$B$4*{IN}$B$6,{IN}$B$13)))")
    # J foreign medical graduates
    if y in m.REG_FMG_OBS:
        ws.cell(row=r, column=10, value=m.REG_FMG_OBS[y])
    elif y > 2025:
        ws.cell(row=r, column=10,
                value=f"=MIN({IN}$B$17,{IN}$B$15+{IN}$B$16*(A{r}-2025))")
    # K entrants
    if y >= 2021:
        ws.cell(row=r, column=11, value=f"=H{r}+I{r}+J{r}")
    else:
        ws.cell(row=r, column=11, value=float(ent[i]))
    # L, M, N cohort-model flows, typed values
    ex = m.EXITS[i]
    ws.cell(row=r, column=12, value=float(ex["retirements"]))
    ws.cell(row=r, column=13, value=float(ex["deaths"]))
    ws.cell(row=r, column=14, value=float(ex["emigration"]))
    ws.cell(row=r, column=15, value=f"=SUM(L{r}:N{r})")
    ws.cell(row=r, column=16, value=f"=K{r}-O{r}")
    # Q active doctors
    if y == 2011:
        ws.cell(row=r, column=17, value=float(act[i]))
    else:
        ws.cell(row=r, column=17, value=f"=Q{r-1}+P{r}")
    ws.cell(row=r, column=18, value=f"=Q{r}/C{r}*10000")
    # S PG qualified output
    if y <= 2025:
        ws.cell(row=r, column=19, value=float(spgr[i]))
    else:
        ws.cell(row=r, column=19, value=f"=G{RW(y - m.LAG_PG)}*{IN}$B$5*{IN}$B$6+{IN}$B$27")
    ws.cell(row=r, column=20, value=float(spex[i]))
    if y == 2011:
        ws.cell(row=r, column=21, value=float(spec[i]))
    else:
        ws.cell(row=r, column=21, value=f"=U{r-1}+S{r}-T{r}")
    ws.cell(row=r, column=22, value=f"=C{r}*{IN}$B$2*{IN}$B$3/10000")
    ws.cell(row=r, column=23, value=f"=Q{r}-V{r}")
    # X, Y: revised projected need, anchored on the modelled 2025 stock, then moved by
    # GDP per capita, out-of-pocket spending and the 65+ population with the
    # published elasticities. Density recursion, times population.
    if y == m.RN_ANCHOR:
        ws.cell(row=r, column=24, value=f"=Q{r}")
        ws.cell(row=r, column=25, value=f"=Q{r}-X{r}")
    elif y > m.RN_ANCHOR:
        gg = f"({IN}$B$35+({IN}$B$36-{IN}$B$35)*(A{r}-2026)/24)"
        ws.cell(row=r, column=24,
                value=f"=X{r-1}/C{r-1}*C{r}*(1+{gg})^{IN}$B$32"
                      f"*(1+{gg}+{IN}$B$37)^{IN}$B$33*(1+{IN}$B$38)^{IN}$B$34")
        ws.cell(row=r, column=25, value=f"=Q{r}-X{r}")
    for j in range(1, 26):
        cl = ws.cell(row=r, column=j)
        cl.font = BODY
        if j >= 3:
            cl.number_format = "0.0" if j == 18 else "#,##0"

ws.freeze_panes = "C2"
last = 1 + N

ch = LineChart()
ch.add_data(Reference(ws, min_col=17, min_row=1, max_row=last), titles_from_data=True)
ch.add_data(Reference(ws, min_col=22, min_row=1, max_row=last), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
style_chart(ch, "Doctors"); mono_lines(ch)
ws.add_chart(ch, "Y2")

ch = LineChart()
ch.add_data(Reference(ws, min_col=18, min_row=1, max_row=last), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
style_chart(ch, "Doctors per 10,000"); mono_lines(ch)
ws.add_chart(ch, "Y23")

for k, t in enumerate([
 "Every cell in columns F, H, K, O, P, Q, R, S, U, V, W, X and Y is a live formula. "
 "Change any parameter on the Inputs sheet and this sheet recalculates.",
 "Typed numbers appear only where a formula is not possible: published population "
 "(column C to 2025), observed seats and registrations to 2025,",
 "and the three cohort-model flows in columns L, M and N, which depend on the full "
 "age distribution of the register and cannot be written as a sheet formula.",
 "Active doctors in column Q is the running identity Q(t) = Q(t-1) + entrants - "
 "exits, which the model verifies to under one doctor in all 39 transitions."]):
    ws.cell(row=last + 2 + k, column=1, value=t).font = BODY

# --------------------------------------------------------------- Scenarios
ws = new_sheet("Scenarios")
rows = []
for scen in (CN, SQ, PP):
    gg, pp_, gpg = m.gov_seats(scen), m.pvt_seats(scen), m.pg_seats(scen)
    aa, ee = m.col(scen, "active"), m.col(scen, "entrants")
    ss = [x["active"] for x in m.SPECIALISTS[scen]]
    for lab, arr, dec in [("Government MBBS seats", gg, 0),
                          ("Private MBBS seats", pp_, 0),
                          ("Total MBBS seats", gg + pp_, 0),
                          ("PG seats", gpg, 0),
                          ("Entrants to register", ee, 0),
                          ("Active doctors", aa, 0),
                          ("Doctors per 10,000", aa / m.POP * 10000, 1),
                          ("Active specialists", ss, 0)]:
        rows.append([scen, lab] + [round(float(arr[m.IDX[y]]), dec)
                                   for y in (2025, 2030, 2040, 2050)])
write_table(ws, ["Scenario", "Metric", "2025", "2030", "2040", "2050"], rows,
            [14, 24, 13, 13, 13, 13],
            {3: "#,##0.#", 4: "#,##0.#", 5: "#,##0.#", 6: "#,##0.#"})

r0 = 2 + len(rows) + 1
sub = [[s, round(float(m.col(s, "active")[m.IDX[2050]]))] for s in (CN, SQ, PP)]
write_table(ws, ["Scenario", "Active doctors 2050"], sub, [14, 24],
            {2: "#,##0"}, start_row=r0)
ch = BarChart(); ch.type = "col"
ch.add_data(Reference(ws, min_col=2, min_row=r0, max_row=r0 + 3), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=r0 + 1, max_row=r0 + 3))
style_chart(ch, "Doctors", w=13, h=8)
ch.series[0].graphicalProperties.solidFill = "808080"
ch.series[0].graphicalProperties.line.solidFill = "000000"
ws.add_chart(ch, "H2")

# ------------------------------------------------------------------- Seats
ws = new_sheet("Seats")
rows = []
for i, y in enumerate(YEARS):
    def v(a):
        return None if a[i] != a[i] else round(float(a[i]))
    rows.append([y, basis(y), v(m.gov_seats(SQ)), v(m.gov_seats(PP)),
                 v(m.pvt_seats(CN)), v(m.pvt_seats(SQ)), v(m.pvt_seats(PP)),
                 v(m.pg_seats(SQ))])
write_table(ws,
    ["Year", "Basis", "Government MBBS (baseline)", "Government MBBS (policy push)",
     "Private MBBS (constrained)", "Private MBBS (status quo)",
     "Private MBBS (policy push)", "PG seats (status quo)"],
    rows, [7, 10, 15, 15, 15, 15, 15, 14],
    {c: "#,##0" for c in range(3, 9)})
ws.freeze_panes = "C2"
last = 1 + N
ch = LineChart()
for cidx in (3, 6, 8):
    ch.add_data(Reference(ws, min_col=cidx, min_row=1, max_row=last), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
style_chart(ch, "Sanctioned seats"); mono_lines(ch)
ws.add_chart(ch, "J2")

# ---------------------------------------------------------------- Entrants
ws = new_sheet("Entrants")
tot_sq = m.gov_seats(SQ) + m.pvt_seats(SQ)
rows = []
for y in YEARS:
    if y < 2021:
        continue
    dom = m.domestic_output(tot_sq, y)
    ext = (m.REG_MBBS_OBS[y] - dom) if y <= 2025 else m.RESIDUAL - m.deemed_overlap(y)
    fm = m.fmg(y)
    rows.append([y, basis(y), round(dom), round(ext), round(fm),
                 round(dom + ext + fm)])
write_table(ws,
    ["Year", "Basis", "Tamil Nadu state-counselling graduates",
     "Deemed universities and out-of-state", "Foreign medical graduates",
     "Total entrants"],
    rows, [7, 10, 22, 22, 20, 13], {c: "#,##0" for c in range(3, 7)})
last = 1 + len(rows)
ch = AreaChart(); ch.grouping = "stacked"; ch.overlap = 100
ch.add_data(Reference(ws, min_col=3, min_row=1, max_col=5, max_row=last),
            titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
style_chart(ch, "New registrations"); mono_areas(ch)
ws.add_chart(ch, "H2")

# ------------------------------------------------------------------- Exits
ws = new_sheet("Exits")
act_sq, ent_sq = m.col(SQ, "active"), m.col(SQ, "entrants")
rows = []
for i, y in enumerate(YEARS):
    if y < 2015:
        continue
    ex = m.EXITS[i]
    tot = ex["retirements"] + ex["deaths"] + ex["emigration"]
    rows.append([y, basis(y), round(ex["retirements"]), round(ex["deaths"]),
                 round(ex["emigration"]), round(tot), round(ent_sq[i]),
                 round(ent_sq[i] - tot), round(tot / act_sq[i] * 100, 2),
                 round(ent_sq[i] / tot, 1)])
write_table(ws,
    ["Year", "Basis", "Retirements", "Deaths", "Migration and out-of-state",
     "Total exits", "Entrants", "Net change", "Exits as % of stock",
     "Entrants per exit"],
    rows, [7, 10, 11, 9, 18, 10, 10, 10, 13, 13],
    {3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "#,##0", 7: "#,##0", 8: "#,##0",
     9: "0.00", 10: "0.0"})
last = 1 + len(rows)
ch = AreaChart(); ch.grouping = "stacked"; ch.overlap = 100
ch.add_data(Reference(ws, min_col=3, min_row=1, max_col=5, max_row=last),
            titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
style_chart(ch, "Exits per year"); mono_areas(ch)
ws.add_chart(ch, "L2")

r0 = last + 2
_r = sum(m.EXITS[i]["retirements"] for i, y in enumerate(YEARS) if y >= 2026)
_d = sum(m.EXITS[i]["deaths"] for i, y in enumerate(YEARS) if y >= 2026)
_mg = sum(m.EXITS[i]["emigration"] for i, y in enumerate(YEARS) if y >= 2026)
_e = sum(ent_sq[i] for i, y in enumerate(YEARS) if y >= 2026)
tote = _r + _d + _mg
write_table(ws, ["Cumulative 2026-2050", "Total", "Share of exits"],
            [["Retirements", round(_r), _r / tote],
             ["Deaths", round(_d), _d / tote],
             ["Migration and out-of-state", round(_mg), _mg / tote],
             ["Total exits", round(tote), 1.0],
             ["Entrants", round(_e), None],
             ["Net addition", round(_e - tote), None]],
            [26, 12, 13], {2: "#,##0", 3: "0.0%"}, start_row=r0)

# ------------------------------------------------------------- Specialists
ws = new_sheet("Specialists")
rows = []
for i, y in enumerate(YEARS):
    sp = m.SPECIALISTS[SQ][i]
    pgv = m.pg_seats(SQ)[i]
    rows.append([y, basis(y), None if pgv != pgv else round(float(pgv)),
                 round(sp["grads"]), round(sp["active"]), round(act_sq[i]),
                 round(sp["active"] / act_sq[i], 4)])
write_table(ws,
    ["Year", "Basis", "PG seats", "PG graduates", "Active specialists",
     "Active doctors", "Specialist share"],
    rows, [7, 10, 11, 12, 14, 13, 13],
    {3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "#,##0", 7: "0.0%"})
ws.freeze_panes = "C2"
last = 1 + N
ch = LineChart()
ch.add_data(Reference(ws, min_col=5, min_row=1, max_row=last), titles_from_data=True)
ch.add_data(Reference(ws, min_col=6, min_row=1, max_row=last), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
style_chart(ch, "Doctors"); mono_lines(ch)
ws.add_chart(ch, "I2")

# ------------------------------------------------------------- Sensitivity
ws = new_sheet("Sensitivity")
_base = m.EMIG_RATE
rows = []
for rate in (0.006, 0.012, 0.020, 0.030):
    m.EMIG_RATE = rate
    rr = m.run_stock(SQ, m.REGISTER)
    a25, a50 = rr[m.IDX[2025]]["active"], rr[m.IDX[2050]]["active"]
    rows.append([rate, 1 - (1 - rate) ** 21, round(a25), round(a50),
                 round(a50 / m.POP[m.IDX[2050]] * 10000, 1),
                 "Central case" if rate == 0.012 else ""])
m.EMIG_RATE = _base
write_table(ws,
    ["Emigration rate, ages 25-45", "Cumulative loss over 21 years",
     "Active doctors 2025", "Active doctors 2050", "Doctors per 10,000 in 2050",
     "Note"],
    rows, [22, 22, 16, 16, 18, 14],
    {1: "0.0%", 2: "0.0%", 3: "#,##0", 4: "#,##0", 5: "0.0"})

r0 = 2 + len(rows) + 1
write_table(ws, ["Comparator", "Doctors per 10,000"],
            [["India, all states, active, approximate", "7 to 9"],
             ["Tamil Nadu, modelled 2025", "19.6"],
             ["United Kingdom", "32"],
             ["OECD average", "35"],
             ["Tamil Nadu, modelled 2050, status quo", "68.7"],
             ["Cuba, world's highest", "84"],
             ["Note", "International figures are indicative orders of "
                      "magnitude for context, not sourced point estimates."]],
            [40, 60], start_row=r0)

# -------------------------------------------------------------- Validation
ws = new_sheet("Validation")
rows = []
for y in (2020, 2021, 2022, 2023, 2024):
    a = m.REGISTER.get(y, 0); b = m.REG_MBBS_OBS[y] + m.REG_FMG_OBS[y]
    rows.append(["1. Register cross-check", y, a, b, abs(a - b) / b,
                 "INDEPENDENT"])
rows.append(["2. PG pipeline", 2024, 2707, 2717, abs(2707 - 2717) / 2717,
             "HELD-OUT TARGET, SHARED PUBLISHER"])
rows.append(["3. PG coverage factor", 2024, 4630, 4629, abs(4630 - 4629) / 4629,
             "NOT INDEPENDENT, INTERNAL CONSISTENCY ONLY"])
end_r = write_table(ws, ["Check", "Year", "Model or scraped value",
                         "Comparison value", "Difference", "Independence"],
                    rows, [24, 8, 20, 18, 12, 40],
                    {3: "#,##0", 4: "#,##0", 5: "0.00%"})
for i in range(2, end_r):
    ws.cell(row=i, column=6).alignment = WRAP
    ws.row_dimensions[i].height = 26

r0 = end_r + 2
rows2 = [
 ["1. Register cross-check", "INDEPENDENT",
  "The scraped National Medical Commission register and the Tamil Nadu Medical "
  "Council's own workbook are separate systems maintained by different bodies. "
  "Neither was used to build the other. This is the only fully independent check of "
  "the three, and it is the strongest evidence that the register inflow is right."],
 ["2. PG pipeline", "HELD-OUT TARGET, SHARED PUBLISHER",
  "Predicted pass-outs are built from seats, the fill rate and the completion rate. "
  "The pass-out figures they are compared against were NOT used to fit any of those "
  "three, so the target is genuinely held out. However, both the seat file and the "
  "pass-out file come from the same publisher, Dr. M.G.R. Medical University, so a "
  "systematic error in that university's reporting would not be caught. Treat this "
  "as a real but partial test."],
 ["3. PG coverage factor", "NOT INDEPENDENT, INTERNAL CONSISTENCY ONLY",
  "The 1.407 factor was itself derived by comparing the MGRMU basis with the "
  "Selection Committee total, and it is then checked against the Selection Committee "
  "total. That is circular. It confirms the scaling was applied correctly and that "
  "the two sources reconcile arithmetically, which is worth knowing, but it is not "
  "evidence that the model predicts anything. It should not be presented as "
  "validation."],
]
end2 = write_table(ws, ["Check", "Independence", "What it does and does not establish"],
                   rows2, [24, 40, 96], start_row=r0)
for i in range(r0 + 1, end2):
    for j in (2, 3):
        ws.cell(row=i, column=j).alignment = WRAP
    ws.row_dimensions[i].height = 76

r1 = end2 + 1
rows3 = [[2021 + k, round(v)] for k, v in enumerate(m._resid)]
rows3.append(["Mean, held flat in projection", round(m.RESIDUAL)])
end3 = write_table(ws, ["External-entrant residual, year", "Value"], rows3,
                   [30, 24], {2: "#,##0"}, start_row=r1)
for k, t in enumerate([
 "A fourth check is internal rather than external: the stock identity. Active "
 "doctors in year t must equal active doctors in t-1 plus entrants minus exits.",
 "The model verifies this to under one doctor across all 39 year transitions, and "
 "the Projection sheet reproduces it as a live formula in column Q.",
 "This catches accounting errors, which it did: it is how the participation-weighting "
 "error in the exit calculation was found. It cannot catch a wrong assumption."]):
    ws.cell(row=end3 + 1 + k, column=1, value=t).font = BODY

# ------------------------------------------------------------ OTHER CADRES
CY = [y for y in c.YEARS if y >= c.FIRST_COMPLETE]
CSEAT_Y = list(range(2021, 2051))
CNAMES = [b["name"] for b in c.BUILT]
CQ = {b["name"]: c.qualified_output(b, "trend") for b in c.BUILT}
CGROUPS = sorted(set(c.GROUP.values()))
CPAR = "'Cadre parameters'!"
CSEAT = "'Cadre seats'!"
CROW = {b["name"]: i + 2 for i, b in enumerate(c.BUILT)}          # row on Cadre parameters
CCOL = {b["name"]: get_column_letter(i + 2) for i, b in enumerate(c.BUILT)}      # Cadre output / by group / nursing
CCOL_SEAT = {b["name"]: get_column_letter(i + 3) for i, b in enumerate(c.BUILT)} # Cadre seats has an extra Basis column

# Cadre parameters first, because the output formulas reference its rates.
ws = new_sheet("Cadre parameters")
rows = [[b["name"], c.GROUP[b["name"]], b["dur"], b["fill"], b["completion"],
         b["comp_basis"],
         round(c.project_seats(b, "frozen")[2050]), round(c.project_seats(b, "trend")[2050])]
        for b in c.BUILT]
end_r = write_table(ws, ["Cadre", "Group", "Course length, years", "Fill rate",
                         "Completion rate", "Completion basis",
                         "Sanctioned seats 2050, frozen path",
                         "Sanctioned seats 2050, trend path"],
                    rows, [24, 16, 14, 12, 14, 30, 18, 18],
                    {4: "0%", 5: "0%", 7: "#,##0", 8: "#,##0"})
for k, t in enumerate([
 "Fill rate is admissions divided by sanctioned seats, pooled across all observed "
 "years for that cadre. Estimated, not assumed.",
 "Completion rate is pass-outs divided by the admissions of the cohort that entered "
 "one course length earlier. Marked observed where a",
 "cohort could be matched inside the observed window, assumed otherwise. Assumed "
 f"cases take {c.DEFAULT_COMPLETION:.0%}, the median of the "
 f"{len(c._obs)} observed values.",
 "Columns D and E are the live inputs to every formula on the Cadre output sheet."]):
    ws.cell(row=end_r + 1 + k, column=1, value=t).font = BODY

# Annual sanctioned seats, the input the output formulas draw on.
ws = new_sheet("Cadre seats")
write_table(ws, ["Year", "Basis"] + CNAMES, [], [7, 15] + [13] * len(CNAMES))
for y in CSEAT_Y:
    r = y - 2019
    ws.cell(row=r, column=1, value=y).font = BODY
    ws.cell(row=r, column=2, value="Observed" if y <= 2025 else "Projected, trend path").font = BODY
    for b in c.BUILT:
        st = c.project_seats(b, "trend")
        cl = ws.cell(row=r, column=CNAMES.index(b["name"]) + 3, value=float(st[y]))
        cl.font = BODY; cl.number_format = "#,##0"
ws.freeze_panes = "C2"
for k, t in enumerate([
 "Sanctioned seats by cadre and year, trend path. Observed to 2025, then the fitted "
 "linear trend continued to 2035 and held flat.",
 "Growth is fitted in LEVELS, never as a compound rate. The frozen-path alternative "
 "is on the Cadre parameters sheet.",
 "This sheet is the seat input that every formula on the Cadre output sheet reads."]):
    ws.cell(row=len(CSEAT_Y) + 3 + k, column=1, value=t).font = BODY

# Qualified output, entirely by formula.
ws = new_sheet("Cadre output")
write_table(ws, ["Year"] + CNAMES + ["Total, these cadres"], [],
            [7] + [13] * len(CNAMES) + [16])
for y in CY:
    r = y - 2025
    ws.cell(row=r, column=1, value=y).font = BODY
    for b in c.BUILT:
        col = CNAMES.index(b["name"]) + 2
        seat_row = (y - b["dur"]) - 2019
        cl = ws.cell(row=r, column=col,
                     value=f"={CSEAT}{CCOL_SEAT[b['name']]}{seat_row}"
                           f"*{CPAR}$D${CROW[b['name']]}*{CPAR}$E${CROW[b['name']]}")
        cl.font = BODY; cl.number_format = "#,##0"
    lastc = get_column_letter(len(CNAMES) + 1)
    cl = ws.cell(row=r, column=len(CNAMES) + 2, value=f"=SUM(B{r}:{lastc}{r})")
    cl.font = BODY; cl.number_format = "#,##0"
ws.freeze_panes = "B2"
_e = len(CY) + 3
for k, t in enumerate([
 "WHAT THIS SHEET IS. The number of people qualifying in each cadre in each year. "
 "It is NOT a count of people practising.",
 "HOW EACH CELL IS BUILT. Sanctioned seats one course length earlier, taken from the "
 "Cadre seats sheet, multiplied by that cadre's",
 "fill rate and completion rate from the Cadre parameters sheet. Every cell is a live "
 "formula, so changing a rate moves the whole column.",
 "WHY IT STARTS AT 2027. Seat data begins in 2021 and the longest course here is six "
 "years, so 2021 to 2026 are only partly covered",
 "and would understate the total. 2027 is the first year in which every cadre has a "
 "source cohort.",
 "WHAT IS MISSING. GNM and ANM nursing appear in no source file, so nursing is degree "
 "level only. See the Not covered sheet."]):
    ws.cell(row=_e + k, column=1, value=t).font = BODY

ws = new_sheet("Cadre by group")
COUT = "'Cadre output'!"
write_table(ws, ["Year"] + CGROUPS + ["Total"], [], [7] + [15] * len(CGROUPS) + [14])
for y in CY:
    r = y - 2025
    ws.cell(row=r, column=1, value=y).font = BODY
    for gi, gname in enumerate(CGROUPS):
        cols = [CCOL[b["name"]] for b in c.BUILT if c.GROUP[b["name"]] == gname]
        expr = "+".join(f"{COUT}{cc}{r}" for cc in cols)
        cl = ws.cell(row=r, column=gi + 2, value=f"={expr}")
        cl.font = BODY; cl.number_format = "#,##0"
    lastc = get_column_letter(len(CGROUPS) + 1)
    cl = ws.cell(row=r, column=len(CGROUPS) + 2, value=f"=SUM(B{r}:{lastc}{r})")
    cl.font = BODY; cl.number_format = "#,##0"
lastrow = len(CY) + 1
ch = LineChart()
ch.add_data(Reference(ws, min_col=2, max_col=1 + len(CGROUPS), min_row=1, max_row=lastrow),
            titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=lastrow))
style_chart(ch, "Qualified output per year"); mono_lines(ch)
ws.add_chart(ch, f"{get_column_letter(len(CGROUPS) + 4)}2")
for k, t in enumerate([
 "WHAT THIS SHEET IS. The Cadre output sheet totalled into six groups. Each cell "
 "simply adds that group's cadre columns for the same year.",
 "Every cell is a live formula pointing at the Cadre output sheet, so the two sheets "
 "can never disagree.",
 "These are annual qualifications, not practising staff, and must not be added to the "
 "doctor stock on the Projection sheet.",
 "GROUP MEMBERSHIP. " + "; ".join(
     f"{gname}: " + ", ".join(b["name"] for b in c.BUILT if c.GROUP[b["name"]] == gname)
     for gname in CGROUPS)]):
    ws.cell(row=lastrow + 2 + k, column=1, value=t).font = BODY

ws = new_sheet("Nursing")
nn = [b["name"] for b in c.BUILT if c.GROUP[b["name"]] == "Nursing"]
write_table(ws, ["Year"] + nn + ["Total degree level nursing"], [],
            [7] + [18] * len(nn) + [22])
for y in CY:
    r = y - 2025
    ws.cell(row=r, column=1, value=y).font = BODY
    for k2, nm in enumerate(nn):
        cl = ws.cell(row=r, column=k2 + 2, value=f"={COUT}{CCOL[nm]}{r}")
        cl.font = BODY; cl.number_format = "#,##0"
    lastc = get_column_letter(len(nn) + 1)
    cl = ws.cell(row=r, column=len(nn) + 2, value=f"=SUM(B{r}:{lastc}{r})")
    cl.font = BODY; cl.number_format = "#,##0"
lastrow = len(CY) + 1
ch = LineChart()
ch.add_data(Reference(ws, min_col=2, max_col=1 + len(nn), min_row=1, max_row=lastrow),
            titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=lastrow))
style_chart(ch, "Qualified output per year"); mono_lines(ch)
ws.add_chart(ch, f"{get_column_letter(len(nn) + 4)}2")
for k, t in enumerate([
 "Degree level nursing only. GNM and ANM are absent from every source file because "
 "they sit with the Tamil Nadu Nurses and Midwives",
 "Council rather than the university, so these figures understate total nursing "
 "output considerably.",
 "Note the fill rates behind these columns: BSc Nursing fills 97 per cent of its "
 "seats, but MSc Nursing 59 per cent and Post Basic 48 per cent."]):
    ws.cell(row=lastrow + 2 + k, column=1, value=t).font = BODY

ws = new_sheet("Data quality")
rows = []
for b in c.BUILT:
    for y in b["dropped"]:
        rows.append([b["name"], y, b["passouts"][y],
                     ", ".join(f"{v:,}" for v in sorted(b["usable"].values())),
                     "Excluded as a partial extract: below half the median of the cadre's "
                     "own pass-out series."])
end_r = write_table(ws, ["Cadre", "Year excluded", "Reported pass-outs", "Years retained",
                         "Reason"], rows, [24, 13, 16, 30, 54], {3: "#,##0"})
for i in range(2, end_r):
    for j in (4, 5):
        ws.cell(row=i, column=j).alignment = WRAP
    ws.row_dimensions[i].height = 28

ws = new_sheet("Not covered")
end_r = write_table(ws, ["Cadre or course", "Why it is not projected"],
                    [[k, v] for k, v in c.EXCLUDED], [30, 96])
for i in range(2, end_r):
    ws.cell(row=i, column=2).alignment = WRAP
    ws.row_dimensions[i].height = 48

# --------------------------------- TOTAL OUTPUT, ALL CADRES INCLUDING MEDICAL
ws = new_sheet("Total output")
CBG = "'Cadre by group'!"
write_table(ws, ["Year", "MBBS (doctors)", "PG medical"] + CGROUPS
            + ["Total qualified output, all cadres"], [],
            [7, 15, 13] + [14] * len(CGROUPS) + [22])
for y in CY:
    r = y - 2025
    ws.cell(row=r, column=1, value=y).font = BODY
    ws.cell(row=r, column=2, value=f"=Projection!H{y - 2009}").font = BODY
    ws.cell(row=r, column=3, value=f"=Projection!S{y - 2009}").font = BODY
    for gi in range(len(CGROUPS)):
        ws.cell(row=r, column=gi + 4,
                value=f"={CBG}{get_column_letter(gi + 2)}{r}").font = BODY
    lastc = get_column_letter(len(CGROUPS) + 3)
    ws.cell(row=r, column=len(CGROUPS) + 4, value=f"=SUM(B{r}:{lastc}{r})").font = BODY
    for j in range(2, len(CGROUPS) + 5):
        ws.cell(row=r, column=j).number_format = "#,##0"
lastrow = len(CY) + 1
ch = LineChart()
ch.add_data(Reference(ws, min_col=len(CGROUPS) + 4, min_row=1, max_row=lastrow),
            titles_from_data=True)
ch.add_data(Reference(ws, min_col=2, min_row=1, max_row=lastrow), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=lastrow))
style_chart(ch, "Qualified output per year"); mono_lines(ch)
ws.add_chart(ch, f"{get_column_letter(len(CGROUPS) + 6)}2")
for k, t in enumerate([
 "Every column here is ANNUAL QUALIFIED OUTPUT, that is people finishing a course, on "
 "a Tamil Nadu seat basis. Every cell is a live formula.",
 "It is NOT a practising workforce count and must not be added to the doctor stock on "
 "the Projection sheet.",
 "The MBBS column is domestic output from Tamil Nadu seats only. The register also "
 "takes in deemed university, out of state and foreign",
 "medical graduates, which is why entrants on the Projection sheet run at roughly "
 "double this column.",
 "A practising stock total across all cadres is not possible from the available data, "
 "because only doctors have a register.",
 "Nursing is degree level only. GNM and ANM are absent from every source file, so "
 "nursing output is understated here."]):
    ws.cell(row=lastrow + 2 + k, column=1, value=t).font = BODY

# -------------------------------------------------------------- Assumptions
# Every judgment call, separated from the observed and fitted parameters, with
# its status stated plainly. Items marked FOR STATE CONFIRMATION are to be put
# to the State in the review meeting.
ws = new_sheet("Assumptions")
A = [
 ("Deemed overlap, 806 graduates per year from 2031", "DERIVED, arithmetic shown",
  "The Selection Committee and MGRMU private MBBS series agree exactly for 2021-22 "
  "and 2022-23, then diverge by 450 seats (2023-24), 550 (2024-25) and 850 (2025-26). "
  "The gap is most likely deemed universities entering the Selection Committee sheet. "
  "Deemed output already sits inside the external-entrant residual, which was "
  "estimated from seat years when the two sources still agreed, so without a "
  "correction those graduates would be counted twice. The arithmetic is: "
  "850 seats x 0.998 fill x 0.95 completion = 806 graduates per year. The 6-year MBBS "
  "lag makes seat year 2025 register in 2031, which is why the full correction starts "
  "then; 2029 and 2030 take the smaller 450 and 550 gaps. Computed live in cell B13 "
  "of the Inputs sheet.", "None. Arithmetic follows from two observed series."),
 ("Private MBBS ceilings: 8,000 / 10,000 / 12,000 seats",
  "EXPERT-JUDGMENT SCENARIO BOUNDS, not derived",
  "These are NOT calculated from faculty norms, bed capacity or the applicant pool. "
  "No arithmetic connects those inputs to these three numbers. They are round-number "
  "bounds chosen to span a plausible range around the 2025 level of 4,750, and the "
  "considerations above are the reasoning for the range being finite rather than the "
  "source of the specific values. They are presented as three scenarios precisely "
  "because the ceiling is not knowable from the data. What the data does support is "
  "that growth is bounded rather than compounding: extrapolating the observed "
  "compound rate gives 183,707 seats by 2050, which is not credible.",
  "FOR STATE CONFIRMATION. The State is better placed to say what private expansion "
  "it will approve. A derived ceiling would need NMC faculty returns and hospital bed "
  "data per college, which we do not hold."),
 ("PG seats capped at 70 per cent of MBBS seats", "OWN ASSUMPTION, not sourced",
  "There is no citation for 70 per cent. It is our own figure, not taken from the "
  "Andhra Pradesh model or any published norm. The reasoning is structural: PG intake "
  "cannot indefinitely exceed the MBBS pipeline that feeds it. For scale, the ratio "
  "was 56 per cent in 2025-26 (5,534 PG against 9,950 MBBS) and national policy has "
  "moved toward parity, so 70 per cent sits between the current position and a 1:1 "
  "target. The cap binds only after 2040 in the central case, so it moves the near-"
  "term projection very little.",
  "FOR STATE CONFIRMATION. Replace with the State's own PG expansion intent if one "
  "exists."),
 ("Government seat paths, including policy push",
  "POLICY SCENARIOS, not extrapolations",
  "Government sets these seats, so these are scenarios about decisions. All three are "
  "built from the observed college-level distribution: 37 colleges, 5,200 seats, mean "
  "141, with sixteen at 100 seats, sixteen at 150, one at 200 and four at the NMC "
  "ceiling of 250. Frozen holds 5,200, a no-action counterfactual. Consolidation "
  "reaches 6,000 by 2035, which is the sixteen 100-seat colleges rising to 150. "
  "Policy push reaches 9,250 by 2045, which is every college reaching the NMC ceiling "
  "of 250 and requires no new college. Ramps are linear and rounded to NMC's 50-seat "
  "approval blocks.",
  "FOR STATE CONFIRMATION. The target levels, the years by which they are reached, "
  "and whether new colleges are planned in addition to intake increases, all need the "
  "State's input. The policy push path in particular is our construction and should "
  "be replaced by the State's actual capacity plan."),
 ("Age at registration, 24 years", "INFERRED, not observed",
  "MBBS entry at 18 plus 5.5 years to qualification. Age is never recorded in the "
  "scraped register, only year of registration, so every age in the model is inferred "
  "from it. Retirement timing is the output most sensitive to this. See the Age "
  "sensitivity sheet: varying the assumption from 22 to 26 moves the 2050 stock by "
  "under 3 per cent but moves cumulative retirements to 2050 by about 18 per cent.",
  "Resolvable with data. The NMC detail records carry birth dates; a stratified "
  "sample of 5,000 to 8,000 records would replace the assumption with an estimate."),
 ("Emigration and out-of-state loss, 1.2 per cent per year, ages 25 to 45",
  "CALIBRATED, not observed",
  "No direct measure of Tamil Nadu doctor out-migration exists. The rate implies "
  "about 22 per cent cumulative loss across the exposed ages. It is the model's most "
  "consequential single parameter.",
  "Resolvable with data. The NMC removedStatus field records removals from the "
  "register and would allow this to be estimated rather than calibrated."),
 ("Revised projected need: growth paths of its three drivers",
  "ASSUMED, FOR STATE CONFIRMATION",
  "The elasticities are published (" + CITE_LIU + ", Table 1): 0.244 to real GDP "
  "per capita in the long run, -0.099 to out-of-pocket spending per capita, 0.516 to "
  "the population aged 65 and over. What is assumed is the path of the drivers: real "
  "GSDP per capita growth of 6.5 per cent in 2026 easing to 4.0 per cent by 2050, "
  "out-of-pocket spending growing with GDP (share constant), and the 65+ population "
  "growing 3.5 per cent a year. The State's own GSDP forecast and the NCP 2019 age "
  "tables should replace these. The line is anchored on the modelled 2025 density, "
  "which is how the paper's country fixed effect works.",
  "Moderate. Demand reaches 305,668 by 2050 on these paths; supply is 517,657. A "
  "faster economy raises the line, but with an elasticity of 0.244 it would take "
  "implausible growth to close the gap."),
 ("Population growth beyond 2025, 0.30 per cent a year held constant",
  "OWN CHOICE, RATE FROM NCP",
  "The NCP series is used verbatim to 2025. Its 2021 to 2025 average rate is then "
  "held constant to 2050, so the population reaches 83.3 million and the WHO "
  "requirement rises with it. The NCP series itself plateaus near 78 million from "
  "2031; that path gives 76.0 million in 2050 and is kept as an alternative in the "
  "dashboard.",
  "Moderate. Density in 2050 is 62.2 per 10,000 on this path against 68.1 on the "
  "NCP plateau; the surplus finding holds on both."),
 ("Completion rate, 95 per cent for medical courses", "BENCHMARK, validated",
  "Taken as the NMC benchmark and also used in the Andhra Pradesh model. It is not "
  "estimated from Tamil Nadu data directly for MBBS because no admission cohort in "
  "the observed window can be matched to its own pass-out year at a 6-year lag. It "
  "was validated indirectly on PG, where the match is possible: predicted 2,707 "
  "against 2,717 observed.", "None outstanding."),
 ("Cadre seat growth: frozen, or trend continued to 2035 then held",
  "OWN ASSUMPTION", "With four or five annual observations per cadre a fitted "
  "ceiling cannot be supported, so two bounding paths are given instead. Holding flat "
  "after 2035 reflects that capacity cannot expand indefinitely against a falling "
  "state population.",
  "FOR STATE CONFIRMATION alongside the medical seat paths."),
]
end_r = write_table(ws, ["Assumption", "Status", "Basis, and the arithmetic where there is any",
                         "Outstanding action"],
                    [list(x) for x in A], [34, 30, 86, 44])
for i in range(2, end_r):
    for j in range(1, 5):
        ws.cell(row=i, column=j).alignment = WRAP
    ws.row_dimensions[i].height = 128

# ---------------------------------------------------------- Age sensitivity
ws = new_sheet("Age sensitivity")
_base_age = m.AGE_AT_REG
rows = []
for _age in (22, 23, 24, 25, 26):
    m.AGE_AT_REG = _age
    rr = m.run_stock(SQ, m.REGISTER)
    ee = m.exit_decomposition(SQ, m.REGISTER)
    cum = sum(ee[i]["retirements"] for i, y in enumerate(YEARS) if y >= 2026)
    rows.append([_age, round(rr[m.IDX[2025]]["active"]), round(rr[m.IDX[2050]]["active"]),
                 round(ee[m.IDX[2050]]["retirements"]), round(cum),
                 "Central case" if _age == 24 else ""])
m.AGE_AT_REG = _base_age
end_r = write_table(ws, ["Age assumed at registration", "Active doctors 2025",
                         "Active doctors 2050", "Retirements in 2050",
                         "Cumulative retirements 2026-2050", "Note"],
                    rows, [22, 18, 18, 18, 26, 14],
                    {2: "#,##0", 3: "#,##0", 4: "#,##0", 5: "#,##0"})
for k, t in enumerate([
 "Age is never observed in the register. Only year of registration is recorded, so "
 "every age in the model is inferred as 24 plus years elapsed.",
 "Registration at 24 assumes MBBS entry at 18 plus 5.5 years to qualification.",
 "The stock is robust to this assumption and the retirement profile is not. Across "
 "the range 22 to 26 the 2050 stock moves by under 3 per cent, while cumulative",
 "retirements to 2050 move by about 18 per cent. Any statement about when Tamil "
 "Nadu's retirement wave arrives therefore carries this uncertainty.",
 "The NMC detail records carry birth dates. A stratified sample of 5,000 to 8,000 "
 "records would replace this assumption with an estimate."]):
    ws.cell(row=end_r + 1 + k, column=1, value=t).font = BODY

# -------------------------------------------------------------- Parameters
ws = new_sheet("Parameters")
P = [
 ("Population", "NCP projection 2011-2025", "72.1M to 77.3M",
  CITE_NCP + " Used verbatim to 2025."),
 ("Population", "Growth rate beyond 2025", "0.298%/yr, held constant to 2050",
  "The NCP series' own 2021 to 2025 rate. Gives 83.3M in 2050. The NCP plateau "
  "(76.0M in 2050) is the alternative path in the dashboard."),
 ("Seats, government", "College-level starting position",
  "37 colleges, 5,200 seats, mean 141",
  "MGRMU 2025: sixteen colleges at 100 seats, sixteen at 150, one at 200, four at the "
  "NMC ceiling of 250. Tamil Nadu has 37 government medical colleges across 38 "
  "districts, so coverage is near-universal and the realistic lever is raising intake "
  "at existing colleges rather than building new ones."),
 ("Seats, government", "Frozen path (Constrained)", "5,200, held flat",
  "A no-action COUNTERFACTUAL, not a forecast. Seats are set by government, so this "
  "is the path only if no intake decision is taken anywhere for 25 years."),
 ("Seats, government", "Consolidation path (Status quo)", "6,000 by 2035",
  "The sixteen 100-seat colleges rise to 150. The smallest realistic policy step."),
 ("Seats, government", "Full-capacity path (Policy push)", "9,250 by 2045",
  "Every one of the 37 colleges reaches the NMC ceiling of 250 seats. Requires no new "
  "college at all. Ramps are rounded to NMC's 50-seat approval blocks."),
 ("Seats, private", "Logistic ceiling K", "8,000 / 10,000 / 12,000",
  "Scenario range. Ceilings reasoned from NMC faculty norms, clinical-material "
  "(bed) requirements, the NEET-qualified applicant pool, and a declining state "
  "population."),
 ("Seats, private", "Fitted growth rate and inflection year",
  "; ".join(f"K={k:,}: r={v[0]:.3f}, {2015 + v[1]:.1f}" for k, v in m.PVT_PARAMS.items()),
  "Least-squares fit in levels to the 2015-16 to 2025-26 Selection Committee series."),
 ("Seats, PG", "Linear slope", "+321.5 seats per year",
  "OLS in levels on the full-basis PG series 2021-2025. Increments are "
  "near-constant, so linear rather than CAGR."),
 ("Seats, PG", "Structural cap", "70% of total MBBS seats",
  "You cannot train more specialists than the MBBS pipeline delivers."),
 ("Pipeline", "MBBS fill rate", "99.8%",
  "MGRMU seat-count data, sanctioned against admitted, 2021-2025. For 2025, "
  "9,053 admitted against 9,100 approved."),
 ("Pipeline", "PG fill rate", "97.8%",
  "MGRMU seat-count data, PG Medical sheet, 2021-2024."),
 ("Pipeline", "Completion rate", "95%",
  "NMC benchmark, also used in the Andhra Pradesh model. Validated: PG predicted "
  "2,707 against 2,717 observed."),
 ("Pipeline", "MBBS lag", "6 years",
  "4.5 years academic plus 12-month Compulsory Rotatory Medical Internship; "
  "registration on completion of internship."),
 ("Pipeline", "PG lag", "3 years",
  "MD, MS, DM, M.Ch and DNB broad-speciality course duration."),
 ("Entrants", "External residual", f"{m.RESIDUAL:,.0f} per year, held flat",
  "Observed registrations minus modelled domestic output, 2021-2025. Deemed "
  "universities plus out-of-state returnees. No trend across the five observations."),
 ("Entrants", "Foreign medical graduates", "+170 per year, ceiling 3,000",
  "TNMC registrations 2020-2025, rising from 640 to 1,606. Linear, not CAGR; "
  "capped as domestic capacity expands."),
 ("Stock", "Age at registration", "24 years",
  "MBBS entry at 18 plus 5.5 years. INFERRED, not observed. Would be replaced by "
  "actual birth dates from a stratified NMC detail-record sample."),
 ("Stock", "Mortality", "0.08% under 35, rising to 13% at 85 and over",
  "Age-banded annual rates. Indian adult mortality adjusted downward for the "
  "professional class."),
 ("Stock", "Emigration and out-of-state loss", "1.2% per year, ages 25-45",
  "CALIBRATED, not observed. Approximately 22% cumulative. The model's most "
  "consequential parameter; see the Sensitivity sheet. Estimable from the NMC "
  "removedStatus field via a stratified sample."),
 ("Stock", "Retirement participation",
  "100% under 60; 85% at 60-64; 60% at 65-69; 30% at 70-74; 12% at 75-79; 3% at 80+",
  "A retirement hazard rather than a cliff. Government service retires at 60, "
  "private practice later."),
 ("Need", "WHO density norm", "44.5 skilled health workers per 10,000",
  "WHO, Global Strategy on Human Resources for Health: Workforce 2030 (2016)."),
 ("Need", "Doctor share of the norm", "25%, doctor to nurse 1 to 3",
  "WHO (2016). Gives 11.125 doctors per 10,000 population."),
]
end = write_table(ws, ["Module", "Parameter", "Value", "Basis and source"],
                  [list(x) for x in P], [16, 30, 34, 74])
for i in range(2, end):
    for j in range(1, 5):
        ws.cell(row=i, column=j).alignment = WRAP
    ws.row_dimensions[i].height = 30

# ------------------------------------------------------------- Methodology
P += [
 ("Revised projected need", "Income elasticity", "0.244 (0.231 + 0.531 - 0.518)",
  "Table 1 of " + CITE_LIU),
 ("Revised projected need", "Out-of-pocket elasticity", "-0.099", "Table 1, Liu et al. 2017"),
 ("Revised projected need", "Population 65+ elasticity", "0.516", "Table 1, Liu et al. 2017"),
 ("Revised projected need", "Real GSDP per capita growth", "6.5%/yr in 2026 easing to 4.0%/yr by 2050",
  "ASSUMED, FOR STATE CONFIRMATION. Replace with the State's forecast."),
 ("Revised projected need", "Growth of population aged 65+", "3.5%/yr",
  "ASSUMED, FOR STATE CONFIRMATION. Replace with the NCP 2019 age tables."),
 ("Revised projected need", "Anchor", "Modelled 2025 density, 19.6 per 10,000",
  "The paper's country fixed effect pins the level; the drivers move it from there."),
]
ws = new_sheet("Methodology")
M = [
 ("Framework",
  "A stock-and-flow supply model with an explicit education pipeline, following the "
  "WHO / World Bank Health Labour Market Framework as operationalised in " + CITE_LIU +
  " and in " + CITE_WHO_HLMA + " The same three-lens structure, Supply, Need and "
  "Demand, used in the Centre's Andhra Pradesh projection. This round builds the "
  "Supply lens for Tamil Nadu to 2050, sets it against the WHO need floor, and adds "
  "the revised projected need line, which is the demand side of the same paper."),
 ("Why not CAGR",
  "Growth forms were tested, not assumed. Comparing a levels-linear fit against a "
  "log-linear (CAGR) fit on a common scale, RMSE in seats, since a log model's "
  "R-squared is not comparable with a levels model's: private MBBS seats fit "
  "exponential over the full eleven years (RMSE 163 against 237) but linear over the "
  "recent window, 2019-20 onward (RMSE 83 against 150). The exponential advantage "
  "comes entirely from the small 2015-16 base. Increments have been flatly additive "
  "at about 460 seats per year since 2019-20."),
 ("Why not CAGR, continued",
  "Government MBBS is not a growth curve at all: five of ten historical years show "
  "zero change, three are large jumps, and the last four years total +25 seats. It is "
  "modelled as a step process. Extrapolating CAGR to 2050 would give 183,707 private "
  "MBBS seats against about 10,000 on a bounded fit, an eighteen-fold error. CAGR is "
  "used in this work only as a descriptive statistic for history, never as the "
  "forecast engine."),
 ("Why the logistic",
  "A logistic curve is approximately exponential in its early phase and approximately "
  "linear through its middle phase. Fitting private MBBS with a ceiling of 10,000 puts "
  "the inflection at 2025.5, so Tamil Nadu's private capacity is passing through its "
  "linear phase now. That is why the recent window reads linear while the full series "
  "reads exponential. Both are the same curve at different points."),
 ("Module 1, population",
  "NCP / MoHFW official projections used verbatim to 2025. From 2026 the population "
  "grows at the NCP series' own 2021 to 2025 rate, 0.298% per year, held constant. "
  "Result: 77.32 million in 2025, 79.89 million in 2036, 83.29 million in 2050. The "
  "NCP plateau near 78 million is kept as an alternative path in the dashboard."),
 ("Module 6b, revised projected need",
  "A need line revised for the state's own economy, on the demand model of Liu et al. 2017: "
  "ln(physicians per 1,000) = -9.882 + 0.231 ln GDPpc(t-1) + 0.531 ln GDPpc(t-4) "
  "- 0.518 ln GDPpc(t-5) - 0.099 ln OOPpc(t-2) + 0.516 ln Pop65(t-3) + country effect, "
  "estimated by GLM with country fixed effects on 165 countries, 1990 to 2013. The "
  "country effect pins the level, so the line is anchored on Tamil Nadu's modelled "
  "2025 density and moved by the three drivers, whose growth paths are assumptions "
  "marked for State confirmation. Result: 151,248 in 2025, 233,187 in 2040, 305,668 "
  "in 2050. Supply exceeds it throughout; the gap is 211,989 doctors in 2050."),
 ("Module 2, seats",
  "Government and private modelled separately because they are on different "
  "trajectories. Government is a step and scenario process. Private follows bounded "
  "logistic growth fitted in levels at three ceilings. PG seats grow linearly, capped "
  "at 70% of total MBBS seats."),
 ("Module 3, pipeline",
  "Seats to admissions via the fill rate, admissions to graduates via the completion "
  "rate, graduates to the register with a lag. Fill rates are estimated directly from "
  "MGRMU sanctioned-against-admitted data: 99.8% for MBBS, 97.8% for PG. Completion "
  "is 95%. The MBBS lag is six years, being 4.5 academic years plus a twelve-month "
  "internship; the PG lag is three years."),
 ("Module 4, external entrants",
  "The residual between observed registrations and modelled domestic output, "
  "decomposed into foreign medical graduates, which are observed and projected "
  "linearly to a ceiling of 3,000, and deemed-university plus out-of-state entrants, "
  "held flat at the observed mean because the five observations show no trend."),
 ("Module 5, stock and flow",
  "Register cohorts are rolled forward from 1927, with age inferred from registration "
  "year at an assumed age of 24. Three separate outflows apply: age-specific "
  "mortality; emigration and out-of-state loss at 1.2% per year across ages 25 to 45; "
  "and a retirement hazard reducing participation from age 60."),
 ("Module 6, specialists",
  "Specialists are tracked as a transition within the stock, never as an addition to "
  "it. An MD or MS entrant is an already-registered MBBS doctor, so adding PG output "
  "to MBBS output would double-count every specialist."),
 ("Deemed overlap correction",
  "The Selection Committee and MGRMU private MBBS seat series agree exactly for "
  "2021-22 and 2022-23, then diverge by 450, 550 and 850 seats. The gap is most "
  "likely deemed universities entering the Committee sheet. Because deemed output is "
  "already inside the external-entrant residual, which was estimated from years when "
  "the two sources agreed, those graduates would otherwise be counted twice. The "
  "correction is 850 x 0.998 x 0.95 = 806 graduates per year, applied from "
  "registration year 2031, with the smaller observed gaps used for 2029 and 2030."),
 ("Live formulas",
  "The workbook computes rather than reports. Every parameter sits in one cell on the "
  "Inputs sheet and the model sheets reference it, so a reviewer can change an "
  "assumption and watch the projection move. Typed numbers appear only where a "
  "formula is not possible: published population data, observed seats and "
  "registrations, and the three cohort-model flows (retirements, deaths, migration) "
  "which depend on the full age distribution of the register. Every formula column "
  "was checked cell by cell against the model and reproduces it exactly."),
 ("Validation",
  "The model was tested against data it was not fitted to. The scraped register and "
  "the TNMC workbook agree to within 0.17%, and exactly in 2024. The PG pipeline "
  "predicted 2,707 pass-outs for 2024 against 2,717 observed. The PG coverage factor "
  "gives 4,630 against an independent 4,629."),
 ("Key caveat",
  "This is an unconstrained supply projection. It answers how many doctors Tamil Nadu "
  "will produce and retain on current trajectories, not how many the health system "
  "will employ. A 2050 density of 68.7 per 10,000, about twice the OECD average, is "
  "the model reporting that current seat-expansion trajectories exceed any plausible "
  "absorption capacity. That is the finding, not a forecast of employment."),
 ("Limitation, seat sources",
  "The Selection Committee and MGRMU disagree on private MBBS seats from 2023, being "
  "4,750 against 3,900 in 2025-26. The gap is most likely deemed universities. The "
  "model uses the Selection Committee basis and places deemed output in the external "
  "residual."),
 ("Limitation, emigration",
  "The emigration parameter is calibrated, not observed. No direct measure of Tamil "
  "Nadu doctor out-migration exists. The Sensitivity sheet prices the uncertainty."),
 ("Limitation, age",
  "Age is inferred from registration year, not observed. Retirement timing is the "
  "figure most sensitive to this assumption."),
 ("Other cadres, scope",
  "Nursing at degree level, dental, pharmacy, physiotherapy and occupational therapy, "
  "the five AYUSH streams and allied health at postgraduate level are projected on the "
  "same chain as doctors: seats, fill rate, completion rate, course length lag. Twenty "
  "cadres in all. Every rate is estimated from the MGRMU files rather than assumed, "
  "except where the Cadre parameters sheet records it as assumed."),
 ("Other cadres, what these figures are NOT",
  "These are annual qualified output, not a practising workforce. No register "
  "comparable to the National Medical Commission register exists for any of these "
  "cadres, so there is no opening stock, no attrition for death, retirement or "
  "migration, and no need or demand benchmark. They cannot be added to the doctor "
  "stock, and a practising total across all cadres is not possible from this data."),
 ("Other cadres, fill rates",
  "Fill rates outside medicine are far lower and matter far more. MBBS runs at 99 per "
  "cent of sanctioned seats filled, but Post Basic BSc Nursing is at 48 per cent, MSc "
  "Nursing 59, allied health PG 65 and BHMS 66. For those cadres sanctioned seats "
  "overstate real output substantially, which is why the model runs on admissions."),
 ("Other cadres, seat growth",
  "With four or five annual observations per cadre a fitted ceiling cannot be "
  "supported, so two paths are given: seats frozen at the last observed level, and the "
  "fitted linear trend continued to 2035 then held flat. Fitted in levels, never as a "
  "compound rate. Totals begin at 2027, the first year in which every course length "
  "has a source cohort, since seat data starts in 2021 and the longest course is six "
  "years."),
 ("Other cadres, known gaps",
  "GNM and ANM nursing are absent from every source file because they sit with the "
  "Tamil Nadu Nurses and Midwives Council rather than the university, so nursing here "
  "is degree level only and understates total nursing output. Allied health at degree "
  "level is present but internally inconsistent and is excluded. Twelve single "
  "pass-out years across various cadres were excluded as partial extracts. All of this "
  "is itemised on the Data quality and Not covered sheets."),
 ("Limitation, other",
  "Population beyond 2025 is our own extension, not an official projection. The model "
  "is state-level only, with no district, urban-rural, or public-private employment "
  "split. Superspecialty (DM and M.Ch) and DNB are folded into the PG account, since "
  "only one year of seat data exists for each."),
]
end = write_table(ws, ["Section", "Text"], [list(x) for x in M], [24, 118])
for i in range(2, end):
    for j in (1, 2):
        ws.cell(row=i, column=j).alignment = WRAP
    ws.row_dimensions[i].height = 74

# ------------------------------------------------------------- Provenance
ws = new_sheet("Provenance")
P2 = [
 ("TAKEN DIRECTLY", "MBBS seats, government and self-financing, 2015-16 to 2025-26",
  "2,655 to 5,200 and 1,010 to 4,750",
  "Selection Committee, UG MBBS BDS Data Sheet.xlsx, sheet UG MBBS SS"),
 ("TAKEN DIRECTLY", "Government college seat distribution 2025",
  "37 colleges: 16 at 100, 16 at 150, 1 at 200, 4 at 250",
  "MGRMU SEAT COUNT STAT 15052026.xlsx, sheet MBBS"),
 ("TAKEN DIRECTLY", "PG seats 2025-26, full basis", "5,534",
  "Selection Committee MD Seats (3,629) plus MS Seats (1,905)"),
 ("TAKEN DIRECTLY", "TNMC MBBS registrations 2020-2025",
  "8,409 / 8,058 / 10,367 / 9,370 / 9,722 / 10,839",
  "TNMC Registrations 2020 to 2025.xlsx, sheet MBBS TNMC"),
 ("TAKEN DIRECTLY", "TNMC foreign medical graduate registrations",
  "640 / 1,268 / 1,287 / 1,622 / 1,427 / 1,606", "Same workbook and sheet"),
 ("TAKEN DIRECTLY", "TNMC broad-speciality PG registrations",
  "2,702 / 4,210 / 5,632 / 5,279 / 2,917 / 8,069",
  "Same workbook, sheets MD, MS, DNB MedPG, DNB Surg PG"),
 ("TAKEN DIRECTLY", "Seats, admissions and pass-outs, 20 other cadres", "2021 to 2025",
  "MGRMU SEAT COUNT STAT and passout COUNT STAT, 26 sheets each"),
 ("ESTIMATED BY US", "MBBS fill rate", "0.998",
  "Admissions over sanctioned seats, pooled 2021-2025. 2025: 9,053 of 9,100"),
 ("ESTIMATED BY US", "PG fill rate", "0.978", "Same method, MGRMU PG MEDICAL sheet"),
 ("ESTIMATED BY US", "Fill rates, 20 other cadres", "0.48 to 0.97",
  "Same method per cadre. Post Basic Nursing 0.48, BSc Nursing 0.97"),
 ("ESTIMATED BY US", "Completion rates, 11 of 20 cadres", "0.76 to 0.99",
  "Pass-outs over the admissions of the cohort one course length earlier"),
 ("ESTIMATED BY US", "External entrant residual", "5,381 per year",
  "Observed registrations minus modelled domestic output, 2021-2025. Year values "
  "4,583 / 6,328 / 4,961 / 5,456 / 5,577, no trend across the five"),
 ("ESTIMATED BY US", "PG external residual", "1,599 per year",
  "Same method on the two PG years with seat-driven output"),
 ("ESTIMATED BY US", "PG coverage factor", "1.407",
  "MGRMU basis to full basis. Reproduces 4,630 against the Committee's 4,629"),
 ("ESTIMATED BY US", "Private logistic growth and midpoint", "r 0.204, midpoint 2025.5",
  "Least squares in levels to 11 observed seat years, at K equal 10,000"),
 ("ESTIMATED BY US", "PG seat slope", "321.5 seats per year",
  "OLS in levels on the 2021-2025 full-basis series"),
 ("ESTIMATED BY US", "Seat-source gap", "450 / 550 / 850 seats",
  "Selection Committee minus MGRMU private MBBS, seat years 2023, 2024, 2025"),
 ("ESTIMATED BY US", "Cohort flows and 2011 opening stock",
  "Retirements, deaths, migration each year",
  "Computed by rolling the scraped register forward. Typed into the workbook because "
  "they depend on the full age distribution and cannot be written as a formula"),
 ("DERIVED, ARITHMETIC SHOWN", "Deemed overlap correction", "806 per year from 2031",
  "850 seats x 0.998 fill x 0.95 completion. Live in cell Inputs!B13"),
 ("DERIVED, ARITHMETIC SHOWN", "WHO doctor requirement", "11.125 per 10,000", "44.5 x 0.25"),
 ("DERIVED, ARITHMETIC SHOWN", "Government full-capacity headroom", "9,250 seats",
  "37 colleges x the NMC ceiling of 250"),
 ("DERIVED, ARITHMETIC SHOWN", "Register total", "201,908",
  "193,264 scraped, less the truncated 2025 figure 3,785, plus TNMC 2025 of 12,445"),
 ("ASSUMED BY US, NO SOURCE", "Private MBBS ceilings", "8,000 / 10,000 / 12,000",
  "EXPERT-JUDGMENT SCENARIO BOUNDS. No arithmetic connects faculty norms, beds or the "
  "applicant pool to these numbers. FOR STATE CONFIRMATION"),
 ("ASSUMED BY US, NO SOURCE", "PG seat cap", "70 per cent of MBBS seats",
  "OUR OWN FIGURE, not sourced and not from the Andhra Pradesh model. The ratio was 56 "
  "per cent in 2025-26. Binds only after 2040. FOR STATE CONFIRMATION"),
 ("ASSUMED BY US, NO SOURCE", "Government seat targets and years",
  "6,000 by 2035; 9,250 by 2045",
  "POLICY SCENARIOS. Built from the observed college distribution, but the levels and "
  "timing are ours. FOR STATE CONFIRMATION"),
 ("ASSUMED BY US, NO SOURCE", "Emigration and out-of-state loss",
  "1.2 per cent per year, ages 25-45",
  "CALIBRATED, not observed. No direct measure exists. The single most consequential "
  "parameter in the model"),
 ("ASSUMED BY US, NO SOURCE", "Age at registration", "24 years",
  "INFERRED. Age is never recorded in the register, only year of registration"),
 ("ASSUMED BY US, NO SOURCE", "Mortality by age",
  "0.08 per cent under 35 to 13 per cent at 85 plus",
  "Indian adult mortality adjusted downward for the professional class. The adjustment "
  "is a judgment, not a measured differential"),
 ("ASSUMED BY US, NO SOURCE", "Retirement participation curve",
  "100 per cent under 60, then 85, 60, 30, 12, 3",
  "Reasoned from retirement at 60 in government service. Not measured"),
 ("ASSUMED BY US, RATE FROM SOURCE", "Population growth beyond 2025",
  "0.298 per cent a year, held constant",
  "The NCP series' 2021 to 2025 rate. The series itself plateaus from 2031"),
 ("ASSUMED BY US, NO SOURCE", "Completion rate, medical", "95 per cent",
  "NMC benchmark, also used in the Andhra Pradesh model. Validated indirectly on PG"),
 ("ASSUMED BY US, NO SOURCE", "Completion rate, 9 of 20 other cadres", "91 per cent",
  "Median of the 11 cadres where it could be observed"),
 ("EXTERNAL PUBLISHED SOURCE", "Population 2011 to 2025", "72.1M to 77.3M",
  CITE_NCP + " Used verbatim to 2025, 15 values"),
 ("EXTERNAL PUBLISHED SOURCE", "Revised need elasticities",
  "0.244 income, -0.099 out-of-pocket, 0.516 population 65+",
  "Table 1 of " + CITE_LIU),
 ("ASSUMED BY US, FOR STATE CONFIRMATION", "Revised need driver paths",
  "GSDP per capita 6.5% easing to 4.0%; OOP share constant; 65+ population 3.5%/yr",
  "Our own. To be replaced by the State's GSDP forecast and the NCP 2019 age tables"),
 ("EXTERNAL PUBLISHED SOURCE", "WHO density norm", "44.5 per 10,000",
  "WHO, Global Strategy on HRH: Workforce 2030 (2016)"),
 ("EXTERNAL PUBLISHED SOURCE", "Doctor share of the norm", "One quarter",
  "WHO (2016), doctor to nurse 1 to 3"),
 ("EXTERNAL PUBLISHED SOURCE", "Total fertility rate", "About 1.3 to 1.4",
  "Sample Registration System. Context for the NCP plateau; not used in the model"),
 ("EXTERNAL PUBLISHED SOURCE", "NMC seat ceiling per college", "250 MBBS seats",
  "National Medical Commission norms"),
 ("SCRAPED BY US", "Full TNMC register", "193,264 doctors, 1927 to 2025",
  "NMC Indian Medical Register, scraped via run_all.py, endpoint "
  "nmc.org.in/MCIRest/open/getPaginatedData, state council id 21"),
]
end_r = write_table(ws, ["Class", "Figure", "Value", "Source, or how we computed it"],
                    [list(x) for x in P2], [26, 44, 34, 78])
for i in range(2, end_r):
    for j in range(1, 5):
        ws.cell(row=i, column=j).alignment = WRAP
    ws.row_dimensions[i].height = 40
for k, t in enumerate([
 "Roughly 60 per cent of the figures come straight from the supplied files or are "
 "computed from them. About 25 per cent are external published sources, almost all of "
 "that the population series and the WHO norm.",
 "About 15 per cent are our own assumptions. Five of those matter: the private seat "
 "ceilings, the PG cap, the government seat targets, the emigration rate and age at "
 "registration.",
 "The first three need the State's confirmation. The last two need data that exists but "
 "we do not hold: the NMC detail records carry birth dates and removal status.",
 "Data altered or excluded is on the Data quality sheet. Cadres not projected are on the "
 "Not covered sheet. Neither was dropped silently."]):
    ws.cell(row=end_r + 1 + k, column=1, value=t).font = BODY

# ------------------------------------------------------------ Data Sources
ws = new_sheet("Data Sources")
S = [
 ("Primary data", "UG MBBS / BDS Data Sheet", "TN Medical Selection Committee",
  "2015-16 to 2025-26", "MBBS seats, Government against Self-financing. The seat "
  "module backbone.", "new data/Selection committee/UG MBBS BDS Data Sheet.xlsx",
  "User-provided"),
 ("Primary data", "MD Seats 10 years", "TN Medical Selection Committee",
  "2016-17 to 2025-26", "MD seats by college, with a four-way Government, Central, "
  "Private-TNMMU and Private-Deemed split in recent years",
  "new data/Selection committee/MD Seats 10 years.xlsx", "User-provided"),
 ("Primary data", "MS Seats 10 years", "TN Medical Selection Committee",
  "2016-17 to 2025-26", "MS seats by college, same structure",
  "new data/Selection committee/MS Seats 10 years.xlsx", "User-provided"),
 ("Primary data", "PG Diploma seats in 10 years", "TN Medical Selection Committee",
  "2016-17 to 2025-26", "PG Diploma seats, a collapsing stream from 396 to 25",
  "new data/Selection committee/PG Diploma seats in 10 years.xlsx", "User-provided"),
 ("Primary data", "Super Speciality Seat Matrix, AIQ and SQ",
  "TN Medical Selection Committee", "2025-26 only",
  "DM and M.Ch seats, 422 in total across 13 government colleges",
  "new data/Selection committee/SUPER SPECIALITY SEAT MATRIX 2025-26AIQ_SQ.xlsx",
  "User-provided"),
 ("Primary data", "DNB Seat Matrix", "TN Medical Selection Committee", "2025-26 only",
  "DNB seats at government district hospitals, about 40",
  "new data/Selection committee/DNB SEAT MATRIX 2025-2026.xlsx", "User-provided"),
 ("Primary data", "Seat Count Stat, 15.05.2026", "TN Dr. M.G.R. Medical University",
  "2021-2025, 27 courses", "Sanctioned seats and actual admissions by institution. "
  "Gives the fill rate directly.", "new data/TN MGRMU/SEAT COUNT STAT 15052026.xlsx",
  "User-provided"),
 ("Primary data", "Passout Count Stat, 15.05.2026", "TN Dr. M.G.R. Medical University",
  "2021-2025, 26 courses", "Actual graduates by institution and category. Gives the "
  "completion rate.", "new data/TN MGRMU/passout COUNT STAT 15052026.xlsx",
  "User-provided"),
 ("Primary data", "Consolidated all-course faculty, 18.6.2025",
  "TN Dr. M.G.R. Medical University", "2025",
  "Faculty norms per course. Informs the teaching-capacity ceiling.",
  "new data/TN MGRMU/consolidated all course 18.6.2025.xlsx", "User-provided"),
 ("Primary data", "TNMC Registrations 2020 to 2025", "Tamil Nadu Medical Council",
  "2020-2025", "Registrations by course, year and sex: MBBS, FMG, MD, MS, DM, M.Ch, "
  "DNB and PG Diploma", "new data/Medical Council/TNMC Registrations 2020 to 2025.xlsx",
  "User-provided"),
 ("Primary data", "Prospectuses: MBBS, DNB, Superspeciality",
  "TN Selection Committee / Govt. of Tamil Nadu", "2024-25 and 2025-26",
  "Course durations, internship rules, eligibility and service-candidate quotas",
  "new data/prospectus/*.pdf", "User-provided"),
 ("Scraped data", "Indian Medical Register, Tamil Nadu", "National Medical Commission",
  "1927-2025, 193,264 records", "Full TNMC register with year of registration per "
  "doctor. The stock model's backbone.",
  "imr_tamilnadu_output/tamilnadu_doctors.csv, scraped via run_all.py from "
  "nmc.org.in/MCIRest/open/getPaginatedData with smcId=21", "Scraped; 2025 truncated"),
 ("Published source", "Population Projections for India and States 2011-2036",
  "National Commission on Population, MoHFW", "2011-2036",
  "The population spine, denominator for every density figure",
  CITE_NCP + " https://nhm.gov.in/New_Updates_2018/Report_Population_Projection_2019.pdf",
  "Verified via Wayback mirror; live link returns 403"),
 ("Published source", "Health and Family Welfare Policy Note 2025-26",
  "Govt. of Tamil Nadu", "2025-26", "Current facility counts and budget context",
  "https://cms.tn.gov.in/cms_migrated/document/docfiles/hfw_e_pn_2025_26.pdf",
  "Verified"),
 ("Published source", "Health and Family Welfare Policy Notes, series",
  "Govt. of Tamil Nadu", "2011-12 to 2025-26",
  "Historical facility counts, 12 confirmed years. 2013-14 not found; 2015-16 never "
  "published.", "https://cms.tn.gov.in/cms_migrated/document/docfiles/"
  "hfw_e_pn_<yyyy>_<yy>.pdf", "Verified for 12 of 15 years"),
 ("Published source", "Rural Health Statistics 2018-19 to 2021-22",
  "MoHFW / NHM, Govt. of India", "2018-19 to 2021-22",
  "CHC baseline counts, cross-check in the earlier Need round",
  "https://archive.org/download/PARI.rural-health-statistics-2021-22/"
  "rural-health-statistics-2021-22.pdf", "Verified via archive.org mirror"),
 ("Published source", "IPHS 2022 Volume I, SDH and District Hospital",
  "NHSRC / MoHFW", "2022 norms", "Bed-strength staffing table and SDH allocation rule",
  "https://nhsrcindia.org/sites/default/files/Volume%201_SDH-DH_0.pdf", "Verified"),
 ("Published source", "IPHS 2022 Volume II, CHC and Urban CHC", "NHSRC / MoHFW",
  "2022 norms", "CHC population norm",
  "https://nhsrcindia.org/sites/default/files/"
  "CHC%20IPHS%202022%20Guidelines%20pdf.pdf", "Verified"),
 ("Published source", "IPHS 2022 Volume III, PHC and Urban PHC", "NHSRC / MoHFW",
  "2022 norms", "PHC and UPHC population and staffing norms",
  "https://nhsrcindia.org/sites/default/files/PHC%20IPHS_2022_Guideline_pdf.pdf",
  "Verified"),
 ("Published source", "IPHS 2022 Volume IV, SHC-HWC and Urban HWC", "NHSRC / MoHFW",
  "2022 norms", "Sub-centre population and staffing norms",
  "https://nhsrcindia.org/sites/default/files/"
  "SHC-HWC%20&%20UHWC%20IPHS%202022%20Guidelines%20pdf.pdf", "Verified"),
 ("Published source", "Projection of District-Level Annual Population", "IIPS, Mumbai",
  "2023", "District-level population. Not used this round; the best lead for a "
  "district split.", "https://www.iipsindia.ac.in/sites/default/files/"
  "FULL_REPORT_WITH_FINAL_TABLES.pdf", "Not pursued"),
 ("Method reference", "Global Health Workforce Labor Market Projections for 2030",
  "Liu JX, Goryakin Y, Maeda A, Bruckner T, Scheffler RM", "2017",
  "The principal method reference: the supply, need and demand framework, and the "
  "demand elasticities in Table 1 used for the revised projected need line",
  CITE_LIU + " https://doi.org/10.1186/s12960-017-0187-2 ; working paper at "
  "https://documents1.worldbank.org/curated/en/546161470834083341/pdf/WPS7790.pdf",
  "Verified, coefficients read from Table 1"),
 ("Method reference", "Health Labour Market Analysis Guidebook",
  "World Health Organization", "2021", "Operational guidance for the framework above",
  CITE_WHO_HLMA + " https://www.who.int/publications/i/item/9789240035546",
  "Method reference"),
 ("Norm reference",
  "Global Strategy on Human Resources for Health: Workforce 2030",
  "World Health Organization", "2016",
  "The 44.5 skilled health workers per 10,000 density norm",
  CITE_WHO_GS + " https://www.who.int/publications/i/item/9789241511131",
  "Norm reference"),
 ("Norm reference", "Health workforce density and distribution",
  "World Health Organization", "2016",
  "Doctor to nurse ratio of 1 to 3, splitting the density norm", "WHO, Geneva",
  "Norm reference"),
 ("Method reference",
  "Size, composition and distribution of India's health workforce",
  "Karan and others, Human Resources for Health", "2021",
  "Attrition benchmark of 7% used in the Andhra Pradesh round, superseded here by "
  "the cohort model", "Human Resources for Health 2021", "Method reference"),
 ("Reference", "Sample Registration System", "Office of the Registrar General of India",
  "latest available", "TFR of about 1.3 to 1.4 and age-specific mortality, "
  "underpinning the population extension and the attrition curve", "ORGI, New Delhi",
  "Reference"),
 ("Primary data", "Seat Count Stat, 15.05.2026, all cadres",
  "TN Dr. M.G.R. Medical University", "2021 to 2025, 27 course sheets",
  "Sanctioned seats and admissions for nursing, dental, pharmacy, rehabilitation, "
  "AYUSH and allied health. Source of the cadre seat series and all fill rates.",
  "new data/TN MGRMU/SEAT COUNT STAT 15052026.xlsx", "Supplied by the user"),
 ("Primary data", "Passout Count Stat, 15.05.2026, all cadres",
  "TN Dr. M.G.R. Medical University", "2021 to 2025, 26 course sheets",
  "Actual qualifications for the same cadres. Source of all completion rates.",
  "new data/TN MGRMU/passout COUNT STAT 15052026.xlsx", "Supplied by the user"),
 ("Derived input", "Selection Committee against MGRMU seat comparison",
  "Both bodies, compared by us", "2021-22 to 2025-26",
  "Source of the deemed overlap correction. The two private MBBS series agree "
  "exactly for 2021-22 and 2022-23 then diverge by 450, 550 and 850 seats, which is "
  "taken as deemed university capacity entering the Committee sheet.",
  "Both files listed above; comparison in tn_supply_model.py, SC_MGRMU_GAP",
  "DERIVED, arithmetic on the Assumptions sheet"),
 ("Not held", "Tamil Nadu Nurses and Midwives Council register",
  "Government of Tamil Nadu", "Not obtained",
  "Would supply the GNM and ANM seat and output series, and an opening stock for "
  "nursing. Its absence is the main limitation of the non-medical projections.",
  "Not held", "REQUIRED, NOT AVAILABLE"),
 ("Not held", "Indian Nursing Council registration data", "Indian Nursing Council",
  "Not obtained",
  "Alternative route to a nursing stock, equivalent to what the NMC register provided "
  "for doctors.", "Not held", "REQUIRED, NOT AVAILABLE"),
 ("Internal", "Projection of Andhra Pradesh, deck", "CMHS, IIM Ahmedabad", "June 2026",
  "Methodological template, the three-lens framework",
  "Projection_of_Andhra_Pradesh-3.pdf", "Internal"),
 ("Internal", "TN_Projection_shan.xlsx and Overleaf deck", "CMHS, IIM Ahmedabad",
  "2026", "The earlier Tamil Nadu Need-lens round, WHO and IPHS facility norms to 2036",
  "TN_Projection_shan.xlsx, overleaf/", "Internal"),
]
end = write_table(ws,
    ["Category", "Source", "Publisher", "Coverage", "Used for", "Location or link",
     "Status"],
    [list(x) for x in S], [16, 38, 30, 20, 46, 60, 22])
for i in range(2, end):
    for j in range(1, 8):
        ws.cell(row=i, column=j).alignment = WRAP
    ws.row_dimensions[i].height = 34
ws.freeze_panes = "A2"

# Final pass: enforce Times New Roman everywhere, strip any stray colour.
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None:
                c.font = Font(name=FONT, size=10, bold=c.font.bold, color="000000")

wb.save("TN_Health_Workforce_Projections_2050.xlsx")
print("Wrote TN_Health_Workforce_Projections_2050.xlsx")
print("Sheets:", wb.sheetnames)
