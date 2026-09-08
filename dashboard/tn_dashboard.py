"""
Tamil Nadu Health Workforce Projection Dashboard.

Run with:  streamlit run dashboard/tn_dashboard.py

Every assumption is editable in the sidebar. "Reset to defaults" restores the
values published in TN_Health_Workforce_Projections_2050.xlsx.
"""
import os, sys, io
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go

import tn_engine as E

st.set_page_config(page_title="Tamil Nadu Health Workforce Projections",
                   page_icon="■", layout="wide",
                   initial_sidebar_state="expanded")

# ---------------------------------------------------------------- styling
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,400;8..60,600;8..60,700&display=swap');
  html, body, [class*="css"], .stMarkdown, p, div, span, label {
      font-family: "Source Serif 4", "Times New Roman", Times, serif; }
  .main .block-container { padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1560px; }

  h1 { font-size: 2.05rem !important; font-weight: 700; letter-spacing: -0.02em;
       color: #0F4761; margin-bottom: 0.15rem !important; }
  .rule { height: 3px; background: linear-gradient(90deg,#0F4761 0%,#0F4761 22%,#dfe5ea 22%);
          margin: 0.35rem 0 0.55rem 0; border-radius: 2px; }
  .subtitle { color: #55606a; font-size: 0.97rem; margin-bottom: 1.35rem; line-height:1.5; }
  h2 { font-size: 1.22rem !important; font-weight: 700; color:#0F4761; margin-top: 1.5rem !important; }
  h3 { font-size: 1.04rem !important; font-weight: 700; color:#1d2a33; margin-top:1.1rem !important; }

  /* metric cards */
  [data-testid="stMetric"] { background:#ffffff; border:1px solid #e3e8ec;
      border-top:3px solid #0F4761; border-radius:7px; padding:0.85rem 1rem 0.7rem 1rem;
      box-shadow:0 1px 3px rgba(16,40,60,0.05); }
  [data-testid="stMetricValue"] { font-size: 1.72rem !important; font-weight: 700;
      color:#0F4761; letter-spacing:-0.02em; }
  [data-testid="stMetricLabel"] p { font-size: 0.8rem !important; color:#55606a;
      font-weight:600; letter-spacing:0.01em; }
  [data-testid="stMetricDelta"] { font-size: 0.78rem !important; }

  /* sidebar */
  section[data-testid="stSidebar"] { background:#f7f9fa; border-right:1px solid #e3e8ec; }
  section[data-testid="stSidebar"] h2 { font-size:0.95rem !important; margin-top:1.15rem !important;
      text-transform:uppercase; letter-spacing:0.07em; color:#0F4761;
      border-bottom:1px solid #dfe5ea; padding-bottom:0.3rem; }
  section[data-testid="stSidebar"] label p { font-size:0.86rem !important; }

  /* tabs */
  .stTabs [data-baseweb="tab-list"] { gap:2px; border-bottom:1px solid #dfe5ea; }
  .stTabs [data-baseweb="tab"] { font-size:0.95rem; font-weight:600; color:#55606a;
      padding:0.55rem 1.05rem; }
  .stTabs [aria-selected="true"] { color:#0F4761 !important; background:#eef3f7;
      border-radius:6px 6px 0 0; }

  /* callouts */
  .note { background:#eef3f7; border-left:4px solid #0F4761; padding:0.8rem 1.05rem;
          font-size:0.9rem; color:#1d2a33; margin:0.7rem 0 1.1rem 0; border-radius:0 6px 6px 0;
          line-height:1.55; }
  .warn { background:#fdf5e8; border-left:4px solid #9a6b00; padding:0.8rem 1.05rem;
          font-size:0.9rem; color:#3d2c00; margin:0.7rem 0 1.1rem 0; border-radius:0 6px 6px 0;
          line-height:1.55; }
  .finding { background:#0F4761; color:#ffffff; padding:0.95rem 1.2rem; border-radius:7px;
             font-size:0.98rem; line-height:1.55; margin:0.2rem 0 1.3rem 0; }
  .finding b { color:#ffffff; }
  .chartcap { color:#6b7680; font-size:0.82rem; margin:-0.7rem 0 1.0rem 0.2rem; }

  [data-testid="stDataFrame"] { border:1px solid #e3e8ec; border-radius:6px; }
  hr { border-color:#e3e8ec; }
  footer, #MainMenu, [data-testid="stToolbar"] { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

INK   = "#0F4761"      # house accent, used for the primary series
ACC   = ["#0F4761", "#C0392B", "#5B8FA8", "#7a7a7a", "#9aa7b1", "#2E6E4F"]
DASH  = ["solid", "dash", "dot", "dashdot", "longdash", "longdashdot"]
FILLS = ["#0F4761", "#5B8FA8", "#A8C2D0", "#7a7a7a"]
SPLIT = 2025.5        # last observed year boundary


def _shade_projection(fig, x):
    """Mark the projected region so observed and projected are never confused."""
    if not x or x[-1] <= SPLIT:
        return
    fig.add_vrect(x0=max(SPLIT, x[0]), x1=x[-1], fillcolor="#0F4761",
                  opacity=0.045, layer="below", line_width=0)
    fig.add_vline(x=SPLIT, line=dict(color="#9aa7b1", width=1, dash="dot"))
    fig.add_annotation(x=SPLIT, y=1.0, yref="paper", text="projected",
                       showarrow=False, xanchor="left", xshift=5, yshift=-6,
                       font=dict(size=10, color="#6b7680"))


def _layout(fig, title, ylab, height, subtitle=None, zero=False, xr=None):
    fig.update_layout(
        title=dict(text=f"<b>{title}</b>" + (f"<br><span style='font-size:11.5px;"
                        f"color:#6b7680;font-weight:400'>{subtitle}</span>" if subtitle else ""),
                   font=dict(size=15.5, color="#1d2a33"), x=0, xanchor="left",
                   y=0.98, yanchor="top"),
        height=height, plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Source Serif 4, Times New Roman, serif", size=12, color="#1d2a33"),
        margin=dict(l=64, r=104, t=72 if subtitle else 52, b=96),
        hovermode="x unified",
        hoverlabel=dict(bgcolor="white", bordercolor="#dfe5ea",
                        font=dict(family="Source Serif 4, serif", size=12)),
        legend=dict(orientation="h", yanchor="top", y=-0.17, xanchor="left", x=0,
                    font=dict(size=11.5), bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(title=None, showgrid=False, linecolor="#c8d1d8", ticks="outside",
                   tickcolor="#c8d1d8", tickfont=dict(size=11.5),
                   range=xr, dtick=10 if xr else None),
        yaxis=dict(title=dict(text=ylab, font=dict(size=11.5, color="#55606a")),
                   gridcolor="#eef1f4", linecolor="#c8d1d8", ticks="outside",
                   tickcolor="#c8d1d8", tickformat=",", tickfont=dict(size=11.5),
                   zeroline=False, rangemode="tozero" if zero else "normal"))
    return fig


def chart(title, ylab, series, x, annotate=True, height=455, subtitle=None, fill_first=True):
    """One chart. First series is the emphasis line and is filled beneath."""
    fig = go.Figure()
    for i, (name, data) in enumerate(series):
        yv = [data.get(k) if isinstance(data, dict) else data[j]
              for j, k in enumerate(x)]
        is_lead = (i == 0)
        fig.add_trace(go.Scatter(
            x=x, y=yv, name=name, mode="lines",
            line=dict(color=ACC[i % len(ACC)], width=2.9 if is_lead else 2.0,
                      dash=DASH[i % len(DASH)], shape="spline", smoothing=0.35),
            fill="tozeroy" if (is_lead and fill_first) else None,
            fillcolor="rgba(15,71,97,0.07)" if (is_lead and fill_first) else None,
            hovertemplate="%{y:,.0f}<extra>" + name + "</extra>"))
        if annotate and yv and yv[-1] is not None and yv[-1] == yv[-1]:
            fig.add_annotation(
                x=x[-1], y=yv[-1], text=f"<b>{yv[-1]:,.0f}</b>", showarrow=False,
                xanchor="left", xshift=8, bgcolor="rgba(255,255,255,0.85)",
                font=dict(size=11.5, color=ACC[i % len(ACC)]))
    _shade_projection(fig, x)
    pad = max(1.5, (x[-1] - x[0]) * 0.055)
    return _layout(fig, title, ylab, height, subtitle, zero=fill_first,
                   xr=[x[0] - 0.4, x[-1] + pad])


def stacked(title, ylab, series, x, height=455, subtitle=None):
    fig = go.Figure()
    for i, (name, data) in enumerate(series):
        yv = [data.get(k) for k in x]
        fig.add_trace(go.Scatter(
            x=x, y=yv, name=name, mode="lines", stackgroup="one",
            line=dict(color="white", width=1.4),
            fillcolor=FILLS[i % len(FILLS)],
            hovertemplate="%{y:,.0f}<extra>" + name + "</extra>"))
    _shade_projection(fig, x)
    return _layout(fig, title, ylab, height, subtitle, zero=True,
                   xr=[x[0], x[-1]])


def df_download(df, label, fname):
    st.download_button(label, df.to_csv(index=False).encode(),
                       file_name=fname, mime="text/csv")


# ------------------------------------------------------------------ sidebar
st.sidebar.markdown("## Projection settings")

if st.sidebar.button("Reset everything to published defaults", width="stretch"):
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.rerun()

D = E.Params()   # published defaults

end_year = st.sidebar.slider("Project to year", 2030, 2070, D.end_year, step=1,
                             key="end_year",
                             help="Horizon. The published workbook runs to 2050.")

st.sidebar.markdown("## Scenario")
gov_choice = st.sidebar.selectbox(
    "Government MBBS seats", list(E.GOV_SCENARIOS.keys()), index=1, key="gov_choice",
    help="Government sets these seats, so these are policy scenarios rather than "
         "extrapolations. Built from the 2025 college distribution: 37 colleges, "
         "sixteen still at 100 seats, only four at the NMC ceiling of 250.")
pvt_choice = st.sidebar.selectbox(
    "Private MBBS seats", list(E.PVT_SCENARIOS.keys()), index=1, key="pvt_choice",
    help="Expert-judgment scenario bounds. No arithmetic derives these ceilings.")

gov_target_d, gov_year_d = E.GOV_SCENARIOS[gov_choice]
pvt_K_d, pvt_r_d, pvt_t0_d = E.PVT_SCENARIOS[pvt_choice]

adv = st.sidebar.toggle("Edit individual assumptions", value=False, key="adv",
                        help="Off: the scenario above sets everything. "
                             "On: every parameter becomes editable.")

if adv:
    st.sidebar.markdown("## Seats")
    gov_target = st.sidebar.number_input("Government seat target", 3000.0, 20000.0,
                                         float(gov_target_d), step=50.0)
    gov_year = st.sidebar.number_input("Government target year", 2026, 2070,
                                       int(gov_year_d or end_year), step=1)
    gov_year = None if gov_target == D.gov_2025 else int(gov_year)
    pvt_K = st.sidebar.number_input("Private ceiling K", 4750.0, 40000.0,
                                    float(pvt_K_d), step=250.0,
                                    help="ASSUMPTION. Not derived from faculty "
                                         "norms, beds or the applicant pool.")
    pvt_r = st.sidebar.number_input("Private logistic growth r", 0.01, 1.0,
                                    float(pvt_r_d), step=0.001, format="%.3f")
    pvt_t0 = st.sidebar.number_input("Private midpoint, years after 2015", 0.0, 40.0,
                                     float(pvt_t0_d), step=0.25)
    pg_cap = st.sidebar.slider("PG cap, share of MBBS seats", 0.30, 1.20,
                               D.pg_cap_share, step=0.01,
                               help="OWN ASSUMPTION, not sourced. Actual ratio "
                                    "was 56 per cent in 2025-26.")
    pg_slope = st.sidebar.number_input("PG seat slope, per year", 0.0, 2000.0,
                                       D.pg_slope, step=5.0)

    st.sidebar.markdown("## Pipeline")
    fill_mbbs = st.sidebar.slider("MBBS fill rate", 0.50, 1.00, D.fill_mbbs, 0.001,
                                  help="OBSERVED from MGRMU sanctioned against admitted.")
    fill_pg = st.sidebar.slider("PG fill rate", 0.50, 1.00, D.fill_pg, 0.001)
    completion = st.sidebar.slider("Completion rate", 0.50, 1.00, D.completion, 0.01,
                                   help="NMC benchmark, validated indirectly on PG.")
    lag_mbbs = st.sidebar.number_input("MBBS lag, years", 4, 8, D.lag_mbbs)
    lag_pg = st.sidebar.number_input("PG lag, years", 2, 6, D.lag_pg)

    st.sidebar.markdown("## Entrants from outside the state system")
    residual = st.sidebar.number_input("External entrants per year", 0.0, 20000.0,
                                       float(D.external_residual), step=50.0,
                                       help="OBSERVED RESIDUAL, 2021 to 2025 mean. "
                                            "Deemed universities and out-of-state returnees.")
    gap_fwd = st.sidebar.number_input("Seat-source gap held forward, seats", 0.0, 3000.0,
                                      D.gap_forward, step=50.0,
                                      help="Drives the deemed overlap correction: "
                                           "gap x fill x completion.")
    fmg_base = st.sidebar.number_input("Foreign graduates, 2025 base", 0.0, 6000.0,
                                       D.fmg_base, step=10.0)
    fmg_inc = st.sidebar.number_input("Foreign graduates, increment per year", 0.0, 1000.0,
                                      D.fmg_increment, step=5.0)
    fmg_cap = st.sidebar.number_input("Foreign graduates, ceiling", 0.0, 20000.0,
                                      D.fmg_ceiling, step=100.0)
    pg_res = st.sidebar.number_input("PG external entrants per year", 0.0, 10000.0,
                                     float(D.pg_residual), step=25.0)

    st.sidebar.markdown("## Stock and attrition")
    age_reg = st.sidebar.slider("Age at registration", 20, 30, D.age_at_registration,
                                help="INFERRED, not observed. Retirement timing is "
                                     "most sensitive to this.")
    emig = st.sidebar.slider("Emigration and out-of-state loss, per year",
                             0.000, 0.060, D.emigration_rate, 0.001, format="%.3f",
                             help="CALIBRATED, not observed. The model's most "
                                  "consequential parameter.")
    emig_lo = st.sidebar.number_input("Emigration, from age", 20, 45, D.emigration_age_lo)
    emig_hi = st.sidebar.number_input("Emigration, to age", 30, 70, D.emigration_age_hi)
    ret60 = st.sidebar.slider("Still practising at 60 to 64", 0.0, 1.0, 0.85, 0.01)
    ret65 = st.sidebar.slider("Still practising at 65 to 69", 0.0, 1.0, 0.60, 0.01)
    ret70 = st.sidebar.slider("Still practising at 70 to 74", 0.0, 1.0, 0.30, 0.01)

    st.sidebar.markdown("## Population and benchmark")
    pop_g = st.sidebar.slider("Population growth by final year", -0.020, 0.010,
                              D.pop_growth_2050, 0.0005, format="%.4f")
    who_norm = st.sidebar.number_input("WHO density norm per 10,000", 10.0, 100.0,
                                       D.who_norm, step=0.5)
    who_share = st.sidebar.slider("Doctor share of the WHO norm", 0.10, 0.60,
                                  D.who_doctor_share, 0.01)
else:
    gov_target, gov_year = gov_target_d, gov_year_d
    pvt_K, pvt_r, pvt_t0 = pvt_K_d, pvt_r_d, pvt_t0_d
    pg_cap, pg_slope = D.pg_cap_share, D.pg_slope
    fill_mbbs, fill_pg, completion = D.fill_mbbs, D.fill_pg, D.completion
    lag_mbbs, lag_pg = D.lag_mbbs, D.lag_pg
    residual, gap_fwd = D.external_residual, D.gap_forward
    fmg_base, fmg_inc, fmg_cap = D.fmg_base, D.fmg_increment, D.fmg_ceiling
    pg_res = D.pg_residual
    age_reg, emig = D.age_at_registration, D.emigration_rate
    emig_lo, emig_hi = D.emigration_age_lo, D.emigration_age_hi
    ret60, ret65, ret70 = 0.85, 0.60, 0.30
    pop_g, who_norm, who_share = D.pop_growth_2050, D.who_norm, D.who_doctor_share

P = E.Params(
    end_year=int(end_year), pop_growth_2050=pop_g,
    gov_target=float(gov_target), gov_target_year=gov_year,
    pvt_ceiling=float(pvt_K), pvt_growth_r=float(pvt_r), pvt_midpoint=float(pvt_t0),
    pg_slope=float(pg_slope), pg_cap_share=float(pg_cap),
    fill_mbbs=float(fill_mbbs), fill_pg=float(fill_pg), completion=float(completion),
    lag_mbbs=int(lag_mbbs), lag_pg=int(lag_pg),
    external_residual=float(residual), gap_forward=float(gap_fwd),
    fmg_base=float(fmg_base), fmg_increment=float(fmg_inc), fmg_ceiling=float(fmg_cap),
    pg_residual=float(pg_res), age_at_registration=int(age_reg),
    emigration_rate=float(emig), emigration_age_lo=int(emig_lo),
    emigration_age_hi=int(emig_hi),
    participation_bands=[(60, 1.00), (65, ret60), (70, ret65), (75, ret70),
                         (80, 0.12), (999, 0.03)],
    who_norm=float(who_norm), who_doctor_share=float(who_share))

is_default = (P.to_dict() == E.Params(end_year=int(end_year)).to_dict()
              and gov_choice == list(E.GOV_SCENARIOS.keys())[1]
              and pvt_choice == list(E.PVT_SCENARIOS.keys())[1])

# ------------------------------------------------------------------- header
st.markdown("# Tamil Nadu Health Workforce Projections")
st.markdown('<div class="rule"></div>', unsafe_allow_html=True)
st.markdown(f'<div class="subtitle">Supply side projection to <b>{P.end_year}</b>, '
            'covering doctors, specialists and twenty other cadres.<br>'
            'Centre for Management of Health Services, Indian Institute of '
            'Management Ahmedabad.</div>', unsafe_allow_html=True)

@st.cache_data(show_spinner="Running the projection")
def run_doc(params_dict):
    return E.run_doctors(E.Params(**params_dict))

try:
    R = run_doc(P.to_dict())
except FileNotFoundError as err:
    st.error(str(err))
    st.stop()

yrs = R["years"]
proj_years = [y for y in yrs if y >= 2011]

if is_default:
    st.markdown('<div class="note"><b>Published defaults.</b> Every parameter is set '
                'to the value in TN_Health_Workforce_Projections_2050.xlsx. '
                'Change anything in the sidebar and every figure below recomputes.</div>',
                unsafe_allow_html=True)
else:
    st.markdown('<div class="warn"><b>Modified from published defaults.</b> These '
                'figures no longer match the workbook. Use "Reset everything to '
                'published defaults" in the sidebar to restore.</div>',
                unsafe_allow_html=True)

ey = P.end_year
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Active doctors, " + str(ey), f"{R['active'][ey]:,.0f}",
          f"{R['active'][ey] - R['active'][2025]:+,.0f} vs 2025")
c2.metric("Per 10,000 people", f"{R['density'][ey]:,.1f}",
          f"{R['density'][ey] - R['density'][2025]:+,.1f}")
c3.metric("Active specialists", f"{R['specialists'][ey]:,.0f}",
          f"{R['specialists'][ey] / R['active'][ey] * 100:.0f}% of stock")
c4.metric("Entrants per year", f"{R['entrants'][ey]:,.0f}")
c5.metric("Surplus vs WHO floor", f"{R['surplus'][ey]:,.0f}",
          f"{R['active'][ey] / R['who_need'][ey]:.1f}x the floor")

_pvt_share = R["pvt_seats"][ey] / R["total_seats"][ey] * 100
_ratio = R["entrants"][ey] / R["exits"][ey]
st.markdown(
    f'<div class="finding"><b>The finding.</b> Tamil Nadu does not have a doctor '
    f'shortage and does not develop one. It crossed the WHO floor around 2016 and '
    f'reaches <b>{R["density"][ey]:,.1f} per 10,000</b> by {ey}, roughly '
    f'{R["active"][ey] / R["who_need"][ey]:.0f} times the floor. The constraint is '
    f'control, not volume: private capacity reaches <b>{_pvt_share:.0f}% of MBBS seats</b> '
    f'by {ey}, and the state still adds <b>{_ratio:.1f} doctors for every one it loses</b>.'
    '</div>', unsafe_allow_html=True)

tabs = st.tabs(["Doctors", "Seats and pipeline", "Exits", "Other cadres",
                "Sensitivity", "Assumptions", "Data and sources"])

# ------------------------------------------------------------------ doctors
with tabs[0]:
    a, b = st.columns(2)
    with a:
        st.plotly_chart(chart("Active doctors against the WHO requirement", "Doctors",
                              [("Active doctors", R["active"]),
                               ("WHO requirement", R["who_need"])], proj_years,
                              subtitle="The WHO figure is a floor for basic coverage, "
                                       "not a target."),
                        width="stretch")
    with b:
        st.plotly_chart(chart("Doctor density", "Doctors per 10,000",
                              [("Modelled density", R["density"]),
                               ("WHO floor", {y: P.who_norm * P.who_doctor_share
                                              for y in proj_years})], proj_years,
                              subtitle="Driven by supply, not by the falling population: "
                                       "the denominator moves under 2 per cent."),
                        width="stretch")
    a, b = st.columns(2)
    with a:
        st.plotly_chart(chart("Active specialists within the stock", "Doctors",
                              [("Active specialists", R["specialists"]),
                               ("All active doctors", R["active"])], proj_years,
                              subtitle="A transition within the stock, never an addition "
                                       "to it. Broad speciality only."),
                        width="stretch")
    with b:
        ent_years = [y for y in proj_years if y >= 2021]
        st.plotly_chart(stacked("Who joins the register each year",
                                "New registrations",
                                [("Tamil Nadu state counselling", R["domestic"]),
                                 ("Deemed and out of state", R["external"]),
                                 ("Foreign medical graduates", R["fmg"])], ent_years,
                                subtitle="Roughly half of all new registrations come from "
                                         "outside the state counselling system."),
                        width="stretch")
    st.markdown("### Annual series")
    doc_df = pd.DataFrame({
        "Year": proj_years,
        "Population": [R["population"][y] for y in proj_years],
        "Government seats": [R["gov_seats"][y] for y in proj_years],
        "Private seats": [R["pvt_seats"][y] for y in proj_years],
        "PG seats": [R["pg_seats"][y] for y in proj_years],
        "Entrants": [R["entrants"][y] for y in proj_years],
        "Total exits": [R["exits"][y] for y in proj_years],
        "Active doctors": [R["active"][y] for y in proj_years],
        "Per 10,000": [R["density"][y] for y in proj_years],
        "Active specialists": [R["specialists"][y] for y in proj_years],
        "WHO need": [R["who_need"][y] for y in proj_years],
        "Surplus": [R["surplus"][y] for y in proj_years]})
    st.dataframe(doc_df.style.format({c: "{:,.0f}" for c in doc_df.columns
                                      if c not in ("Year", "Per 10,000")}
                                     | {"Per 10,000": "{:,.1f}", "Year": "{:.0f}"}),
                 width="stretch", height=340)
    df_download(doc_df, "Download the doctor series as CSV", "tn_doctors.csv")

# -------------------------------------------------------- seats and pipeline
with tabs[1]:
    a, b = st.columns(2)
    with a:
        st.plotly_chart(chart("MBBS seats, government against private", "Sanctioned seats",
                              [("Government", R["gov_seats"]),
                               ("Private", R["pvt_seats"]),
                               ("Total", R["total_seats"])],
                              [y for y in proj_years if y >= 2015], fill_first=False,
                              subtitle="Government moved 25 seats in four years. All growth "
                                       "since 2021 has been private."),
                        width="stretch")
    with b:
        st.plotly_chart(chart("Postgraduate medical seats", "Sanctioned seats",
                              [("PG seats", R["pg_seats"])],
                              [y for y in proj_years if y >= 2021],
                              subtitle=f"Linear in levels, capped at {P.pg_cap_share:.0%} "
                                       "of MBBS seats. The cap is our own assumption."),
                        width="stretch")
    st.markdown('<div class="note"><b>Why government is a scenario and private is a '
                'curve.</b> Government seats moved twenty five in four years, so they '
                'are modelled as a policy decision, not a trend. Private seats follow a '
                'bounded logistic fitted in levels. Neither uses a compound growth rate: '
                'extrapolating the observed compound rate would give roughly 183,700 '
                'private seats by 2050.</div>', unsafe_allow_html=True)
    st.markdown("### The pipeline, end to end")
    pl = pd.DataFrame({
        "Step": ["Sanctioned seats", "times fill rate", "times completion rate",
                 "lagged by course length", "plus external entrants",
                 "plus foreign graduates", "equals entrants to the register"],
        "Value at " + str(ey): [
            f"{R['total_seats'][ey - P.lag_mbbs]:,.0f} (seat year {ey - P.lag_mbbs})",
            f"x {P.fill_mbbs:.3f}", f"x {P.completion:.2f}",
            f"{P.lag_mbbs} years", f"+ {R['external'][ey]:,.0f}",
            f"+ {R['fmg'][ey]:,.0f}", f"= {R['entrants'][ey]:,.0f}"],
        "Status": ["Observed to 2025, then scenario", "OBSERVED from MGRMU",
                   "NMC benchmark", "4.5 academic years plus internship",
                   "OBSERVED residual, less the deemed overlap",
                   "OBSERVED trend, capped", "Derived"]})
    st.dataframe(pl, width="stretch", hide_index=True)

# -------------------------------------------------------------------- exits
with tabs[2]:
    a, b = st.columns(2)
    with a:
        st.plotly_chart(stacked("Exits from the workforce each year", "Exits",
                                [("Retirement", R["retirements"]),
                                 ("Mortality", R["deaths"]),
                                 ("Migration and out of state", R["migration"])],
                                proj_years,
                                subtitle="Migration dominates. The 2010s expansion cohorts "
                                         "do not reach 60 within this window."),
                        width="stretch")
    with b:
        st.plotly_chart(chart("Entrants against exits", "People per year",
                              [("Entrants", R["entrants"]),
                               ("Total exits", R["exits"])],
                              [y for y in proj_years if y >= 2015],
                              subtitle="The gap between the two is why the stock compounds."),
                        width="stretch")
    cum_r = sum(R["retirements"][y] for y in proj_years if y >= 2026)
    cum_d = sum(R["deaths"][y] for y in proj_years if y >= 2026)
    cum_m = sum(R["migration"][y] for y in proj_years if y >= 2026)
    tot = cum_r + cum_d + cum_m
    x1, x2, x3, x4 = st.columns(4)
    x1.metric("Cumulative retirements", f"{cum_r:,.0f}", f"{cum_r/tot*100:.0f}% of exits")
    x2.metric("Cumulative deaths", f"{cum_d:,.0f}", f"{cum_d/tot*100:.0f}% of exits")
    x3.metric("Cumulative migration", f"{cum_m:,.0f}", f"{cum_m/tot*100:.0f}% of exits")
    x4.metric("Entrants per exit, " + str(ey),
              f"{R['entrants'][ey] / R['exits'][ey]:.1f}x")
    st.markdown('<div class="note"><b>Migration dominates, not retirement.</b> The large '
                'cohorts registered during the 2010s expansion do not reach 60 within '
                'this window, so the retirement wave falls largely outside it. Note that '
                'no exit figure here is observed: all three rest on assumed schedules.</div>',
                unsafe_allow_html=True)

# ------------------------------------------------------------- other cadres
with tabs[3]:
    st.markdown('<div class="warn"><b>These are annual qualifications, not a practising '
                'workforce.</b> No register comparable to the NMC register exists for '
                'these cadres, so there is no stock, no attrition and no benchmark. They '
                'cannot be added to the doctor numbers. Nursing is degree level only: '
                'GNM and ANM appear in no source file.</div>', unsafe_allow_html=True)
    cpath = st.radio("Cadre seat path", ["trend", "frozen"], horizontal=True,
                     format_func=lambda s: ("Fitted trend to 2035, then held"
                                            if s == "trend" else
                                            "Frozen at the last observed level"))
    try:
        C = E.run_cadres(end_year=P.end_year, path=cpath)
    except FileNotFoundError as err:
        st.error(str(err)); st.stop()
    first = C["first_complete"]
    cy = list(range(first, P.end_year + 1))
    groups = sorted({v["group"] for v in C["results"].values()})
    gser = {g: {y: sum(v["output"][y] for v in C["results"].values()
                       if v["group"] == g and v["output"][y] == v["output"][y])
                for y in cy} for g in groups}
    a, b = st.columns(2)
    with a:
        st.plotly_chart(chart("Qualified output by cadre group", "People per year",
                              [(g, gser[g]) for g in groups], cy, fill_first=False,
                              subtitle="Annual qualifications, not practising staff."),
                        width="stretch")
    with b:
        fills = sorted(((v["fill"], k) for k, v in C["results"].items()))
        fig = go.Figure(go.Bar(
            x=[f * 100 for f, _ in fills], y=[k for _, k in fills], orientation="h",
            marker=dict(color=[("#C0392B" if f < 0.70 else "#0F4761") for f, _ in fills],
                        line=dict(color="white", width=0.6)),
            hovertemplate="%{y}: %{x:.0f}%<extra></extra>"))
        fig.update_layout(
            title=dict(text="<b>Seats actually filled, by cadre</b><br>"
                            "<span style='font-size:11.5px;color:#6b7680;font-weight:400'>"
                            "Red marks cadres below 70 per cent, where sanctioned seats "
                            "overstate real output.</span>",
                       font=dict(size=15.5, color="#1d2a33"), x=0, xanchor="left"),
            height=455, plot_bgcolor="white", paper_bgcolor="white",
            font=dict(family="Source Serif 4, serif", size=11, color="#1d2a33"),
            margin=dict(l=190, r=40, t=76, b=46),
            xaxis=dict(title="Per cent of sanctioned seats filled", gridcolor="#eef1f4",
                       linecolor="#c8d1d8", range=[0, 105]),
            yaxis=dict(linecolor="#c8d1d8"))
        st.plotly_chart(fig, width="stretch")
    st.markdown('<div class="note"><b>Fill rates outside medicine are nothing like '
                'medicine.</b> MBBS fills 99 per cent of sanctioned seats. Post Basic '
                'BSc Nursing fills 48 per cent and MSc Nursing 59 per cent. Any plan '
                'built on sanctioned nursing or AYUSH seats overstates supply by a third '
                'or more, which is why this model runs on admissions.</div>',
                unsafe_allow_html=True)
    cdf = pd.DataFrame({"Year": cy} | {n: [v["output"][y] for y in cy]
                                       for n, v in C["results"].items()})
    cdf["Total, these cadres"] = cdf.drop(columns=["Year"]).sum(axis=1)
    st.dataframe(cdf.style.format({c: "{:,.0f}" for c in cdf.columns if c != "Year"}
                                  | {"Year": "{:.0f}"}),
                 width="stretch", height=320)
    df_download(cdf, "Download the cadre series as CSV", "tn_cadres.csv")
    st.markdown("### Cadre parameters, and what is not covered")
    par = pd.DataFrame([{"Cadre": n, "Group": v["group"], "Course years": v["dur"],
                         "Fill rate": f"{v['fill']:.0%}",
                         "Completion": f"{v['completion']:.0%}",
                         "Completion basis": v["comp_basis"],
                         "Pass-out years excluded": ", ".join(str(d) for d in v["dropped"]) or "none"}
                        for n, v in C["results"].items()])
    st.dataframe(par, width="stretch", hide_index=True, height=300)
    st.dataframe(pd.DataFrame(E.NOT_COVERED, columns=["Not covered", "Why"]),
                 width="stretch", hide_index=True)

# -------------------------------------------------------------- sensitivity
with tabs[4]:
    st.markdown("### The two parameters that carry the most weight")
    a, b = st.columns(2)
    with a:
        st.markdown("**Emigration and out-of-state loss**")
        rows = []
        for rate in (0.006, 0.012, 0.020, 0.030):
            rr = run_doc(E.Params(**(P.to_dict() | {"emigration_rate": rate})).to_dict())
            rows.append({"Rate per year": f"{rate:.1%}",
                         "Cumulative loss over 21 years": f"{1-(1-rate)**21:.0%}",
                         f"Active {ey}": f"{rr['active'][ey]:,.0f}",
                         f"Density {ey}": f"{rr['density'][ey]:,.1f}",
                         "Note": "Central case" if abs(rate - 0.012) < 1e-9 else ""})
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
        st.caption("CALIBRATED, not observed. No direct measure of Tamil Nadu doctor "
                   "out-migration exists. The surplus finding survives every rate tested.")
    with b:
        st.markdown("**Age at registration**")
        rows = []
        for age in (22, 23, 24, 25, 26):
            rr = run_doc(E.Params(**(P.to_dict() | {"age_at_registration": age})).to_dict())
            cum = sum(rr["retirements"][y] for y in rr["years"] if y >= 2026)
            rows.append({"Age assumed": age, f"Active {ey}": f"{rr['active'][ey]:,.0f}",
                         f"Retirements in {ey}": f"{rr['retirements'][ey]:,.0f}",
                         "Cumulative retirements": f"{cum:,.0f}",
                         "Note": "Central case" if age == 24 else ""})
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
        st.caption("INFERRED, not observed. Age is never recorded in the register, only "
                   "year of registration. The stock is robust to this; the retirement "
                   "profile is not.")
    st.markdown("### Scenario comparison")
    rows = []
    for gname, (gt, gy) in E.GOV_SCENARIOS.items():
        for pname, (K, r_, t0_) in E.PVT_SCENARIOS.items():
            rr = run_doc(E.Params(**(P.to_dict() | {
                "gov_target": gt, "gov_target_year": gy, "pvt_ceiling": K,
                "pvt_growth_r": r_, "pvt_midpoint": t0_})).to_dict())
            rows.append({"Government path": gname, "Private path": pname,
                         f"Total MBBS seats {ey}": f"{rr['total_seats'][ey]:,.0f}",
                         f"Active doctors {ey}": f"{rr['active'][ey]:,.0f}",
                         f"Density {ey}": f"{rr['density'][ey]:,.1f}"})
    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True, height=350)

# --------------------------------------------------------------- assumptions
with tabs[5]:
    st.markdown("### Every assumption, and what kind it is")
    st.markdown('<div class="note">Three of these are decisions only the State can make '
                'and are marked for confirmation. Two are resolvable with data that '
                'exists but we do not hold. The rest are technical.</div>',
                unsafe_allow_html=True)
    A = [
        ("Private MBBS ceiling", f"{P.pvt_ceiling:,.0f} seats", "ASSUMED, no source",
         "Expert-judgment scenario bound. No arithmetic connects faculty norms, beds or "
         "the applicant pool to this number.", "FOR STATE CONFIRMATION"),
        ("PG seat cap", f"{P.pg_cap_share:.0%} of MBBS seats", "ASSUMED, no source",
         "Our own figure, not from the Andhra Pradesh model. Actual ratio was 56 per "
         "cent in 2025-26. Binds only after 2040.", "FOR STATE CONFIRMATION"),
        ("Government seat target", f"{P.gov_target:,.0f} by {P.gov_target_year or 'n/a'}",
         "POLICY SCENARIO",
         "Built from the observed college distribution, but the level and timing are ours.",
         "FOR STATE CONFIRMATION"),
        ("Emigration rate", f"{P.emigration_rate:.1%} per year, ages "
         f"{P.emigration_age_lo} to {P.emigration_age_hi}", "CALIBRATED, not observed",
         "No direct measure exists. The single most consequential parameter.",
         "Resolvable: NMC removedStatus field"),
        ("Age at registration", f"{P.age_at_registration} years", "INFERRED, not observed",
         "Age is never recorded, only year of registration.",
         "Resolvable: NMC birth dates"),
        ("Mortality by age", "0.08% under 35 to 13% at 85 plus", "ASSUMED",
         "Indian adult mortality adjusted downward for the professional class. The "
         "adjustment is a judgment, not a measured differential.", "Low priority"),
        ("Retirement participation", "100% under 60, then 85, 60, 30, 12, 3", "ASSUMED",
         "Reasoned from retirement at 60 in government service.", "Low priority"),
        ("Population growth beyond 2036", f"falling to {P.pop_growth_2050:.2%}", "ASSUMED",
         "Our own extension. The official series stops at 2036.",
         "Low impact, under 2 per cent on the denominator"),
        ("Foreign graduate ceiling", f"{P.fmg_ceiling:,.0f} per year", "ASSUMED",
         "Roughly double the 2025 level.", "Low priority"),
        ("Completion rate", f"{P.completion:.0%}", "BENCHMARK",
         "NMC benchmark. Validated indirectly on PG: predicted 2,707 against 2,717 "
         "observed.", "Validated"),
        ("External entrants", f"{P.external_residual:,.0f} per year",
         "OBSERVED, assumed to persist",
         "The value is measured across 2021 to 2025 with no trend. Holding it flat for "
         "25 years is the assumption.", "Monitor"),
        ("Seat-source gap", f"{P.gap_forward:,.0f} seats",
         "OBSERVED, assumed to persist",
         f"Drives the deemed overlap of {P.gap_forward * P.fill_mbbs * P.completion:,.0f} "
         "per year from 2031. Inferred to be deemed universities.",
         "Ask the State to confirm"),
    ]
    st.dataframe(pd.DataFrame(A, columns=["Assumption", "Current value", "Status",
                                          "Basis", "Action"]),
                 width="stretch", hide_index=True, height=470)

# ----------------------------------------------------------- data and sources
with tabs[6]:
    st.markdown("### Where every number comes from")
    S = [
        ("TAKEN DIRECTLY", "MBBS seats, government and private, 2015-16 to 2025-26",
         "Selection Committee, UG MBBS BDS Data Sheet.xlsx"),
        ("TAKEN DIRECTLY", "Government college seat distribution, 2025",
         "MGRMU SEAT COUNT STAT 15052026.xlsx, sheet MBBS"),
        ("TAKEN DIRECTLY", "TNMC registrations: MBBS, foreign graduates, PG, 2020 to 2025",
         "TNMC Registrations 2020 to 2025.xlsx"),
        ("TAKEN DIRECTLY", "Seats, admissions and pass-outs for 20 cadres",
         "MGRMU SEAT COUNT STAT and passout COUNT STAT, 26 sheets each"),
        ("ESTIMATED BY US", "Fill rates, all cadres",
         "Admissions over sanctioned seats, pooled across observed years"),
        ("ESTIMATED BY US", "Completion rates, 11 of 20 cadres",
         "Pass-outs over the admissions of the cohort one course length earlier"),
        ("ESTIMATED BY US", "External entrant residual, 5,381 per year",
         "Observed registrations minus modelled domestic output, 2021 to 2025"),
        ("ESTIMATED BY US", "Private logistic r and midpoint",
         "Least squares in levels on 11 observed seat years"),
        ("DERIVED", "Deemed overlap, 806 per year from 2031",
         "850 seats x 0.998 fill x 0.95 completion"),
        ("DERIVED", "WHO doctor requirement, 11.125 per 10,000", "44.5 x 0.25"),
        ("EXTERNAL SOURCE", "Population 2011 to 2036",
         "National Commission on Population and MoHFW (2019), used verbatim"),
        ("EXTERNAL SOURCE", "WHO density norm 44.5 and doctor share one quarter",
         "WHO, Global Strategy on HRH: Workforce 2030 (2016)"),
        ("SCRAPED BY US", "TNMC register, 193,264 doctors, 1927 to 2025",
         "NMC Indian Medical Register, endpoint nmc.org.in/MCIRest, council id 21"),
    ]
    st.dataframe(pd.DataFrame(S, columns=["Class", "Figure", "Source"]),
                 width="stretch", hide_index=True, height=420)
    st.markdown("### Validation, and how independent each check actually is")
    V = [("Register cross-check", "INDEPENDENT",
          "Scraped NMC register against the TNMC workbook. Separate systems, different "
          "bodies. Agreement within 0.17 per cent, exact in 2024."),
         ("PG pipeline", "HELD-OUT TARGET, SHARED PUBLISHER",
          "Predicted 2,707 pass-outs against 2,717 observed. The target was not used to "
          "fit the rates, but both files come from the same university."),
         ("PG coverage factor", "NOT INDEPENDENT",
          "The 1.407 factor was derived by comparing the two sources and is then checked "
          "against one of them. Internal consistency only, not validation."),
         ("Stock identity", "INTERNAL",
          "Active doctors equal last year plus entrants minus exits, to under one doctor "
          "across every transition. Catches accounting errors, not wrong assumptions.")]
    st.dataframe(pd.DataFrame(V, columns=["Check", "Independence", "What it establishes"]),
                 width="stretch", hide_index=True)
    st.markdown("### Current parameter set")
    st.dataframe(pd.DataFrame([{"Parameter": k, "Value": v}
                               for k, v in P.to_dict().items()
                               if not isinstance(v, list)]),
                 width="stretch", hide_index=True, height=300)

st.markdown("---")
st.caption("Method: WHO and World Bank Health Labour Market Framework, following Liu, "
           "Goryakin, Maeda, Bruckner and Scheffler, World Bank Policy Research Working "
           "Paper 7790, Human Resources for Health 2017;15:11. Growth is fitted in "
           "levels throughout, never as a compound annual rate. This is a supply "
           "projection: it describes what the state will produce and retain, not what "
           "the health system will employ.")
