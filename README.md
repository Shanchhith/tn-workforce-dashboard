# Tamil Nadu Health Workforce Projections

An interactive supply side projection of Tamil Nadu's health workforce to 2050,
covering doctors, specialists and twenty other cadres.

Centre for Management of Health Services, Indian Institute of Management Ahmedabad.

## Live dashboard

Deployed on Streamlit Community Cloud. Every assumption in the sidebar is
editable and every figure recomputes from it.

## What it projects

**Doctors.** A stock and flow model with an explicit education pipeline: seats,
then fill rate, then completion rate, then a course length lag, then the medical
register. The stock is rolled forward from 201,908 registration records with age
specific mortality, migration and a retirement participation curve.

**Twenty other cadres.** Nursing at degree level, dental, pharmacy,
physiotherapy and occupational therapy, the five AYUSH streams, and allied
health at postgraduate level, on the same pipeline chain.

## Method

WHO and World Bank Health Labour Market Framework, following Liu, Goryakin,
Maeda, Bruckner and Scheffler, *Global Health Workforce Labor Market Projections
for 2030*, World Bank Policy Research Working Paper 7790, published in *Human
Resources for Health* 2017;15:11.

Growth is fitted **in levels throughout, never as a compound annual rate**.
Extrapolating the observed compound rate for private medical seats would give
roughly 183,700 seats by 2050, which is not credible. Government seats are
modelled as a policy scenario rather than a trend, because government sets them
directly and they moved twenty five seats in four years.

## Reading it correctly

**The doctor figures are a practising workforce. The other cadres are not.**
No register comparable to the National Medical Commission register exists for
nursing or the allied cadres, so for those there is no stock, no attrition and
no benchmark. Those figures are annual qualifications only and must not be added
to the doctor numbers.

**This is a supply projection.** It describes what the state will produce and
retain, not what the health system will employ.

**Nursing is degree level only.** GNM and ANM sit with the Tamil Nadu Nurses and
Midwives Council rather than the university and appear in no source file, so
total nursing output is understated here.

**Three assumptions are marked FOR STATE CONFIRMATION** on the Assumptions tab:
the private seat ceilings, the postgraduate cap, and the government seat
targets. They are our judgment, not the State's position.

## What is published here, and what is not

This repository contains **no personal data and no raw source workbooks**.

| File | Contents |
|---|---|
| `dashboard/data/register_by_year.csv` | Registrations per year, 1927 to 2025. Ninety nine numbers. |
| `dashboard/data/cadre_series.json` | Sanctioned seats, admissions and pass-outs per cadre and year. |
| `TN_Health_Workforce_Projections_2050.xlsx` | The published workbook, 24 sheets, 1,523 live formulas. |
| `tn_projection_2050.csv`, `tn_cadre_projection_2050.csv` | The annual output series. |

**Deliberately excluded.** The scraped medical register held doctor names,
father names and registration numbers for 193,264 individuals. It is not in this
repository and not in its history. The model reduces it to a count per year and
produces identical results from that aggregate, verified series by series. The
source workbooks supplied by the Selection Committee, the Dr. M.G.R. Medical
University and the Medical Council are likewise not published; only the derived
series the model uses are.

`tn_supply_model.py`, `tn_cadre_model.py` and `dashboard/verify_engine.py` are
included so the method can be read and audited, but they read the raw sources
and therefore will not run from this repository alone. The dashboard does not
depend on them.

## Running locally

    pip install -r requirements.txt
    streamlit run dashboard/tn_dashboard.py

## Sources

Seat and pass-out data from the Tamil Nadu Medical Selection Committee and the
Tamil Nadu Dr. M.G.R. Medical University. Registrations from the Tamil Nadu
Medical Council and the National Medical Commission Indian Medical Register.
Population from the National Commission on Population and MoHFW, *Population
Projections for India and States 2011-2036* (2019). Density norm from WHO,
*Global Strategy on Human Resources for Health: Workforce 2030* (2016).

Full provenance, including which figures are observed, fitted, derived or
assumed, is on the Data and sources and Assumptions tabs of the dashboard.
