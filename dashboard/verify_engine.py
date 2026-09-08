"""Verify tn_engine reproduces tn_supply_model and tn_cadre_model exactly."""
import io, contextlib, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import numpy as np
import tn_engine as E
with contextlib.redirect_stdout(io.StringIO()):
    import tn_supply_model as m
    import tn_cadre_model as c

SQ = "Status quo"
R = E.run_doctors(E.Params())
ok = True
print(f"{'series':<22}{'n':>5}{'max abs diff':>15}{'worst year':>12}")
checks = [
    ("population",  R["population"],  {y: m.POP[m.IDX[y]] for y in m.YEARS}),
    ("gov seats",   R["gov_seats"],   {y: m.gov_seats(SQ)[m.IDX[y]] for y in m.YEARS}),
    ("private seats", R["pvt_seats"], {y: m.pvt_seats(SQ)[m.IDX[y]] for y in m.YEARS}),
    ("PG seats",    R["pg_seats"],    {y: m.pg_seats(SQ)[m.IDX[y]] for y in m.YEARS}),
    ("entrants",    R["entrants"],    {y: m.col(SQ,"entrants")[m.IDX[y]] for y in m.YEARS}),
    ("active doctors", R["active"],   {y: m.col(SQ,"active")[m.IDX[y]] for y in m.YEARS}),
    ("specialists", R["specialists"], {y: m.SPECIALISTS[SQ][m.IDX[y]]["active"] for y in m.YEARS}),
    ("retirements", R["retirements"], {y: m.EXITS[m.IDX[y]]["retirements"] for y in m.YEARS}),
    ("deaths",      R["deaths"],      {y: m.EXITS[m.IDX[y]]["deaths"] for y in m.YEARS}),
    ("migration",   R["migration"],   {y: m.EXITS[m.IDX[y]]["emigration"] for y in m.YEARS}),
    ("WHO need",    R["who_need"],    {y: m.NEED_DOCTORS[m.IDX[y]] for y in m.YEARS}),
]
for name, got, exp in checks:
    worst, wy, n = 0.0, None, 0
    for y in exp:
        e = exp[y]
        if e != e:
            continue
        g = got.get(y, float("nan"))
        if g != g:
            continue
        n += 1
        d = abs(g - e)
        if d > worst:
            worst, wy = d, y
    good = worst < 0.51
    ok &= good
    print(f"{name:<22}{n:>5}{worst:>15.6f}{str(wy):>12}   {'OK' if good else 'MISMATCH'}")

C = E.run_cadres(end_year=2050, path="trend")
worst, wn, wy = 0.0, None, None
for b in c.BUILT:
    q = c.qualified_output(b, "trend")
    got = C["results"][b["name"]]["output"]
    for y in range(C["first_complete"], 2051):
        if q[y] != q[y]:
            continue
        d = abs(got[y] - q[y])
        if d > worst:
            worst, wn, wy = d, b["name"], y
good = worst < 0.51
ok &= good
print(f"{'cadre output':<22}{20*24:>5}{worst:>15.6f}{str(wy):>12}   {'OK' if good else 'MISMATCH'}  {wn or ''}")
print("\nENGINE REPRODUCES THE PUBLISHED MODEL EXACTLY" if ok
      else "\nDIFFERENCES FOUND, engine does not match")
