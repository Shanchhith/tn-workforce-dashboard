"""Extract PG registrations by speciality from the TNMC workbook.
Writes an aggregate CSV (no personal data) the engine reads."""
import openpyxl, re, csv, collections, sys
SRC = "new data/Medical Council/TNMC Registrations 2020 to 2025.xlsx"
OUT = "dashboard/data/pg_registrations_by_speciality.csv"
YEARS = list(range(2020, 2026))
# The four sheets that make up the model's broad-speciality PG total. Verified:
# their sum equals REG_PG_OBS exactly in all six years. Superspecialty (DM,
# M.Ch, DNB SS) and PG Diploma are outside it, by design.
SHEETS = {"MD TNMC": "MD", "MS TNMC ": "MS", "DNB MedPG TNMC": "DNB", "DNB Surg PG TNMC": "DNB"}
FIX = {
    "Obstetrics and Gynecology": "Obstetrics and Gynaecology",
    "Otorhinolaryngology": "ENT (Otorhinolaryngology)",
    "Orthopedics": "Orthopaedics",
    "Dermatology, Venerology and Leprosy": "Dermatology, Venereology and Leprosy",
    "Radio Diagnosis": "Radiodiagnosis",
    "(Respiratory Diseases)": "Respiratory Medicine",
    "Tuberculosis and Respiratory Diseases": "Respiratory Medicine",
    "(Medical Genetics)": "Medical Genetics",
    "Anaesthesiology": "Anesthesiology",
}
def canon(name):
    n = " ".join(str(name).split())
    n = re.sub(r"^(M\.?D\.?|M\.?S\.?|DNB)\s*", "", n, flags=re.I).strip(" .-")
    n = n.replace("&", "and")
    return FIX.get(n, n)

def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    rows = collections.defaultdict(lambda: {y: [0.0, 0.0] for y in YEARS})
    awards = collections.defaultdict(set)
    for sh, tag in SHEETS.items():
        ws = wb[sh]
        for r in ws.iter_rows(min_row=5, values_only=True):
            nm = r[1]
            if not nm or not str(nm).strip():
                continue
            s = " ".join(str(nm).split())
            if s.lower().startswith(("total", "s.no")):
                continue
            c = canon(s); awards[c].add(tag)
            for i, y in enumerate(YEARS):
                m, f = r[2 + 2 * i], r[3 + 2 * i]
                rows[c][y][0] += float(m) if isinstance(m, (int, float)) else 0.0
                rows[c][y][1] += float(f) if isinstance(f, (int, float)) else 0.0
    with open(OUT, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["speciality", "awards", "year", "male", "female", "total"])
        for c in sorted(rows):
            for y in YEARS:
                m, f = rows[c][y]
                w.writerow([c, "+".join(sorted(awards[c])), y, int(m), int(f), int(m + f)])
    grand = sum(sum(v[y]) for v in rows.values() for y in YEARS)
    print(f"wrote {OUT}: {len(rows)} specialities, {grand:,.0f} registrations 2020-2025")
    return grand

if __name__ == "__main__":
    g = main()
    sys.exit(0 if abs(g - 28809) < 1 else 1)
