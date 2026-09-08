# Tamil Nadu Health Workforce Projection Dashboard

## Running it

Double-click **`Run Dashboard.command`** in the project folder. It opens in your
browser at `http://localhost:8520`. Close the terminal window to stop it.

Or from a terminal:

    cd "Tamil Nadu Projections"
    streamlit run dashboard/tn_dashboard.py

## What is in it

| File | Purpose |
|---|---|
| `tn_engine.py` | The projection engine. A pure, parameterised version of the published model with no side effects. Every assumption is an argument. |
| `tn_dashboard.py` | The interface. |
| `verify_engine.py` | Proves the engine reproduces the published model exactly. Run it after any change. |

## Verification

    python3 dashboard/verify_engine.py

It compares eleven doctor series and all 480 cadre cells against
`tn_supply_model.py` and `tn_cadre_model.py`. Every one must read 0.000000.

## How to use it

The sidebar opens on **published defaults**, which reproduce
`TN_Health_Workforce_Projections_2050.xlsx` exactly. A banner confirms this.

- **Project to year** moves the horizon anywhere from 2030 to 2070.
- **Scenario** dropdowns pick a government seat path and a private ceiling.
- **Edit individual assumptions** turns on every parameter as an editable field.
- **Reset everything to published defaults** restores the workbook values.

The banner turns amber the moment anything differs from the published values, so
a modified run can never be mistaken for the official one.

Charts are interactive: hover for values, drag to zoom, double-click to reset,
and use the camera icon to save a PNG. Every table has a CSV download.

## A caution worth repeating

The **Other cadres** tab reports annual qualifications, not a practising
workforce. No register comparable to the NMC register exists for those cadres,
so there is no stock, no attrition and no benchmark. Those figures must not be
added to the doctor numbers. Nursing is degree level only, because GNM and ANM
appear in no source file.

## If files show as 0 bytes

Dropbox has the project folder set to online only. Right-click the folder in
Finder and choose **Make Available Offline**. The dashboard reads the register
and the MGRMU workbooks at runtime and will show a clear error if they are not
materialised.

---

## Putting it online

The dashboard is a normal Streamlit app, so any of these work. They are listed
cheapest and simplest first.

### 1. Streamlit Community Cloud (free, easiest)

Needs a public or private GitHub repository. Push `dashboard/`, the two model
scripts, and the data the engine reads, then point
`share.streamlit.io` at `dashboard/tn_dashboard.py`.

Add a `requirements.txt` at the repository root:

    streamlit
    numpy
    pandas
    plotly
    openpyxl

Caution: the register CSV is 15 MB and the MGRMU workbooks are about 1 MB. That
is fine for GitHub, but the whole repository becomes readable by anyone you
share the link with. Use a **private** repository, and note that free Community
Cloud apps are public by default unless you restrict viewers by email.

### 2. Share it on your own network, no hosting at all

If the State team is on the same network, or you are presenting from your laptop:

    streamlit run dashboard/tn_dashboard.py --server.address 0.0.0.0

Streamlit prints a Network URL such as `http://10.5.0.89:8520`. Anyone on the
same network can open it. Nothing leaves your machine. This is the safest option
for a meeting and needs no setup.

### 3. A private server, for a permanent internal link

Any small cloud instance, or an IIMA server. Install the requirements, run the
app behind nginx with a password, and keep the data files on the server rather
than in a repository. This is the right answer if the State is to have standing
access.

### 4. Not recommended: ngrok or similar tunnels

They expose your laptop to the public internet on a URL that is easy to share
further than you intend. Given the data is unpublished state workforce material,
avoid it.

### Before publishing anywhere

Two things to settle first. The workbook and this dashboard contain unpublished
Tamil Nadu government data, so confirm with the State what may go on a public
host. And the three assumptions marked FOR STATE CONFIRMATION are still ours,
so a published dashboard should either carry that label prominently, as it
currently does on the Assumptions tab, or wait until the State has signed off.
