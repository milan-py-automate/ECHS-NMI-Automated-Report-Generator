# ECHS NMI Report Studio

A corporate-style, stateless Excel reporting web application with **HTML + CSS + JavaScript** and **Python/Flask**. Deploy from GitHub to **Vercel**.

## Repository layout

```
app.py
engine.py
requirements.txt
vercel.json
templates/index.html
static/style.css
static/app.js
```

## Deploy using GitHub + Vercel

1. Create a new GitHub repository and upload **all** files/folders in this project, preserving the structure.
2. Open the [Vercel dashboard](https://vercel.com/new), choose **Import Git Repository**, and select this repository.
3. The root directory is `/`; Vercel detects the Python Flask application from `app.py` and `requirements.txt`. Framework preset can be **Other**. Do not set a static-output directory.
4. Recommended: add an environment variable `APP_PASSWORD` (choose a strong password) in Vercel Project Settings → Environment Variables. This adds password verification for `/api/preview` and `/api/generate`; share the URL only with authorized users. **Without this variable, the app is publicly accessible.** The UI routes and the health route remain public.
5. Click **Deploy**. Test `/api/health`, then upload your files and try the workflow.

> Vercel's Python runtime supports Flask auto-detection and doesn't need legacy `builds` / `routes` rules. `vercel.json` declares an empty framework preset. `app.py` is the application entry point.

## Local test

```bash
python -m venv .venv
# Activate your virtual environment
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000

## Processing rules

- Upload multiple portal-generated Raw files (`.xls`, `.xlsx`, `.xlsb`, `.csv`); header automatically detected in the first 35 rows. **Part 1 retains every valid row**, even when Claim IDs repeat.
- Raw Claim IDs are normalized to numbers (only positive integers with <=15 digits are accepted). Dates are real Excel date values formatted `DD-MM-YYYY`.
- For master reconciliation, each Claim ID is treated as **active** if present in any uploaded Raw file. Older master rows remain in place, including duplicates.
- All matching existing master rows become `Pending`; all old IDs absent from Raw become `Done`.
- A claim ID absent from the previous master is appended **once** (choosing the latest NMI Date among Raw duplicates).
- Existing master columns **A–J and L–M** are kept unchanged; **K** status and **N/O** formulas are updated. New entries get hospital name, sponsor `ECHS`, IP/OP, Claim ID, patient name, claim amount, NMI Date and latest user remarks. Column **D stays empty** for new entries by design.
- `Settings!B2` contains the report reference date for the Days to Pending and Aging formulas.
- If any old `Pending` records will change to `Done`, the UI requires an explicit checkbox confirmation.
- The report downloads a ZIP containing **Part 1 Combined Raw** and **Part 2 Updated Master**.

## Limitations and security

- **Critical:** upload *all* relevant hospital/region Raw files. Missing files can wrongly change a previously Pending claim to Done.
- **Vercel Function limits:** requests and responses are limited to about **4.5 MB**. The UI caps the combined file upload at 3.9 MB and the backend caps generated ZIP at 4.2 MB. Large workloads require a different backend or direct object storage.
- Reports are processed in memory for each request, not stored in a database or permanent filesystem. They pass through your Vercel Function. This **does not** make the app a certified medical-data processor.
- Do not put patient information, report files, passwords, or secrets into GitHub. This project contains no real patient data.
- Use Vercel authentication controls or protected deployment / network rules for sensitive patient claims. `APP_PASSWORD` is only a basic shared-secret gate and is not a full user-authentication system. Consult your organization's applicable data policies before processing patient information in hosted services.
- An XLSM master is supported and macros are retained in the XLSM result. Files with complex unsupported Excel features should be validated in Excel afterward.
- Opening the workbook in Microsoft Excel recalculates formulas. Downloaded files may initially lack formula caches in non-Excel viewers.
- The app intentionally does not store report history or upload files across page reloads. Users must keep each downloaded ZIP locally.

## Notes

- Dependency packages specified in `requirements.txt`.
- Python 3.12 is a reasonable default for local testing and deployment.
- Frontend uses Tailwind CSS 2.2.19 utility stylesheet (CDN) plus bespoke corporate CSS (no JS framework) for a dependable deploy, and supports responsive screens, drag/drop, preview, danger confirmation, error messages, and download.
