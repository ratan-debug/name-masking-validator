# Name Masking Validator

A local Flask web app for uploading an Excel workbook, validating masked names against original names, and downloading an updated workbook with match analysis.

## Folder Structure

```text
.
|-- app.py
|-- processor.py
|-- requirements.txt
|-- run_workspace.ps1
|-- templates/
|   `-- index.html
|-- static/
|   |-- app.js
|   `-- styles.css
|-- storage/
|   |-- uploads/
|   `-- outputs/
`-- logs/
    `-- app.log
```

## Input Excel Format

The workbook should contain these columns:

| Original Name | Masked Name |
| --- | --- |
| SUJAN MISTRI | S***N M****I |
| MADHU MISHRA | K*C B******N P*T L*D |
| PRATIK DNYANDEV JADHAV | D******V J****V |

If the exact column names are missing, the app falls back to the first two columns.

## Output Columns

The processed workbook adds:

| Match Status | Confidence % | Match Details | Final Output |
| --- | --- | --- | --- |
| FULL_MATCH | 0-100 | star count, first/last visible letters, token matching detail | reconstructed in masked token order |
| PARTIAL_MATCH | 0-100 | strong subset/token alignment detail | best readable reconstruction from original |
| NO_MATCH | 0-100 | weak/no alignment detail | masked value or weak candidate, depending on available evidence |

## Updated Matching Behavior

The masked value is treated as the priority pattern.

- Confirmed full matches reconstruct readable output using the masked token order.
- A single masked token can map to adjacent original words when the pattern supports it.
- `HUKUM DEV PASWAN` plus `H******V P****N` becomes `HUKUMDEV PASWAN`.
- `MOHAN R` plus `R.****N` becomes `R MOHAN`, not `MOHAN R`.
- `KARIYAMMA .` plus `K*******A` becomes `KARIYAMMA`, with punctuation removed.
- Weak cases are marked `NO_MATCH` instead of overusing `PARTIAL_MATCH`.
- Strong subset cases can still be `PARTIAL_MATCH`, for example `P****K` against `PRATIK DNYANDEV JADHAV` gives `PRATIK`.

The processor compares:

- number of masked stars
- first visible masked character
- last visible masked character
- masked word count vs original word count
- token-level pattern matching
- visible character order
- RapidFuzz fuzzy similarity

## Installation

Open PowerShell in this folder.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

Port 70 may require an elevated PowerShell window on Windows.

```powershell
python app.py
```

Then open:

```text
http://localhost:70
```

Alternative Flask command:

```powershell
flask --app app run --host=0.0.0.0 --port=70
```

## Free Live Hosting

The app is ready for Render free web service hosting.

1. Push this folder to a GitHub repository.
2. Open Render and create a new Web Service from that repository.
3. Use:

```text
Build Command: pip install -r requirements.txt
Start Command: gunicorn app:app
```

Render will provide a public URL like:

```text
https://name-masking-validator.onrender.com
```

Free Render services may sleep after inactivity, so the first request can take extra time.

## Notes

- Supported uploads: `.xlsx`, `.xlsm`
- Maximum upload size: 100 MB
- Processed files are saved under `storage/outputs`
- Uploads are saved under `storage/uploads`
- Logs are written to `logs/app.log`
- The app runs fully offline after Python packages are installed
