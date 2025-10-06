<p align="center">
  <img src="https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=google&logoColor=white" alt="Google Apps Script">
  <img src="https://img.shields.io/badge/JavaScript-ES5-F7DF1E?logo=javascript&logoColor=000" alt="JavaScript">
  <img src="https://img.shields.io/badge/Google%20Sheets-34A853?logo=google-sheets&logoColor=white" alt="Google Sheets">
  <img src="https://img.shields.io/badge/Regex-6f42c1" alt="Regex">
</p>

<h1 align="center">Planning Portal Scraper (Google Sheets + Apps Script)</h1>

<p align="center">
  One-click scraping of UK council planning application details into Google Sheets.<br>
  Built to turn 10k+ raw planning links into a clean, analysis-ready dataset.
</p>

<p align="center">
  <b>✅ Final code and example datasets are included in this repository.</b>
</p>

---

## ✨ What this does

- Reads application **URLs** (with `keyVal=...`) from **Column B** of a Google Sheet  
- Visits **Summary**, **Contacts** (and **Documents** only if needed)  
- Extracts and writes to **Columns J → O**:
  - **Application Received**
  - **Decision Issued Date**
  - **Documents Associated** &nbsp;*(e.g. “There are 14 documents associated with this application.”)*
  - **Agent Name**, **Agent Email**, **Agent Telephone**
- Adds runtime metadata in **Columns P → Q**:
  - **Status** (`OK`, `RETRY 1m (...)`, `ERROR ...`)
  - **Next Try After** (timestamp for quick backoff)
- Runs in **continuous passes** and can **auto-restart until you cancel**

---

## 🧭 Why I built it

A friend needed tailored datasets for research and analysis to support planning permission applications.  
He had **raw exports** (URL, address, subject, basic meta) downloaded from a **premium property planning research service**, but the exports were missing key fields and doing it by hand for **10,000+** records was impossible.

This Apps Script runs **inside Google Sheets**, scrapes the missing fields from council portals (Idox “Public Access” style), and builds a clean dataset.

---

## 🧰 Tools & Technologies

- **Google Apps Script** (serverless JS in Google Workspace)  
  `UrlFetchApp`, `SpreadsheetApp`, `ScriptProperties`, `LockService`, time-based **Triggers**
- **JavaScript (ES5)** + **Regular Expressions** for resilient HTML parsing
- **Google Sheets** as the UI & datastore
- **GitHub** for versioning and portfolio docs

> `IMPORTXML` was not reliable because many council pages are dynamic or blocked; server-side fetch + regex parsing is more dependable.

---

## 📊 Output Schema (J → Q)

| Col | Header                  | Description |
|-----|-------------------------|-------------|
| J   | Application Received    | From Summary tab |
| K   | Decision Issued Date    | From Summary tab |
| L   | Documents Associated    | Full sentence; singular/plural handled; tries Summary first, then Documents |
| M   | Agent Name              | From Contacts; heading text or labeled cell |
| N   | Agent Email             | Labeled field or email pattern |
| O   | Agent Telephone         | Handles “Telephone number” label, `tel:` links, UK formats |
| P   | Status                  | `OK`, `RETRY 1m (…)`, or `ERROR …` |
| Q   | Next Try After          | Timestamp used to skip a row temporarily |

---

## 🏗️ How it works

1. **Input**  
   Paste application URLs into **Column B** (`B2` downward). Each link contains a unique `keyVal`.

2. **Per-row flow**  
   - Fetch **Summary** → extract dates and attempt the documents sentence  
   - If documents sentence is missing → fetch **Documents** and extract  
   - Fetch **Contacts** → parse Agent name/email/telephone  
   - Write results to **J–O**; mark **Status = OK**

3. **Scale & resilience**  
   - Each Apps Script run works ≈ **5.9 minutes** (Google’s limit)  
   - Short retries with small backoff; rows that still fail are marked **RETRY 1m** and **skipped immediately** (no delay for the rest)  
   - A **pointer** persists across runs so the next pass resumes exactly where it stopped

---

## 🚀 Quick Start

1. **Create a Google Sheet** and put your URLs in **Column B** from **B2**.
2. **Extensions → Apps Script** → paste the contents of `Code.gs` from this repo → **Save**.
3. Reload the sheet. A **Planning Scraper** menu appears.
4. Choose one of:
   - **Run now (continuous pass)** — processes rows immediately for ~6 min  
   - **Start “Run until canceled”** — auto-restarts every minute until you stop it  
   - **Scrape active row** — only the currently selected row
5. Watch **J–Q** fill in. Rows with temporary issues show `RETRY 1m (...)` and will be picked up on the next pass.

---

> **Data provenance**: The raw dataset (URLs, addresses, subjects, basic metadata) was provided by my friend from a **premium property planning research service**.  
> This script **enriches** that data with fields scraped from council portals to create a tailored dataset.  
> **Both the final code and example datasets are attached to this GitHub repository.**

---

## ⚙️ Configuration

Inside `Code.gs`, tweak these to taste:

```js
const OUTPUT_START_COLUMN = 10;   // Column J
const TIME_BUDGET_MS = 355000;    // ~5.9 min per pass, stays under Google limit
const MAX_ROWS_PER_RUN = 1500;    // Safety cap per pass
````

* Want to **refresh** already filled rows? Remove the “skip if J has data” check in the batch loop.
* Output elsewhere? Change `OUTPUT_START_COLUMN`.
* Seeing throttling? Lower `TIME_BUDGET_MS` or add a tiny `Utilities.sleep()` between rows.

---

## 🧩 Notes & Limitations

* **Be respectful** of site terms and local laws. This is for **research and analysis** use.
* Council HTML can change; update regexes if a portal’s markup is customized:

  * `extractLabel_` (Summary fields)
  * `extractAgentDetails_` (Contacts/Agent block)
  * `extractDocsSentenceFromOnePage_` (documents count)
* Apps Script enforces quotas. The “run until canceled” mode chunks work into safe passes, auto-continuing in the background.

---

## 📚 What I learned

* Why spreadsheet XPath (`IMPORTXML`) often fails on modern, JS-heavy portals, and how **server-side fetch + regex** solves it
* Building **resume-safe**, **idempotent** workflows with **Script Properties**, **Locks**, and **Triggers**
* Crafting robust parsers for Idox variants (`<dt>/<dd>`, `<th>/<td>`, plain heading text)
* Designing a simple **in-Sheets UX**: custom menu, status columns, 1-minute backoff, “run until canceled”

---

## 📸 Screenshots 

Some Screenshots Below:

<img src="https://github.com/nabirudd/idox-planning-enricher/blob/main/Images/Code%20JS.png?raw=true" alt="Setup" height="500" hspace="70">
<img src="https://github.com/nabirudd/idox-planning-enricher/blob/main/Images/Functions%20Button.png?raw=true" alt="Setup" height="500" hspace="70">
