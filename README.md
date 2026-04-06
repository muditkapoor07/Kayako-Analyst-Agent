# M&A Analyst Agent

An AI-powered due diligence assistant that analyses any company's data room and produces investment-ready reports — automatically.

---

## What It Does

Point it at a folder of company files (Excel, CSV, PDF, Word) and it behaves like a junior analyst:

- Reads and understands all the documents
- Calculates key financial metrics (growth rates, margins, churn, retention, etc.)
- Identifies risks and red flags
- Writes a full due diligence report
- Answers follow-up questions about the deal

No manual work. No copy-pasting numbers. Just point it at a folder and ask.

---

## Two Ways to Use It

### 1. Terminal Agent (private, local)
Your personal analyst. Run it from the command line on any deal folder on your machine.

```bash
python ma_agent.py --deal "C:/deals/AnyCompany"
```

### 2. Web App (shareable, browser)
Upload files via browser, get analysis, download report. Share the URL with anyone.

```bash
streamlit run app.py
```

---

## Key Features

- **Works on any company** — drop any data room folder or upload any files, it figures out the rest
- **Reads any file type** — Excel, CSV, Word, PDF, plain text, JSON
- **Calculates automatically** — CAGR, retention curves, concentration ratios, EBITDA, margins
- **Writes reports** — full markdown analysis saved to the deal folder or downloadable from browser
- **Remembers past deals** — recall any previously analysed company by name (terminal agent)
- **Ask follow-up questions** — interactive Q&A after the initial analysis
- **No setup per deal** — same command, any company, every time

---

## How It Works

```
You run:   python ma_agent.py --deal "path/to/company/folder"

Agent:     1. Discovers all files in the folder
           2. Reads the relevant ones (CIM, financials, retention, etc.)
           3. Runs calculations in Python
           4. Writes a full due diligence report
           5. Saves it to folder/outputs/ma_analysis.md
           6. Waits for your follow-up questions
```

The agent decides what to read, what to calculate, and what to flag — you just give it a folder.

---

## Project Structure

```
/Kayako-Analyst-Agent
 │
 ├── ma_agent.py          # Terminal agent — run locally on any deal folder
 ├── app.py               # Web app — upload files, get analysis in browser
 ├── requirements.txt     # Python dependencies
 ├── README.md            # This file
 ├── claude.md            # Agent instructions
 │
 ├── Sample Deal: Kayako
 │   ├── Kayako_Confidential_Information_Memorandum.xlsx
 │   ├── Kayako_Customer_Retention_Analysis.xlsx
 │   ├── Kayako_Churn_Reduction_Plan.xlsx
 │   ├── Kayako_Top20_Revenue_Concentration.xlsx
 │   ├── Kayako_Sales_Forecast_Pipeline.xlsx
 │   └── Kayako_One_PAGER.xlsx
 │
 └── Sample Deal: ZenDeskly
     ├── ZenDeskly_Confidential_Information_Memorandum.xlsx
     ├── ZenDeskly_Customer_Retention_Analysis.xlsx
     ├── ZenDeskly_Churn_Reduction_Plan.xlsx
     ├── ZenDeskly_Top20_Revenue_Concentration.xlsx
     ├── ZenDeskly_Sales_Forecast_Pipeline.xlsx
     └── ZenDeskly_One_PAGER.xlsx
```

---

## Installation

**Step 1 — Clone the repo**
```bash
git clone https://github.com/muditkapoor07/Kayako-Analyst-Agent.git
cd Kayako-Analyst-Agent
```

**Step 2 — Install dependencies**
```bash
pip install -r requirements.txt
```

**Step 3 — Add your API key**

Get a free key at [console.anthropic.com](https://console.anthropic.com) → API Keys → Create Key

```bash
# Windows
setx ANTHROPIC_API_KEY "sk-ant-your-key-here"

# Mac / Linux
export ANTHROPIC_API_KEY="sk-ant-your-key-here"
```

---

## Terminal Agent Usage

### Analyse a deal folder
```bash
python ma_agent.py --deal "C:/deals/Kayako"
python ma_agent.py --deal "C:/deals/ZenDeskly"
```

### Focus on a specific area
```bash
python ma_agent.py --deal "C:/deals/Kayako" --task "focus on churn risk only"
```

### List all past deals from memory
```bash
python ma_agent.py --deals
```

### Recall a past deal
```bash
python ma_agent.py --recall "Kayako"
```

---

## Web App Usage

**Run locally:**
```bash
streamlit run app.py
```
Opens at `http://localhost:8501`

**Deploy to Streamlit Cloud (free):**
1. Go to [share.streamlit.io](https://share.streamlit.io) → connect this GitHub repo
2. Select `app.py` as the main file
3. Go to **Settings → Secrets** and add:
```toml
ANTHROPIC_API_KEY = "sk-ant-your-key-here"
```
4. Share the URL — anyone can upload files and get analysis, no key needed on their end

**How the web app works:**
1. Upload any company's data room files (Excel, CSV, Word, TXT)
2. Click **Run Analysis** — agent reads files and calculates autonomously
3. Download the full report as a markdown file
4. Ask follow-up questions in the chat below

---

## Example

```bash
python ma_agent.py --deal "C:/Kayako-Analyst-Agent"
```

**What happens:**
1. Agent lists all files in the folder
2. Reads CIM, retention analysis, churn plan, revenue concentration
3. Calculates 5-year revenue CAGR, ARR growth, cohort retention curves
4. Flags top risks (customer concentration, churn trend, EBITDA quality)
5. Saves full report to `outputs/ma_analysis.md`
6. You ask: *"What's the biggest acquisition risk?"* → instant answer

**Report covers:**
- Revenue & ARR trends
- Gross margin and EBITDA expansion
- Customer count and ARPU growth
- Cohort retention analysis
- Top customer revenue concentration
- Sales pipeline and forecast
- Key risks and diligence questions for management

---

## Supported File Types

| Format | Extensions |
|--------|-----------|
| Excel  | `.xlsx`, `.xls` |
| CSV    | `.csv` |
| Word   | `.docx` |
| Text   | `.txt`, `.md` |
| JSON   | `.json` |

---

## Sample Data Rooms Included

Two artificial data rooms are included for testing:

| Company | Files |
|---------|-------|
| **Kayako** | CIM, Retention, Churn Plan, Revenue Concentration, Sales Forecast, One Pager |
| **ZenDeskly** | CIM, Retention, Churn Plan, Revenue Concentration, Sales Forecast, One Pager |

All data is **artificial and for demonstration purposes only**.

---

## Requirements

- Python 3.9+
- An Anthropic API key ([get one free](https://console.anthropic.com))
- Dependencies in `requirements.txt`: `anthropic`, `streamlit`, `openpyxl`, `httpx`

---

## Who Is This For?

- **M&A analysts** speeding up initial due diligence
- **Investors** screening deals before deep dives
- **Associates** generating first-pass analysis on new targets
- **Anyone** who receives data room files and needs fast answers

---

## Notes

- The agent never modifies your original files — it only reads them
- Terminal agent reports saved to `outputs/` inside your deal folder
- Deal memory stored locally at `~/.ma_analyst/deals/`
- Works on Windows, Mac, and Linux
- All sample data is artificial
