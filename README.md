# M&A Analyst Agent

An AI-powered due diligence assistant that analyses any company's data room and produces investment-ready reports — automatically.

---

## What It Does

Point it at a folder of company files (Excel, CSV, PDF, Word) and it behaves like a junior analyst:

- Reads and understands all the documents
- Calculates key financial metrics (growth rates, margins, churn, etc.)
- Identifies risks and red flags
- Writes a full due diligence report
- Answers follow-up questions about the deal

No manual work. No copy-pasting numbers. Just point it at a folder and ask.

---

## Key Features

- **Works on any company** — drop any data room folder, it figures out the rest
- **Reads any file type** — Excel, CSV, Word, PDF, plain text
- **Calculates automatically** — CAGR, retention curves, concentration ratios, EBITDA
- **Writes reports** — saves a full markdown analysis to the deal folder
- **Remembers past deals** — recall any previously analysed company by name
- **Ask follow-up questions** — interactive Q&A after the initial analysis
- **No setup per deal** — same command, any company, every time

---

## How It Works

```
You run:   python ma_agent.py --deal "path/to/company/folder"

Agent:     1. Discovers all files in the folder
           2. Reads the relevant ones (CIM, financials, retention, etc.)
           3. Runs calculations in Python
           4. Writes a due diligence report
           5. Saves it to folder/outputs/ma_analysis.md
           6. Waits for your follow-up questions
```

The agent decides what to read, what to calculate, and what to flag — you just give it a folder.

---

## Project Structure

```
/Kayako-Due-Deligence
 ├── ma_agent.py          # The M&A analyst agent (use this)
 ├── app.py               # Streamlit chat interface (browser demo)
 ├── requirements.txt     # Python dependencies
 ├── claude.md            # Agent instructions
 │
 ├── Kayako_Confidential_Information_Memorandum.xlsx
 ├── Kayako_Customer_Retention_Analysis.xlsx
 ├── Kayako_Churn_Reduction_Plan.xlsx
 ├── Kayako_Top20_Revenue_Concentration.xlsx
 ├── Kayako_Sales_Forecast_Pipeline.xlsx
 └── Kayako_One_PAGER.xlsx
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

## How to Use

### Analyse a deal
```bash
python ma_agent.py --deal "C:/deals/CompanyName"
```

### Focus on a specific area
```bash
python ma_agent.py --deal "C:/deals/CompanyName" --task "focus on churn risk only"
```

### List all past deals
```bash
python ma_agent.py --deals
```

### Recall a past deal
```bash
python ma_agent.py --recall "Kayako"
```

---

## Example

```bash
python ma_agent.py --deal "C:/deals/Kayako"
```

**What happens:**
1. Agent lists all files in the Kayako folder
2. Reads the CIM, retention analysis, churn plan, revenue concentration
3. Calculates 5-year revenue CAGR, ARR growth, cohort retention curves
4. Flags top risks (customer concentration, churn trend, EBITDA quality)
5. Saves full report to `C:/deals/Kayako/outputs/ma_analysis.md`
6. You can then ask: *"What's the biggest acquisition risk?"*

**Output report covers:**
- Revenue & ARR trends (2019–2024)
- Gross margin and EBITDA expansion
- Customer count and ARPU growth
- Cohort retention analysis
- Top 20 customer concentration
- Sales pipeline and forecast
- Key risks and diligence questions

---

## Supported File Types

| Format | Examples |
|--------|---------|
| Excel  | `.xlsx`, `.xls` |
| CSV    | `.csv` |
| Word   | `.docx` |
| PDF    | `.pdf` |
| Text   | `.txt`, `.md` |
| JSON   | `.json` |

---

## Browser Demo (Streamlit)

To run the interactive chat interface in your browser:

```bash
streamlit run app.py
```

Opens at `http://localhost:8501` — enter your API key in the sidebar and ask questions about the Kayako data room.

---

## Requirements

- Python 3.9+
- An Anthropic API key ([get one free](https://console.anthropic.com))
- Dependencies: `anthropic`, `streamlit`, `openpyxl` (see `requirements.txt`)

---

## Who Is This For?

- **M&A analysts** who want to speed up initial due diligence
- **Investors** screening deals quickly before deep dives
- **Associates** generating first-pass analysis on new targets
- **Anyone** who receives data room files and needs fast answers

---

## Notes

- The agent never modifies your original files — it only reads them
- Reports are saved to an `outputs/` subfolder inside your deal folder
- Deal memory is stored locally at `~/.ma_analyst/deals/`
- Works on Windows, Mac, and Linux
