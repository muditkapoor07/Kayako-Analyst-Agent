"""
Kayako M&A Analyst — Interactive Q&A Chat App
Powered by Claude. Loads all data room files at startup.
Run with: streamlit run app.py
"""

import os
import streamlit as st
import anthropic
import openpyxl

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Kayako M&A Analyst",
    page_icon="📊",
    layout="centered",
)

# ── Data room loader ──────────────────────────────────────────────────────────
DATA_ROOM_DIR = os.path.dirname(os.path.abspath(__file__))

FILES = {
    "Confidential Information Memorandum": "Kayako_Confidential_Information_Memorandum.xlsx",
    "Customer Retention Analysis":         "Kayako_Customer_Retention_Analysis.xlsx",
    "Churn Reduction Plan":                "Kayako_Churn_Reduction_Plan.xlsx",
    "One Pager":                           "Kayako_One_PAGER.xlsx",
    "Sales Forecast & Pipeline":           "Kayako_Sales_Forecast_Pipeline.xlsx",
    "Top 20 Revenue Concentration":        "Kayako_Top20_Revenue_Concentration.xlsx",
}


def load_excel(path: str) -> str:
    """Read all sheets from an Excel file and return as plain text."""
    wb = openpyxl.load_workbook(path, data_only=True)
    parts = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            if any(cell is not None for cell in row):
                rows.append("\t".join("" if c is None else str(c) for c in row))
        if rows:
            parts.append(f"### Sheet: {sheet_name}\n" + "\n".join(rows))
    return "\n\n".join(parts)


@st.cache_resource(show_spinner="Loading data room files…")
def build_context() -> str:
    sections = []
    for label, filename in FILES.items():
        path = os.path.join(DATA_ROOM_DIR, filename)
        if os.path.exists(path):
            content = load_excel(path)
            sections.append(f"## {label}\n\n{content}")
        else:
            sections.append(f"## {label}\n\n[File not found: {filename}]")
    return "\n\n{'='*80}\n\n".join(sections)


SYSTEM_PROMPT = """You are an expert M&A analyst with deep experience in SaaS company due diligence. \
You have been given access to the complete Kayako data room, which contains the following documents:

- Confidential Information Memorandum (CIM) — financial summary 2019–2024
- Customer Retention Analysis — cohort retention data
- Churn Reduction Plan — management's churn initiatives
- One Pager — high-level company summary
- Sales Forecast & Pipeline — forward-looking revenue projections
- Top 20 Revenue Concentration — top customer ARR breakdown

The full contents of all these files are provided below. Answer questions accurately and concisely \
based solely on this data. When relevant, cite specific numbers, tables, or file names. \
If a question cannot be answered from the available data, say so clearly and suggest what \
additional information would be needed.

Maintain a professional, executive-level tone. For financial figures always include units (USD, %, etc.).

--- DATA ROOM CONTENTS ---

{context}
"""

# ── UI ────────────────────────────────────────────────────────────────────────
st.title("📊 Kayako M&A Analyst")
st.caption("Ask anything about the Kayako data room — financials, churn, retention, pipeline, and more.")

# Load data
context = build_context()
system = SYSTEM_PROMPT.format(context=context)

# API key input (sidebar)
with st.sidebar:
    st.header("Configuration")
    api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        value=os.environ.get("ANTHROPIC_API_KEY", ""),
        help="Your key stays in your browser session only.",
    )
    st.divider()
    st.markdown("**Data Room Files Loaded:**")
    for label in FILES:
        path = os.path.join(DATA_ROOM_DIR, FILES[label])
        icon = "✅" if os.path.exists(path) else "❌"
        st.markdown(f"{icon} {label}")
    st.divider()
    if st.button("🗑️ Clear conversation"):
        st.session_state.messages = []
        st.rerun()
    st.caption("Powered by Claude Opus 4.6")

# Initialise chat history
if "messages" not in st.session_state:
    st.session_state.messages = []

# Render existing messages
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Suggested starter questions
if not st.session_state.messages:
    st.markdown("**Try asking:**")
    starters = [
        "What is Kayako's ARR growth trend from 2019 to 2024?",
        "What are the top 5 revenue concentration risks?",
        "What is the biggest acquisition risk?",
        "How bad is churn and what is management doing about it?",
        "Summarise the sales forecast and pipeline outlook.",
    ]
    cols = st.columns(1)
    for q in starters:
        if st.button(q, use_container_width=True):
            st.session_state.messages.append({"role": "user", "content": q})
            st.rerun()

# Chat input
if prompt := st.chat_input("Ask a question about the Kayako data room…"):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # Generate response
    if not api_key:
        with st.chat_message("assistant"):
            st.warning("Please enter your Anthropic API key in the sidebar.")
    else:
        with st.chat_message("assistant"):
            client = anthropic.Anthropic(api_key=api_key)
            response_placeholder = st.empty()
            full_response = ""

            with client.messages.stream(
                model="claude-opus-4-6",
                max_tokens=4096,
                system=system,
                messages=[
                    {"role": m["role"], "content": m["content"]}
                    for m in st.session_state.messages
                ],
            ) as stream:
                for text in stream.text_stream:
                    full_response += text
                    response_placeholder.markdown(full_response + "▌")
            response_placeholder.markdown(full_response)

        st.session_state.messages.append({"role": "assistant", "content": full_response})
