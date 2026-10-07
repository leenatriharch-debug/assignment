from __future__ import annotations

import json
import os
import csv
import io
from typing import Any

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from workflow_agent.engine import DEFAULT_WORKBOOK, load_workflow_catalog, run_agent, select_workflow


load_dotenv()
st.set_page_config(page_title="Workflow Studio", page_icon="W", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Manrope:wght@400;500;600;700;800&display=swap');
    :root { --ink:#193832; --muted:#62766f; --green:#176b55; --lime:#c8ed79; --coral:#e27658; --paper:#f2f5f0; --line:#d7dfd7; }
    html, body, [class*="css"] { font-family:'Manrope','Segoe UI',sans-serif; color:var(--ink); }
    .stApp { background:var(--paper); }
    [data-testid="stHeader"] { background:rgba(242,245,240,.94); }
    [data-testid="stSidebar"] { background:#e7eee8; border-right:1px solid var(--line); }
    [data-testid="stSidebar"] > div { padding-top:1.2rem; }
    h1,h2,h3 { color:var(--ink); letter-spacing:0; }
    h1 { font-size:2rem; font-weight:800; }
    h2 { font-size:1.12rem; font-weight:700; }
    .eyebrow { color:var(--green); font-family:'DM Mono',monospace; font-size:.72rem; letter-spacing:.08em; text-transform:uppercase; }
    .masthead { border-bottom:1px solid var(--line); padding:0 0 1rem 0; margin-bottom:1.15rem; }
    .brand { display:flex; align-items:center; gap:.7rem; font-weight:800; font-size:1.12rem; }
    .brand-mark { background:var(--green); color:var(--lime); width:2rem; height:2rem; display:grid; place-items:center; font-family:'DM Mono',monospace; }
    .subtle { color:var(--muted); font-size:.88rem; }
    div[data-testid="stMetric"] { background:#fff; border:1px solid var(--line); border-radius:4px; padding:.8rem 1rem; }
    div[data-testid="stMetricLabel"] { color:var(--muted); }
    div[data-testid="stMetricValue"] { color:var(--green); }
    .stButton > button[kind="primary"] { background:var(--green); border:0; color:white; font-weight:700; }
    .stButton > button[kind="primary"]:hover { background:#105642; color:white; }
    .stDownloadButton > button { border-color:var(--green); color:var(--green); }
    code, pre, .stCode { font-family:'DM Mono',Consolas,monospace!important; }
    [data-testid="stDataFrame"] { border:1px solid var(--line); }
    </style>
    """,
    unsafe_allow_html=True,
)


def file_upload(label: str, input_key: str, filename_key: str | None = None) -> dict[str, Any]:
    uploaded = st.file_uploader(label, type=["csv", "xlsx"], key=f"upload_{input_key}")
    if uploaded is None:
        return {}
    values = {input_key: uploaded.getvalue()}
    if filename_key:
        values[filename_key] = uploaded.name
    else:
        values[f"{input_key}_filename"] = uploaded.name
    return values


def collect_inputs(workflow_id: str) -> dict[str, Any]:
    values: dict[str, Any] = {}
    if workflow_id == "WF001":
        override = st.checkbox("Use one minimum-stock threshold for every row")
        if override:
            values["minimum_stock"] = st.number_input("Minimum stock", min_value=0, value=10, step=1)
        values.update(file_upload("Inventory file (optional; fixture used by default)", "inventory_file"))
    elif workflow_id == "WF002":
        values["tolerance_percent"] = st.number_input("Flag differences over (%)", min_value=0.0, value=10.0, step=1.0)
        left, right = st.columns(2)
        with left:
            values.update(file_upload("Product prices (optional)", "products_file"))
        with right:
            values.update(file_upload("Vendor price list (optional)", "vendor_prices_file"))
    elif workflow_id == "WF003":
        values.update(file_upload("Vendor CSV/XLSX (optional)", "vendor_file", "vendor_filename"))
    elif workflow_id == "WF004":
        left, right = st.columns(2)
        with left:
            values["product_name"] = st.text_input("Product name")
            values["category"] = st.text_input("Category")
            values["material"] = st.text_input("Material")
        with right:
            values["attributes"] = st.text_input("Other known attributes")
            values["color"] = st.text_input("Color")
            values["target_audience"] = st.text_input("Target audience")
        if os.getenv("OPENAI_API_KEY"):
            values["use_openai"] = st.checkbox("Use OpenAI for copy generation", value=False)
    elif workflow_id == "WF005":
        left, right = st.columns(2)
        with left:
            values["order_id"] = st.text_input("Order ID", placeholder="ORD-1001")
        with right:
            values["customer_email"] = st.text_input("Customer email")
        values.update(file_upload("Order data (optional)", "orders_file"))
    elif workflow_id == "WF006":
        values["similarity_threshold"] = st.slider("Possible-duplicate name similarity", min_value=0.5, max_value=1.0, value=0.84, step=0.01)
        values.update(file_upload("Product catalog (optional)", "catalog_file"))
    elif workflow_id == "WF007":
        left, right = st.columns(2)
        with left:
            values["campaign_goal"] = st.text_input("Campaign goal")
            values["product_list"] = st.text_input("Products")
            values["target_audience"] = st.text_input("Target audience")
        with right:
            values["promotion"] = st.text_input("Promotion")
            values["start_date"] = st.text_input("Start date (YYYY-MM-DD)")
            values["end_date"] = st.text_input("End date (YYYY-MM-DD)")
        if os.getenv("OPENAI_API_KEY"):
            values["use_openai"] = st.checkbox("Use OpenAI for brief generation", value=False)
    elif workflow_id == "WF008":
        values["category"] = st.text_input("Default category", value="General")
        values.update(file_upload("Keyword CSV/XLSX (optional)", "keywords_file"))
    elif workflow_id == "WF009":
        values["task_description"] = st.text_area("Task description", height=85)
        values["required_skills"] = st.text_input("Required skills (comma-separated)")
        left, right = st.columns(2)
        with left:
            default_priority = 0 if "urgent" in st.session_state.get("request", "").lower() else 2
            values["priority"] = st.selectbox("Priority", ["urgent", "high", "normal", "low"], index=default_priority)
        with right:
            values["deadline"] = st.text_input("Deadline")
        values.update(file_upload("Employee list (optional)", "employees_file"))
    elif workflow_id == "WF010":
        values["failure_threshold_percent"] = st.number_input("Failure-rate threshold (%)", min_value=0.0, value=10.0, step=1.0)
        values["slow_step_threshold_seconds"] = st.number_input("Slow-step threshold (seconds)", min_value=0.0, value=30.0, step=1.0)
        values.update(file_upload("Execution logs (optional)", "logs_file"))
    return values


def csv_export(rows: list[dict[str, Any]]) -> str:
    output = io.StringIO(newline="")
    headers = list(dict.fromkeys(key for row in rows for key in row))
    writer = csv.DictWriter(output, fieldnames=headers, extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return output.getvalue()


def render_result(result: dict[str, Any]) -> None:
    status = result.get("status", "unknown")
    if status in {"completed", "plan_only"}:
        st.success(f"{result.get('workflow_id', '')} · {result.get('workflow_name', '')} · {status.replace('_', ' ')}")
    elif status in {"needs_input", "needs_clarification"}:
        st.warning(result.get("message", "More information is required."))
    elif status in {"not_found", "escalated"}:
        st.info(result.get("message", status.replace("_", " ")))
    else:
        st.error(result.get("message", "The workflow could not be completed."))

    if "metrics" in result:
        metrics = result["metrics"]
        columns = st.columns(4)
        for column, label, value in zip(columns, ("Runs", "Failures", "Failure rate", "Average seconds"),
                                        (metrics["runs"], metrics["failures"], f"{metrics['failure_rate_percent']}%", metrics["average_execution_seconds"])):
            column.metric(label, value)
    for key, label in (("items_to_restock", "Restock list"), ("price_report", "Price validation"),
                       ("cleaned_rows", "Cleaned rows"), ("invalid_rows", "Invalid rows"),
                       ("duplicate_groups", "Duplicate groups"), ("keyword_report", "Keyword report"),
                       ("ranked_candidates", "Candidate ranking"), ("slow_steps", "Slow steps"),
                       ("frequent_errors", "Frequent errors")):
        if key in result and result[key]:
            st.subheader(label)
            st.dataframe(pd.json_normalize(result[key]), width="stretch", hide_index=True)
    if "assignment" in result:
        st.subheader("Recommended assignment")
        st.json(result["assignment"])
    if "order" in result:
        st.subheader("Order and shipment")
        st.json(result["order"])
    if "content" in result:
        st.subheader("Generated product content")
        st.json(result["content"])
    if "brief" in result:
        st.subheader("Campaign brief")
        st.json(result["brief"])
    if "summary" in result:
        st.caption(f"Vendor file: {result['summary']['received']} rows · {result['summary']['valid']} valid · {result['summary']['invalid']} invalid")
    if "recommendations" in result:
        st.subheader("Recommendations")
        for recommendation in result["recommendations"]:
            st.write(f"- {recommendation}")
    if result.get("missing_fields"):
        st.caption("Required before generation: " + ", ".join(result["missing_fields"]))
    if result.get("candidates"):
        st.caption("Possible workflows: " + ", ".join(result["candidates"]))
    if result.get("execution_steps"):
        st.write("**Registered steps:** " + result["execution_steps"])
    if result.get("api_calls"):
        st.subheader("External API calls")
        st.json(result["api_calls"])
    if result.get("tools_used"):
        st.caption("Tools: " + " · ".join(result["tools_used"]))
    if result.get("cleaned_rows"):
        st.download_button("Download cleaned vendor CSV", csv_export(result["cleaned_rows"]),
                           file_name="cleaned_vendor_data.csv", mime="text/csv", width="content")
    if result.get("keyword_report"):
        st.download_button("Download keyword report CSV", csv_export(result["keyword_report"]),
                           file_name="keyword_report.csv", mime="text/csv", width="content")
    st.download_button("Download result JSON", json.dumps(result, indent=2, default=str),
                       file_name=f"{result.get('workflow_id', 'workflow').lower()}_result.json", mime="application/json", width="content")


catalog = load_workflow_catalog(DEFAULT_WORKBOOK)
with st.sidebar:
    st.markdown('<div class="eyebrow">Assessment build · 10 workflows</div>', unsafe_allow_html=True)
    workflow_options = {"Auto-detect from request": None}
    workflow_options.update({f"{item.workflow_id} · {item.name}": item.workflow_id for item in catalog.values()})
    selected_label = st.selectbox("Workflow", list(workflow_options))
    selected_id = workflow_options[selected_label]
    st.divider()
    if st.button("Load selected sample request", width="stretch"):
        examples = {
            "WF001": "Which products need restocking?",
            "WF002": "Find products where vendor price differs by more than 10%.",
            "WF003": "Process this vendor spreadsheet and show invalid rows.",
            "WF004": "Generate SEO content for this product.",
            "WF005": "Where is order ORD-1001?",
            "WF006": "Find likely duplicate products in the catalog.",
            "WF007": "Create a campaign brief for the new collection.",
            "WF008": "Classify these keywords and map them to pages.",
            "WF009": "Assign this urgent task to the best available developer.",
            "WF010": "Which workflows are failing most often?",
        }
        st.session_state["request"] = examples.get(selected_id, "")

st.markdown(
    '<div class="masthead"><div class="brand"><span class="brand-mark">W</span> Workflow Studio</div>'
    '<div class="subtle">Workbook-driven operations agent · LangGraph execution</div></div>',
    unsafe_allow_html=True,
)
left, right = st.columns([1.2, 0.8], gap="large")
with left:
    st.markdown('<div class="eyebrow">01 / Request</div>', unsafe_allow_html=True)
    request = st.text_area("Describe the operation", key="request", height=90,
                           placeholder="Try: Which products need restocking?")
with right:
    st.markdown('<div class="eyebrow">02 / Workflow definition</div>', unsafe_allow_html=True)
    if selected_id:
        definition = catalog[selected_id]
        st.markdown(f"**{definition.workflow_id} · {definition.name}**")
        st.caption(f"{definition.trigger}\n\nInputs: {definition.inputs}")
    else:
        st.markdown("**Automatic routing**")
        st.caption("The router compares your request with the trigger and name of every workflow in the Excel catalog.")

st.markdown('<div class="eyebrow">03 / Workflow inputs</div>', unsafe_allow_html=True)
input_workflow = selected_id
if input_workflow is None:
    sample_request = st.session_state.get("request", "")
    inferred_id, _ = select_workflow(sample_request, catalog)
    input_workflow = inferred_id or "WF001"
inputs = collect_inputs(input_workflow)

if st.button("Run workflow", type="primary", width="content"):
    if not request.strip():
        st.warning("Enter a request to route or select a workflow from the sidebar.")
    else:
        try:
            with st.spinner("Routing request and executing workflow tools..."):
                result = run_agent(request, inputs, selected_id)
            st.divider()
            st.markdown('<div class="eyebrow">04 / Result</div>', unsafe_allow_html=True)
            render_result(result)
        except Exception as exc:
            st.error(f"Workflow failed: {type(exc).__name__}: {exc}")
            st.caption("Check the input file format and required columns, then retry.")

st.divider()
st.caption(f"Catalog source: {DEFAULT_WORKBOOK.name} · {len(catalog)} registered workflows · Sample data is synthetic")