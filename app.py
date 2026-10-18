from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from excel_report import build_excel_report, build_profile_excel_report
from mis_core import analyze_sales, build_report, validate_sales_result_schema
from report_templates import REPORT_TYPES, analyze_template, default_mappings, mapping_candidates, validate_template_result_schema
from spreadsheet_core import load_spreadsheet, profile_spreadsheet, validate_profile_schema


BASE_DIR = Path(__file__).parent
SAMPLE_FILE = BASE_DIR / "sample_sales.csv"
MAX_UPLOAD_SIZE_MB = 50
INR = lambda value: f"₹{value:,.2f}"


def _sales_issues_with_profile(sales_result: dict, profile: dict) -> pd.DataFrame:
    """Combine sales checks with generic checks on otherwise unmapped fields."""
    sales_issues = sales_result["issues"].copy()
    mapped_columns = set(sales_result["mapping"].values())
    generic_records = []
    for record in profile["issues"].to_dict("records"):
        column = str(record["Column"])
        issue = str(record["Issue"])
        if issue == "Exact duplicate row":
            continue  # analyze_sales reports this with its established sales schema
        if column in mapped_columns:
            continue  # the sales analyzer gives the recognized field a specific check
        generic_records.append({
            "Row": record["Row"],
            "Issue": issue,
            "Column": column,
            "Value": record["Original Value"],
            "Severity": record["Severity"],
            "Explanation": record["Explanation"],
            "Recommended Action": record["Recommended Action"],
        })
    if generic_records:
        sales_issues = pd.concat([sales_issues, pd.DataFrame(generic_records)], ignore_index=True)
    if not sales_issues.empty:
        sales_issues = sales_issues.drop_duplicates(subset=["Row", "Issue", "Column"], keep="first")
        sales_issues = sales_issues.sort_values(["Row", "Column", "Issue"], kind="stable").reset_index(drop=True)
    return sales_issues


def _csv_with_issue_flags(data: pd.DataFrame, issues: pd.DataFrame) -> bytes:
    flagged = data.copy()
    flags: dict[int, list[str]] = {}
    for row in issues.to_dict("records"):
        flags.setdefault(int(row["Row"]), []).append(str(row["Issue"]))
    flagged["MIS Data Quality Issues"] = ["; ".join(dict.fromkeys(flags.get(index + 2, []))) for index in range(len(flagged))]
    return flagged.to_csv(index=False).encode("utf-8-sig")


def _template_text_report(profile: dict, template_result: dict, issues: pd.DataFrame) -> str:
    lines = [
        f"MIS AUTOPILOT | {template_result['template'].upper()} SUMMARY",
        "=" * (25 + len(template_result["template"])),
        f"Source: {profile['source_name']}",
        f"Rows: {profile['row_count']:,}",
        f"Columns: {profile['column_count']:,}",
        f"Exact duplicate rows: {profile['duplicate_count']:,} (retained in original data)",
        f"Data-quality findings: {len(issues):,}",
        "",
        "CONFIRMED COLUMN MAPPINGS",
        "-" * 27,
    ]
    lines.extend(f"{role.replace('_', ' ').title()}: {column}" for role, column in template_result["mapping"].items())
    lines.extend(["", "METRICS", "-" * 7])
    lines.extend(f"{metric['label']}: {metric['value']}" for metric in template_result["metrics"])
    if template_result["missing_requirements"]:
        lines.extend(["", "METRICS NOT CALCULATED", "-" * 22])
        lines.extend(f"- {message}" for message in template_result["missing_requirements"])
    lines.extend(["", *template_result["notes"]])
    return "\n".join(lines) + "\n"


st.set_page_config(page_title="MIS Autopilot", layout="wide")
st.title("MIS Autopilot")
st.caption("A spreadsheet MIS assistant for small businesses. Choose a report template that matches your data.")

with st.sidebar:
    st.subheader("Analyze a spreadsheet")
    report_type = st.selectbox("Report template", REPORT_TYPES, index=0)
    uploaded = st.file_uploader("Upload Excel or CSV", type=["csv", "xlsx", "xlsm", "xls"])
    st.caption(f"Files up to {MAX_UPLOAD_SIZE_MB} MB are accepted. The uploaded data is processed by this app session.")
    st.divider()
    st.markdown("**Demo data**")
    st.caption("With no upload, the included mock sales file is selected. Upload a file or choose a matching template to analyze other business data.")

sheet_name: str | None = None
if uploaded is None and not SAMPLE_FILE.is_file():
    st.info("The optional demo file is not present. Upload a CSV or Excel workbook to begin.")
    st.stop()

try:
    if uploaded is None:
        raw_data, _ = load_spreadsheet(SAMPLE_FILE.read_bytes(), SAMPLE_FILE.name)
        source_name = "sample_sales.csv (mock data)"
    else:
        source_name = uploaded.name
        payload = uploaded.getvalue()
        raw_data, available_sheets = load_spreadsheet(payload, uploaded.name)
        if available_sheets:
            sheet_name = st.sidebar.selectbox("Worksheet", available_sheets)
            if sheet_name != available_sheets[0]:
                raw_data, _ = load_spreadsheet(payload, uploaded.name, sheet_name=sheet_name)
    raw_data = raw_data.copy().reset_index(drop=True)
    if not len(raw_data.columns):
        st.warning("This file has no header row. Add one header row and try again.")
        st.stop()
except (ValueError, OSError) as exc:
    st.error(f"Could not read this file: {exc}", icon="⚠️")
    st.info("Supported formats are .csv, .xlsx, .xlsm and .xls. Excel workbooks can contain multiple worksheets.")
    st.stop()

profile = validate_profile_schema(profile_spreadsheet(raw_data, source_name))

if uploaded is None:
    st.info("Showing the included mock sales dataset. Upload a spreadsheet from the sidebar to analyze your data.")

candidates = mapping_candidates(raw_data.columns, report_type)
mapping = default_mappings(candidates)
ambiguous_roles: list[str] = []
with st.sidebar.expander("Confirm column mappings", expanded=report_type != "General spreadsheet analysis"):
    for role, options in candidates.items():
        if not options:
            continue
        label = role.replace("_", " ").title()
        if len(options) == 1:
            chosen = st.selectbox(label, options, index=0, key=f"mapping_{report_type}_{role}")
            mapping[role] = chosen
        else:
            ambiguous_roles.append(role)
            selection_options = ["— choose a column —", *options]
            chosen = st.selectbox(label, selection_options, index=0, key=f"mapping_{report_type}_{role}")
            if chosen != selection_options[0]:
                mapping[role] = chosen
            else:
                mapping.pop(role, None)
                st.caption(f"Several columns could mean {role.replace('_', ' ')}. Select the correct one to use this metric.")

mapping_conflicts = {
    column: [role for role, mapped_column in mapping.items() if mapped_column == column]
    for column in set(mapping.values())
    if sum(mapped_column == column for mapped_column in mapping.values()) > 1
}
template_result = validate_template_result_schema(analyze_template(raw_data, report_type, mapping, source_name, profile))

if report_type == "Sales" and template_result["ready"] and template_result["sales_result"] is not None:
    sales_result = validate_sales_result_schema(template_result["sales_result"])
    visible_issues = _sales_issues_with_profile(sales_result, profile)
else:
    sales_result = None
    visible_issues = profile["issues"].copy()

if mapping_conflicts:
    conflicts = "; ".join(f"{column}: {', '.join(roles)}" for column, roles in mapping_conflicts.items())
    st.warning(f"A source column is mapped to more than one field ({conflicts}). Confirm distinct columns in the sidebar before using these metrics.")
    template_result["ready"] = False

if template_result["missing_requirements"]:
    st.warning("Some template metrics are unavailable: " + " ".join(template_result["missing_requirements"]))

mapping_text = " · ".join(f"{role.replace('_', ' ').title()}: {column}" for role, column in mapping.items())
if mapping_text:
    st.caption(f"Confirmed mappings — {mapping_text}")
if sheet_name:
    st.caption(f"Worksheet: {sheet_name}")

with st.container(horizontal=True):
    if sales_result is not None:
        st.metric("Total sales", INR(sales_result["validated_total_sales"]), border=True)
        st.metric(sales_result["order_label"], f"{sales_result['order_count']:,}", border=True)
        st.metric("Average order value", INR(sales_result["average_order_value"]) if sales_result["average_order_value"] is not None else "—", border=True)
        st.metric("Quality findings", f"{len(visible_issues):,}", border=True)
    else:
        st.metric("Source rows", f"{profile['row_count']:,}", border=True)
        st.metric("Columns", f"{profile['column_count']:,}", border=True)
        st.metric("Exact duplicate rows", f"{profile['duplicate_count']:,}", border=True)
        st.metric("Quality findings", f"{len(visible_issues):,}", border=True)

overview, profile_tab, quality, cleaned = st.tabs(["MIS overview", "Data profile", "Data quality", "Cleaned data"])

with overview:
    if sales_result is not None and template_result["ready"]:
        left, right = st.columns([1.2, 1])
        with left:
            st.subheader("Top products by validated sales")
            if sales_result["top_products"]:
                products = pd.DataFrame(sales_result["top_products"]).set_index("Product")
                st.bar_chart(products, height=300, alt="Sales by top product")
                st.dataframe(products.rename(columns={"Sales": "Sales (₹)"}), alt="Top five products ranked by validated sales", hide_index=False)
            else:
                st.info("A product field is not mapped, so product rankings are unavailable.")
        with right:
            st.subheader("At a glance")
            st.write(f"**Source:** {source_name}")
            st.write(f"**Records analyzed:** {sales_result['record_count']:,}")
            st.write(f"**Exact duplicate rows:** {sales_result['duplicate_count']:,}")
            st.write(f"**Repeated transaction candidates (Record ID ignored):** {sales_result['repeated_transaction_count']:,}")
            st.write(f"**Worksheet:** {sheet_name or 'Default / first sheet'}")
            st.caption("Source rows and values are retained. Validated sales summaries exclude exact full-row duplicates; repeated transaction candidates remain included pending review.")
    else:
        st.subheader(f"{report_type} overview")
        if not template_result["ready"]:
            st.info("The file is profiled and remains available below. Confirm the required columns in the sidebar to calculate this template's metrics.")
        else:
            metric_columns = st.columns(min(max(len(template_result["metrics"]), 1), 4))
            for index, metric in enumerate(template_result["metrics"]):
                value = metric["value"]
                if value is None:
                    shown = "—"
                elif metric["kind"] == "currency":
                    shown = INR(value)
                elif metric["kind"] == "percent":
                    shown = f"{value:.1%}"
                elif isinstance(value, (int, float)):
                    shown = f"{value:,.2f}" if isinstance(value, float) and not value.is_integer() else f"{value:,.0f}"
                else:
                    shown = str(value)
                with metric_columns[index % len(metric_columns)]:
                    st.metric(metric["label"], shown, help=metric.get("note") or None, border=True)
            if template_result["notes"]:
                for note in template_result["notes"]:
                    st.caption(note)
            for title, summary in template_result["summaries"].items():
                st.subheader(title)
                st.dataframe(summary, hide_index=True, width="stretch")
                numeric_cols = summary.select_dtypes(include="number").columns.tolist()
                if len(summary) and numeric_cols:
                    st.bar_chart(summary.set_index(summary.columns[0])[numeric_cols[:2]], height=260)
        if report_type == "General spreadsheet analysis":
            st.subheader("Numeric summaries")
            if profile["numeric_summary"].empty:
                st.info("No numeric columns were detected confidently.")
            else:
                st.dataframe(profile["numeric_summary"], hide_index=True, width="stretch")
            for column, summary in profile["category_summaries"].items():
                st.subheader(f"Top values: {column}")
                st.bar_chart(summary.set_index("Category"), height=220)

    st.divider()
    report_content = build_report({**sales_result, "issues": visible_issues}) if sales_result is not None else _template_text_report(profile, template_result, visible_issues)
    if report_type == "Sales" and sales_result is not None:
        excel_builder = lambda: build_excel_report({**sales_result, "issues": visible_issues})
        excel_name = "mis_autopilot_sales_report.xlsx"
    else:
        excel_builder = lambda: build_profile_excel_report(profile, template_result)
        safe_template = report_type.lower().replace(" ", "_").replace("&", "and")
        excel_name = f"mis_autopilot_{safe_template}_report.xlsx"
    st.download_button(
        "Download Excel MIS Report", data=excel_builder, file_name=excel_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
    )
    st.download_button(
        "Download MIS report", data=report_content, file_name="mis_autopilot_report.txt",
        mime="text/plain", width="stretch",
    )

with profile_tab:
    st.subheader("Uploaded data profile")
    st.write(f"**Source rows:** {profile['row_count']:,} · **Nonblank rows:** {profile['nonblank_row_count']:,} · **Blank rows:** {profile['blank_row_count']:,} · **Columns:** {profile['column_count']:,}")
    st.caption("Detected types are best-effort hints. Original cell values remain unchanged in the preview and downloads.")
    st.dataframe(profile["column_profile"], hide_index=True, width="stretch")
    st.subheader("Original data preview")
    st.dataframe(profile["original_data"].head(100), hide_index=True, width="stretch", height=380)

with quality:
    st.subheader("Data-quality findings")
    st.caption("Source row numbers count the header as row 1. Suspicious rows are flagged and retained; each row may have multiple findings.")
    if visible_issues.empty:
        st.success("No issues were found by the applicable checks.")
    else:
        if "Issue" in visible_issues:
            counts = visible_issues["Issue"].value_counts().rename_axis("Issue").reset_index(name="Count")
            st.dataframe(counts, hide_index=True, width="stretch")
        st.dataframe(visible_issues, hide_index=True, width="stretch", height=420)

with cleaned:
    if sales_result is not None:
        st.subheader("Derived sales view")
        st.caption("This derived preview parses recognized dates and numeric fields and excludes exact full-row duplicates. The Original data preview and Excel report preserve every uploaded row and source value; duplicate rows remain visible in the issue table.")
        st.dataframe(sales_result["cleaned_data"], hide_index=True, width="stretch", height=380)
        clean_csv = sales_result["cleaned_data"].to_csv(index=False).encode("utf-8-sig")
        st.download_button("Download derived sales CSV", data=clean_csv, file_name="derived_sales_data.csv", mime="text/csv", width="content")
    else:
        st.subheader("Original data retained")
        st.caption("No generic cleaning or row removal is performed. Download the original values with an appended issue-flags column for review.")
        st.dataframe(profile["original_data"], hide_index=True, width="stretch", height=380)
        flagged_csv = _csv_with_issue_flags(profile["original_data"], visible_issues)
        st.download_button("Download data with issue flags", data=flagged_csv, file_name="data_with_quality_flags.csv", mime="text/csv", width="content")
