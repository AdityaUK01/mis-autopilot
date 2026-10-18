"""Column detection and lightweight sales data-quality analysis for MIS Autopilot."""

from __future__ import annotations

import math
import re
from typing import Any

import pandas as pd


ALIASES: dict[str, set[str]] = {
    "date": {"date", "order date", "sale date", "sales date", "invoice date", "transaction date", "bill date"},
    "product": {"product", "product name", "item", "item name", "sku", "description", "product description"},
    "quantity": {"quantity", "qty", "units", "unit sold", "units sold", "quantity sold"},
    "unit_price": {"unit price", "price", "selling price", "sale price", "rate", "price per unit"},
    "total_sales": {"total", "sales", "sales amount", "total amount", "line total", "revenue", "amount", "total sales", "net sales"},
    "region": {"region", "area", "territory", "state", "city", "location", "market"},
    "record_id": {"record id", "record number", "row id", "line id", "record key"},
    "order_id": {"order id", "order number", "order no", "invoice number", "invoice no", "transaction id", "sales id", "bill number", "bill no"},
}

REQUIRED_FOR_QUALITY = ("date", "product", "quantity", "unit_price", "total_sales", "region", "order_id")
SALES_RESULT_SCHEMA_VERSION = 1
SALES_RESULT_REQUIRED_KEYS = frozenset({
    "schema_version", "source_name", "original_data", "data", "cleaned_data", "mapping",
    "missing_columns", "issues", "raw_sales_by_row", "validated_sales_by_row",
    "raw_total_sales", "validated_total_sales", "total_sales", "record_count",
    "validated_record_count", "order_count", "unique_order_count", "order_label",
    "average_order_value", "aov_sales", "aov_denominator", "top_products",
    "product_summary", "region_summary", "region_available", "sales_trend",
    "sales_trend_period", "duplicate_count", "repeated_transaction_count",
})


def validate_sales_result_schema(result: dict[str, Any]) -> dict[str, Any]:
    """Enforce the documented producer/consumer contract for sales analysis."""
    missing = sorted(SALES_RESULT_REQUIRED_KEYS.difference(result))
    version = result.get("schema_version")
    if version != SALES_RESULT_SCHEMA_VERSION or missing:
        details = []
        if version != SALES_RESULT_SCHEMA_VERSION:
            details.append(f"expected schema version {SALES_RESULT_SCHEMA_VERSION}, got {version!r}")
        if missing:
            details.append("missing keys: " + ", ".join(missing))
        raise ValueError("Invalid sales analysis result schema (" + "; ".join(details) + ").")
    return result


def normalize_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").strip().lower()).strip()


def detect_columns(columns: list[Any] | pd.Index) -> dict[str, str]:
    """Map canonical field names to source headers, preferring exact aliases."""
    source_headers = list(columns)
    normalized = {header: normalize_header(header) for header in source_headers}
    mapping: dict[str, str] = {}
    used: set[Any] = set()
    for field, aliases in ALIASES.items():
        match = next((header for header in source_headers if header not in used and normalized[header] in aliases), None)
        if match is None:
            # Permit clear multiword variants such as "Product (SKU)" without
            # accidentally mapping a generic column such as "Sales Region" as sales.
            match = next((header for header in source_headers if header not in used and any(
                normalized[header].startswith(alias + " ") or normalized[header].endswith(" " + alias)
                for alias in aliases if len(alias.split()) > 1
            )), None)
        if match is not None:
            mapping[field] = str(match)
            used.add(match)
    return mapping


def parse_number(value: Any) -> float | None:
    if pd.isna(value) or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        result = float(value)
        return result if math.isfinite(result) else None
    raw = str(value).strip().replace(",", "")
    raw = re.sub(r"^[^\d+\-.]+", "", raw)
    raw = re.sub(r"[^\d.\-]+$", "", raw)
    try:
        result = float(raw)
        return result if math.isfinite(result) else None
    except ValueError:
        return None


def _issue_description(issue: str, column: str) -> tuple[str, str, str]:
    field = issue.removeprefix("Missing ").replace("_", " ").lower()
    if issue.startswith("Missing"):
        severity = "Warning" if field in {"region", "order id"} else "Critical"
        return severity, f"The recognized {field} value is blank.", "Fill in the value from the source record; do not guess."
    if issue == "Invalid date":
        return "Critical", "This value could not be interpreted as a valid date.", "Correct it to a real date, preferably YYYY-MM-DD."
    if issue.startswith("Invalid "):
        return "Critical", f"The value in {column} is not a usable number.", "Enter a numeric value and keep currency symbols or notes out of the cell."
    if issue == "Negative quantity":
        return "Critical", "Quantity is below zero and may reverse or distort sales totals.", "Check the source invoice; represent returns separately if that is the intended meaning."
    if issue == "Negative unit price":
        return "Critical", "Unit price is below zero.", "Verify the price and represent discounts or returns in their own fields."
    if issue == "Negative total sales":
        return "Critical", "The recorded line total is below zero.", "Check the source invoice and confirm whether this is a return or a sign error."
    if issue == "Total sales does not match quantity × unit price":
        return "Critical", "Recorded sales differs from quantity × unit price by more than the rounding tolerance; discounts and taxes are not included in this simple check.", "Verify the amount and account for discounts, tax, or other adjustments before correcting it."
    if issue == "Repeated transaction candidate":
        return "Warning", "This row matches an earlier row across every source column except the recognized record ID. It is a candidate duplicate, not an exact full-row duplicate.", "Compare both rows with the source invoice. Keep both rows until the duplicate is confirmed."
    if issue == "Duplicate record":
        return "Warning", "This row exactly repeats an earlier nonblank record. It remains in Sales Data but is excluded from validated summary totals.", "Confirm whether it is a duplicate transaction; correct or remove it only after checking the source."
    return "Warning", "Review this data-quality finding.", "Check the value against its source document."


def _summary_by_field(
    data: pd.DataFrame,
    mapping: dict[str, str],
    field: str,
    raw_sales: pd.Series,
    validated_sales: pd.Series,
    duplicate_mask: pd.Series,
    missing_label: str,
) -> list[dict[str, Any]]:
    if field in mapping:
        group = data[mapping[field]].map(lambda value: missing_label if pd.isna(value) or not str(value).strip() else str(value).strip())
    else:
        group = pd.Series(missing_label, index=data.index, dtype="object")
    frame = pd.DataFrame({
        "Group": group,
        "Raw Records": 1,
        "Validated Records": (~duplicate_mask).astype(int),
        "Raw Sales": raw_sales,
        "Validated Sales": validated_sales,
    }, index=data.index)
    summary = frame.groupby("Group", dropna=False, sort=False).agg({
        "Raw Records": "sum",
        "Validated Records": "sum",
        "Raw Sales": "sum",
        "Validated Sales": "sum",
    }).reset_index()
    summary = summary.sort_values(["Validated Sales", "Group"], ascending=[False, True], kind="stable")
    return [
        {
            "Label": row["Group"],
            "Raw Records": int(row["Raw Records"]),
            "Validated Records": int(row["Validated Records"]),
            "Raw Sales": float(row["Raw Sales"]),
            "Validated Sales": float(row["Validated Sales"]),
        }
        for _, row in summary.iterrows()
    ]


def analyze_sales(
    data: pd.DataFrame,
    source_name: str = "Uploaded data",
    mapping_override: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Analyze sales while retaining an unmodified copy of the entire source table."""
    original_data = data.copy().reset_index(drop=True)
    if original_data.columns.has_duplicates:
        from spreadsheet_core import unique_headers

        original_data.columns = unique_headers(list(original_data.columns))
    active_mask = pd.Series(False, index=original_data.index)
    for column in original_data.columns:
        values = original_data[column]
        active_mask |= values.notna() & values.astype(str).str.strip().ne("")
    data = original_data.loc[active_mask].copy()
    mapping = detect_columns(data.columns)
    if mapping_override:
        valid_fields = set(ALIASES)
        mapping.update({field: column for field, column in mapping_override.items() if field in valid_fields and column in data.columns})
        # If the user confirms a mapping, do not let one source column be used
        # for two different semantic roles without an explicit reason.
        seen_columns: set[str] = set()
        for field in list(mapping):
            column = mapping[field]
            if column in seen_columns and field not in (mapping_override or {}):
                mapping.pop(field)
            else:
                seen_columns.add(column)
    issues: list[dict[str, Any]] = []

    def add_issue(
        row_index: int,
        issue: str,
        column: str,
        value: Any = "",
        explanation_override: str | None = None,
    ) -> None:
        severity, explanation, action = _issue_description(issue, column)
        issues.append({
            "Row": int(row_index) + 2,
            "Issue": issue,
            "Column": column,
            "Value": "" if pd.isna(value) else str(value),
            "Severity": severity,
            "Explanation": explanation_override or explanation,
            "Recommended Action": action,
        })

    for field in REQUIRED_FOR_QUALITY:
        if field not in mapping:
            continue
        source_col = mapping[field]
        blank = data[source_col].isna() | data[source_col].astype(str).str.strip().eq("")
        for index in data.index[blank]:
            add_issue(index, f"Missing {field.replace('_', ' ')}", source_col)

    parsed_dates: pd.Series | None = None
    if "date" in mapping:
        col = mapping["date"]
        parsed_dates = pd.to_datetime(data[col], errors="coerce", dayfirst=True, format="mixed")
        nonblank = data[col].notna() & data[col].astype(str).str.strip().ne("")
        for index in data.index[nonblank & parsed_dates.isna()]:
            add_issue(index, "Invalid date", col, data.at[index, col])

    numeric: dict[str, pd.Series] = {}
    for field in ("quantity", "unit_price", "total_sales"):
        if field not in mapping:
            continue
        col = mapping[field]
        values = data[col].map(parse_number)
        numeric[field] = values
        nonblank = data[col].notna() & data[col].astype(str).str.strip().ne("")
        for index in data.index[nonblank & values.isna()]:
            add_issue(index, f"Invalid {field.replace('_', ' ')}", col, data.at[index, col])
        if field in {"quantity", "unit_price", "total_sales"}:
            for index in data.index[(values < 0).fillna(False)]:
                issue = "Negative " + field.replace("_", " ")
                add_issue(index, issue, col, data.at[index, col])

    if all(field in numeric for field in ("quantity", "unit_price", "total_sales")):
        expected = numeric["quantity"] * numeric["unit_price"]
        actual = numeric["total_sales"]
        mismatch = expected.notna() & actual.notna() & ~(
            (expected - actual).abs() <= expected.abs().mul(0.01).clip(lower=0.02)
        )
        for index in data.index[mismatch]:
            add_issue(
                index,
                "Total sales does not match quantity × unit price",
                mapping["total_sales"],
                data.at[index, mapping["total_sales"]],
                f"Recorded total {actual.at[index]:g} differs from quantity × unit price ({expected.at[index]:g}) beyond the rounding tolerance.",
            )

    # These checks use the parsed source values before the cleaned date/numeric
    # columns are written below. Exact duplicates compare every source column.
    duplicate_mask = data.duplicated(keep="first")
    repeated_transaction_mask = pd.Series(False, index=data.index, dtype=bool)
    record_id_column = mapping.get("record_id")
    if record_id_column is not None and len(data.columns) > 1:
        transaction_columns = [column for column in data.columns if column != record_id_column]
        repeated_transaction_mask = data.duplicated(subset=transaction_columns, keep="first") & ~duplicate_mask

    for index in data.index[duplicate_mask]:
        duplicate_value = " | ".join(f"{column}={value}" for column, value in data.loc[index].items())
        add_issue(index, "Duplicate record", "Entire row", duplicate_value)
    for index in data.index[repeated_transaction_mask]:
        record_id_value = data.at[index, record_id_column]
        add_issue(
            index,
            "Repeated transaction candidate",
            f"All columns except {record_id_column}",
            f"{record_id_column}={record_id_value}",
        )

    issues_df = pd.DataFrame(issues, columns=[
        "Row", "Issue", "Column", "Value", "Severity", "Explanation", "Recommended Action",
    ])
    clean = data.loc[~duplicate_mask].copy()
    if parsed_dates is not None:
        clean[mapping["date"]] = parsed_dates.loc[clean.index].dt.strftime("%Y-%m-%d").fillna("")
    for field, values in numeric.items():
        clean[mapping[field]] = values.loc[clean.index]

    if "total_sales" in numeric:
        raw_sales = numeric["total_sales"].fillna(0.0)
    elif all(field in numeric for field in ("quantity", "unit_price")):
        raw_sales = (numeric["quantity"] * numeric["unit_price"]).fillna(0.0)
    else:
        raw_sales = pd.Series(0.0, index=data.index)
    validated_sales = raw_sales.loc[~duplicate_mask]
    raw_total_sales = float(raw_sales.sum())
    validated_total_sales = float(validated_sales.sum())

    if "order_id" in mapping:
        order_ids = data[mapping["order_id"]].dropna().astype(str).str.strip()
        unique_order_count = int(order_ids[order_ids.ne("")].nunique())
        order_count = unique_order_count
        order_label = "Unique orders"
        aov_denominator = unique_order_count
        identified_order_rows = data[mapping["order_id"]].notna() & data[mapping["order_id"]].astype(str).str.strip().ne("")
        aov_sales = float(validated_sales.loc[identified_order_rows.loc[validated_sales.index]].sum())
    else:
        unique_order_count = None
        order_count = len(clean)
        order_label = "Records"
        aov_denominator = len(clean)
        aov_sales = validated_total_sales
    average_order_value = aov_sales / aov_denominator if aov_denominator else None

    product_summary = _summary_by_field(data, mapping, "product", raw_sales, validated_sales, duplicate_mask, "(Missing product)")
    region_values_available = False
    if "region" in mapping:
        region_values = data[mapping["region"]]
        region_values_available = bool((region_values.notna() & region_values.astype(str).str.strip().ne("")).any())
    region_summary = (
        _summary_by_field(data, mapping, "region", raw_sales, validated_sales, duplicate_mask, "(Missing region)")
        if region_values_available else []
    )

    sales_trend: list[dict[str, Any]] = []
    sales_trend_period = "Date"
    if parsed_dates is not None:
        dated_mask = parsed_dates.notna()
        valid_dates = parsed_dates.loc[dated_mask]
        if not valid_dates.empty:
            if valid_dates.dt.normalize().nunique() > 31:
                trend_periods = valid_dates.dt.to_period("M").dt.to_timestamp()
                sales_trend_period = "Month"
            else:
                trend_periods = valid_dates.dt.normalize()
            trend_frame = pd.DataFrame({
                "Period": trend_periods,
                "Raw Sales": raw_sales.loc[valid_dates.index],
                "Validated Sales": validated_sales.reindex(valid_dates.index).fillna(0.0),
            }, index=valid_dates.index)
            trend_frame = trend_frame.groupby("Period", sort=True).sum(numeric_only=True).reset_index()
            sales_trend = [
                {"Period": row["Period"].to_pydatetime(), "Raw Sales": float(row["Raw Sales"]), "Validated Sales": float(row["Validated Sales"])}
                for _, row in trend_frame.iterrows()
            ]

    result = {
        "schema_version": SALES_RESULT_SCHEMA_VERSION,
        "source_name": source_name,
        "original_data": original_data,
        "data": data,
        "cleaned_data": clean.reset_index(drop=True),
        "mapping": mapping,
        "missing_columns": [field for field in REQUIRED_FOR_QUALITY if field not in mapping],
        "issues": issues_df,
        "raw_sales_by_row": raw_sales,
        "validated_sales_by_row": validated_sales,
        "raw_total_sales": raw_total_sales,
        "validated_total_sales": validated_total_sales,
        "total_sales": validated_total_sales,
        "record_count": len(data),
        "validated_record_count": len(clean),
        "order_count": order_count,
        "unique_order_count": unique_order_count,
        "order_label": order_label,
        "average_order_value": average_order_value,
        "aov_sales": aov_sales,
        "aov_denominator": aov_denominator,
        "top_products": [
            {"Product": item["Label"], "Sales": item["Validated Sales"]}
            for item in product_summary[:5]
        ],
        "product_summary": product_summary,
        "region_summary": region_summary,
        "region_available": region_values_available,
        "sales_trend": sales_trend,
        "sales_trend_period": sales_trend_period,
        "duplicate_count": int(duplicate_mask.sum()),
        "repeated_transaction_count": int(repeated_transaction_mask.sum()),
    }
    return validate_sales_result_schema(result)


def build_report(result: dict[str, Any]) -> str:
    """Build the existing small plain-text MIS report, now distinguishing sales bases."""
    validate_sales_result_schema(result)
    aov = result["average_order_value"]
    lines = [
        "MIS AUTOPILOT | SALES SUMMARY",
        "=" * 34,
        f"Source: {result['source_name']}",
        f"Raw total sales: INR {result['raw_total_sales']:,.2f}",
        f"Validated total sales: INR {result['validated_total_sales']:,.2f}",
        f"{result['order_label']}: {result['order_count']:,}",
        f"Average order value: {'INR ' + format(aov, ',.2f') if aov is not None else 'Not available'}",
        f"Data rows: {result['record_count']:,}",
        f"Data quality issues: {len(result['issues']):,}",
        f"Exact full-row duplicates excluded from validated totals: {result['duplicate_count']:,}",
        f"Repeated transaction candidates (record ID ignored): {result['repeated_transaction_count']:,}",
        "Candidates remain visible and included until reviewed; validated totals exclude exact full-row duplicates only.",
        "",
        "TOP PRODUCTS BY VALIDATED SALES",
        "-" * 32,
    ]
    if result["top_products"]:
        lines.extend(f"{index}. {item['Product']}: INR {item['Sales']:,.2f}" for index, item in enumerate(result["top_products"], start=1))
    else:
        lines.append("Product column was not recognized.")
    lines += ["", "COLUMN MAPPING", "-" * 14]
    lines.extend(f"{field.replace('_', ' ').title()}: {column}" for field, column in result["mapping"].items())
    if result["missing_columns"]:
        lines += ["", "Not recognized: " + ", ".join(field.replace("_", " ") for field in result["missing_columns"])]
    return "\n".join(lines) + "\n"
