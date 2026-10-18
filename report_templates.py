"""Guided, explicit business-template calculations over generic spreadsheet data."""

from __future__ import annotations

import math
import re
from typing import Any

import pandas as pd


REPORT_TYPES = [
    "Sales",
    "Finance",
    "HR and attendance",
    "Inventory",
    "Operations",
    "General spreadsheet analysis",
]
TEMPLATE_RESULT_SCHEMA_VERSION = 1
TEMPLATE_RESULT_REQUIRED_KEYS = frozenset({
    "schema_version", "template", "mapping", "ready", "missing_requirements",
    "metrics", "summaries", "notes", "sales_result",
})


def validate_template_result_schema(result: dict[str, Any]) -> dict[str, Any]:
    missing = sorted(TEMPLATE_RESULT_REQUIRED_KEYS.difference(result))
    if result.get("schema_version") != TEMPLATE_RESULT_SCHEMA_VERSION or missing:
        details = []
        if result.get("schema_version") != TEMPLATE_RESULT_SCHEMA_VERSION:
            details.append(f"expected schema version {TEMPLATE_RESULT_SCHEMA_VERSION}, got {result.get('schema_version')!r}")
        if missing:
            details.append("missing keys: " + ", ".join(missing))
        raise ValueError("Invalid template analysis result schema (" + "; ".join(details) + ").")
    return result


ROLE_ALIASES: dict[str, dict[str, set[str]]] = {
    "Sales": {
        "date": {"date", "order date", "sale date", "sales date", "invoice date", "transaction date", "bill date"},
        "product": {"product", "product name", "item", "item name", "sku", "description", "product description"},
        "quantity": {"quantity", "qty", "units", "unit sold", "units sold", "quantity sold"},
        "unit_price": {"unit price", "price", "selling price", "sale price", "rate", "price per unit"},
        "total_sales": {"total", "sales", "sales amount", "total amount", "line total", "revenue", "amount", "total sales", "net sales"},
        "region": {"region", "area", "territory", "state", "city", "location", "market"},
        "order_id": {"order id", "order number", "order no", "invoice number", "invoice no", "transaction id", "sales id", "bill number", "bill no"},
    },
    "Finance": {
        "expense": {"expense", "expenses", "expense amount", "total expense", "total expenses", "spend"},
        "amount": {"amount", "transaction amount", "amount paid", "amount spent"},
        "actual": {"actual", "actual amount", "actual value", "actual spend", "actual expenses"},
        "budget": {"budget", "budget amount", "planned amount", "planned spend"},
        "category": {"category", "expense category", "account category", "cost category", "ledger category"},
        "date": {"date", "transaction date", "invoice date", "payment date"},
    },
    "HR and attendance": {
        "employee": {"employee id", "staff id", "employee name", "staff name", "employee", "staff", "person"},
        "attendance_status": {"attendance status", "attendance", "status", "present absent", "presence"},
        "date": {"date", "attendance date", "work date", "shift date"},
    },
    "Inventory": {
        "item": {"item", "item name", "product", "product name", "sku", "stock item"},
        "stock_quantity": {"stock quantity", "quantity on hand", "current stock", "stock on hand", "available quantity", "qty available"},
        "threshold": {"reorder level", "minimum stock", "min stock", "stock threshold", "reorder point", "safety stock"},
        "region": {"warehouse", "location", "store", "region"},
    },
    "Operations": {
        "actual": {"actual", "actual output", "completed", "completed units", "actual quantity", "actual amount", "output actual"},
        "target": {"target", "target output", "planned", "planned output", "target quantity", "target amount"},
        "status": {"status", "task status", "work status", "completion status"},
        "owner": {"owner", "employee", "staff", "team", "operator", "assigned to"},
        "date": {"date", "work date", "operation date", "shift date"},
    },
    "General spreadsheet analysis": {},
}


def _normalize(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()


def mapping_candidates(columns: list[Any] | pd.Index, report_type: str) -> dict[str, list[str]]:
    """Suggest exact/clear alias mappings; leave ambiguity for user selection."""
    aliases = ROLE_ALIASES.get(report_type, {})
    normalized = {str(column): _normalize(column) for column in columns}
    output: dict[str, list[str]] = {}
    for role, role_aliases in aliases.items():
        matches = [column for column, header in normalized.items() if header in role_aliases]
        if not matches:
            # The known role phrase may be embedded in a descriptive heading,
            # but don't match generic words such as "amount" or "date" loosely.
            long_aliases = {alias for alias in role_aliases if len(alias.split()) > 1}
            matches = [
                column for column, header in normalized.items()
                if any(header.startswith(alias + " ") or header.endswith(" " + alias) for alias in long_aliases)
            ]
        output[role] = matches
    return output


def default_mappings(candidates: dict[str, list[str]]) -> dict[str, str]:
    """Only auto-map a role when one unambiguous candidate exists."""
    return {role: matches[0] for role, matches in candidates.items() if len(matches) == 1}


def _number(value: Any) -> float | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        return number if math.isfinite(number) else None
    text = re.sub(r"[^\d+\-.]", "", str(value).replace(",", ""))
    try:
        number = float(text)
        return number if math.isfinite(number) else None
    except (ValueError, OverflowError):
        return None


def _numeric_series(data: pd.DataFrame, column: str) -> pd.Series:
    return data[column].map(_number)


def _clean_group(values: pd.Series, missing_label: str = "(Missing)") -> pd.Series:
    return values.map(lambda value: missing_label if pd.isna(value) or not str(value).strip() else str(value).strip())


def analyze_template(
    data: pd.DataFrame,
    report_type: str,
    mapping: dict[str, str] | None = None,
    source_name: str = "Uploaded data",
    profile: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Calculate only metrics supported by mapped source columns."""
    mapping = {role: column for role, column in (mapping or {}).items() if column in data.columns}
    result: dict[str, Any] = {
        "schema_version": TEMPLATE_RESULT_SCHEMA_VERSION,
        "template": report_type,
        "mapping": mapping,
        "ready": True,
        "missing_requirements": [],
        "metrics": [],
        "summaries": {},
        "notes": [],
        "sales_result": None,
    }

    def metric(label: str, value: Any, kind: str = "number", note: str = "") -> None:
        result["metrics"].append({"label": label, "value": value, "kind": kind, "note": note})

    def require(condition: bool, message: str) -> bool:
        if not condition:
            result["ready"] = False
            result["missing_requirements"].append(message)
        return condition

    if report_type == "Sales":
        have_sales = "total_sales" in mapping or ("quantity" in mapping and "unit_price" in mapping)
        require(have_sales, "Map a Total Sales column, or map both Quantity and Unit Price.")
        if have_sales:
            from mis_core import analyze_sales

            sales_result = analyze_sales(data, source_name, mapping_override=mapping)
            result["sales_result"] = sales_result
            metric("Total sales", sales_result["validated_total_sales"], "currency", "Exact full-row duplicates excluded; other records remain included.")
            metric(sales_result["order_label"], sales_result["order_count"], "number")
            metric("Average order value", sales_result["average_order_value"], "currency", "Per-record average when no order ID is mapped.")
            metric("Total records", sales_result["record_count"], "number")
            result["summaries"]["Product Summary"] = pd.DataFrame(sales_result["product_summary"])
            if sales_result["region_summary"]:
                result["summaries"]["Region Summary"] = pd.DataFrame(sales_result["region_summary"])
            result["notes"].append("Sales totals use the explicitly mapped total column, or quantity × unit price when no total column is mapped. Exact duplicate rows remain in Sales Data and are excluded from validated summary totals.")
        return validate_template_result_schema(result)

    if report_type == "Finance":
        value_role = next((role for role in ("expense", "actual", "amount") if role in mapping), None)
        require(value_role is not None, "Map an Expense, Actual, or Amount column to calculate finance totals.")
        if value_role:
            column = mapping[value_role]
            amounts = _numeric_series(data, column)
            label = {"expense": "Mapped expense total", "actual": "Mapped actual total", "amount": "Mapped amount total"}[value_role]
            amounts_total = amounts.sum(min_count=1)
            metric(label, float(amounts_total) if pd.notna(amounts_total) else None, "currency", "Uses the selected source column; no expense/income sign assumptions are made.")
            metric("Records with numeric values", int(amounts.notna().sum()), "number")
            if "budget" in mapping and "actual" in mapping:
                budget_column = mapping["budget"]
                actual_column = mapping.get("actual", column)
                budget_values = _numeric_series(data, budget_column)
                actual_values = _numeric_series(data, actual_column)
                frame = pd.DataFrame({"Budget": budget_values, "Actual": actual_values}).dropna()
                if not frame.empty:
                    total_budget, total_actual = float(frame["Budget"].sum()), float(frame["Actual"].sum())
                    metric("Budget total", total_budget, "currency")
                    metric("Actual total", total_actual, "currency")
                    metric("Actual less budget", total_actual - total_budget, "currency")
                    metric("Budget used", total_actual / total_budget if total_budget else None, "percent", "Actual divided by budget; unavailable when budget total is zero.")
            if "category" in mapping:
                groups = _clean_group(data[mapping["category"]])
                result["summaries"]["Category Totals"] = pd.DataFrame({"Category": groups, "Records": 1, "Total": amounts.fillna(0)}).groupby("Category", as_index=False, sort=False).agg({"Records": "sum", "Total": "sum"}).sort_values("Total", ascending=False, kind="stable")
        result["notes"].append("All source rows, including duplicate candidates, remain included in finance metrics. Totals use only numeric values in the confirmed mapping.")
        return validate_template_result_schema(result)

    if report_type == "HR and attendance":
        employee_column = mapping.get("employee")
        require(employee_column is not None, "Map an Employee ID, Employee Name, or equivalent field to calculate headcount.")
        if employee_column:
            employees = data[employee_column].map(lambda value: str(value).strip() if not pd.isna(value) else "")
            employees = employees.loc[employees.ne("")]
            metric("Unique employees", int(employees.nunique()), "number", f"Distinct values in {employee_column}.")
        status_column = mapping.get("attendance_status")
        if status_column:
            statuses = _clean_group(data[status_column])
            status_summary = statuses.value_counts(dropna=False).rename_axis("Attendance status").reset_index(name="Records")
            result["summaries"]["Attendance Summary"] = status_summary
            metric("Attendance records", int(statuses.ne("(Missing)").sum()), "number")
            absent_values = {"absent", "not present", "unattended", "no show"}
            absent_count = int(statuses.map(lambda value: _normalize(value) in absent_values).sum())
            if any(_normalize(value) in absent_values for value in statuses):
                metric("Absence records", absent_count, "number", "Counts only explicit absence status labels.")
            else:
                result["notes"].append("No explicit absence status label was found; an absence count is not inferred.")
        else:
            result["notes"].append("Map an attendance/status column to show attendance and absence summaries.")
        return validate_template_result_schema(result)

    if report_type == "Inventory":
        item_column = mapping.get("item")
        stock_column = mapping.get("stock_quantity")
        require(item_column is not None, "Map an Item/Product/SKU field to summarize inventory.")
        require(stock_column is not None, "Map an explicit Stock Quantity / Quantity on Hand field.")
        if item_column and stock_column:
            stock = _numeric_series(data, stock_column)
            stock_total = stock.sum(min_count=1)
            metric("Total mapped stock units", float(stock_total) if pd.notna(stock_total) else None, "number", "Sum of numeric stock quantities; source rows are retained.")
            metric("Items with numeric stock", int(stock.notna().sum()), "number")
            groups = _clean_group(data[item_column])
            summary_frame = pd.DataFrame({"Item": groups, "Records": 1, "Stock Quantity": stock.fillna(0)})
            summary = summary_frame.groupby("Item", as_index=False, sort=False).agg({"Records": "sum", "Stock Quantity": "sum"})
            threshold_column = mapping.get("threshold")
            if threshold_column:
                threshold = _numeric_series(data, threshold_column)
                low = stock.notna() & threshold.notna() & (stock <= threshold)
                summary_frame["Low Stock Records"] = low.astype(int)
                metric("Low-stock records", int(low.sum()), "number", "Stock quantity is at or below the mapped threshold.")
                summary = summary.merge(summary_frame.groupby("Item", as_index=False)["Low Stock Records"].sum(), on="Item", how="left")
                result["notes"].append(f"Low stock uses {stock_column} ≤ {threshold_column}.")
            else:
                result["notes"].append("Map a reorder/minimum-stock threshold to calculate low-stock indicators.")
            result["summaries"]["Inventory Summary"] = summary.sort_values("Stock Quantity", ascending=False, kind="stable")
        return validate_template_result_schema(result)

    if report_type == "Operations":
        actual_column, target_column = mapping.get("actual"), mapping.get("target")
        require(actual_column is not None, "Map an explicit Actual/Completed output field.")
        require(target_column is not None, "Map a Target/Planned output field.")
        if actual_column and target_column:
            actual = _numeric_series(data, actual_column)
            target = _numeric_series(data, target_column)
            pairs = pd.DataFrame({"Actual": actual, "Target": target}).dropna()
            if not pairs.empty:
                actual_total, target_total = float(pairs["Actual"].sum()), float(pairs["Target"].sum())
                metric("Actual total", actual_total, "number")
                metric("Target total", target_total, "number")
                metric("Target variance", actual_total - target_total, "number")
                metric("Target achievement", actual_total / target_total if target_total else None, "percent", "Actual divided by target; unavailable when target total is zero.")
                if "status" in mapping:
                    result["summaries"]["Status Summary"] = _clean_group(data[mapping["status"]]).value_counts().rename_axis("Status").reset_index(name="Records")
                if "owner" in mapping:
                    groups = _clean_group(data[mapping["owner"]])
                    result["summaries"]["Output by Owner"] = pd.DataFrame({"Owner": groups, "Actual": actual.fillna(0), "Target": target.fillna(0)}).groupby("Owner", as_index=False, sort=False).sum(numeric_only=True)
            else:
                result["notes"].append("No rows contain both a numeric actual and target value.")
        result["notes"].append("Operational productivity is reported only for the explicitly mapped actual and target fields.")
        return validate_template_result_schema(result)

    # General spreadsheet analysis needs no business-field mappings.
    require(report_type == "General spreadsheet analysis", "Choose a supported report template.")
    metric("Rows", len(data), "number")
    metric("Columns", len(data.columns), "number")
    missing_cells = int(profile["column_profile"]["Missing values"].sum()) if profile is not None and not profile["column_profile"].empty else 0
    metric("Missing cells", missing_cells, "number")
    if profile is not None:
        if not profile["numeric_summary"].empty:
            result["summaries"]["Numeric Summary"] = profile["numeric_summary"].copy()
        for column, summary in profile["category_summaries"].items():
            result["summaries"][f"{column} Breakdown"] = summary.copy()
    result["notes"].append("General summaries are descriptive and include every source record. Use the data-quality table to review questionable values.")
    return validate_template_result_schema(result)
