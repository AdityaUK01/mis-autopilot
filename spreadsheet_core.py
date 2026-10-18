"""Format-agnostic spreadsheet loading, profiling, and basic validation.

This module deliberately retains source cell values. Its ``data``/``original_data``
members are the uploaded table with only headers made unique when necessary.
Issues are annotations; no suspicious record is removed.
"""

from __future__ import annotations

from collections import Counter
from io import BytesIO
import math
import re
from typing import Any

import pandas as pd


PROFILE_ISSUE_COLUMNS = [
    "Row", "Column", "Issue", "Original Value", "Severity", "Explanation", "Recommended Action",
]
PROFILE_RESULT_SCHEMA_VERSION = 1
PROFILE_RESULT_REQUIRED_KEYS = frozenset({
    "schema_version", "source_name", "original_data", "data", "row_count",
    "nonblank_row_count", "blank_row_count", "column_count", "column_profile",
    "detected_types", "numeric_summary", "category_summaries", "issues", "duplicate_count",
})


def validate_profile_schema(profile: dict[str, Any]) -> dict[str, Any]:
    missing = sorted(PROFILE_RESULT_REQUIRED_KEYS.difference(profile))
    if profile.get("schema_version") != PROFILE_RESULT_SCHEMA_VERSION or missing:
        details = []
        if profile.get("schema_version") != PROFILE_RESULT_SCHEMA_VERSION:
            details.append(f"expected schema version {PROFILE_RESULT_SCHEMA_VERSION}, got {profile.get('schema_version')!r}")
        if missing:
            details.append("missing keys: " + ", ".join(missing))
        raise ValueError("Invalid spreadsheet profile result schema (" + "; ".join(details) + ").")
    return profile


def unique_headers(headers: list[Any]) -> list[str]:
    """Return readable unique headers without dropping source columns."""
    counts: Counter[str] = Counter()
    output: list[str] = []
    for index, raw in enumerate(headers, start=1):
        label = " ".join(str(raw if raw is not None else "").split()) or f"Column {index}"
        counts[label] += 1
        output.append(label if counts[label] == 1 else f"{label} [{counts[label]}]")
    return output


def _headerless_to_frame(table: pd.DataFrame) -> pd.DataFrame:
    if table.empty and len(table.columns) == 0:
        return pd.DataFrame()
    rows = table.values.tolist()
    if not rows:
        return pd.DataFrame()
    headers = unique_headers(list(rows[0]))
    body = rows[1:]
    return pd.DataFrame(body, columns=headers).reset_index(drop=True)


def load_spreadsheet(payload: bytes, filename: str, sheet_name: str | None = None) -> tuple[pd.DataFrame, list[str]]:
    """Load CSV/XLS/XLSX/XLSM preserving values and duplicate header columns.

    Returns ``(data, worksheet_names)``. Excel files are read headerless so
    duplicate source headings can be retained as uniquely labelled columns.
    """
    suffix = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    if suffix not in {"csv", "xlsx", "xlsm", "xls"}:
        raise ValueError("Unsupported file format. Upload a .csv, .xlsx, .xlsm, or .xls workbook.")
    if not payload:
        raise ValueError("The selected file is empty.")

    if suffix == "csv":
        from pandas.errors import EmptyDataError

        try:
            try:
                table = pd.read_csv(BytesIO(payload), header=None, dtype=object, keep_default_na=False, skip_blank_lines=False)
            except UnicodeDecodeError:
                table = pd.read_csv(BytesIO(payload), header=None, encoding="latin-1", dtype=object, keep_default_na=False, skip_blank_lines=False)
        except EmptyDataError:
            raise ValueError("The CSV file has no header row or data.") from None
        except Exception as exc:
            raise ValueError("The CSV could not be read. Check that it is a valid CSV with consistent rows.") from exc
        frame = _headerless_to_frame(table)
        if len(frame.columns) == 0:
            raise ValueError("The CSV file has no header row or data.")
        return frame, []

    engine = "xlrd" if suffix == "xls" else "openpyxl"
    try:
        excel = pd.ExcelFile(BytesIO(payload), engine=engine)
    except ImportError as exc:
        if suffix == "xls":
            raise ValueError("Reading .xls files needs xlrd. Install the project requirements, then restart the app.") from exc
        raise
    except Exception as exc:
        raise ValueError("The Excel workbook could not be opened. Check that it is a valid, unprotected workbook.") from exc
    sheets = list(excel.sheet_names)
    if not sheets:
        raise ValueError("The workbook does not contain a worksheet.")
    selected = sheet_name or sheets[0]
    if selected not in sheets:
        raise ValueError(f"Worksheet {selected!r} is not in this workbook. Choose one of: {', '.join(sheets)}")
    try:
        table = pd.read_excel(excel, sheet_name=selected, header=None, dtype=object, keep_default_na=False)
    except ImportError as exc:
        if suffix == "xls":
            raise ValueError("Reading .xls files needs xlrd. Install the project requirements, then restart the app.") from exc
        raise
    except Exception as exc:
        raise ValueError(f"Worksheet {selected!r} could not be read. Check that the workbook is not damaged or password-protected.") from exc
    return _headerless_to_frame(table), sheets


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(value, str) and not value.strip()


def _num(value: Any) -> float | None:
    if _is_blank(value) or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        try:
            number = float(value)
            return number if math.isfinite(number) else None
        except (ValueError, OverflowError):
            return None
    text = str(value).strip().replace(",", "")
    text = re.sub(r"^[^\d+\-.]+", "", text)
    text = re.sub(r"[^\d.\-]+$", "", text)
    try:
        number = float(text)
        return number if math.isfinite(number) else None
    except (ValueError, OverflowError):
        return None


def _date_header(name: str) -> bool:
    normalized = re.sub(r"[^a-z0-9]+", " ", name.lower()).strip()
    return bool(re.search(r"\b(date|datetime|timestamp|time|period)\b", normalized))


def _identifier_header(name: str) -> bool:
    normalized = re.sub(r"[^a-z0-9]+", " ", name.lower()).strip()
    return bool(re.search(r"\b(id|code|phone|mobile|postal|pincode|zip)\b", normalized))


def _numeric_header(name: str) -> bool:
    normalized = re.sub(r"[^a-z0-9]+", " ", name.lower()).strip()
    words = {"amount", "price", "sales", "revenue", "cost", "quantity", "qty", "units", "stock", "rate", "hours", "salary", "wage", "budget", "actual", "target", "count", "expense", "balance", "total"}
    return bool(words.intersection(normalized.split()))


def _type_and_values(series: pd.Series, name: str) -> tuple[str, pd.Series | None, pd.Series | None]:
    nonblank = series.loc[~series.map(_is_blank)]
    if nonblank.empty:
        return "empty", None, None
    if _date_header(name):
        parsed = pd.to_datetime(series, errors="coerce", dayfirst=True, format="mixed")
        return "date", None, parsed
    numeric_values = series.map(_num)
    numeric_count = int(numeric_values.loc[~series.map(_is_blank)].notna().sum())
    ratio = numeric_count / len(nonblank)
    if not _identifier_header(name) and (
        pd.api.types.is_numeric_dtype(series) or _numeric_header(name) or ratio >= 0.8
    ):
        return "numeric", numeric_values, None
    parsed = pd.to_datetime(series, errors="coerce", dayfirst=True, format="mixed")
    parsed_count = int(parsed.loc[~series.map(_is_blank)].notna().sum())
    if parsed_count / len(nonblank) >= 0.8:
        return "date", None, parsed
    return "text", None, None


def profile_spreadsheet(data: pd.DataFrame, source_name: str = "Uploaded data") -> dict[str, Any]:
    """Profile/validate a generic table; preserve all rows and original values."""
    original = data.copy().reset_index(drop=True)
    if original.columns.has_duplicates or any(not str(column).strip() for column in original.columns):
        original.columns = unique_headers(list(original.columns))
    issues: list[dict[str, Any]] = []

    def add_issue(row_index: int, column: str, issue: str, value: Any, severity: str, explanation: str, action: str) -> None:
        issues.append({
            "Row": row_index + 2,
            "Column": column,
            "Issue": issue,
            "Original Value": "" if _is_blank(value) else str(value),
            "Severity": severity,
            "Explanation": explanation,
            "Recommended Action": action,
        })

    row_nonblank = pd.Series(False, index=original.index)
    for col in original.columns:
        row_nonblank |= ~original[col].map(_is_blank)
    blank_rows = int((~row_nonblank).sum())
    profile_rows: list[dict[str, Any]] = []
    type_map: dict[str, str] = {}
    numeric_summary: list[dict[str, Any]] = []
    categorical_summaries: dict[str, pd.DataFrame] = {}

    for column in original.columns:
        values = original[column]
        blanks = values.map(_is_blank)
        detected, parsed_numbers, parsed_dates = _type_and_values(values, str(column))
        type_map[str(column)] = detected
        nonblank = values.loc[~blanks]
        profile_rows.append({
            "Column": str(column),
            "Detected type": detected,
            "Non-empty values": int((~blanks).sum()),
            "Missing values": int(blanks.sum()),
            "Distinct values": int(nonblank.nunique(dropna=False)),
            "Sample values": ", ".join(str(value)[:40] for value in nonblank.head(3)),
        })
        for index in original.index[blanks & row_nonblank]:
            add_issue(index, str(column), "Missing value", values.at[index], "Warning", "A nonblank record has no value in this column.", "Check the source record; fill the value only when it can be verified.")
        if detected == "numeric" and parsed_numbers is not None:
            nonblank_bad = ~blanks & parsed_numbers.isna()
            for index in original.index[nonblank_bad]:
                add_issue(index, str(column), "Invalid numeric value", values.at[index], "Error", "This value is not a usable number for a column inferred as numeric.", "Check the source and enter a numeric value; keep the original value for review.")
            usable = parsed_numbers.dropna()
            if not usable.empty:
                numeric_summary.append({
                    "Column": str(column), "Records": len(usable), "Sum": float(usable.sum()),
                    "Average": float(usable.mean()), "Minimum": float(usable.min()), "Maximum": float(usable.max()),
                })
        elif detected == "date" and parsed_dates is not None:
            nonblank_bad = ~blanks & parsed_dates.isna()
            for index in original.index[nonblank_bad]:
                add_issue(index, str(column), "Invalid date", values.at[index], "Error", "This value could not be interpreted as a valid date.", "Correct it to a real date after checking the source.")
        elif detected == "text" and nonblank.nunique() <= min(100, max(20, len(nonblank) // 3)):
            counts = nonblank.astype(str).value_counts().head(10)
            categorical_summaries[str(column)] = counts.rename_axis("Category").reset_index(name="Records")

    if len(original.columns):
        duplicate_mask = original.loc[row_nonblank].duplicated(keep="first")
        duplicate_indices = original.loc[row_nonblank].index[duplicate_mask]
    else:
        duplicate_indices = pd.Index([])
    for index in duplicate_indices:
        add_issue(index, "Entire row", "Exact duplicate row", " | ".join(str(v) for v in original.loc[index].tolist()), "Warning", "This row exactly matches an earlier nonblank source row. It has been retained.", "Compare both source records; exclude only after confirming duplication.")

    result = {
        "schema_version": PROFILE_RESULT_SCHEMA_VERSION,
        "source_name": source_name,
        "original_data": original,
        "data": original,
        "row_count": len(original),
        "nonblank_row_count": int(row_nonblank.sum()),
        "blank_row_count": blank_rows,
        "column_count": len(original.columns),
        "column_profile": pd.DataFrame(profile_rows, columns=["Column", "Detected type", "Non-empty values", "Missing values", "Distinct values", "Sample values"]),
        "detected_types": type_map,
        "numeric_summary": pd.DataFrame(numeric_summary, columns=["Column", "Records", "Sum", "Average", "Minimum", "Maximum"]),
        "category_summaries": categorical_summaries,
        "issues": pd.DataFrame(issues, columns=PROFILE_ISSUE_COLUMNS),
        "duplicate_count": len(duplicate_indices),
    }
    return validate_profile_schema(result)
