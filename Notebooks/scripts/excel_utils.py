"""
Excel utilities for reading files with leading junk rows before the actual table.

Primary function:
- read_excel_table: Reads an Excel sheet, auto-detects the header row by scanning down
  until a plausible header is found, and returns a clean pandas DataFrame.

Design goals:
- Be resilient to a few empty/title rows before the table starts.
- Avoid hard-coding specific broker column names.
- Keep parameters simple but allow tuning when needed.
"""
from __future__ import annotations

from typing import Any, Iterable, List, Optional, Sequence, Tuple, Union

import pandas as pd
import warnings
import os
import re
import unicodedata


def _normalize_colname(val: Any, space_policy: str = "preserve") -> str:
    """Normalize a column name robustly (Unicode-aware).

    - Applies NFKC normalization to canonicalize characters.
    - Removes zero-width/formatting joiners (e.g., U+200B..U+200D, U+2060, U+FEFF).
    - Converts all Unicode space-separators (Zs), as well as ASCII whitespace, to ASCII spaces.
    - Collapses whitespace runs and strips ends.
    - Applies space_policy to internal spaces: "preserve" (default), "underscore", or "remove".
    """
    # 1) Coerce to string
    if pd.isna(val):
        s = ""
    else:
        s = str(val)

    # 2) Unicode normalize (canonical composition/compatibility)
    s = unicodedata.normalize("NFKC", s)

    # 3) Remove zero-width/formatting characters explicitly
    #    (ZWSP U+200B, ZWNJ U+200C, ZWJ U+200D, WJ U+2060, BOM/ZWNBSP U+FEFF)
    zero_width = {
        "\u200B",  # ZERO WIDTH SPACE
        "\u200C",  # ZERO WIDTH NON-JOINER
        "\u200D",  # ZERO WIDTH JOINER
        "\u2060",  # WORD JOINER
        "\ufeff",  # ZERO WIDTH NO-BREAK SPACE / BOM
    }
    # Also normalize the common NBSP to regular space first
    s = s.replace("\u00A0", " ")
    for ch in zero_width:
        s = s.replace(ch, "")

    # 4) Map all Unicode space separators to ASCII space
    #    (covers U+2000..U+200A, U+202F, U+205F, U+3000, etc.)
    def to_ascii_space(c: str) -> str:
        if c in "\t\n\r\v\f":
            return " "
        cat = unicodedata.category(c)
        if cat == "Zs":  # Separator, space
            return " "
        return c

    s = "".join(to_ascii_space(c) for c in s)

    # 5) Collapse whitespace and strip
    s = re.sub(r"\s+", " ", s).strip()

    # 6) Apply space policy
    if space_policy not in {"preserve", "underscore", "remove"}:
        space_policy = "preserve"
    if space_policy == "underscore":
        s = s.replace(" ", "_")
    elif space_policy == "remove":
        s = s.replace(" ", "")

    return s


def _non_empty_count(row: Sequence[Any]) -> int:
    return sum(1 for v in row if pd.notna(v) and str(v).strip() != "")


def _looks_like_header_row(row: Sequence[Any]) -> bool:
    """Heuristic to decide whether a row could be a header row.

    Criteria:
    - At least 2 non-empty cells.
    - Majority of non-empty cells contain non-numeric strings (labels).
    - Header values are relatively short (< 64 chars) and not all repeated.
    """
    values = [str(v).strip() for v in row if pd.notna(v) and str(v).strip() != ""]
    if len(values) < 2:
        return False
    # Most cells should be non-numeric-like
    def is_number_like(s: str) -> bool:
        try:
            float(s.replace(",", ""))
            return True
        except Exception:
            return False
    non_numeric = [v for v in values if not is_number_like(v)]
    if len(non_numeric) < max(2, (len(values) + 1) // 2):
        return False
    # Length check and uniqueness
    if any(len(v) > 64 for v in values):
        return False
    if len(set(values)) <= 1:
        return False
    return True


def _row_width(row: Sequence[Any]) -> int:
    """Approximate width as the rightmost non-empty cell index + 1."""
    width = 0
    for idx, v in enumerate(row):
        if pd.notna(v) and str(v).strip() != "":
            width = idx + 1
    return width


def read_excel_table(
    path: str,
    sheet_name: Union[int, str, None] = 0,
    max_scan_rows: int = 100,
    min_header_cols: int = 2,
    dtype: Optional[Any] = None,
    engine: Optional[str] = None,
    space_policy: str = "preserve",
) -> pd.DataFrame:
    """Read an Excel file that may have leading junk rows and return the table.

    The function reads the sheet with header=None first, scans the first
    `max_scan_rows` rows looking for a plausible header row, then returns a
    DataFrame with that row as headers and subsequent rows as data. Columns
    that are entirely empty are dropped. Header names are stripped of
    whitespace and deduplicated.

    Parameters:
    - path: Path to the Excel file.
    - sheet_name: Sheet name or index (passed to pandas). Default 0.
    - max_scan_rows: How many initial rows to scan for the header.
    - min_header_cols: Minimal number of non-empty cells required for a header.
    - dtype: Optional dtype for the returned DataFrame. If None, pandas infers.
    - engine: Optional engine override for pandas.read_excel.

    Returns:
    - pandas.DataFrame with cleaned header and data rows.

    Notes:
    - space_policy controls how internal spaces in header names are handled after
      Unicode normalization:
        * "preserve": keep single spaces (default; maintains backward compatibility)
        * "underscore": replace spaces with underscore
        * "remove": remove spaces entirely (e.g., "IP Address" -> "IPAddress")

    Raises:
    - ValueError if a header row cannot be determined.
    """
    # Determine engine if not provided based on file extension
    eff_engine = engine
    if eff_engine is None:
        ext = os.path.splitext(str(path))[1].lower()
        if ext == ".xls":
            eff_engine = "xlrd"
        else:
            # default for modern Excel formats
            eff_engine = "openpyxl"

    # Read everything as object first to analyze rows; suppress noisy openpyxl default-style warnings
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            category=UserWarning,
            message=r"Workbook contains no default style, apply openpyxl's default",
            module=r"openpyxl\.styles\.stylesheet",
        )
        # Try reading with the chosen engine; if the file is misnamed (e.g., XLSX content with .xls
        # extension), fall back to the alternate engine automatically.
        def _attempt_read(engine_name: Optional[str]):
            return pd.read_excel(
                path,
                sheet_name=sheet_name,
                header=None,
                dtype=object,
                engine=engine_name,
            )

        try:
            raw = _attempt_read(eff_engine)
        except Exception as e:
            msg = str(e)
            lower = msg.lower()
            fallback_raw = None
            # Case 1: tried xlrd but content is actually XLSX
            if (eff_engine == "xlrd" and ("xlsx file; not supported" in lower or "not supported" in lower and "xlsx" in lower)) or (
                eff_engine == "xlrd" and ("unsupported format, or corrupt file" in lower or "zip file" in lower)
            ):
                try:
                    fallback_raw = _attempt_read("openpyxl")
                except Exception:
                    pass
            # Case 2: tried openpyxl but content is actually legacy XLS
            elif eff_engine == "openpyxl" and ("file is not a zip file" in lower or "openpyxl does not support" in lower or "is not a zip" in lower):
                try:
                    fallback_raw = _attempt_read("xlrd")
                except Exception:
                    pass
            if fallback_raw is not None:
                raw = fallback_raw
            else:
                # Re-raise original error if no suitable fallback succeeded
                raise
    if raw.empty:
        raise ValueError("The Excel sheet is empty.")

    scan_limit = min(len(raw), max_scan_rows)

    header_idx: Optional[int] = None
    header_width: Optional[int] = None

    for i in range(scan_limit):
        row = list(raw.iloc[i].tolist())
        if _non_empty_count(row) < min_header_cols:
            continue
        if not _looks_like_header_row(row):
            continue
        # Check that at least one of the next two rows has some data in the same width
        width = _row_width(row)
        if width < min_header_cols:
            continue
        ok = False
        for j in (i + 1, i + 2):
            if j < len(raw):
                next_row = list(raw.iloc[j].tolist())
                # Consider it data-like if it has any non-empty value within width
                if _non_empty_count(next_row[:width]) >= 1:
                    ok = True
                    break
        if ok:
            header_idx = i
            header_width = width
            break

    if header_idx is None:
        # Fallback: try using the very first row as header
        first = list(raw.iloc[0].tolist())
        non_empty_first = _non_empty_count(first)
        if non_empty_first >= min_header_cols:
            header_idx = 0
            header_width = max(min_header_cols, _row_width(first))
        elif non_empty_first >= 1:
            # Relaxed fallback: allow a single non-empty header cell if followed by data
            width = _row_width(first)
            if width < 1:
                raise ValueError("Could not detect a header row in the Excel sheet.")
            ok = False
            for j in (1, 2):
                if j < len(raw):
                    next_row = list(raw.iloc[j].tolist())
                    if _non_empty_count(next_row[:width]) >= 1:
                        ok = True
                        break
            if ok:
                header_idx = 0
                header_width = width
            else:
                raise ValueError("Could not detect a header row in the Excel sheet.")
        else:
            raise ValueError("Could not detect a header row in the Excel sheet.")

    # Before building DataFrame, expand header_width to include immediate data width if needed
    if header_width is None:
        header_width = _row_width(list(raw.iloc[header_idx].tolist()))
    # Consider the next 1-2 rows to capture data width beyond header's visible width
    data_width = 0
    for j in (header_idx + 1, header_idx + 2):
        if j < len(raw):
            data_width = max(data_width, _row_width(list(raw.iloc[j].tolist())))
    header_width = max(header_width, data_width)

    # Build DataFrame starting after header_idx
    data = raw.iloc[header_idx + 1 :].copy()
    # Trim columns to header_width and drop fully-empty columns
    if header_width is not None:
        data = data.iloc[:, :header_width]

    # Extract raw header values without coercing NaNs to the string 'nan'
    header_vals = raw.iloc[header_idx].tolist()[: data.shape[1]]
    # Clean header names (robust stripping and normalization)
    clean_headers: List[str] = []
    seen = {}
    for h in header_vals:
        h = _normalize_colname(h, space_policy=space_policy)
        if h == "":
            h = "Unnamed"
        base = h
        k = 1
        while h in seen:
            k += 1
            h = f"{base}_{k}"
        seen[h] = True
        clean_headers.append(h)

    # Assign and ensure final column names are normalized/stripped
    data.columns = [ _normalize_colname(c, space_policy=space_policy) for c in clean_headers ]
    # Drop rows that are completely empty
    data = data.dropna(how="all")

    # Let pandas infer dtypes if requested (dtype=None) or enforce dtype
    if dtype is not None:
        try:
            data = data.astype(dtype)
        except Exception:
            # If dtype casting fails, keep original types for robustness
            pass
    else:
        data = data.convert_dtypes()

    return data


__all__ = ["read_excel_table"]
