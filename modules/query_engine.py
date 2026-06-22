"""
query_engine.py — Natural language query engine over the results DataFrame.
Implemented entirely with Pandas pattern matching — no external AI APIs.
"""

import re
import pandas as pd


def _normalise(text: str) -> str:
    return text.lower().strip()


def _extract_day(text: str) -> str | None:
    """Extract a day number like '15', '05', '1' from query text."""
    m = re.search(r'\b(\d{1,2})\b', text)
    return m.group(1).zfill(2) if m else None


def _extract_station(text: str, df: pd.DataFrame) -> str | None:
    """Find a station name mentioned in the query."""
    stations = df["Station"].dropna().unique()
    for st in stations:
        if st.lower() in text.lower():
            return st
    return None


def run_query(query: str, df: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    """
    Parse a natural language query and return (filtered_df, explanation).
    Supports patterns:
        - tested on [day]
        - manual review / not tested
        - station [NAME]
        - most reviews / highest review count
        - all points
        - tested / all tested
    """
    q = _normalise(query)
    explanation = ""

    # ── "tested on [day]" ────────────────────────────────────────────────────
    if re.search(r'tested.{0,15}on|on.{0,5}day|date\s+\d', q):
        day = _extract_day(q)
        if day:
            mask = df["Dates"].apply(
                lambda d: day in [x.zfill(2) for x in str(d).split(",") if x.strip()]
            ) & (df["Status"] == "Tested")
            result = df[mask]
            explanation = f"Points tested on day {day}."
            return result, explanation
        return df[df["Status"] == "Tested"], "All tested points (no day specified)."

    # ── "manual review" / "review" ────────────────────────────────────────────
    if any(kw in q for kw in ["manual review", "review", "not tested", "unverified", "exception"]):
        # Check if asking which station has MOST reviews
        if any(kw in q for kw in ["most", "highest", "maximum", "max", "top"]):
            counts = df[df["Status"] == "Manual Review Required"].groupby("Station").size()
            if counts.empty:
                return pd.DataFrame(), "No manual review records found."
            top_station = counts.idxmax()
            explanation = f"Station '{top_station}' has the most manual review records ({counts.max()})."
            result = df[(df["Station"] == top_station) & (df["Status"] == "Manual Review Required")]
            return result, explanation

        result = df[df["Status"] == "Manual Review Required"]
        explanation = "All records marked as Manual Review Required."
        return result, explanation

    # ── "station [NAME]" / activity for a station ────────────────────────────
    station = _extract_station(q, df)
    if station:
        result = df[df["Station"].str.upper() == station.upper()]
        explanation = f"All records for station '{station}'."
        return result, explanation

    # ── "all tested" / "tested points" ───────────────────────────────────────
    if any(kw in q for kw in ["all tested", "tested points", "show tested", "list tested"]):
        result = df[df["Status"] == "Tested"]
        explanation = "All records classified as Tested."
        return result, explanation

    # ── "all points" / broad "show all" ──────────────────────────────────────
    if any(kw in q for kw in ["all points", "show all", "everything", "full report", "complete"]):
        return df.copy(), "Full report — all records."

    # ── Fallback: fuzzy keyword search across string columns ─────────────────
    tokens = [t for t in q.split() if len(t) > 2]
    if tokens:
        mask = pd.Series([False] * len(df), index=df.index)
        for token in tokens:
            for col in df.select_dtypes(include="object").columns:
                mask |= df[col].astype(str).str.contains(token, case=False, na=False)
        result = df[mask]
        if not result.empty:
            explanation = f"Records matching keywords: {', '.join(tokens)}."
            return result, explanation

    explanation = (
        "Query not understood. Try: 'show all manual review points', "
        "'points tested on 15', 'show station DPH', 'which station has most reviews'."
    )
    return pd.DataFrame(), explanation
