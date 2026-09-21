from __future__ import annotations

import io
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Image as RLImage
from reportlab.platypus import (
    KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

st.set_page_config(page_title="Bank Market Monitor", page_icon="🏦", layout="wide")

DEFAULT_FILE = Path(__file__).with_name("Bank Monitoring Model .xlsx")
PERIODS = ["5 day", "2 week", "3 month", "1 year"]
HEATMAP_SORT_OPTIONS = [
    "Largest absolute move", "5 day", "2 week", "3 month", "1 year",
    "Alphabetical", "Region",
]

# --- Shared visual identity -------------------------------------------------
NAVY = "#17365D"
INK = "#33404D"
MUTED = "#7A8794"
RULE = "#D9DEE4"
GRID = "#EDF0F3"
PANEL = "#F2F5F8"
GOOD = "#548235"
BAD = "#C00000"
WATCH = "#C55A11"
NEUTRAL = "#7F7F7F"

REGION_COLORS = {
    "United States": "#17365D",
    "Canada": "#2E75B6",
    "Euro Area": "#8FB8DE",
    "United Kingdom": "#7030A0",
    "Other Europe": "#70AD47",
    "Japan": "#C55A11",
    "Other Asia": "#ED7D31",
    "Other": "#7F7F7F",
}

REGION_ALIASES = {
    "america": "United States",
    "us": "United States",
    "u.s.": "United States",
    "united states": "United States",
    "canada": "Canada",
    "euro area": "Euro Area",
    "eurozone": "Euro Area",
    "other euro": "Other Europe",
    "other europe": "Other Europe",
    "u.k.": "United Kingdom",
    "uk": "United Kingdom",
    "united kingdom": "United Kingdom",
    "japan": "Japan",
    "asia": "Other Asia",
    "asia/pacific": "Other Asia",
    "asia-pacific": "Other Asia",
}

BANK_REGION_OVERRIDES = {
    "Citigroup": "United States",
    "Bank of America": "United States",
    "JPMorgan Chase": "United States",
    "Goldman Sachs": "United States",
    "Morgan Stanley": "United States",
    "Wells Fargo": "United States",
    "State Street Bank": "United States",
    "BNP Paribas SA": "Euro Area",
    "Credit Agricole SA": "Euro Area",
    "Societe Generale": "Euro Area",
    "Deutsche Bank AG": "Euro Area",
    "Commerzbank": "Euro Area",
    "UBS Group AG": "Other Europe",
    "Danske": "Other Europe",
    "Barclays Bank PLC": "United Kingdom",
    "HSBC Holdings PLC": "United Kingdom",
    "RBS Group PLC": "United Kingdom",
    "Standard Chartered": "United Kingdom",
    "Mitsubishi": "Japan",
    "Sumitomo": "Japan",
    "Mizuho": "Japan",
    "Daiwa": "Japan",
    "Nomura": "Japan",
    "Bank of Nova Scotia": "Canada",
    "BMO": "Canada",
    "CIBC": "Canada",
    "TD": "Canada",
    "RBC": "Canada",
    "Nationale": "Canada",
    "Laurentian": "Canada",
}

COLUMN_MAP = {
    0: "Bank",
    1: "Region Raw",
    2: "CDS Now",
    3: "CDS 5 day",
    4: "CDS 2 week",
    5: "CDS 3 month",
    6: "CDS 1 year",
    7: "Equity 5 day",
    8: "Equity 2 week",
    9: "Equity 3 month",
    10: "Equity 1 year",
}

METRIC_INFO = {
    "CDS": {
        "columns": {p: f"CDS {p}" for p in PERIODS},
        "unit": "bps",
        "description": "Positive changes indicate wider bank CDS spreads and higher perceived credit risk.",
    },
    "Equity": {
        "columns": {p: f"Equity {p}" for p in PERIODS},
        "unit": "%",
        "description": "Negative returns indicate weaker bank equity performance.",
    },
}

DEFAULT_CDS_THRESHOLDS = {"5 day": 3.0, "2 week": 3.0, "3 month": 10.0, "1 year": 20.0}
DEFAULT_EQUITY_THRESHOLDS = {"5 day": 5.0, "2 week": 7.5, "3 month": 15.0, "1 year": 25.0}


# --- Data loading -----------------------------------------------------------
def _source_bytes(uploaded_file) -> tuple[bytes, str]:
    if uploaded_file is not None:
        return uploaded_file.getvalue(), f"uploaded file · {uploaded_file.name}"
    if DEFAULT_FILE.exists():
        stamp = DEFAULT_FILE.stat().st_mtime
        return DEFAULT_FILE.read_bytes(), f"default workbook · saved {pd.to_datetime(stamp, unit='s'):%Y-%m-%d %H:%M}"
    raise FileNotFoundError("Upload the Bank Monitoring Model Excel workbook.")


def _clean_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return " ".join(str(value).replace("\xa0", " ").split()).strip()


def _normalize_region(bank: str, raw_region: object) -> str:
    if bank in BANK_REGION_OVERRIDES:
        return BANK_REGION_OVERRIDES[bank]
    key = _clean_text(raw_region).lower()
    return REGION_ALIASES.get(key, _clean_text(raw_region) or "Other")


@st.cache_data(show_spinner=False)
def load_bank_monitor(file_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(
        io.BytesIO(file_bytes), sheet_name="Report", header=None,
        usecols="A:K", engine="openpyxl"
    )
    rows: list[dict] = []
    for _, row in raw.iterrows():
        bank = _clean_text(row.iloc[0])
        if not bank or bank in {"Bank", "Distribution"}:
            continue

        numeric = pd.to_numeric(row.iloc[2:11], errors="coerce")
        if numeric.isna().all():
            continue

        record = {COLUMN_MAP[i]: row.iloc[i] for i in range(11)}
        record["Bank"] = bank
        record["Region"] = _normalize_region(bank, record.pop("Region Raw"))
        rows.append(record)

    df = pd.DataFrame(rows)
    numeric_cols = [c for c in df.columns if c not in {"Bank", "Region"}]
    df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors="coerce")

    equity_cols = [f"Equity {p}" for p in PERIODS]
    df[equity_cols] = df[equity_cols] * 100
    return df


def fmt(value: float, unit: str = "", decimals: int = 1) -> str:
    if pd.isna(value):
        return "N/A"
    sign = "+" if value > 0 else ""
    suffix = f" {unit}" if unit else ""
    return f"{sign}{value:,.{decimals}f}{suffix}"


# --- Signal logic -----------------------------------------------------------
def signal_label(cds_change: float, equity_return: float,
                 cds_threshold: float, equity_threshold: float) -> str:
    if pd.isna(cds_change) or pd.isna(equity_return):
        return "Insufficient data"
    if cds_change >= cds_threshold and equity_return <= -equity_threshold:
        return "Deteriorating"
    if cds_change <= -cds_threshold and equity_return >= equity_threshold:
        return "Improving"
    if cds_change >= cds_threshold or equity_return <= -equity_threshold:
        return "Watch"
    if cds_change <= -cds_threshold or equity_return >= equity_threshold:
        return "Positive"
    return "Mixed / stable"


def build_watchlist(df: pd.DataFrame, period: str,
                    cds_threshold: float, equity_threshold: float,
                    include_positive: bool = False) -> pd.DataFrame:
    cds_col, eq_col = f"CDS {period}", f"Equity {period}"
    out = df[["Bank", "Region", cds_col, eq_col]].copy()
    out["Signal"] = [
        signal_label(cds, eq, cds_threshold, equity_threshold)
        for cds, eq in zip(out[cds_col], out[eq_col])
    ]
    if include_positive:
        keep = {"Deteriorating", "Improving", "Watch"}
    else:
        keep = {"Deteriorating", "Watch"}
    out = out[out["Signal"].isin(keep)].copy()
    rank = {"Deteriorating": 0, "Watch": 1, "Improving": 2, "Positive": 3}
    out["_rank"] = out["Signal"].map(rank).fillna(9)
    out["_severity"] = (
        pd.to_numeric(out[cds_col], errors="coerce").fillna(0).clip(lower=0)
        + (-pd.to_numeric(out[eq_col], errors="coerce").fillna(0)).clip(lower=0)
    )
    return out.sort_values(["_rank", "_severity"], ascending=[True, False]).drop(columns=["_rank", "_severity"])


def simultaneous_deterioration(df: pd.DataFrame, period: str) -> pd.DataFrame:
    cds_col, eq_col = f"CDS {period}", f"Equity {period}"
    return df[(df[cds_col] > 0) & (df[eq_col] < 0)].copy()


def regional_summary(df: pd.DataFrame, period: str,
                     cds_threshold: float, equity_threshold: float) -> pd.DataFrame:
    cds_col, eq_col = f"CDS {period}", f"Equity {period}"
    rows = []
    for region, grp in df.groupby("Region", sort=False):
        usable = grp.dropna(subset=[cds_col, eq_col])
        if usable.empty:
            continue
        deteriorating = sum(
            signal_label(c, e, cds_threshold, equity_threshold) == "Deteriorating"
            for c, e in zip(usable[cds_col], usable[eq_col])
        )
        rows.append({
            "Region": region,
            f"Avg CDS {period}": usable[cds_col].mean(),
            f"Avg Equity {period}": usable[eq_col].mean(),
            "Deteriorating": deteriorating,
            "Banks": len(usable),
            "Deteriorating %": 100 * deteriorating / len(usable),
        })
    return pd.DataFrame(rows).sort_values(
        ["Deteriorating %", f"Avg CDS {period}"],
        ascending=[False, False]
    )


def _direction_score(cds_change: float, equity_return: float) -> int | None:
    if pd.isna(cds_change) or pd.isna(equity_return):
        return None
    cds_score = 1 if cds_change > 0 else (-1 if cds_change < 0 else 0)
    eq_score = 1 if equity_return < 0 else (-1 if equity_return > 0 else 0)
    return cds_score + eq_score


def _trend_label(current_score: int | None, compare_score: int | None) -> str:
    if current_score is None or compare_score is None:
        return "Insufficient data"
    if current_score >= 1 and compare_score >= 1:
        return "Deterioration accelerating" if current_score > compare_score else "Persistent weakness"
    if current_score <= -1 and compare_score <= -1:
        return "Improvement accelerating" if current_score < compare_score else "Persistent improvement"
    if current_score > compare_score:
        return "Recent deterioration"
    if current_score < compare_score:
        return "Recent improvement"
    return "Mixed / stable"


def comparison_table(df: pd.DataFrame, primary_period: str, compare_period: str) -> pd.DataFrame:
    p_cds, p_eq = f"CDS {primary_period}", f"Equity {primary_period}"
    c_cds, c_eq = f"CDS {compare_period}", f"Equity {compare_period}"
    rows = []
    for _, row in df.iterrows():
        current = _direction_score(row[p_cds], row[p_eq])
        compare = _direction_score(row[c_cds], row[c_eq])
        rows.append({
            "Bank": row["Bank"],
            "Region": row["Region"],
            f"CDS {primary_period}": row[p_cds],
            f"Equity {primary_period}": row[p_eq],
            f"CDS {compare_period}": row[c_cds],
            f"Equity {compare_period}": row[c_eq],
            "Trend": _trend_label(current, compare),
            "_current_score": current if current is not None else -99,
            "_delta_score": (current - compare) if current is not None and compare is not None else -99,
        })
    out = pd.DataFrame(rows)
    return out.sort_values(["_current_score", "_delta_score"], ascending=False).drop(columns=["_current_score", "_delta_score"])


# --- Charts -----------------------------------------------------------------
def chart_ranked(df: pd.DataFrame, metric: str, period: str, top_n: int) -> go.Figure:
    col = METRIC_INFO[metric]["columns"][period]
    plot_df = df[["Bank", "Region", col]].dropna().copy()
    plot_df = plot_df.reindex(plot_df[col].abs().sort_values(ascending=False).index).head(top_n)
    plot_df = plot_df.sort_values(col)

    def bar_color(v: float) -> str:
        if metric == "CDS":
            return BAD if v > 0 else (GOOD if v < 0 else NEUTRAL)
        return GOOD if v > 0 else (BAD if v < 0 else NEUTRAL)

    fig = go.Figure(go.Bar(
        x=plot_df[col],
        y=plot_df["Bank"],
        orientation="h",
        marker_color=[bar_color(v) for v in plot_df[col]],
        customdata=np.column_stack([plot_df["Region"]]),
        hovertemplate=(
            "<b>%{y}</b><br>"
            f"{metric} change: %{{x:.2f}} {METRIC_INFO[metric]['unit']}<br>"
            "Region: %{customdata[0]}<extra></extra>"
        ),
    ))
    fig.add_vline(x=0, line_width=1, line_color=MUTED)
    fig.update_layout(
        title=f"Largest {metric} moves — {period}",
        xaxis_title=f"Change ({METRIC_INFO[metric]['unit']})",
        yaxis_title="",
        height=max(430, 31 * len(plot_df)),
        margin=dict(l=10, r=10, t=60, b=10),
        showlegend=False,
    )
    return fig


def chart_risk_matrix(df: pd.DataFrame, period: str, label_top: int = 8) -> go.Figure:
    x_col = f"CDS {period}"
    y_col = f"Equity {period}"
    plot_df = df.dropna(subset=[x_col, y_col]).copy()
    fig = go.Figure()

    if plot_df.empty:
        fig.update_layout(title=f"CDS vs equity risk matrix — {period}")
        return fig

    x_limit = max(float(plot_df[x_col].abs().max()) * 1.18, 1.0)
    y_limit = max(float(plot_df[y_col].abs().max()) * 1.18, 1.0)

    quadrant_shapes = [
        dict(type="rect", x0=-x_limit, x1=0, y0=0, y1=y_limit,
             fillcolor="rgba(84,130,53,0.10)", line_width=0, layer="below"),
        dict(type="rect", x0=0, x1=x_limit, y0=-y_limit, y1=0,
             fillcolor="rgba(192,0,0,0.10)", line_width=0, layer="below"),
        dict(type="rect", x0=0, x1=x_limit, y0=0, y1=y_limit,
             fillcolor="rgba(197,90,17,0.06)", line_width=0, layer="below"),
        dict(type="rect", x0=-x_limit, x1=0, y0=-y_limit, y1=0,
             fillcolor="rgba(197,90,17,0.06)", line_width=0, layer="below"),
    ]

    for region, grp in plot_df.groupby("Region", sort=False):
        sizes = np.full(len(grp), 12.0)
        if "CDS Now" in grp and grp["CDS Now"].notna().any():
            cds_now = grp["CDS Now"].fillna(grp["CDS Now"].median()).clip(lower=0)
            max_cds = max(float(cds_now.max()), 1.0)
            sizes = 9 + 16 * np.sqrt(cds_now / max_cds)
        fig.add_trace(go.Scatter(
            x=grp[x_col], y=grp[y_col], mode="markers", name=region,
            marker=dict(
                size=sizes,
                color=REGION_COLORS.get(region, NEUTRAL),
                line=dict(color="white", width=0.8),
                opacity=0.9,
            ),
            customdata=np.column_stack([grp["Bank"], grp["Region"]]),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Region: %{customdata[1]}<br>"
                "CDS: %{x:+.2f} bps<br>"
                "Equity: %{y:+.2f}%<extra></extra>"
            ),
        ))

    x_scale = plot_df[x_col].abs().max() or 1
    y_scale = plot_df[y_col].abs().max() or 1
    magnitude = ((plot_df[x_col] / x_scale) ** 2 + (plot_df[y_col] / y_scale) ** 2) ** 0.5
    for _, row in plot_df.loc[magnitude.nlargest(min(label_top, len(plot_df))).index].iterrows():
        fig.add_annotation(
            x=row[x_col], y=row[y_col], text=row["Bank"],
            showarrow=False, yshift=12, font=dict(size=9, color=INK),
        )

    annotations = [
        (-0.62 * x_limit, 0.82 * y_limit, "Improving"),
        (0.62 * x_limit, -0.82 * y_limit, "Deteriorating"),
        (0.62 * x_limit, 0.82 * y_limit, "Mixed"),
        (-0.62 * x_limit, -0.82 * y_limit, "Mixed"),
    ]
    for x, y, text in annotations:
        fig.add_annotation(
            x=x, y=y, text=f"<b>{text}</b>", showarrow=False,
            font=dict(size=11, color=MUTED),
        )

    fig.update_layout(
        title=f"CDS vs equity risk matrix — {period}",
        xaxis=dict(title="CDS change (bps)", range=[-x_limit, x_limit], zeroline=False, gridcolor=GRID),
        yaxis=dict(title="Equity return (%)", range=[-y_limit, y_limit], zeroline=False, gridcolor=GRID),
        shapes=quadrant_shapes,
        height=560,
        legend_title_text="Region",
        margin=dict(l=10, r=10, t=60, b=10),
    )
    fig.add_vline(x=0, line_width=1, line_dash="dot", line_color=MUTED)
    fig.add_hline(y=0, line_width=1, line_dash="dot", line_color=MUTED)
    return fig


def chart_heatmap(df: pd.DataFrame, metric: str, sort_by: str = "Largest absolute move",
                  max_rows: int | None = None) -> go.Figure:
    columns = [METRIC_INFO[metric]["columns"][p] for p in PERIODS]
    work = df[["Bank", "Region", *columns]].copy()

    if sort_by == "Alphabetical":
        work = work.sort_values("Bank")
    elif sort_by == "Region":
        work = work.sort_values(["Region", "Bank"])
    elif sort_by in PERIODS:
        col = METRIC_INFO[metric]["columns"][sort_by]
        work = work.sort_values(col, ascending=(metric == "Equity"), na_position="last")
    else:
        work["_abs"] = work[columns].abs().max(axis=1)
        work = work.sort_values("_abs", ascending=False).drop(columns="_abs")

    if max_rows is not None:
        work = work.head(max_rows)

    labels = work["Bank"].astype(str)
    if sort_by == "Region":
        labels = work["Region"].astype(str) + " · " + labels

    heat = work[columns].copy()
    heat.columns = PERIODS
    heat.index = labels

    fig = go.Figure(go.Heatmap(
        z=heat.to_numpy(),
        x=heat.columns.tolist(),
        y=heat.index.tolist(),
        colorscale="RdYlGn" if metric == "Equity" else "RdYlGn_r",
        zmid=0,
        colorbar=dict(title=""),
        text=np.round(heat.to_numpy(), 1) if len(heat) <= 28 else None,
        texttemplate="%{text:.1f}" if len(heat) <= 28 else None,
        hovertemplate=(
            "<b>%{y}</b><br>Period: %{x}<br>"
            f"{metric} change: %{{z:+.2f}} {METRIC_INFO[metric]['unit']}<extra></extra>"
        ),
    ))
    fig.update_layout(
        title=f"{metric} changes across periods ({METRIC_INFO[metric]['unit']})",
        height=max(520, 25 * len(heat)),
        margin=dict(l=10, r=10, t=60, b=10),
        xaxis_title="",
        yaxis_title="",
    )
    return fig


def chart_region_summary(summary: pd.DataFrame, period: str) -> go.Figure:
    cds_col, eq_col = f"Avg CDS {period}", f"Avg Equity {period}"
    if summary.empty:
        return go.Figure()
    plot_df = summary.sort_values(cds_col)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=plot_df[cds_col], y=plot_df["Region"], orientation="h",
        name="Avg CDS change (bps)", marker_color=[
            BAD if v > 0 else GOOD for v in plot_df[cds_col]
        ],
        hovertemplate="%{y}<br>Avg CDS: %{x:+.2f} bps<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=plot_df[eq_col], y=plot_df["Region"], mode="markers",
        name="Avg equity return (%)", xaxis="x2",
        marker=dict(size=12, symbol="diamond", color=[
            GOOD if v > 0 else BAD for v in plot_df[eq_col]
        ]),
        hovertemplate="%{y}<br>Avg Equity: %{x:+.2f}%<extra></extra>",
    ))
    fig.update_layout(
        title=f"Regional market summary — {period}",
        height=max(380, 48 * len(plot_df)),
        xaxis=dict(title="Average CDS change (bps)", side="bottom", zeroline=True, zerolinecolor=RULE),
        xaxis2=dict(title="Average equity return (%)", overlaying="x", side="top", zeroline=True, zerolinecolor=RULE),
        yaxis_title="",
        legend=dict(orientation="h", y=1.16, x=0),
        margin=dict(l=10, r=10, t=80, b=20),
    )
    return fig


def chart_bank_detail(df: pd.DataFrame, bank: str) -> go.Figure:
    row = df[df["Bank"] == bank]
    if row.empty:
        return go.Figure()
    row = row.iloc[0]
    cds_vals = [row[f"CDS {p}"] for p in PERIODS]
    eq_vals = [row[f"Equity {p}"] for p in PERIODS]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=PERIODS, y=cds_vals, name="CDS change (bps)",
        marker_color=[BAD if (not pd.isna(v) and v > 0) else GOOD for v in cds_vals],
        yaxis="y",
        hovertemplate="%{x}<br>CDS: %{y:+.2f} bps<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=PERIODS, y=eq_vals, name="Equity return (%)",
        mode="lines+markers", yaxis="y2",
        line=dict(color=NAVY, width=2.5),
        marker=dict(size=9),
        hovertemplate="%{x}<br>Equity: %{y:+.2f}%<extra></extra>",
    ))
    fig.update_layout(
        title=f"{bank} — multi-period market profile",
        height=500,
        yaxis=dict(title="CDS change (bps)", zeroline=True, zerolinecolor=RULE),
        yaxis2=dict(title="Equity return (%)", overlaying="y", side="right", zeroline=True, zerolinecolor=RULE),
        legend=dict(orientation="h", y=1.12, x=0),
        margin=dict(l=10, r=10, t=80, b=10),
    )
    return fig


# --- Executive summary helpers ----------------------------------------------
def period_extremes(df: pd.DataFrame, period: str) -> dict:
    cds_col, eq_col = f"CDS {period}", f"Equity {period}"
    out: dict = {}
    cds = df.dropna(subset=[cds_col])
    equity = df.dropna(subset=[eq_col])
    if not cds.empty:
        out["cds_wide"] = cds.loc[cds[cds_col].idxmax()]
        out["cds_tight"] = cds.loc[cds[cds_col].idxmin()]
    if not equity.empty:
        out["eq_strong"] = equity.loc[equity[eq_col].idxmax()]
        out["eq_weak"] = equity.loc[equity[eq_col].idxmin()]
    return out


def generate_summary(df: pd.DataFrame, period: str, compare_period: str,
                     cds_threshold: float, equity_threshold: float) -> list[str]:
    statements: list[str] = []
    cds_col, eq_col = f"CDS {period}", f"Equity {period}"
    usable = df.dropna(subset=[cds_col, eq_col])
    if usable.empty:
        return ["No complete CDS/equity observations are available for the selected filters."]

    wider = int((usable[cds_col] > 0).sum())
    equity_down = int((usable[eq_col] < 0).sum())
    both = int(((usable[cds_col] > 0) & (usable[eq_col] < 0)).sum())
    statements.append(
        f"{wider} of {len(usable)} banks recorded CDS widening, {equity_down} posted equity declines, "
        f"and {both} showed both signals over the {period} window."
    )

    extremes = period_extremes(usable, period)
    if "cds_wide" in extremes and "eq_weak" in extremes:
        statements.append(
            f"{extremes['cds_wide']['Bank']} had the largest CDS widening at "
            f"{fmt(extremes['cds_wide'][cds_col], 'bps')}; "
            f"{extremes['eq_weak']['Bank']} had the weakest equity return at "
            f"{fmt(extremes['eq_weak'][eq_col], '%')}."
        )

    reg = regional_summary(usable, period, cds_threshold, equity_threshold)
    if not reg.empty:
        top = reg.iloc[0]
        statements.append(
            f"{top['Region']} has the highest share of threshold-based deterioration flags: "
            f"{int(top['Deteriorating'])} of {int(top['Banks'])} banks "
            f"({top['Deteriorating %']:.0f}%)."
        )

    comp = comparison_table(usable, period, compare_period)
    counts = comp["Trend"].value_counts()
    accelerating = int(counts.get("Deterioration accelerating", 0) + counts.get("Recent deterioration", 0))
    improving = int(counts.get("Improvement accelerating", 0) + counts.get("Recent improvement", 0))
    statements.append(
        f"Relative to the {compare_period} window, {accelerating} banks show a more adverse directional mix "
        f"and {improving} show a more favorable directional mix."
    )
    return statements


# --- PDF helpers -------------------------------------------------------------
PAGE_SIZE = landscape(letter)
MARGIN_X, MARGIN_TOP, MARGIN_BOTTOM = 32, 30, 38
PDF_MAX_ROWS = 24
TITLE_H, TITLE_GAP = 15, 8


def _prepare_for_print(fig: go.Figure, width: int, height: int) -> go.Figure:
    output = go.Figure(fig)
    is_heatmap = bool(output.data) and isinstance(output.data[0], go.Heatmap)
    output.update_layout(
        title=None, template="plotly_white", width=width, height=height,
        font=dict(family="Helvetica, Arial, sans-serif", size=10.5, color=INK),
        paper_bgcolor="white", plot_bgcolor="white",
        legend=dict(
            orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1.0,
            title_text="", font=dict(size=9), bgcolor="rgba(0,0,0,0)",
        ),
        margin=dict(l=8, r=70 if is_heatmap else 8, t=28, b=34),
    )
    output.update_xaxes(
        automargin=True, gridcolor=GRID, linecolor=RULE,
        ticks="outside", tickcolor=RULE, ticklen=4,
    )
    output.update_yaxes(
        automargin=True, gridcolor=GRID, linecolor=RULE,
    )
    return output


def _fig_to_png(fig: go.Figure, width: int, height: int, scale: int = 2) -> bytes:
    return _prepare_for_print(fig, width, height).to_image(
        format="png", width=width, height=height, scale=scale
    )


def _footer(canvas, doc):
    canvas.saveState()
    page_w, _ = PAGE_SIZE
    canvas.setStrokeColor(colors.HexColor(RULE))
    canvas.setLineWidth(0.5)
    canvas.line(MARGIN_X, 27, page_w - MARGIN_X, 27)
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(colors.HexColor(MUTED))
    canvas.drawString(MARGIN_X, 16, doc.footer_note)
    canvas.drawRightString(page_w - MARGIN_X, 16, str(canvas.getPageNumber()))
    canvas.restoreState()


def _pdf_styles():
    styles = getSampleStyleSheet()
    custom = {
        "TitleX": dict(fontName="Helvetica-Bold", fontSize=19, leading=22, textColor=colors.HexColor(NAVY)),
        "SubX": dict(fontName="Helvetica", fontSize=9.5, leading=12, textColor=colors.HexColor(MUTED)),
        "SectionX": dict(fontName="Helvetica-Bold", fontSize=12.5, leading=15, textColor=colors.HexColor(NAVY)),
        "BodyX": dict(fontName="Helvetica", fontSize=9.5, leading=13, textColor=colors.HexColor(INK)),
        "BulletX": dict(fontName="Helvetica", fontSize=9.4, leading=13, leftIndent=12, bulletIndent=2, textColor=colors.HexColor(INK)),
        "CardLabelX": dict(fontName="Helvetica-Bold", fontSize=7, leading=9, textColor=colors.HexColor(MUTED)),
        "CardValueX": dict(fontName="Helvetica-Bold", fontSize=12.5, leading=15, textColor=colors.HexColor(NAVY)),
    }
    for name, kwargs in custom.items():
        if name not in styles:
            styles.add(ParagraphStyle(name=name, parent=styles["BodyText"], **kwargs))
    return styles


def _metric_card(label: str, value: str, styles, width: float) -> Table:
    card = Table([
        [Paragraph(label.upper(), styles["CardLabelX"])],
        [Paragraph(value, styles["CardValueX"])],
    ], colWidths=[width], rowHeights=[18, 34])
    card.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(PANEL)),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor(RULE)),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return card


def _cards_grid(items: list[tuple[str, str]], styles, usable_width: float) -> Table:
    col_w = usable_width / 3
    rows = []
    for i in range(0, len(items), 3):
        cells = [_metric_card(label, value, styles, col_w - 8) for label, value in items[i:i+3]]
        while len(cells) < 3:
            cells.append("")
        rows.append(cells)
    grid = Table(rows, colWidths=[col_w] * 3)
    grid.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return grid


def _df_table(df: pd.DataFrame, widths: list[float] | None = None,
              font_size: float = 7.4, max_rows: int = 22) -> Table:
    if df.empty:
        data = [["No observations match the current filters."]]
        table = Table(data, colWidths=[700])
        table.setStyle(TableStyle([
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor(MUTED)),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
        ]))
        return table

    shown = df.head(max_rows).copy()
    data = [list(shown.columns)] + shown.astype(str).values.tolist()
    table = Table(data, colWidths=widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(NAVY)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor(RULE)),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FAFBFC")]),
        ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor(INK)),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return table


def _format_watchlist_for_pdf(watch: pd.DataFrame, period: str) -> pd.DataFrame:
    cds_col, eq_col = f"CDS {period}", f"Equity {period}"
    out = watch[["Bank", "Region", cds_col, eq_col, "Signal"]].copy()
    out[cds_col] = out[cds_col].map(lambda x: fmt(x, "bps"))
    out[eq_col] = out[eq_col].map(lambda x: fmt(x, "%"))
    return out.rename(columns={cds_col: "CDS", eq_col: "Equity"})


def _format_regional_for_pdf(reg: pd.DataFrame, period: str) -> pd.DataFrame:
    if reg.empty:
        return reg
    cds_col, eq_col = f"Avg CDS {period}", f"Avg Equity {period}"
    out = reg[["Region", cds_col, eq_col, "Deteriorating", "Banks", "Deteriorating %"]].copy()
    out[cds_col] = out[cds_col].map(lambda x: fmt(x, "bps"))
    out[eq_col] = out[eq_col].map(lambda x: fmt(x, "%"))
    out["Deteriorating"] = out["Deteriorating"].astype(int).astype(str) + "/" + out["Banks"].astype(int).astype(str)
    out["Deteriorating %"] = out["Deteriorating %"].map(lambda x: f"{x:.0f}%")
    return out.drop(columns="Banks").rename(columns={
        cds_col: "Avg CDS", eq_col: "Avg Equity", "Deteriorating": "Flags"
    })


def _format_comparison_for_pdf(comp: pd.DataFrame, primary_period: str, compare_period: str) -> pd.DataFrame:
    p_cds, p_eq = f"CDS {primary_period}", f"Equity {primary_period}"
    c_cds, c_eq = f"CDS {compare_period}", f"Equity {compare_period}"
    out = comp[["Bank", p_cds, p_eq, c_cds, c_eq, "Trend"]].copy()
    for col in [p_cds, c_cds]:
        out[col] = out[col].map(lambda x: fmt(x, "bps"))
    for col in [p_eq, c_eq]:
        out[col] = out[col].map(lambda x: fmt(x, "%"))
    return out.rename(columns={
        p_cds: f"CDS {primary_period}",
        p_eq: f"Eq {primary_period}",
        c_cds: f"CDS {compare_period}",
        c_eq: f"Eq {compare_period}",
    })


def make_pdf(
    df: pd.DataFrame,
    period: str,
    compare_period: str,
    selected_regions: Iterable[str],
    top_n: int,
    cds_threshold: float,
    equity_threshold: float,
    heatmap_sort: str,
    selected_bank: str,
    analyst_notes: list[str],
) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=PAGE_SIZE,
        leftMargin=MARGIN_X, rightMargin=MARGIN_X,
        topMargin=MARGIN_TOP, bottomMargin=MARGIN_BOTTOM,
        title="Bank Market Monitor — Chart Pack",
    )
    doc.footer_note = (
        f"Bank Market Monitor · {period} vs {compare_period} · "
        f"{', '.join(selected_regions)}"
    )
    styles = _pdf_styles()
    story: list = []

    cds_col, eq_col = f"CDS {period}", f"Equity {period}"
    extremes = period_extremes(df, period)
    both = simultaneous_deterioration(df, period)
    watch = build_watchlist(df, period, cds_threshold, equity_threshold, include_positive=True)
    reg = regional_summary(df, period, cds_threshold, equity_threshold)
    comp = comparison_table(df, period, compare_period)

    story.extend([
        Paragraph("Bank Market Monitor", styles["TitleX"]),
        Paragraph(
            f"Executive summary · primary window: {period} · comparison window: {compare_period}",
            styles["SubX"],
        ),
        Spacer(1, 12),
    ])

    cards = [
        ("Banks monitored", str(len(df))),
        ("Largest CDS widening", (
            f"{extremes['cds_wide']['Bank']} · {fmt(extremes['cds_wide'][cds_col], 'bps')}"
            if "cds_wide" in extremes else "N/A"
        )),
        ("Largest CDS tightening", (
            f"{extremes['cds_tight']['Bank']} · {fmt(extremes['cds_tight'][cds_col], 'bps')}"
            if "cds_tight" in extremes else "N/A"
        )),
        ("Largest equity decline", (
            f"{extremes['eq_weak']['Bank']} · {fmt(extremes['eq_weak'][eq_col], '%')}"
            if "eq_weak" in extremes else "N/A"
        )),
        ("Largest equity gain", (
            f"{extremes['eq_strong']['Bank']} · {fmt(extremes['eq_strong'][eq_col], '%')}"
            if "eq_strong" in extremes else "N/A"
        )),
        ("CDS up + equity down", f"{len(both)} banks"),
    ]
    story.append(_cards_grid(cards, styles, doc.width))
    story.extend([Spacer(1, 7), Paragraph("Key observations", styles["SectionX"]), Spacer(1, 4)])
    for text in generate_summary(df, period, compare_period, cds_threshold, equity_threshold):
        story.append(Paragraph(text, styles["BulletX"], bulletText="•"))
    if analyst_notes:
        story.extend([Spacer(1, 6), Paragraph("Analyst notes", styles["SectionX"]), Spacer(1, 3)])
        for text in analyst_notes:
            story.append(Paragraph(text, styles["BulletX"], bulletText="•"))
    story.append(PageBreak())

    story.extend([
        Paragraph("Risk watchlist", styles["TitleX"]),
        Paragraph(
            f"Flags use CDS widening ≥ {cds_threshold:.1f} bps and equity decline ≥ {equity_threshold:.1f}% "
            f"over the {period} window.",
            styles["SubX"],
        ),
        Spacer(1, 10),
        _df_table(_format_watchlist_for_pdf(watch, period),
                  widths=[190, 110, 90, 90, 130], font_size=8.1, max_rows=24),
        PageBreak(),
    ])

    story.extend([
        Paragraph("Regional and period comparison", styles["TitleX"]),
        Spacer(1, 8),
        Paragraph("Regional summary", styles["SectionX"]),
        Spacer(1, 5),
        _df_table(_format_regional_for_pdf(reg, period),
                  widths=[145, 110, 110, 80, 90], font_size=8.0, max_rows=12),
        Spacer(1, 14),
        Paragraph(f"Directional comparison: {period} vs {compare_period}", styles["SectionX"]),
        Spacer(1, 5),
        _df_table(_format_comparison_for_pdf(comp, period, compare_period),
                  widths=[145, 90, 90, 90, 90, 150], font_size=6.9, max_rows=14),
        PageBreak(),
    ])

    panel_w = int(doc.width * 0.88)
    panel_h = int((doc.height - 48) * 0.88)
    figures = [
        chart_risk_matrix(df, period, label_top=8),
        chart_region_summary(reg, period),
        chart_ranked(df, "CDS", period, min(top_n, PDF_MAX_ROWS)),
        chart_ranked(df, "Equity", period, min(top_n, PDF_MAX_ROWS)),
        chart_heatmap(df, "CDS", sort_by=heatmap_sort, max_rows=PDF_MAX_ROWS),
        chart_heatmap(df, "Equity", sort_by=heatmap_sort, max_rows=PDF_MAX_ROWS),
        chart_bank_detail(df, selected_bank),
    ]

    for idx, fig in enumerate(figures):
        heading = (fig.layout.title.text or "").strip()
        png = _fig_to_png(fig, panel_w, panel_h)
        image = RLImage(io.BytesIO(png), width=panel_w, height=panel_h)
        image.hAlign = "CENTER"
        story.append(KeepTogether([
            Paragraph(heading, styles["SectionX"]),
            Spacer(1, TITLE_GAP),
            image,
        ]))
        if idx < len(figures) - 1:
            story.append(PageBreak())

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return buffer.getvalue()


# --- App --------------------------------------------------------------------
st.title("🏦 Bank Market Monitor")
st.caption("Bank credit and equity monitoring with automated risk signals, regional context, and report-ready outputs.")

with st.sidebar:
    st.header("Data & controls")
    uploaded = st.file_uploader("Upload updated Bank Monitoring Model", type=["xlsx"])

    if st.button(
        "🔄 Reload data", use_container_width=True,
        help="Re-read the workbook and rebuild all views from the latest saved values."
    ):
        st.cache_data.clear()
        st.rerun()

    try:
        source, source_label = _source_bytes(uploaded)
        data = load_bank_monitor(source)
    except Exception as exc:
        st.error(f"Could not read the workbook: {exc}")
        st.stop()

    st.caption(f"Source: {source_label}")

    st.markdown("### Filters")
    available_regions = list(dict.fromkeys(data["Region"].dropna().tolist()))
    selected_regions = st.multiselect("Regions", available_regions, default=available_regions)
    primary_period = st.selectbox("Primary period", PERIODS, index=1)
    compare_options = [p for p in PERIODS if p != primary_period]
    compare_period = st.selectbox(
        "Compare against", compare_options,
        index=min(1, len(compare_options) - 1),
    )
    top_n = st.slider("Banks in ranked charts", 5, max(5, min(30, len(data))), min(15, len(data)))
    search = st.text_input("Search bank")

    st.markdown("### Risk signals")
    cds_threshold = st.number_input(
        "CDS widening threshold (bps)", min_value=0.0, step=0.5,
        value=float(DEFAULT_CDS_THRESHOLDS[primary_period]),
        help="Used for watchlist and regional deterioration flags."
    )
    equity_threshold = st.number_input(
        "Equity decline threshold (%)", min_value=0.0, step=0.5,
        value=float(DEFAULT_EQUITY_THRESHOLDS[primary_period]),
        help="Used for watchlist and regional deterioration flags."
    )

    st.markdown("### Heatmap")
    heatmap_sort = st.selectbox("Sort rows by", HEATMAP_SORT_OPTIONS)

    with st.expander("Report commentary"):
        analyst_notes_raw = st.text_area(
            "Analyst notes (one per line)", height=130,
            placeholder="Optional observations to include on the PDF executive-summary page."
        )

filtered = data[data["Region"].isin(selected_regions)].copy()
if search:
    filtered = filtered[filtered["Bank"].str.contains(search, case=False, na=False)]
if filtered.empty:
    st.warning("No banks match the selected filters.")
    st.stop()

analyst_notes = [line.strip() for line in analyst_notes_raw.splitlines() if line.strip()]
cds_col, eq_col = f"CDS {primary_period}", f"Equity {primary_period}"
extremes = period_extremes(filtered, primary_period)
both = simultaneous_deterioration(filtered, primary_period)
watch = build_watchlist(filtered, primary_period, cds_threshold, equity_threshold, include_positive=True)
reg_summary = regional_summary(filtered, primary_period, cds_threshold, equity_threshold)
comp_table = comparison_table(filtered, primary_period, compare_period)

k1, k2, k3 = st.columns(3)
k1.metric("Banks monitored", len(filtered))
k2.metric(
    "Largest CDS widening",
    extremes["cds_wide"]["Bank"] if "cds_wide" in extremes else "N/A",
    fmt(extremes["cds_wide"][cds_col], "bps") if "cds_wide" in extremes else None,
    delta_color="inverse",
)
k3.metric(
    "Largest CDS tightening",
    extremes["cds_tight"]["Bank"] if "cds_tight" in extremes else "N/A",
    fmt(extremes["cds_tight"][cds_col], "bps") if "cds_tight" in extremes else None,
)

k4, k5, k6 = st.columns(3)
k4.metric(
    "Largest equity decline",
    extremes["eq_weak"]["Bank"] if "eq_weak" in extremes else "N/A",
    fmt(extremes["eq_weak"][eq_col], "%") if "eq_weak" in extremes else None,
    delta_color="inverse",
)
k5.metric(
    "Largest equity gain",
    extremes["eq_strong"]["Bank"] if "eq_strong" in extremes else "N/A",
    fmt(extremes["eq_strong"][eq_col], "%") if "eq_strong" in extremes else None,
)
k6.metric("CDS up + equity down", len(both), help="Banks with both CDS widening and a negative equity return.")

overview_tab, moves_tab, heatmap_tab, bank_tab, data_tab, report_tab = st.tabs([
    "Overview", "Market Moves", "Heatmaps", "Bank Detail", "Data Explorer", "Report"
])

with overview_tab:
    st.subheader(f"Executive summary — {primary_period}")
    for statement in generate_summary(
        filtered, primary_period, compare_period, cds_threshold, equity_threshold
    ):
        st.markdown(f"- {statement}")
    if analyst_notes:
        st.caption("Analyst notes")
        for note in analyst_notes:
            st.markdown(f"- {note}")

    st.markdown("### Risk watchlist")
    st.caption(
        f"Deteriorating = CDS widening ≥ {cds_threshold:.1f} bps and equity decline ≥ {equity_threshold:.1f}%. "
        "Watch = either threshold is breached. Improving = both favorable thresholds are met."
    )
    watch_display = watch.rename(columns={
        cds_col: f"CDS {primary_period} (bps)",
        eq_col: f"Equity {primary_period} (%)",
    })
    if watch_display.empty:
        st.success("No banks breach the selected deterioration/watch thresholds.")
    else:
        st.dataframe(
            watch_display.style.format({
                f"CDS {primary_period} (bps)": "{:+.1f}",
                f"Equity {primary_period} (%)": "{:+.1f}",
            }, na_rep="—"),
            use_container_width=True, hide_index=True,
        )

    st.markdown("### CDS vs equity risk matrix")
    st.plotly_chart(chart_risk_matrix(filtered, primary_period), use_container_width=True)

    st.markdown("### Regional summary")
    if not reg_summary.empty:
        reg_display = reg_summary.copy()
        st.dataframe(
            reg_display.style.format({
                f"Avg CDS {primary_period}": "{:+.1f}",
                f"Avg Equity {primary_period}": "{:+.1f}%",
                "Deteriorating %": "{:.0f}%",
            }),
            use_container_width=True, hide_index=True,
        )
        st.plotly_chart(chart_region_summary(reg_summary, primary_period), use_container_width=True)

    st.markdown(f"### Period comparison — {primary_period} vs {compare_period}")
    comp_display = comp_table.copy()
    st.dataframe(
        comp_display.style.format({
            f"CDS {primary_period}": "{:+.1f}",
            f"Equity {primary_period}": "{:+.1f}%",
            f"CDS {compare_period}": "{:+.1f}",
            f"Equity {compare_period}": "{:+.1f}%",
        }, na_rep="—"),
        use_container_width=True, hide_index=True,
    )

with moves_tab:
    c1, c2 = st.columns(2)
    with c1:
        st.caption(METRIC_INFO["CDS"]["description"])
        st.plotly_chart(
            chart_ranked(filtered, "CDS", primary_period, top_n),
            use_container_width=True,
        )
    with c2:
        st.caption(METRIC_INFO["Equity"]["description"])
        st.plotly_chart(
            chart_ranked(filtered, "Equity", primary_period, top_n),
            use_container_width=True,
        )
    st.info(
        "Ranked-chart colors indicate direction rather than geography: "
        "red = adverse move, green = favorable move. Region remains available in hover details."
    )

with heatmap_tab:
    st.caption(
        "Heatmaps use risk-consistent colors: CDS widening is red and tightening is green; "
        "equity gains are green and declines are red."
    )
    h1, h2 = st.columns(2)
    with h1:
        st.plotly_chart(
            chart_heatmap(filtered, "CDS", sort_by=heatmap_sort),
            use_container_width=True,
        )
    with h2:
        st.plotly_chart(
            chart_heatmap(filtered, "Equity", sort_by=heatmap_sort),
            use_container_width=True,
        )

with bank_tab:
    bank_choices = sorted(filtered["Bank"].dropna().unique().tolist())
    selected_bank = st.selectbox("Financial institution", bank_choices)
    bank_row = filtered[filtered["Bank"] == selected_bank].iloc[0]

    b1, b2, b3 = st.columns(3)
    b1.metric("Region", bank_row["Region"])
    b2.metric("Current CDS", fmt(bank_row["CDS Now"], "bps"))
    bank_signal = signal_label(
        bank_row[cds_col], bank_row[eq_col], cds_threshold, equity_threshold
    )
    b3.metric("Current signal", bank_signal)

    bank_table = pd.DataFrame({
        "Period": PERIODS,
        "CDS change (bps)": [bank_row[f"CDS {p}"] for p in PERIODS],
        "Equity return (%)": [bank_row[f"Equity {p}"] for p in PERIODS],
    })
    st.dataframe(
        bank_table.style.format({
            "CDS change (bps)": "{:+.1f}",
            "Equity return (%)": "{:+.1f}%",
        }, na_rep="—"),
        use_container_width=True, hide_index=True,
    )
    st.plotly_chart(chart_bank_detail(filtered, selected_bank), use_container_width=True)

with data_tab:
    display_columns = [
        "Bank", "Region", "CDS Now",
        *[f"CDS {p}" for p in PERIODS],
        *[f"Equity {p}" for p in PERIODS],
    ]
    table = filtered[display_columns].sort_values(cds_col, ascending=False)
    formatters = {"CDS Now": "{:.1f}"}
    formatters.update({f"CDS {p}": "{:+.1f}" for p in PERIODS})
    formatters.update({f"Equity {p}": "{:+.1f}%" for p in PERIODS})
    st.dataframe(
        table.style.format(formatters, na_rep="—"),
        use_container_width=True, hide_index=True,
    )
    csv = table.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download filtered data (CSV)", csv,
        "bank_market_snapshot.csv", "text/csv",
    )

with report_tab:
    st.subheader("Downloadable monitoring pack")
    st.write(
        "The PDF mirrors the dashboard: executive KPIs and commentary, risk watchlist, "
        "regional summary, period comparison, risk matrix, ranked charts, both heatmaps, "
        "and the selected bank detail."
    )
    st.caption(
        f"Current export settings: {primary_period} vs {compare_period}; "
        f"heatmap sort = {heatmap_sort}; selected bank = {selected_bank}."
    )
    if st.button("Build PDF", type="primary"):
        with st.spinner("Rendering dashboard views into PDF…"):
            try:
                st.session_state["bank_pdf_bytes"] = make_pdf(
                    filtered,
                    primary_period,
                    compare_period,
                    selected_regions,
                    top_n,
                    cds_threshold,
                    equity_threshold,
                    heatmap_sort,
                    selected_bank,
                    analyst_notes,
                )
            except Exception as exc:
                st.session_state.pop("bank_pdf_bytes", None)
                st.error(
                    f"Could not render the PDF: {exc}\n\n"
                    "Static Plotly export requires the pinned Kaleido package in requirements.txt."
                )
    if st.session_state.get("bank_pdf_bytes"):
        st.download_button(
            "Download PDF", st.session_state["bank_pdf_bytes"],
            "bank_market_charts.pdf", "application/pdf",
        )

    with st.expander("Methodology and interpretation"):
        st.markdown(f"""
        **CDS:** Positive changes mean wider spreads and generally higher perceived bank credit risk.  
        **Equity:** Negative returns indicate weaker market performance.  
        **Watchlist:** A bank is *Deteriorating* when both the CDS widening threshold and equity-decline threshold are breached; *Watch* means one threshold is breached.  
        **Risk matrix:** Bottom-right (CDS wider + equity lower) represents simultaneous market deterioration; top-left represents simultaneous improvement.  
        **Period comparison:** Directional labels compare the sign mix of CDS and equity changes between the primary and comparison windows; they are not a credit rating or forecast.  
        **Current thresholds:** {cds_threshold:.1f} bps CDS widening and {equity_threshold:.1f}% equity decline.
        """)

st.caption(
    "Data are read from the cached values saved in the Excel workbook. Save the workbook in Excel "
    "(after refreshing the Capital IQ and LSEG plug-ins), then re-upload it or press “Reload data”."
)
