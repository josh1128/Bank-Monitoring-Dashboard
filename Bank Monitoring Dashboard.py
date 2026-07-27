from __future__ import annotations

import io
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd
import plotly.express as px
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

# --- Shared visual identity -------------------------------------------------
NAVY = "#17365D"
INK = "#33404D"
MUTED = "#7A8794"
RULE = "#D9DEE4"
GRID = "#EDF0F3"
PANEL = "#F2F5F8"
GOOD = "#548235"
BAD = "#C00000"

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

# The workbook contains inconsistent region strings and a few incorrect labels.
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

# Overrides are applied before the workbook's region field because the Report
# sheet currently classifies U.K. banks as Euro Area and Canada as America.
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


def _source_bytes(uploaded_file) -> tuple[bytes, str]:
    """Return the workbook bytes and a short label describing the source."""
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
        # Country/index separator rows have no market observations and are skipped.
        if numeric.isna().all():
            continue

        record = {COLUMN_MAP[i]: row.iloc[i] for i in range(11)}
        record["Bank"] = bank
        record["Region"] = _normalize_region(bank, record.pop("Region Raw"))
        rows.append(record)

    df = pd.DataFrame(rows)
    numeric_cols = [c for c in df.columns if c not in {"Bank", "Region"}]
    df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors="coerce")

    # Equity returns are stored as decimals in Excel; convert to percentage points.
    equity_cols = [f"Equity {p}" for p in PERIODS]
    df[equity_cols] = df[equity_cols] * 100
    return df


def fmt(value: float, unit: str = "", decimals: int = 1) -> str:
    if pd.isna(value):
        return "N/A"
    sign = "+" if value > 0 else ""
    suffix = f" {unit}" if unit else ""
    return f"{sign}{value:,.{decimals}f}{suffix}"


# --- Charts -----------------------------------------------------------------
def chart_ranked(df: pd.DataFrame, metric: str, period: str, top_n: int) -> go.Figure:
    col = METRIC_INFO[metric]["columns"][period]
    plot_df = df[["Bank", "Region", col]].dropna().copy()
    plot_df = plot_df.reindex(plot_df[col].abs().sort_values(ascending=False).index).head(top_n)
    plot_df = plot_df.sort_values(col)
    fig = px.bar(
        plot_df, x=col, y="Bank", color="Region",
        color_discrete_map=REGION_COLORS, orientation="h",
        labels={col: f"Change ({METRIC_INFO[metric]['unit']})", "Bank": ""},
        title=f"Largest {metric} moves — {period}",
        hover_data={"Region": True, col: ":.2f"},
    )
    fig.add_vline(x=0, line_width=1, line_color=MUTED)
    fig.update_layout(
        height=max(430, 31 * len(plot_df)), legend_title_text="Region",
        margin=dict(l=10, r=10, t=60, b=10),
    )
    return fig


def chart_scatter(df: pd.DataFrame, period: str, label_top: int = 0) -> go.Figure:
    x_col = f"CDS {period}"
    y_col = f"Equity {period}"
    plot_df = df.dropna(subset=[x_col, y_col, "CDS Now"]).copy()
    fig = px.scatter(
        plot_df, x=x_col, y=y_col, color="Region",
        color_discrete_map=REGION_COLORS, size="CDS Now", size_max=26,
        hover_name="Bank",
        labels={x_col: "CDS change (bps)", y_col: "Equity return (%)", "CDS Now": "CDS now"},
        title=f"Credit versus equity performance — {period}",
    )
    fig.add_hline(y=0, line_width=1, line_dash="dot", line_color=MUTED)
    fig.add_vline(x=0, line_width=1, line_dash="dot", line_color=MUTED)
    if label_top and not plot_df.empty:
        x_scale = plot_df[x_col].abs().max() or 1
        y_scale = plot_df[y_col].abs().max() or 1
        magnitude = ((plot_df[x_col] / x_scale) ** 2 + (plot_df[y_col] / y_scale) ** 2) ** 0.5
        for _, row in plot_df.loc[magnitude.nlargest(label_top).index].iterrows():
            fig.add_annotation(
                x=row[x_col], y=row[y_col], text=row["Bank"],
                showarrow=False, yshift=13, font=dict(size=9.5, color=INK),
            )
    fig.update_layout(height=520, legend_title_text="Region", margin=dict(l=10, r=10, t=60, b=10))
    return fig


def chart_heatmap(df: pd.DataFrame, metric: str, max_rows: int | None = None) -> go.Figure:
    columns = [METRIC_INFO[metric]["columns"][p] for p in PERIODS]
    heat = df.set_index("Bank")[columns]
    heat.columns = PERIODS
    heat = heat.loc[heat.abs().max(axis=1).sort_values(ascending=False).index]
    if max_rows is not None:
        heat = heat.head(max_rows)
    fig = px.imshow(
        heat, aspect="auto", color_continuous_scale="RdYlGn_r",
        color_continuous_midpoint=0, text_auto=".1f" if len(heat) <= 28 else False,
        title=f"{metric} changes across periods ({METRIC_INFO[metric]['unit']})",
    )
    fig.update_layout(height=max(520, 25 * len(heat)), margin=dict(l=10, r=10, t=60, b=10))
    fig.update_coloraxes(colorbar_title_text="")
    return fig


# --- Executive summary helpers ----------------------------------------------
def period_extremes(df: pd.DataFrame, period: str) -> dict:
    """Return the banks at the extremes of CDS and equity moves for a period."""
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


def cds_takeaways(df: pd.DataFrame, period: str) -> list[str]:
    """Auto-generated CDS commentary used when the CDS box is left blank."""
    statements: list[str] = []
    col = f"CDS {period}"
    cds = df.dropna(subset=[col])
    if not cds.empty:
        low, high = cds[col].min(), cds[col].max()
        wider = cds.loc[cds[col].idxmax()]
        tighter = cds.loc[cds[col].idxmin()]
        statements.append(
            f"Bank CDS moves ranged from {fmt(low, 'bps')} to {fmt(high, 'bps')} "
            f"over the {period} window."
        )
        statements.append(
            f"{wider['Bank']} saw the largest widening at {fmt(wider[col], 'bps')}; "
            f"{tighter['Bank']} tightened the most at {fmt(tighter[col], 'bps')}."
        )
    return statements


def equity_takeaways(df: pd.DataFrame, period: str) -> list[str]:
    """Auto-generated equity commentary used when the Equity box is left blank."""
    statements: list[str] = []
    col = f"Equity {period}"
    equity = df.dropna(subset=[col])
    if not equity.empty:
        strongest = equity.loc[equity[col].idxmax()]
        weakest = equity.loc[equity[col].idxmin()]
        statements.append(
            f"{strongest['Bank']} led equity performance at {fmt(strongest[col], '%')}; "
            f"{weakest['Bank']} was weakest at {fmt(weakest[col], '%')}."
        )
        statements.append(
            f"Equity moves ranged from {fmt(equity[col].min(), '%')} "
            f"to {fmt(equity[col].max(), '%')} over the {period} window."
        )
    return statements


def generate_summary(df: pd.DataFrame, period: str) -> list[str]:
    """Combined CDS + Equity commentary, used for the on-screen overview."""
    return cds_takeaways(df, period) + equity_takeaways(df, period)


# --- PDF chart pack ---------------------------------------------------------
PAGE_SIZE = landscape(letter)
MARGIN_X, MARGIN_TOP, MARGIN_BOTTOM = 32, 30, 40
PDF_MAX_ROWS = 26
TITLE_H, TITLE_GAP = 15, 8


def _prepare_for_print(fig: go.Figure, width: int, height: int) -> go.Figure:
    output = go.Figure(fig)
    is_heatmap = bool(output.data) and isinstance(output.data[0], go.Heatmap)
    output.update_layout(
        title=None, template="plotly_white", width=width, height=height,
        font=dict(family="Helvetica, Arial, sans-serif", size=10.5, color=INK),
        paper_bgcolor="white", plot_bgcolor="white", showlegend=not is_heatmap,
        legend=dict(
            orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1.0,
            title_text="", font=dict(size=9.5), bgcolor="rgba(0,0,0,0)",
        ),
        margin=dict(l=8, r=76 if is_heatmap else 8, t=10 if is_heatmap else 30, b=34),
        coloraxis_colorbar=dict(thickness=10, len=0.65, outlinewidth=0, tickfont=dict(size=9)),
    )
    output.update_xaxes(
        automargin=True, gridcolor=GRID, zeroline=False, linecolor=RULE,
        ticks="outside", tickcolor=RULE, ticklen=4, title_font=dict(size=10),
    )
    output.update_yaxes(
        automargin=True, gridcolor=GRID, zeroline=False,
        linecolor=RULE, title_font=dict(size=10),
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
    canvas.line(MARGIN_X, 28, page_w - MARGIN_X, 28)
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(colors.HexColor(MUTED))
    canvas.drawString(MARGIN_X, 17, doc.footer_note)
    canvas.drawRightString(page_w - MARGIN_X, 17, str(canvas.getPageNumber()))
    canvas.restoreState()


def _exec_styles(styles) -> None:
    """Register the paragraph styles used by the executive-summary page."""
    definitions = {
        "ExecTitle": dict(
            fontName="Helvetica-Bold", fontSize=19, leading=23,
            textColor=colors.HexColor(NAVY), spaceAfter=2,
        ),
        "ExecSub": dict(
            fontName="Helvetica", fontSize=10.5, leading=14,
            textColor=colors.HexColor(MUTED),
        ),
        "ExecHeading": dict(
            fontName="Helvetica-Bold", fontSize=13, leading=16,
            textColor=colors.HexColor(NAVY), spaceBefore=6, spaceAfter=6,
        ),
        "ExecBullet": dict(
            fontName="Helvetica", fontSize=10.5, leading=15,
            textColor=colors.HexColor(INK), leftIndent=14, bulletIndent=2, spaceAfter=4,
        ),
        "CardLabel": dict(
            fontName="Helvetica-Bold", fontSize=7, leading=9,
            textColor=colors.HexColor(MUTED),
        ),
        "CardName": dict(
            fontName="Helvetica-Bold", fontSize=12.5, leading=15,
            textColor=colors.HexColor(NAVY),
        ),
        "CardValue": dict(fontName="Helvetica-Bold", fontSize=14, leading=17),
    }
    for name, kwargs in definitions.items():
        if name not in styles:
            styles.add(ParagraphStyle(name=name, parent=styles["BodyText"], **kwargs))


def _metric_card(label: str, name: str, value_str: str, accent: str,
                 width: float, styles) -> Table:
    lab = Paragraph(label.upper(), styles["CardLabel"])
    nm = Paragraph(name, styles["CardName"])
    val = Paragraph(f'<font color="{accent}">{value_str}</font>', styles["CardValue"])
    card = Table([[lab], [nm], [val]], colWidths=[width])
    card.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(PANEL)),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor(RULE)),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("TOPPADDING", (0, 0), (0, 0), 11),
        ("BOTTOMPADDING", (0, 0), (0, 0), 1),
        ("TOPPADDING", (0, 1), (0, 1), 0),
        ("BOTTOMPADDING", (0, 1), (0, 1), 1),
        ("TOPPADDING", (0, 2), (0, 2), 2),
        ("BOTTOMPADDING", (0, 2), (0, 2), 12),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return card


def _card_row(cards: list, col_w: float) -> Table:
    row = Table([cards], colWidths=[col_w] * len(cards))
    row.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-2, -1), 10),
        ("RIGHTPADDING", (-1, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return row


def _ruled_lines(width: float, count: int = 4, row_h: float = 22) -> Table:
    """Blank ruled lines to write commentary on by hand."""
    table = Table([[""] for _ in range(count)], colWidths=[width],
                  rowHeights=[row_h] * count)
    table.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -1), 0.5, colors.HexColor(RULE)),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return table


def _section_heading(text: str, styles, width: float) -> Table:
    heading = Table([[Paragraph(text, styles["ExecHeading"])]], colWidths=[width])
    heading.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -1), 1, colors.HexColor(NAVY)),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return heading


def _extreme_card(extremes: dict, key: str, label: str, column: str, unit: str,
                  accent: str, suffix: str, width: float, styles) -> Table:
    row = extremes.get(key)
    if row is None:
        return _metric_card(f"{label} · {suffix}", "N/A", "N/A", MUTED, width, styles)
    return _metric_card(
        f"{label} · {suffix}", str(row["Bank"]),
        fmt(row[column], unit), accent, width, styles,
    )


def _notes_flowables(notes: list[str], styles, width: float,
                     blank_lines: int = 4) -> list:
    """Render typed commentary as bullets, or blank ruled lines if none given."""
    out: list = [Spacer(1, 8), Paragraph("NOTES", styles["CardLabel"]), Spacer(1, 4)]
    if notes:
        for statement in notes:
            out.append(Paragraph(statement, styles["ExecBullet"], bulletText="•"))
    else:
        out.append(_ruled_lines(width, count=blank_lines))
    return out


def _metric_section(title: str, df: pd.DataFrame, display_periods: list[str],
                    specs: list[tuple], notes: list[str], styles,
                    usable_width: float) -> list:
    """One executive-summary block (CDS or Equity): heading, cards, then notes."""
    col_w = usable_width / 4.0
    card_w = col_w - 10
    flow: list = [_section_heading(title, styles, usable_width), Spacer(1, 8)]

    cards: list = []
    for display_period in display_periods:
        extremes = period_extremes(df, display_period)
        suffix = display_period.upper()
        for key, label, unit, accent, column_prefix in specs:
            cards.append(_extreme_card(
                extremes, key, label, f"{column_prefix} {display_period}",
                unit, accent, suffix, card_w, styles,
            ))
    flow.append(_card_row(cards, col_w))
    flow.extend(_notes_flowables(notes, styles, usable_width))
    return flow


def _exec_summary_flowables(df: pd.DataFrame, period: str, cds_notes: list[str],
                            equity_notes: list[str], styles,
                            usable_width: float) -> list:
    flow: list = [
        Paragraph(f"Executive Summary — Last {period}", styles["ExecTitle"]),
        Paragraph(
            "Largest market movements across bank credit and equity markets.",
            styles["ExecSub"],
        ),
        Spacer(1, 16),
    ]

    # Selected period first, then 5 day as a secondary reference.
    display_periods = [period] + (["5 day"] if period != "5 day" else [])

    cds_specs = [
        ("cds_wide", "Largest CDS widening", "bps", BAD, "CDS"),
        ("cds_tight", "Largest CDS tightening", "bps", GOOD, "CDS"),
    ]
    equity_specs = [
        ("eq_strong", "Strongest equity", "%", GOOD, "Equity"),
        ("eq_weak", "Weakest equity", "%", BAD, "Equity"),
    ]

    flow.extend(_metric_section("CDS", df, display_periods, cds_specs,
                                cds_notes, styles, usable_width))
    flow.append(Spacer(1, 18))
    flow.extend(_metric_section("Equity", df, display_periods, equity_specs,
                                equity_notes, styles, usable_width))
    return flow


def make_pdf(
    df: pd.DataFrame, period: str, selected_regions: Iterable[str],
    top_n: int, cds_notes: list[str], equity_notes: list[str],
) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=PAGE_SIZE, leftMargin=MARGIN_X, rightMargin=MARGIN_X,
        topMargin=MARGIN_TOP, bottomMargin=MARGIN_BOTTOM,
        title="Bank Market Monitor — Chart Pack",
    )
    doc.footer_note = f"Bank Market Monitor  ·  {period}  ·  {', '.join(selected_regions)}"
    styles = getSampleStyleSheet()
    _exec_styles(styles)
    title_style = ParagraphStyle(
        name="ChartTitle", parent=styles["BodyText"], fontName="Helvetica-Bold",
        fontSize=13, leading=TITLE_H, textColor=colors.HexColor(NAVY),
        spaceBefore=0, spaceAfter=0,
    )
    shrink = 0.86
    panel_w = int(doc.width * shrink)
    panel_h = int((doc.height - TITLE_H - TITLE_GAP) * shrink)
    cap = min(top_n, PDF_MAX_ROWS)

    figures: list[go.Figure] = [
        chart_scatter(df, period, label_top=6),
        chart_ranked(df, "CDS", period, cap),
        chart_ranked(df, "Equity", period, cap),
        chart_heatmap(df, "CDS", max_rows=PDF_MAX_ROWS),
        chart_heatmap(df, "Equity", max_rows=PDF_MAX_ROWS),
    ]

    # Page 1: executive summary.
    story: list = _exec_summary_flowables(
        df, period, cds_notes, equity_notes, styles, doc.width
    )
    story.append(PageBreak())

    for index, fig in enumerate(figures):
        heading = (fig.layout.title.text or "").strip()
        png = _fig_to_png(fig, panel_w, panel_h)
        image = RLImage(io.BytesIO(png), width=panel_w, height=panel_h)
        image.hAlign = "CENTER"
        story.append(KeepTogether([
            Paragraph(heading, title_style), Spacer(1, TITLE_GAP), image,
        ]))
        if index < len(figures) - 1:
            story.append(PageBreak())

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return buffer.getvalue()


# --- App --------------------------------------------------------------------
st.title("🏦 Bank Market Monitor")
st.caption("Interactive monitoring of bank CDS spreads and equity performance.")

with st.sidebar:
    st.header("Data and filters")
    uploaded = st.file_uploader("Upload updated Bank Monitoring Model", type=["xlsx"])

    # Clearing the cache forces the workbook to be re-read on the next run.
    if st.button("🔄 Reload data", use_container_width=True,
                 help="Re-read the workbook and rebuild every chart from the latest saved values."):
        st.cache_data.clear()
        st.rerun()

    try:
        source, source_label = _source_bytes(uploaded)
        data = load_bank_monitor(source)
    except Exception as exc:
        st.error(f"Could not read the workbook: {exc}")
        st.stop()

    st.caption(f"Source: {source_label}")

    available_regions = list(dict.fromkeys(data["Region"].dropna().tolist()))
    selected_regions = st.multiselect("Regions", available_regions, default=available_regions)
    period = st.selectbox("Comparison period", PERIODS, index=1)
    top_n = st.slider("Banks in ranked charts", 5, max(5, min(30, len(data))), min(15, len(data)))
    search = st.text_input("Search bank")

    st.markdown("---")
    st.subheader("Executive summary")
    auto_fill = st.checkbox(
        "Auto-fill commentary from data", value=True,
        help="When on, the CDS and Equity notes are generated from the largest moves. "
             "Turn off to leave blank lines you can fill in by hand.",
    )
    cds_notes_raw = st.text_area(
        "CDS commentary (one per line)", height=120,
        placeholder="Leave blank to auto-fill or to get blank lines…",
        help="Appears under the CDS section of the executive-summary page.",
    )
    equity_notes_raw = st.text_area(
        "Equity commentary (one per line)", height=120,
        placeholder="Leave blank to auto-fill or to get blank lines…",
        help="Appears under the Equity section of the executive-summary page.",
    )

filtered = data[data["Region"].isin(selected_regions)].copy()
if search:
    filtered = filtered[filtered["Bank"].str.contains(search, case=False, na=False)]
if filtered.empty:
    st.warning("No banks match the selected filters.")
    st.stop()

# Typed commentary wins; otherwise auto-fill (if enabled) or leave blank for manual entry.
cds_notes_manual = [line.strip() for line in cds_notes_raw.splitlines() if line.strip()]
equity_notes_manual = [line.strip() for line in equity_notes_raw.splitlines() if line.strip()]
cds_notes = cds_notes_manual or (cds_takeaways(filtered, period) if auto_fill else [])
equity_notes = equity_notes_manual or (equity_takeaways(filtered, period) if auto_fill else [])
takeaways = cds_notes + equity_notes

cds_col, eq_col = f"CDS {period}", f"Equity {period}"
worst_cds = filtered.loc[filtered[cds_col].idxmax()] if filtered[cds_col].notna().any() else None
worst_equity = filtered.loc[filtered[eq_col].idxmin()] if filtered[eq_col].notna().any() else None
best_equity = filtered.loc[filtered[eq_col].idxmax()] if filtered[eq_col].notna().any() else None

k1, k2, k3, k4 = st.columns(4)
k1.metric("Banks monitored", len(filtered))
k2.metric(
    "Largest CDS widening", worst_cds["Bank"] if worst_cds is not None else "N/A",
    fmt(worst_cds[cds_col], "bps") if worst_cds is not None else None,
)
k3.metric(
    "Weakest equity market", worst_equity["Bank"] if worst_equity is not None else "N/A",
    fmt(worst_equity[eq_col], "%") if worst_equity is not None else None,
    delta_color="inverse",
)
k4.metric(
    "Strongest equity market", best_equity["Bank"] if best_equity is not None else "N/A",
    fmt(best_equity[eq_col], "%") if best_equity is not None else None,
)

overview_tab, ranked_tab, heatmap_tab, data_tab, report_tab = st.tabs([
    "Overview", "Ranked charts", "Heatmap", "Data explorer", "Report",
])

with overview_tab:
    st.subheader(f"Executive summary — {period}")
    col_cds, col_eq = st.columns(2)
    with col_cds:
        st.markdown("**CDS**")
        if cds_notes:
            for statement in cds_notes:
                st.markdown(f"- {statement}")
        else:
            st.caption("_To be completed._")
    with col_eq:
        st.markdown("**Equity**")
        if equity_notes:
            for statement in equity_notes:
                st.markdown(f"- {statement}")
        else:
            st.caption("_To be completed._")
    st.plotly_chart(chart_scatter(filtered, period), use_container_width=True)

with ranked_tab:
    metric = st.radio("Metric", ["CDS", "Equity"], horizontal=True)
    st.caption(METRIC_INFO[metric]["description"])
    st.plotly_chart(chart_ranked(filtered, metric, period, top_n), use_container_width=True)

with heatmap_tab:
    heat_metric = st.selectbox("Heatmap metric", ["CDS", "Equity"])
    st.plotly_chart(chart_heatmap(filtered, heat_metric), use_container_width=True)

with data_tab:
    display_columns = ["Bank", "Region", "CDS Now", cds_col, eq_col]
    table = filtered[display_columns].sort_values(cds_col, ascending=False)
    st.dataframe(
        table.style.format({
            "CDS Now": "{:.1f}", cds_col: "{:+.1f}", eq_col: "{:+.1f}%",
        }, na_rep="—"),
        use_container_width=True, hide_index=True,
    )
    csv = table.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download filtered data (CSV)", csv,
        "bank_market_snapshot.csv", "text/csv",
    )

with report_tab:
    st.subheader("Downloadable chart pack")
    st.write(
        "A print-ready PDF that opens with an executive-summary page, followed by the "
        "CDS, equity, and heatmap visualizations."
    )
    st.caption(
        "The executive-summary page is split into a CDS section and an Equity section. "
        "Each shows your commentary, the auto-filled notes, or — if you leave a box blank "
        "with auto-fill off — blank lines to complete by hand."
    )
    if st.button("Build PDF", type="primary"):
        with st.spinner("Rendering charts…"):
            try:
                st.session_state["bank_pdf_bytes"] = make_pdf(
                    filtered, period, selected_regions, top_n, cds_notes, equity_notes
                )
            except Exception as exc:
                st.session_state.pop("bank_pdf_bytes", None)
                st.error(
                    f"Could not render the charts to PDF: {exc}\n\n"
                    "Static Plotly export requires the `kaleido` package."
                )
    if st.session_state.get("bank_pdf_bytes"):
        st.download_button(
            "Download PDF", st.session_state["bank_pdf_bytes"],
            "bank_market_charts.pdf", "application/pdf",
        )
    with st.expander("Methodology and interpretation"):
        st.markdown("""
        **CDS:** Higher spreads or positive spread changes normally indicate increased perceived bank credit risk.  
        **Equity:** Negative returns indicate weaker market sentiment toward the bank.  
        **Commentary:** The executive summary has separate CDS and Equity sections. Type notes into the sidebar boxes, let them auto-fill from the largest moves, or leave them blank (with auto-fill off) to print blank lines you can complete by hand.
        """)

st.caption(
    "Data are read from the cached values saved in the Excel workbook. Save the workbook in Excel "
    "(after refreshing the Capital IQ and LSEG plug-ins), then re-upload it or press “Reload data”."
)
