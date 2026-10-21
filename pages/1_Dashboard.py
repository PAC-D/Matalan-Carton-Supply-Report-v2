import altair as alt
import pandas as pd
import streamlit as st

import core
from web_ui import header, kpi_grid, page

PALETTE = ["#2869B1", "#00AEEF", "#F15A24", "#1FAF7A", "#F0A202", "#DD6E9F"]
TYPE_COLORS = ["#2869B1", "#00AEEF"]


def stacked_chart(data, color, colors, tooltip_title):
    order = [f"Q{quarter} ({core.QUARTER_MONTHS[quarter]})"
             for quarter in sorted(data["quarter"].unique())]
    return (
        alt.Chart(data)
        .mark_bar()
        .encode(
            x=alt.X("quarter_label:N", title=None, sort=order,
                    axis=alt.Axis(labelAngle=0, labelPadding=10, tickSize=0,
                                  domainColor="#CBD5E1")),
            y=alt.Y("sum(cartons):Q", title="Cartons",
                    axis=alt.Axis(format="~s", tickCount=6, domain=False,
                                  tickSize=0, labelPadding=8, titlePadding=12)),
            color=alt.Color(f"{color}:N", title=None,
                            scale=alt.Scale(range=colors),
                            legend=alt.Legend(orient="top", direction="horizontal",
                                              symbolType="square", labelLimit=230)),
            tooltip=[
                alt.Tooltip("quarter_label:N", title="Quarter"),
                alt.Tooltip(f"{color}:N", title=tooltip_title),
                alt.Tooltip("sum(cartons):Q", title="Cartons", format=",.0f"),
            ],
        )
        .properties(height=300)
        .configure_view(stroke=None)
        .configure_axis(gridColor="#E8EDF3", labelColor="#64748B",
                        labelFontSize=11, titleColor="#64748B")
        .configure_legend(labelColor="#475569", labelFontSize=11, symbolSize=90)
    )


def rate_text(value):
    return f"{value * 100:g}%" if isinstance(value, (int, float)) else value


def summary_table(data):
    table = data.rename(columns={
        "packaging_supplier": "Packaging supplier",
        "quarter_label": "Quarter",
        "carton_type": "Carton type",
        "pacd_rate": "PACD rate",
        "rebate_rate": "Rebate rate",
        "cartons": "Carton QTY",
        "total_amount": "Total amount",
        "pacd": "PACD",
        "matalan_rebate": "Matalan rebate",
    }).copy()
    table["PACD rate"] = table["PACD rate"].map(rate_text)
    table["Rebate rate"] = table["Rebate rate"].map(rate_text)
    shown = ["Packaging supplier", "Quarter", "Carton type", "PACD rate",
             "Rebate rate", "Carton QTY", "Total amount", "PACD", "Matalan rebate"]
    st.dataframe(
        table[shown], hide_index=True, height=min(520, 38 + 35 * len(table)),
        column_config={
            "Packaging supplier": st.column_config.TextColumn(pinned=True),
            "Carton QTY": st.column_config.NumberColumn(format="localized"),
            **{column: st.column_config.NumberColumn(format="$ %.2f")
               for column in ("Total amount", "PACD", "Matalan rebate")},
        },
        key="live_summary_table",
    )


def cross_reference_section(con, selected_suppliers):
    with st.container(border=True, key="card_cross_reference"):
        st.subheader("Tracker and eBiz cross-reference")
        st.caption("Matches the complete Tracker and eBiz datasets, exactly like the generated Cross Reference report.")
        if core.ebiz_status(con) is None:
            st.info("No eBiz data is loaded. Import all current vendor exports to enable reconciliation.")
            return
        rows = [row for row in core.cross_reference_overview(con)
                if row["supplier_key"] in selected_suppliers]
        if not rows:
            st.info("No reconciliation data exists for the selected packaging suppliers.")
            return
        matched = sum(row["matches"] for row in rows)
        missing = sum(row["not_in_ebiz"] for row in rows)
        mismatched = sum(row["quantity_mismatches"] for row in rows)
        ebiz_only = sum(row["ebiz_only"] for row in rows)
        kpi_grid([
            {"label": "Matched orders", "value": f"{matched:,}", "tone": "ok"},
            {"label": "Missing from eBiz", "value": f"{missing:,}",
             "tone": "warn" if missing else "ok"},
            {"label": "Quantity mismatches", "value": f"{mismatched:,}",
             "tone": "warn" if mismatched else "ok"},
            {"label": "eBiz-only orders", "value": f"{ebiz_only:,}",
             "tone": "warn" if ebiz_only else "ok"},
        ])
        table = pd.DataFrame(rows).rename(columns={
            "packaging_supplier": "Packaging supplier",
            "tracker_orders": "Tracker orders",
            "matches": "Matched",
            "not_in_ebiz": "Missing from eBiz",
            "quantity_mismatches": "QTY mismatches",
            "ebiz_only": "eBiz only",
        })
        st.dataframe(
            table[["Packaging supplier", "Tracker orders", "Matched",
                   "Missing from eBiz", "QTY mismatches", "eBiz only"]],
            hide_index=True,
            column_config={"Packaging supplier": st.column_config.TextColumn(pinned=True)},
            key="cross_reference_table",
        )


@st.fragment
def live_summary():
    con = core.connect()
    years = core.fiscal_years(con)
    year = st.selectbox("Fiscal year", years, index=len(years) - 1,
                        key="dashboard_fiscal_year", width=220)
    rows = core.summary_rows(con, year)
    if not rows:
        st.info("No Summary data exists for this fiscal year.")
        return
    data = pd.DataFrame(rows)
    supplier_names = (data[["supplier_key", "packaging_supplier"]]
                      .drop_duplicates().set_index("supplier_key")["packaging_supplier"].to_dict())
    selected = st.multiselect(
        "Packaging suppliers", list(supplier_names), default=list(supplier_names),
        format_func=supplier_names.get, key="dashboard_summary_suppliers",
        placeholder="Choose packaging suppliers",
    )
    if not selected:
        st.warning("Choose at least one packaging supplier.")
        return
    data = data[data["supplier_key"].isin(selected)].copy()

    standard = int(data.loc[data["carton_type"] == "Standard", "cartons"].sum())
    custom = int(data.loc[data["carton_type"] == "Custom", "cartons"].sum())
    total = standard + custom
    kpi_grid([
        {"label": "Total cartons", "value": f"{total:,}", "note": f"Fiscal year {year}"},
        {"label": "Standard cartons", "value": f"{standard:,}",
         "note": f"{standard / total:.1%} of total" if total else "No cartons"},
        {"label": "Custom cartons", "value": f"{custom:,}",
         "note": f"{custom / total:.1%} of total" if total else "No cartons"},
        {"label": "Packaging suppliers", "value": f"{len(selected):,}",
         "note": "Included in this view"},
    ])

    with st.container(border=True, key="card_quarterly_suppliers"):
        st.subheader("Quarterly carton sales by packaging supplier")
        st.altair_chart(
            stacked_chart(data, "packaging_supplier", PALETTE, "Packaging supplier"),
            key="quarterly_supplier_chart",
        )
        st.caption("The live chart uses the same fiscal-quarter totals as the generated Summary workbook.")

    left, right = st.columns([2, 3], gap="medium")
    with left.container(border=True, key="card_carton_types"):
        st.subheader("Standard and custom cartons")
        st.altair_chart(
            stacked_chart(data, "carton_type", TYPE_COLORS, "Carton type"),
            key="quarterly_type_chart",
        )
    with right.container(border=True, key="card_live_summary"):
        st.subheader("Live Summary sheet")
        st.caption("These are the rows used to build the Summary Excel report.")
        summary_table(data)

    cross_reference_section(con, selected)


@page("Dashboard")
def main():
    con = core.connect()
    years = core.fiscal_years(con)
    header(
        "Dashboard",
        "A live view of the Summary report and Tracker-to-eBiz reconciliation.",
        "Data source",
        "Live tracker database" if years else "No tracker data",
    )
    if not years:
        st.info("No tracker data yet. Import a workbook on the Data import page.")
        return
    live_summary()


main()
