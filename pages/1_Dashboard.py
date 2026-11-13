import altair as alt
import pandas as pd
import streamlit as st

import core
from web_ui import (
    dashboard_header,
    dashboard_interactive_metric_grid,
    interactive_kpi_grid,
    page,
    search_frame,
    table_search,
)


SUPPLIER_PALETTE = [
    "#B6193B", "#173B6D", "#D84A67", "#31598D", "#7C102A", "#6E88AD",
]
TYPE_COLORS = ["#173B6D", "#B6193B"]


def stacked_chart(data, color, colors, tooltip_title, height=270):
    year = int(data["fiscal_year"].iloc[0])
    order = [
        f"Q{quarter} ({core.quarter_months(year, quarter)})"
        for quarter in sorted(data["quarter"].unique())
    ]
    return (
        alt.Chart(data)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
        .encode(
            x=alt.X(
                "quarter_label:N", title=None, sort=order,
                axis=alt.Axis(labelAngle=0, labelPadding=9, tickSize=0, domain=False),
            ),
            y=alt.Y(
                "sum(cartons):Q", title="Cartons",
                axis=alt.Axis(
                    format="~s", tickCount=5, domain=False, tickSize=0,
                    labelPadding=7, titlePadding=10,
                ),
            ),
            color=alt.Color(
                f"{color}:N", title=None, scale=alt.Scale(range=colors),
                legend=alt.Legend(
                    orient="top", direction="horizontal", symbolType="circle",
                    symbolSize=85, labelLimit=240, offset=4,
                ),
            ),
            tooltip=[
                alt.Tooltip("quarter_label:N", title="Quarter"),
                alt.Tooltip(f"{color}:N", title=tooltip_title),
                alt.Tooltip("sum(cartons):Q", title="Cartons", format=",.0f"),
            ],
        )
        .properties(height=height)
        .configure_view(stroke=None)
        .configure_axis(
            gridColor="#E9EEF5", gridOpacity=1, labelColor="#687A92",
            labelFont="Inter", labelFontSize=11, titleColor="#687A92",
            titleFont="Inter", titleFontSize=11,
        )
        .configure_legend(
            labelColor="#52647D", labelFont="Inter", labelFontSize=11,
            symbolStrokeWidth=0,
        )
    )


def rate_text(value):
    return f"{value * 100:g}%" if isinstance(value, (int, float)) else value


def summary_table(data):
    table = data.rename(columns={
        "packaging_supplier": "Packaging supplier",
        "quarter_label": "Quarter",
        "carton_type": "Carton type",
        "pacd_rate": "PACD Rate",
        "rebate_rate": "Matalan Rate",
        "cartons": "Carton QTY",
        "total_amount": "Total amount",
        "pacd": "PACD Rebate",
        "matalan_rebate": "Matalan rebate",
    }).copy()
    table["PACD Rate"] = table["PACD Rate"].map(rate_text)
    table["Matalan Rate"] = table["Matalan Rate"].map(rate_text)
    shown = [
        "Packaging supplier", "Quarter", "Carton type", "PACD Rate",
        "Matalan Rate", "Carton QTY", "Total amount", "PACD Rebate", "Matalan rebate",
    ]
    query = table_search("Search live summary", "dashboard_summary_search")
    table = search_frame(table[shown], query)
    if table.empty:
        st.info("No summary rows match your search.")
        return
    st.dataframe(
        table,
        hide_index=True,
        height=min(420, 38 + 35 * len(table)),
        width="stretch",
        column_config={
            "Packaging supplier": st.column_config.TextColumn(pinned=True),
            "Carton QTY": st.column_config.NumberColumn(format="localized"),
            **{
                column: st.column_config.NumberColumn(format="$ %.2f")
                for column in ("Total amount", "PACD Rebate", "Matalan rebate")
            },
        },
        key="live_summary_table",
    )


def supplier_totals_table(data):
    table = (
        data.groupby("packaging_supplier", as_index=False)[
            ["cartons", "total_amount", "pacd", "matalan_rebate"]
        ]
        .sum()
        .rename(columns={
            "packaging_supplier": "Packaging supplier",
            "cartons": "Carton QTY",
            "total_amount": "Total amount",
            "pacd": "PACD Rebate",
            "matalan_rebate": "Matalan rebate",
        })
    )
    table = search_frame(
        table,
        table_search("Search supplier totals", "dashboard_supplier_totals_search"),
    )
    if table.empty:
        st.info("No supplier totals match your search.")
        return
    st.dataframe(
        table,
        hide_index=True,
        height=min(420, 38 + 35 * len(table)),
        width="stretch",
        column_config={
            "Packaging supplier": st.column_config.TextColumn(pinned=True),
            "Carton QTY": st.column_config.NumberColumn(format="localized"),
            **{
                column: st.column_config.NumberColumn(format="$ %.2f")
                for column in ("Total amount", "PACD Rebate", "Matalan rebate")
            },
        },
        key="dashboard_supplier_totals_table",
    )


def cross_reference_section(con, selected_suppliers):
    with st.container(border=True, key="dashboard_panel_cross_reference"):
        st.subheader("Tracker and eBiz cross-reference")
        st.caption("Selected suppliers · complete reconciliation dataset")
        if core.ebiz_status(con) is None:
            st.info("No eBiz data is loaded. Import vendor exports to enable reconciliation.")
            return

        rows = [
            row for row in core.cross_reference_overview(con)
            if row["supplier_key"] in selected_suppliers
        ]
        if not rows:
            st.info("No reconciliation data exists for the selected packaging suppliers.")
            return

        active = interactive_kpi_grid([
            {
                "id": "all", "label": "Tracker orders",
                "value": f"{sum(row['tracker_orders'] for row in rows):,}",
            },
            {
                "id": "matches", "label": "Matched",
                "value": f"{sum(row['matches'] for row in rows):,}",
            },
            {
                "id": "not_in_ebiz", "label": "Missing in eBiz",
                "value": f"{sum(row['not_in_ebiz'] for row in rows):,}",
            },
            {
                "id": "quantity_mismatches", "label": "QTY mismatch",
                "value": f"{sum(row['quantity_mismatches'] for row in rows):,}",
            },
            {
                "id": "ebiz_only", "label": "eBiz only",
                "value": f"{sum(row['ebiz_only'] for row in rows):,}",
            },
        ], "dashboard_reconciliation")

        table = pd.DataFrame(rows).rename(columns={
            "packaging_supplier": "Packaging supplier",
            "tracker_orders": "Tracker orders",
            "matches": "Matched",
            "not_in_ebiz": "Missing from eBiz",
            "quantity_mismatches": "QTY mismatches",
            "ebiz_only": "eBiz only",
        })
        shown = table[[
            "Packaging supplier", "Tracker orders", "Matched",
            "Missing from eBiz", "QTY mismatches", "eBiz only",
        ]]
        active_column = {
            "matches": "Matched",
            "not_in_ebiz": "Missing from eBiz",
            "quantity_mismatches": "QTY mismatches",
            "ebiz_only": "eBiz only",
        }.get(active)
        if active_column:
            shown = shown[shown[active_column] > 0]
        shown = search_frame(
            shown,
            table_search("Search reconciliation", "dashboard_reconciliation_search"),
        )
        if shown.empty:
            st.info("No reconciliation rows match the active metric and search.")
            return
        st.dataframe(
            shown,
            hide_index=True,
            height=min(310, 38 + 35 * len(shown)),
            width="stretch",
            column_config={"Packaging supplier": st.column_config.TextColumn(pinned=True)},
            key="cross_reference_table",
        )


@st.fragment
@page
def live_summary(years):
    con = core.connect()

    with st.container(key="dashboard_filter_bar"):
        year_column, quarter_column, supplier_column = st.columns(
            [1.05, 1.25, 3.7], gap="small"
        )
        with year_column:
            year = st.selectbox(
                "Fiscal year", years, index=len(years) - 1,
                key="dashboard_fiscal_year",
            )

        rows = core.summary_rows(con, year)
        if not rows:
            st.info("No summary data exists for this fiscal year.")
            return
        year_data = pd.DataFrame(rows)

        available_quarters = sorted(year_data["quarter"].unique())
        with quarter_column:
            quarter = st.selectbox(
                "Quarter",
                [None, *available_quarters],
                format_func=lambda value: "All quarters" if value is None else (
                    f"Q{value} · {core.quarter_months(year, value)}"
                ),
                key=f"dashboard_quarter_{year}",
            )

        supplier_names = (
            year_data[["supplier_key", "packaging_supplier"]]
            .drop_duplicates()
            .set_index("supplier_key")["packaging_supplier"]
            .to_dict()
        )
        with supplier_column:
            selected = st.multiselect(
                "Packaging suppliers",
                list(supplier_names),
                default=list(supplier_names),
                format_func=supplier_names.get,
                key=f"dashboard_summary_suppliers_{year}",
                placeholder="Select packaging suppliers",
            )

    if not selected:
        st.warning("Select at least one packaging supplier.")
        return

    data = year_data[year_data["supplier_key"].isin(selected)].copy()
    if quarter is not None:
        data = data[data["quarter"] == quarter].copy()
    if data.empty:
        st.info("No summary data matches the selected filters.")
        return

    total_cartons = int(data["cartons"].sum())
    standard = int(data.loc[data["carton_type"] == "Standard", "cartons"].sum())
    custom = int(data.loc[data["carton_type"] == "Custom", "cartons"].sum())
    period_label = "All quarters" if quarter is None else f"Q{quarter}"

    active = dashboard_interactive_metric_grid([
        {
            "id": "all", "label": "Total cartons", "value": f"{total_cartons:,}",
            "note": f"{period_label} · active supplier selection", "tone": "navy",
        },
        {
            "id": "suppliers", "label": "Packaging suppliers", "value": f"{len(selected):,}",
            "note": "Included in this view", "tone": "red",
        },
        {
            "id": "Standard", "label": "Standard cartons", "value": f"{standard:,}",
            "note": f"{period_label} · {len(selected)} supplier{'s' if len(selected) != 1 else ''}",
        },
        {
            "id": "Custom", "label": "Custom cartons", "value": f"{custom:,}",
            "note": f"{period_label} · {len(selected)} supplier{'s' if len(selected) != 1 else ''}",
        },
    ], "dashboard_summary")

    shown_data = (
        data[data["carton_type"] == active].copy()
        if active in {"Standard", "Custom"}
        else data
    )
    supplier_label = "packaging supplier" if len(selected) == 1 else "packaging suppliers"
    quarterly_title = {
        "all": "Quarterly total carton sales by packaging supplier",
        "suppliers": f"Quarterly carton sales across {len(selected):,} {supplier_label}",
        "Standard": "Quarterly standard carton sales by packaging supplier",
        "Custom": "Quarterly custom carton sales by packaging supplier",
    }[active]

    st.html("<div class='mc-dashboard-section-label'>Volume overview</div>")
    with st.container(border=True, key="dashboard_panel_quarterly_suppliers"):
        st.subheader(quarterly_title)
        st.caption(f"{period_label} · {len(selected)} packaging supplier{'s' if len(selected) != 1 else ''}")
        st.altair_chart(
            stacked_chart(shown_data, "packaging_supplier", SUPPLIER_PALETTE, "Packaging supplier"),
            key="quarterly_supplier_chart",
            width="stretch",
        )

    left, right = st.columns([2, 3], gap="small")
    with left.container(border=True, key="dashboard_panel_carton_types"):
        st.subheader("Standard and custom cartons")
        st.caption("Carton mix across the active selection")
        st.altair_chart(
            stacked_chart(shown_data, "carton_type", TYPE_COLORS, "Carton type", height=310),
            key="quarterly_type_chart",
            width="stretch",
        )
    with right.container(border=True, key="dashboard_panel_live_summary"):
        if active == "suppliers":
            st.subheader("Packaging supplier totals")
            st.caption("Aggregated totals for the active period and supplier selection")
            supplier_totals_table(data)
        else:
            st.subheader("Live Summary sheet")
            st.caption("Searchable rows from the active fiscal year, quarter, and supplier selection")
            summary_table(shown_data)

    cross_reference_section(con, selected)


@page
def main():
    con = core.connect()
    years = core.fiscal_years(con)
    dashboard_header()
    if not years:
        st.info("No tracker data yet. Import a workbook on the Data import page.")
        return
    live_summary(years)


main()
