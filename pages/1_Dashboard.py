import altair as alt
import pandas as pd
import streamlit as st

import core
from web_ui import header, kpi_grid, page

PALETTE = [
    "#2869B1", "#00AEEF", "#F15A24", "#1FAF7A", "#F0A202",
    "#DD6E9F", "#4A3AA7", "#64748B", "#37A0A4", "#8B5CF6",
]


def month_label(value):
    return pd.to_datetime(f"{value}-01").strftime("%b %Y")


def monthly_chart(data, suppliers, month_order, y_max):
    return (
        alt.Chart(data)
        .mark_bar()
        .encode(
            x=alt.X(
                "month_label:N",
                title=None,
                sort=month_order,
                axis=alt.Axis(
                    labelAngle=0, labelPadding=10, tickSize=0,
                    domainColor="#CBD5E1",
                ),
            ),
            y=alt.Y(
                "sum(cartons):Q",
                title="Cartons",
                scale=alt.Scale(domain=[0, y_max], nice=True),
                axis=alt.Axis(
                    format="~s", tickCount=5, domain=False,
                    tickSize=0, labelPadding=8, titlePadding=12,
                ),
            ),
            color=alt.Color(
                "supplier_key:N",
                title=None,
                sort=suppliers,
                scale=alt.Scale(range=PALETTE),
                legend=alt.Legend(
                    orient="top", direction="horizontal", columns=5,
                    symbolType="square", labelLimit=210, offset=10,
                ),
            ),
            tooltip=[
                alt.Tooltip("month_label:N", title="Month"),
                alt.Tooltip("supplier_key:N", title="Packaging supplier"),
                alt.Tooltip("sum(cartons):Q", title="Cartons", format=",.0f"),
            ],
        )
        .properties(height=300)
        .configure_view(stroke=None)
        .configure_axis(
            gridColor="#E8EDF3", gridOpacity=1, labelColor="#64748B",
            labelFontSize=11, titleColor="#64748B",
        )
        .configure_legend(
            labelColor="#475569", labelFontSize=11, symbolSize=90, padding=0,
        )
    )


def ranking_table(rows, entity_label):
    data = pd.DataFrame(rows).rename(columns={"name": entity_label, "cartons": "Cartons"})
    st.dataframe(
        data,
        width="stretch",
        hide_index=True,
        column_config={
            entity_label: st.column_config.TextColumn(entity_label),
            "Cartons": st.column_config.NumberColumn("Cartons", format="localized"),
        },
        height=min(390, 38 + 35 * len(data)),
    )


@page("Dashboard")
def main():
    con = core.connect()
    stats = core.dashboard_stats(con)
    period = (f"{stats['first_date']} to {stats['last_date']}"
              if stats["first_date"] else "No tracker data")
    header(
        "Dashboard",
        "Carton shipment activity, data quality and supplier coverage.",
        "Shipment period",
        period,
    )

    quality_issues = stats["undated"] + (stats["unrated"] or 0)
    kpi_grid([
        {"label": "Total cartons", "value": f"{stats['cartons']:,.0f}",
         "note": "Across all dated shipments"},
        {"label": "Tracker rows", "value": f"{stats['rows']:,}",
         "note": "Standard and custom cartons"},
        {"label": "Packaging suppliers", "value": f"{stats['suppliers']:,}",
         "note": "Active in the shipment period"},
        {"label": "Data quality",
         "value": "Clean" if not quality_issues else f"{quality_issues:,} issues",
         "note": f"{stats['undated']:,} undated · {stats['unrated'] or 0:,} without rates",
         "tone": "ok" if not quality_issues else "warn"},
    ])

    monthly = pd.DataFrame(core.monthly_cartons(con))
    with st.container(border=True, key="card_monthly"):
        st.subheader("Monthly cartons")
        if monthly.empty:
            st.info("No tracker data yet. Import a workbook on the Data import page.")
        else:
            suppliers = sorted(monthly["supplier_key"].dropna().astype(str).unique())
            months = sorted(monthly["month"].dropna().astype(str).unique())
            supplier_col, month_col = st.columns([3, 2], gap="medium")
            selected_suppliers = supplier_col.multiselect(
                "Packaging suppliers",
                suppliers,
                default=suppliers,
                placeholder="Choose packaging suppliers",
                key="dashboard_supplier_filter",
            )
            if len(months) > 1:
                first_month, last_month = month_col.select_slider(
                    "Months",
                    options=months,
                    value=(months[0], months[-1]),
                    format_func=month_label,
                    key="dashboard_month_filter",
                )
            else:
                month_col.selectbox(
                    "Months", months, format_func=month_label,
                    disabled=True, key="dashboard_single_month",
                )
                first_month = last_month = months[0]

            filtered = monthly[
                monthly["supplier_key"].isin(selected_suppliers)
                & monthly["month"].between(first_month, last_month)
            ].copy()
            if filtered.empty:
                st.info("Select at least one packaging supplier with cartons in this month range.")
            else:
                visible_months = [m for m in months if first_month <= m <= last_month]
                month_order = [month_label(m) for m in visible_months]
                filtered["month_label"] = filtered["month"].map(month_label)
                peak = float(filtered.groupby("month")["cartons"].sum().max())
                y_max = max(1.0, peak * 1.08)
                st.altair_chart(
                    monthly_chart(filtered, selected_suppliers, month_order, y_max),
                    width="stretch",
                    key="monthly_cartons_chart",
                )
                st.caption(
                    "Stacked by packaging supplier. The carton scale adjusts to the selected data."
                )

    rankings = [
        ("Top suppliers", core.top_tracker_entities(con, "supplier_name"),
         "Supplier", "card_top_suppliers"),
        ("Top factories", core.top_tracker_entities(con, "factory_name"),
         "Factory", "card_top_factories"),
    ]
    rankings = [item for item in rankings if item[1]]
    if rankings:
        columns = st.columns(len(rankings), gap="medium") if len(rankings) > 1 else [st]
        for column, (title, rows, label, key) in zip(columns, rankings):
            with column.container(border=True, key=key):
                st.subheader(title)
                ranking_table(rows, label)


main()
