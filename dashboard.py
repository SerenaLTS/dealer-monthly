import os
import re

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Monthly Dealer Dashboard",
    layout="wide",
)


MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

ENQUIRY_COLS = [
    "TestDrive",
    "OfferEnquiry",
    "FleetEnquiry",
    "ShowroomEnquiry",
    "New",
    "Used",
    "Tradein",
]

ENQUIRY_LABELS = {
    "TestDrive": "Test Drive",
    "OfferEnquiry": "Offer Enquiry",
    "FleetEnquiry": "Fleet Enquiry",
    "ShowroomEnquiry": "Showroom Enquiry",
    "New": "New",
    "Used": "Used",
    "Tradein": "Trade-in",
}

SALES_TYPE_LABELS = {
    "21": "Private Local Delivery",
    "47": "Large Fleet",
    "48": "Fleet",
    "59": "Dealer Demo",
    "33": "Local Government",
    "38": "Company Capitalisation",
}

SALES_TYPE_GROUPS = {
    "21": "Private",
    "47": "Fleet",
    "48": "Fleet",
    "33": "Fleet",
    "38": "Fleet",
    "59": "Dealer Demo",
}


def normalise_code(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def normalise_name(value):
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().lower().split())


def parse_sales_column(column_name):
    text = str(column_name)
    match = re.match(r"^(\d{2})(.+)$", text)
    if not match:
        return "", "Other", text

    sales_type_code, model = match.groups()
    return (
        sales_type_code,
        SALES_TYPE_LABELS.get(sales_type_code, f"Type {sales_type_code}"),
        model,
        SALES_TYPE_GROUPS.get(sales_type_code, "Other"),
    )


def make_sales_long(source, group_cols, sales_cols):
    if not sales_cols:
        return pd.DataFrame(columns=group_cols + ["sales_column", "sales", "sales_type_code", "sales_type", "model"])

    grouped = (
        source.groupby(group_cols, dropna=False)[sales_cols]
        .sum()
        .reset_index()
    )
    melted = grouped.melt(
        id_vars=group_cols,
        value_vars=sales_cols,
        var_name="sales_column",
        value_name="sales",
    )
    dimensions = pd.DataFrame(
        [
            {
                "sales_column": col,
                "sales_type_code": parse_sales_column(col)[0],
                "sales_type": parse_sales_column(col)[1],
                "model": parse_sales_column(col)[2],
                "sales_group": parse_sales_column(col)[3],
            }
            for col in sales_cols
        ]
    )
    melted = melted.merge(dimensions, on="sales_column", how="left")
    return melted[melted["sales"] > 0].copy()


@st.cache_data
def load_data():
    database_path = "database.xlsx"
    dealer_info_path = "dealer_info.xlsx"

    if not os.path.exists(database_path):
        raise FileNotFoundError(f"Cannot find {database_path}")
    if not os.path.exists(dealer_info_path):
        raise FileNotFoundError(f"Cannot find {dealer_info_path}")

    df = pd.read_excel(database_path)
    dealer_info = pd.read_excel(dealer_info_path)

    required_db_cols = {"dealer_code", "Dealername", "month"}
    missing_db_cols = required_db_cols - set(df.columns)
    if missing_db_cols:
        raise ValueError(f"database.xlsx is missing columns: {sorted(missing_db_cols)}")

    dealer_name_col = next(
        (col for col in ["final_dealer_name", "dealer_name", "Dealername", "Name"] if col in dealer_info.columns),
        None,
    )
    state_col = next(
        (col for col in ["final_state", "dealer_state", "state", "State"] if col in dealer_info.columns),
        None,
    )
    if dealer_name_col is None or state_col is None:
        raise ValueError(
            "dealer_info.xlsx must include a dealer name column and a state column. "
            "Accepted dealer name columns: final_dealer_name, dealer_name, Dealername, Name. "
            "Accepted state columns: final_state, dealer_state, state, State."
        )
    if "dealer_code" not in dealer_info.columns:
        dealer_info["dealer_code"] = ""

    df["dealer_code_key"] = df["dealer_code"].apply(normalise_code)
    df["dealer_name_raw"] = df["Dealername"].astype(str).str.strip()
    df["dealer_name_key"] = df["dealer_name_raw"].apply(normalise_name)
    df["month"] = df["month"].astype(str).str.strip()
    df["month_order"] = df["month"].map({month: i for i, month in enumerate(MONTH_ORDER, start=1)})

    dealer_info["dealer_code_key"] = dealer_info["dealer_code"].apply(normalise_code)
    dealer_info["dealer_name_info"] = dealer_info[dealer_name_col].astype(str).str.strip()
    dealer_info["dealer_state_info"] = dealer_info[state_col].astype(str).str.strip()
    dealer_info["info_name_key"] = dealer_info["dealer_name_info"].apply(normalise_name)
    if "dealer_name_database" in dealer_info.columns:
        dealer_info["database_name_key"] = dealer_info["dealer_name_database"].apply(normalise_name)
    else:
        dealer_info["database_name_key"] = ""

    by_code = (
        dealer_info[["dealer_code_key", "dealer_name_info", "dealer_state_info"]]
        .dropna(how="all")
        .drop_duplicates("dealer_code_key")
    )
    by_code = by_code[by_code["dealer_code_key"] != ""]

    df = df.merge(by_code, on="dealer_code_key", how="left")

    info_name_lookup = dealer_info[["info_name_key", "dealer_name_info", "dealer_state_info"]].rename(
        columns={"info_name_key": "dealer_name_key"}
    )
    database_name_lookup = dealer_info[["database_name_key", "dealer_name_info", "dealer_state_info"]].rename(
        columns={"database_name_key": "dealer_name_key"}
    )
    name_lookup = (
        pd.concat([info_name_lookup, database_name_lookup], ignore_index=True)
        .drop_duplicates("dealer_name_key")
    )
    name_lookup = name_lookup[name_lookup["dealer_name_key"] != ""]

    by_name = df[["dealer_name_key"]].merge(name_lookup, on="dealer_name_key", how="left")
    df["dealer_name"] = df["dealer_name_info"].fillna(by_name["dealer_name_info"]).fillna(df["dealer_name_raw"])
    df["dealer_state"] = df["dealer_state_info"].fillna(by_name["dealer_state_info"]).fillna("Unassigned")
    df["dealer_state"] = df["dealer_state"].astype(str).str.strip().replace({"": "Unassigned", "nan": "Unassigned"})
    if "active" not in df.columns:
        df["active"] = "Inactive"
    df["active"] = df["active"].astype(str).str.strip().str.title()
    df["active"] = df["active"].where(df["active"].isin(["Active", "Inactive"]), "Inactive")
    df["active_flag"] = df["active"].eq("Active")

    for col in ENQUIRY_COLS:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    id_cols = {
        "dealer_code",
        "Dealername",
        "month",
        "dealer_code_key",
        "dealer_name_raw",
        "dealer_name_key",
        "month_order",
        "dealer_name_info",
        "dealer_state_info",
        "dealer_name",
        "dealer_state",
        "active",
        "active_flag",
    }
    sales_cols = [col for col in df.columns if col not in id_cols and col not in ENQUIRY_COLS]
    for col in sales_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    private_sales_cols = [col for col in sales_cols if parse_sales_column(col)[3] == "Private"]
    fleet_sales_cols = [col for col in sales_cols if parse_sales_column(col)[3] == "Fleet"]
    demo_sales_cols = [col for col in sales_cols if parse_sales_column(col)[3] == "Dealer Demo"]

    df["total_enquiry"] = df[ENQUIRY_COLS].sum(axis=1)
    df["total_sales"] = df[sales_cols].sum(axis=1)
    df["private_enquiry"] = df["total_enquiry"] - df["FleetEnquiry"]
    df["fleet_enquiry"] = df["FleetEnquiry"]
    df["private_sales"] = df[private_sales_cols].sum(axis=1) if private_sales_cols else 0
    df["fleet_sales"] = df[fleet_sales_cols].sum(axis=1) if fleet_sales_cols else 0
    df["dealer_demo_sales"] = df[demo_sales_cols].sum(axis=1) if demo_sales_cols else 0
    df["conversion_rate"] = (df["total_sales"] / df["total_enquiry"]).where(df["total_enquiry"] > 0, 0)
    df["private_conversion_rate"] = (df["private_sales"] / df["private_enquiry"]).where(df["private_enquiry"] > 0, 0)
    df["fleet_conversion_rate"] = (df["fleet_sales"] / df["fleet_enquiry"]).where(df["fleet_enquiry"] > 0, 0)

    return df, sales_cols


def format_int(value):
    return f"{int(value):,}"


def format_rate_columns(frame, columns):
    formatted = frame.copy()
    for col in columns:
        if col in formatted.columns:
            formatted[col] = formatted[col].map(lambda value: f"{value:.1%}" if pd.notna(value) else "")
    return formatted


df, sales_cols = load_data()

st.title("Monthly Dealer Enquiry & Sales Dashboard")

with st.expander("Data source & definitions"):
    st.markdown(
        """
        **Data sources**
        - `database.xlsx`: monthly sales and dealer enquiry fact table
        - `dealer_info.xlsx`: dealer name and state mapping table

        **Definitions**
        - Dealer enquiry = TestDrive + OfferEnquiry + FleetEnquiry + ShowroomEnquiry + New + Used + Tradein
        - Sales columns combine two dimensions: first two digits = sales/customer type, remaining text = model
        - Conversion rate = Sales / Dealer enquiry
        - Private conversion = Private Local Delivery sales / non-FleetEnquiry enquiries
        - Fleet conversion = Fleet, Large Fleet, Local Government, and Company Capitalisation sales / FleetEnquiry
        - Dealer Demo sales are shown in sales totals and sales breakdowns, but excluded from Private/Fleet conversion
        - Fleet conversion should be treated as directional only, because FleetEnquiry may not capture every fleet lead that later becomes a fleet sale
        """
    )

if st.sidebar.button("Force reload data"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.header("Filters")

available_months = sorted(df["month"].dropna().unique(), key=lambda x: MONTH_ORDER.index(x) if x in MONTH_ORDER else 99)
selected_months = st.sidebar.multiselect("Month", available_months, default=available_months)

available_states = sorted(df["dealer_state"].dropna().unique())
selected_states = st.sidebar.multiselect("State", available_states, default=available_states)

active_options = ["Active", "Inactive"]
selected_active_status = st.sidebar.multiselect("Active status", active_options, default=active_options)

filtered = df.copy()
if selected_months:
    filtered = filtered[filtered["month"].isin(selected_months)]
if selected_states:
    filtered = filtered[filtered["dealer_state"].isin(selected_states)]
if selected_active_status:
    filtered = filtered[filtered["active"].isin(selected_active_status)]

if filtered.empty:
    st.warning("No data matches the selected filters.")
    st.stop()

total_enquiry = filtered["total_enquiry"].sum()
total_sales = filtered["total_sales"].sum()
private_enquiry = filtered["private_enquiry"].sum()
private_sales = filtered["private_sales"].sum()
fleet_enquiry = filtered["fleet_enquiry"].sum()
fleet_sales = filtered["fleet_sales"].sum()
total_dealers = filtered["dealer_name"].nunique()
active_dealers = filtered.loc[filtered["active_flag"], "dealer_name"].nunique()
active_states = filtered["dealer_state"].nunique()
conversion_rate = total_sales / total_enquiry if total_enquiry else 0
private_conversion_rate = private_sales / private_enquiry if private_enquiry else 0
fleet_conversion_rate = fleet_sales / fleet_enquiry if fleet_enquiry else 0

kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
kpi1.metric("Dealer Enquiry", format_int(total_enquiry))
kpi2.metric("Sales", format_int(total_sales))
kpi3.metric("Conversion", f"{conversion_rate:.1%}")
kpi4.metric("Private Conv.", f"{private_conversion_rate:.1%}")
kpi5.metric("Fleet Conv.", f"{fleet_conversion_rate:.1%}")

kpi6, kpi7, kpi8 = st.columns(3)
kpi6.metric("Dealers", format_int(total_dealers))
kpi7.metric("Active Dealers", format_int(active_dealers))
kpi8.metric("States", format_int(active_states))

st.markdown("---")

monthly = (
    filtered.groupby(["month", "month_order"], dropna=False)[["total_enquiry", "total_sales"] + ENQUIRY_COLS]
    .sum()
    .reset_index()
    .sort_values("month_order")
)
monthly_enquiry_melted = monthly.melt(
    id_vars=["month", "month_order"],
    value_vars=ENQUIRY_COLS,
    var_name="channel",
    value_name="enquiries",
)
monthly_enquiry_melted["channel"] = monthly_enquiry_melted["channel"].map(ENQUIRY_LABELS)

st.subheader("Monthly Summary - Filtered All Dealers")
st.caption("Scope: all dealers matching the Month and State filters.")
summary_left, summary_right = st.columns([2, 1])

with summary_left:
    fig_monthly_enquiry = px.bar(
        monthly_enquiry_melted,
        x="month",
        y="enquiries",
        color="channel",
        title="Monthly Dealer Enquiry Breakdown by Channel",
        labels={"month": "Month", "enquiries": "Dealer Enquiry", "channel": "Channel"},
    )
    st.plotly_chart(fig_monthly_enquiry, use_container_width=True)

with summary_right:
    fig_monthly_sales = px.bar(
        monthly,
        x="month",
        y="total_sales",
        title="Monthly Sales Total",
        labels={"month": "Month", "total_sales": "Sales"},
    )
    st.plotly_chart(fig_monthly_sales, use_container_width=True)

left, right = st.columns(2)

state_summary = (
    filtered.groupby("dealer_state", dropna=False)
    .agg(
        dealer_enquiry=("total_enquiry", "sum"),
        sales=("total_sales", "sum"),
        private_enquiry=("private_enquiry", "sum"),
        private_sales=("private_sales", "sum"),
        fleet_enquiry=("fleet_enquiry", "sum"),
        fleet_sales=("fleet_sales", "sum"),
        dealers=("dealer_name", "nunique"),
        active_dealers=("dealer_name", lambda x: x[filtered.loc[x.index, "active_flag"]].nunique()),
    )
    .reset_index()
    .sort_values("dealer_enquiry", ascending=False)
)
state_summary["inactive_dealers"] = state_summary["dealers"] - state_summary["active_dealers"]
state_summary["conversion_rate"] = (
    state_summary["sales"] / state_summary["dealer_enquiry"]
).where(state_summary["dealer_enquiry"] > 0, 0)
state_summary["private_conversion_rate"] = (
    state_summary["private_sales"] / state_summary["private_enquiry"]
).where(state_summary["private_enquiry"] > 0, 0)
state_summary["fleet_conversion_rate"] = (
    state_summary["fleet_sales"] / state_summary["fleet_enquiry"]
).where(state_summary["fleet_enquiry"] > 0, 0)

with left:
    fig_state_enquiry = px.bar(
        state_summary,
        x="dealer_state",
        y="dealer_enquiry",
        title="Dealer Enquiry by State",
        labels={"dealer_state": "State", "dealer_enquiry": "Dealer Enquiry"},
    )
    st.plotly_chart(fig_state_enquiry, use_container_width=True)

with right:
    fig_state_sales = px.bar(
        state_summary.sort_values("sales", ascending=False),
        x="dealer_state",
        y="sales",
        title="Sales by State",
        labels={"dealer_state": "State", "sales": "Sales"},
    )
    st.plotly_chart(fig_state_sales, use_container_width=True)

st.subheader("State Summary - Filtered All Dealers")
st.caption("Scope: all dealers matching the Month and State filters, grouped by state.")
st.dataframe(
    format_rate_columns(
        state_summary.rename(
            columns={
                "dealer_state": "State",
                "dealer_enquiry": "Dealer Enquiry",
                "sales": "Sales",
                "private_enquiry": "Private Enquiry",
                "private_sales": "Private Sales",
                "private_conversion_rate": "Private Conversion Rate",
                "fleet_enquiry": "Fleet Enquiry",
                "fleet_sales": "Fleet Sales",
                "fleet_conversion_rate": "Fleet Conversion Rate",
                "dealers": "Dealers",
                "active_dealers": "Active Dealers",
                "inactive_dealers": "Inactive Dealers",
                "conversion_rate": "Conversion Rate",
            }
        ),
        ["Conversion Rate", "Private Conversion Rate", "Fleet Conversion Rate"],
    ),
    use_container_width=True,
    hide_index=True,
)

st.subheader("Dealer Performance - Filtered All Dealers")
st.caption("Scope: all dealers matching the Month and State filters, grouped by dealer.")

dealer_summary = (
    filtered.groupby(["dealer_name", "dealer_state", "active"], dropna=False)
    .agg(
        dealer_enquiry=("total_enquiry", "sum"),
        sales=("total_sales", "sum"),
        private_enquiry=("private_enquiry", "sum"),
        private_sales=("private_sales", "sum"),
        fleet_enquiry=("fleet_enquiry", "sum"),
        fleet_sales=("fleet_sales", "sum"),
        active_months=("month", "nunique"),
    )
    .reset_index()
    .sort_values(["sales", "dealer_enquiry"], ascending=False)
)
dealer_summary["conversion_rate"] = (
    dealer_summary["sales"] / dealer_summary["dealer_enquiry"]
).where(dealer_summary["dealer_enquiry"] > 0, 0)
dealer_summary["private_conversion_rate"] = (
    dealer_summary["private_sales"] / dealer_summary["private_enquiry"]
).where(dealer_summary["private_enquiry"] > 0, 0)
dealer_summary["fleet_conversion_rate"] = (
    dealer_summary["fleet_sales"] / dealer_summary["fleet_enquiry"]
).where(dealer_summary["fleet_enquiry"] > 0, 0)

fig_dealer = px.bar(
    dealer_summary.head(30),
    x="dealer_name",
    y=["dealer_enquiry", "sales"],
    barmode="group",
    title="Top Dealer Performance",
    labels={"dealer_name": "Dealer", "value": "Count", "variable": "Metric"},
)
fig_dealer.update_layout(xaxis_tickangle=-45)
st.plotly_chart(fig_dealer, use_container_width=True)

st.dataframe(
    format_rate_columns(
        dealer_summary.rename(
            columns={
                "dealer_name": "Dealer",
                "dealer_state": "State",
                "active": "Active",
                "dealer_enquiry": "Dealer Enquiry",
                "sales": "Sales",
                "private_enquiry": "Private Enquiry",
                "private_sales": "Private Sales",
                "private_conversion_rate": "Private Conversion Rate",
                "fleet_enquiry": "Fleet Enquiry",
                "fleet_sales": "Fleet Sales",
                "fleet_conversion_rate": "Fleet Conversion Rate",
                "active_months": "Active Months",
                "conversion_rate": "Conversion Rate",
            }
        ),
        ["Conversion Rate", "Private Conversion Rate", "Fleet Conversion Rate"],
    ),
    use_container_width=True,
    hide_index=True,
)

st.subheader("Dealer Drill-through - Selected Dealer Only")
st.caption("Scope: one selected dealer, using the current Month and State filters.")

dealer_for_view = st.selectbox(
    "Select a dealer",
    options=dealer_summary["dealer_name"].dropna().drop_duplicates().tolist(),
)

dealer_df = filtered[filtered["dealer_name"] == dealer_for_view].copy()
dealer_state = ", ".join(sorted(dealer_df["dealer_state"].dropna().unique()))
dealer_active = ", ".join(sorted(dealer_df["active"].dropna().unique()))
st.caption(f"State: {dealer_state} | Active status: {dealer_active}")

dealer_monthly = (
    dealer_df.groupby(["month", "month_order"], dropna=False)[
        [
            "total_enquiry",
            "total_sales",
            "private_enquiry",
            "private_sales",
            "fleet_enquiry",
            "fleet_sales",
        ]
    ]
    .sum()
    .reset_index()
    .sort_values("month_order")
)
dealer_monthly_melted = dealer_monthly.melt(
    id_vars=["month", "month_order"],
    value_vars=["total_enquiry", "total_sales"],
    var_name="metric",
    value_name="value",
)
dealer_monthly_melted["metric"] = dealer_monthly_melted["metric"].map(
    {"total_enquiry": "Dealer Enquiry", "total_sales": "Sales"}
)

fig_dealer_month = px.bar(
    dealer_monthly_melted,
    x="month",
    y="value",
    color="metric",
    barmode="group",
    title=f"Monthly Enquiry and Sales - {dealer_for_view}",
    labels={"month": "Month", "value": "Count", "metric": "Metric"},
)
st.plotly_chart(fig_dealer_month, use_container_width=True)

drill_left, drill_right = st.columns(2)

dealer_enquiry_breakdown = (
    dealer_df.groupby(["month", "month_order"], dropna=False)[ENQUIRY_COLS]
    .sum()
    .reset_index()
    .sort_values("month_order")
)
dealer_enquiry_melted = dealer_enquiry_breakdown.melt(
    id_vars=["month", "month_order"],
    value_vars=ENQUIRY_COLS,
    var_name="enquiry_type",
    value_name="enquiries",
)
dealer_enquiry_melted["enquiry_type"] = dealer_enquiry_melted["enquiry_type"].map(ENQUIRY_LABELS)

with drill_left:
    fig_dealer_enquiry = px.bar(
        dealer_enquiry_melted,
        x="month",
        y="enquiries",
        color="enquiry_type",
        title="Dealer Enquiry Breakdown",
        labels={"month": "Month", "enquiries": "Dealer Enquiry", "enquiry_type": "Type"},
    )
    st.plotly_chart(fig_dealer_enquiry, use_container_width=True)

dealer_sales_melted = make_sales_long(dealer_df, ["month", "month_order"], sales_cols)

with drill_right:
    if dealer_sales_melted.empty:
        st.info("No sales for this dealer in the selected filters.")
    else:
        dealer_sales_by_model = (
            dealer_sales_melted.groupby(["month", "month_order", "model"], dropna=False)["sales"]
            .sum()
            .reset_index()
            .sort_values("month_order")
        )
        fig_dealer_sales = px.bar(
            dealer_sales_by_model,
            x="month",
            y="sales",
            color="model",
            title="Dealer Sales Breakdown by Model",
            labels={"month": "Month", "sales": "Sales", "model": "Model"},
        )
        st.plotly_chart(fig_dealer_sales, use_container_width=True)

if not dealer_sales_melted.empty:
    st.markdown("#### Sales Type Detail - Selected Dealer Only")
    st.caption(f"Scope: {dealer_for_view} only.")
    sales_type_left, sales_type_right = st.columns(2)

    dealer_sales_by_type = (
        dealer_sales_melted.groupby(["month", "month_order", "sales_type"], dropna=False)["sales"]
        .sum()
        .reset_index()
        .sort_values("month_order")
    )
    with sales_type_left:
        fig_dealer_sales_type = px.bar(
            dealer_sales_by_type,
            x="month",
            y="sales",
            color="sales_type",
            title="Dealer Sales Breakdown by Sales Type",
            labels={"month": "Month", "sales": "Sales", "sales_type": "Sales Type"},
        )
        st.plotly_chart(fig_dealer_sales_type, use_container_width=True)

    dealer_sales_matrix = (
        dealer_sales_melted.groupby(["sales_type", "model"], dropna=False)["sales"]
        .sum()
        .reset_index()
    )
    with sales_type_right:
        fig_sales_matrix = px.density_heatmap(
            dealer_sales_matrix,
            x="model",
            y="sales_type",
            z="sales",
            histfunc="sum",
            title="Sales Mix: Model x Sales Type",
            labels={"model": "Model", "sales_type": "Sales Type", "sales": "Sales"},
            color_continuous_scale="Blues",
        )
        st.plotly_chart(fig_sales_matrix, use_container_width=True)

    sales_detail_table = (
        dealer_sales_melted.groupby(["month", "month_order", "model", "sales_group", "sales_type"], dropna=False)["sales"]
        .sum()
        .reset_index()
        .sort_values(["month_order", "model", "sales_group", "sales_type"])
    )
    st.markdown("##### Sales Detail Table - Selected Dealer Only")
    st.dataframe(
        sales_detail_table[["month", "model", "sales_group", "sales_type", "sales"]].rename(
            columns={
                "month": "Month",
                "model": "Model",
                "sales_group": "Sales Group",
                "sales_type": "Sales Type",
                "sales": "Sales",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

dealer_detail = dealer_monthly.copy()
dealer_detail["conversion_rate"] = (
    dealer_detail["total_sales"] / dealer_detail["total_enquiry"]
).where(dealer_detail["total_enquiry"] > 0, 0)
dealer_detail["private_conversion_rate"] = (
    dealer_detail["private_sales"] / dealer_detail["private_enquiry"]
).where(dealer_detail["private_enquiry"] > 0, 0)
dealer_detail["fleet_conversion_rate"] = (
    dealer_detail["fleet_sales"] / dealer_detail["fleet_enquiry"]
).where(dealer_detail["fleet_enquiry"] > 0, 0)
st.markdown("##### Monthly Summary Table - Selected Dealer Only")
st.dataframe(
    format_rate_columns(
        dealer_detail[
            [
                "month",
                "total_enquiry",
                "total_sales",
                "conversion_rate",
                "private_enquiry",
                "private_sales",
                "private_conversion_rate",
                "fleet_enquiry",
                "fleet_sales",
                "fleet_conversion_rate",
            ]
        ].rename(
            columns={
                "month": "Month",
                "total_enquiry": "Dealer Enquiry",
                "total_sales": "Sales",
                "conversion_rate": "Conversion Rate",
                "private_enquiry": "Private Enquiry",
                "private_sales": "Private Sales",
                "private_conversion_rate": "Private Conversion Rate",
                "fleet_enquiry": "Fleet Enquiry",
                "fleet_sales": "Fleet Sales",
                "fleet_conversion_rate": "Fleet Conversion Rate",
            }
        ),
        ["Conversion Rate", "Private Conversion Rate", "Fleet Conversion Rate"],
    ),
    use_container_width=True,
    hide_index=True,
)

with st.expander("View all filtered raw data"):
    st.caption(
        "Scope: all raw rows matching the Month and State filters. "
        "This table is not limited to the selected dealer above."
    )
    detail_cols = ["dealer_name", "dealer_state", "active", "dealer_code", "month", "total_enquiry", "total_sales"] + ENQUIRY_COLS + sales_cols
    detail_cols = [col for col in detail_cols if col in filtered.columns]
    raw_detail = filtered[detail_cols].sort_values(["dealer_state", "dealer_name", "month"])
    st.dataframe(raw_detail, use_container_width=True, hide_index=True)
