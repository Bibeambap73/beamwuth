# =========================================================
# IMPORTS
# =========================================================
import streamlit as st
import pandas as pd
import plotly.express as px
import io
import zipfile

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="ABC Costing Dashboard",
    layout="wide"
)

st.title("Activity-Based Costing (ABC) Dashboard")

# =========================================================
# FILE UPLOAD
# =========================================================
uploaded_file = st.file_uploader(
    "Upload Excel File (.xlsx)",
    type=["xlsx"]
)

if uploaded_file is None:
    st.stop()

# =========================================================
# LOAD DATA
# =========================================================
sheets = pd.read_excel(uploaded_file, sheet_name=None)
sheet_names = list(sheets.keys())

cost_sheet = next((s for s in sheet_names if "cost" in s.lower()), sheet_names[0])
event_sheet = next((s for s in sheet_names if "event" in s.lower()), sheet_names[1])

cost_df = sheets[cost_sheet].copy()
event_df = sheets[event_sheet].copy()

# =========================================================
# VALIDATION
# =========================================================
required_cols = {"Activity", "Type", "Total_Cost", "Driver"}
if not required_cols.issubset(cost_df.columns):
    st.error("Costpools sheet ต้องมี Activity, Type, Total_Cost, Driver")
    st.stop()

# =========================================================
# AUTO REMOVE TOTAL ROW IF EXISTS
# =========================================================
first_col = event_df.columns[0]
last_row_value = str(event_df.iloc[-1][first_col]).lower()

if "total" in last_row_value:
    event_df = event_df.iloc[:-1]

# =========================================================
# AUTO CALCULATE DRIVER UNITS
# =========================================================
driver_data = event_df.iloc[:, 1:]
numeric_cols = driver_data.select_dtypes(include="number").columns

driver_units = driver_data[numeric_cols].sum().to_dict()

def get_driver_units(driver):
    if pd.isna(driver):
        return 0
    return driver_units.get(str(driver).strip(), 0)

cost_df["Driver_Units"] = cost_df["Driver"].apply(get_driver_units)

cost_df["RatePerDriverUnit"] = cost_df.apply(
    lambda r: r["Total_Cost"] / r["Driver_Units"]
    if r["Driver_Units"] != 0 else 0,
    axis=1
)

# =========================================================
# COST ALLOCATION
# =========================================================
events_clean = event_df.copy()
costs_data = []

for _, flight_row in events_clean.iterrows():
    flight_name = flight_row.iloc[0]
    row_cost = {"Flight": str(flight_name)}

    for _, act in cost_df.iterrows():
        row_cost[act["Activity"]] = (
            flight_row.get(act["Driver"], 0) * act["RatePerDriverUnit"]
        )

    costs_data.append(row_cost)

costs_df = pd.DataFrame(costs_data)
costs_df["Total_Cost_Per_Flight"] = costs_df.drop(
    columns=["Flight"]
).sum(axis=1)

# =========================================================
# SUMMARY BY TYPE
# =========================================================
summary = pd.DataFrame()
summary["Flight"] = costs_df["Flight"]

for t in cost_df["Type"].unique():
    acts = cost_df.loc[cost_df["Type"] == t, "Activity"]
    summary[f"{t}_Cost"] = costs_df[acts].sum(axis=1)

summary["Total_Cost_Per_Flight"] = costs_df["Total_Cost_Per_Flight"]

# =========================================================
# PREP DATA FOR VISUAL
# =========================================================
viz_df = events_clean.copy()
viz_df["Flight"] = viz_df.iloc[:, 0].astype(str)

if "Departure Time" in viz_df.columns:
    viz_df["Departure Time"] = pd.to_datetime(
        viz_df["Departure Time"],
        errors="coerce"
    )
else:
    viz_df["Departure Time"] = pd.NaT

def get_time_period(t):
    if pd.isna(t):
        return "Unknown"
    hour = t.hour
    if 5 <= hour < 12:
        return "Morning"
    elif 12 <= hour < 17:
        return "Afternoon"
    elif 17 <= hour < 21:
        return "Evening"
    else:
        return "Night"

viz_df["Time Period"] = viz_df["Departure Time"].apply(get_time_period)

viz_df = viz_df.merge(
    summary[["Flight", "Total_Cost_Per_Flight"]],
    on="Flight",
    how="left"
)

# =========================================================
# SIDEBAR FILTERS
# =========================================================
st.sidebar.header("Filters")

selected_continent = st.sidebar.selectbox(
    "Continent",
    ["All"] + sorted(viz_df["Continent"].dropna().unique())
)

selected_dest = st.sidebar.selectbox(
    "Destination Code",
    ["All"] + sorted(viz_df["Destination Code"].dropna().unique())
)

selected_period = st.sidebar.selectbox(
    "Time Period",
    ["All", "Morning", "Afternoon", "Evening", "Night", "Unknown"]
)

filtered_viz = viz_df.copy()

if selected_continent != "All":
    filtered_viz = filtered_viz[
        filtered_viz["Continent"] == selected_continent
    ]

if selected_dest != "All":
    filtered_viz = filtered_viz[
        filtered_viz["Destination Code"] == selected_dest
    ]

if selected_period != "All":
    filtered_viz = filtered_viz[
        filtered_viz["Time Period"] == selected_period
    ]

filtered_summary = summary[
    summary["Flight"].isin(filtered_viz["Flight"])
]

# =========================================================
# TABLE SECTION
# =========================================================
st.markdown("---")
st.header("ABC Calculation Tables")

with st.expander("Costpools & Driver Rates"):
    st.dataframe(cost_df.round(3), use_container_width=True)

with st.expander("Cost Allocation (Flight × Activity)"):
    st.dataframe(costs_df.round(2), use_container_width=True)

with st.expander("Cost Summary by Type"):
    st.dataframe(summary.round(2), use_container_width=True)

# =========================================================
# ================= DASHBOARD SECTION =====================
# =========================================================
st.markdown("---")
st.header("ABC Cost Analysis Dashboard")

# KPI
k1, k2, k3, k4 = st.columns(4)

total_cost = filtered_summary["Total_Cost_Per_Flight"].sum()
avg_cost = filtered_summary["Total_Cost_Per_Flight"].mean()

k1.metric("Total Flights", len(filtered_summary))
k2.metric("Total Cost", f"{total_cost:,.0f}")
k3.metric("Average Cost / Flight", f"{avg_cost:,.0f}")

if len(filtered_summary) > 0:
    top_row = filtered_summary.sort_values(
        "Total_Cost_Per_Flight",
        ascending=False
    ).iloc[0]

    k4.metric(
        "Highest Cost Flight",
        f"{top_row['Total_Cost_Per_Flight']:,.0f}",
        delta=f"Flight: {top_row['Flight']}"
    )

# =========================================================
# EXECUTIVE GRAPHS
# =========================================================
st.markdown("### Executive Overview")

colA, colB, colC = st.columns(3)

# ---------------------------------------------------------
# 1️⃣ Cost Trend
# ---------------------------------------------------------
# ---------------------------------------------------------
# 1️⃣ Cost Trend
# ---------------------------------------------------------
with colA:
    st.markdown("#### Cost Trend")

    if filtered_viz["Departure Time"].notna().any():
        trend_df = filtered_viz.sort_values("Departure Time")
        x_axis = "Departure Time"
    else:
        trend_df = filtered_summary.sort_values("Flight")
        x_axis = "Flight"

    fig_trend = px.line(
        trend_df,
        x=x_axis,
        y="Total_Cost_Per_Flight",
        markers=True
    )

    st.plotly_chart(fig_trend, use_container_width=True)

# ---------------------------------------------------------
# 2️⃣ Cost by Continent
# ---------------------------------------------------------
with colB:
    st.markdown("#### Cost by Continent")

    continent_df = filtered_viz.groupby(
        "Continent"
    )["Total_Cost_Per_Flight"].sum().reset_index()

    fig_continent = px.bar(
        continent_df,
        x="Continent",
        y="Total_Cost_Per_Flight"
    )

    st.plotly_chart(fig_continent, use_container_width=True)

# ---------------------------------------------------------
# 3️⃣ Cost Breakdown (Top Flight)
# ---------------------------------------------------------
with colC:
    st.markdown("#### Cost Breakdown (Top Flight)")

    if len(filtered_summary) > 0:

        top_flight_name = filtered_summary.sort_values(
            "Total_Cost_Per_Flight",
            ascending=False
        ).iloc[0]["Flight"]

        top_flight_data = summary[
            summary["Flight"] == top_flight_name
        ]

        pie_df = top_flight_data.melt(
            id_vars=["Flight"],
            var_name="Cost Type",
            value_name="Cost"
        )

        pie_df = pie_df[
            pie_df["Cost Type"] != "Total_Cost_Per_Flight"
        ]

        fig_pie = px.pie(
            pie_df,
            names="Cost Type",
            values="Cost"
        )

        st.plotly_chart(fig_pie, use_container_width=True)

# Top 5 Flights
st.subheader("Top 5 Flights by Total Cost")

top5 = filtered_summary.sort_values(
    "Total_Cost_Per_Flight",
    ascending=False
).head(5)

melted = top5.melt(
    id_vars=["Flight"],
    var_name="Cost Type",
    value_name="Cost"
)

melted = melted[melted["Cost Type"] != "Total_Cost_Per_Flight"]

st.plotly_chart(
    px.bar(
        melted,
        x="Flight",
        y="Cost",
        color="Cost Type",
        barmode="stack"
    ),
    use_container_width=True
)

# =========================================================
# Cost by Activity (All Flights)
# =========================================================
#st.markdown("### Cost by Activity (All Flights)")

activity_totals = costs_df.drop(
    columns=["Flight", "Total_Cost_Per_Flight"]
).sum().reset_index()

activity_totals.columns = ["Activity", "Total Cost"]

activity_totals = activity_totals.sort_values(
    "Total Cost",
    ascending=False
)

#st.plotly_chart(
    #px.bar(
        #activity_totals,
        #x="Activity",
        #y="Total Cost"
    #),
    #use_container_width=True
#)

st.markdown("### Pareto Analysis (Activity Cost 80/20)")

pareto_df = activity_totals.copy()
pareto_df["Cumulative %"] = (
    pareto_df["Total Cost"].cumsum() /
    pareto_df["Total Cost"].sum()
) * 100

fig_pareto = px.bar(
    pareto_df,
    x="Activity",
    y="Total Cost"
)

fig_pareto.add_scatter(
    x=pareto_df["Activity"],
    y=pareto_df["Cumulative %"],
    mode="lines+markers",
    name="Cumulative %",
    yaxis="y2"
)

fig_pareto.update_layout(
    yaxis2=dict(
        overlaying="y",
        side="right",
        title="Cumulative %",
        range=[0, 100]
    )
)

st.plotly_chart(fig_pareto, use_container_width=True)



# Cost by Time Period
st.markdown("### Cost by Time Period")

time_df = filtered_viz.groupby(
    "Time Period"
)["Total_Cost_Per_Flight"].sum().reset_index()
order = ["Morning", "Afternoon", "Evening", "Night", "Unknown"]

time_df["Time Period"] = pd.Categorical(
    time_df["Time Period"],
    categories=order,
    ordered=True
)

time_df = time_df.sort_values("Time Period")

st.plotly_chart(
    px.bar(
        time_df,
        x="Time Period",
        y="Total_Cost_Per_Flight"
    ),
    use_container_width=True
)

# Cost by Destination
st.markdown("### Cost by Destination Code")

dest_df = filtered_viz.groupby(
    "Destination Code"
)["Total_Cost_Per_Flight"].sum().reset_index()

dest_df = dest_df.sort_values(
    "Total_Cost_Per_Flight",
    ascending=False
).head(10)

st.plotly_chart(
    px.bar(
        dest_df,
        x="Destination Code",
        y="Total_Cost_Per_Flight"
    ),
    use_container_width=True
)

# =========================================================
# EXPORT
# =========================================================
st.markdown("---")
st.header("Export Reports")

output = io.BytesIO()

with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    cost_df.to_excel(writer, index=False, sheet_name="Costpools")
    event_df.to_excel(writer, index=False, sheet_name="Events")
    costs_df.to_excel(writer, index=False, sheet_name="Cost_Allocation")
    summary.to_excel(writer, index=False, sheet_name="Summary")

st.download_button(
    "Download Final ABC Report (Excel)",
    data=output.getvalue(),
    file_name="ABC_Final_Report.xlsx"
)

# =========================================================
# EXPORT CSV (ZIP)
# =========================================================
csv_buffer = io.BytesIO()

with zipfile.ZipFile(csv_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
    zipf.writestr("Costpools.csv", cost_df.to_csv(index=False))
    zipf.writestr("Events.csv", event_df.to_csv(index=False))
    zipf.writestr("Cost_Allocation.csv", costs_df.to_csv(index=False))
    zipf.writestr("Summary.csv", summary.to_csv(index=False))

st.download_button(
    "Download ABC Data (CSV ZIP)",
    data=csv_buffer.getvalue(),
    file_name="ABC_Data_CSV.zip",
    mime="application/zip"
)