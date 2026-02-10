# =========================================================
# IMPORTS
# =========================================================
import streamlit as st
import pandas as pd
import plotly.express as px
import io

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
# LOAD DATA (OLD LOGIC)
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
# DRIVER UNITS (TOTAL ROW)
# =========================================================
total_row = event_df.iloc[-1]
numeric_cols = event_df.select_dtypes(include="number").columns

driver_units = {
    str(k).strip(): float(v)
    for k, v in total_row[numeric_cols].to_dict().items()
    if pd.notna(v)
}

def get_driver_units(driver):
    if pd.isna(driver):
        return 0
    return driver_units.get(str(driver).strip(), 0)

cost_df["Driver_Units"] = cost_df["Driver"].apply(get_driver_units)
cost_df["RatePerDriverUnit"] = cost_df.apply(
    lambda r: r["Total_Cost"] / r["Driver_Units"] if r["Driver_Units"] != 0 else 0,
    axis=1
)

# =========================================================
# REMOVE TOTAL ROW FROM EVENTS
# =========================================================
events_clean = event_df.iloc[:-1].copy()

# =========================================================
# COST ALLOCATION
# =========================================================
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

# =========================================================
# CREATE TIME PERIOD FROM DEPARTURE TIME
# =========================================================
viz_df["Departure Time"] = pd.to_datetime(
    viz_df["Departure Time"],
    errors="coerce"
)

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
# ================= NEW VISUAL SECTION ====================
# =========================================================
st.markdown("---")
st.header("ABC Cost Analysis Dashboard")

# ---------------- KPI ----------------
k1, k2, k3, k4 = st.columns(4)

k1.metric("Total Flights", len(filtered_summary))
k2.metric(
    "Total Cost",
    f"{filtered_summary['Total_Cost_Per_Flight'].sum():,.0f}"
)
k3.metric(
    "Average Cost / Flight",
    f"{filtered_summary['Total_Cost_Per_Flight'].mean():,.0f}"
)
k4.metric(
    "Highest Cost Flight",
    f"{filtered_summary['Total_Cost_Per_Flight'].max():,.0f}"
)

# ---------------- MAIN BAR ----------------
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

melted = melted[
    melted["Cost Type"] != "Total_Cost_Per_Flight"
]

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

st.markdown("### Cost by Time Period")

time_df = filtered_viz.groupby(
    "Time Period"
)["Total_Cost_Per_Flight"].sum().reset_index()

st.plotly_chart(
    px.bar(
        time_df,
        x="Time Period",
        y="Total_Cost_Per_Flight",
        title="Total Cost by Time Period"
    ),
    use_container_width=True
)

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
        y="Total_Cost_Per_Flight",
        title="Top 10 Destination Codes by Total Cost"
    ),
    use_container_width=True
)


# ---------------- SUPPORTING VISUALS ----------------
c1, c2, c3 = st.columns(3)

with c1:
    st.subheader("Cost Trend")
    st.plotly_chart(
        px.line(
            filtered_summary,
            x="Flight",
            y="Total_Cost_Per_Flight",
            markers=True
        ),
        use_container_width=True
    )

with c2:
    st.subheader("Cost by Continent")
    cont_df = filtered_viz.groupby(
        "Continent"
    )["Total_Cost_Per_Flight"].sum().reset_index()

    st.plotly_chart(
        px.bar(
            cont_df,
            x="Continent",
            y="Total_Cost_Per_Flight"
        ),
        use_container_width=True
    )

with c3:
    st.subheader("Cost Breakdown (Top Flight)")
    top_flight = top5.iloc[0]

    pie_df = top_flight.drop(
        ["Flight", "Total_Cost_Per_Flight"]
    ).reset_index()

    pie_df.columns = ["Cost Type", "Cost"]

    st.plotly_chart(
        px.pie(
            pie_df,
            values="Cost",
            names="Cost Type"
        ),
        use_container_width=True
    )



# =========================================================
# EXPORT
# =========================================================
st.markdown("---")
output = io.BytesIO()

with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    cost_df.to_excel(writer, index=False, sheet_name="Costpools")
    event_df.to_excel(writer, index=False, sheet_name="Events")
    costs_df.to_excel(writer, index=False, sheet_name="Cost_Allocation")
    summary.to_excel(writer, index=False, sheet_name="Summary")

st.download_button(
    "Download Final ABC Report",
    data=output.getvalue(),
    file_name="ABC_Final_Report.xlsx"
)

# =========================================================
# EXPORT CSV (ZIP)
# =========================================================
import zipfile

csv_buffer = io.BytesIO()

with zipfile.ZipFile(csv_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
    zipf.writestr(
        "Costpools.csv",
        cost_df.to_csv(index=False)
    )
    zipf.writestr(
        "Events.csv",
        event_df.to_csv(index=False)
    )
    zipf.writestr(
        "Cost_Allocation.csv",
        costs_df.to_csv(index=False)
    )
    zipf.writestr(
        "Summary.csv",
        summary.to_csv(index=False)
    )

st.download_button(
    "Download ABC Data (CSV)",
    data=csv_buffer.getvalue(),
    file_name="ABC_Data_CSV.zip",
    mime="application/zip"
)

