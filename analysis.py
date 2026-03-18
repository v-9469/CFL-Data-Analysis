import pandas as pd
import altair as alt
import streamlit as st
import re
from pathlib import Path

# --- CONFIGURATION ---
FILE_PATH = Path(r"D:\projects and files\CFL\CFL_External Data Pack_Phase1.xlsx")

st.set_page_config(page_title="Demand Forecasting Deep Dive", layout="wide")

# --- MASTER TIMELINE ORDER ---
MASTER_QUARTERS = [
    "FY23 Q1", "FY23 Q2", "FY23 Q3", "FY23 Q4",
    "FY24 Q1", "FY24 Q2", "FY24 Q3", "FY24 Q4",
    "FY25 Q1", "FY25 Q2", "FY25 Q3", "FY25 Q4",
    "FY26 Q1", "FY26 Q2"
]

# --- HELPER FUNCTIONS ---
def extract_and_format_quarter(col_name):
    """Intelligently pulls the quarter out of messy headers like 'VMS Units 2023Q1'"""
    col_str = str(col_name).strip()
    match = re.search(r'(20\d{2}Q[1-4])', col_str, re.IGNORECASE)
    if match:
        q_str = match.group(1).upper()
        return f"FY{q_str[2:4]} Q{q_str[-1]}"
    
    match_fy = re.search(r'(FY\d{2})\s*(Q[1-4])', col_str, re.IGNORECASE)
    if match_fy:
        return f"{match_fy.group(1).upper()} {match_fy.group(2).upper()}"
        
    return col_str

def clean_pct(val):
    if pd.isna(val) or val == 'nan': return None 
    if isinstance(val, str):
        return float(val.replace('%', '').strip()) / 100
    return float(val)

def normalize_string(series):
    return series.astype(str).str.upper().str.strip().str.replace(r'\s+', ' ', regex=True)

def clean_units(series):
    return pd.to_numeric(series.astype(str).str.replace(',', '', regex=False), errors="coerce").fillna(0)


# --- DATA LOADERS ---
@st.cache_data
def load_forecast_timeline():
    sheet_name = "Data Pack - Actual Bookings"
    
    raw_top = pd.read_excel(FILE_PATH, sheet_name=sheet_name, header=[0, 2])
    prod_col = [c for c in raw_top.columns if "Product Name" in str(c[0])][0]
    actual_cols = [c for c in raw_top.columns if str(c[0]).strip() == "ACTUAL UNITS"]
    
    df_actuals = raw_top[[prod_col] + actual_cols].copy()
    df_actuals.columns = ["Product"] + [c[1] for c in actual_cols]
    df_actuals = df_actuals.dropna(subset=["Product"])
    df_actuals["Product"] = normalize_string(df_actuals["Product"])
    
    df_actuals = df_actuals.melt(id_vars="Product", var_name="FYQTR", value_name="Actual_Units")
    df_actuals["Actual_Units"] = clean_units(df_actuals["Actual_Units"]) 
    df_actuals = df_actuals[df_actuals["Actual_Units"] >= 1].copy()
    
    raw_flat = pd.read_excel(FILE_PATH, sheet_name=sheet_name, header=None)
    mask = raw_flat.astype(str).apply(lambda x: x.str.contains("Forecast Accuracy", na=False))
    if not mask.any().any(): return df_actuals
        
    start_idx = mask.any(axis=1).idxmax()
    table2 = raw_flat.iloc[start_idx:].copy().reset_index(drop=True)
    table2.iloc[0] = table2.iloc[0].ffill() 
    table2.iloc[1] = table2.iloc[1].ffill()
    
    headers = [f"{str(table2.iloc[0, c]).strip()}|{str(table2.iloc[1, c]).strip()}|{str(table2.iloc[2, c]).strip()}" for c in range(table2.shape[1])]
    table2.columns = headers
    df2 = table2.drop([0, 1, 2]).reset_index(drop=True)
    
    prod_col_bot = [c for c in df2.columns if "Product Name" in c][0]
    df2 = df2.rename(columns={prod_col_bot: "Product"})
    df2 = df2.dropna(subset=["Product"])
    df2["Product"] = normalize_string(df2["Product"])
    
    teams = {"Demand": "Demand Planning", "Marketing": "Marketing Team", "Data Science": "Data Science Team"}
    merged = df_actuals.copy()
    
    for name, search in teams.items():
        team_cols = [c for c in df2.columns if search in c and "Accuracy" in c.split("|")[-1]]
        if not team_cols: continue
        
        df_team = df2[["Product"] + team_cols].melt(id_vars="Product", var_name="Raw_Col", value_name=f"{name}_Pct")
        df_team["FYQTR"] = df_team["Raw_Col"].apply(lambda x: x.split("|")[1].replace("FY20", "FY").strip())
        df_team[f"{name}_Pct"] = df_team[f"{name}_Pct"].apply(clean_pct)
        df_team["Product"] = normalize_string(df_team["Product"])
        
        merged = pd.merge(merged, df_team[["Product", "FYQTR", f"{name}_Pct"]], on=["Product", "FYQTR"], how="left")
        merged[f"{name} Forecast"] = merged["Actual_Units"] * merged[f"{name}_Pct"]

    return merged

@st.cache_data
def load_scms_data():
    raw = pd.read_excel(FILE_PATH, sheet_name="SCMS", header=None)
    
    plid_mask = raw.astype(str).apply(lambda x: x.str.contains("PLID", case=False, na=False))
    plid_idx = plid_mask.any(axis=1).idxmax() if plid_mask.any().any() else 0

    q_mask = raw.astype(str).apply(lambda x: x.str.contains(r"20\d{2}Q[1-4]|FY\d{2}\s*Q[1-4]", case=False, na=False))
    q_idx = q_mask.any(axis=1).idxmax() if q_mask.any().any() else plid_idx

    h1 = raw.iloc[plid_idx].astype(str).replace('nan', '')
    h2 = raw.iloc[q_idx].astype(str).replace('nan', '')
    raw.columns = [ (a + " " + b).strip() for a,b in zip(h1, h2) ]

    data_start = max(plid_idx, q_idx) + 1
    raw = raw.iloc[data_start:].reset_index(drop=True)

    prod_col = [c for c in raw.columns if "PLID" in str(c).upper() or "PRODUCT" in str(c).upper()][0]
    scms_col = [c for c in raw.columns if "COVERAGE" in str(c).upper() or "SCMS" in str(c).upper()][0]

    raw = raw.rename(columns={prod_col: "Product", scms_col: "Sales Coverage Code"})
    raw = raw.dropna(subset=["Product", "Sales Coverage Code"])
    raw["Product"] = normalize_string(raw["Product"])

    q_cols = [c for c in raw.columns if re.search(r"20\d{2}Q[1-4]|FY\d{2}\s*Q[1-4]", c, re.IGNORECASE)]
    df = raw.melt(id_vars=["Product", "Sales Coverage Code"], value_vars=q_cols, var_name="Raw_QTR", value_name="Units")
    
    df["FYQTR"] = df["Raw_QTR"].apply(extract_and_format_quarter)
    df["Units"] = clean_units(df["Units"]) 
    return df[df["Units"] > 0]

@st.cache_data
def load_vms_data():
    raw = pd.read_excel(FILE_PATH, sheet_name="VMS", header=None)
    
    plid_mask = raw.astype(str).apply(lambda x: x.str.contains("PLID", case=False, na=False))
    plid_idx = plid_mask.any(axis=1).idxmax() if plid_mask.any().any() else 0

    q_mask = raw.astype(str).apply(lambda x: x.str.contains(r"20\d{2}Q[1-4]|FY\d{2}\s*Q[1-4]", case=False, na=False))
    q_idx = q_mask.any(axis=1).idxmax() if q_mask.any().any() else plid_idx

    h1 = raw.iloc[plid_idx].astype(str).replace('nan', '')
    h2 = raw.iloc[q_idx].astype(str).replace('nan', '')
    raw.columns = [ (a + " " + b).strip() for a,b in zip(h1, h2) ]

    data_start = max(plid_idx, q_idx) + 1
    raw = raw.iloc[data_start:].reset_index(drop=True)

    prod_col = [c for c in raw.columns if "PLID" in str(c).upper() or "PRODUCT" in str(c).upper()][0]
    vms_col = [c for c in raw.columns if "VMS" in str(c).upper() or "VERTICAL" in str(c).upper() or "TOP NAME" in str(c).upper()][0]

    raw = raw.rename(columns={prod_col: "Product", vms_col: "Vertical"})
    raw = raw.dropna(subset=["Product", "Vertical"])
    raw["Product"] = normalize_string(raw["Product"])

    q_cols = [c for c in raw.columns if re.search(r"20\d{2}Q[1-4]|FY\d{2}\s*Q[1-4]", c, re.IGNORECASE)]
    df = raw.melt(id_vars=["Product", "Vertical"], value_vars=q_cols, var_name="Raw_QTR", value_name="Units")
    
    df["FYQTR"] = df["Raw_QTR"].apply(extract_and_format_quarter)
    df["Units"] = clean_units(df["Units"]) 
    return df[df["Units"] > 0]

@st.cache_data
def load_big_deals_data():
    raw = pd.read_excel(FILE_PATH, sheet_name="Big Deal", header=[0, 1])
    prod_col = [c for c in raw.columns if "PLID" in str(c[0]) or "PLID" in str(c[1])][0]
    
    bd_cols = [c for c in raw.columns if "Big Deals" in str(c[0])]
    df_bd = raw[[prod_col] + bd_cols].copy()
    df_bd.columns = ["Product"] + [str(c[1]).strip() for c in bd_cols]
    df_bd = df_bd.melt(id_vars="Product", var_name="Raw_QTR", value_name="Big Deal Units")
    
    avg_cols = [c for c in raw.columns if "Avg Deals" in str(c[0])]
    df_avg = raw[[prod_col] + avg_cols].copy()
    df_avg.columns = ["Product"] + [str(c[1]).strip() for c in avg_cols]
    df_avg = df_avg.melt(id_vars="Product", var_name="Raw_QTR", value_name="Run-Rate (Avg) Units")
    
    merged = pd.merge(df_bd, df_avg, on=["Product", "Raw_QTR"], how="outer")
    merged["Product"] = normalize_string(merged["Product"])
    merged["FYQTR"] = merged["Raw_QTR"].apply(extract_and_format_quarter)
    
    for c in ["Big Deal Units", "Run-Rate (Avg) Units"]:
        merged[c] = clean_units(merged[c])
        
    df = merged.melt(id_vars=["Product", "FYQTR"], value_vars=["Big Deal Units", "Run-Rate (Avg) Units"], var_name="Deal Type", value_name="Units")
    return df[df["Units"] > 0]


# --- MAIN APP UI ---
st.title("🚀 Demand Forecasting Deep Dive")

if not FILE_PATH.exists():
    st.error(f"Excel file not found at: {FILE_PATH}")
    st.stop()

# Load Data
df_main = load_forecast_timeline()
df_scms = load_scms_data()
df_vms = load_vms_data()
df_bd = load_big_deals_data()

# Product Selection
products = sorted(df_main["Product"].unique())
selected_product = st.selectbox("Select a Product to Analyze", options=products)

st.divider()

# --- SECTION 1: MASTER TIMELINE ---
st.header("1. Forecast Accuracy Timeline")
st.markdown("Use this to determine which team's modeling is historically the most reliable for this specific product.")

prod_main = df_main[df_main["Product"] == selected_product].copy()
metrics = ["Actual_Units", "Demand Forecast", "Marketing Forecast", "Data Science Forecast"]
available_metrics = [m for m in metrics if m in prod_main.columns]

plot_main = prod_main.melt(id_vars=["FYQTR"], value_vars=available_metrics, var_name="Metric", value_name="Units").dropna()

color_scale = alt.Scale(domain=metrics, range=["#1f77b4", "#2ca02c", "#ff7f0e", "#9467bd"])
chart_main = (
    alt.Chart(plot_main)
    .mark_line(point=True, strokeWidth=3)
    .encode(
        x=alt.X("FYQTR:N", title="Fiscal Quarter", sort=MASTER_QUARTERS),
        y=alt.Y("Units:Q", title="Volume (Units)"),
        color=alt.Color("Metric:N", scale=color_scale),
        tooltip=["FYQTR", "Metric", alt.Tooltip("Units:Q", format=",.0f")]
    ).properties(height=400).interactive()
)
st.altair_chart(chart_main, use_container_width=True)

st.divider()

# --- SECTION 2: DEMAND DRIVERS (SCMS & VMS) ---
st.header("2. Demand Drivers (Segments & Trends)")
st.markdown("Analyze how different customer segments and industries are driving your sales volume over time. **Hover over the top right corner of any chart to view it in Full Screen!**")

# Filter data for selected product
prod_scms = df_scms[df_scms["Product"] == selected_product]
prod_vms = df_vms[df_vms["Product"] == selected_product]

# Pre-calculate Total Actual Units for baselines in both charts
df_total = df_main[df_main["Product"] == selected_product][["FYQTR", "Actual_Units"]].dropna()
df_total = df_total.rename(columns={"Actual_Units": "Units"})
df_total["Category"] = "Total Actual Units"

# --- SCMS VISUALIZATIONS (INSIDE EXPANDER) ---
with st.expander("📊 Deep Dive: Sales Coverage (SCMS)", expanded=True):
    if not prod_scms.empty:
        # SCMS Bar Chart
        st.subheader("Percentage Contribution")
        chart_scms = alt.Chart(prod_scms).mark_bar().encode(
            x=alt.X("FYQTR:N", title="Quarter", sort=MASTER_QUARTERS),
            y=alt.Y("sum(Units):Q", stack="normalize", title="SCMS Contribution (%)", axis=alt.Axis(format='%')),
            color=alt.Color("Sales Coverage Code:N", title="SCMS Segment"),
            tooltip=["FYQTR", "Sales Coverage Code", alt.Tooltip("sum(Units):Q", title="Raw Units", format=",.0f")]
        ).properties(height=400).interactive()
        st.altair_chart(chart_scms, use_container_width=True)
        
        st.divider()
        
        # SCMS Line Chart
        st.subheader("Coverage Trend (Raw Volume Drivers)")
        available_scms = sorted(prod_scms["Sales Coverage Code"].unique())
        top_scms = prod_scms.groupby("Sales Coverage Code")["Units"].sum().nlargest(3).index.tolist()
        
        selected_scms = st.multiselect(
            "Select SCMS Segments to Compare against Total Sales:",
            options=available_scms,
            default=top_scms,
            key="scms_multiselect"
        )

        if selected_scms:
            df_scms_sub = prod_scms[prod_scms["Sales Coverage Code"].isin(selected_scms)][["FYQTR", "Units", "Sales Coverage Code"]]
            df_scms_sub = df_scms_sub.rename(columns={"Sales Coverage Code": "Category"})
            
            plot_trend_scms = pd.concat([df_total, df_scms_sub], ignore_index=True)

            chart_scms_line = alt.Chart(plot_trend_scms).mark_line(point=True, strokeWidth=3).encode(
                x=alt.X("FYQTR:N", title="Quarter", sort=MASTER_QUARTERS),
                y=alt.Y("Units:Q", title="Volume (Units Sold)"),
                color=alt.Color("Category:N", title="Metric / Segment", scale=alt.Scale(scheme='category10')),
                strokeDash=alt.condition(
                    alt.datum.Category == 'Total Actual Units',
                    alt.value([5, 5]),  
                    alt.value([0])      
                ),
                tooltip=["FYQTR", "Category", alt.Tooltip("Units:Q", title="Units Sold", format=",.0f")]
            ).properties(height=450).interactive()
            
            st.altair_chart(chart_scms_line, use_container_width=True)
        else:
            st.info("Please select at least one SCMS segment to view the trend.")
    else:
        st.info("No SCMS data available for this product.")

st.write("") # Quick spacing between expanders

# --- VMS VISUALIZATIONS (INSIDE EXPANDER) ---
with st.expander("🏢 Deep Dive: Industry Vertical (VMS)", expanded=True):
    if not prod_vms.empty:
        # VMS Bar Chart
        st.subheader("Percentage Contribution")
        chart_vms_pct = alt.Chart(prod_vms).mark_bar().encode(
            x=alt.X("FYQTR:N", title="Quarter", sort=MASTER_QUARTERS),
            y=alt.Y("sum(Units):Q", stack="normalize", title="Sector Contribution (%)", axis=alt.Axis(format='%')),
            color=alt.Color("Vertical:N", title="Industry", scale=alt.Scale(scheme='category20')),
            tooltip=["FYQTR", "Vertical", alt.Tooltip("sum(Units):Q", title="Raw Units", format=",.0f")]
        ).properties(height=400).interactive()
        st.altair_chart(chart_vms_pct, use_container_width=True)
        
        st.divider()
        
        # VMS Line Chart
        st.subheader("Vertical Trend (Raw Volume Drivers)")
        available_sectors = sorted(prod_vms["Vertical"].unique())
        top_sectors = prod_vms.groupby("Vertical")["Units"].sum().nlargest(3).index.tolist()
        
        selected_sectors = st.multiselect(
            "Select Industry Sectors to Compare against Total Sales:",
            options=available_sectors,
            default=top_sectors,
            key="vms_multiselect"
        )

        if selected_sectors:
            df_vms_sub = prod_vms[prod_vms["Vertical"].isin(selected_sectors)][["FYQTR", "Units", "Vertical"]]
            df_vms_sub = df_vms_sub.rename(columns={"Vertical": "Category"})
            
            plot_trend_vms = pd.concat([df_total, df_vms_sub], ignore_index=True)

            chart_vms_line = alt.Chart(plot_trend_vms).mark_line(point=True, strokeWidth=3).encode(
                x=alt.X("FYQTR:N", title="Quarter", sort=MASTER_QUARTERS),
                y=alt.Y("Units:Q", title="Volume (Units Sold)"),
                color=alt.Color("Category:N", title="Metric / Industry", scale=alt.Scale(scheme='category20')),
                strokeDash=alt.condition(
                    alt.datum.Category == 'Total Actual Units',
                    alt.value([5, 5]),  
                    alt.value([0])      
                ),
                tooltip=["FYQTR", "Category", alt.Tooltip("Units:Q", title="Units Sold", format=",.0f")]
            ).properties(height=450).interactive()
            
            st.altair_chart(chart_vms_line, use_container_width=True)
        else:
            st.info("Please select at least one sector to view the trend.")
    else:
        st.info("No VMS data available for this product.")

st.divider()

# --- SECTION 3: RUN-RATE VS ANOMALIES ---
st.header("3. Run-Rate vs. Big Deals")
st.markdown("Strip away the 'Big Deals' to see the true underlying run-rate of the business. Base next quarter's baseline on the **Run-Rate Trendline**, not the total volume.")
st.caption("🟡 **How to read the Yellow Line:** This line tracks your true baseline (Run-Rate). If total units spike but the yellow line stays flat, the spike was an anomaly. If the yellow line grows, your core business is truly expanding!")

prod_bd = df_bd[df_bd["Product"] == selected_product]

if not prod_bd.empty:
    bars = alt.Chart(prod_bd).mark_bar(size=40).encode(
        x=alt.X("FYQTR:N", title="Quarter", sort=MASTER_QUARTERS),
        y=alt.Y("sum(Units):Q", title="Total Units Sold", stack=True),
        color=alt.Color("Deal Type:N", scale=alt.Scale(domain=["Run-Rate (Avg) Units", "Big Deal Units"], range=["#2ca02c", "#d62728"])),
        order=alt.Order("Deal Type:N", sort="descending"), 
        tooltip=["FYQTR", "Deal Type", alt.Tooltip("sum(Units):Q", title="Units", format=",.0f")]
    )

    df_run_rate = prod_bd[prod_bd["Deal Type"] == "Run-Rate (Avg) Units"]
    line = alt.Chart(df_run_rate).mark_line(
        color="#ffcc00", 
        strokeWidth=4, 
        point=alt.OverlayMarkDef(color="#ffcc00", size=100) 
    ).encode(
        x=alt.X("FYQTR:N", sort=MASTER_QUARTERS),
        y=alt.Y("sum(Units):Q"),
        tooltip=["FYQTR", alt.Tooltip("sum(Units):Q", title="Baseline Run-Rate", format=",.0f")]
    )

    chart_bd = (bars + line).properties(height=450).interactive()
    st.altair_chart(chart_bd, use_container_width=True)
else:
    st.info("No Big Deal data available for this product.")