"""
Well Dashboard - ONGC Drilling Fluid Services
Streamlit application for monitoring well drilling operations.
Dynamically reads data from all_wells_data.json (117 wells).
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import os
from datetime import datetime, date

from data_loader import (
    build_wells_dataframe, build_phases_dataframe,
    build_complications_dataframe, build_npt_summary_dataframe,
    build_chemicals_dataframe, build_cost_analysis_dataframe,
    get_chemical_totals, scan_for_new_wells, load_wells_json,
    MUD_TYPE_COLORS,
)

WELL_CARDS_DIR = os.environ.get("WELL_CARDS_DIR", os.path.join(os.path.dirname(__file__), "well_cards"))

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Well Dashboard",
    page_icon="\U0001f6e2\ufe0f",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .main-header {
        background: linear-gradient(135deg, #1a237e 0%, #283593 50%, #3949ab 100%);
        padding: 20px 30px; border-radius: 12px; margin-bottom: 20px;
        color: white; display: flex; align-items: center; gap: 20px;
    }
    .main-header h1 { margin: 0; font-size: 2.2rem; font-weight: 700; letter-spacing: 1px; }
    .main-header p { margin: 4px 0 0 0; font-size: 0.95rem; opacity: 0.85; }
    .section-header {
        background: linear-gradient(90deg, #1a237e, #3949ab);
        color: white; padding: 10px 16px; border-radius: 8px;
        font-weight: 600; font-size: 1rem; margin: 16px 0 10px 0;
    }
    div[data-testid="stDataFrame"] table { font-size: 12px; }
    .stTabs [data-baseweb="tab"] { font-weight: 500; font-size: 14px; }
    .stTabs [aria-selected="true"] { color: #1a237e !important; }
</style>
""", unsafe_allow_html=True)

# ── Load data ─────────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data():
    return {
        "wells": build_wells_dataframe(),
        "phases": build_phases_dataframe(),
        "mud_loss": build_complications_dataframe("mud_loss"),
        "well_activity": build_complications_dataframe("well_activity"),
        "stuck_up": build_complications_dataframe("stuck_up"),
        "npt": build_npt_summary_dataframe(),
        "chemicals": build_chemicals_dataframe(),
        "cost": build_cost_analysis_dataframe(),
        "chem_totals": get_chemical_totals(),
    }

data = load_data()
wells_df = data["wells"]
phases_df = data["phases"]
npt_df = data["npt"]

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <div>
        <h1>\U0001f6e2\ufe0f Well Dashboard</h1>
        <p>ONGC \u00b7 Drilling Fluid Services \u00b7 All Assets \u00b7 117 Wells Loaded</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Sidebar Filters ──────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### \u2699\ufe0f Filters")

    all_wells = sorted(wells_df["Well Name"].tolist())
    sel_wells = st.multiselect("Well Name", all_wells, default=[], key="filter_wells",
                                placeholder="All Wells")

    all_assets = sorted(wells_df["Asset"].unique().tolist())
    sel_assets = st.multiselect("Asset / Area", all_assets, default=[], key="filter_assets",
                                 placeholder="All Assets")

    all_phases = sorted(phases_df["Phase"].unique().tolist())
    sel_phases = st.multiselect("Phase", all_phases, default=[], key="filter_phases",
                                 placeholder="All Phases")

    all_mud_types = sorted(phases_df["Mud Type"].unique().tolist())
    sel_mud_types = st.multiselect("Mud Type", all_mud_types, default=[], key="filter_mud",
                                    placeholder="All Mud Types")

    st.markdown("---")
    st.markdown("**Date Range**")
    date_from = st.date_input("From", value=date(2024, 1, 1), key="date_from")
    date_to = st.date_input("To", value=date.today(), key="date_to")

    st.markdown("---")
    st.markdown("**Complication Filters**")
    loss_types = sorted(set(
        str(r) for r in data["mud_loss"]["Type of Loss/Stuck Up"].unique()
        if r and str(r) not in ('', 'nan', '0')
    )) if not data["mud_loss"].empty else []
    sel_loss = st.multiselect("Loss Type", loss_types, default=[], key="filter_loss",
                               placeholder="All Types")

    stuck_types = sorted(set(
        str(r) for r in data["stuck_up"]["Type of Loss/Stuck Up"].unique()
        if r and str(r) not in ('', 'nan', '0')
    )) if not data["stuck_up"].empty else []
    sel_stuck = st.multiselect("Stuck Up Type", stuck_types, default=[], key="filter_stuck",
                                placeholder="All Types")

    st.markdown("---")
    if st.button("\U0001f504 Refresh Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# ── Apply filters ────────────────────────────────────────────────────────────
def apply_filters(df, well_col="Well Name", asset_col="Asset"):
    mask = pd.Series([True] * len(df), index=df.index)
    if sel_wells:
        mask &= df[well_col].isin(sel_wells)
    if sel_assets and asset_col in df.columns:
        mask &= df[asset_col].isin(sel_assets)
    return df[mask]

filtered_wells = apply_filters(wells_df)
filtered_phases = apply_filters(phases_df)
if sel_phases:
    filtered_phases = filtered_phases[filtered_phases["Phase"].isin(sel_phases)]
if sel_mud_types:
    filtered_phases = filtered_phases[filtered_phases["Mud Type"].isin(sel_mud_types)]
filtered_npt = apply_filters(npt_df)
filtered_chemicals = apply_filters(data["chemicals"], well_col="well_name", asset_col="asset")

# ── KPI Metrics ──────────────────────────────────────────────────────────────
col1, col2, col3, col4, col5, col6 = st.columns(6)
with col1:
    st.metric("Total Wells", len(filtered_wells))
with col2:
    st.metric("Total Phases", len(filtered_phases))
with col3:
    st.metric("Total Meterage (m)", f"{filtered_wells['Meterage (m)'].sum():,.0f}")
with col4:
    total_cost = filtered_wells["Total Cost (INR)"].sum()
    st.metric("Total Cost (INR Cr)", f"{total_cost/1e7:.2f}")
with col5:
    total_mud = filtered_wells["Total Mud Handled (bbl)"].sum()
    st.metric("Total Mud (bbl)", f"{total_mud:,.0f}")
with col6:
    total_npt = filtered_wells["Total NPT (Hrs)"].sum()
    st.metric("Total NPT (Hrs)", f"{total_npt:.0f}")

st.markdown("---")

# ── TABS ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "\U0001f4ca Overview", "\u23f1\ufe0f NPT Details", "\U0001f5fa\ufe0f Well Coverage",
    "\U0001f4b0 Cost Analysis", "\U0001f9ea Chemical Analysis", "\u26a0\ufe0f Complications", "\U0001f4cd Well Location"
])

# ═══════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ═══════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-header">\U0001f4ca Total Wells & Phases</div>', unsafe_allow_html=True)
    col_left, col_right = st.columns(2)

    with col_left:
        asset_counts = filtered_wells.groupby("Asset").size().reset_index(name="Count")
        fig_asset = px.bar(asset_counts, x="Asset", y="Count", color="Asset",
                           title="Number of Wells by Asset", text="Count",
                           color_discrete_sequence=px.colors.qualitative.Set2)
        fig_asset.update_traces(textposition="outside")
        fig_asset.update_layout(showlegend=False, height=320, margin=dict(t=40, b=20))
        st.plotly_chart(fig_asset, use_container_width=True)

    with col_right:
        cat_counts = filtered_wells.groupby("Category").size().reset_index(name="Count")
        fig_cat = px.bar(cat_counts, x="Category", y="Count", color="Category",
                         title="Wells by Category", text="Count",
                         color_discrete_sequence=px.colors.qualitative.Set1)
        fig_cat.update_traces(textposition="outside")
        fig_cat.update_layout(showlegend=False, height=320, margin=dict(t=40, b=20))
        st.plotly_chart(fig_cat, use_container_width=True)

    st.markdown('<div class="section-header">\U0001f4cf Well Count by Phase & Hole Size</div>', unsafe_allow_html=True)

    hole_phase = filtered_phases[filtered_phases["Hole Size"] != "N/A"].copy()
    hole_size_order = ['42"', '36"', '26"', '17.5"', '14-3/4"', '12.25"', '8.5"', '6"']

    col_a, col_b = st.columns(2)
    with col_a:
        if not hole_phase.empty:
            hs_counts = hole_phase.groupby("Hole Size").agg(
                Count=("Well Name", "nunique")).reset_index()
            fig_hs = px.bar(hs_counts, x="Hole Size", y="Count",
                            title="Unique Wells per Hole Size", text="Count", color="Hole Size",
                            color_discrete_sequence=px.colors.sequential.Blues_r)
            fig_hs.update_traces(textposition="outside")
            fig_hs.update_layout(showlegend=False, height=340, margin=dict(t=40, b=20))
            st.plotly_chart(fig_hs, use_container_width=True)

    with col_b:
        mud_counts = filtered_phases.groupby("Mud Type").size().reset_index(name="Phase Count")
        mud_counts = mud_counts.sort_values("Phase Count", ascending=True)
        fig_mud = px.bar(mud_counts, y="Mud Type", x="Phase Count",
                         title="Number of Phases by Mud Type", text="Phase Count",
                         orientation="h", color="Mud Type",
                         color_discrete_map=MUD_TYPE_COLORS)
        fig_mud.update_traces(textposition="outside")
        fig_mud.update_layout(showlegend=False, height=340, margin=dict(t=40, b=20))
        st.plotly_chart(fig_mud, use_container_width=True)

    # Grouped bar: Phase distribution
    st.markdown('<div class="section-header">\U0001f500 Phase Distribution by Hole Size & Well (Top 30)</div>', unsafe_allow_html=True)
    ph_well_hs = filtered_phases[filtered_phases["Hole Size"] != "N/A"].groupby(
        ["Well Name", "Hole Size"]).size().reset_index(name="Count")
    # Limit to top 30 wells by total count
    top_wells = ph_well_hs.groupby("Well Name")["Count"].sum().nlargest(30).index
    ph_well_hs_top = ph_well_hs[ph_well_hs["Well Name"].isin(top_wells)]
    fig_grouped = px.bar(ph_well_hs_top, x="Well Name", y="Count", color="Hole Size",
                         title="Phase Count per Well by Hole Size",
                         barmode="stack", category_orders={"Hole Size": hole_size_order},
                         color_discrete_sequence=px.colors.qualitative.Plotly)
    fig_grouped.update_layout(height=400, margin=dict(t=40, b=100), xaxis_tickangle=-45)
    st.plotly_chart(fig_grouped, use_container_width=True)


# ═══════════════════════════════════════════════════════════
# TAB 2 — NPT DETAILS
# ═══════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-header">\u23f1\ufe0f Well-wise NPT Details</div>', unsafe_allow_html=True)

    npt_display = filtered_npt[[
        "Well Name", "Asset", "Phase",
        "Mud Loss (Hrs)", "Activity (Check-up Hrs)", "Unplanned Waiting (Hrs)",
        "Stuck Up (Hrs)", "Total NPT (Hrs)"
    ]].copy()

    # Only show wells with NPT > 0
    npt_display_nonzero = npt_display[npt_display["Total NPT (Hrs)"] > 0]

    total_row = pd.DataFrame([{
        "Well Name": "TOTAL", "Asset": "", "Phase": "",
        "Mud Loss (Hrs)": npt_display_nonzero["Mud Loss (Hrs)"].sum(),
        "Activity (Check-up Hrs)": npt_display_nonzero["Activity (Check-up Hrs)"].sum(),
        "Unplanned Waiting (Hrs)": npt_display_nonzero["Unplanned Waiting (Hrs)"].sum(),
        "Stuck Up (Hrs)": npt_display_nonzero["Stuck Up (Hrs)"].sum(),
        "Total NPT (Hrs)": npt_display_nonzero["Total NPT (Hrs)"].sum(),
    }])
    npt_with_total = pd.concat([npt_display_nonzero, total_row], ignore_index=True)
    st.dataframe(npt_with_total, use_container_width=True, height=400)

    col_npt1, col_npt2 = st.columns(2)
    with col_npt1:
        st.markdown('<div class="section-header">\U0001f967 Total NPT Breakdown</div>', unsafe_allow_html=True)
        npt_totals = {
            "Mud Loss": npt_display["Mud Loss (Hrs)"].sum(),
            "Well Activity": npt_display["Activity (Check-up Hrs)"].sum(),
            "Unplanned Waiting": npt_display["Unplanned Waiting (Hrs)"].sum(),
            "Stuck Up": npt_display["Stuck Up (Hrs)"].sum(),
        }
        npt_totals = {k: v for k, v in npt_totals.items() if v > 0}
        if npt_totals:
            fig_npt_pie = px.pie(names=list(npt_totals.keys()), values=list(npt_totals.values()),
                                 title="Total NPT Breakdown (Hrs)",
                                 color_discrete_sequence=["#e53935", "#ff9800", "#2196f3", "#9c27b0"],
                                 hole=0.4)
            fig_npt_pie.update_traces(textposition="inside", textinfo="percent+label")
            fig_npt_pie.update_layout(height=380, margin=dict(t=50, b=20))
            st.plotly_chart(fig_npt_pie, use_container_width=True)

    with col_npt2:
        st.markdown('<div class="section-header">\U0001f4ca NPT by Well (Top 20)</div>', unsafe_allow_html=True)
        npt_well = filtered_wells[["Well Name", "Mud Loss NPT (Hrs)", "Activity NPT (Hrs)",
                                   "Unplanned Waiting NPT (Hrs)", "Stuck Up NPT (Hrs)"]].copy()
        npt_well["Total"] = npt_well.iloc[:, 1:].sum(axis=1)
        npt_well = npt_well[npt_well["Total"] > 0].nlargest(20, "Total").drop("Total", axis=1)
        if not npt_well.empty:
            fig_npt_well = px.bar(
                npt_well.melt(id_vars="Well Name", var_name="NPT Type", value_name="Hours"),
                x="Hours", y="Well Name", color="NPT Type", orientation="h",
                title="NPT Hours by Well (Top 20)",
                color_discrete_sequence=["#e53935", "#ff9800", "#2196f3", "#9c27b0"])
            fig_npt_well.update_layout(height=500, margin=dict(t=50, b=20))
            st.plotly_chart(fig_npt_well, use_container_width=True)

    # Top Activities Contributing to NPT
    st.markdown('<div class="section-header">\U0001f3c6 Top Activities Contributing to NPT</div>', unsafe_allow_html=True)
    mud_loss_all = data["mud_loss"]
    if not mud_loss_all.empty:
        activity_npt = mud_loss_all.groupby("Operation in Brief").size().reset_index(name="Count")
        activity_npt = activity_npt[activity_npt["Operation in Brief"].str.strip() != ""].sort_values("Count", ascending=True).tail(15)
        if not activity_npt.empty:
            fig_activities = px.bar(activity_npt, y="Operation in Brief", x="Count",
                                    title="Top Activities in Complications",
                                    orientation="h", text="Count", color="Count",
                                    color_continuous_scale="Reds")
            fig_activities.update_traces(textposition="outside")
            fig_activities.update_layout(height=max(300, len(activity_npt) * 35 + 80),
                                         margin=dict(t=50, b=20), coloraxis_showscale=False)
            st.plotly_chart(fig_activities, use_container_width=True)

    csv_npt = npt_display.to_csv(index=False).encode()
    st.download_button("\u2b07\ufe0f Export NPT Data (CSV)", csv_npt, "npt_data.csv", "text/csv")


# ═══════════════════════════════════════════════════════════
# TAB 3 — WELL COVERAGE
# ═══════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-header">\U0001f5fa\ufe0f Well Coverage Summary</div>', unsafe_allow_html=True)

    coverage_cols = [
        "Well Name", "Asset", "Field", "Category",
        "Max Depth (m)", "Total Mud Handled (bbl)",
        "Cost per Meter (INR)", "Cost per Barrel (INR)",
        "Total Cost (INR)", "Spud Date", "TD Date",
        "Planned Days", "Actual Days", "Time Variance (days)"
    ]
    coverage_df = filtered_wells[coverage_cols].copy()
    coverage_df["Total Cost (INR Cr)"] = (coverage_df["Total Cost (INR)"] / 1e7).round(3)

    st.dataframe(
        coverage_df.style.format({
            "Max Depth (m)": "{:,.0f}",
            "Total Mud Handled (bbl)": "{:,.0f}",
            "Cost per Meter (INR)": "{:,.0f}",
            "Cost per Barrel (INR)": "{:,.0f}",
            "Total Cost (INR)": "{:,.0f}",
            "Total Cost (INR Cr)": "{:.3f}",
            "Time Variance (days)": "{:+.1f}",
        }),
        use_container_width=True, height=500
    )

    st.markdown('<div class="section-header">\U0001f4c8 Well Coverage KPIs</div>', unsafe_allow_html=True)
    col_kpi1, col_kpi2 = st.columns(2)

    with col_kpi1:
        top30_depth = filtered_wells.nlargest(30, "Max Depth (m)")
        fig_depth = px.bar(top30_depth, x="Well Name", y="Max Depth (m)", color="Asset",
                           title="Maximum Depth by Well (Top 30)", text="Max Depth (m)",
                           color_discrete_sequence=px.colors.qualitative.Set1)
        fig_depth.update_traces(textposition="outside", texttemplate="%{text:,.0f}")
        fig_depth.update_layout(height=400, margin=dict(t=50, b=100), xaxis_tickangle=-45)
        st.plotly_chart(fig_depth, use_container_width=True)

    with col_kpi2:
        top30_mud = filtered_wells.nlargest(30, "Total Mud Handled (bbl)")
        fig_mud = px.bar(top30_mud, x="Well Name", y="Total Mud Handled (bbl)", color="Asset",
                         title="Total Mud Handled (Top 30)", text="Total Mud Handled (bbl)",
                         color_discrete_sequence=px.colors.qualitative.Set1)
        fig_mud.update_traces(textposition="outside", texttemplate="%{text:,.0f}")
        fig_mud.update_layout(height=400, margin=dict(t=50, b=100), xaxis_tickangle=-45)
        st.plotly_chart(fig_mud, use_container_width=True)

    csv_cov = coverage_df.to_csv(index=False).encode()
    st.download_button("\u2b07\ufe0f Export Well Coverage (CSV)", csv_cov, "well_coverage.csv", "text/csv")


# ═══════════════════════════════════════════════════════════
# TAB 4 — COST ANALYSIS
# ═══════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-header">\U0001f4b0 Cost Analysis Table</div>', unsafe_allow_html=True)

    cost_summary_df = filtered_wells[[
        "Well Name", "Asset", "Total Cost (INR)", "Cost per Meter (INR)", "Cost per Barrel (INR)"
    ]].copy()
    cost_summary_df["Total Cost (INR Cr)"] = (cost_summary_df["Total Cost (INR)"] / 1e7).round(3)

    st.dataframe(
        cost_summary_df.style.format({
            "Total Cost (INR Cr)": "{:.3f}",
            "Total Cost (INR)": "{:,.0f}",
            "Cost per Meter (INR)": "{:,.0f}",
            "Cost per Barrel (INR)": "{:,.0f}",
        }),
        use_container_width=True, height=400
    )

    # Phase-wise cost table
    st.markdown('<div class="section-header">\U0001f4cb Phase-wise Cost per Metre, Cost per Barrel & Mud Type</div>',
                unsafe_allow_html=True)
    phase_cost_display = filtered_phases[[
        "Well Name", "Asset", "Phase", "Hole Size", "Mud Type",
        "Actual Cost (INR)", "Cost per Meter (INR)", "Cost per Barrel (INR)",
        "Interval (m)", "Planned Days", "Actual Days"
    ]].copy()

    st.dataframe(
        phase_cost_display.style.format({
            "Actual Cost (INR)": "{:,.0f}",
            "Cost per Meter (INR)": "{:,.0f}",
            "Cost per Barrel (INR)": "{:,.0f}",
            "Interval (m)": "{:,.0f}",
        }),
        use_container_width=True, height=420
    )

    # Cost charts
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        top30_cpm = filtered_wells[filtered_wells["Cost per Meter (INR)"] > 0].nlargest(30, "Cost per Meter (INR)")
        fig_cpm = px.bar(top30_cpm, x="Well Name", y="Cost per Meter (INR)", color="Asset",
                         title="Cost per Meter (Top 30)", text="Cost per Meter (INR)",
                         color_discrete_sequence=px.colors.qualitative.Pastel)
        fig_cpm.update_traces(textposition="outside", texttemplate="%{text:,.0f}")
        fig_cpm.update_layout(height=400, margin=dict(t=50, b=100), xaxis_tickangle=-45)
        st.plotly_chart(fig_cpm, use_container_width=True)

    with col_c2:
        top30_cpb = filtered_wells[filtered_wells["Cost per Barrel (INR)"] > 0].nlargest(30, "Cost per Barrel (INR)")
        fig_cpb = px.bar(top30_cpb, x="Well Name", y="Cost per Barrel (INR)", color="Asset",
                         title="Cost per Barrel (Top 30)", text="Cost per Barrel (INR)",
                         color_discrete_sequence=px.colors.qualitative.Pastel)
        fig_cpb.update_traces(textposition="outside", texttemplate="%{text:,.0f}")
        fig_cpb.update_layout(height=400, margin=dict(t=50, b=100), xaxis_tickangle=-45)
        st.plotly_chart(fig_cpb, use_container_width=True)

    # Total cost comparison
    st.markdown('<div class="section-header">\U0001f4b5 Total Cost Comparison (Top 30)</div>', unsafe_allow_html=True)
    cost_comp = filtered_wells[["Well Name", "Asset", "Total Cost (INR)"]].copy()
    cost_comp["Total Cost (INR Cr)"] = cost_comp["Total Cost (INR)"] / 1e7
    cost_comp_top = cost_comp.nlargest(30, "Total Cost (INR Cr)")
    fig_total_cost = px.bar(cost_comp_top, x="Well Name", y="Total Cost (INR Cr)", color="Asset",
                             title="Total Cost by Well (INR Cr)", text="Total Cost (INR Cr)",
                             color_discrete_sequence=px.colors.qualitative.Set2)
    fig_total_cost.update_traces(textposition="outside", texttemplate="%{text:.2f}")
    fig_total_cost.update_layout(height=420, margin=dict(t=50, b=100), xaxis_tickangle=-45)
    st.plotly_chart(fig_total_cost, use_container_width=True)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        cost_summary_df.to_excel(writer, sheet_name="Well Cost Summary", index=False)
        phase_cost_display.to_excel(writer, sheet_name="Phase Cost Detail", index=False)
    st.download_button("\u2b07\ufe0f Export Cost Analysis (Excel)", buf.getvalue(),
                       "cost_analysis.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ═══════════════════════════════════════════════════════════
# TAB 5 — CHEMICAL ANALYSIS
# ═══════════════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="section-header">\U0001f9ea Chemical Cost Analysis Table</div>', unsafe_allow_html=True)

    chem_df = filtered_chemicals.copy() if not filtered_chemicals.empty else data["chemicals"].copy()

    if not chem_df.empty:
        chem_display = chem_df[[
            "well_name", "phase", "asset", "chemical_name", "unit_size", "consumption_kg", "actual_cost_inr"
        ]].copy()
        chem_display.columns = [
            "Well Name", "Phase", "Asset", "Chemical Name", "Unit Size",
            "Consumption", "Actual Cost (INR)"
        ]
        st.dataframe(
            chem_display.style.format({
                "Consumption": "{:,.0f}",
                "Actual Cost (INR)": "{:,.0f}",
            }),
            use_container_width=True, height=400
        )

    # Pie charts for key chemicals
    st.markdown('<div class="section-header">\U0001f967 Annual Chemical Consumption by Phase (Pie Charts)</div>',
                unsafe_allow_html=True)

    chem_totals = data["chem_totals"]

    # Build list of all chemicals with data, sorted by total consumption descending
    all_available_chemicals = sorted(
        [c for c, v in chem_totals.items() if v.get('total_kg', 0) > 0],
        key=lambda c: chem_totals[c]['total_kg'], reverse=True
    )

    if all_available_chemicals:
        # Dropdown to select which chemical to view
        col_sel1, col_sel2 = st.columns([1, 2])
        with col_sel1:
            selected_chemical = st.selectbox(
                "Select Chemical",
                options=["Show Top 7"] + all_available_chemicals,
                index=0,
                key="chem_pie_selector"
            )

        if selected_chemical == "Show Top 7":
            # Show top 7 chemicals by consumption as a grid of pie charts
            chemicals_to_plot = all_available_chemicals[:7]
            cols_per_row = 3
            for row_start in range(0, len(chemicals_to_plot), cols_per_row):
                cols = st.columns(min(cols_per_row, len(chemicals_to_plot) - row_start))
                for idx, col in enumerate(cols):
                    chem_idx = row_start + idx
                    if chem_idx >= len(chemicals_to_plot):
                        break
                    chem_name = chemicals_to_plot[chem_idx]
                    chem_item = chem_totals[chem_name]
                    phases_data = chem_item.get('phases', {})
                    sorted_phases = sorted(phases_data.items(), key=lambda x: x[1], reverse=True)[:10]
                    if sorted_phases:
                        labels = [p[0] for p in sorted_phases]
                        values = [p[1] for p in sorted_phases]
                        fig_pie = px.pie(names=labels, values=values,
                                         title=f"{chem_name}<br><sub>Total: {chem_item['total_kg']:,.0f}</sub>",
                                         hole=0.35, color_discrete_sequence=px.colors.qualitative.Set3)
                        fig_pie.update_traces(textposition="inside", textinfo="percent+label", textfont_size=9)
                        fig_pie.update_layout(height=300, margin=dict(t=60, b=10, l=10, r=10), showlegend=False)
                        with col:
                            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            # Show detailed pie chart for the selected chemical
            chem_item = chem_totals[selected_chemical]
            phases_data = chem_item.get('phases', {})
            sorted_phases = sorted(phases_data.items(), key=lambda x: x[1], reverse=True)[:15]
            if sorted_phases:
                labels = [p[0] for p in sorted_phases]
                values = [p[1] for p in sorted_phases]

                col_pie, col_info = st.columns([2, 1])
                with col_pie:
                    fig_pie = px.pie(names=labels, values=values,
                                     title=f"{selected_chemical} — Consumption by Phase<br>"
                                           f"<sub>Total: {chem_item['total_kg']:,.0f} | "
                                           f"Total Cost: INR {chem_item.get('total_cost', 0):,.0f}</sub>",
                                     hole=0.4, color_discrete_sequence=px.colors.qualitative.Set3)
                    fig_pie.update_traces(textposition="inside", textinfo="percent+label", textfont_size=10)
                    fig_pie.update_layout(height=450, margin=dict(t=70, b=20, l=20, r=20))
                    st.plotly_chart(fig_pie, use_container_width=True)

                with col_info:
                    st.markdown(f"**{selected_chemical}**")
                    st.markdown(f"- **Total Consumption:** {chem_item['total_kg']:,.0f}")
                    st.markdown(f"- **Total Cost (INR):** {chem_item.get('total_cost', 0):,.0f}")
                    st.markdown(f"- **Phases using it:** {len(phases_data)}")
                    st.markdown("---")
                    st.markdown("**Top Phases:**")
                    for lbl, val in sorted_phases[:10]:
                        pct = val / chem_item['total_kg'] * 100 if chem_item['total_kg'] > 0 else 0
                        st.markdown(f"- {lbl}: {val:,.0f} ({pct:.1f}%)")
            else:
                st.info(f"No phase-wise data available for {selected_chemical}.")
    else:
        st.info("No chemical consumption data available.")

    # Chemical consumption by well
    if not chem_df.empty:
        st.markdown('<div class="section-header">\U0001f4ca Chemical Consumption by Well (Top 20)</div>', unsafe_allow_html=True)
        chem_well_grp = chem_df.groupby(["well_name", "chemical_name"])["consumption_kg"].sum().reset_index()
        top_chem_wells = chem_well_grp.groupby("well_name")["consumption_kg"].sum().nlargest(20).index
        chem_well_top = chem_well_grp[chem_well_grp["well_name"].isin(top_chem_wells)]
        fig_chem_well = px.bar(chem_well_top, x="well_name", y="consumption_kg", color="chemical_name",
                               title="Chemical Consumption by Well", barmode="stack",
                               color_discrete_sequence=px.colors.qualitative.Pastel)
        fig_chem_well.update_layout(height=420, margin=dict(t=50, b=100),
                                    xaxis_tickangle=-45, legend_title="Chemical")
        st.plotly_chart(fig_chem_well, use_container_width=True)

    if not chem_df.empty:
        csv_chem = chem_display.to_csv(index=False).encode()
        st.download_button("\u2b07\ufe0f Export Chemical Data (CSV)", csv_chem, "chemical_data.csv", "text/csv")


# ═══════════════════════════════════════════════════════════
# TAB 6 — COMPLICATIONS
# ═══════════════════════════════════════════════════════════
with tab6:
    comp_cols = [
        "Well Name", "Phase", "Date of Occurrence",
        "Depth of Occurrence (m)", "Mud System", "Operation in Brief",
        "Type of Loss/Stuck Up", "Formation Info",
        "Type of Pill/Action", "Mud Volume Lost (bbl)"
    ]

    # MUD LOSS TABLE
    st.markdown('<div class="section-header">\U0001f30a Mud Loss Events</div>', unsafe_allow_html=True)
    mud_loss_df = data["mud_loss"].copy()
    if sel_wells:
        mud_loss_df = mud_loss_df[mud_loss_df["Well Name"].isin(sel_wells)]
    if sel_assets:
        mud_loss_df = mud_loss_df[mud_loss_df["Asset"].isin(sel_assets)]
    if sel_loss:
        mud_loss_df = mud_loss_df[mud_loss_df["Type of Loss/Stuck Up"].isin(sel_loss)]

    if not mud_loss_df.empty:
        available_cols = [c for c in comp_cols if c in mud_loss_df.columns]
        st.dataframe(mud_loss_df[available_cols], use_container_width=True, height=350)

        col_ml1, col_ml2, col_ml3 = st.columns(3)
        with col_ml1:
            st.metric("Total Mud Loss Events", len(mud_loss_df))
        with col_ml2:
            st.metric("Total Volume Lost (bbl)", f"{mud_loss_df['Mud Volume Lost (bbl)'].sum():,.0f}")
        with col_ml3:
            st.metric("Wells Affected", mud_loss_df["Well Name"].nunique())

        col_ml_g1, col_ml_g2 = st.columns(2)
        with col_ml_g1:
            # Events by Well
            ml_well_counts = mud_loss_df.groupby("Well Name")["Mud Volume Lost (bbl)"].sum().reset_index()
            ml_well_counts = ml_well_counts.sort_values("Mud Volume Lost (bbl)", ascending=True).tail(15)
            if not ml_well_counts.empty:
                fig_ml_well = px.bar(ml_well_counts, y="Well Name", x="Mud Volume Lost (bbl)",
                                     title="Mud Volume Lost by Well (Top 15)",
                                     orientation="h", text="Mud Volume Lost (bbl)", color="Mud Volume Lost (bbl)",
                                     color_continuous_scale="Reds")
                fig_ml_well.update_traces(textposition="outside", texttemplate="%{text:,.0f}")
                fig_ml_well.update_layout(height=max(300, len(ml_well_counts) * 30 + 80),
                                          margin=dict(t=50, b=20), coloraxis_showscale=False)
                st.plotly_chart(fig_ml_well, use_container_width=True)

        with col_ml_g2:
            if "Formation Info" in mud_loss_df.columns:
                form_loss = mud_loss_df.groupby("Formation Info")["Mud Volume Lost (bbl)"].sum().reset_index()
                form_loss = form_loss[form_loss["Formation Info"].str.strip() != ""].sort_values("Mud Volume Lost (bbl)", ascending=True).tail(15)
                if not form_loss.empty:
                    fig_form = px.bar(form_loss, y="Formation Info", x="Mud Volume Lost (bbl)",
                                      title="Mud Volume Lost by Formation (Top 15)",
                                      orientation="h", text="Mud Volume Lost (bbl)",
                                      color="Mud Volume Lost (bbl)", color_continuous_scale="YlOrRd")
                    fig_form.update_traces(textposition="outside", texttemplate="%{text:,.0f}")
                    fig_form.update_layout(height=max(300, len(form_loss) * 30 + 80),
                                           margin=dict(t=50, b=20), coloraxis_showscale=False)
                    st.plotly_chart(fig_form, use_container_width=True)

        # Loss type pie chart and depth distribution
        col_ml_g3, col_ml_g4 = st.columns(2)
        with col_ml_g3:
            type_col = "Type of Loss/Stuck Up"
            if type_col in mud_loss_df.columns:
                ml_types = mud_loss_df.groupby(type_col).size().reset_index(name="Count")
                ml_types = ml_types[ml_types[type_col].astype(str).str.strip() != ""]
                if not ml_types.empty:
                    fig_ml_pie = px.pie(ml_types, names=type_col, values="Count",
                                        title="Mud Loss Events by Type",
                                        hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
                    fig_ml_pie.update_traces(textposition="inside", textinfo="percent+label")
                    fig_ml_pie.update_layout(height=380, margin=dict(t=50, b=20))
                    st.plotly_chart(fig_ml_pie, use_container_width=True)

        with col_ml_g4:
            if "Depth of Occurrence (m)" in mud_loss_df.columns:
                ml_depth = mud_loss_df[mud_loss_df["Depth of Occurrence (m)"] > 0].copy()
                if not ml_depth.empty:
                    fig_ml_depth = px.histogram(ml_depth, x="Depth of Occurrence (m)", nbins=20,
                                                title="Mud Loss Events - Depth Distribution",
                                                color_discrete_sequence=["#e53935"])
                    fig_ml_depth.update_layout(height=380, margin=dict(t=50, b=20))
                    st.plotly_chart(fig_ml_depth, use_container_width=True)
    else:
        st.info("No mud loss events for selected filters.")

    st.markdown("---")

    # WELL ACTIVITY TABLE
    st.markdown('<div class="section-header">\U0001f527 Well Activity Events</div>', unsafe_allow_html=True)
    wa_df = data["well_activity"].copy()
    if sel_wells:
        wa_df = wa_df[wa_df["Well Name"].isin(sel_wells)]
    if sel_assets:
        wa_df = wa_df[wa_df["Asset"].isin(sel_assets)]

    if not wa_df.empty:
        available_cols = [c for c in comp_cols if c in wa_df.columns]
        st.dataframe(wa_df[available_cols], use_container_width=True, height=300)

        col_wa1, col_wa2, col_wa3 = st.columns(3)
        with col_wa1:
            st.metric("Total Well Activity Events", len(wa_df))
        with col_wa2:
            st.metric("Wells Affected", wa_df["Well Name"].nunique())
        with col_wa3:
            npt_col_wa = "NPT (Hrs)" if "NPT (Hrs)" in wa_df.columns else None
            if npt_col_wa:
                st.metric("Total NPT (Hrs)", f"{wa_df[npt_col_wa].sum():,.1f}")

        col_wa_g1, col_wa_g2 = st.columns(2)
        with col_wa_g1:
            # Events by Well
            wa_well_counts = wa_df.groupby("Well Name").size().reset_index(name="Events")
            wa_well_counts = wa_well_counts.sort_values("Events", ascending=True).tail(15)
            if not wa_well_counts.empty:
                fig_wa_well = px.bar(wa_well_counts, y="Well Name", x="Events",
                                     title="Well Activity Events by Well (Top 15)",
                                     orientation="h", text="Events", color="Events",
                                     color_continuous_scale="Oranges")
                fig_wa_well.update_traces(textposition="outside")
                fig_wa_well.update_layout(height=max(300, len(wa_well_counts) * 30 + 80),
                                          margin=dict(t=50, b=20), coloraxis_showscale=False)
                st.plotly_chart(fig_wa_well, use_container_width=True)

        with col_wa_g2:
            # Events by Operation
            if "Operation in Brief" in wa_df.columns:
                wa_ops = wa_df.groupby("Operation in Brief").size().reset_index(name="Count")
                wa_ops = wa_ops[wa_ops["Operation in Brief"].str.strip() != ""].sort_values("Count", ascending=True).tail(15)
                if not wa_ops.empty:
                    fig_wa_ops = px.bar(wa_ops, y="Operation in Brief", x="Count",
                                        title="Well Activity by Operation Type (Top 15)",
                                        orientation="h", text="Count", color="Count",
                                        color_continuous_scale="Blues")
                    fig_wa_ops.update_traces(textposition="outside")
                    fig_wa_ops.update_layout(height=max(300, len(wa_ops) * 30 + 80),
                                             margin=dict(t=50, b=20), coloraxis_showscale=False)
                    st.plotly_chart(fig_wa_ops, use_container_width=True)

        col_wa_g3, col_wa_g4 = st.columns(2)
        with col_wa_g3:
            # Well Activity by Formation
            if "Formation Info" in wa_df.columns:
                wa_form = wa_df.groupby("Formation Info").size().reset_index(name="Events")
                wa_form = wa_form[wa_form["Formation Info"].astype(str).str.strip() != ""].sort_values("Events", ascending=True).tail(15)
                if not wa_form.empty:
                    fig_wa_form = px.bar(wa_form, y="Formation Info", x="Events",
                                         title="Well Activity Events by Formation (Top 15)",
                                         orientation="h", text="Events", color="Events",
                                         color_continuous_scale="Teal")
                    fig_wa_form.update_traces(textposition="outside")
                    fig_wa_form.update_layout(height=max(300, len(wa_form) * 30 + 80),
                                              margin=dict(t=50, b=20), coloraxis_showscale=False)
                    st.plotly_chart(fig_wa_form, use_container_width=True)

        with col_wa_g4:
            # Depth distribution of well activity events
            if "Depth of Occurrence (m)" in wa_df.columns:
                wa_depth = wa_df[wa_df["Depth of Occurrence (m)"] > 0].copy()
                if not wa_depth.empty:
                    fig_wa_depth = px.histogram(wa_depth, x="Depth of Occurrence (m)", nbins=20,
                                                title="Well Activity Events - Depth Distribution",
                                                color_discrete_sequence=["#ff9800"])
                    fig_wa_depth.update_layout(height=300, margin=dict(t=50, b=20))
                    st.plotly_chart(fig_wa_depth, use_container_width=True)
    else:
        st.info("No well activity events for selected filters.")

    st.markdown("---")

    # STUCK UP TABLE
    st.markdown('<div class="section-header">\u2693 Stuck Up Events</div>', unsafe_allow_html=True)
    su_df = data["stuck_up"].copy()
    if sel_wells:
        su_df = su_df[su_df["Well Name"].isin(sel_wells)]
    if sel_assets:
        su_df = su_df[su_df["Asset"].isin(sel_assets)]
    if sel_stuck:
        su_df = su_df[su_df["Type of Loss/Stuck Up"].isin(sel_stuck)]

    if not su_df.empty:
        available_cols = [c for c in comp_cols if c in su_df.columns]
        st.dataframe(su_df[available_cols], use_container_width=True, height=300)

        col_su1, col_su2, col_su3 = st.columns(3)
        with col_su1:
            st.metric("Total Stuck Up Events", len(su_df))
        with col_su2:
            st.metric("Wells Affected", su_df["Well Name"].nunique())
        with col_su3:
            npt_col_su = "NPT (Hrs)" if "NPT (Hrs)" in su_df.columns else None
            if npt_col_su:
                st.metric("Total NPT (Hrs)", f"{su_df[npt_col_su].sum():,.1f}")

        col_su_g1, col_su_g2 = st.columns(2)
        with col_su_g1:
            # Events by Well
            su_well_counts = su_df.groupby("Well Name").size().reset_index(name="Events")
            su_well_counts = su_well_counts.sort_values("Events", ascending=True).tail(15)
            if not su_well_counts.empty:
                fig_su_well = px.bar(su_well_counts, y="Well Name", x="Events",
                                     title="Stuck Up Events by Well (Top 15)",
                                     orientation="h", text="Events", color="Events",
                                     color_continuous_scale="Purples")
                fig_su_well.update_traces(textposition="outside")
                fig_su_well.update_layout(height=max(300, len(su_well_counts) * 30 + 80),
                                          margin=dict(t=50, b=20), coloraxis_showscale=False)
                st.plotly_chart(fig_su_well, use_container_width=True)

        with col_su_g2:
            # Events by Type of Stuck Up
            type_col = "Type of Loss/Stuck Up"
            if type_col in su_df.columns:
                su_types = su_df.groupby(type_col).size().reset_index(name="Count")
                su_types = su_types[su_types[type_col].astype(str).str.strip() != ""].sort_values("Count", ascending=True).tail(15)
                if not su_types.empty:
                    fig_su_type = px.bar(su_types, y=type_col, x="Count",
                                         title="Stuck Up by Type (Top 15)",
                                         orientation="h", text="Count", color="Count",
                                         color_continuous_scale="Reds")
                    fig_su_type.update_traces(textposition="outside")
                    fig_su_type.update_layout(height=max(300, len(su_types) * 30 + 80),
                                              margin=dict(t=50, b=20), coloraxis_showscale=False)
                    st.plotly_chart(fig_su_type, use_container_width=True)

        col_su_g3, col_su_g4 = st.columns(2)
        with col_su_g3:
            # Stuck Up by Operation
            if "Operation in Brief" in su_df.columns:
                su_ops = su_df.groupby("Operation in Brief").size().reset_index(name="Count")
                su_ops = su_ops[su_ops["Operation in Brief"].str.strip() != ""].sort_values("Count", ascending=True).tail(15)
                if not su_ops.empty:
                    fig_su_ops = px.bar(su_ops, y="Operation in Brief", x="Count",
                                        title="Stuck Up by Operation Type (Top 15)",
                                        orientation="h", text="Count", color="Count",
                                        color_continuous_scale="YlOrRd")
                    fig_su_ops.update_traces(textposition="outside")
                    fig_su_ops.update_layout(height=max(300, len(su_ops) * 30 + 80),
                                             margin=dict(t=50, b=20), coloraxis_showscale=False)
                    st.plotly_chart(fig_su_ops, use_container_width=True)

        with col_su_g4:
            # Stuck Up by Formation
            if "Formation Info" in su_df.columns:
                su_form = su_df.groupby("Formation Info").size().reset_index(name="Events")
                su_form = su_form[su_form["Formation Info"].astype(str).str.strip() != ""].sort_values("Events", ascending=True).tail(15)
                if not su_form.empty:
                    fig_su_form = px.bar(su_form, y="Formation Info", x="Events",
                                         title="Stuck Up Events by Formation (Top 15)",
                                         orientation="h", text="Events", color="Events",
                                         color_continuous_scale="Purples")
                    fig_su_form.update_traces(textposition="outside")
                    fig_su_form.update_layout(height=max(300, len(su_form) * 30 + 80),
                                              margin=dict(t=50, b=20), coloraxis_showscale=False)
                    st.plotly_chart(fig_su_form, use_container_width=True)

        # Depth distribution row
        col_su_g5, col_su_g6 = st.columns(2)
        with col_su_g5:
            # Depth distribution of stuck up events
            if "Depth of Occurrence (m)" in su_df.columns:
                su_depth = su_df[su_df["Depth of Occurrence (m)"] > 0].copy()
                if not su_depth.empty:
                    fig_su_depth = px.histogram(su_depth, x="Depth of Occurrence (m)", nbins=20,
                                                title="Stuck Up Events - Depth Distribution",
                                                color_discrete_sequence=["#9c27b0"])
                    fig_su_depth.update_layout(height=300, margin=dict(t=50, b=20))
                    st.plotly_chart(fig_su_depth, use_container_width=True)
    else:
        st.info("No stuck up events for selected filters.")

    # Export all complications
    buf_comp = io.BytesIO()
    with pd.ExcelWriter(buf_comp, engine='openpyxl') as writer:
        if not data["mud_loss"].empty:
            data["mud_loss"].to_excel(writer, sheet_name="Mud Loss", index=False)
        if not data["well_activity"].empty:
            data["well_activity"].to_excel(writer, sheet_name="Well Activity", index=False)
        if not data["stuck_up"].empty:
            data["stuck_up"].to_excel(writer, sheet_name="Stuck Up", index=False)
    st.download_button("\u2b07\ufe0f Export All Complications (Excel)", buf_comp.getvalue(),
                       "complications.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ═══════════════════════════════════════════════════════════
# TAB 7 — WELL LOCATION MAP (Enhanced)
# ═══════════════════════════════════════════════════════════

# Asset color scheme for consistent coloring
ASSET_COLORS = {
    "B&S Asset": "#e63946",
    "Mumbai High": "#457b9d",
    "Neelam-Heera": "#2a9d8f",
    "Exploratory": "#e9c46a",
    "Deep Water": "#f4a261",
    "Other": "#6c757d",
}

with tab7:
    _well_count = len(filtered_wells)
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #0d1b2a 0%, #1b263b 50%, #415a77 100%);
         padding: 18px 24px; border-radius: 12px; margin-bottom: 16px; color: white;">
        <h2 style="margin:0; font-size:1.6rem;">\U0001f30d Well Location Intelligence</h2>
        <p style="margin:4px 0 0 0; opacity:0.8; font-size:0.9rem;">
            Interactive map with {_well_count} wells \u00b7 Asset-colored clusters \u00b7 Complication overlay \u00b7 Click for details
        </p>
    </div>
    """, unsafe_allow_html=True)

    # ── Search by Lat/Lon ─────────────────────────────────────────────────
    st.markdown('<div class="section-header">\U0001f50d Search Location & Nearby Wells</div>', unsafe_allow_html=True)

    col_search1, col_search2, col_search3, col_search4 = st.columns([1, 1, 1, 1])
    with col_search1:
        search_lat = st.number_input("Search Latitude", value=19.4, min_value=-90.0, max_value=90.0,
                                      step=0.01, format="%.4f", key="search_lat")
    with col_search2:
        search_lon = st.number_input("Search Longitude", value=71.5, min_value=-180.0, max_value=180.0,
                                      step=0.01, format="%.4f", key="search_lon")
    with col_search3:
        search_radius = st.selectbox("Radius (km)", [10, 25, 50, 100, 200, 500], index=2, key="search_radius")
    with col_search4:
        do_search = st.button("\U0001f50e Find Nearby Wells", use_container_width=True)

    # ── Build map data with complication info ──────────────────────────────
    map_cols = ["Well Name", "Asset", "Field", "Latitude", "Longitude",
                "Max Depth (m)", "Total Cost (INR)", "Category",
                "Total NPT (Hrs)", "Mud Loss Events", "Well Activity Events",
                "Stuck Up Events", "Total Complications", "Total Mud Handled (bbl)",
                "Cost per Meter (INR)", "Well Status"]
    available_map_cols = [c for c in map_cols if c in filtered_wells.columns]
    map_df = filtered_wells[available_map_cols].copy()
    map_df["Total Cost (INR Cr)"] = (map_df["Total Cost (INR)"] / 1e7).round(3)

    # Build rich hover text with complication details
    def build_hover_text(row):
        lines = [
            f"<b>{row['Well Name']}</b>",
            f"Asset: {row['Asset']} | Field: {row.get('Field', 'N/A')}",
            f"Category: {row.get('Category', 'N/A')} | Status: {row.get('Well Status', 'N/A')}",
            f"Max Depth: {row.get('Max Depth (m)', 0):,.0f} m",
            f"Total Cost: INR {row.get('Total Cost (INR Cr)', 0):.3f} Cr",
            f"Mud Handled: {row.get('Total Mud Handled (bbl)', 0):,.0f} bbl",
            f"Cost/m: INR {row.get('Cost per Meter (INR)', 0):,.0f}",
            f"NPT: {row.get('Total NPT (Hrs)', 0):.1f} hrs",
            "---",
        ]
        ml = row.get('Mud Loss Events', 0)
        wa = row.get('Well Activity Events', 0)
        su = row.get('Stuck Up Events', 0)
        total_comp = ml + wa + su
        if total_comp > 0:
            lines.append(f"<b>COMPLICATIONS ({total_comp} events):</b>")
            if ml > 0:
                lines.append(f"  Mud Loss: {ml} events")
            if wa > 0:
                lines.append(f"  Well Activity: {wa} events")
            if su > 0:
                lines.append(f"  Stuck Up: {su} events")
        else:
            lines.append("No complications recorded")
        return "<br>".join(lines)

    valid_map = map_df[(map_df["Latitude"] != 0) & (map_df["Longitude"] != 0)].copy()

    if not valid_map.empty:
        valid_map["hover_text"] = valid_map.apply(build_hover_text, axis=1)

        # Complication severity for marker size
        valid_map["Marker Size"] = valid_map.get("Total Complications", pd.Series([0]*len(valid_map))).fillna(0).clip(lower=0) + 8

        # Determine if well has complications for icon shape
        valid_map["Has Complications"] = valid_map.get("Total Complications", pd.Series([0]*len(valid_map))).fillna(0) > 0

        # ── Map Style Selection ──────────────────────────────────────
        col_style1, col_style2 = st.columns([1, 3])
        with col_style1:
            map_style = st.selectbox("Map Style", [
                "open-street-map", "carto-positron", "carto-darkmatter"
            ], index=0, key="map_style")

        # ── Main Map with Asset-colored markers ──────────────────────
        st.markdown('<div class="section-header">\U0001f5fa\ufe0f Interactive Well Map ({} wells plotted)</div>'.format(len(valid_map)),
                    unsafe_allow_html=True)

        fig_map = go.Figure()

        # Plot each asset as a separate trace for legend grouping (clustering effect)
        for asset_name in sorted(valid_map["Asset"].unique()):
            asset_data = valid_map[valid_map["Asset"] == asset_name]
            color = ASSET_COLORS.get(asset_name, "#6c757d")

            # Wells WITHOUT complications
            no_comp = asset_data[~asset_data["Has Complications"]]
            if not no_comp.empty:
                fig_map.add_trace(go.Scattermap(
                    lat=no_comp["Latitude"], lon=no_comp["Longitude"],
                    mode="markers",
                    marker=dict(size=10, color=color, opacity=0.85),
                    text=no_comp["hover_text"],
                    hoverinfo="text",
                    name=f"{asset_name}",
                    legendgroup=asset_name,
                ))

            # Wells WITH complications (larger markers with border)
            has_comp = asset_data[asset_data["Has Complications"]]
            if not has_comp.empty:
                fig_map.add_trace(go.Scattermap(
                    lat=has_comp["Latitude"], lon=has_comp["Longitude"],
                    mode="markers",
                    marker=dict(
                        size=has_comp["Marker Size"].tolist(),
                        color=color,
                        opacity=0.95,
                        symbol="triangle",
                    ),
                    text=has_comp["hover_text"],
                    hoverinfo="text",
                    name=f"{asset_name} (Complications)",
                    legendgroup=asset_name,
                ))

        # ── Search pin & nearby wells ─────────────────────────────────
        if do_search:
            # Add search location marker
            fig_map.add_trace(go.Scattermap(
                lat=[search_lat], lon=[search_lon],
                mode="markers+text",
                marker=dict(size=18, color="#ff0000", symbol="star"),
                text=["Search Location"],
                textposition="top center",
                hoverinfo="text",
                hovertext=f"<b>Search Point</b><br>Lat: {search_lat:.4f}<br>Lon: {search_lon:.4f}",
                name="Search Location",
                showlegend=True,
            ))

        # Calculate center
        center_lat = valid_map["Latitude"].mean()
        center_lon = valid_map["Longitude"].mean()

        fig_map.update_layout(
            map=dict(
                style=map_style,
                center=dict(lat=center_lat, lon=center_lon),
                zoom=5.5,
            ),
            height=700,
            margin=dict(t=10, b=10, l=10, r=10),
            legend=dict(
                bgcolor="rgba(255,255,255,0.9)",
                bordercolor="#ccc",
                borderwidth=1,
                font=dict(size=11),
                yanchor="top", y=0.99,
                xanchor="left", x=0.01,
                itemsizing="constant",
            ),
            hoverlabel=dict(
                bgcolor="rgba(26,35,126,0.95)",
                font_size=12,
                font_family="Inter, sans-serif",
                font_color="white",
                bordercolor="#3949ab",
            ),
        )

        st.plotly_chart(fig_map, use_container_width=True)

        # ── Legend explanation ────────────────────────────────────────
        st.markdown("""
        <div style="display:flex; gap:24px; flex-wrap:wrap; padding:8px 16px;
                    background:#f8f9fa; border-radius:8px; font-size:0.85rem; color:#333;">
            <span><b>Circle</b> = No Complications</span>
            <span><b>Triangle</b> = Has Complications (size = severity)</span>
            <span><b>Red Star</b> = Search Location</span>
        </div>
        """, unsafe_allow_html=True)

        # ── Nearby Wells Results (if searched) ───────────────────────
        if do_search:
            st.markdown('<div class="section-header">\U0001f4cd Nearby Wells within {} km of ({:.4f}, {:.4f})</div>'.format(
                search_radius, search_lat, search_lon), unsafe_allow_html=True)

            # Haversine distance calculation
            from math import radians, sin, cos, sqrt, atan2
            def haversine(lat1, lon1, lat2, lon2):
                R = 6371  # Earth radius in km
                dlat = radians(lat2 - lat1)
                dlon = radians(lon2 - lon1)
                a = sin(dlat/2)**2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon/2)**2
                return R * 2 * atan2(sqrt(a), sqrt(1-a))

            valid_map["Distance (km)"] = valid_map.apply(
                lambda r: round(haversine(search_lat, search_lon, r["Latitude"], r["Longitude"]), 2), axis=1)
            nearby = valid_map[valid_map["Distance (km)"] <= search_radius].sort_values("Distance (km)")

            if not nearby.empty:
                st.success(f"Found **{len(nearby)} wells** within {search_radius} km")
                nearby_display = nearby[["Well Name", "Asset", "Field", "Category", "Distance (km)",
                                          "Max Depth (m)", "Total Cost (INR Cr)",
                                          "Total Complications", "Total NPT (Hrs)"]].copy()
                st.dataframe(
                    nearby_display.style.format({
                        "Distance (km)": "{:.2f}",
                        "Max Depth (m)": "{:,.0f}",
                        "Total Cost (INR Cr)": "{:.3f}",
                        "Total NPT (Hrs)": "{:.1f}",
                    }).background_gradient(subset=["Distance (km)"], cmap="RdYlGn_r"),
                    use_container_width=True, height=min(400, len(nearby) * 40 + 60)
                )
            else:
                st.warning(f"No wells found within {search_radius} km of the search location.")

        # ── Well Detail Panel (click-to-view) ────────────────────────
        st.markdown('<div class="section-header">\U0001f4cb Well Detail Panel</div>', unsafe_allow_html=True)

        well_options = sorted(valid_map["Well Name"].tolist())
        selected_well = st.selectbox("Select a well to view details", ["-- Select --"] + well_options,
                                      key="well_detail_select")

        if selected_well != "-- Select --":
            w_row = filtered_wells[filtered_wells["Well Name"] == selected_well].iloc[0]

            col_d1, col_d2, col_d3 = st.columns(3)
            with col_d1:
                st.markdown(f"""
                <div style="background:linear-gradient(135deg,#1a237e,#3949ab); color:white;
                     padding:16px; border-radius:10px;">
                    <h3 style="margin:0;">{selected_well}</h3>
                    <p style="margin:4px 0; opacity:0.9;">Asset: {w_row['Asset']} | Field: {w_row.get('Field','N/A')}</p>
                    <p style="margin:4px 0; opacity:0.9;">Category: {w_row.get('Category','N/A')}</p>
                    <p style="margin:4px 0; opacity:0.9;">Status: {w_row.get('Well Status','N/A')}</p>
                    <p style="margin:4px 0; opacity:0.9;">Coordinates: ({w_row['Latitude']:.4f}, {w_row['Longitude']:.4f})</p>
                </div>
                """, unsafe_allow_html=True)

            with col_d2:
                st.markdown("""<div style="background:#f0f4f8; padding:16px; border-radius:10px;">""", unsafe_allow_html=True)
                st.metric("Max Depth", f"{w_row.get('Max Depth (m)', 0):,.0f} m")
                st.metric("Total Cost", f"INR {w_row.get('Total Cost (INR)', 0)/1e7:.3f} Cr")
                st.metric("Cost per Meter", f"INR {w_row.get('Cost per Meter (INR)', 0):,.0f}")
                st.markdown("</div>", unsafe_allow_html=True)

            with col_d3:
                st.markdown("""<div style="background:#f0f4f8; padding:16px; border-radius:10px;">""", unsafe_allow_html=True)
                st.metric("Total Mud", f"{w_row.get('Total Mud Handled (bbl)', 0):,.0f} bbl")
                st.metric("Total NPT", f"{w_row.get('Total NPT (Hrs)', 0):.1f} hrs")
                st.metric("Planned vs Actual Days", f"{w_row.get('Planned Days',0):.0f} / {w_row.get('Actual Days',0):.0f}")
                st.markdown("</div>", unsafe_allow_html=True)

            # Complication details for selected well
            comp_total = (w_row.get('Mud Loss Events', 0) + w_row.get('Well Activity Events', 0)
                         + w_row.get('Stuck Up Events', 0))
            if comp_total > 0:
                st.markdown(f'<div class="section-header" style="background:linear-gradient(90deg,#b71c1c,#e53935);">'
                            f'\u26a0\ufe0f Complications for {selected_well} ({comp_total} events)</div>',
                            unsafe_allow_html=True)

                comp_col1, comp_col2, comp_col3 = st.columns(3)
                with comp_col1:
                    ml_count = w_row.get('Mud Loss Events', 0)
                    st.metric("\U0001f30a Mud Loss Events", ml_count)
                with comp_col2:
                    wa_count = w_row.get('Well Activity Events', 0)
                    st.metric("\U0001f527 Well Activity Events", wa_count)
                with comp_col3:
                    su_count = w_row.get('Stuck Up Events', 0)
                    st.metric("\u2693 Stuck Up Events", su_count)

                # Show actual complication records
                for comp_type, comp_label, comp_icon in [
                    ("mud_loss", "Mud Loss", "\U0001f30a"),
                    ("well_activity", "Well Activity", "\U0001f527"),
                    ("stuck_up", "Stuck Up", "\u2693"),
                ]:
                    comp_data = data[comp_type]
                    if not comp_data.empty:
                        well_comp = comp_data[comp_data["Well Name"] == selected_well]
                        if not well_comp.empty:
                            st.markdown(f"**{comp_icon} {comp_label} Details:**")
                            display_cols = [c for c in ["Phase", "Date of Occurrence", "Depth of Occurrence (m)",
                                            "Mud System", "Operation in Brief", "Type of Loss/Stuck Up",
                                            "Formation Info", "Mud Volume Lost (bbl)"] if c in well_comp.columns]
                            st.dataframe(well_comp[display_cols], use_container_width=True,
                                         height=min(200, len(well_comp) * 40 + 50))
            else:
                st.success(f"No complications recorded for {selected_well}")

        # ── Asset Cluster Summary ────────────────────────────────────
        st.markdown('<div class="section-header">\U0001f4ca Asset Cluster Summary</div>', unsafe_allow_html=True)

        col_cluster1, col_cluster2 = st.columns(2)
        with col_cluster1:
            asset_summary = valid_map.groupby("Asset").agg(
                Wells=("Well Name", "count"),
                Complications=("Total Complications", "sum"),
                Avg_Lat=("Latitude", "mean"),
                Avg_Lon=("Longitude", "mean"),
            ).reset_index()
            fig_asset_map = px.pie(asset_summary, names="Asset", values="Wells",
                                   title="Wells Distribution by Asset",
                                   color="Asset", color_discrete_map=ASSET_COLORS,
                                   hole=0.4)
            fig_asset_map.update_traces(textposition="inside", textinfo="percent+label+value")
            fig_asset_map.update_layout(height=350, margin=dict(t=50, b=20))
            st.plotly_chart(fig_asset_map, use_container_width=True)

        with col_cluster2:
            comp_by_asset = valid_map.groupby("Asset").agg(
                Mud_Loss=("Mud Loss Events", "sum"),
                Well_Activity=("Well Activity Events", "sum"),
                Stuck_Up=("Stuck Up Events", "sum"),
            ).reset_index()
            comp_melted = comp_by_asset.melt(id_vars="Asset", var_name="Comp Type", value_name="Events")
            fig_comp_asset = px.bar(comp_melted, x="Asset", y="Events", color="Comp Type",
                                    title="Complications by Asset",
                                    barmode="stack",
                                    color_discrete_sequence=["#e53935", "#ff9800", "#9c27b0"])
            fig_comp_asset.update_layout(height=350, margin=dict(t=50, b=20))
            st.plotly_chart(fig_comp_asset, use_container_width=True)

    else:
        st.info("No wells with valid coordinates to display on map.")

    # ── Well Location Table ──────────────────────────────────────────
    st.markdown('<div class="section-header">\U0001f4ca Well Location Table</div>', unsafe_allow_html=True)
    loc_cols = ["Well Name", "Asset", "Field", "Category", "Latitude", "Longitude",
                "Max Depth (m)", "Total Cost (INR Cr)", "Total Complications", "Total NPT (Hrs)"]
    loc_available = [c for c in loc_cols if c in map_df.columns]
    loc_display = map_df[loc_available].copy()

    format_dict = {"Latitude": "{:.6f}", "Longitude": "{:.6f}", "Max Depth (m)": "{:,.0f}",
                   "Total Cost (INR Cr)": "{:.3f}", "Total NPT (Hrs)": "{:.1f}"}
    format_dict = {k: v for k, v in format_dict.items() if k in loc_display.columns}
    st.dataframe(loc_display.style.format(format_dict), use_container_width=True, height=400)

    # ── Automation & Export ──────────────────────────────────────────
    st.markdown('<div class="section-header">\u26a1 Automation \u2014 New Well Card Detection</div>', unsafe_allow_html=True)
    st.markdown(f"**Monitoring folder:** `{WELL_CARDS_DIR}`")
    st.markdown(f"**Currently loaded:** {len(wells_df)} wells")

    if st.button("\U0001f50d Scan for New Wells"):
        new = scan_for_new_wells(WELL_CARDS_DIR)
        if new:
            st.success(f"Found {len(new)} new well(s):")
            for nw in new:
                st.write(f"\u2022 {nw['well_name']}")
        else:
            st.info("No new wells detected.")

    st.markdown("---")
    st.markdown("### \U0001f4e5 Export All Dashboard Data")
    buf_all = io.BytesIO()
    with pd.ExcelWriter(buf_all, engine='openpyxl') as writer:
        filtered_wells.to_excel(writer, sheet_name="Wells Summary", index=False)
        filtered_phases.to_excel(writer, sheet_name="Phase Details", index=False)
        npt_display.to_excel(writer, sheet_name="NPT Details", index=False)
        if not data["mud_loss"].empty:
            data["mud_loss"].to_excel(writer, sheet_name="Mud Loss", index=False)
        if not data["well_activity"].empty:
            data["well_activity"].to_excel(writer, sheet_name="Well Activity", index=False)
        if not data["stuck_up"].empty:
            data["stuck_up"].to_excel(writer, sheet_name="Stuck Up", index=False)
        if not data["chemicals"].empty:
            data["chemicals"].to_excel(writer, sheet_name="Chemicals", index=False)

    st.download_button(
        "\u2b07\ufe0f Export Full Dashboard Data (Excel)",
        buf_all.getvalue(),
        f"well_dashboard_export_{date.today().strftime('%Y%m%d')}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    f"<div style='text-align:center;color:#999;font-size:0.8rem;'>"
    f"Well Dashboard \u2022 ONGC Drilling Fluid Services \u2022 "
    f"Last updated: {datetime.now().strftime('%d %b %Y %H:%M')} \u2022 "
    f"{len(wells_df)} wells loaded"
    f"</div>",
    unsafe_allow_html=True
)
