import streamlit as st
import pandas as pd
import numpy as np
import pulp
import io
import altair as alt

st.set_page_config(page_title="PwC Value Creation: CM Digital Twin", layout="wide")

st.title("🏭 Strategy & Value Creation: CM Factory Twin")
st.markdown("**Context:** We are the Contract Manufacturer (CM). Our objective is to evaluate our CapEx architectures (Dedicated vs. Flexible) against erratic Purchase Orders from our Wholesaler client (e.g., Implus).")

# --- 1. CONFIGURATION & YOUR DEFAULT DATA ---
WEEKS = [f"W{i+1}" for i in range(13)]
DEFAULT_PRODUCTS = ["Lifting Straps", "Weight Belts", "Knee Sleeves", "Gloves"]

# Demand Parameters (Mimicking your Quarterly_Forecast.csv)
DEMAND_PARAMS = {
    "Lifting Straps": {"mean": 3252, "std": 400},
    "Weight Belts": {"mean": 1800, "std": 350},
    "Knee Sleeves": {"mean": 1000, "std": 300},
    "Gloves": {"mean": 3000, "std": 600}
}

# SCENARIO A: Dedicated Line Matrix
SCENARIO_A = {
    "Lifting Straps": {"price": 5, "cost": 2, "fine": 1.25, "premium": 0.75, "rm_cbm": 0.01, "fg_cbm": 0.02, "L1": {"rate": 80, "time": 1, "cost": 100}, "L2": {"rate": 40, "time": 2, "cost": 300}, "L3": {"rate": 40, "time": 2, "cost": 300}},
    "Weight Belts": {"price": 15, "cost": 6, "fine": 3.75, "premium": 2.25, "rm_cbm": 0.03, "fg_cbm": 0.05, "L1": {"rate": 20, "time": 8, "cost": 1500}, "L2": {"rate": 40, "time": 1, "cost": 100}, "L3": {"rate": 20, "time": 2, "cost": 300}},
    "Knee Sleeves": {"price": 12, "cost": 4, "fine": 3.0, "premium": 1.8, "rm_cbm": 0.02, "fg_cbm": 0.03, "L1": {"rate": 30, "time": 8, "cost": 1500}, "L2": {"rate": 30, "time": 2, "cost": 300}, "L3": {"rate": 60, "time": 1, "cost": 100}},
    "Gloves": {"price": 6, "cost": 3, "fine": 1.5, "premium": 0.9, "rm_cbm": 0.01, "fg_cbm": 0.01, "L1": {"rate": 30, "time": 8, "cost": 1500}, "L2": {"rate": 40, "time": 2, "cost": 300}, "L3": {"rate": 40, "time": 2, "cost": 300}}
}

# SCENARIO B: Flexible Matrix
SCENARIO_B = {
    "Lifting Straps": {"price": 5, "cost": 2, "fine": 1.25, "premium": 0.75, "rm_cbm": 0.01, "fg_cbm": 0.02, "L1": {"rate": 60, "time": 2, "cost": 300}, "L2": {"rate": 60, "time": 2, "cost": 300}, "L3": {"rate": 60, "time": 2, "cost": 300}},
    "Weight Belts": {"price": 15, "cost": 6, "fine": 3.75, "premium": 2.25, "rm_cbm": 0.03, "fg_cbm": 0.05, "L1": {"rate": 30, "time": 2, "cost": 300}, "L2": {"rate": 30, "time": 2, "cost": 300}, "L3": {"rate": 30, "time": 2, "cost": 300}},
    "Knee Sleeves": {"price": 12, "cost": 4, "fine": 3.0, "premium": 1.8, "rm_cbm": 0.02, "fg_cbm": 0.03, "L1": {"rate": 45, "time": 2, "cost": 300}, "L2": {"rate": 45, "time": 2, "cost": 300}, "L3": {"rate": 45, "time": 2, "cost": 300}},
    "Gloves": {"price": 6, "cost": 3, "fine": 1.5, "premium": 0.9, "rm_cbm": 0.01, "fg_cbm": 0.01, "L1": {"rate": 30, "time": 2, "cost": 300}, "L2": {"rate": 40, "time": 2, "cost": 300}, "L3": {"rate": 40, "time": 2, "cost": 300}}
}

# --- 2. EXCEL TEMPLATE GENERATOR ---
def generate_excel_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        def build_matrix_df(matrix_data):
            rows = []
            for p, data in matrix_data.items():
                row = {
                    "Product": p, "Price": data["price"], "Mat_Cost": data["cost"], 
                    "RM_Vol_CBM": data["rm_cbm"], "FG_Vol_CBM": data["fg_cbm"], 
                    "SLA_Fine": data["fine"], "Rush_Premium": data["premium"]
                }
                for l in ["L1", "L2", "L3"]:
                    row[f"{l}_Rate"] = data[l]["rate"]
                    row[f"{l}_Setup_Time"] = data[l]["time"]
                    row[f"{l}_Setup_Cost"] = data[l]["cost"]
                rows.append(row)
            return pd.DataFrame(rows)
            
        build_matrix_df(SCENARIO_A).to_excel(writer, sheet_name="Scenario_A_Master", index=False)
        build_matrix_df(SCENARIO_B).to_excel(writer, sheet_name="Scenario_B_Master", index=False)
        
        # Matches your Quarterly_Forecast structure
        f_data = [{"Product": p, "Mean Weekly Demand": DEMAND_PARAMS[p]["mean"], "St Dev": DEMAND_PARAMS[p]["std"]} for p in DEFAULT_PRODUCTS]
        pd.DataFrame(f_data).to_excel(writer, sheet_name="Quarterly_Forecast", index=False)
    return output.getvalue()

# --- 3. STATE MANAGEMENT FOR APPLES-TO-APPLES DEMAND ---
if 'demand_locked' not in st.session_state:
    st.session_state.demand_locked = False
    st.session_state.upfront_forecast = {}
    st.session_state.weekly_chase = {}

def generate_stochastic_demand(products_list, params_dict):
    upfront, chase = {}, {}
    for p in products_list:
        mean = params_dict[p]["mean"]
        std = params_dict[p]["std"]
        upfront[p], chase[p] = {}, {}
        for w in WEEKS:
            actual_demand = max(0, int(np.random.normal(mean, std)))
            upfront[p][w] = mean
            chase[p][w] = max(0, actual_demand - mean)
    st.session_state.upfront_forecast = upfront
    st.session_state.weekly_chase = chase
    st.session_state.demand_locked = True

# --- 4. SIDEBAR CONTROLS ---
with st.sidebar:
    st.header("📂 Data Integration")
    st.download_button("📥 Download Master Template", data=generate_excel_template(), file_name="cm_digital_twin.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    uploaded_file = st.file_uploader("Upload Configured Excel", type=["xlsx"])
    
    st.markdown("---")
    
    # STEP 1: DEMAND GENERATION
    st.header("Step 1: Lock Implus Demand")
    st.info("Generate a single stochastic path first. This ensures Scenario A and B face the exact same PO volatility.")
    
    # Process Uploaded Demand Params
    ACTIVE_PRODUCTS = DEFAULT_PRODUCTS
    ACTIVE_DEMAND_PARAMS = DEMAND_PARAMS
    if uploaded_file is not None:
        try:
            f_df = pd.read_excel(uploaded_file, sheet_name="Quarterly_Forecast")
            ACTIVE_PRODUCTS = f_df["Product"].tolist()
            ACTIVE_DEMAND_PARAMS = {row["Product"]: {"mean": row["Mean Weekly Demand"], "std": row["St Dev"]} for _, row in f_df.iterrows()}
        except Exception:
            st.warning("Could not read Quarterly_Forecast sheet. Using default parameters.")

    if st.button("🎲 Generate Apples-to-Apples Demand Path"):
        generate_stochastic_demand(ACTIVE_PRODUCTS, ACTIVE_DEMAND_PARAMS)
        st.success("Demand Path Locked! Ready for Scenario testing.")

    # Generate a default one on first load so the app doesn't break
    if not st.session_state.demand_locked:
        np.random.seed(42) # Seeded fallback
        generate_stochastic_demand(ACTIVE_PRODUCTS, ACTIVE_DEMAND_PARAMS)
    
    st.markdown("---")
    
    # STEP 2: SCENARIO EXECUTION
    st.header("Step 2: Run CM Architecture")
    with st.form("control_panel"):
        scen_a_name = st.text_input("Scenario A:", "Dedicated Line")
        scen_a_capex = st.number_input(f"{scen_a_name} CapEx", value=5000000, step=500000)
        
        scen_b_name = st.text_input("Scenario B:", "Flexible SMED")
        scen_b_capex = st.number_input(f"{scen_b_name} CapEx", value=7500000, step=500000)
        
        active_scenario = st.radio("Active Simulation:", [scen_a_name, scen_b_name])
        
        st.subheader("Factory & Working Capital")
        num_lines = st.slider("Active Lines", 1, 3, 3)
        weekly_capacity = st.slider("Weekly Cap. Per Line (Hrs)", 80, 168, 120)
        labor_rate = st.slider("Direct Labor Rate (£/Hr)", 10, 50, 20)
        fixed_line_cost = st.slider("Fixed Line Overhead (£/Wk)", 1000, 10000, 3000)
        wc_rate = st.slider("Working Capital Rate (%/Wk)", 0.0, 1.0, 0.2) / 100.0
        
        st.subheader("Physical Warehouse Limits")
        max_rm_cbm = st.slider("Max RM Storage (CBM)", 100, 5000, 1000)
        max_fg_cbm = st.slider("Max FG Storage (CBM)", 500, 10000, 2500)
        wh_cbm_cost = st.slider("Warehouse Lease (£/CBM/Wk)", 0.5, 5.0, 1.5)
        
        rollover_pct = 0.50
        corp_tax = 0.25 
        submitted = st.form_submit_button("🚀 Run Scenario Optimization")

# --- 5. DATA PROCESSING (MATRIX SELECTION) ---
try:
    if uploaded_file is not None:
        target_sheet = "Scenario_A_Master" if active_scenario == scen_a_name else "Scenario_B_Master"
        eco_df = pd.read_excel(uploaded_file, sheet_name=target_sheet)
        FINANCIALS = {}
        for _, row in eco_df.iterrows():
            FINANCIALS[row["Product"]] = {
                "price": row["Price"], "cost": row["Mat_Cost"], 
                "rm_cbm": row.get("RM_Vol_CBM", 0.02), "fg_cbm": row.get("FG_Vol_CBM", 0.03),
                "fine": row["SLA_Fine"], "premium": row["Rush_Premium"],
                "L1": {"rate": row.get("L1_Rate", 0), "time": row.get("L1_Setup_Time", 0), "cost": row.get("L1_Setup_Cost", 0)},
                "L2": {"rate": row.get("L2_Rate", 0), "time": row.get("L2_Setup_Time", 0), "cost": row.get("L2_Setup_Cost", 0)},
                "L3": {"rate": row.get("L3_Rate", 0), "time": row.get("L3_Setup_Time", 0), "cost": row.get("L3_Setup_Cost", 0)}
            }
    else:
        FINANCIALS = SCENARIO_A if active_scenario == scen_a_name else SCENARIO_B
except Exception as e:
    st.error(f"❌ Excel Parsing Error. {str(e)}")
    st.stop()

LINES = [f"L{i+1}" for i in range(num_lines)]

# Pull locked demand from Session State
UPFRONT_FORECAST = st.session_state.upfront_forecast
WEEKLY_CHASE = st.session_state.weekly_chase

# --- 6. THE MILP SOLVER ---
def optimize_operations(lines, capacity_limit):
    prob = pulp.LpProblem("Digital_Twin", pulp.LpMaximize)
    
    prod_line = pulp.LpVariable.dicts("ProdLine", (ACTIVE_PRODUCTS, lines, WEEKS), lowBound=0)
    setup_line = pulp.LpVariable.dicts("SetupLine", (ACTIVE_PRODUCTS, lines, WEEKS), cat=pulp.LpBinary)
    line_active = pulp.LpVariable.dicts("LineActive", (lines, WEEKS), cat=pulp.LpBinary) 
    
    total_prod = pulp.LpVariable.dicts("TotalProd", (ACTIVE_PRODUCTS, WEEKS), lowBound=0)
    rm_purchased = pulp.LpVariable.dicts("RMPurchased", (ACTIVE_PRODUCTS, WEEKS), lowBound=0)
    rm_inv = pulp.LpVariable.dicts("RMInv", (ACTIVE_PRODUCTS, WEEKS), lowBound=0)
    fg_inv = pulp.LpVariable.dicts("FGInv", (ACTIVE_PRODUCTS, WEEKS), lowBound=0)
    sold = pulp.LpVariable.dicts("Sold", (ACTIVE_PRODUCTS, WEEKS), lowBound=0)
    shortage = pulp.LpVariable.dicts("Shortage", (ACTIVE_PRODUCTS, WEEKS), lowBound=0)
    rollover = pulp.LpVariable.dicts("Rollover", (ACTIVE_PRODUCTS, WEEKS), lowBound=0)
    expedited_sold = pulp.LpVariable.dicts("ExpeditedSold", (ACTIVE_PRODUCTS, WEEKS), lowBound=0)
    
    for p in ACTIVE_PRODUCTS:
        for w in WEEKS: prob += total_prod[p][w] == pulp.lpSum([prod_line[p][l][w] for l in lines])
    
    # Financials (CM Perspective)
    revenue = pulp.lpSum([sold[p][w] * FINANCIALS[p]["price"] for p in ACTIVE_PRODUCTS for w in WEEKS])
    expedite_rev = pulp.lpSum([expedited_sold[p][w] * FINANCIALS[p]["premium"] for p in ACTIVE_PRODUCTS for w in WEEKS])
    materials_cost = pulp.lpSum([rm_purchased[p][w] * FINANCIALS[p]["cost"] for p in ACTIVE_PRODUCTS for w in WEEKS])
    capital_cost = pulp.lpSum([(fg_inv[p][w] * FINANCIALS[p]["cost"] * wc_rate) + (rm_inv[p][w] * FINANCIALS[p]["cost"] * wc_rate) for p in ACTIVE_PRODUCTS for w in WEEKS])
    sla_fines = pulp.lpSum([shortage[p][w] * FINANCIALS[p]["fine"] for p in ACTIVE_PRODUCTS for w in WEEKS])
    setup_fees = pulp.lpSum([setup_line[p][l][w] * FINANCIALS[p][l]["cost"] for p in ACTIVE_PRODUCTS for l in lines for w in WEEKS])
    
    labor_hours = []
    for p in ACTIVE_PRODUCTS:
        for l in lines:
            rate = FINANCIALS[p][l]["rate"]
            time = FINANCIALS[p][l]["time"]
            for w in WEEKS:
                if rate > 0: labor_hours.append((prod_line[p][l][w]/rate) + (setup_line[p][l][w]*time))
                else: 
                    prob += prod_line[p][l][w] == 0
                    prob += setup_line[p][l][w] == 0
                    
    labor_cost = pulp.lpSum(labor_hours) * labor_rate
    overhead_cost = pulp.lpSum([line_active[l][w] * fixed_line_cost for l in lines for w in WEEKS])
    
    prob += revenue + expedite_rev - (materials_cost + capital_cost + sla_fines + setup_fees + labor_cost + overhead_cost)

    for p in ACTIVE_PRODUCTS:
        for i, w in enumerate(WEEKS):
            # Asian RM Sourcing Rule: Deliveries only arrive once a month (W1, W5, W9, W13)
            if i % 4 != 0:
                prob += rm_purchased[p][w] == 0

            if i == 0: prob += rm_inv[p][w] == rm_purchased[p][w] - total_prod[p][w]
            else: prob += rm_inv[p][w] == rm_inv[p][WEEKS[i-1]] + rm_purchased[p][w] - total_prod[p][w]

            total_dem = UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w] + (rollover[p][WEEKS[i-1]] if i > 0 else 0)
            prob += sold[p][w] <= total_dem
            
            if i == 0: 
                prob += sold[p][w] <= total_prod[p][w]
                prob += total_prod[p][w] - sold[p][w] == fg_inv[p][w]
            else: 
                prob += sold[p][w] <= fg_inv[p][WEEKS[i-1]] + total_prod[p][w]
                prob += fg_inv[p][WEEKS[i-1]] + total_prod[p][w] - sold[p][w] == fg_inv[p][w]
                
            prob += expedited_sold[p][w] <= WEEKLY_CHASE[p][w]
            prob += expedited_sold[p][w] <= sold[p][w]
            prob += shortage[p][w] == total_dem - sold[p][w]
            prob += rollover[p][w] == shortage[p][w] * rollover_pct
            
            for l in lines:
                if FINANCIALS[p][l]["rate"] > 0:
                    prob += prod_line[p][l][w] <= (capacity_limit * FINANCIALS[p][l]["rate"]) * setup_line[p][l][w]
                    prob += setup_line[p][l][w] <= line_active[l][w]

    for w in WEEKS:
        prob += pulp.lpSum([rm_inv[p][w] * FINANCIALS[p]["rm_cbm"] for p in ACTIVE_PRODUCTS]) <= max_rm_cbm
        prob += pulp.lpSum([fg_inv[p][w] * FINANCIALS[p]["fg_cbm"] for p in ACTIVE_PRODUCTS]) <= max_fg_cbm
        
        for l in lines:
            time_used = [(prod_line[p][l][w]/FINANCIALS[p][l]["rate"]) + (setup_line[p][l][w]*FINANCIALS[p][l]["time"]) for p in ACTIVE_PRODUCTS if FINANCIALS[p][l]["rate"] > 0]
            prob += pulp.lpSum(time_used) <= capacity_limit * line_active[l][w]

    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=15))
    return prob, total_prod, fg_inv, rm_inv, sold, shortage, prod_line, setup_line, expedited_sold, line_active

# --- 7. EXECUTION & KPIs ---
def get_val(var): return var.varValue if var.varValue is not None else 0

with st.spinner(f"Simulating Scenario: {active_scenario}..."):
    prob, total_prod, fg_inv, rm_inv, sold, shortage, prod_line, setup_line, expedited_sold, line_active = optimize_operations(LINES, weekly_capacity)

rev = sum([get_val(sold[p][w]) * FINANCIALS[p]["price"] for p in ACTIVE_PRODUCTS for w in WEEKS])
expedite_rev = sum([get_val(expedited_sold[p][w]) * FINANCIALS[p]["premium"] for p in ACTIVE_PRODUCTS for w in WEEKS])
topline = rev + expedite_rev

mat_cost = sum([get_val(total_prod[p][w]) * FINANCIALS[p]["cost"] for p in ACTIVE_PRODUCTS for w in WEEKS])
setups = sum([get_val(setup_line[p][l][w]) * FINANCIALS[p][l]["cost"] for p in ACTIVE_PRODUCTS for l in LINES for w in WEEKS])
capital_hold_cost = sum([(get_val(fg_inv[p][w]) + get_val(rm_inv[p][w])) * FINANCIALS[p]["cost"] * wc_rate for p in ACTIVE_PRODUCTS for w in WEEKS])
fixed_wh_lease = (max_rm_cbm + max_fg_cbm) * wh_cbm_cost * len(WEEKS)
fines = sum([get_val(shortage[p][w]) * FINANCIALS[p]["fine"] for p in ACTIVE_PRODUCTS for w in WEEKS])
labor = sum([((get_val(prod_line[p][l][w]) / FINANCIALS[p][l]["rate"]) + (get_val(setup_line[p][l][w]) * FINANCIALS[p][l]["time"])) * labor_rate for p in ACTIVE_PRODUCTS for l in LINES for w in WEEKS if FINANCIALS[p][l]["rate"] > 0])
overhead = sum([get_val(line_active[l][w]) * fixed_line_cost for l in LINES for w in WEEKS])

ebit = topline - (mat_cost + setups + capital_hold_cost + fixed_wh_lease + fines + labor + overhead)
nopat = ebit * (1 - corp_tax)
annualized_nopat = nopat * 4 
active_capex = scen_a_capex if active_scenario == scen_a_name else scen_b_capex
roic = (annualized_nopat / active_capex) * 100 if active_capex > 0 else 0

glob_dem = sum([UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w] for p in ACTIVE_PRODUCTS for w in WEEKS])
glob_sold = sum([get_val(sold[p][w]) for p in ACTIVE_PRODUCTS for w in WEEKS])
service_level = (glob_sold / glob_dem * 100) if glob_dem > 0 else 0

total_factory_hrs = weekly_capacity * len(WEEKS) * num_lines
used_factory_hrs = sum([((get_val(prod_line[p][l][w]) / FINANCIALS[p][l]["rate"]) + (get_val(setup_line[p][l][w]) * FINANCIALS[p][l]["time"])) for p in ACTIVE_PRODUCTS for l in LINES for w in WEEKS if FINANCIALS[p][l]["rate"] > 0])
factory_utilization = (used_factory_hrs / total_factory_hrs * 100) if total_factory_hrs > 0 else 0

# --- 8. VISUAL DASHBOARDS ---
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 CFO CapEx & ROIC", "📈 CM to Implus Execution", "⚙️ Machine Routing", "💰 Unit Economics", "📥 Volumes & Data"])

with tab1:
    st.subheader(f"Investment Profile: {active_scenario}")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Required CapEx", f"£{active_capex:,.0f}")
    c2.metric("Quarterly EBIT", f"£{ebit:,.0f}")
    c3.metric("Annualized NOPAT", f"£{annualized_nopat:,.0f}")
    c4.metric("Annualized ROIC", f"{roic:.1f}%")
    
    
    
    st.markdown("---")
    def pct(val): return f"{(val / topline) * 100:.1f}%" if topline > 0 else "0.0%"
    pl_df = pd.DataFrame({
        "Financial Line Item": ["CM Revenue (from Implus POs)", "Rush Surcharges (Agile Upside)", "TOTAL CM TOPLINE", "Materials (COGS)", "Direct Labor", "Line Overhead", "Setup Costs", "Fixed WH Lease (Sunk)", "WACC Holding Cost", "Implus SLA Penalties", "QUARTERLY EBIT", "Corporate Taxes", "QUARTERLY NOPAT"],
        "Amount (£)": [f"£{rev:,.0f}", f"£{expedite_rev:,.0f}", f"£{topline:,.0f}", f"-£{mat_cost:,.0f}", f"-£{labor:,.0f}", f"-£{overhead:,.0f}", f"-£{setups:,.0f}", f"-£{fixed_wh_lease:,.0f}", f"-£{capital_hold_cost:,.0f}", f"-£{fines:,.0f}", f"£{ebit:,.0f}", f"-£{ebit*corp_tax:,.0f}", f"£{nopat:,.0f}"],
        "% of Topline": [pct(rev), pct(expedite_rev), "100.0%", pct(-mat_cost), pct(-labor), pct(-overhead), pct(-setups), pct(-fixed_wh_lease), pct(-capital_hold_cost), pct(-fines), pct(ebit), pct(-ebit*corp_tax), pct(nopat)]
    })
    st.table(pl_df)

with tab2:
    st.subheader("Quarterly Supply & Demand Matching")
    
    k1, k2, k3 = st.columns(3)
    k1.metric("Overall Service Level", f"{service_level:.1f}%")
    k2.metric("Factory Utilization", f"{factory_utilization:.1f}%")
    k3.metric("Total Implus POs", f"{glob_dem:,.0f} units")
    st.markdown("---")
    
    match_data = []
    for p in ACTIVE_PRODUCTS:
        dem = sum([UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w] for w in WEEKS])
        prd = sum([get_val(total_prod[p][w]) for w in WEEKS])
        sld = sum([get_val(sold[p][w]) for w in WEEKS])
        sl = (sld / dem * 100) if dem > 0 else 0
        match_data.append({"Product": p, "Implus Demanded": dem, "CM Produced": prd, "Shipped to Implus": sld, "Service Level": sl})
        
    match_df = pd.DataFrame(match_data)
    chart_df = match_df.melt(id_vars=["Product"], value_vars=["Implus Demanded", "Shipped to Implus"], var_name="Metric", value_name="Units")
    chart = alt.Chart(chart_df).mark_bar().encode(x=alt.X('Product:N', title=''), y=alt.Y('Units:Q', title='Units'), color='Metric:N', xOffset='Metric:N').properties(height=350, title="Quarterly Demand vs. Fulfillment by Product")
    st.altair_chart(chart, use_container_width=True)
    st.dataframe(match_df.style.format({"Service Level": "{:.1f}%"}), use_container_width=True)

with tab3:
    st.subheader(f"Machine Routing Map")
    
    for l in LINES:
        st.markdown(f"#### 🏭 {l}")
        l_data = []
        for w in WEEKS:
            if get_val(line_active[l][w]) == 0:
                l_data.append({"Week": w, "Status": "🛑 SHUT DOWN"})
                continue
            row = {"Week": w, "Status": "🟢 ACTIVE"}
            total_l_hrs = 0
            for p in ACTIVE_PRODUCTS:
                rate = FINANCIALS[p][l]["rate"]
                if rate > 0:
                    p_hrs = get_val(prod_line[p][l][w]) / rate
                    s_hrs = get_val(setup_line[p][l][w]) * FINANCIALS[p][l]["time"]
                    total_l_hrs += (p_hrs + s_hrs)
                    status = []
                    if p_hrs > 0: status.append(f"Prod: {p_hrs:.1f}h")
                    if s_hrs > 0: status.append(f"Setup: {s_hrs:.1f}h")
                    row[p] = " | ".join(status) if status else "-"
                else: row[p] = "Incompatible"
            row["Hrs Used"] = f"{total_l_hrs:.1f} / {weekly_capacity}"
            l_data.append(row)
        st.dataframe(pd.DataFrame(l_data), use_container_width=True)

with tab4:
    st.subheader("Blended Unit Economics")
    eco_data = []
    for p, data in FINANCIALS.items():
        best_rate = max([data[l]["rate"] for l in LINES if data[l]["rate"] > 0] + [1]) 
        unit_labor = labor_rate / best_rate
        unit_overhead = fixed_line_cost / (weekly_capacity * best_rate) 
        cm_gbp = data['price'] - data['cost'] - unit_labor - unit_overhead
        cm_pct = (cm_gbp / data['price']) * 100
        
        eco_data.append({
            "Product": p,
            "CM Price to Implus": f"£{data['price']:.2f}",
            "Material Cost": f"£{data['cost']:.2f}",
            "Labor Cost (Est.)": f"£{unit_labor:.2f}",
            "Overhead (Est.)": f"£{unit_overhead:.2f}",
            "Cont. Margin (£)": f"£{cm_gbp:.2f}",
            "Cont. Margin (%)": f"{cm_pct:.1f}%"
        })
    st.table(pd.DataFrame(eco_data))

with tab5:
    st.subheader("Volumetric Storage Constraints")
    st.info("💡 **Asian Sourcing Constraint:** RM Shipments only arrive once a month (W1, W5, W9). Watch the RM utilization spike on delivery weeks and drain down as production consumes it.")
    
    
    wh_data = []
    for w in WEEKS:
        rm_vol = sum([get_val(rm_inv[p][w]) * FINANCIALS[p]["rm_cbm"] for p in ACTIVE_PRODUCTS])
        fg_vol = sum([get_val(fg_inv[p][w]) * FINANCIALS[p]["fg_cbm"] for p in ACTIVE_PRODUCTS])
        wh_data.append({
            "Week": w, 
            "RM Volume (CBM)": f"{rm_vol:.1f}", "RM Utilization": f"{(rm_vol/max_rm_cbm)*100:.1f}%",
            "FG Volume (CBM)": f"{fg_vol:.1f}", "FG Utilization": f"{(fg_vol/max_fg_cbm)*100:.1f}%"
        })
    st.dataframe(pd.DataFrame(wh_data), use_container_width=True)
    
    st.markdown("---")
    st.subheader("Data Downloads (Apple-to-Apples Path)")
    dem_df = []
    for p in ACTIVE_PRODUCTS:
        for w in WEEKS:
            dem_df.append({
                "Product": p, "Week": w, 
                "Base Forecast": UPFRONT_FORECAST[p][w], 
                "Stochastic Surge": WEEKLY_CHASE[p][w], 
                "Total Implus PO": UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w]
            })
    st.download_button(label="📥 Download Active Demand Path (CSV)", data=pd.DataFrame(dem_df).to_csv(index=False).encode('utf-8'), file_name="locked_stochastic_demand.csv", mime="text/csv")        def build_matrix_df(matrix_data):
            rows = []
            for p, data in matrix_data.items():
                row = {
                    "Product": p, "Price": data["price"], "Mat_Cost": data["cost"], 
                    "RM_Vol_CBM": data["rm_cbm"], "FG_Vol_CBM": data["fg_cbm"], 
                    "Mean_Demand": data["mean_dem"], "Std_Dev": data["std_dev"], 
                    "SLA_Fine": data["fine"], "Rush_Premium": data["premium"]
                }
                for l in ["L1", "L2", "L3"]:
                    row[f"{l}_Rate"] = data[l]["rate"]
                    row[f"{l}_Setup_Time"] = data[l]["time"]
                    row[f"{l}_Setup_Cost"] = data[l]["cost"]
                rows.append(row)
            return pd.DataFrame(rows)
            
        build_matrix_df(SCENARIO_A).to_excel(writer, sheet_name="Scenario_A_Master", index=False)
        build_matrix_df(SCENARIO_B).to_excel(writer, sheet_name="Scenario_B_Master", index=False)
    return output.getvalue()

# --- 3. SIDEBAR CONTROLS ---
with st.sidebar:
    st.header("📂 1. Data Integration")
    st.download_button("📥 Download Master Template", data=generate_excel_template(), file_name="stochastic_digital_twin.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    uploaded_file = st.file_uploader("Upload Configured Excel", type=["xlsx"])
    
    st.markdown("---")
    st.header("🏢 2. Executive Controls")
    with st.form("control_panel"):
        st.subheader("Scenario Architectures")
        scen_a_name = st.text_input("Scenario A:", "Dedicated Line")
        scen_a_capex = st.number_input(f"{scen_a_name} CapEx", value=5000000, step=500000)
        
        scen_b_name = st.text_input("Scenario B:", "Flexible")
        scen_b_capex = st.number_input(f"{scen_b_name} CapEx", value=7500000, step=500000)
        
        active_scenario = st.radio("Active Simulation:", [scen_a_name, scen_b_name])
        
        st.subheader("Factory & Working Capital")
        num_lines = st.slider("Active Lines", 1, 3, 3)
        weekly_capacity = st.slider("Weekly Cap. Per Line (Hrs)", 80, 168, 120)
        labor_rate = st.slider("Direct Labor Rate (£/Hr)", 10, 50, 20)
        fixed_line_cost = st.slider("Fixed Line Overhead (£/Wk)", 1000, 10000, 3000)
        
        wc_rate = st.slider("Working Capital Rate (%/Wk)", 0.0, 1.0, 0.2) / 100.0
        
        st.subheader("Physical Warehouse Limits")
        max_rm_cbm = st.slider("Max RM Storage (CBM)", 100, 5000, 1000)
        max_fg_cbm = st.slider("Max FG Storage (CBM)", 500, 10000, 2500)
        wh_cbm_cost = st.slider("Warehouse Lease (£/CBM/Wk)", 0.5, 5.0, 1.5)
        
        rollover_pct = 0.50 # Hardcoded to save UI space
        corp_tax = 0.25 # Fixed 25% 
        submitted = st.form_submit_button("🚀 Run 13-Week Monte Carlo")

# --- 4. DATA PROCESSING & STOCHASTIC GENERATION ---
try:
    if uploaded_file is not None:
        target_sheet = "Scenario_A_Master" if active_scenario == scen_a_name else "Scenario_B_Master"
        eco_df = pd.read_excel(uploaded_file, sheet_name=target_sheet)
        PRODUCTS = eco_df["Product"].tolist()
        FINANCIALS = {}
        for _, row in eco_df.iterrows():
            FINANCIALS[row["Product"]] = {
                "price": row["Price"], "cost": row["Mat_Cost"], 
                # .get() prevents crashes if uploading older excel files
                "rm_cbm": row.get("RM_Vol_CBM", 0.02), "fg_cbm": row.get("FG_Vol_CBM", 0.03),
                "fine": row["SLA_Fine"], "premium": row["Rush_Premium"],
                "mean_dem": row.get("Mean_Demand", 1500), "std_dev": row.get("Std_Dev", 300),
                "L1": {"rate": row.get("L1_Rate", 0), "time": row.get("L1_Setup_Time", 0), "cost": row.get("L1_Setup_Cost", 0)},
                "L2": {"rate": row.get("L2_Rate", 0), "time": row.get("L2_Setup_Time", 0), "cost": row.get("L2_Setup_Cost", 0)},
                "L3": {"rate": row.get("L3_Rate", 0), "time": row.get("L3_Setup_Time", 0), "cost": row.get("L3_Setup_Cost", 0)}
            }
    else:
        PRODUCTS = DEFAULT_PRODUCTS
        FINANCIALS = SCENARIO_A if active_scenario == scen_a_name else SCENARIO_B

    np.random.seed(42) 
    UPFRONT_FORECAST, WEEKLY_CHASE = {}, {}
    for p in PRODUCTS:
        mean = FINANCIALS[p]["mean_dem"]
        std = FINANCIALS[p]["std_dev"]
        UPFRONT_FORECAST[p], WEEKLY_CHASE[p] = {}, {}
        for w in WEEKS:
            actual_demand = max(0, int(np.random.normal(mean, std)))
            UPFRONT_FORECAST[p][w] = mean
            WEEKLY_CHASE[p][w] = max(0, actual_demand - mean)
except Exception as e:
    st.error(f"❌ Excel Parsing Error. {str(e)}")
    st.stop()

LINES = [f"L{i+1}" for i in range(num_lines)]

# --- 5. THE MILP SOLVER (WITH RM/FG STAGES) ---
def optimize_operations(lines, capacity_limit):
    prob = pulp.LpProblem("Digital_Twin", pulp.LpMaximize)
    
    prod_line = pulp.LpVariable.dicts("ProdLine", (PRODUCTS, lines, WEEKS), lowBound=0)
    setup_line = pulp.LpVariable.dicts("SetupLine", (PRODUCTS, lines, WEEKS), cat=pulp.LpBinary)
    line_active = pulp.LpVariable.dicts("LineActive", (lines, WEEKS), cat=pulp.LpBinary) 
    
    total_prod = pulp.LpVariable.dicts("TotalProd", (PRODUCTS, WEEKS), lowBound=0)
    rm_purchased = pulp.LpVariable.dicts("RMPurchased", (PRODUCTS, WEEKS), lowBound=0)
    rm_inv = pulp.LpVariable.dicts("RMInv", (PRODUCTS, WEEKS), lowBound=0)
    fg_inv = pulp.LpVariable.dicts("FGInv", (PRODUCTS, WEEKS), lowBound=0)
    sold = pulp.LpVariable.dicts("Sold", (PRODUCTS, WEEKS), lowBound=0)
    shortage = pulp.LpVariable.dicts("Shortage", (PRODUCTS, WEEKS), lowBound=0)
    rollover = pulp.LpVariable.dicts("Rollover", (PRODUCTS, WEEKS), lowBound=0)
    expedited_sold = pulp.LpVariable.dicts("ExpeditedSold", (PRODUCTS, WEEKS), lowBound=0)
    
    for p in PRODUCTS:
        for w in WEEKS: prob += total_prod[p][w] == pulp.lpSum([prod_line[p][l][w] for l in lines])
    
    revenue = pulp.lpSum([sold[p][w] * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    expedite_rev = pulp.lpSum([expedited_sold[p][w] * FINANCIALS[p]["premium"] for p in PRODUCTS for w in WEEKS])
    
    # Material cost triggers upon PURCHASE, not production (Cash flow reality)
    materials_cost = pulp.lpSum([rm_purchased[p][w] * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    
    # Capital Holding Cost: WACC applied to value of RM and FG sitting in warehouses
    capital_cost = pulp.lpSum([(fg_inv[p][w] * FINANCIALS[p]["cost"] * wc_rate) + (rm_inv[p][w] * FINANCIALS[p]["cost"] * wc_rate) for p in PRODUCTS for w in WEEKS])
    
    sla_fines = pulp.lpSum([shortage[p][w] * FINANCIALS[p]["fine"] for p in PRODUCTS for w in WEEKS])
    setup_fees = pulp.lpSum([setup_line[p][l][w] * FINANCIALS[p][l]["cost"] for p in PRODUCTS for l in lines for w in WEEKS])
    
    labor_hours = []
    for p in PRODUCTS:
        for l in lines:
            rate = FINANCIALS[p][l]["rate"]
            time = FINANCIALS[p][l]["time"]
            for w in WEEKS:
                if rate > 0: labor_hours.append((prod_line[p][l][w]/rate) + (setup_line[p][l][w]*time))
                else: 
                    prob += prod_line[p][l][w] == 0
                    prob += setup_line[p][l][w] == 0
                    
    labor_cost = pulp.lpSum(labor_hours) * labor_rate
    overhead_cost = pulp.lpSum([line_active[l][w] * fixed_line_cost for l in lines for w in WEEKS])
    
    prob += revenue + expedite_rev - (materials_cost + capital_cost + sla_fines + setup_fees + labor_cost + overhead_cost)

    for p in PRODUCTS:
        for i, w in enumerate(WEEKS):
            # RM Inventory Balance
            if i == 0: prob += rm_inv[p][w] == rm_purchased[p][w] - total_prod[p][w]
            else: prob += rm_inv[p][w] == rm_inv[p][WEEKS[i-1]] + rm_purchased[p][w] - total_prod[p][w]

            # FG Inventory Balance
            total_dem = UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w] + (rollover[p][WEEKS[i-1]] if i > 0 else 0)
            prob += sold[p][w] <= total_dem
            
            if i == 0: 
                prob += sold[p][w] <= total_prod[p][w]
                prob += total_prod[p][w] - sold[p][w] == fg_inv[p][w]
            else: 
                prob += sold[p][w] <= fg_inv[p][WEEKS[i-1]] + total_prod[p][w]
                prob += fg_inv[p][WEEKS[i-1]] + total_prod[p][w] - sold[p][w] == fg_inv[p][w]
                
            prob += expedited_sold[p][w] <= WEEKLY_CHASE[p][w]
            prob += expedited_sold[p][w] <= sold[p][w]
            prob += shortage[p][w] == total_dem - sold[p][w]
            prob += rollover[p][w] == shortage[p][w] * rollover_pct
            
            for l in lines:
                if FINANCIALS[p][l]["rate"] > 0:
                    prob += prod_line[p][l][w] <= (capacity_limit * FINANCIALS[p][l]["rate"]) * setup_line[p][l][w]
                    prob += setup_line[p][l][w] <= line_active[l][w]

    for w in WEEKS:
        # VOLUMETRIC CONSTRAINTS
        prob += pulp.lpSum([rm_inv[p][w] * FINANCIALS[p]["rm_cbm"] for p in PRODUCTS]) <= max_rm_cbm
        prob += pulp.lpSum([fg_inv[p][w] * FINANCIALS[p]["fg_cbm"] for p in PRODUCTS]) <= max_fg_cbm
        
        for l in lines:
            time_used = [(prod_line[p][l][w]/FINANCIALS[p][l]["rate"]) + (setup_line[p][l][w]*FINANCIALS[p][l]["time"]) for p in PRODUCTS if FINANCIALS[p][l]["rate"] > 0]
            prob += pulp.lpSum(time_used) <= capacity_limit * line_active[l][w]

    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=15))
    return prob, total_prod, fg_inv, rm_inv, sold, shortage, prod_line, setup_line, expedited_sold, line_active, rollover

# --- 6. EXECUTION ---
def get_val(var): return var.varValue if var.varValue is not None else 0

with st.spinner(f"Simulating 13-Week Quarter: {active_scenario}..."):
    prob, total_prod, fg_inv, rm_inv, sold, shortage, prod_line, setup_line, expedited_sold, line_active, rollover = optimize_operations(LINES, weekly_capacity)

# --- 7. TABS & VISUALS ---
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 1. CFO CapEx & ROIC", "📈 2. Operations & Matching", "⚙️ 3. Machine Routing", "💰 4. Unit Economics", "📥 5. Volumes & Downloads"])

# Calculate Global KPIs using exactly matching variables
rev = sum([get_val(sold[p][w]) * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
expedite_rev = sum([get_val(expedited_sold[p][w]) * FINANCIALS[p]["premium"] for p in PRODUCTS for w in WEEKS])
topline = rev + expedite_rev

# Material cost based on actual production consumption in this simplified P&L view
mat_cost = sum([get_val(total_prod[p][w]) * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
setups = sum([get_val(setup_line[p][l][w]) * FINANCIALS[p][l]["cost"] for p in PRODUCTS for l in LINES for w in WEEKS])
capital_hold_cost = sum([(get_val(fg_inv[p][w]) + get_val(rm_inv[p][w])) * FINANCIALS[p]["cost"] * wc_rate for p in PRODUCTS for w in WEEKS])
fixed_wh_lease = (max_rm_cbm + max_fg_cbm) * wh_cbm_cost * len(WEEKS)

fines = sum([get_val(shortage[p][w]) * FINANCIALS[p]["fine"] for p in PRODUCTS for w in WEEKS])
labor = sum([((get_val(prod_line[p][l][w]) / FINANCIALS[p][l]["rate"]) + (get_val(setup_line[p][l][w]) * FINANCIALS[p][l]["time"])) * labor_rate for p in PRODUCTS for l in LINES for w in WEEKS if FINANCIALS[p][l]["rate"] > 0])
overhead = sum([get_val(line_active[l][w]) * fixed_line_cost for l in LINES for w in WEEKS])

ebit = topline - (mat_cost + setups + capital_hold_cost + fixed_wh_lease + fines + labor + overhead)
nopat = ebit * (1 - corp_tax)
annualized_nopat = nopat * 4 
active_capex = scen_a_capex if active_scenario == scen_a_name else scen_b_capex
roic = (annualized_nopat / active_capex) * 100 if active_capex > 0 else 0

with tab1:
    st.subheader(f"Investment Profile: {active_scenario}")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Required CapEx", f"£{active_capex:,.0f}")
    c2.metric("Quarterly EBIT", f"£{ebit:,.0f}")
    c3.metric("Annualized NOPAT", f"£{annualized_nopat:,.0f}")
    c4.metric("Annualized ROIC", f"{roic:.1f}%")
    
    
    
    st.markdown("---")
    def pct(val): return f"{(val / topline) * 100:.1f}%" if topline > 0 else "0.0%"
    pl_df = pd.DataFrame({
        "Financial Line Item": ["Quarterly Base Revenue", "Quarterly Rush Surcharges", "QUARTERLY TOPLINE", "Materials (COGS)", "Direct Labor", "Line Overhead", "Setup Costs", "Fixed WH Lease (Sunk)", "WACC Holding Cost", "SLA Penalties", "QUARTERLY EBIT", "Corporate Taxes", "QUARTERLY NOPAT"],
        "Amount (£)": [f"£{rev:,.0f}", f"£{expedite_rev:,.0f}", f"£{topline:,.0f}", f"-£{mat_cost:,.0f}", f"-£{labor:,.0f}", f"-£{overhead:,.0f}", f"-£{setups:,.0f}", f"-£{fixed_wh_lease:,.0f}", f"-£{capital_hold_cost:,.0f}", f"-£{fines:,.0f}", f"£{ebit:,.0f}", f"-£{ebit*corp_tax:,.0f}", f"£{nopat:,.0f}"],
        "% of Topline": [pct(rev), pct(expedite_rev), "100.0%", pct(-mat_cost), pct(-labor), pct(-overhead), pct(-setups), pct(-fixed_wh_lease), pct(-capital_hold_cost), pct(-fines), pct(ebit), pct(-ebit*corp_tax), pct(nopat)]
    })
    st.table(pl_df)

with tab2:
    st.subheader("Quarterly Supply & Demand Matching")
    match_data = []
    for p in PRODUCTS:
        dem = sum([UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w] for w in WEEKS])
        prd = sum([get_val(total_prod[p][w]) for w in WEEKS])
        sld = sum([get_val(sold[p][w]) for w in WEEKS])
        sl = (sld / dem * 100) if dem > 0 else 0
        match_data.append({"Product": p, "Total Demanded": dem, "Total Produced": prd, "Total Delivered": sld, "Service Level": sl})
        
    match_df = pd.DataFrame(match_data)
    
    chart_df = match_df.melt(id_vars=["Product"], value_vars=["Total Demanded", "Total Delivered"], var_name="Metric", value_name="Units")
    chart = alt.Chart(chart_df).mark_bar().encode(
        x=alt.X('Product:N', title=''),
        y=alt.Y('Units:Q', title='Units'),
        color='Metric:N',
        xOffset='Metric:N'
    ).properties(height=350, title="Quarterly Demand vs. Fulfillment by Product")
    st.altair_chart(chart, use_container_width=True)
    
    st.dataframe(match_df.style.format({"Service Level": "{:.1f}%"}), use_container_width=True)

with tab3:
    st.subheader(f"Machine Routing Map")
    
    for l in LINES:
        st.markdown(f"#### 🏭 {l}")
        l_data = []
        for w in WEEKS:
            if get_val(line_active[l][w]) == 0:
                l_data.append({"Week": w, "Status": "🛑 SHUT DOWN"})
                continue
            row = {"Week": w, "Status": "🟢 ACTIVE"}
            total_l_hrs = 0
            for p in PRODUCTS:
                rate = FINANCIALS[p][l]["rate"]
                if rate > 0:
                    p_hrs = get_val(prod_line[p][l][w]) / rate
                    s_hrs = get_val(setup_line[p][l][w]) * FINANCIALS[p][l]["time"]
                    total_l_hrs += (p_hrs + s_hrs)
                    status = []
                    if p_hrs > 0: status.append(f"Prod: {p_hrs:.1f}h")
                    if s_hrs > 0: status.append(f"Setup: {s_hrs:.1f}h")
                    row[p] = " | ".join(status) if status else "-"
                else: row[p] = "Incompatible"
            row["Hrs Used"] = f"{total_l_hrs:.1f} / {weekly_capacity}"
            l_data.append(row)
        st.dataframe(pd.DataFrame(l_data), use_container_width=True)

with tab4:
    st.subheader("Blended Unit Economics")
    eco_data = []
    for p, data in FINANCIALS.items():
        best_rate = max([data[l]["rate"] for l in LINES if data[l]["rate"] > 0] + [1]) 
        unit_labor = labor_rate / best_rate
        unit_overhead = fixed_line_cost / (weekly_capacity * best_rate) 
        cm_gbp = data['price'] - data['cost'] - unit_labor - unit_overhead
        cm_pct = (cm_gbp / data['price']) * 100
        
        eco_data.append({
            "Product": p,
            "Price": f"£{data['price']:.2f}",
            "Material Cost": f"£{data['cost']:.2f}",
            "Labor Cost (Est.)": f"£{unit_labor:.2f}",
            "Overhead (Est.)": f"£{unit_overhead:.2f}",
            "Cont. Margin (£)": f"£{cm_gbp:.2f}",
            "Cont. Margin (%)": f"{cm_pct:.1f}%"
        })
    st.table(pd.DataFrame(eco_data))

with tab5:
    st.subheader("Volumetric Storage Constraints")
    st.markdown(f"**Max Capacities:** Raw Materials ({max_rm_cbm} CBM) | Finished Goods ({max_fg_cbm} CBM)")
    
    wh_data = []
    for w in WEEKS:
        rm_vol = sum([get_val(rm_inv[p][w]) * FINANCIALS[p]["rm_cbm"] for p in PRODUCTS])
        fg_vol = sum([get_val(fg_inv[p][w]) * FINANCIALS[p]["fg_cbm"] for p in PRODUCTS])
        wh_data.append({
            "Week": w, 
            "RM Volume (CBM)": f"{rm_vol:.1f}", "RM Utilization": f"{(rm_vol/max_rm_cbm)*100:.1f}%",
            "FG Volume (CBM)": f"{fg_vol:.1f}", "FG Utilization": f"{(fg_vol/max_fg_cbm)*100:.1f}%"
        })
    st.dataframe(pd.DataFrame(wh_data), use_container_width=True)
    
    st.markdown("---")
    st.subheader("Data Downloads")
    dem_df = []
    for p in PRODUCTS:
        for w in WEEKS:
            dem_df.append({
                "Product": p, "Week": w, 
                "Mean Forecast": UPFRONT_FORECAST[p][w], 
                "Stochastic Spike": WEEKLY_CHASE[p][w], 
                "Total Demand": UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w]
            })
    st.download_button(label="📥 Download Simulated 13-Week Demand (CSV)", data=pd.DataFrame(dem_df).to_csv(index=False).encode('utf-8'), file_name="simulated_demand.csv", mime="text/csv")
