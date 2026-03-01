import streamlit as st
import pandas as pd
import numpy as np
import pulp
import io

st.set_page_config(page_title="PwC Value Creation: Digital Twin", layout="wide")

st.title("🏭 Strategy & Value Creation: Stochastic Factory Twin")
st.markdown("**Objective:** 13-Week Quarterly CapEx evaluation using Mean/StDev demand simulation and volumetric warehouse constraints.")

# --- 1. CONFIGURATION (13 Weeks = 1 Quarter) ---
WEEKS = [f"W{i+1}" for i in range(13)]

# CRITICAL FIX: Ensure DEFAULT_PRODUCTS is defined before any function calls it
DEFAULT_PRODUCTS = ["Lifting Straps", "Weight Belts", "Knee Sleeves"]

# SCENARIO A: Dedicated (High Efficiency, Rigid)
SCENARIO_A = {
    "Lifting Straps": {"price": 25, "cost": 10, "fine": 10, "premium": 3, "vol_cbm": 0.05, "mean_dem": 3500, "std_dev": 200, "L1": {"rate": 80, "time": 1, "cost": 100}, "L2": {"rate": 20, "time": 24, "cost": 10000}, "L3": {"rate": 20, "time": 24, "cost": 10000}},
    "Weight Belts": {"price": 85, "cost": 45, "fine": 30, "premium": 15, "vol_cbm": 0.20, "mean_dem": 1200, "std_dev": 600, "L1": {"rate": 10, "time": 24, "cost": 10000}, "L2": {"rate": 40, "time": 1, "cost": 100}, "L3": {"rate": 10, "time": 24, "cost": 10000}},
    "Knee Sleeves": {"price": 45, "cost": 22, "fine": 15, "premium": 5, "vol_cbm": 0.10, "mean_dem": 2000, "std_dev": 400, "L1": {"rate": 15, "time": 24, "cost": 10000}, "L2": {"rate": 15, "time": 24, "cost": 10000}, "L3": {"rate": 60, "time": 1, "cost": 100}}
}

# SCENARIO B: Flexible (Agile SMED lines)
SCENARIO_B = {
    "Lifting Straps": {"price": 25, "cost": 10, "fine": 10, "premium": 3, "vol_cbm": 0.05, "mean_dem": 3500, "std_dev": 200, "L1": {"rate": 50, "time": 2, "cost": 500}, "L2": {"rate": 50, "time": 2, "cost": 500}, "L3": {"rate": 50, "time": 2, "cost": 500}},
    "Weight Belts": {"price": 85, "cost": 45, "fine": 30, "premium": 15, "vol_cbm": 0.20, "mean_dem": 1200, "std_dev": 600, "L1": {"rate": 25, "time": 2, "cost": 500}, "L2": {"rate": 25, "time": 2, "cost": 500}, "L3": {"rate": 25, "time": 2, "cost": 500}},
    "Knee Sleeves": {"price": 45, "cost": 22, "fine": 15, "premium": 5, "vol_cbm": 0.10, "mean_dem": 2000, "std_dev": 400, "L1": {"rate": 40, "time": 2, "cost": 500}, "L2": {"rate": 40, "time": 2, "cost": 500}, "L3": {"rate": 40, "time": 2, "cost": 500}}
}

@st.cache_data
def get_default_demand():
    np.random.seed(42)
    q_forecast, w_chase = {}, {}
    for p in DEFAULT_PRODUCTS:
        base = 3500 if "Straps" in p else (1200 if "Belts" in p else 2000)
        spike = 2000 if "Straps" in p else (1500 if "Belts" in p else 2000)
        q_forecast[p] = {WEEKS[i]: np.random.randint(int(base*0.9), int(base*1.1)) for i in range(13)}
        w_chase[p] = {WEEKS[i]: (spike if 4 < i < 8 else 0) for i in range(13)}
    return q_forecast, w_chase

DEFAULT_FORECAST, DEFAULT_CHASE = get_default_demand()

# --- 2. EXCEL TEMPLATE GENERATOR ---
def generate_excel_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        def build_matrix_df(matrix_data):
            rows = []
            for p, data in matrix_data.items():
                row = {"Product": p, "Price": data["price"], "Mat_Cost": data["cost"], "Unit_Vol_CBM": data["vol_cbm"], "Mean_Demand": data["mean_dem"], "Std_Dev": data["std_dev"], "SLA_Fine": data["fine"], "Rush_Premium": data["premium"]}
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
        scen_a_name = st.text_input("Scenario A:", "Dedicated Architecture")
        scen_a_capex = st.number_input(f"{scen_a_name} CapEx", value=5000000, step=500000)
        
        scen_b_name = st.text_input("Scenario B:", "Flexible Architecture")
        scen_b_capex = st.number_input(f"{scen_b_name} CapEx", value=7500000, step=500000)
        
        active_scenario = st.radio("Active Simulation:", [scen_a_name, scen_b_name])
        
        st.subheader("Factory & Warehouse Parameters")
        num_lines = st.slider("Active Lines", 1, 3, 3)
        weekly_capacity = st.slider("Weekly Cap. Per Line (Hrs)", 80, 168, 120)
        labor_rate = st.slider("Direct Labor Rate (£/Hr)", 10, 50, 20)
        fixed_line_cost = st.slider("Fixed Line Overhead (£/Wk)", 1000, 10000, 3000)
        
        # Volumetric Storage Inputs
        max_fg_cbm = st.slider("Max FG Storage Capacity (CBM)", 500, 5000, 2000)
        wh_cbm_cost = st.slider("Warehouse Cost (£/CBM/Wk)", 1.0, 10.0, 4.0)
        
        corp_tax = 0.25 # Fixed 25% for simplicity
        submitted = st.form_submit_button("🚀 Run 13-Week Monte Carlo")

# --- 4. DATA PROCESSING & STOCHASTIC DEMAND GENERATION ---
try:
    if uploaded_file is not None:
        target_sheet = "Scenario_A_Master" if active_scenario == scen_a_name else "Scenario_B_Master"
        eco_df = pd.read_excel(uploaded_file, sheet_name=target_sheet)
        PRODUCTS = eco_df["Product"].tolist()
        
        FINANCIALS = {}
        for _, row in eco_df.iterrows():
            FINANCIALS[row["Product"]] = {
                "price": row["Price"], "cost": row["Mat_Cost"], "vol_cbm": row["Unit_Vol_CBM"],
                "fine": row["SLA_Fine"], "premium": row["Rush_Premium"],
                "mean_dem": row["Mean_Demand"], "std_dev": row["Std_Dev"],
                "L1": {"rate": row.get("L1_Rate", 0), "time": row.get("L1_Setup_Time", 0), "cost": row.get("L1_Setup_Cost", 0)},
                "L2": {"rate": row.get("L2_Rate", 0), "time": row.get("L2_Setup_Time", 0), "cost": row.get("L2_Setup_Cost", 0)},
                "L3": {"rate": row.get("L3_Rate", 0), "time": row.get("L3_Setup_Time", 0), "cost": row.get("L3_Setup_Cost", 0)}
            }
    else:
        PRODUCTS = DEFAULT_PRODUCTS
        FINANCIALS = SCENARIO_A if active_scenario == scen_a_name else SCENARIO_B

    # STOCHASTIC SIMULATION: Generate 13 weeks of demand using Mean and StDev
    np.random.seed(42) # Seeded for consistent demo purposes
    UPFRONT_FORECAST = {}
    WEEKLY_CHASE = {}
    
    for p in PRODUCTS:
        mean = FINANCIALS[p]["mean_dem"]
        std = FINANCIALS[p]["std_dev"]
        UPFRONT_FORECAST[p] = {}
        WEEKLY_CHASE[p] = {}
        for w in WEEKS:
            # Generate simulated actual demand (floor at 0)
            actual_demand = max(0, int(np.random.normal(mean, std)))
            
            # The "Forecast" is the established Mean
            UPFRONT_FORECAST[p][w] = mean
            # The "Chase" is anything above the mean
            WEEKLY_CHASE[p][w] = max(0, actual_demand - mean)
            
except Exception as e:
    st.error(f"❌ Excel Parsing Error. {str(e)}")
    st.stop()

LINES = [f"L{i+1}" for i in range(num_lines)]

# --- 5. THE MILP SOLVER ---
def optimize_operations(lines, capacity_limit):
    prob = pulp.LpProblem("Digital_Twin", pulp.LpMaximize)
    
    prod_line = pulp.LpVariable.dicts("ProdLine", (PRODUCTS, lines, WEEKS), lowBound=0)
    setup_line = pulp.LpVariable.dicts("SetupLine", (PRODUCTS, lines, WEEKS), cat=pulp.LpBinary)
    line_active = pulp.LpVariable.dicts("LineActive", (lines, WEEKS), cat=pulp.LpBinary) 
    
    total_prod = pulp.LpVariable.dicts("TotalProd", (PRODUCTS, WEEKS), lowBound=0)
    inv = pulp.LpVariable.dicts("Inv", (PRODUCTS, WEEKS), lowBound=0)
    sold = pulp.LpVariable.dicts("Sold", (PRODUCTS, WEEKS), lowBound=0)
    shortage = pulp.LpVariable.dicts("Shortage", (PRODUCTS, WEEKS), lowBound=0)
    expedited_sold = pulp.LpVariable.dicts("ExpeditedSold", (PRODUCTS, WEEKS), lowBound=0)
    
    for p in PRODUCTS:
        for w in WEEKS: prob += total_prod[p][w] == pulp.lpSum([prod_line[p][l][w] for l in lines])
    
    revenue = pulp.lpSum([sold[p][w] * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    expedite_rev = pulp.lpSum([expedited_sold[p][w] * FINANCIALS[p]["premium"] for p in PRODUCTS for w in WEEKS])
    materials_cost = pulp.lpSum([total_prod[p][w] * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    
    # CALCULATED HOLDING COST: Unit Volume * Warehouse Cost per CBM
    inv_cost = pulp.lpSum([inv[p][w] * (FINANCIALS[p]["vol_cbm"] * wh_cbm_cost) for p in PRODUCTS for w in WEEKS])
    
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
    
    prob += revenue + expedite_rev - (materials_cost + inv_cost + sla_fines + setup_fees + labor_cost + overhead_cost)

    for p in PRODUCTS:
        for i, w in enumerate(WEEKS):
            total_dem = UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w]
            prob += sold[p][w] <= total_dem
            
            if i == 0: 
                prob += sold[p][w] <= total_prod[p][w]
                prob += total_prod[p][w] - sold[p][w] == inv[p][w]
            else: 
                prob += sold[p][w] <= inv[p][WEEKS[i-1]] + total_prod[p][w]
                prob += inv[p][WEEKS[i-1]] + total_prod[p][w] - sold[p][w] == inv[p][w]
                
            prob += expedited_sold[p][w] <= WEEKLY_CHASE[p][w]
            prob += expedited_sold[p][w] <= sold[p][w]
            prob += shortage[p][w] == total_dem - sold[p][w]
            
            for l in lines:
                if FINANCIALS[p][l]["rate"] > 0:
                    prob += prod_line[p][l][w] <= (capacity_limit * FINANCIALS[p][l]["rate"]) * setup_line[p][l][w]
                    prob += setup_line[p][l][w] <= line_active[l][w]

    for w in WEEKS:
        # VOLUMETRIC STORAGE CONSTRAINT: Total inventory volume cannot exceed warehouse capacity
        prob += pulp.lpSum([inv[p][w] * FINANCIALS[p]["vol_cbm"] for p in PRODUCTS]) <= max_fg_cbm
        
        for l in lines:
            time_used = [(prod_line[p][l][w]/FINANCIALS[p][l]["rate"]) + (setup_line[p][l][w]*FINANCIALS[p][l]["time"]) for p in PRODUCTS if FINANCIALS[p][l]["rate"] > 0]
            prob += pulp.lpSum(time_used) <= capacity_limit * line_active[l][w]

    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=15))
    return prob, total_prod, inv, sold, shortage, prod_line, setup_line, expedited_sold, line_active

# --- 6. EXECUTION ---
def get_val(var): return var.varValue if var.varValue is not None else 0

with st.spinner(f"Simulating 13-Week Stochastic Quarter: {active_scenario}..."):
    prob_status, prod_v, inv_v, sold_v, short_v, p_line_v, s_line_v, expedite_v, active_v = optimize_operations(LINES, weekly_capacity)

# --- 7. TABS & VISUALS ---
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 1. CFO CapEx & ROIC", "📈 2. Quarterly Ops P&L", "⚙️ 3. Machine Routing", "💰 4. Unit Economics", "📥 5. Data & Downloads"])

# Calculate Global KPIs
rev = sum([get_val(sold_v[p][w]) * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
expedite_rev = sum([get_val(expedite_v[p][w]) * FINANCIALS[p]["premium"] for p in PRODUCTS for w in WEEKS])
topline = rev + expedite_rev

mat_cost = sum([get_val(prod_v[p][w]) * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
setups = sum([get_val(s_line_v[p][l][w]) * FINANCIALS[p][l]["cost"] for p in PRODUCTS for l in LINES for w in WEEKS])
holding = sum([get_val(inv_v[p][w]) * (FINANCIALS[p]["vol_cbm"] * wh_cbm_cost) for p in PRODUCTS for w in WEEKS])
fines = sum([get_val(short_v[p][w]) * FINANCIALS[p]["fine"] for p in PRODUCTS for w in WEEKS])

labor = sum([((get_val(p_line_v[p][l][w]) / FINANCIALS[p][l]["rate"]) + (get_val(s_line_v[p][l][w]) * FINANCIALS[p][l]["time"])) * labor_rate for p in PRODUCTS for l in LINES for w in WEEKS if FINANCIALS[p][l]["rate"] > 0])
overhead = sum([get_val(active_v[l][w]) * fixed_line_cost for l in LINES for w in WEEKS])

ebit = topline - (mat_cost + setups + holding + fines + labor + overhead)
nopat = ebit * (1 - corp_tax)
annualized_nopat = nopat * 4 # 4 quarters in a year
active_capex = scen_a_capex if active_scenario == scen_a_name else scen_b_capex
roic = (annualized_nopat / active_capex) * 100 if active_capex > 0 else 0

glob_dem = sum([UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w] for p in PRODUCTS for w in WEEKS])
glob_sold = sum([get_val(sold_v[p][w]) for p in PRODUCTS for w in WEEKS])
service_level = (glob_sold / glob_dem * 100) if glob_dem > 0 else 0

total_factory_hrs = weekly_capacity * len(WEEKS) * num_lines
used_factory_hrs = sum([((get_val(p_line_v[p][l][w]) / FINANCIALS[p][l]["rate"]) + (get_val(s_line_v[p][l][w]) * FINANCIALS[p][l]["time"])) for p in PRODUCTS for l in LINES for w in WEEKS if FINANCIALS[p][l]["rate"] > 0])
factory_utilization = (used_factory_hrs / total_factory_hrs * 100) if total_factory_hrs > 0 else 0

with tab1:
    st.subheader(f"Investment Profile: {active_scenario}")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Required CapEx", f"£{active_capex:,.0f}")
    c2.metric("Quarterly EBIT", f"£{ebit:,.0f}")
    c3.metric("Annualized NOPAT", f"£{annualized_nopat:,.0f}")
    c4.metric("Annualized ROIC", f"{roic:.1f}%")
    
    

with tab2:
    st.subheader(f"Quarterly Operations Dashboard (13 Weeks)")
    k1, k2, k3 = st.columns(3)
    k1.metric("Quarterly Net Contribution (EBIT)", f"£{ebit:,.0f}")
    k2.metric("Quarterly Service Level", f"{service_level:.1f}%")
    k3.metric("Quarterly Factory Utilization", f"{factory_utilization:.1f}%")
    
    st.markdown("---")
    def pct(val): return f"{(val / topline) * 100:.1f}%" if topline > 0 else "0.0%"
    pl_df = pd.DataFrame({
        "Financial Line Item": ["Quarterly Base Revenue", "Quarterly Rush Surcharges", "QUARTERLY TOPLINE", "Materials (COGS)", "Direct Labor", "Line Overhead", "Setup Costs", "Warehouse Holding Costs", "SLA Penalties", "QUARTERLY EBIT", "Corporate Taxes", "QUARTERLY NOPAT"],
        "Amount (£)": [f"£{rev:,.0f}", f"£{expedite_rev:,.0f}", f"£{topline:,.0f}", f"-£{mat_cost:,.0f}", f"-£{labor:,.0f}", f"-£{overhead:,.0f}", f"-£{setups:,.0f}", f"-£{holding:,.0f}", f"-£{fines:,.0f}", f"£{ebit:,.0f}", f"-£{ebit*corp_tax:,.0f}", f"£{nopat:,.0f}"],
        "% of Topline": [pct(rev), pct(expedite_rev), "100.0%", pct(-mat_cost), pct(-labor), pct(-overhead), pct(-setups), pct(-holding), pct(-fines), pct(ebit), pct(-ebit*corp_tax), pct(nopat)]
    })
    st.table(pl_df)

with tab3:
    st.subheader(f"Machine Routing Map")
    
    for l in LINES:
        st.markdown(f"#### 🏭 {l}")
        l_data = []
        for w in WEEKS:
            if get_val(active_v[l][w]) == 0:
                l_data.append({"Week": w, "Status": "🛑 SHUT DOWN"})
                continue
            row = {"Week": w, "Status": "🟢 ACTIVE"}
            total_l_hrs = 0
            for p in PRODUCTS:
                rate = FINANCIALS[p][l]["rate"]
                if rate > 0:
                    p_hrs = get_val(p_line_v[p][l][w]) / rate
                    s_hrs = get_val(s_line_v[p][l][w]) * FINANCIALS[p][l]["time"]
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
    st.markdown("Shows per-unit profitability. *Note: Labor cost is averaged assuming routing to the most efficient compatible line.*")
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
    st.subheader("Data Downloads & Audit")
    
    dem_df = []
    for p in PRODUCTS:
        for w in WEEKS:
            dem_df.append({
                "Product": p, "Week": w, 
                "Mean Forecast": UPFRONT_FORECAST[p][w], 
                "Stochastic Spike": WEEKLY_CHASE[p][w], 
                "Total Simulated Demand": UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w]
            })
    dem_csv = pd.DataFrame(dem_df).to_csv(index=False).encode('utf-8')
    
    mc_df = []
    for p, data in FINANCIALS.items():
        row = {"Product": p}
        for l in LINES:
            if l in data:
                row[f"{l} Rate"] = data[l]['rate']
                row[f"{l} Setup Time"] = data[l]['time']
                row[f"{l} Setup Cost"] = data[l]['cost']
        mc_df.append(row)
    mc_csv = pd.DataFrame(mc_df).to_csv(index=False).encode('utf-8')

    st.download_button(label="📥 Download Simulated 13-Week Demand (CSV)", data=dem_csv, file_name="simulated_demand.csv", mime="text/csv")
    st.download_button(label="📥 Download Machine Constraints (CSV)", data=mc_csv, file_name="machine_constraints.csv", mime="text/csv")
    
    
    st.markdown("---")
    st.markdown(f"**Warehouse Volumetric Check (Max Capacity: {max_fg_cbm} CBM)**")
    wh_data = []
    for w in WEEKS:
        vol = sum([get_val(inv_v[p][w]) * FINANCIALS[p]["vol_cbm"] for p in PRODUCTS])
        wh_data.append({"Week": w, "FG Inventory Volume (CBM)": f"{vol:.1f} / {max_fg_cbm}", "Utilization": f"{(vol/max_fg_cbm)*100:.1f}%"})
    st.dataframe(pd.DataFrame(wh_data), use_container_width=True)
