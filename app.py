import streamlit as st
import pandas as pd
import numpy as np
import pulp
import io

st.set_page_config(page_title="PwC Value Creation: Digital Twin", layout="wide")

st.title("🏭 Strategy & Value Creation: CapEx Digital Twin")
st.markdown("**Objective:** Bridge S&OP execution with CFO metrics. Evaluate factory physical architectures based on Annualized ROIC.")

# --- 1. DEFAULT SCENARIO DATA ---
DEFAULT_WEEKS = [f"W{i+1}" for i in range(12)]
DEFAULT_PRODUCTS = ["Lifting Straps", "Weight Belts", "Knee Sleeves"]

# SCENARIO A: Dedicated (Cheaper CapEx, Rigid Operations)
SCENARIO_A = {
    "Lifting Straps": {"price": 25, "cost": 10, "fine": 10, "premium": 3, "L1": {"rate": 80, "time": 1, "cost": 100}, "L2": {"rate": 20, "time": 24, "cost": 10000}, "L3": {"rate": 20, "time": 24, "cost": 10000}},
    "Weight Belts": {"price": 85, "cost": 45, "fine": 30, "premium": 15, "L1": {"rate": 10, "time": 24, "cost": 10000}, "L2": {"rate": 40, "time": 1, "cost": 100}, "L3": {"rate": 10, "time": 24, "cost": 10000}},
    "Knee Sleeves": {"price": 45, "cost": 22, "fine": 15, "premium": 5, "L1": {"rate": 15, "time": 24, "cost": 10000}, "L2": {"rate": 15, "time": 24, "cost": 10000}, "L3": {"rate": 60, "time": 1, "cost": 100}}
}

# SCENARIO B: Flexible (Expensive CapEx, Agile Operations)
SCENARIO_B = {
    "Lifting Straps": {"price": 25, "cost": 10, "fine": 10, "premium": 3, "L1": {"rate": 50, "time": 2, "cost": 500}, "L2": {"rate": 50, "time": 2, "cost": 500}, "L3": {"rate": 50, "time": 2, "cost": 500}},
    "Weight Belts": {"price": 85, "cost": 45, "fine": 30, "premium": 15, "L1": {"rate": 25, "time": 2, "cost": 500}, "L2": {"rate": 25, "time": 2, "cost": 500}, "L3": {"rate": 25, "time": 2, "cost": 500}},
    "Knee Sleeves": {"price": 45, "cost": 22, "fine": 15, "premium": 5, "L1": {"rate": 40, "time": 2, "cost": 500}, "L2": {"rate": 40, "time": 2, "cost": 500}, "L3": {"rate": 40, "time": 2, "cost": 500}}
}

@st.cache_data
def get_default_demand():
    np.random.seed(42)
    q_forecast, w_chase = {}, {}
    for p in DEFAULT_PRODUCTS:
        base = 3500 if "Straps" in p else (1200 if "Belts" in p else 2000)
        spike = 2000 if "Straps" in p else (1500 if "Belts" in p else 2000)
        q_forecast[p] = {DEFAULT_WEEKS[i]: np.random.randint(int(base*0.9), int(base*1.1)) for i in range(12)}
        w_chase[p] = {DEFAULT_WEEKS[i]: (spike if 4 < i < 8 else 0) for i in range(12)}
    return q_forecast, w_chase

DEFAULT_FORECAST, DEFAULT_CHASE = get_default_demand()

# --- 2. SIDEBAR CONTROLS ---
with st.sidebar:
    uploaded_file = st.file_uploader("Upload Configured Excel (Optional)", type=["xlsx"])
    
    st.markdown("---")
    st.header("🏢 CFO & S&OP Controls")
    with st.form("control_panel"):
        
        st.subheader("1. Architecture CapEx (£)")
        scen_a_name = st.text_input("Name Scenario A:", "Dedicated/Rigid Architecture")
        scen_a_capex = st.number_input(f"{scen_a_name} CapEx", value=5000000, step=500000)
        
        scen_b_name = st.text_input("Name Scenario B:", "Flexible/SMED Architecture")
        scen_b_capex = st.number_input(f"{scen_b_name} CapEx", value=7500000, step=500000)
        
        active_scenario = st.radio("Active Simulation:", [scen_a_name, scen_b_name])
        
        st.subheader("2. Financials & Layout")
        corp_tax = st.slider("Corporate Tax Rate (%)", 0, 40, 25) / 100.0
        num_lines = st.slider("Active Lines", 1, 3, 3)
        weekly_capacity = st.slider("Weekly Cap. Per Line (Hrs)", 80, 168, 120)
        labor_rate = st.slider("Direct Labor Rate (£/Hr)", 10, 50, 20)
        fixed_line_cost = st.slider("Fixed Line Overhead (£/Wk)", 1000, 10000, 3000)
        holding_cost = st.slider("Holding Cost (£/unit/wk)", 0.1, 2.0, 0.2)
        rollover_pct = st.slider("Demand Rollover (%)", 0, 100, 50) / 100.0
        
        submitted = st.form_submit_button("🚀 Run Digital Twin")

# --- 3. DATA PROCESSING ---
try:
    if uploaded_file is not None:
        target_sheet = "Scenario_A_Matrix" if active_scenario == scen_a_name else "Scenario_B_Matrix"
        eco_df = pd.read_excel(uploaded_file, sheet_name=target_sheet)
        f_df = pd.read_excel(uploaded_file, sheet_name="Quarterly_Forecast")
        c_df = pd.read_excel(uploaded_file, sheet_name="Weekly_Chase")
        
        PRODUCTS = eco_df["Product"].tolist()
        WEEKS = [col for col in f_df.columns if col != "Product"]
        
        FINANCIALS, UPFRONT_FORECAST, WEEKLY_CHASE = {}, {}, {}
        for _, row in eco_df.iterrows():
            FINANCIALS[row["Product"]] = {
                "price": row["Price"], "cost": row["Cost"], "fine": row["SLA_Fine"], "premium": row["Rush_Premium"],
                "L1": {"rate": row.get("L1_Rate", 0), "time": row.get("L1_Setup_Time", 0), "cost": row.get("L1_Setup_Cost", 0)},
                "L2": {"rate": row.get("L2_Rate", 0), "time": row.get("L2_Setup_Time", 0), "cost": row.get("L2_Setup_Cost", 0)},
                "L3": {"rate": row.get("L3_Rate", 0), "time": row.get("L3_Setup_Time", 0), "cost": row.get("L3_Setup_Cost", 0)}
            }
        for _, row in f_df.iterrows(): UPFRONT_FORECAST[row["Product"]] = {w: row[w] for w in WEEKS}
        for _, row in c_df.iterrows(): WEEKLY_CHASE[row["Product"]] = {w: row[w] for w in WEEKS}
    else:
        PRODUCTS, WEEKS = DEFAULT_PRODUCTS, DEFAULT_WEEKS
        FINANCIALS = SCENARIO_A if active_scenario == scen_a_name else SCENARIO_B
        UPFRONT_FORECAST, WEEKLY_CHASE = DEFAULT_FORECAST, DEFAULT_CHASE
except Exception as e:
    st.error(f"❌ Excel Parsing Error. {str(e)}")
    st.stop()

LINES = [f"L{i+1}" for i in range(num_lines)]

# --- 4. THE MILP SOLVER ---
def optimize_operations(lines, capacity_limit):
    prob = pulp.LpProblem("Digital_Twin", pulp.LpMaximize)
    
    prod_line = pulp.LpVariable.dicts("ProdLine", (PRODUCTS, lines, WEEKS), lowBound=0)
    setup_line = pulp.LpVariable.dicts("SetupLine", (PRODUCTS, lines, WEEKS), cat=pulp.LpBinary)
    line_active = pulp.LpVariable.dicts("LineActive", (lines, WEEKS), cat=pulp.LpBinary) 
    
    total_prod = pulp.LpVariable.dicts("TotalProd", (PRODUCTS, WEEKS), lowBound=0)
    inv = pulp.LpVariable.dicts("Inv", (PRODUCTS, WEEKS), lowBound=0)
    sold = pulp.LpVariable.dicts("Sold", (PRODUCTS, WEEKS), lowBound=0)
    shortage = pulp.LpVariable.dicts("Shortage", (PRODUCTS, WEEKS), lowBound=0)
    rollover = pulp.LpVariable.dicts("Rollover", (PRODUCTS, WEEKS), lowBound=0)
    expedited_sold = pulp.LpVariable.dicts("ExpeditedSold", (PRODUCTS, WEEKS), lowBound=0)
    
    for p in PRODUCTS:
        for w in WEEKS: prob += total_prod[p][w] == pulp.lpSum([prod_line[p][l][w] for l in lines])
    
    revenue = pulp.lpSum([sold[p][w] * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    expedite_rev = pulp.lpSum([expedited_sold[p][w] * FINANCIALS[p]["premium"] for p in PRODUCTS for w in WEEKS])
    materials_cost = pulp.lpSum([total_prod[p][w] * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    inv_cost = pulp.lpSum([inv[p][w] * holding_cost for p in PRODUCTS for w in WEEKS])
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
    
    # Objective: EBIT (Operating Profit)
    prob += revenue + expedite_rev - (materials_cost + inv_cost + sla_fines + setup_fees + labor_cost + overhead_cost)

    for p in PRODUCTS:
        for i, w in enumerate(WEEKS):
            total_dem = UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w] + (rollover[p][WEEKS[i-1]] if i > 0 else 0)
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
            prob += rollover[p][w] == shortage[p][w] * rollover_pct
            
            for l in lines:
                if FINANCIALS[p][l]["rate"] > 0:
                    prob += prod_line[p][l][w] <= (capacity_limit * FINANCIALS[p][l]["rate"]) * setup_line[p][l][w]
                    prob += setup_line[p][l][w] <= line_active[l][w]

    for l in lines:
        for w in WEEKS:
            time_used = [(prod_line[p][l][w]/FINANCIALS[p][l]["rate"]) + (setup_line[p][l][w]*FINANCIALS[p][l]["time"]) for p in PRODUCTS if FINANCIALS[p][l]["rate"] > 0]
            prob += pulp.lpSum(time_used) <= capacity_limit * line_active[l][w]

    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=15))
    return prob, total_prod, inv, sold, shortage, rollover, prod_line, setup_line, expedited_sold

# --- 5. EXECUTION ---
def get_val(var): return var.varValue if var.varValue is not None else 0

with st.spinner(f"Simulating CapEx Scenario: {active_scenario}..."):
    prob_status, prod_v, inv_v, sold_v, short_v, roll_v, p_line_v, s_line_v, expedite_v = optimize_operations(LINES, weekly_capacity)

# --- 6. VISUALS ---
tab1, tab2, tab3 = st.tabs(["📊 CFO CapEx & ROIC", "📈 Operations P&L", "⚙️ S&OP Factory Audit"])

with tab1:
    st.subheader(f"Investment Profile: {active_scenario}")
    
    # Financial Math
    rev = sum([get_val(sold_v[p][w]) * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    expedite_rev = sum([get_val(expedite_v[p][w]) * FINANCIALS[p]["premium"] for p in PRODUCTS for w in WEEKS])
    topline = rev + expedite_rev
    
    costs = sum([get_val(prod_v[p][w]) * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS]) + \
            sum([get_val(s_line_v[p][l][w]) * FINANCIALS[p][l]["cost"] for p in PRODUCTS for l in LINES for w in WEEKS]) + \
            sum([get_val(inv_v[p][w]) * holding_cost for p in PRODUCTS for w in WEEKS]) + \
            sum([get_val(short_v[p][w]) * FINANCIALS[p]["fine"] for p in PRODUCTS for w in WEEKS]) + \
            sum([((get_val(p_line_v[p][l][w]) / FINANCIALS[p][l]["rate"]) + (get_val(s_line_v[p][l][w]) * FINANCIALS[p][l]["time"])) * labor_rate for p in PRODUCTS for l in LINES for w in WEEKS if FINANCIALS[p][l]["rate"] > 0]) + \
            (len(LINES) * len(WEEKS) * fixed_line_cost)
            
    ebit = topline - costs
    nopat = ebit * (1 - corp_tax)
    
    # Annualize the 12-week NOPAT (52 weeks / 12 weeks = 4.33 multiplier)
    annualized_nopat = nopat * 4.33
    active_capex = scen_a_capex if active_scenario == scen_a_name else scen_b_capex
    roic = (annualized_nopat / active_capex) * 100 if active_capex > 0 else 0
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Required CapEx", f"£{active_capex:,.0f}")
    c2.metric("Quarterly EBIT", f"£{ebit:,.0f}")
    c3.metric("Annualized NOPAT", f"£{annualized_nopat:,.0f}")
    c4.metric("Annualized ROIC", f"{roic:.1f}%")
    
    
    st.info("💡 **The CFO Pitch:** Scenario A might cost £2.5M less to build, but its rigid architecture fails under S&OP volatility, leading to massive fines. By spending the extra CapEx on the Flexible SMED architecture in Scenario B, the factory captures upside premiums, avoids SLA penalties, and ultimately delivers a mathematically superior Return on Invested Capital (ROIC).")

with tab2:
    st.subheader(f"Quarterly P&L ({active_scenario})")
    
    # Service Level Logic
    glob_dem = sum([UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w] + (get_val(roll_v[p][WEEKS[WEEKS.index(w)-1]]) if WEEKS.index(w) > 0 else 0) for p in PRODUCTS for w in WEEKS])
    glob_sold = sum([get_val(sold_v[p][w]) for p in PRODUCTS for w in WEEKS])
    sl = (glob_sold / glob_dem * 100) if glob_dem > 0 else 0
    
    st.metric("Global Service Level (%)", f"{sl:.1f}%")
    st.markdown("---")
    
    # Re-calculate buckets for display
    mat_cost = sum([get_val(prod_v[p][w]) * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    setups = sum([get_val(s_line_v[p][l][w]) * FINANCIALS[p][l]["cost"] for p in PRODUCTS for l in LINES for w in WEEKS])
    labor = sum([((get_val(p_line_v[p][l][w]) / FINANCIALS[p][l]["rate"]) + (get_val(s_line_v[p][l][w]) * FINANCIALS[p][l]["time"])) * labor_rate for p in PRODUCTS for l in LINES for w in WEEKS if FINANCIALS[p][l]["rate"] > 0])
    overhead = len(LINES) * len(WEEKS) * fixed_line_cost
    holding = sum([get_val(inv_v[p][w]) * holding_cost for p in PRODUCTS for w in WEEKS])
    fines = sum([get_val(short_v[p][w]) * FINANCIALS[p]["fine"] for p in PRODUCTS for w in WEEKS])
    
    pl_df = pd.DataFrame({
        "Financial Line Item": ["Base Revenue", "Rush Surcharges", "TOTAL TOPLINE", "Materials (COGS)", "Direct Labor", "Line Overhead", "Setup Costs", "Holding Costs", "SLA Penalties", "EBIT (Operating Profit)", "Corporate Taxes", "NOPAT"],
        "Amount (£)": [f"£{rev:,.0f}", f"£{expedite_rev:,.0f}", f"£{topline:,.0f}", f"-£{mat_cost:,.0f}", f"-£{labor:,.0f}", f"-£{overhead:,.0f}", f"-£{setups:,.0f}", f"-£{holding:,.0f}", f"-£{fines:,.0f}", f"£{ebit:,.0f}", f"-£{ebit*corp_tax:,.0f}", f"£{nopat:,.0f}"]
    })
    st.table(pl_df)

with tab3:
    st.subheader("Line Routing & S&OP Audit")
    sku = st.selectbox("Select SKU to Audit", PRODUCTS)
    
    audit_data = []
    for i, w in enumerate(WEEKS):
        dem = UPFRONT_FORECAST[sku][w] + WEEKLY_CHASE[sku][w] + (get_val(roll_v[sku][WEEKS[i-1]]) if i > 0 else 0)
        
        # Where was it made?
        routing = []
        for l in LINES:
            if get_val(p_line_v[sku][l][w]) > 0: routing.append(f"{l}: {int(get_val(p_line_v[sku][l][w]))}")
        
        audit_data.append({
            "Week": w,
            "Total Demand": int(dem),
            "Line Routing": " | ".join(routing) if routing else "No Prod",
            "Units Delivered": int(get_val(sold_v[sku][w])),
            "SLA Penalty": f"£{int(get_val(short_v[sku][w]) * FINANCIALS[sku]['fine'])}"
        })
    st.dataframe(pd.DataFrame(audit_data), use_container_width=True)
