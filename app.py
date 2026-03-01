import streamlit as st
import pandas as pd
import numpy as np
import pulp

st.set_page_config(page_title="PwC Value Creation: Operations Masterclass", layout="wide")

st.title("💎 Strategy & Value Creation: High-Utilization Optimizer")
st.markdown("**Assumption:** Single shared production line running at ~80% baseline utilization.")

# --- 1. CONFIGURATION (12-Week Planning Horizon) ---
WEEKS = [f"W{i+1}" for i in range(12)]
PRODUCTS = ["Lifting Straps", "Weight Belts", "Knee Sleeves"]

FINANCIALS = {
    "Lifting Straps": {"price": 25, "cost": 10, "rate": 60, "co_time": 2, "co_cost": 500},
    "Weight Belts":  {"price": 85, "cost": 45, "rate": 25,  "co_time": 8, "co_cost": 2500},
    "Knee Sleeves":  {"price": 45, "cost": 22, "rate": 40,  "co_time": 4, "co_cost": 1200}
}

# --- 2. SIDEBAR CONTROLS ---
with st.sidebar:
    st.header("🏢 Operating Strategy")
    with st.form("control_panel"):
        mode = st.radio("Model", ["Legacy: Max 1 Setup/Week (Big Batches)", "MILP: Value Optimized (Agile)"])
        
        st.header("⚙️ Capacity")
        weekly_capacity = st.slider("Weekly Machine Hours", 80, 168, 120)
        
        st.header("📊 Demand & Commercial Risk")
        rollover_pct = st.slider("Demand Rollover (%)", 0, 100, 50) / 100.0
        stockout_fine = st.slider("Lost Sale Penalty (£/unit)", 0, 50, 15)
        holding_cost = st.slider("Holding Cost (£/unit/wk)", 0.1, 2.0, 0.2)
        agility_premium = st.slider("Agility Premium (£/unit in stock)", 0, 15, 5) # THE RETURN OF AGILITY
        
        submitted = st.form_submit_button("🚀 Run Optimization")

# --- 3. BASE DEMAND GENERATION ---
@st.cache_data
def get_base_demand():
    np.random.seed(42)
    data = {}
    for p in PRODUCTS:
        base = np.random.randint(800, 1200, 12)
        peak = [800 if 4 < i < 8 else 0 for i in range(12)] # Peak hits in Weeks 5-7
        data[p] = {WEEKS[i]: base[i] + peak[i] for i in range(12)}
    return data

BASE_DEMAND = get_base_demand()

# --- 4. THE SOLVER ---
def optimize_operations(strat, capacity_limit):
    prob = pulp.LpProblem("Value_Model", pulp.LpMaximize)
    
    prod = pulp.LpVariable.dicts("Prod", (PRODUCTS, WEEKS), lowBound=0)
    inv = pulp.LpVariable.dicts("Inv", (PRODUCTS, WEEKS), lowBound=0)
    sold = pulp.LpVariable.dicts("Sold", (PRODUCTS, WEEKS), lowBound=0)
    shortage = pulp.LpVariable.dicts("Shortage", (PRODUCTS, WEEKS), lowBound=0)
    rollover = pulp.LpVariable.dicts("Rollover", (PRODUCTS, WEEKS), lowBound=0)
    setup = pulp.LpVariable.dicts("Setup", (PRODUCTS, WEEKS), cat=pulp.LpBinary)
    
    # Objective: Maximize Profit (Now including Agility Premium)
    revenue = pulp.lpSum([sold[p][w] * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    agility_rev = pulp.lpSum([inv[p][w] * agility_premium for p in PRODUCTS for w in WEEKS])
    
    costs = pulp.lpSum([
        prod[p][w] * FINANCIALS[p]["cost"] + 
        inv[p][w] * holding_cost + 
        setup[p][w] * FINANCIALS[p]["co_cost"] +
        shortage[p][w] * stockout_fine 
        for p in PRODUCTS for w in WEEKS
    ])
    prob += revenue + agility_rev - costs

    for p in PRODUCTS:
        for i, w in enumerate(WEEKS):
            if i == 0: total_demand = BASE_DEMAND[p][w]
            else: total_demand = BASE_DEMAND[p][w] + rollover[p][WEEKS[i-1]]
            
            prob += sold[p][w] <= total_demand
            if i == 0: prob += sold[p][w] <= prod[p][w]
            else: prob += sold[p][w] <= inv[p][WEEKS[i-1]] + prod[p][w]
                
            prob += shortage[p][w] == total_demand - sold[p][w]
            prob += rollover[p][w] == shortage[p][w] * rollover_pct
            
            if i == 0: prob += prod[p][w] - sold[p][w] == inv[p][w]
            else: prob += inv[p][WEEKS[i-1]] + prod[p][w] - sold[p][w] == inv[p][w]
            
            max_possible_production = capacity_limit * FINANCIALS[p]["rate"]
            prob += prod[p][w] <= max_possible_production * setup[p][w]

    for w in WEEKS:
        prob += pulp.lpSum([(prod[p][w]/FINANCIALS[p]["rate"]) + (setup[p][w]*FINANCIALS[p]["co_time"]) for p in PRODUCTS]) <= capacity_limit
        
        if strat == "Legacy: Max 1 Setup/Week (Big Batches)":
            prob += pulp.lpSum([setup[p][w] for p in PRODUCTS]) <= 1

    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=10))
    return prob, prod, inv, sold, shortage, rollover, setup

# --- 5. EXECUTION ---
def get_val(var):
    return var.varValue if var.varValue is not None else 0

with st.spinner("Crunching the MILP Matrix (This may take up to 10 seconds)..."):
    prob_status, prod_v, inv_v, sold_v, short_v, roll_v, setup_v = optimize_operations(mode, weekly_capacity)

if pulp.LpStatus[prob_status.status] != 'Optimal':
    st.warning(f"⚠️ **Optimality Gap Detected:** Status: {pulp.LpStatus[prob_status.status]}. Displaying best margin found within time limit.")
else:
    st.success("✅ **Optimal Schedule Found**")

# --- 6. TABS & VISUALS ---
tab1, tab2, tab3, tab4 = st.tabs(["💰 Item Economics", "⚙️ Line Utilization", "📦 Supply/Demand Audit", "📈 Strategic P&L"])

with tab1:
    st.subheader("SKU Financial & Operational Profiles")
    eco_data = []
    for p, metrics in FINANCIALS.items():
        unit_margin = metrics['price'] - metrics['cost']
        hourly_profit = unit_margin * metrics['rate']
        
        eco_data.append({
            "Product": p,
            "Unit Price": f"£{metrics['price']:.2f}",
            "Unit Cost": f"£{metrics['cost']:.2f}",
            "Unit Margin": f"£{unit_margin:.2f}",
            "Units / Hour": metrics['rate'],
            "Hourly Profitability": f"£{hourly_profit:,.2f}",
            "Setup Time (Hrs)": metrics['co_time'],
            "Setup Cost": f"£{metrics['co_cost']}"
        })
    st.table(pd.DataFrame(eco_data))
    
    
    
    st.info("💡 **Consultant Insight:** Hourly Profitability is the most critical metric for the MILP. When capacity is constrained, the solver will prioritize products that generate the most £ per hour of machine time.")

with tab2:
    st.subheader(f"Factory Schedule & Line Utilization ({mode})")
    sched_data = []
    total_factory_hrs = 0
    used_factory_hrs = 0
    
    for w in WEEKS:
        row = {"Week": w}
        total_hrs = 0
        for p in PRODUCTS:
            p_hrs = get_val(prod_v[p][w]) / FINANCIALS[p]["rate"]
            s_hrs = get_val(setup_v[p][w]) * FINANCIALS[p]["co_time"]
            total_hrs += (p_hrs + s_hrs)
            
            status = []
            if p_hrs > 0: status.append(f"Prod: {p_hrs:.1f}h")
            if s_hrs > 0: status.append(f"Setup: {s_hrs:.1f}h")
            row[p] = " | ".join(status) if status else "-"
            
        row["Total Hrs Used"] = f"{total_hrs:.1f}"
        row["Utilization"] = f"{(total_hrs / weekly_capacity) * 100:.1f}%"
        sched_data.append(row)
        
        used_factory_hrs += total_hrs
        total_factory_hrs += weekly_capacity
        
    st.dataframe(pd.DataFrame(sched_data), use_container_width=True)

with tab3:
    st.subheader("Transparent Inventory & Fulfillment Audit")
    sku = st.selectbox("Select SKU to Audit", PRODUCTS)
    
    audit_data = []
    total_sku_demand = 0
    total_sku_sold = 0
    
    for i, w in enumerate(WEEKS):
        base = BASE_DEMAND[sku][w]
        roll = get_val(roll_v[sku][WEEKS[i-1]]) if i > 0 else 0
        tot_dem = base + roll
        sold = get_val(sold_v[sku][w])
        
        total_sku_demand += tot_dem
        total_sku_sold += sold
        
        audit_data.append({
            "Week": w,
            "Base Demand": int(base),
            "+ Rollover In": int(roll),
            "= Total Demand": int(tot_dem),
            "Production": int(get_val(prod_v[sku][w])),
            "Units Sold": int(sold),
            "Ending Inventory": int(get_val(inv_v[sku][w])),
            "Missed (Shortage)": int(get_val(short_v[sku][w]))
        })
    st.dataframe(pd.DataFrame(audit_data), use_container_width=True)

with tab4:
    st.subheader(f"Executive Summary: {mode}")
    
    # Calculate global metrics for the KPIs
    rev = sum([get_val(sold_v[p][w]) * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    agility = sum([get_val(inv_v[p][w]) * agility_premium for p in PRODUCTS for w in WEEKS])
    cogs = sum([get_val(prod_v[p][w]) * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    setups = sum([get_val(setup_v[p][w]) * FINANCIALS[p]["co_cost"] for p in PRODUCTS for w in WEEKS])
    holding = sum([get_val(inv_v[p][w]) * holding_cost for p in PRODUCTS for w in WEEKS])
    fines = sum([get_val(short_v[p][w]) * stockout_fine for p in PRODUCTS for w in WEEKS])
    
    net = (rev + agility) - (cogs + setups + holding + fines)
    
    # Aggregate Service Level
    glob_demand = sum([BASE_DEMAND[p][w] + (get_val(roll_v[p][WEEKS[WEEKS.index(w)-1]]) if WEEKS.index(w) > 0 else 0) for p in PRODUCTS for w in WEEKS])
    glob_sold = sum([get_val(sold_v[p][w]) for p in PRODUCTS for w in WEEKS])
    service_level = (glob_sold / glob_demand * 100) if glob_demand > 0 else 0
    overall_utilization = (used_factory_hrs / total_factory_hrs * 100) if total_factory_hrs > 0 else 0
    
    # --- EXECUTIVE KPI DASHBOARD ---
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Net Contribution (£)", f"£{net:,.0f}")
    kpi2.metric("Overall Service Level (%)", f"{service_level:.1f}%")
    kpi3.metric("Factory Utilization (%)", f"{overall_utilization:.1f}%")
    
    st.markdown("---")
    
    pl_df = pd.DataFrame({
        "Financial Line Item": [
            "Gross Sales Revenue", 
            "Agility Premium Earned", 
            "Total COGS", 
            "Changeover Expenses", 
            "Inventory Carrying Costs", 
            "Stockout/Late Fines", 
            "NET CONTRIBUTION"
        ],
        "Amount (£)": [rev, agility, -cogs, -setups, -holding, -fines, net]
    })
    st.table(pl_df)
