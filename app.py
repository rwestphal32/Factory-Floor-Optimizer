import streamlit as st
import pandas as pd
import numpy as np
import pulp

st.set_page_config(page_title="PwC Value Creation: Operations Masterclass", layout="wide")

st.title("💎 Strategy & Value Creation: High-Utilization Optimizer")
st.markdown("**Assumption:** Single shared production line running at ~80% baseline utilization.")

# --- 1. CONFIGURATION ---
WEEKS = [f"W{i+1}" for i in range(24)]
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
        
        st.header("📊 Demand & Risk")
        rollover_pct = st.slider("Demand Rollover (%)", 0, 100, 50) / 100.0
        stockout_fine = st.slider("Lost Sale Penalty (£/unit)", 0, 50, 15)
        holding_cost = st.slider("Holding Cost (£/unit/wk)", 0.1, 2.0, 0.2)
        
        submitted = st.form_submit_button("🚀 Run Optimization")

# --- 3. BASE DEMAND GENERATION ---
@st.cache_data
def get_base_demand():
    np.random.seed(42)
    data = {}
    for p in PRODUCTS:
        base = np.random.randint(800, 1200, 24)
        peak = [800 if 10 < i < 16 else 0 for i in range(24)] # Holiday Peak
        data[p] = {WEEKS[i]: base[i] + peak[i] for i in range(24)}
    return data

BASE_DEMAND = get_base_demand()

# --- 4. THE SOLVER ---
def optimize_operations(strat):
    prob = pulp.LpProblem("Value_Model", pulp.LpMaximize)
    
    prod = pulp.LpVariable.dicts("Prod", (PRODUCTS, WEEKS), lowBound=0)
    inv = pulp.LpVariable.dicts("Inv", (PRODUCTS, WEEKS), lowBound=0)
    sold = pulp.LpVariable.dicts("Sold", (PRODUCTS, WEEKS), lowBound=0)
    shortage = pulp.LpVariable.dicts("Shortage", (PRODUCTS, WEEKS), lowBound=0)
    rollover = pulp.LpVariable.dicts("Rollover", (PRODUCTS, WEEKS), lowBound=0)
    setup = pulp.LpVariable.dicts("Setup", (PRODUCTS, WEEKS), cat=pulp.LpBinary)
    
    revenue = pulp.lpSum([sold[p][w] * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    costs = pulp.lpSum([
        prod[p][w] * FINANCIALS[p]["cost"] + 
        inv[p][w] * holding_cost + 
        setup[p][w] * FINANCIALS[p]["co_cost"] +
        shortage[p][w] * stockout_fine 
        for p in PRODUCTS for w in WEEKS
    ])
    prob += revenue - costs

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
                
            prob += prod[p][w] <= 20000 * setup[p][w]

    for w in WEEKS:
        # Physical Capacity Constraint
        prob += pulp.lpSum([(prod[p][w]/FINANCIALS[p]["rate"]) + (setup[p][w]*FINANCIALS[p]["co_time"]) for p in PRODUCTS]) <= weekly_capacity
        
        # Legacy Logic
        if strat == "Legacy: Max 1 Setup/Week (Big Batches)":
            prob += pulp.lpSum([setup[p][w] for p in PRODUCTS]) <= 1

    # THE FIX: Added a 10-second time limit so it never hangs
    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=10))
    return prob, prod, inv, sold, short, roll, setup

# Helper function to prevent crashes if solver hits time limit
def get_val(var):
    return var.varValue if var.varValue is not None else 0

# --- 5. EXECUTION ---
# UI Spinner so you know it's working
with st.spinner("Crunching the MILP Matrix (This may take up to 10 seconds)..."):
    prob_status, prod_v, inv_v, sold_v, short_v, roll_v, setup_v = optimize_operations(mode)

# Warn if the time limit was hit
if pulp.LpStatus[prob_status.status] != 'Optimal':
    st.warning(f"Solver Status: {pulp.LpStatus[prob_status.status]}. The problem is highly constrained. Displaying the best solution found within the time limit.")

# --- 6. TABS & VISUALS ---
tab1, tab2, tab3, tab4 = st.tabs(["💰 Item Economics", "⚙️ Line Utilization", "📦 Supply/Demand Audit", "📈 P&L Statement"])

with tab1:
    st.subheader("SKU Financial & Operational Profiles")
    eco_data = []
    for p, metrics in FINANCIALS.items():
        eco_data.append({
            "Product": p,
            "Unit Margin": f"£{metrics['price'] - metrics['cost']:.2f}",
            "Units / Hour": metrics['rate'],
            "Setup Time (Hrs)": metrics['co_time'],
            "Setup Cost": f"£{metrics['co_cost']}"
        })
    st.table(pd.DataFrame(eco_data))

with tab2:
    st.subheader(f"Factory Schedule & Line Utilization ({mode})")
    sched_data = []
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
        
    st.dataframe(pd.DataFrame(sched_data), use_container_width=True)

with tab3:
    st.subheader("Transparent Inventory & Fulfillment Audit")
    sku = st.selectbox("Select SKU to Audit", PRODUCTS)
    
    audit_data = []
    for i, w in enumerate(WEEKS):
        base = BASE_DEMAND[sku][w]
        roll = get_val(roll_v[sku][WEEKS[i-1]]) if i > 0 else 0
        tot_dem = base + roll
        
        audit_data.append({
            "Week": w,
            "Base Demand": int(base),
            "+ Rollover In": int(roll),
            "= Total Demand": int(tot_dem),
            "Production": int(get_val(prod_v[sku][w])),
            "Units Sold": int(get_val(sold_v[sku][w])),
            "Ending Inventory": int(get_val(inv_v[sku][w])),
            "Missed (Shortage)": int(get_val(short_v[sku][w]))
        })
    st.dataframe(pd.DataFrame(audit_data), use_container_width=True)

with tab4:
    st.subheader(f"Strategic P&L ({mode})")
    
    rev = sum([get_val(sold_v[p][w]) * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    cogs = sum([get_val(prod_v[p][w]) * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    setups = sum([get_val(setup_v[p][w]) * FINANCIALS[p]["co_cost"] for p in PRODUCTS for w in WEEKS])
    holding = sum([get_val(inv_v[p][w]) * holding_cost for p in PRODUCTS for w in WEEKS])
    fines = sum([get_val(short_v[p][w]) * stockout_fine for p in PRODUCTS for w in WEEKS])
    
    net = rev - (cogs + setups + holding + fines)
    
    pl_df = pd.DataFrame({
        "Financial Line Item": ["Gross Sales Revenue", "Total COGS", "Changeover Expenses", "Inventory Carrying Costs", "Stockout/Late Fines", "NET CONTRIBUTION"],
        "Amount (£)": [rev, -cogs, -setups, -holding, -fines, net]
    })
    st.table(pl_df)
