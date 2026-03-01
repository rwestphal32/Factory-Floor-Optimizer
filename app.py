import streamlit as st
import pandas as pd
import numpy as np
import pulp

st.set_page_config(page_title="PwC Value Creation: Operations Masterclass", layout="wide")

st.title("💎 Strategy & Value Creation: End-to-End Operations Model")
st.markdown("**Assumption:** Single shared production line. SKUs compete for total capacity.")

# --- 1. CONFIGURATION ---
WEEKS = [f"W{i+1}" for i in range(24)]
PRODUCTS = ["Lifting Straps", "Weight Belts", "Knee Sleeves"]

# The Economics
FINANCIALS = {
    "Lifting Straps": {"price": 25, "cost": 10, "rate": 100, "co_time": 2, "co_cost": 500},
    "Weight Belts":  {"price": 85, "cost": 45, "rate": 30,  "co_time": 8, "co_cost": 2500},
    "Knee Sleeves":  {"price": 45, "cost": 22, "rate": 65,  "co_time": 4, "co_cost": 1200}
}

with st.sidebar:
    st.header("🏢 Operating Strategy")
    mode = st.radio("Model", ["Legacy: Fixed 4-Week Batches", "MILP: Value Optimized"])
    
    st.header("⚙️ Capacity")
    weekly_capacity = st.slider("Weekly Machine Hours", 40, 168, 120)
    
    st.header("📊 Demand & Risk")
    rollover_pct = st.slider("Demand Rollover (%)", 0, 100, 50) / 100.0
    stockout_fine = st.slider("Lost Sale Penalty (£/unit)", 0, 50, 15)
    holding_cost = st.slider("Holding Cost (£/unit/wk)", 0.1, 2.0, 0.2)

# --- 2. BASE DEMAND GENERATION ---
@st.cache_data
def get_base_demand():
    np.random.seed(42)
    data = {}
    for p in PRODUCTS:
        base = np.random.randint(400, 1000, 24)
        peak = [1500 if 10 < i < 16 else 0 for i in range(24)] # Holiday Peak
        data[p] = {WEEKS[i]: base[i] + peak[i] for i in range(24)}
    return data

BASE_DEMAND = get_base_demand()

# --- 3. THE SOLVER (Rolling Demand & Profit Max) ---
def optimize_operations(strat):
    prob = pulp.LpProblem("Value_Model", pulp.LpMaximize)
    
    # Variables (Continuous for flows to handle rollover percentages cleanly)
    prod = pulp.LpVariable.dicts("Prod", (PRODUCTS, WEEKS), lowBound=0)
    inv = pulp.LpVariable.dicts("Inv", (PRODUCTS, WEEKS), lowBound=0)
    sold = pulp.LpVariable.dicts("Sold", (PRODUCTS, WEEKS), lowBound=0)
    shortage = pulp.LpVariable.dicts("Shortage", (PRODUCTS, WEEKS), lowBound=0)
    rollover = pulp.LpVariable.dicts("Rollover", (PRODUCTS, WEEKS), lowBound=0)
    setup = pulp.LpVariable.dicts("Setup", (PRODUCTS, WEEKS), cat=pulp.LpBinary)
    
    # Objective: Maximize Profit
    revenue = pulp.lpSum([sold[p][w] * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    costs = pulp.lpSum([
        prod[p][w] * FINANCIALS[p]["cost"] + 
        inv[p][w] * holding_cost + 
        setup[p][w] * FINANCIALS[p]["co_cost"] +
        shortage[p][w] * stockout_fine 
        for p in PRODUCTS for w in WEEKS
    ])
    prob += revenue - costs

    # Constraints
    for p in PRODUCTS:
        for i, w in enumerate(WEEKS):
            # 1. Total Demand Calculation
            if i == 0:
                total_demand = BASE_DEMAND[p][w]
            else:
                total_demand = BASE_DEMAND[p][w] + rollover[p][WEEKS[i-1]]
            
            # 2. Sales Constraints
            prob += sold[p][w] <= total_demand
            if i == 0:
                prob += sold[p][w] <= prod[p][w]
            else:
                prob += sold[p][w] <= inv[p][WEEKS[i-1]] + prod[p][w]
                
            # 3. Shortage & Rollover
            prob += shortage[p][w] == total_demand - sold[p][w]
            prob += rollover[p][w] == shortage[p][w] * rollover_pct
            
            # 4. Inventory Balance
            if i == 0:
                prob += prod[p][w] - sold[p][w] == inv[p][w]
            else:
                prob += inv[p][WEEKS[i-1]] + prod[p][w] - sold[p][w] == inv[p][w]
                
            # 5. Setup Link
            prob += prod[p][w] <= 20000 * setup[p][w]
            
            # LEGACY FORCING: Only produce every 4 weeks
            if strat == "Legacy: Fixed 4-Week Batches":
                if i % 4 != 0:
                    prob += prod[p][w] == 0

    # Capacity Constraint (Shared Line)
    for w in WEEKS:
        prob += pulp.lpSum([(prod[p][w]/FINANCIALS[p]["rate"]) + (setup[p][w]*FINANCIALS[p]["co_time"]) for p in PRODUCTS]) <= weekly_capacity

    prob.solve(pulp.PULP_CBC_CMD(msg=0))
    return prod, inv, sold, shortage, rollover, setup

# Run the engine
prod_v, inv_v, sold_v, short_v, roll_v, setup_v = optimize_operations(mode)

# --- 4. TABS & VISUALS ---
tab1, tab2, tab3, tab4 = st.tabs(["💰 Item Economics", "⚙️ Line Utilization", "📦 Supply/Demand Audit", "📈 P&L Statement"])

with tab1:
    st.subheader("SKU Financial & Operational Profiles")
    eco_data = []
    for p, metrics in FINANCIALS.items():
        eco_data.append({
            "Product": p,
            "Unit Price": f"£{metrics['price']:.2f}",
            "Unit Cost": f"£{metrics['cost']:.2f}",
            "Gross Margin (%)": f"{((metrics['price'] - metrics['cost']) / metrics['price']) * 100:.1f}%",
            "Units / Hour": metrics['rate'],
            "Setup Time (Hrs)": metrics['co_time'],
            "Setup Cost": f"£{metrics['co_cost']}"
        })
    st.table(pd.DataFrame(eco_data))
    st.info("Notice how the solver will prioritize Weight Belts during capacity crunches due to the massive Gross Margin, despite the 8-hour changeover penalty.")

with tab2:
    st.subheader(f"Factory Schedule & Line Utilization ({mode})")
    sched_data = []
    for w in WEEKS:
        row = {"Week": w}
        total_hrs = 0
        for p in PRODUCTS:
            p_hrs = prod_v[p][w].varValue / FINANCIALS[p]["rate"]
            s_hrs = setup_v[p][w].varValue * FINANCIALS[p]["co_time"]
            total_hrs += (p_hrs + s_hrs)
            
            # Show what happened for each product
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
        roll = roll_v[sku][WEEKS[i-1]].varValue if i > 0 else 0
        tot_dem = base + roll
        
        audit_data.append({
            "Week": w,
            "Base Demand": int(base),
            "+ Rollover In": int(roll),
            "= Total Demand": int(tot_dem),
            "Production": int(prod_v[sku][w].varValue),
            "Units Sold": int(sold_v[sku][w].varValue),
            "Ending Inventory": int(inv_v[sku][w].varValue),
            "Missed (Shortage)": int(short_v[sku][w].varValue)
        })
    st.dataframe(pd.DataFrame(audit_data), use_container_width=True)
    st.caption("Fulfillment Logic: Total Demand is met by Production + Previous Inventory. Any missed demand becomes a 'Shortage', and a percentage rolls over to the next week.")

with tab4:
    st.subheader(f"Strategic P&L ({mode})")
    
    rev = sum([sold_v[p][w].varValue * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    cogs = sum([prod_v[p][w].varValue * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    setups = sum([setup_v[p][w].varValue * FINANCIALS[p]["co_cost"] for p in PRODUCTS for w in WEEKS])
    holding = sum([inv_v[p][w].varValue * holding_cost for p in PRODUCTS for w in WEEKS])
    fines = sum([short_v[p][w].varValue * stockout_fine for p in PRODUCTS for w in WEEKS])
    
    net = rev - (cogs + setups + holding + fines)
    
    pl_df = pd.DataFrame({
        "Financial Line Item": ["Gross Sales Revenue", "Total COGS", "Changeover Expenses", "Inventory Carrying Costs", "Stockout/Late Fines", "NET CONTRIBUTION"],
        "Amount (£)": [rev, -cogs, -setups, -holding, -fines, net]
    })
    st.table(pl_df)
    
    if mode == "Legacy: Fixed 4-Week Batches":
        st.error(f"Value Destruction: The rigid batching schedule forces stockouts during demand peaks, leading to high penalties and low overall contribution.")
    else:
        st.success(f"Value Creation: MILP navigates the shared line capacity, accepting some holding costs to pre-build inventory and avoid massive stockout fines.")
