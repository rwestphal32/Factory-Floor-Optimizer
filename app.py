import streamlit as st
import pandas as pd
import pulp

st.set_page_config(page_title="PwC Value Creation: Harbinger Case", layout="wide")

st.title("💎 Strategy & Value Creation: Multi-SKU Optimizer")
st.markdown("### Proving the Margin Delta: Legacy Batching vs. MILP Optimization")

# --- 1. CONFIGURATION & INPUTS ---
PRODUCTS = ["Lifting Straps", "Weight Belts", "Knee Sleeves"]
WEEKS = ["Week 1", "Week 2", "Week 3", "Week 4"]

# The "Nightmare" Demand: Base + Random Spikes
DEMAND = {
    "Lifting Straps": {"Week 1": 1500, "Week 2": 3500, "Week 3": 1500, "Week 4": 4000}, # Spike W2/W4
    "Weight Belts": {"Week 1": 800, "Week 2": 800, "Week 3": 2500, "Week 4": 800},   # Spike W3
    "Knee Sleeves": {"Week 1": 1000, "Week 2": 1000, "Week 3": 1000, "Week 4": 1000} # Flat
}

# Production Specs
RATE = {"Lifting Straps": 100, "Weight Belts": 40, "Knee Sleeves": 60}
UNIT_PRICE = {"Lifting Straps": 25, "Weight Belts": 60, "Knee Sleeves": 40}
UNIT_COST = {"Lifting Straps": 12, "Weight Belts": 35, "Knee Sleeves": 20}

with st.sidebar:
    st.header("🏢 Strategy Toggle")
    # THE "OLD WAY" vs "NEW WAY"
    mode = st.radio("Operating Model", ["Legacy (Fixed Large Batches)", "PwC Value-Optimized (MILP)"])
    
    st.header("⚙️ Factory Constraints")
    weekly_hours = st.slider("Weekly Capacity (Machine Hours)", 40, 168, 100)
    co_time = st.number_input("Changeover Downtime (Hours)", value=6)
    co_cost = st.number_input("Changeover Setup Cost (£)", value=1500)
    
    st.header("💰 Financial Levers")
    hold_cost = st.slider("Holding Cost (£/unit/week)", 0.1, 2.0, 0.4)
    agility_fee = st.slider("Agility Premium (£/unit)", 0, 10, 5)
    penalty = st.slider("Stockout Fine/Penalty (£/unit)", 5, 50, 20)

# --- 2. THE SOLVER ENGINE ---
def run_pwc_model(strategy):
    prob = pulp.LpProblem("ValueCreation", pulp.LpMinimize)
    
    # Variables
    prod = pulp.LpVariable.dicts("Prod", (PRODUCTS, WEEKS), 0, cat=pulp.LpInteger)
    inv = pulp.LpVariable.dicts("Inv", (PRODUCTS, WEEKS), 0, cat=pulp.LpInteger)
    setup = pulp.LpVariable.dicts("Setup", (PRODUCTS, WEEKS), cat=pulp.LpBinary)
    short = pulp.LpVariable.dicts("Shortage", (PRODUCTS, WEEKS), 0, cat=pulp.LpInteger)
    
    # Objective: Minimize (Setup + Holding + Penalties) - (Agility Revenue)
    s_cost = pulp.lpSum([setup[p][w] * co_cost for p in PRODUCTS for w in WEEKS])
    h_cost = pulp.lpSum([inv[p][w] * hold_cost for p in PRODUCTS for w in WEEKS])
    p_cost = pulp.lpSum([short[p][w] * penalty for p in PRODUCTS for w in WEEKS])
    
    # Agility Revenue: 
    # Simplified: If we carry stock into a week with a spike, we charge the fee
    a_rev = pulp.lpSum([inv[p][w] * agility_fee for p in PRODUCTS for w in WEEKS])
    
    prob += s_cost + h_cost + p_cost - a_rev

    for p in PRODUCTS:
        prev_inv = 0
        total_needed = sum(DEMAND[p].values())
        for w in WEEKS:
            # Inventory Balance: prev + prod - demand = inv - shortage
            prob += prev_inv + prod[p][w] - DEMAND[p][w] == inv[p][w] - short[p][w]
            
            # Setup Link
            prob += prod[p][w] <= 10000 * setup[p][w]
            
            # LEGACY MODE FORCING: 
            # In legacy, you run everything for that product in its first big demand week
            if strategy == "Legacy (Fixed Large Batches)":
                if w == "Week 1":
                    prob += prod[p][w] == total_needed
                else:
                    prob += prod[p][w] == 0
            
            prev_inv = inv[p][w]

    # Capacity Constraint (Shared across all 3 SKUs)
    for w in WEEKS:
        prob += pulp.lpSum([(prod[p][w]/RATE[p]) + (setup[p][w]*co_time) for p in PRODUCTS]) <= weekly_hours

    prob.solve(pulp.PULP_CBC_CMD(msg=0))
    return prod, inv, setup, short

# --- 3. EXECUTION & DISPLAY ---
prod_v, inv_v, setup_v, short_v = run_pwc_model(mode)

col1, col2 = st.columns([2, 1])

with col1:
    st.subheader(f"📊 {mode}: Production Schedule")
    # Production Table
    p_rows = []
    for p in PRODUCTS:
        row = {"Product": p}
        for w in WEEKS: row[w] = int(prod_v[p][w].varValue)
        p_rows.append(row)
    st.table(pd.DataFrame(p_rows))

    st.subheader("🚨 Stockout & Fulfillment Audit")
    s_rows = []
    for p in PRODUCTS:
        row = {"Product": p}
        for w in WEEKS: row[w] = int(short_v[p][w].varValue)
        s_rows.append(row)
    st.table(pd.DataFrame(s_rows))

with col2:
    st.subheader("💰 The Financial Delta")
    
    # Math for P&L
    total_rev = sum([DEMAND[p][w] * UNIT_PRICE[p] for p in PRODUCTS for w in WEEKS])
    total_cogs = sum([DEMAND[p][w] * UNIT_COST[p] for p in PRODUCTS for w in WEEKS])
    total_setup = sum([setup_v[p][w].varValue * co_cost for p in PRODUCTS for w in WEEKS])
    total_hold = sum([inv_v[p][w].varValue * hold_cost for p in PRODUCTS for w in WEEKS])
    total_pen = sum([short_v[p][w].varValue * penalty for p in PRODUCTS for w in WEEKS])
    total_agility = sum([inv_v[p][w].varValue * agility_fee for p in PRODUCTS for w in WEEKS])
    
    net_contrib = (total_rev + total_agility) - (total_cogs + total_setup + total_hold + total_pen)
    
    pl_data = {
        "Line Item": ["Base Revenue", "Agility Fees", "COGS", "Setup/Changeover", "Holding Costs", "Stockout Penalties", "NET MARGIN"],
        "Value (£)": [total_rev, total_agility, -total_cogs, -total_setup, -total_hold, -total_pen, net_contrib]
    }
    st.table(pd.DataFrame(pl_data))

    # CAPACITY CHECK
    st.subheader("⚙️ Capacity Utilization")
    for w in WEEKS:
        hrs = sum([(prod_v[p][w].varValue/RATE[p]) + (setup_v[p][w].varValue*co_time) for p in PRODUCTS])
        st.write(f"{w}: {hrs:.1f} / {weekly_hours} hrs")
        st.progress(min(1.0, hrs/weekly_hours))
