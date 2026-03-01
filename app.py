import streamlit as st
import pandas as pd
import numpy as np
import pulp

st.set_page_config(page_title="PwC Value Creation: 6-Month Strategic Optimizer", layout="wide")

st.title("💎 Strategy & Value Creation: 6-Month Operations Masterclass")
st.markdown("### Managing Volatility, Lead Times, and Margin Recovery")

# --- 1. CONFIGURATION: 24-WEEK HORIZON ---
WEEKS = [f"W{i+1}" for i in range(24)]
PRODUCTS = ["Lifting Straps", "Weight Belts", "Knee Sleeves"]

# Financial Profiles per SKU
FINANCIALS = {
    "Lifting Straps": {"price": 25, "cost": 10, "rate": 100, "co_time": 2, "co_cost": 500},
    "Weight Belts":  {"price": 85, "cost": 45, "rate": 30,  "co_time": 8, "co_cost": 2500},
    "Knee Sleeves":  {"price": 45, "cost": 22, "rate": 65,  "co_time": 4, "co_cost": 1200}
}

# --- 2. SIDEBAR CONTROLS ---
with st.sidebar:
    st.header("🏢 Strategic Parameters")
    mode = st.radio("Operating Strategy", ["Legacy: 3.5 FMOC (Static Batching)", "MILP: Optimized (Integrated Value)"])
    
    st.header("⚙️ Capacity & Constraints")
    weekly_capacity = st.slider("Weekly Machine Hours", 40, 168, 120)
    
    st.header("💰 Risk & Agility")
    stockout_penalty = st.slider("Late/Stockout Penalty (£/unit)", 5, 50, 25)
    holding_cost = st.slider("Weekly Inventory Tax (£/unit)", 0.05, 1.0, 0.25)
    agility_fee = st.slider("In-Stock Agility Premium (£)", 0, 15, 8)

# --- 3. DEMAND GENERATION (Seasonality) ---
@st.cache_data
def get_demand_data():
    data = {}
    for p in PRODUCTS:
        # Base demand + a massive "Peak" around weeks 12-18 (Retail Holiday Prep)
        base = np.random.randint(500, 1500, 24)
        peak = [2000 if 11 < i < 19 else 0 for i in range(24)]
        data[p] = {WEEKS[i]: base[i] + peak[i] for i in range(24)}
    return data

DEMAND = get_demand_data()

# --- 4. THE OPTIMIZER ENGINE ---
def run_strategic_optimization(strat):
    prob = pulp.LpProblem("Value_Creation_24W", pulp.LpMinimize)
    
    # Variables
    prod = pulp.LpVariable.dicts("Prod", (PRODUCTS, WEEKS), 0, cat=pulp.LpInteger)
    inv = pulp.LpVariable.dicts("Inv", (PRODUCTS, WEEKS), 0, cat=pulp.LpInteger)
    setup = pulp.LpVariable.dicts("Setup", (PRODUCTS, WEEKS), cat=pulp.LpBinary)
    short = pulp.LpVariable.dicts("Short", (PRODUCTS, WEEKS), 0, cat=pulp.LpInteger)
    
    # Financial Objective
    total_setup = pulp.lpSum([setup[p][w] * FINANCIALS[p]["co_cost"] for p in PRODUCTS for w in WEEKS])
    total_hold = pulp.lpSum([inv[p][w] * holding_cost for p in PRODUCTS for w in WEEKS])
    total_pen = pulp.lpSum([short[p][w] * stockout_penalty for p in PRODUCTS for w in WEEKS])
    # Agility Premium earned on units held into a week
    total_agile = pulp.lpSum([inv[p][w] * agility_fee for p in PRODUCTS for w in WEEKS])
    
    prob += total_setup + total_hold + total_pen - total_agile

    for p in PRODUCTS:
        prev_inv = 0
        for w in WEEKS:
            # Flow: Prev + Prod - Demand = Inv - Short
            prob += prev_inv + prod[p][w] - DEMAND[p][w] == inv[p][w] - short[p][w]
            prob += prod[p][w] <= 20000 * setup[p][w] # Setup Link
            
            # LEGACY LOGIC: Force a huge batch every 4 weeks (to 'save' changeovers)
            if strat == "Legacy: 3.5 FMOC (Static Batching)":
                if WEEKS.index(w) % 4 != 0:
                    prob += prod[p][w] == 0
            
            prev_inv = inv[p][w]

    for w in WEEKS:
        # Machine Hour Constraint
        prob += pulp.lpSum([(prod[p][w]/FINANCIALS[p]["rate"]) + (setup[p][w]*FINANCIALS[p]["co_time"]) for p in PRODUCTS]) <= weekly_capacity

    prob.solve(pulp.PULP_CBC_CMD(msg=0))
    return prod, inv, setup, short

# --- 5. VISUALIZING THE RESULTS ---
prod_v, inv_v, setup_v, short_v = run_strategic_optimization(mode)

tab1, tab2 = st.tabs(["📊 Inventory & Fulfillment Audit", "💰 6-Month Strategic P&L"])

with tab1:
    st.subheader(f"Transparency Report: {mode}")
    selected_sku = st.selectbox("Select Product to Audit", PRODUCTS)
    
    audit_data = []
    for w in WEEKS:
        audit_data.append({
            "Week": w,
            "Demand": DEMAND[selected_sku][w],
            "Production": int(prod_v[selected_sku][w].varValue),
            "On-Hand (End)": int(inv_v[selected_sku][w].varValue),
            "Fulfillment Gap": int(short_v[selected_sku][w].varValue),
            "Capacity Used (Hrs)": round((prod_v[selected_sku][w].varValue/FINANCIALS[selected_sku]["rate"]) + (setup_v[selected_sku][w].varValue*FINANCIALS[selected_sku]["co_time"]), 1)
        })
    st.table(pd.DataFrame(audit_data))

with tab2:
    st.subheader("Consolidated Financial Performance")
    
    # Aggregating all data
    base_rev = sum([DEMAND[p][w] * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    cogs = sum([DEMAND[p][w] * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    
    s_cost = sum([setup_v[p][w].varValue * FINANCIALS[p]["co_cost"] for p in PRODUCTS for w in WEEKS])
    h_cost = sum([inv_v[p][w].varValue * holding_cost for p in PRODUCTS for w in WEEKS])
    p_cost = sum([short_v[p][w].varValue * stockout_penalty for p in PRODUCTS for w in WEEKS])
    a_rev = sum([inv_v[p][w].varValue * agility_fee for p in PRODUCTS for w in WEEKS])
    
    net_contrib = (base_rev + a_rev) - (cogs + s_cost + h_cost + p_cost)
    
    pl_rows = [
        {"Item": "Base Revenue", "Amount": f"£{base_rev:,.0f}"},
        {"Item": "Agility Fees (Bonus)", "Amount": f"£{a_rev:,.0f}"},
        {"Item": "Cost of Goods Sold (COGS)", "Amount": f"-£{cogs:,.0f}"},
        {"Item": "Changeover (Setup) Costs", "Amount": f"-£{s_cost:,.0f}"},
        {"Item": "Inventory Holding Costs", "Amount": f"-£{h_cost:,.0f}"},
        {"Item": "Stockout Penalties (MABD Fines)", "Amount": f"-£{p_cost:,.0f}"},
        {"Item": "NET OPERATING PROFIT", "Amount": f"£{net_contrib:,.0f}"}
    ]
    st.table(pd.DataFrame(pl_rows))

    # Efficiency Comparison
    st.write("### 🏗️ Strategic Interpretation")
    if mode == "Legacy: 3.5 FMOC (Static Batching)":
        st.error("LEGACY RISK: Notice the massive Stockout Penalties during the W12-W18 peak. By 'Batching' to save changeover costs, you physically ran out of time and units.")
    else:
        st.success("MILP VALUE: The solver pre-produced inventory in W1-W10 (the valley) to survive the W12 peak. It 'paid' the holding cost to avoid the £25 stockout penalty.")
