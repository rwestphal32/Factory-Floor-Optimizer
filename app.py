import streamlit as st
import pandas as pd
import numpy as np
import pulp

st.set_page_config(page_title="PwC Value Creation: Strategic Ops", layout="wide")

st.title("💎 Strategy & Value Creation: The 'Harbinger' Case Study")
st.markdown("### Managing 6-Month Lead Times vs. Stochastic Retail Demand")

# --- 1. GLOBAL SETTINGS ---
with st.sidebar:
    st.header("1. Financial Parameters")
    unit_price = st.number_input("Wholesale Unit Price (£)", value=25.0)
    unit_cost = st.number_input("Manufacturing Cost (£)", value=10.0)
    
    st.header("2. Logistics & Agility")
    ocean_cost = 1.50 # Sunk cost from 6 months ago
    air_cost = st.number_input("Emergency Air Freight (£/unit)", value=8.50)
    stockout_penalty = st.number_input("Stockout Penalty (Lost GP + Fines)", value=15.0)
    holding_cost = st.slider("Weekly Holding Cost (£/unit)", 0.1, 1.0, 0.3)
    
    agility_premium = st.slider("Agility Premium (£/unit)", 0.0, 10.0, 4.0)
    st.info("Premium earned ONLY when fulfilling a 'Spike' from existing safety stock.")

# --- 2. THE DATA: RE-CREATING THE "NIGHTMARE" ---
# Demand = Base Forecast + Random Spike
WEEKS = ["Week 1", "Week 2", "Week 3", "Week 4"]
base_demand = [1500, 1500, 1500, 1500]

# Simulate a "DSG Spike" in Week 2 and 4
np.random.seed(42)
spikes = [0, 2000, 0, 1200]
actual_demand = [base_demand[i] + spikes[i] for i in range(4)]

# --- TAB 1: THE TACTICAL AUDIT (FINANCIALS) ---
tab1, tab2 = st.tabs(["📊 Tactical P&L: The Buffer Strategy", "🏭 Factory Schedule: Future Orders"])

with tab1:
    st.subheader("Financial Impact of Inventory Resilience")
    
    # User Control: How much Safety Stock did we decide to bring in 6 months ago?
    ss_target = st.select_slider("Strategy: Target Safety Stock (Units)", options=[0, 500, 1500, 3000, 5000], value=1500)

    # Simulation Logic
    inv = ss_target
    results = []
    
    for i in range(4):
        d = actual_demand[i]
        s = spikes[i]
        
        # Fulfillment
        from_stock = min(d, inv)
        shortage = d - from_stock
        
        # Agility Revenue (Earned if we cover the spike from stock)
        agility_units = min(s, from_stock) if s > 0 else 0
        agility_rev = agility_units * agility_premium
        
        # Mitigation: Do we air-freight to stop the stockout?
        # Let's assume we air-freight 60% of any shortage
        air_units = shortage * 0.6
        lost_sales = shortage - air_units
        
        # Financials
        revenue = (from_stock + air_units) * unit_price
        cogs = (from_stock + air_units) * unit_cost
        freight_hit = (from_stock * ocean_cost) + (air_units * air_cost)
        holding_hit = inv * holding_cost
        penalty_hit = lost_sales * stockout_penalty
        
        margin = (revenue + agility_rev) - (cogs + freight_hit + holding_hit + penalty_hit)
        
        results.append({
            "Week": WEEKS[i],
            "Total Demand": d,
            "Fulfillment Rate (%)": round((from_stock + air_units)/d * 100, 1),
            "Agility Units": int(agility_units),
            "Air Freight": int(air_units),
            "Lost Sales": int(lost_sales),
            "Net Margin (£)": round(margin, 2)
        })
        
        inv = max(0, inv - from_stock)

    st.table(pd.DataFrame(results))
    
    total_margin = sum([r["Net Margin (£)"] for r in results])
    st.metric("Total Net Contribution", f"£{total_margin:,.2f}")
    
    

# --- TAB 2: THE FACTORY SCHEDULE (FUTURE) ---
with tab2:
    st.subheader("6-Month Planning: Optimizing Future Batch Sizes")
    st.markdown("Placing orders *today* for the next 4-week cycle arriving in 6 months.")
    
    # Setup MILP for Batching
    prob = pulp.LpProblem("Batch_Optimizer", pulp.LpMinimize)
    
    # Variables
    prod = pulp.LpVariable.dicts("Prod", WEEKS, lowBound=0, cat=pulp.LpInteger)
    setup = pulp.LpVariable.dicts("Setup", WEEKS, cat=pulp.LpBinary)
    inv_f = pulp.LpVariable.dicts("Future_Inv", WEEKS, lowBound=0, cat=pulp.LpInteger)
    
    setup_cost = 2000 # Cost to switch the line
    
    # Objective: Min (Holding + Setup)
    prob += pulp.lpSum([inv_f[w] * holding_cost + setup[w] * setup_cost for w in WEEKS])
    
    # Constraints
    prev_f = 0
    for w in WEEKS:
        prob += prev_f + prod[w] - actual_demand[WEEKS.index(w)] == inv_f[w]
        prob += prod[w] <= 20000 * setup[w] # Setup Link
        prev_f = inv_f[w]
        
    prob.solve(pulp.PULP_CBC_CMD(msg=0))
    
    # Schedule Display
    sched_data = []
    for w in WEEKS:
        sched_data.append({"Week": w, "Production Order (Units)": int(prod[w].varValue), "Setup Required": "Yes" if setup[w].varValue > 0 else "No"})
    
    st.write("### ⏱️ Optimized Production Run")
    st.table(pd.DataFrame(sched_data))
    
    st.info("""
    **PwC Consultant Insight:** In the Tactical Audit (Tab 1), look at how a low safety stock target 
    leads to negative margins because the Air Freight and Stockout Penalties 'bleed' the P&L dry. 
    This proves that 'Lean' (zero stock) is actually the most expensive way to run a 6-month supply chain.
    """)
