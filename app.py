import streamlit as st
import pandas as pd
import pulp

st.set_page_config(page_title="PwC Value Creation Optimizer", layout="wide")

st.title("💎 Supply Chain Value Creation & Margin Optimizer")
st.markdown("### Integrated Decision Support: Procurement, Production & Agility")

# --- 1. THE COMMERCIAL PARAMETERS ---
with st.sidebar:
    st.header("Financial Levers")
    rm_hold_cost = st.slider("Raw Material Holding Cost (£/unit/week)", 0.01, 0.50, 0.05)
    fg_hold_cost = st.slider("Finished Good Holding Cost (£/unit/week)", 0.10, 2.00, 0.50)
    
    st.header("Procurement Logic")
    bulk_threshold = st.number_input("Bulk Discount Threshold (Units)", value=5000)
    bulk_discount_pct = st.slider("Bulk Discount (%)", 0, 30, 15)
    base_material_cost = 10.00
    
    st.header("Commercial Opportunity")
    agility_premium = st.slider("Customer 'Agility Fee' (£/unit sold from stock)", 0.0, 5.0, 2.0)
    st.info("Clients pay extra for 48-hour fulfillment. Is it worth the holding cost?")

# --- 2. DATA SETUP (Multi-Period Demand) ---
weeks = ["Week 1", "Week 2", "Week 3", "Week 4"]
demand = {"Week 1": 1000, "Week 2": 4500, "Week 3": 2000, "Week 4": 6000}
late_penalty = 5.00 # Cost per unit per week late

# --- 3. THE MILP SOLVER ---
def solve_value_chain():
    prob = pulp.LpProblem("Value_Creation", pulp.LpMinimize)

    # VARIABLES
    # Procurement
    buy = pulp.LpVariable.dicts("Buy_Qty", weeks, lowBound=0)
    is_bulk = pulp.LpVariable.dicts("Is_Bulk", weeks, cat=pulp.LpBinary) # 1 if buy > threshold
    
    # Inventory & Production
    produce = pulp.LpVariable.dicts("Produce_Qty", weeks, lowBound=0)
    inv_fg = pulp.LpVariable.dicts("Inv_FG", weeks, lowBound=0)
    inv_rm = pulp.LpVariable.dicts("Inv_RM", weeks, lowBound=0)
    sold_from_stock = pulp.LpVariable.dicts("Agility_Units", weeks, lowBound=0)
    
    # OBJECTIVE: Minimize (Material Cost - Bulk Savings) + Holding Costs + Late Penalties - Agility Premium
    material_spend = pulp.lpSum([buy[t] * base_material_cost for t in weeks])
    bulk_savings = pulp.lpSum([is_bulk[t] * (bulk_threshold * base_material_cost * (bulk_discount_pct/100)) for t in weeks])
    holding_costs = pulp.lpSum([inv_rm[t] * rm_hold_cost + inv_fg[t] * fg_hold_cost for t in weeks])
    revenue_boost = pulp.lpSum([sold_from_stock[t] * agility_premium for t in weeks])
    
    prob += material_spend - bulk_savings + holding_costs - revenue_boost

    # CONSTRAINTS
    prev_rm = 0
    prev_fg = 0
    for t in weeks:
        # 1. Bulk Discount Logic (If Buy > Threshold, is_bulk = 1)
        prob += buy[t] >= bulk_threshold * is_bulk[t]
        
        # 2. Raw Material Balance: Prev RM + Buy - Produce = Current RM
        prob += prev_rm + buy[t] - produce[t] == inv_rm[t]
        
        # 3. Finished Good Balance: Prev FG + Produce - Demand = Current FG
        # Note: We allow FG to be negative to represent "late/backorder" (simplified)
        prob += prev_fg + produce[t] - demand[t] == inv_fg[t]
        
        # 4. Agility Premium: Only units already in stock from 'prev_fg' count for the premium
        prob += sold_from_stock[t] <= prev_fg
        prob += sold_from_stock[t] <= demand[t]
        
        prev_rm = inv_rm[t]
        prev_fg = inv_fg[t]

    prob.solve(pulp.PULP_CBC_CMD(msg=0))
    return prob, buy, produce, inv_fg, sold_from_stock

# --- 4. EXECUTION & DASHBOARD ---
if st.button("🚀 Calculate Strategic Value"):
    prob, buy, produce, inv_fg, sold_from_stock = solve_value_chain()
    
    st.subheader("Results: The Optimized Value Chain")
    
    # Data Table
    res_data = []
    for t in weeks:
        res_data.append({
            "Period": t,
            "Demand": demand[t],
            "Purchased (RM)": buy[t].varValue,
            "Produced (FG)": produce[t].varValue,
            "Closing FG Inv": inv_fg[t].varValue,
            "Agility Units": sold_from_stock[t].varValue
        })
    df_res = pd.DataFrame(res_data)
    st.table(df_res)
    
    # Value Creation Metric
    total_premium = sum([sold_from_stock[t].varValue * agility_premium for t in weeks])
    st.metric("Total 'Agility Premium' Earned", f"£{total_premium:,.2f}")
    
    st.markdown("""
    ### Consulting Insight:
    The model is deciding whether to buy in bulk in Week 1 (to save on material costs) 
    vs. holding that stock (which costs money). It is also calculating if it's worth 
    producing early to 'sit' on inventory so we can charge the **Agility Fee**.
    """)
