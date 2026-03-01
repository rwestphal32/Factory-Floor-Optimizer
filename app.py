import streamlit as st
import pandas as pd
import numpy as np
import pulp
import io

st.set_page_config(page_title="PwC Value Creation: CM Operations", layout="wide")

st.title("🏭 Strategy & Value Creation: CM Operations Optimizer")
st.markdown("**Context:** Contract Manufacturer (CM) optimizing production, labor, and SLA contracts for a Wholesaler client (e.g., Implus).")

# --- 1. DEFAULT DATA ---
DEFAULT_WEEKS = [f"W{i+1}" for i in range(12)]
DEFAULT_PRODUCTS = ["Lifting Straps", "Weight Belts", "Knee Sleeves"]

DEFAULT_FINANCIALS = {
    "Lifting Straps": {"price": 25, "cost": 10, "rate": 60, "co_time": 2, "co_cost": 500, "fine": 10, "premium": 3},
    "Weight Belts":  {"price": 85, "cost": 45, "rate": 25,  "co_time": 8, "co_cost": 2500, "fine": 30, "premium": 15},
    "Knee Sleeves":  {"price": 45, "cost": 22, "rate": 40,  "co_time": 4, "co_cost": 1200, "fine": 15, "premium": 5}
}

@st.cache_data
def get_default_demand():
    np.random.seed(42)
    q_forecast, w_chase = {}, {}
    for p in DEFAULT_PRODUCTS:
        q_forecast[p] = {DEFAULT_WEEKS[i]: np.random.randint(800, 1200) for i in range(12)}
        w_chase[p] = {DEFAULT_WEEKS[i]: (800 if 4 < i < 8 else 0) for i in range(12)}
    return q_forecast, w_chase

DEFAULT_FORECAST, DEFAULT_CHASE = get_default_demand()

def generate_excel_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        eco_df = pd.DataFrame([
            {"Product": "Lifting Straps", "Price": 25, "Materials_Cost": 10, "Rate": 60, "Setup_Time": 2, "Setup_Cost": 500, "SLA_Penalty": 10, "Rush_Surcharge": 3},
            {"Product": "Weight Belts", "Price": 85, "Materials_Cost": 45, "Rate": 25, "Setup_Time": 8, "Setup_Cost": 2500, "SLA_Penalty": 30, "Rush_Surcharge": 15},
            {"Product": "Knee Sleeves", "Price": 45, "Materials_Cost": 22, "Rate": 40, "Setup_Time": 4, "Setup_Cost": 1200, "SLA_Penalty": 15, "Rush_Surcharge": 5}
        ])
        eco_df.to_excel(writer, sheet_name="Economics", index=False)
        
        f_data = {"Product": DEFAULT_PRODUCTS}
        for w in DEFAULT_WEEKS: f_data[w] = [DEFAULT_FORECAST[p][w] for p in DEFAULT_PRODUCTS]
        pd.DataFrame(f_data).to_excel(writer, sheet_name="Quarterly_Forecast", index=False)

        c_data = {"Product": DEFAULT_PRODUCTS}
        for w in DEFAULT_WEEKS: c_data[w] = [DEFAULT_CHASE[p][w] for p in DEFAULT_PRODUCTS]
        pd.DataFrame(c_data).to_excel(writer, sheet_name="Weekly_Chase", index=False)
    return output.getvalue()

# --- 2. SIDEBAR CONTROLS ---
with st.sidebar:
    st.header("📂 1. Data Input")
    st.download_button("📥 Download CM Excel Template", data=generate_excel_template(), file_name="cm_optimization_template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    uploaded_file = st.file_uploader("Upload filled Excel Template", type=["xlsx"])
    
    st.markdown("---")
    st.header("🏢 2. CM Factory Parameters")
    with st.form("control_panel"):
        mode = st.radio("Model", ["Legacy: Min 40-Hour Run", "MILP: Value Optimized"])
        
        st.subheader("Labor & Overhead")
        num_lines = st.slider("Available Production Lines", 1, 3, 2)
        weekly_capacity = st.slider("Max Capacity Per Line (Hrs/Wk)", 80, 168, 120)
        labor_rate = st.slider("Direct Labor Rate (£/Hr)", 10, 50, 20)
        fixed_line_cost = st.slider("Fixed Line Overhead (£/Wk active)", 1000, 10000, 3000)
        
        st.subheader("Global Contracts")
        holding_cost = st.slider("Holding Cost (£/unit/wk)", 0.1, 2.0, 0.2)
        rollover_pct = st.slider("Demand Rollover (%)", 0, 100, 50) / 100.0
        
        submitted = st.form_submit_button("🚀 Run Factory Optimization")

# --- 3. DATA PROCESSING ---
try:
    if uploaded_file is not None:
        eco_df = pd.read_excel(uploaded_file, sheet_name="Economics")
        f_df = pd.read_excel(uploaded_file, sheet_name="Quarterly_Forecast")
        c_df = pd.read_excel(uploaded_file, sheet_name="Weekly_Chase")
        
        PRODUCTS = eco_df["Product"].tolist()
        WEEKS = [col for col in f_df.columns if col != "Product"]
        
        FINANCIALS, UPFRONT_FORECAST, WEEKLY_CHASE = {}, {}, {}
        for _, row in eco_df.iterrows():
            FINANCIALS[row["Product"]] = {
                "price": row["Price"], "cost": row["Materials_Cost"], 
                "rate": row["Rate"], "co_time": row["Setup_Time"], "co_cost": row["Setup_Cost"],
                "fine": row["SLA_Penalty"], "premium": row["Rush_Surcharge"]
            }
        for _, row in f_df.iterrows(): UPFRONT_FORECAST[row["Product"]] = {w: row[w] for w in WEEKS}
        for _, row in c_df.iterrows(): WEEKLY_CHASE[row["Product"]] = {w: row[w] for w in WEEKS}
    else:
        PRODUCTS, WEEKS = DEFAULT_PRODUCTS, DEFAULT_WEEKS
        FINANCIALS, UPFRONT_FORECAST, WEEKLY_CHASE = DEFAULT_FINANCIALS, DEFAULT_FORECAST, DEFAULT_CHASE
except Exception as e:
    st.error(f"❌ Error reading Excel. ({str(e)})")
    st.stop()

LINES = [f"Line {i+1}" for i in range(num_lines)]

# --- 4. THE SOLVER ---
def optimize_operations(strat, lines, capacity_limit):
    prob = pulp.LpProblem("CM_Value_Model", pulp.LpMaximize)
    
    prod_line = pulp.LpVariable.dicts("ProdLine", (PRODUCTS, lines, WEEKS), lowBound=0)
    setup_line = pulp.LpVariable.dicts("SetupLine", (PRODUCTS, lines, WEEKS), cat=pulp.LpBinary)
    line_active = pulp.LpVariable.dicts("LineActive", (lines, WEEKS), cat=pulp.LpBinary) # NEW: Tracks if line is turned on
    
    total_prod = pulp.LpVariable.dicts("TotalProd", (PRODUCTS, WEEKS), lowBound=0)
    inv = pulp.LpVariable.dicts("Inv", (PRODUCTS, WEEKS), lowBound=0)
    sold = pulp.LpVariable.dicts("Sold", (PRODUCTS, WEEKS), lowBound=0)
    shortage = pulp.LpVariable.dicts("Shortage", (PRODUCTS, WEEKS), lowBound=0)
    rollover = pulp.LpVariable.dicts("Rollover", (PRODUCTS, WEEKS), lowBound=0)
    expedited_sold = pulp.LpVariable.dicts("ExpeditedSold", (PRODUCTS, WEEKS), lowBound=0)
    
    for p in PRODUCTS:
        for w in WEEKS:
            prob += total_prod[p][w] == pulp.lpSum([prod_line[p][l][w] for l in lines])
    
    # Financials
    revenue = pulp.lpSum([sold[p][w] * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    expedite_rev = pulp.lpSum([expedited_sold[p][w] * FINANCIALS[p]["premium"] for p in PRODUCTS for w in WEEKS])
    
    materials_cost = pulp.lpSum([total_prod[p][w] * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    inv_cost = pulp.lpSum([inv[p][w] * holding_cost for p in PRODUCTS for w in WEEKS])
    sla_fines = pulp.lpSum([shortage[p][w] * FINANCIALS[p]["fine"] for p in PRODUCTS for w in WEEKS])
    setup_fees = pulp.lpSum([setup_line[p][l][w] * FINANCIALS[p]["co_cost"] for p in PRODUCTS for l in lines for w in WEEKS])
    
    # NEW: Direct Labor & Fixed Overhead
    total_hours = pulp.lpSum([(prod_line[p][l][w]/FINANCIALS[p]["rate"]) + (setup_line[p][l][w]*FINANCIALS[p]["co_time"]) for p in PRODUCTS for l in lines for w in WEEKS])
    labor_cost = total_hours * labor_rate
    overhead_cost = pulp.lpSum([line_active[l][w] * fixed_line_cost for l in lines for w in WEEKS])
    
    prob += revenue + expedite_rev - (materials_cost + inv_cost + sla_fines + setup_fees + labor_cost + overhead_cost)

    for p in PRODUCTS:
        for i, w in enumerate(WEEKS):
            if i == 0: total_demand = UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w]
            else: total_demand = UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w] + rollover[p][WEEKS[i-1]]
            
            prob += sold[p][w] <= total_demand
            if i == 0: 
                prob += sold[p][w] <= total_prod[p][w]
                prob += total_prod[p][w] - sold[p][w] == inv[p][w]
            else: 
                prob += sold[p][w] <= inv[p][WEEKS[i-1]] + total_prod[p][w]
                prob += inv[p][WEEKS[i-1]] + total_prod[p][w] - sold[p][w] == inv[p][w]
                
            prob += expedited_sold[p][w] <= WEEKLY_CHASE[p][w]
            prob += expedited_sold[p][w] <= sold[p][w]
            prob += shortage[p][w] == total_demand - sold[p][w]
            prob += rollover[p][w] == shortage[p][w] * rollover_pct
            
            for l in lines:
                max_prod = capacity_limit * FINANCIALS[p]["rate"]
                prob += prod_line[p][l][w] <= max_prod * setup_line[p][l][w]
                # Force line_active to 1 if any setup occurs on that line
                prob += setup_line[p][l][w] <= line_active[l][w]

                if strat == "Legacy: Min 40-Hour Run":
                    min_batch_qty = 40 * FINANCIALS[p]["rate"]
                    prob += prod_line[p][l][w] >= min_batch_qty * setup_line[p][l][w]

    for l in lines:
        for w in WEEKS:
            prob += pulp.lpSum([(prod_line[p][l][w]/FINANCIALS[p]["rate"]) + (setup_line[p][l][w]*FINANCIALS[p]["co_time"]) for p in PRODUCTS]) <= capacity_limit * line_active[l][w]

    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=15))
    return prob, total_prod, inv, sold, shortage, rollover, prod_line, setup_line, expedited_sold, line_active

# --- 5. EXECUTION ---
def get_val(var):
    return var.varValue if var.varValue is not None else 0

with st.spinner(f"Optimizing {len(PRODUCTS)} SKUs across {num_lines} CM Lines..."):
    prob_status, prod_v, inv_v, sold_v, short_v, roll_v, p_line_v, s_line_v, expedite_v, active_v = optimize_operations(mode, LINES, weekly_capacity)

if pulp.LpStatus[prob_status.status] != 'Optimal':
    st.warning(f"⚠️ **Optimality Gap:** Status: {pulp.LpStatus[prob_status.status]}. Scale hit time limit; displaying best margin.")

# --- 6. VISUALS ---
tab1, tab2, tab3 = st.tabs(["📈 CM P&L Statement", "⚙️ Factory Routing", "📦 CM to Wholesaler Audit"])

with tab1:
    st.subheader(f"Contract Manufacturer P&L ({mode})")
    
    rev = sum([get_val(sold_v[p][w]) * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    expedite_rev = sum([get_val(expedite_v[p][w]) * FINANCIALS[p]["premium"] for p in PRODUCTS for w in WEEKS])
    
    mat_cost = sum([get_val(prod_v[p][w]) * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    setups = sum([get_val(s_line_v[p][l][w]) * FINANCIALS[p]["co_cost"] for p in PRODUCTS for l in LINES for w in WEEKS])
    holding = sum([get_val(inv_v[p][w]) * holding_cost for p in PRODUCTS for w in WEEKS])
    fines = sum([get_val(short_v[p][w]) * FINANCIALS[p]["fine"] for p in PRODUCTS for w in WEEKS])
    
    # Labor & Overhead Calculation
    tot_hrs = sum([(get_val(p_line_v[p][l][w]) / FINANCIALS[p]["rate"]) + (get_val(s_line_v[p][l][w]) * FINANCIALS[p]["co_time"]) for p in PRODUCTS for l in LINES for w in WEEKS])
    labor = tot_hrs * labor_rate
    overhead = sum([get_val(active_v[l][w]) * fixed_line_cost for l in LINES for w in WEEKS])
    
    net = (rev + expedite_rev) - (mat_cost + setups + holding + fines + labor + overhead)
    
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("CM Net Contribution (£)", f"£{net:,.0f}")
    kpi2.metric("Total Labor & Overhead (£)", f"£{(labor + overhead):,.0f}")
    
    active_line_weeks = sum([get_val(active_v[l][w]) for l in LINES for w in WEEKS])
    total_line_weeks = len(LINES) * len(WEEKS)
    kpi3.metric("Line Activation Rate", f"{(active_line_weeks/total_line_weeks)*100:.0f}%")
    
    st.markdown("---")
    pl_df = pd.DataFrame({
        "Financial Line Item": ["Wholesale Revenue", "Rush Order Surcharges", "Materials Cost (COGS)", "Direct Labor (£/Hr)", "Fixed Line Overhead (£/Wk)", "Setup Fees", "Holding Costs", "SLA Shortage Penalties", "NET CONTRIBUTION"],
        "Amount (£)": [rev, expedite_rev, -mat_cost, -labor, -overhead, -setups, -holding, -fines, net]
    })
    st.table(pl_df)

with tab2:
    st.subheader(f"CM Multi-Line Routing ({mode})")
    for l in LINES:
        st.markdown(f"#### 🏭 {l}")
        l_data = []
        for w in WEEKS:
            if get_val(active_v[l][w]) == 0:
                l_data.append({"Week": w, "Status": "🛑 LINE SHUT DOWN (Overhead Saved)"})
                continue
                
            row = {"Week": w, "Status": "🟢 ACTIVE"}
            total_l_hrs = 0
            for p in PRODUCTS:
                p_hrs = get_val(p_line_v[p][l][w]) / FINANCIALS[p]["rate"]
                s_hrs = get_val(s_line_v[p][l][w]) * FINANCIALS[p]["co_time"]
                total_l_hrs += (p_hrs + s_hrs)
                
                status = []
                if p_hrs > 0: status.append(f"Prod: {p_hrs:.1f}h")
                if s_hrs > 0: status.append(f"Setup: {s_hrs:.1f}h")
                row[p] = " | ".join(status) if status else "-"
                
            row["Hrs Used"] = f"{total_l_hrs:.1f} / {weekly_capacity}"
            row["Labor Cost"] = f"£{total_l_hrs * labor_rate:,.0f}"
            l_data.append(row)
        st.dataframe(pd.DataFrame(l_data), use_container_width=True)

with tab3:
    st.subheader("CM to Wholesaler Delivery Audit")
    sku = st.selectbox("Select SKU to Audit", PRODUCTS)
    
    audit_data = []
    for i, w in enumerate(WEEKS):
        forecast = UPFRONT_FORECAST[sku][w]
        chase = WEEKLY_CHASE[sku][w]
        roll = get_val(roll_v[sku][WEEKS[i-1]]) if i > 0 else 0
        tot_dem = forecast + chase + roll
        
        audit_data.append({
            "Week": w,
            "Wholesaler Forecast": int(forecast),
            "+ Rush Orders": int(chase),
            "= Total Implus Demand": int(tot_dem),
            "CM Total Production": int(get_val(prod_v[sku][w])),
            "Units Delivered": int(get_val(sold_v[sku][w])),
            "Rush Surcharge Earned": int(get_val(expedite_v[sku][w])),
            "CM Warehouse Inv": int(get_val(inv_v[sku][w])),
            "SLA Missed (Penalty)": int(get_val(short_v[sku][w]))
        })
    st.dataframe(pd.DataFrame(audit_data), use_container_width=True)
