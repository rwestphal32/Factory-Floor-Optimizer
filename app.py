import streamlit as st
import pandas as pd
import numpy as np
import pulp
import io

st.set_page_config(page_title="PwC Value Creation: Digital Twin", layout="wide")

st.title("🏭 Strategy & Value Creation: Factory Digital Twin")
st.markdown("**Architecture:** Parallel routing highlighting the *Efficiency vs. Flexibility* trade-off for Functional vs. Volatile products.")

# --- 1. DIGITAL TWIN DEMO DATA (Fisher Framework) ---
DEFAULT_WEEKS = [f"W{i+1}" for i in range(12)]
DEFAULT_PRODUCTS = ["Lifting Straps (Stable)", "Weight Belts (Volatile)", "Knee Sleeves (Volatile)"]

# Complex Product-Machine Relationship Matrix
DEFAULT_FINANCIALS = {
    "Lifting Straps (Stable)": {
        "price": 25, "cost": 10, "fine": 10, "premium": 3,
        "L1": {"rate": 80, "time": 2, "cost": 500},    # Native to Dedicated L1 (Highly Efficient)
        "L2": {"rate": 40, "time": 4, "cost": 800},    # Compatible with Flex lines but slower
        "L3": {"rate": 40, "time": 4, "cost": 800}
    },
    "Weight Belts (Volatile)": {
        "price": 85, "cost": 45, "fine": 30, "premium": 15,
        "L1": {"rate": 20, "time": 24, "cost": 10000}, # MASSIVE PENALTY to re-tool the dedicated line
        "L2": {"rate": 30, "time": 6, "cost": 1500},   # Native to Flex lines
        "L3": {"rate": 30, "time": 6, "cost": 1500}
    },
    "Knee Sleeves (Volatile)": {
        "price": 45, "cost": 22, "fine": 15, "premium": 5,
        "L1": {"rate": 20, "time": 24, "cost": 10000}, # MASSIVE PENALTY to re-tool the dedicated line
        "L2": {"rate": 40, "time": 4, "cost": 1000},   # Native to Flex lines
        "L3": {"rate": 40, "time": 4, "cost": 1000}
    }
}

@st.cache_data
def get_default_demand():
    np.random.seed(42)
    q_forecast, w_chase = {}, {}
    # Stable product gets flat demand. Volatile gets spikes.
    for p in DEFAULT_PRODUCTS:
        if "Stable" in p:
            q_forecast[p] = {DEFAULT_WEEKS[i]: 3500 for i in range(12)}
            w_chase[p] = {DEFAULT_WEEKS[i]: 0 for i in range(12)}
        else:
            base = 1200 if "Belts" in p else 2000
            spike = 1500 if "Belts" in p else 2000
            q_forecast[p] = {DEFAULT_WEEKS[i]: np.random.randint(int(base*0.9), int(base*1.1)) for i in range(12)}
            w_chase[p] = {DEFAULT_WEEKS[i]: (spike if 4 < i < 8 else 0) for i in range(12)}
    return q_forecast, w_chase

DEFAULT_FORECAST, DEFAULT_CHASE = get_default_demand()

def generate_excel_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        rows = []
        for p, data in DEFAULT_FINANCIALS.items():
            row = {"Product": p, "Price": data["price"], "Cost": data["cost"], "SLA_Fine": data["fine"], "Rush_Premium": data["premium"]}
            for l in ["L1", "L2", "L3"]:
                row[f"{l}_Rate"] = data[l]["rate"]
                row[f"{l}_Setup_Time"] = data[l]["time"]
                row[f"{l}_Setup_Cost"] = data[l]["cost"]
            rows.append(row)
        pd.DataFrame(rows).to_excel(writer, sheet_name="Economics_Matrix", index=False)
        
        f_data = {"Product": DEFAULT_PRODUCTS}
        for w in DEFAULT_WEEKS: f_data[w] = [DEFAULT_FORECAST[p][w] for p in DEFAULT_PRODUCTS]
        pd.DataFrame(f_data).to_excel(writer, sheet_name="Quarterly_Forecast", index=False)

        c_data = {"Product": DEFAULT_PRODUCTS}
        for w in DEFAULT_WEEKS: c_data[w] = [DEFAULT_CHASE[p][w] for p in DEFAULT_PRODUCTS]
        pd.DataFrame(c_data).to_excel(writer, sheet_name="Weekly_Chase", index=False)
    return output.getvalue()

# --- 2. SIDEBAR CONTROLS ---
with st.sidebar:
    st.header("📂 1. Digital Twin Data")
    st.download_button("📥 Download Twin Excel Template", data=generate_excel_template(), file_name="digital_twin_matrix.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    uploaded_file = st.file_uploader("Upload Configured Excel", type=["xlsx"])
    st.caption("Upload manipulating speeds, setup times, and setup costs per individual line.")
    
    st.markdown("---")
    st.header("🏢 2. Exec Controls")
    with st.form("control_panel"):
        mode = st.radio("S&OP Policy", ["Legacy: Min 40-Hour Run", "MILP: Value Optimized"])
        num_lines = st.slider("Active Production Lines", 1, 3, 3)
        weekly_capacity = st.slider("Weekly Capacity Per Line (Hrs)", 80, 168, 120)
        labor_rate = st.slider("Direct Labor Rate (£/Hr)", 10, 50, 20)
        fixed_line_cost = st.slider("Fixed Line Overhead (£/Wk active)", 1000, 10000, 3000)
        holding_cost = st.slider("Holding Cost (£/unit/wk)", 0.1, 2.0, 0.2)
        rollover_pct = st.slider("Demand Rollover (%)", 0, 100, 50) / 100.0
        submitted = st.form_submit_button("🚀 Run Twin Simulation")

# --- 3. DATA PROCESSING ---
try:
    if uploaded_file is not None:
        eco_df = pd.read_excel(uploaded_file, sheet_name="Economics_Matrix")
        f_df = pd.read_excel(uploaded_file, sheet_name="Quarterly_Forecast")
        c_df = pd.read_excel(uploaded_file, sheet_name="Weekly_Chase")
        
        PRODUCTS = eco_df["Product"].tolist()
        WEEKS = [col for col in f_df.columns if col != "Product"]
        
        FINANCIALS, UPFRONT_FORECAST, WEEKLY_CHASE = {}, {}, {}
        for _, row in eco_df.iterrows():
            FINANCIALS[row["Product"]] = {
                "price": row["Price"], "cost": row["Cost"], "fine": row["SLA_Fine"], "premium": row["Rush_Premium"],
                "L1": {"rate": row["L1_Rate"], "time": row["L1_Setup_Time"], "cost": row["L1_Setup_Cost"]},
                "L2": {"rate": row["L2_Rate"], "time": row["L2_Setup_Time"], "cost": row["L2_Setup_Cost"]},
                "L3": {"rate": row["L3_Rate"], "time": row["L3_Setup_Time"], "cost": row["L3_Setup_Cost"]}
            }
        for _, row in f_df.iterrows(): UPFRONT_FORECAST[row["Product"]] = {w: row[w] for w in WEEKS}
        for _, row in c_df.iterrows(): WEEKLY_CHASE[row["Product"]] = {w: row[w] for w in WEEKS}
    else:
        PRODUCTS, WEEKS = DEFAULT_PRODUCTS, DEFAULT_WEEKS
        FINANCIALS, UPFRONT_FORECAST, WEEKLY_CHASE = DEFAULT_FINANCIALS, DEFAULT_FORECAST, DEFAULT_CHASE
except Exception as e:
    st.error(f"❌ Excel Parsing Error. ({str(e)})")
    st.stop()

LINES = [f"L{i+1}" for i in range(num_lines)]

# --- 4. THE 3D ROUTING SOLVER ---
def optimize_operations(strat, lines, capacity_limit):
    prob = pulp.LpProblem("Digital_Twin_Solver", pulp.LpMaximize)
    
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
        for w in WEEKS:
            prob += total_prod[p][w] == pulp.lpSum([prod_line[p][l][w] for l in lines])
    
    revenue = pulp.lpSum([sold[p][w] * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    expedite_rev = pulp.lpSum([expedited_sold[p][w] * FINANCIALS[p]["premium"] for p in PRODUCTS for w in WEEKS])
    materials_cost = pulp.lpSum([total_prod[p][w] * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    inv_cost = pulp.lpSum([inv[p][w] * holding_cost for p in PRODUCTS for w in WEEKS])
    sla_fines = pulp.lpSum([shortage[p][w] * FINANCIALS[p]["fine"] for p in PRODUCTS for w in WEEKS])
    
    # Costs are now pulled from the specific line metrics
    setup_fees = pulp.lpSum([setup_line[p][l][w] * FINANCIALS[p][l]["cost"] for p in PRODUCTS for l in lines for w in WEEKS])
    
    labor_hours = []
    for p in PRODUCTS:
        for l in lines:
            rate = FINANCIALS[p][l]["rate"]
            time = FINANCIALS[p][l]["time"]
            for w in WEEKS:
                if rate > 0:
                    labor_hours.append((prod_line[p][l][w]/rate) + (setup_line[p][l][w]*time))
                else:
                    prob += prod_line[p][l][w] == 0
                    prob += setup_line[p][l][w] == 0
                    
    labor_cost = pulp.lpSum(labor_hours) * labor_rate
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
                rate = FINANCIALS[p][l]["rate"]
                if rate > 0:
                    max_prod = capacity_limit * rate
                    prob += prod_line[p][l][w] <= max_prod * setup_line[p][l][w]
                    prob += setup_line[p][l][w] <= line_active[l][w]

                    if strat == "Legacy: Min 40-Hour Run":
                        prob += prod_line[p][l][w] >= (40 * rate) * setup_line[p][l][w]

    for l in lines:
        for w in WEEKS:
            time_used = []
            for p in PRODUCTS:
                rate = FINANCIALS[p][l]["rate"]
                time = FINANCIALS[p][l]["time"]
                if rate > 0:
                    time_used.append((prod_line[p][l][w]/rate) + (setup_line[p][l][w]*time))
            prob += pulp.lpSum(time_used) <= capacity_limit * line_active[l][w]

    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=15))
    return prob, total_prod, inv, sold, shortage, rollover, prod_line, setup_line, expedited_sold, line_active

# --- 5. EXECUTION ---
def get_val(var):
    return var.varValue if var.varValue is not None else 0

with st.spinner(f"Simulating Digital Twin: {len(PRODUCTS)} SKUs across {num_lines} Lines..."):
    prob_status, prod_v, inv_v, sold_v, short_v, roll_v, p_line_v, s_line_v, expedite_v, active_v = optimize_operations(mode, LINES, weekly_capacity)

if pulp.LpStatus[prob_status.status] != 'Optimal':
    st.warning(f"⚠️ **Optimality Gap:** Displaying best margin found within 15 seconds.")

# --- 6. VISUALS ---
tab1, tab2, tab3, tab4 = st.tabs(["📈 Percent-Yield P&L", "⚙️ Twin Factory Routing", "📦 S&OP Service Level", "📊 Machine Matrix (Excel Data)"])

with tab1:
    st.subheader(f"Contract Manufacturer P&L ({mode})")
    
    rev = sum([get_val(sold_v[p][w]) * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    expedite_rev = sum([get_val(expedite_v[p][w]) * FINANCIALS[p]["premium"] for p in PRODUCTS for w in WEEKS])
    total_topline = rev + expedite_rev
    
    mat_cost = sum([get_val(prod_v[p][w]) * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    setups = sum([get_val(s_line_v[p][l][w]) * FINANCIALS[p][l]["cost"] for p in PRODUCTS for l in LINES for w in WEEKS])
    holding = sum([get_val(inv_v[p][w]) * holding_cost for p in PRODUCTS for w in WEEKS])
    fines = sum([get_val(short_v[p][w]) * FINANCIALS[p]["fine"] for p in PRODUCTS for w in WEEKS])
    
    tot_hrs = sum([(get_val(p_line_v[p][l][w]) / FINANCIALS[p][l]["rate"]) + (get_val(s_line_v[p][l][w]) * FINANCIALS[p][l]["time"]) for p in PRODUCTS for l in LINES for w in WEEKS if FINANCIALS[p][l]["rate"] > 0])
    labor = tot_hrs * labor_rate
    overhead = sum([get_val(active_v[l][w]) * fixed_line_cost for l in LINES for w in WEEKS])
    
    net = total_topline - (mat_cost + setups + holding + fines + labor + overhead)
    
    # --- PROMINENT SERVICE LEVEL DASHBOARD ---
    glob_demand = sum([UPFRONT_FORECAST[p][w] + WEEKLY_CHASE[p][w] + (get_val(roll_v[p][WEEKS[WEEKS.index(w)-1]]) if WEEKS.index(w) > 0 else 0) for p in PRODUCTS for w in WEEKS])
    glob_sold = sum([get_val(sold_v[p][w]) for p in PRODUCTS for w in WEEKS])
    service_level = (glob_sold / glob_demand * 100) if glob_demand > 0 else 0
    
    active_line_weeks = sum([get_val(active_v[l][w]) for l in LINES for w in WEEKS])
    total_line_weeks = len(LINES) * len(WEEKS)
    
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("GLOBAL SERVICE LEVEL (%)", f"{service_level:.1f}%")
    kpi2.metric("Net Contribution (£)", f"£{net:,.0f}")
    kpi3.metric("Line Activation Rate", f"{(active_line_weeks/total_line_weeks)*100:.0f}%")
    
    st.markdown("---")
    def pct(val): return f"{(val / total_topline) * 100:.1f}%" if total_topline > 0 else "0.0%"
    
    pl_df = pd.DataFrame({
        "Financial Line Item": ["Wholesale Base Revenue", "Rush Order Surcharges", "TOTAL TOPLINE REVENUE", "Materials Cost (COGS)", "Direct Labor", "Fixed Line Overhead", "Dynamic Setup Penalties", "Holding Costs", "SLA Shortage Penalties", "NET CONTRIBUTION"],
        "Amount (£)": [f"£{rev:,.0f}", f"£{expedite_rev:,.0f}", f"£{total_topline:,.0f}", f"-£{mat_cost:,.0f}", f"-£{labor:,.0f}", f"-£{overhead:,.0f}", f"-£{setups:,.0f}", f"-£{holding:,.0f}", f"-£{fines:,.0f}", f"£{net:,.0f}"],
        "% of Revenue": [pct(rev), pct(expedite_rev), "100.0%", pct(-mat_cost), pct(-labor), pct(-overhead), pct(-setups), pct(-holding), pct(-fines), pct(net)]
    })
    st.table(pl_df)

with tab2:
    st.subheader(f"Digital Twin Line Routing ({mode})")
    
    st.info("💡 **Fisher Trade-off:** Watch how the solver inherently locks 'Lifting Straps (Stable)' onto Dedicated Line 1 to maximize efficiency, while dynamically routing the Volatile products onto Flex Lines 2 & 3 to chase spikes without incurring massive setup penalties.")
    
    for l in LINES:
        st.markdown(f"#### 🏭 {l}")
        l_data = []
        for w in WEEKS:
            if get_val(active_v[l][w]) == 0:
                l_data.append({"Week": w, "Status": "🛑 SHUT DOWN (Overhead Saved)"})
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
                else:
                    row[p] = "Incompatible"
                
            row["Hrs Used"] = f"{total_l_hrs:.1f} / {weekly_capacity}"
            l_data.append(row)
        st.dataframe(pd.DataFrame(l_data), use_container_width=True)

with tab3:
    st.subheader("Service Level & S&OP Audit")
    sku = st.selectbox("Select SKU to Audit", PRODUCTS)
    
    # SKU specific service level calculation
    sku_dem = sum([UPFRONT_FORECAST[sku][w] + WEEKLY_CHASE[sku][w] for w in WEEKS])
    sku_sold = sum([get_val(sold_v[sku][w]) for w in WEEKS])
    st.metric(f"🎯 {sku} Service Level", f"{(sku_sold/sku_dem*100) if sku_dem > 0 else 0:.1f}%")
    
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
            "SLA Missed (Penalty)": int(get_val(short_v[sku][w]))
        })
    st.dataframe(pd.DataFrame(audit_data), use_container_width=True)

with tab4:
    st.subheader("Digital Twin Constraint Matrix")
    st.markdown("This data is driven entirely by the uploaded Excel Twin Matrix. Notice the massive setup penalties designed to create 'Dedicated' vs 'Flexible' lines.")
    eco_data = []
    for p, data in FINANCIALS.items():
        row = {"Product": p, "SLA Fine": f"£{data['fine']}", "Rush Premium": f"£{data['premium']}"}
        for l in ["L1", "L2", "L3"]:
            if l in data:
                row[f"{l} Rate"] = f"{data[l]['rate']}/h"
                row[f"{l} Setup Time"] = f"{data[l]['time']}h"
                row[f"{l} Setup Cost"] = f"£{data[l]['cost']}"
        eco_data.append(row)
    st.table(pd.DataFrame(eco_data))
