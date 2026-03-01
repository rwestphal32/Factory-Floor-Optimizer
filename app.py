import streamlit as st
import pandas as pd
import numpy as np
import pulp
import io

st.set_page_config(page_title="PwC Value Creation: CM Operations", layout="wide")

st.title("🏭 Strategy & Value Creation: CM Operations Optimizer")
st.markdown("**Context:** Contract Manufacturer (CM) optimizing production, routing, and SLA contracts for a Wholesaler client.")

# --- 1. DEFAULT DATA (High Utilization 2-Line Baseline) ---
DEFAULT_WEEKS = [f"W{i+1}" for i in range(12)]
DEFAULT_PRODUCTS = ["Lifting Straps", "Weight Belts", "Knee Sleeves"]

# THE FIX: True Hybrid Penalty applied to Line 3
DEFAULT_FINANCIALS = {
    "Lifting Straps": {"price": 25, "cost": 10, "co_time": 2, "co_cost": 500, "fine": 10, "premium": 3, "rates": {"Line 1": 40, "Line 2": 0, "Line 3": 25}},
    "Weight Belts":  {"price": 85, "cost": 45, "co_time": 8, "co_cost": 2500, "fine": 30, "premium": 15, "rates": {"Line 1": 0, "Line 2": 15, "Line 3": 10}},
    "Knee Sleeves":  {"price": 45, "cost": 22, "co_time": 4, "co_cost": 1200, "fine": 15, "premium": 5, "rates": {"Line 1": 35, "Line 2": 35, "Line 3": 20}}
}

@st.cache_data
def get_default_demand():
    np.random.seed(42)
    q_forecast, w_chase = {}, {}
    for p in DEFAULT_PRODUCTS:
        base = 3000 if p == "Lifting Straps" else (1200 if p == "Weight Belts" else 2000)
        spike = 2000 if p == "Lifting Straps" else (1000 if p == "Weight Belts" else 1500)
        
        q_forecast[p] = {DEFAULT_WEEKS[i]: np.random.randint(int(base*0.8), int(base*1.2)) for i in range(12)}
        w_chase[p] = {DEFAULT_WEEKS[i]: (spike if 4 < i < 8 else 0) for i in range(12)}
    return q_forecast, w_chase

DEFAULT_FORECAST, DEFAULT_CHASE = get_default_demand()

def generate_excel_template():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        eco_df = pd.DataFrame([
            {"Product": "Lifting Straps", "Price": 25, "Materials_Cost": 10, "Setup_Time": 2, "Setup_Cost": 500, "SLA_Penalty": 10, "Rush_Surcharge": 3, "Line 1 Rate": 40, "Line 2 Rate": 0, "Line 3 Rate": 25},
            {"Product": "Weight Belts", "Price": 85, "Materials_Cost": 45, "Setup_Time": 8, "Setup_Cost": 2500, "SLA_Penalty": 30, "Rush_Surcharge": 15, "Line 1 Rate": 0, "Line 2 Rate": 15, "Line 3 Rate": 10},
            {"Product": "Knee Sleeves", "Price": 45, "Materials_Cost": 22, "Setup_Time": 4, "Setup_Cost": 1200, "SLA_Penalty": 15, "Rush_Surcharge": 5, "Line 1 Rate": 35, "Line 2 Rate": 35, "Line 3 Rate": 20}
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
    st.download_button("📥 Download CM Excel Template", data=generate_excel_template(), file_name="cm_routing_template.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    uploaded_file = st.file_uploader("Upload filled Excel Template", type=["xlsx"])
    
    st.markdown("---")
    st.header("🏢 2. CM Factory Parameters")
    with st.form("control_panel"):
        mode = st.radio("Model", ["Legacy: Min 40-Hour Run", "MILP: Value Optimized"])
        
        st.subheader("Labor, Overhead & Lines")
        num_lines = st.slider("Available Production Lines", 1, 3, 2)
        weekly_capacity = st.slider("Max Capacity Per Line (Hrs/Wk)", 80, 168, 120)
        labor_rate = st.slider("Flat Direct Labor (£/Hr per line)", 10, 50, 20)
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
                "co_time": row["Setup_Time"], "co_cost": row["Setup_Cost"],
                "fine": row["SLA_Penalty"], "premium": row["Rush_Surcharge"],
                "rates": {"Line 1": row.get("Line 1 Rate", 0), "Line 2": row.get("Line 2 Rate", 0), "Line 3": row.get("Line 3 Rate", 0)}
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
    
    # Financials
    revenue = pulp.lpSum([sold[p][w] * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    expedite_rev = pulp.lpSum([expedited_sold[p][w] * FINANCIALS[p]["premium"] for p in PRODUCTS for w in WEEKS])
    
    materials_cost = pulp.lpSum([total_prod[p][w] * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    inv_cost = pulp.lpSum([inv[p][w] * holding_cost for p in PRODUCTS for w in WEEKS])
    sla_fines = pulp.lpSum([shortage[p][w] * FINANCIALS[p]["fine"] for p in PRODUCTS for w in WEEKS])
    setup_fees = pulp.lpSum([setup_line[p][l][w] * FINANCIALS[p]["co_cost"] for p in PRODUCTS for l in lines for w in WEEKS])
    
    labor_hours = []
    for p in PRODUCTS:
        for l in lines:
            rate = FINANCIALS[p]["rates"].get(l, 0)
            for w in WEEKS:
                if rate > 0:
                    labor_hours.append((prod_line[p][l][w]/rate) + (setup_line[p][l][w]*FINANCIALS[p]["co_time"]))
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
                rate = FINANCIALS[p]["rates"].get(l, 0)
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
                rate = FINANCIALS[p]["rates"].get(l, 0)
                if rate > 0:
                    time_used.append((prod_line[p][l][w]/rate) + (setup_line[p][l][w]*FINANCIALS[p]["co_time"]))
            prob += pulp.lpSum(time_used) <= capacity_limit * line_active[l][w]

    prob.solve(pulp.PULP_CBC_CMD(msg=0, timeLimit=15))
    return prob, total_prod, inv, sold, shortage, rollover, prod_line, setup_line, expedited_sold, line_active

# --- 5. EXECUTION ---
def get_val(var):
    return var.varValue if var.varValue is not None else 0

with st.spinner(f"Routing {len(PRODUCTS)} SKUs across {num_lines} Line(s)..."):
    prob_status, prod_v, inv_v, sold_v, short_v, roll_v, p_line_v, s_line_v, expedite_v, active_v = optimize_operations(mode, LINES, weekly_capacity)

if pulp.LpStatus[prob_status.status] != 'Optimal':
    st.warning(f"⚠️ **Optimality Gap:** Status: {pulp.LpStatus[prob_status.status]}. Matrix hit 15s limit; displaying best margin.")

# --- 6. VISUALS ---
tab1, tab2, tab3, tab4 = st.tabs(["📈 Percent-Yield P&L", "⚙️ Factory Routing", "📦 S&OP Audit", "💰 SKU Economics"])

with tab1:
    st.subheader(f"Contract Manufacturer P&L ({mode})")
    
    rev = sum([get_val(sold_v[p][w]) * FINANCIALS[p]["price"] for p in PRODUCTS for w in WEEKS])
    expedite_rev = sum([get_val(expedite_v[p][w]) * FINANCIALS[p]["premium"] for p in PRODUCTS for w in WEEKS])
    total_topline = rev + expedite_rev
    
    mat_cost = sum([get_val(prod_v[p][w]) * FINANCIALS[p]["cost"] for p in PRODUCTS for w in WEEKS])
    setups = sum([get_val(s_line_v[p][l][w]) * FINANCIALS[p]["co_cost"] for p in PRODUCTS for l in LINES for w in WEEKS])
    holding = sum([get_val(inv_v[p][w]) * holding_cost for p in PRODUCTS for w in WEEKS])
    fines = sum([get_val(short_v[p][w]) * FINANCIALS[p]["fine"] for p in PRODUCTS for w in WEEKS])
    
    tot_hrs = sum([(get_val(p_line_v[p][l][w]) / FINANCIALS[p]["rates"].get(l, 1)) + (get_val(s_line_v[p][l][w]) * FINANCIALS[p]["co_time"]) for p in PRODUCTS for l in LINES for w in WEEKS if FINANCIALS[p]["rates"].get(l, 0) > 0])
    labor = tot_hrs * labor_rate
    overhead = sum([get_val(active_v[l][w]) * fixed_line_cost for l in LINES for w in WEEKS])
    
    net = total_topline - (mat_cost + setups + holding + fines + labor + overhead)
    
    def pct(val):
        if total_topline == 0: return "0.0%"
        return f"{(val / total_topline) * 100:.1f}%"
    
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Net Contribution (£)", f"£{net:,.0f}")
    kpi2.metric("Total Labor & Overhead (£)", f"£{(labor + overhead):,.0f}")
    
    active_line_weeks = sum([get_val(active_v[l][w]) for l in LINES for w in WEEKS])
    total_line_weeks = len(LINES) * len(WEEKS)
    kpi3.metric("Line Activation Rate", f"{(active_line_weeks/total_line_weeks)*100:.0f}%")
    
    st.markdown("---")
    pl_df = pd.DataFrame({
        "Financial Line Item": [
            "Wholesale Base Revenue", 
            "Rush Order Surcharges", 
            "TOTAL TOPLINE REVENUE",
            "Materials Cost (COGS)", 
            "Direct Labor (Flat Rate)", 
            "Fixed Line Overhead", 
            "Setup Fees", 
            "Holding Costs", 
            "SLA Shortage Penalties", 
            "NET CONTRIBUTION"
        ],
        "Amount (£)": [
            f"£{rev:,.0f}", f"£{expedite_rev:,.0f}", f"£{total_topline:,.0f}",
            f"-£{mat_cost:,.0f}", f"-£{labor:,.0f}", f"-£{overhead:,.0f}", 
            f"-£{setups:,.0f}", f"-£{holding:,.0f}", f"-£{fines:,.0f}", f"£{net:,.0f}"
        ],
        "% of Revenue": [
            pct(rev), pct(expedite_rev), "100.0%",
            pct(-mat_cost), pct(-labor), pct(-overhead),
            pct(-setups), pct(-holding), pct(-fines), pct(net)
        ]
    })
    st.table(pl_df)

with tab2:
    st.subheader(f"CM Line-Specific Routing ({mode})")
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
                rate = FINANCIALS[p]["rates"].get(l, 0)
                if rate > 0:
                    p_hrs = get_val(p_line_v[p][l][w]) / rate
                    s_hrs = get_val(s_line_v[p][l][w]) * FINANCIALS[p]["co_time"]
                    total_l_hrs += (p_hrs + s_hrs)
                    
                    status = []
                    if p_hrs > 0: status.append(f"Prod: {p_hrs:.1f}h")
                    if s_hrs > 0: status.append(f"Setup: {s_hrs:.1f}h")
                    row[p] = " | ".join(status) if status else "-"
                else:
                    row[p] = "Incompatible"
                
            row["Hrs Used"] = f"{total_l_hrs:.1f} / {weekly_capacity}"
            row["Utilization"] = f"{(total_l_hrs / weekly_capacity) * 100:.1f}%"
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

with tab4:
    st.subheader("Dynamic SKU Routing Economics")
    eco_data = []
    for p, metrics in FINANCIALS.items():
        rates_str = " | ".join([f"{l}: {r}/h" for l, r in metrics["rates"].items() if r > 0])
        eco_data.append({
            "Product": p,
            "Unit Margin": f"£{metrics['price'] - metrics['cost']:.2f}",
            "Compatible Lines & Speeds": rates_str,
            "Setup Time / Cost": f"{metrics['co_time']}h / £{metrics['co_cost']}",
            "SLA Fine": f"£{metrics['fine']}",
            "Rush Premium": f"£{metrics['premium']}"
        })
    st.table(pd.DataFrame(eco_data))
