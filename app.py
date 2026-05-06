import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, date
import numpy as np

# ─── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Loan Portfolio KPI Dashboard",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─── CUSTOM CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600&family=Playfair+Display:wght@600;700&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

.stApp { background: #0a0e1a; }

/* Hide streamlit chrome */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 1.5rem 2rem 2rem 2rem; max-width: 1600px; }

/* HEADER */
.dash-header {
    background: #111827;
    border: 1px solid #1e2d45;
    border-radius: 12px;
    padding: 18px 28px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 24px;
}
.dash-title { font-family: 'Playfair Display', serif; font-size: 22px; color: #e8c96e; }
.dash-sub { font-size: 12px; color: #64748b; letter-spacing: 0.06em; text-transform: uppercase; margin-top: 2px; }
.live-badge {
    display: inline-flex; align-items: center; gap: 6px;
    background: rgba(34,197,94,0.1); border: 1px solid rgba(34,197,94,0.25);
    color: #22c55e; font-size: 11px; font-weight: 600; letter-spacing: 0.08em;
    text-transform: uppercase; padding: 4px 12px; border-radius: 20px;
}

/* KPI CARDS */
.kpi-card {
    background: #111827;
    border: 1px solid #1e2d45;
    border-radius: 12px;
    padding: 20px;
    position: relative;
    overflow: hidden;
    height: 100%;
}
.kpi-card::before {
    content: ''; position: absolute;
    top: 0; left: 0; right: 0; height: 2px;
    background: linear-gradient(90deg, #c9a84c, transparent);
}
.kpi-card.green::before { background: linear-gradient(90deg, #22c55e, transparent); }
.kpi-card.red::before { background: linear-gradient(90deg, #ef4444, transparent); }
.kpi-card.blue::before { background: linear-gradient(90deg, #3b82f6, transparent); }
.kpi-icon { font-size: 20px; margin-bottom: 8px; }
.kpi-label { font-size: 11px; color: #64748b; letter-spacing: 0.05em; text-transform: uppercase; font-weight: 500; margin-bottom: 6px; }
.kpi-value { font-family: 'DM Mono', monospace; font-size: 28px; font-weight: 500; color: #e2e8f0; line-height: 1; margin-bottom: 6px; }
.kpi-sub { font-size: 11px; color: #64748b; }
.kpi-up { color: #22c55e; }
.kpi-down { color: #ef4444; }
.kpi-warn { color: #f59e0b; }

/* PANELS */
.panel {
    background: #111827;
    border: 1px solid #1e2d45;
    border-radius: 12px;
    padding: 22px;
    margin-bottom: 0;
    height: 100%;
}
.panel-title { font-size: 13px; font-weight: 600; color: #e2e8f0; margin-bottom: 16px; }

/* SECTION LABEL */
.section-label {
    font-family: 'DM Mono', monospace;
    font-size: 10px; letter-spacing: 0.12em;
    text-transform: uppercase; color: #64748b;
    margin-bottom: 14px; margin-top: 8px;
}

/* BADGES */
.badge { display: inline-block; padding: 3px 9px; border-radius: 20px; font-size: 10px; font-weight: 600; }
.badge-green { background: rgba(34,197,94,0.12); color: #22c55e; }
.badge-red { background: rgba(239,68,68,0.12); color: #ef4444; }
.badge-amber { background: rgba(245,158,11,0.12); color: #f59e0b; }
.badge-blue { background: rgba(59,130,246,0.12); color: #3b82f6; }

/* GOAL ROWS */
.goal-row { display: flex; align-items: center; gap: 12px; margin-bottom: 12px; }
.goal-lbl { font-size: 12px; color: #94a3b8; width: 190px; flex-shrink: 0; }
.goal-pct-lbl { font-family: 'DM Mono', monospace; font-size: 12px; width: 50px; text-align: right; }

/* DIVIDER */
.divider { height: 1px; background: #1e2d45; margin: 4px 0 16px 0; }

/* FILTER PILLS */
.stRadio > div { flex-direction: row; flex-wrap: wrap; gap: 8px; }
.stRadio label { background: #1a2236; border: 1px solid #1e2d45; border-radius: 20px; padding: 4px 14px; font-size: 12px; color: #94a3b8; cursor: pointer; }
</style>
""", unsafe_allow_html=True)

# ─── LOAD DATA ─────────────────────────────────────────────────────────────────
@st.cache_data
def load_data(uploaded_file=None):
    """Load from uploaded Excel file, or fall back to demo data."""
    if uploaded_file is not None:
        try:
            xl = pd.ExcelFile(uploaded_file)
            sheets = xl.sheet_names
            # Try to read the pipeline/loan sheet
            if "Pipeline Tracker" in sheets:
                loans = xl.parse("Pipeline Tracker", header=None)
                # Find header row (look for 'Borrower')
                for i, row in loans.iterrows():
                    if 'Borrower' in row.values:
                        loans.columns = loans.iloc[i]
                        loans = loans.iloc[i+1:].reset_index(drop=True)
                        break

            # Clean up column names
            loans.columns = [str(c).strip() if c else f'Col_{i}' for i, c in enumerate(loans.columns)]

            # Convert numeric columns
            numeric_cols = ['Note Amount', 'Disbursed', 'Left to Fund', 'Interest Rate',
                            'Monthly Pmt', 'Days To Maturity', 'Days Since Close']
            for col in numeric_cols:
                if col in loans.columns:
                    loans[col] = pd.to_numeric(loans[col], errors='coerce').fillna(0)

            # Rename columns to match app expectations
            rename_map = {
                'Monthly Pmt': 'Monthly Pmt',
                'Days Since Close': 'Days To Maturity',
            }
            loans = loans.rename(columns=rename_map)

            # Add computed columns if missing
            if 'Left to Fund' not in loans.columns and 'Note Amount' in loans.columns:
                loans['Left to Fund'] = loans['Note Amount'] - loans['Disbursed']
            if 'Pct Funded' not in loans.columns and 'Note Amount' in loans.columns:
                loans['Pct Funded'] = loans.apply(
                    lambda r: r['Disbursed'] / r['Note Amount'] if r['Note Amount'] > 0 else 0, axis=1)
            if 'Risk' not in loans.columns:
                loans['Risk'] = '🟢 OK'
            if 'Status' not in loans.columns:
                loans['Status'] = 'Active'
            if 'Loan #' not in loans.columns and 'Loan#' in loans.columns:
                loans = loans.rename(columns={'Loan#': 'Loan #'})

            # Drop empty rows
            loans = loans[loans['Borrower'].notna() & (loans['Borrower'] != 0) & (loans['Borrower'] != '')].reset_index(drop=True)

            return loans, True
        except Exception as e:
            st.warning(f"Could not read file: {e}. Using demo data.")
            return demo_loans(), False
    return demo_loans(), False

def demo_loans():
    """Generate realistic demo loan data."""
    data = {
        'Borrower': ['James Whitmore','Maria Chen','Derek Okafor','Laura Simmons','TJ Palmer LLC',
                     'Mark Donovan','Sofia Reyes','Tony Crawford','Bright Star LLC','Peter Nguyen',
                     'Rachel Brooks','Horizon Props','Kevin Walsh','Diane Foster','Apex Capital LLC'],
        'Loan #': ['1001','1002','1003','1004','1005','1006','1007','1008','1009','1010','1011','1012','1013','1014','1015'],
        'Property': ['412 Cedar Ave, Pittsburgh','1832 Elm St, Bethel Park','520 Broad Blvd, Monroeville',
                     '1943 Park St, Overbrook','511 California Ave, Pittsburgh','307 Cypress St, West Newton',
                     '621 Oberdick Dr, McKeesport','18 Center St, Fayette City','133 Spring St, Aliquippa',
                     '800 3rd Ave, Hyde Park','2604 Ivyglen St, Pittsburgh','1615 McClure St, Homestead',
                     '740 Main St, Carnegie','215 Oak Ave, Monroeville','988 Ridge Rd, Pittsburgh'],
        'Note Amount': [147500,112000,105000,130000,287000,105000,192500,135000,60000,108000,129500,192500,175000,155000,245000],
        'Disbursed':   [130200,112000, 95000, 66650,239123,104150,175950, 84300, 25000, 80500, 74695,162434,175000,120000,244500],
        'Interest Rate': [0.13,0.13,0.13,0.13,0.13,0.13,0.13,0.13,0.13,0.13,0.13,0.13,0.13,0.15,0.13],
        'Monthly Pmt': [1597.92,1213.33,1137.5,1408.33,3109.17,1137.5,2085.42,1462.5,650,1170,1402.92,2085.42,1895.83,1500,2654.17],
        'Status': ['Active - Drawing','Fully Funded','Active - Drawing','Active - Drawing','Active - Drawing',
                   'Active - Drawing','Active - Drawing','Active - Drawing','Active - Drawing','Active - Drawing',
                   'Active - Drawing','Active - Drawing','Fully Funded','Active - Drawing','Active - Drawing'],
        'Days To Maturity': [44,147,153,163,163,155,195,196,214,187,214,203,181,234,168],
        'Risk': ['🟠 Urgent','🟢 OK','🟢 OK','🟢 OK','🟢 OK','🟢 OK','🟢 OK','🟢 OK','🟢 OK','🟢 OK','🟢 OK','🟢 OK','🟢 OK','🟢 OK','🟢 OK'],
        'Lending Entity': ['Dannic Properties LLC','ACP Debt Fund','Dannic Properties LLC','ACP Debt Fund',
                           'Dannic Properties LLC','Dannic Properties LLC','ACP Debt Fund','ACP Debt Fund',
                           'ACP Debt Fund','ACP Debt Fund','ACP Debt Fund','ACP Debt Fund',
                           'Dannic Properties LLC','ACP Debt Fund','Dannic Properties LLC'],
    }
    df = pd.DataFrame(data)
    df['Left to Fund'] = df['Note Amount'] - df['Disbursed']
    df['Pct Funded'] = df['Disbursed'] / df['Note Amount']
    return df

def demo_kpis():
    return [
        {'Category':'Active Loans','KPI':'Loans Funded','MTD':4,'Target':80,'MTD_fmt':'4','Target_fmt':'80'},
        {'Category':'Active Loans','KPI':'Total Loan Value Funded','MTD':6952500,'Target':8000000,'MTD_fmt':'$6.95M','Target_fmt':'$8M'},
        {'Category':'Pipeline','KPI':'# of Loans in Pipeline','MTD':11,'Target':150,'MTD_fmt':'11','Target_fmt':'150'},
        {'Category':'Pipeline','KPI':'Pipeline Value $','MTD':1917000,'Target':8000000,'MTD_fmt':'$1.92M','Target_fmt':'$8M'},
        {'Category':'Pipeline','KPI':'Term Sheets Sent','MTD':12,'Target':150,'MTD_fmt':'12','Target_fmt':'150'},
        {'Category':'Draw Activity','KPI':'# of Draws','MTD':9,'Target':25,'MTD_fmt':'9','Target_fmt':'25'},
        {'Category':'Draw Activity','KPI':'Total Draw Amount Funded','MTD':6144375,'Target':8000000,'MTD_fmt':'$6.14M','Target_fmt':'$8M'},
        {'Category':'Closings','KPI':'Loans Closing This Month','MTD':1,'Target':10,'MTD_fmt':'1','Target_fmt':'10'},
        {'Category':'Closings','KPI':'Closed Loan Amount','MTD':122500,'Target':8000000,'MTD_fmt':'$122K','Target_fmt':'$8M'},
        {'Category':'Revenue','KPI':'Interest Collected','MTD':58420,'Target':80000,'MTD_fmt':'$58.4K','Target_fmt':'$80K'},
        {'Category':'Revenue','KPI':'Fees Collected','MTD':6200,'Target':None,'MTD_fmt':'$6.2K','Target_fmt':'—'},
        {'Category':'Risk','KPI':'Total Past Due Amount','MTD':94000,'Target':0,'MTD_fmt':'$94K','Target_fmt':'$0'},
        {'Category':'Capital','KPI':'Total Capital Deployed','MTD':6144375,'Target':8000000,'MTD_fmt':'$6.14M','Target_fmt':'$8M'},
        {'Category':'Capital','KPI':'Pipeline Capital Demand','MTD':1917000,'Target':8000000,'MTD_fmt':'$1.92M','Target_fmt':'$8M'},
    ]

# ─── HELPERS ───────────────────────────────────────────────────────────────────
def fmt_currency(n):
    if n >= 1_000_000: return f"${n/1_000_000:.1f}M"
    if n >= 1_000: return f"${n/1_000:.0f}K"
    return f"${n:,.0f}"

def status_badge(pct):
    if pct is None: return "—"
    if pct >= 75: return "⚠ Near Target"
    if pct >= 20: return "In Progress"
    return "✘ Below Target"

def status_color(pct):
    if pct is None: return "#64748b"
    if pct >= 75: return "#f59e0b"
    if pct >= 20: return "#3b82f6"
    return "#ef4444"

# ─── LOAD ──────────────────────────────────────────────────────────────────────
df, is_real = load_data()
kpis = demo_kpis()

# ─── HEADER ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="dash-header">
  <div>
    <div class="dash-title">🏦 Capital Bridge Partners</div>
    <div class="dash-sub">Loan Portfolio · KPI Dashboard</div>
  </div>
  <div style="display:flex;align-items:center;gap:16px;">
    <div class="live-badge">● {'Live Data' if is_real else 'Demo Mode'}</div>
    <div style="font-family:'DM Mono',monospace;font-size:12px;color:#64748b;">{datetime.now().strftime('%b %d, %Y')}</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─── UPLOAD YOUR FILE ──────────────────────────────────────────────────────────
with st.expander("📂 Connect Your Excel File", expanded=not is_real):
    st.markdown("Upload your **Lending_KPI_Dashboard.xlsx** to see your real data.")
    uploaded = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")
    if uploaded:
        df, is_real = load_data(uploaded)
        st.success("✅ File loaded! Dashboard updated with your data.")

# ─── PORTFOLIO METRICS ─────────────────────────────────────────────────────────
st.markdown('<div class="section-label">Portfolio Overview</div>', unsafe_allow_html=True)

total_note = df['Note Amount'].sum()
total_disbursed = df['Disbursed'].sum()
left_to_fund = df['Left to Fund'].sum()
active_loans = len(df)
past_due = (df['Days To Maturity'] <= 0).sum() if 'Days To Maturity' in df.columns else 2
pct_funded = total_disbursed / total_note * 100 if total_note > 0 else 0

c1, c2, c3, c4, c5 = st.columns(5)
cards = [
    (c1, "🏛", "Active Loans", str(active_loans), "+3 from last month", ""),
    (c2, "💰", "Total Note Value", fmt_currency(total_note), "Across all active loans", "green"),
    (c3, "✅", "Total Disbursed", fmt_currency(total_disbursed), f"{pct_funded:.1f}% funded", ""),
    (c4, "⏳", "Left to Fund", fmt_currency(left_to_fund), "Pending draws", "blue"),
    (c5, "⚠️", "Past Due Loans", str(past_due), f"Require attention", "red"),
]

for col, icon, label, val, sub, cls in cards:
    with col:
        st.markdown(f"""
        <div class="kpi-card {cls}">
          <div class="kpi-icon">{icon}</div>
          <div class="kpi-label">{label}</div>
          <div class="kpi-value">{val}</div>
          <div class="kpi-sub">{sub}</div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<div style='margin-top:20px;'></div>", unsafe_allow_html=True)

# ─── GOAL PROGRESS + TREND CHART ──────────────────────────────────────────────
st.markdown('<div class="section-label">Monthly Performance</div>', unsafe_allow_html=True)
col_goal, col_chart = st.columns([1.1, 0.9])

with col_goal:
    st.markdown('<div class="panel"><div class="panel-title">📈 Monthly Goal Progress</div>', unsafe_allow_html=True)

    fig_goals = go.Figure()
    goal_labels, goal_pcts, goal_colors = [], [], []

    for k in kpis:
        if k['Target'] and k['Target'] > 0:
            pct = min((k['MTD'] / k['Target']) * 100, 100)
            color = '#22c55e' if pct >= 75 else '#c9a84c' if pct >= 30 else '#ef4444'
            goal_labels.append(k['KPI'][:28])
            goal_pcts.append(round(pct, 1))
            goal_colors.append(color)

    fig_goals = go.Figure(go.Bar(
        x=goal_pcts,
        y=goal_labels,
        orientation='h',
        marker=dict(color=goal_colors, line=dict(width=0)),
        text=[f"{p}%" for p in goal_pcts],
        textposition='outside',
        textfont=dict(family='DM Mono', size=10, color='#94a3b8'),
    ))

    fig_goals.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=0, r=60, t=0, b=0),
        height=340,
        xaxis=dict(range=[0, 110], showgrid=False, showticklabels=False, zeroline=False),
        yaxis=dict(tickfont=dict(family='DM Sans', size=11, color='#94a3b8'), gridcolor='#1e2d45'),
        bargap=0.35,
    )

    st.plotly_chart(fig_goals, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)

with col_chart:
    st.markdown('<div class="panel"><div class="panel-title">📊 Weekly Draw Activity</div>', unsafe_allow_html=True)

    weeks = ['Wk 1', 'Wk 2', 'Wk 3', 'Wk 4', 'Wk 5']
    values = [820000, 940000, 1050000, 1110000, 1240000]

    fig_bar = go.Figure(go.Bar(
        x=weeks, y=values,
        marker=dict(
            color=values,
            colorscale=[[0, 'rgba(201,168,76,0.3)'], [1, 'rgba(201,168,76,0.9)']],
            line=dict(width=0)
        ),
        text=[fmt_currency(v) for v in values],
        textposition='outside',
        textfont=dict(family='DM Mono', size=10, color='#94a3b8'),
    ))
    fig_bar.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=0, r=0, t=10, b=0),
        height=160,
        xaxis=dict(tickfont=dict(family='DM Sans', size=11, color='#64748b'), showgrid=False, zeroline=False),
        yaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
        bargap=0.3,
    )
    st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar': False})

    # Revenue mini cards
    r1, r2 = st.columns(2)
    with r1:
        st.markdown("""
        <div style="background:#1a2236;border-radius:8px;padding:12px;">
          <div style="font-size:11px;color:#64748b;margin-bottom:4px;">Interest Collected</div>
          <div style="font-family:'DM Mono',monospace;font-size:20px;color:#e8c96e;">$58,420</div>
        </div>""", unsafe_allow_html=True)
    with r2:
        st.markdown("""
        <div style="background:#1a2236;border-radius:8px;padding:12px;">
          <div style="font-size:11px;color:#64748b;margin-bottom:4px;">Fees Collected</div>
          <div style="font-family:'DM Mono',monospace;font-size:20px;color:#e8c96e;">$6,200</div>
        </div>""", unsafe_allow_html=True)

    # Funding pie
    st.markdown("<div style='margin-top:12px;'></div>", unsafe_allow_html=True)
    fully = (df['Pct Funded'] >= 0.99).sum()
    drawing = (df['Pct Funded'] < 0.99).sum()
    fig_pie = go.Figure(go.Pie(
        labels=['Fully Funded', 'Active Drawing'],
        values=[fully, drawing],
        hole=0.65,
        marker=dict(colors=['#22c55e', '#c9a84c'], line=dict(width=0)),
        textfont=dict(family='DM Sans', size=11, color='white'),
    ))
    fig_pie.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=0, r=0, t=0, b=0),
        height=120,
        showlegend=True,
        legend=dict(font=dict(family='DM Sans', size=10, color='#94a3b8'), orientation='h', x=0.1, y=-0.1),
        annotations=[dict(text=f'{fully}/{active_loans}', x=0.5, y=0.5, font=dict(family='DM Mono', size=16, color='#e2e8f0'), showarrow=False)],
    )
    st.plotly_chart(fig_pie, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("<div style='margin-top:20px;'></div>", unsafe_allow_html=True)

# ─── DRAW STATUS TABLE ─────────────────────────────────────────────────────────
st.markdown('<div class="section-label">Draw Status & Loan Detail</div>', unsafe_allow_html=True)

col_draw, col_mat = st.columns([1.6, 1])

with col_draw:
    draw_df = df[['Borrower','Property','Note Amount','Disbursed','Pct Funded','Interest Rate','Status']].copy()
    draw_df['Note Amount'] = draw_df['Note Amount'].apply(fmt_currency)
    draw_df['Disbursed'] = draw_df['Disbursed'].apply(fmt_currency)
    draw_df['Pct Funded'] = (draw_df['Pct Funded'] * 100).round(1).astype(str) + '%'
    draw_df['Interest Rate'] = (draw_df['Interest Rate'] * 100).round(0).astype(int).astype(str) + '%'

    st.markdown('<div class="panel"><div class="panel-title">🔄 Draw Status</div>', unsafe_allow_html=True)
    st.dataframe(
        draw_df,
        use_container_width=True,
        height=340,
        hide_index=True,
        column_config={
            "Borrower": st.column_config.TextColumn("Borrower", width=130),
            "Property": st.column_config.TextColumn("Property", width=200),
            "Note Amount": st.column_config.TextColumn("Note Amt", width=80),
            "Disbursed": st.column_config.TextColumn("Disbursed", width=80),
            "Pct Funded": st.column_config.TextColumn("Funded", width=60),
            "Interest Rate": st.column_config.TextColumn("Rate", width=50),
            "Status": st.column_config.TextColumn("Status", width=130),
        }
    )
    st.markdown('</div>', unsafe_allow_html=True)

with col_mat:
    st.markdown('<div class="panel"><div class="panel-title">⏰ Nearing Maturity</div>', unsafe_allow_html=True)

    if 'Days To Maturity' in df.columns:
        mat_df = df[['Borrower','Loan #','Days To Maturity','Risk']].copy()
        mat_df = mat_df.sort_values('Days To Maturity').head(10)

        fig_mat = go.Figure(go.Bar(
            x=mat_df['Days To Maturity'],
            y=mat_df['Borrower'],
            orientation='h',
            marker=dict(
                color=mat_df['Days To Maturity'],
                colorscale=[[0,'#ef4444'],[0.3,'#f59e0b'],[1,'#22c55e']],
                cmin=0, cmax=300,
                line=dict(width=0),
            ),
            text=mat_df['Days To Maturity'].astype(str) + 'd',
            textposition='outside',
            textfont=dict(family='DM Mono', size=10, color='#94a3b8'),
        ))
        fig_mat.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=0, r=40, t=0, b=0),
            height=340,
            xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
            yaxis=dict(tickfont=dict(family='DM Sans', size=10, color='#94a3b8'), gridcolor='#1e2d45'),
            bargap=0.3,
        )
        st.plotly_chart(fig_mat, use_container_width=True, config={'displayModeBar': False})

    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("<div style='margin-top:20px;'></div>", unsafe_allow_html=True)

# ─── BORROWER CONCENTRATION + KPI BREAKDOWN ───────────────────────────────────
st.markdown('<div class="section-label">Concentration & Full KPI Breakdown</div>', unsafe_allow_html=True)
col_conc, col_kpi = st.columns(2)

with col_conc:
    st.markdown('<div class="panel"><div class="panel-title">👤 Borrower Concentration</div>', unsafe_allow_html=True)

    conc = df.groupby('Borrower').agg(
        Loans=('Loan #', 'count'),
        Total_Note=('Note Amount', 'sum'),
        Disbursed=('Disbursed', 'sum'),
        Monthly_Income=('Monthly Pmt', 'sum'),
    ).reset_index()
    conc['Pct of Portfolio'] = conc['Total_Note'] / conc['Total_Note'].sum()
    conc = conc.sort_values('Total_Note', ascending=False)

    fig_conc = go.Figure(go.Bar(
        x=conc['Pct of Portfolio'] * 100,
        y=conc['Borrower'],
        orientation='h',
        marker=dict(
            color=conc['Pct of Portfolio'] * 100,
            colorscale=[[0,'#3b82f6'],[0.5,'#c9a84c'],[1,'#ef4444']],
            cmin=0, cmax=20,
            line=dict(width=0),
        ),
        text=(conc['Pct of Portfolio'] * 100).round(1).astype(str) + '%',
        textposition='outside',
        textfont=dict(family='DM Mono', size=10, color='#94a3b8'),
    ))
    fig_conc.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=0, r=50, t=0, b=0),
        height=360,
        xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
        yaxis=dict(tickfont=dict(family='DM Sans', size=10, color='#94a3b8')),
        bargap=0.3,
    )
    st.plotly_chart(fig_conc, use_container_width=True, config={'displayModeBar': False})
    st.markdown('</div>', unsafe_allow_html=True)

with col_kpi:
    st.markdown('<div class="panel"><div class="panel-title">📋 Full KPI Breakdown</div>', unsafe_allow_html=True)

    kpi_df = pd.DataFrame(kpis)
    kpi_df['% of Goal'] = kpi_df.apply(
        lambda r: f"{min((r['MTD']/r['Target'])*100, 100):.1f}%" if r['Target'] else '—', axis=1
    )
    kpi_df['Status'] = kpi_df.apply(
        lambda r: status_badge((r['MTD']/r['Target'])*100 if r['Target'] else None), axis=1
    )

    display_df = kpi_df[['Category','KPI','MTD_fmt','Target_fmt','% of Goal','Status']].copy()
    display_df.columns = ['Category','KPI','MTD','Target','% Goal','Status']

    st.dataframe(
        display_df,
        use_container_width=True,
        height=360,
        hide_index=True,
        column_config={
            "Category": st.column_config.TextColumn("Category", width=100),
            "KPI": st.column_config.TextColumn("KPI", width=180),
            "MTD": st.column_config.TextColumn("MTD", width=70),
            "Target": st.column_config.TextColumn("Target", width=70),
            "% Goal": st.column_config.TextColumn("% Goal", width=60),
            "Status": st.column_config.TextColumn("Status", width=100),
        }
    )
    st.markdown('</div>', unsafe_allow_html=True)

# ─── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center;padding:20px;color:#334155;font-size:11px;margin-top:12px;border-top:1px solid #1e2d45;">
  Capital Bridge Partners LLC · Loan Portfolio Dashboard · Built with Streamlit
</div>
""", unsafe_allow_html=True)
