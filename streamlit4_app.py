import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# 1. Config & Ultra-Modern Executive Styling
st.set_page_config(page_title="Executive Device Analytics", page_icon="🏥", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700;800&display=swap');
    
    /* บังคับตัวหนังสือดำเข้มและชัดเจนที่สุด */
    html, body, [class*="css"] { 
        font-family: 'Sarabun', sans-serif; 
        color: #000000 !important; 
    }
    
    .stApp { background-color: #ffffff; }

    /* ซ่อนเมนู GitHub และ Toolbar มุมขวา/ล่าง */
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    .stDeployButton {display:none;}
    [data-testid="stStatusWidget"] {visibility: hidden;}
    .stAppHeader {display:none;}
    
    .main-title { 
        color: #0f172a; font-size: 3.5rem; font-weight: 800; 
        text-align: center; margin-bottom: 40px;
    }

    /* MEGA KPI Card */
    .mega-kpi-container {
        background: #f8fafc;
        padding: 3.5rem 1rem;
        border-radius: 40px;
        text-align: center;
        box-shadow: 0 15px 30px rgba(0,0,0,0.06);
        border: 2px solid #e2e8f0;
        min-height: 380px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    
    .mega-label { color: #475569; font-size: 1.8rem; font-weight: 700; text-transform: uppercase; margin-bottom: 20px; }
    .mega-value { color: #000000; font-size: 6.5rem; font-weight: 800; margin: 0; line-height: 0.9; }
    .mega-delta-tag { margin-top: 2.5rem; padding: 12px 30px; border-radius: 100px; display: inline-block; font-weight: 800; font-size: 2.2rem; }

    .stButton>button {
        border-radius: 20px; background: #0f172a; color: white !important;
        font-weight: 800; border: none; padding: 1.5rem 2rem; font-size: 1.5rem;
    }
    </style>
    """, unsafe_allow_html=True)

# --- Keywords Configuration ---
keywords = ["Ventilator", "Foley", "Central line", "Port A Cath"]

def get_safe_total(df_in):
    d_cols = [c for c in df_in.columns if any(k.lower() in str(c).lower() for k in keywords)]
    if not d_cols: return 0, []
    df_c = df_in.copy()
    for c in d_cols: 
        df_c[c] = pd.to_numeric(df_c[c], errors='coerce').fillna(0)
    return int(df_c[d_cols].values.sum()), d_cols

def process_file_summary(uploaded_file):
    excel_file = pd.ExcelFile(uploaded_file)
    data = []
    for s in excel_file.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=s).dropna(how='all')
        total_sum, _ = get_safe_total(df)
        data.append({'Ward': s, 'Total_Days': int(total_sum)})
    return pd.DataFrame(data)

def color_growth(val):
    color = '#059669' if val > 0 else '#dc2626' if val < 0 else '#64748b'
    return f'color: {color}; font-weight: 800; font-size: 1.2rem;'

# --- Sidebar ---
st.sidebar.markdown("<h1 style='text-align:center;'>🏥 SYSTEM</h1>", unsafe_allow_html=True)
file_1 = st.sidebar.file_uploader("📂 เดือนที่ 1 (Previous Month)", type=["xlsx"], key="f1")
file_2 = st.sidebar.file_uploader("📂 เดือนที่ 2 (Current Month)", type=["xlsx"], key="f2")

if file_1 and file_2:
    page = st.sidebar.selectbox("🎯 เมนูเลือกดูข้อมูล", ["📊 Utilization Analytics", "📄 Data Editor"])

    df_stats_1 = process_file_summary(file_1)
    df_stats_2 = process_file_summary(file_2)
    df_compare = pd.merge(df_stats_1, df_stats_2, on='Ward', how='outer', suffixes=('_M1', '_M2')).fillna(0)
    
    df_compare['Diff'] = (df_compare['Total_Days_M2'] - df_compare['Total_Days_M1']).astype(int)
    df_compare['% Growth'] = ((df_compare['Diff'] / df_compare['Total_Days_M1']) * 100).replace([float('inf'), -float('inf')], 0).fillna(0).round(0).astype(int)

    if page == "📊 Utilization Analytics":
        st.markdown("<h1 class='main-title'>Executive Utilization Insights</h1>", unsafe_allow_html=True)

        t1, t2 = int(df_compare['Total_Days_M1'].sum()), int(df_compare['Total_Days_M2'].sum())
        diff, growth = t2 - t1, int((t2-t1)/t1*100) if t1 != 0 else 0
        bg_color = "#d1fae5" if diff >= 0 else "#fee2e2"
        text_color = "#047857" if diff >= 0 else "#b91c1c"

        # MEGA KPI SECTION
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"<div class='mega-kpi-container'><p class='mega-label'>Previous Total</p><p class='mega-value'>{t1:,}</p></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='mega-kpi-container'><p class='mega-label'>Current Total</p><p class='mega-value' style='color:#2563eb;'>{t2:,}</p></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"""<div class='mega-kpi-container'>
                <p class='mega-label'>Variance</p>
                <p class='mega-value' style='color:{text_color};'>{diff:+,}</p>
                <div class='mega-delta-tag' style='background:{bg_color}; color:{text_color};'>{growth:+,}% Change</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br><br>", unsafe_allow_html=True)
        
        # --- กราฟที่ 1: กราฟเส้นเปรียบเทียบจำนวนวัน ---
        st.subheader("📈 เปรียบเทียบปริมาณการใช้งานรายแผนก (M1 vs M2)")
        fig_line1 = px.line(df_compare, x='Ward', y=['Total_Days_M1', 'Total_Days_M2'],
                            markers=True, # เปิดให้มีจุดบนเส้น
                            color_discrete_sequence=['#94a3b8', '#0f172a'],
                            labels={'value': 'จำนวนวัน (Days)', 'variable': 'เดือน'})
        
        # ตั้งค่าให้โชว์ตัวเลขบนจุด (lines+markers+text)
        fig_line1.update_traces(mode='lines+markers+text', texttemplate='%{y}', textposition='top center', 
                                textfont_size=18, textfont_weight="bold", marker=dict(size=10, line=dict(width=2, color='white')))
        
        # ใส่ชื่อแกน X และ Y
        fig_line1.update_layout(
            font=dict(size=18, color="#000"), 
            plot_bgcolor='rgba(0,0,0,0)', 
            xaxis=dict(title="แผนก (Ward)", tickfont=dict(size=18, color="#000", weight='bold'), showgrid=True, gridcolor='#f1f5f9'),
            yaxis=dict(title="จำนวนวันการใช้งาน (Total Days)", showgrid=True, gridcolor='#f1f5f9'),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_line1, use_container_width=True)

        # --- กราฟที่ 2: กราฟเส้นอัตราการเติบโต (%) ---
        st.subheader("📈 อัตราการเปลี่ยนแปลงรายแผนก (%)")
        fig_line2 = px.line(df_compare, x='Ward', y='% Growth',
                            markers=True, 
                            color_discrete_sequence=['#2563eb'])
        
        # ตั้งค่าให้โชว์ตัวเลข % บนจุด
        fig_line2.update_traces(mode='lines+markers+text', texttemplate='%{y}%', textposition='top center', 
                                textfont_size=18, textfont_weight="bold", marker=dict(size=10, line=dict(width=2, color='white')))
        
        # ใส่ชื่อแกน X และ Y
        fig_line2.update_layout(
            font=dict(size=18, color="#000"), 
            plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(title="แผนก (Ward)", tickfont=dict(size=18, color="#000", weight='bold'), showgrid=True, gridcolor='#f1f5f9'),
            yaxis=dict(title="อัตราการเติบโต (Growth %)", showgrid=True, gridcolor='#f1f5f9')
        )
        st.plotly_chart(fig_line2, use_container_width=True)

        # ตารางข้อมูล (แก้ applymap เป็น map แล้ว)
        st.markdown("---")
        st.subheader("📋 รายละเอียดข้อมูลเปรียบเทียบ")
        styled_df = df_compare[['Ward', 'Total_Days_M1', 'Total_Days_M2', 'Diff', '% Growth']].style.format(precision=0).map(color_growth, subset=['% Growth'])
        st.dataframe(styled_df, use_container_width=True)

    elif page == "📄 Data Editor":
        excel_2 = pd.ExcelFile(file_2)
        selected_sheet = st.sidebar.selectbox("เลือกวอร์ด:", excel_2.sheet_names)
        st.markdown(f"<h1 class='main-title'>Editor: {selected_sheet}</h1>", unsafe_allow_html=True)
        df_raw = pd.read_excel(file_2, sheet_name=selected_sheet).dropna(how='all')
        edited_df = st.data_editor(df_raw, use_container_width=True, hide_index=True)
        
        # Live สรุปยอดด้านล่าง
        total_val, device_cols = get_safe_total(edited_df)
        if device_cols:
            st.markdown("### 🧮 Live สรุปยอด")
            summary_row = edited_df[device_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum().astype(int)
            m_cols = st.columns(len(device_cols) + 1)
            for i, col in enumerate(device_cols):
                m_cols[i].metric(col, f"{summary_row[col]:,}")
            m_cols[-1].metric("GRAND TOTAL", f"{total_val:,}")
            
            # Export ทุกแผนก
            st.markdown("---")
            if st.button("📥 รวบรวมข้อมูลทุกแผนกเป็นไฟล์เดียว"):
                with st.spinner('กำลังประมวลผล...'):
                    bulk_buf = io.BytesIO()
                    with pd.ExcelWriter(bulk_buf, engine='xlsxwriter') as writer:
                        for s in excel_2.sheet_names:
                            s_df = pd.read_excel(file_2, sheet_name=s).dropna(how='all')
                            _, s_cols = get_safe_total(s_df)
                            if s_cols:
                                for c in s_cols: s_df[c] = pd.to_numeric(s_df[c], errors='coerce').fillna(0).astype(int)
                                s_sum = s_df[s_cols].sum().to_frame().T
                                s_sum.index = [len(s_df)]; s_df = pd.concat([s_df, s_sum])
                                s_df.iloc[-1, 0] = "GRAND TOTAL"
                            s_df.to_excel(writer, sheet_name=s, index=False)
                    st.download_button("📥 คลิกเพื่อดาวน์โหลด", data=bulk_buf.getvalue(), file_name="Hospital_Report.xlsx")

else:
    st.info("กรุณาอัปโหลดไฟล์ Excel ทั้ง 2 เดือนเพื่อเริ่มการวิเคราะห์")
