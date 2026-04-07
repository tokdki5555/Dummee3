import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

# 1. Config & Premium Luxury Styling
st.set_page_config(page_title="Executive Device Analytics", page_icon="🏥", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f4f7f9; }
    .main-title { 
        color: #1a3a5f; font-family: 'Sarabun', sans-serif; 
        font-weight: 800; text-align: center; margin-bottom: 30px;
        letter-spacing: 1px;
    }
    /* Luxury Card Style */
    .kpi-box {
        background: white; padding: 25px; border-radius: 20px; text-align: center;
        box-shadow: 0 10px 25px rgba(0,0,0,0.05);
        border-top: 5px solid #1a3a5f;
        transition: 0.3s;
    }
    .kpi-box:hover { transform: translateY(-5px); box-shadow: 0 15px 35px rgba(0,0,0,0.1); }
    
    /* Button Styling */
    .stButton>button {
        border-radius: 15px; background: linear-gradient(135deg, #1a3a5f 0%, #2c5282 100%);
        color: white; border: none; height: 3.5em; font-weight: bold; width: 100%;
        box-shadow: 0 4px 15px rgba(26, 58, 95, 0.2);
    }
    </style>
    """, unsafe_allow_html=True)

# --- Sidebar ---
st.sidebar.markdown("<h2 style='text-align:center; color:#1a3a5f;'>🏥 Device JAN 5400</h2>", unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("📂 อัปโหลดไฟล์ Excel", type=["xlsx"])

if uploaded_file:
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_names = excel_file.sheet_names
    selected_sheet = st.sidebar.selectbox("เลือกแผนก (Ward):", sheet_names)
    
    st.sidebar.markdown("---")
    page = st.sidebar.radio("Navigation:", ["📄 Data Editor", "📊 Executive Analytics"])

    keywords = ["Ventilator", "Foley", "Central line", "Port A Cath"]

    def get_safe_total(df_in):
        d_cols = [c for c in df_in.columns if any(k.lower() in str(c).lower() for k in keywords)]
        if not d_cols: return 0, []
        df_c = df_in.copy()
        for c in d_cols: df_c[c] = pd.to_numeric(df_c[c], errors='coerce').fillna(0)
        return df_c[d_cols].values.sum(), d_cols

    # --- หน้า 1: Preview & Edit ---
    if page == "📄 Data Editor":
        st.markdown(f"<h1 class='main-title'>📄 แผนก: {selected_sheet}</h1>", unsafe_allow_html=True)
        
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet).dropna(how='all')
        for col in df.columns:
            if 'date' in col.lower():
                try: df[col] = pd.to_datetime(df[col]).dt.date
                except: pass
        
        edited_df = st.data_editor(df, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        total_val, device_cols = get_safe_total(edited_df)
        
        # แสดงผลสรุปรายอุปกรณ์ในหน้านี้
        st.subheader("📊 สรุปยอดอุปกรณ์รายแผนก")
        if device_cols:
            cols_grid = st.columns(len(device_cols))
            sum_vals = edited_df[device_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum()
            for i, col_name in enumerate(device_cols):
                with cols_grid[i % len(device_cols)]:
                    st.metric(label=col_name, value=f"{int(sum_vals[col_name]):,}")
        
        st.markdown("---")
        if st.button("📥 Export All Wards (รวบรวมทุกแผนกพร้อมยอดรวม)"):
            with st.spinner('กำลังจัดเตรียมไฟล์...'):
                all_dfs = {}
                for s in sheet_names:
                    df_s = pd.read_excel(uploaded_file, sheet_name=s).dropna(how='all')
                    v, cols = get_safe_total(df_s)
                    if cols:
                        for c in cols: df_s[c] = pd.to_numeric(df_s[c], errors='coerce').fillna(0)
                        t_row = df_s[cols].sum().to_frame().T
                        t_row.index = [len(df_s)]
                        df_final = pd.concat([df_s, t_row], axis=0)
                        df_final.iloc[-1, 0] = "GRAND TOTAL"
                        all_dfs[s] = df_final
                    else: all_dfs[s] = df_s
                
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                    for s, data in all_dfs.items(): data.to_excel(writer, sheet_name=s, index=False)
                st.download_button("📥 คลิกเพื่อดาวน์โหลดรายงานสมบูรณ์", data=buf.getvalue(), file_name="Full_Device_Report.xlsx")

    # --- หน้า 2: Dashboard ---
    elif page == "📊 Executive Analytics":
        st.markdown("<h1 class='main-title'>📊 Executive Summary Dashboard</h1>", unsafe_allow_html=True)
        
        ward_data = []
        for s in sheet_names:
            df_t = pd.read_excel(uploaded_file, sheet_name=s).dropna(how='all')
            total_sum, _ = get_safe_total(df_t)
            ward_data.append({'Ward': s, 'Total_Days': total_sum})
        
        df_stats = pd.DataFrame(ward_data)
        grand_total = df_stats['Total_Days'].sum()
        avg_per_ward = df_stats['Total_Days'].mean()
        
        if grand_total > 0:
            df_stats['Proportion_%'] = (df_stats['Total_Days'] / grand_total * 100).round(1)
            
            # --- KPI Cards สไตล์ใหม่ (เพิ่มมิติ) ---
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f"<div class='kpi-box'><p style='color:#6c757d; font-weight:bold;'>GRAND TOTAL</p><h1 style='color:#1a3a5f;'>{int(grand_total):,}</h1><p style='font-size:0.8em; color:#adb5bd;'>วันรวมทุกวอร์ด</p></div>", unsafe_allow_html=True)
            with c2:
                st.markdown(f"<div class='kpi-box' style='border-top-color:#4a90e2;'><p style='color:#6c757d; font-weight:bold;'>AVERAGE / WARD</p><h1 style='color:#4a90e2;'>{avg_per_ward:.1f}</h1><p style='font-size:0.8em; color:#adb5bd;'>ค่าเฉลี่ยภาระงาน</p></div>", unsafe_allow_html=True)
            with c3:
                max_w = df_stats.loc[df_stats['Total_Days'].idxmax()]
                st.markdown(f"<div class='kpi-box' style='border-top-color:#f39c12;'><p style='color:#6c757d; font-weight:bold;'>TOP USAGE WARD</p><h1 style='color:#f39c12;'>{max_w['Ward']}</h1><p style='font-size:0.8em; color:#adb5bd;'>สัดส่วน {max_w['Proportion_%']}%</p></div>", unsafe_allow_html=True)

            # Charts
            col_l, col_r = st.columns(2)
            with col_l:
                fig_bar = px.bar(df_stats, x='Ward', y='Total_Days', color='Total_Days', color_continuous_scale='Blues',
                                 text_auto='.2s', title="Total Days by Ward")
                st.plotly_chart(fig_bar, use_container_width=True)
            with col_r:
                fig_pie = px.pie(df_stats, values='Total_Days', names='Ward', hole=0.5, 
                                 color_discrete_sequence=px.colors.qualitative.Pastel, title="Usage Distribution (%)")
                st.plotly_chart(fig_pie, use_container_width=True)

            # --- ปุ่ม Export Dashboard ที่มีค่าสถิติรวม ---
            st.markdown("---")
            st.subheader("📥 Download Summary Data")
            
            # สร้าง DataFrame พิเศษสำหรับ Export Dashboard
            df_export = df_stats.copy()
            # เพิ่มแถว Grand Total และ Average ลงไปในไฟล์ Excel
            summary_rows = pd.DataFrame([
                {'Ward': '---', 'Total_Days': None, 'Proportion_%': None},
                {'Ward': 'GRAND TOTAL', 'Total_Days': grand_total, 'Proportion_%': 100.0},
                {'Ward': 'AVERAGE PER WARD', 'Total_Days': avg_per_ward, 'Proportion_%': ''}
            ])
            df_final_export = pd.concat([df_export, summary_rows], ignore_index=True)
            
            buf_sum = io.BytesIO()
            df_final_export.to_excel(buf_sum, index=False)
            
            st.download_button(
                label="📥 Export Dashboard Summary ",
                data=buf_sum.getvalue(),
                file_name=f"Dashboard_Summary_{datetime.now().strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.ms-excel"
            )
        else:
            st.warning("กรุณาอัปโหลดไฟล์ที่มีข้อมูลตัวเลขในคอลัมน์อุปกรณ์")

else:
    st.markdown("<div style='text-align:center; margin-top:100px;'><h1>🏦 Smart Device JAN 6500</h1><p>Luxury Web-based Data Management System</p></div>", unsafe_allow_html=True)
