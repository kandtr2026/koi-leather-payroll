import streamlit as st
import pandas as pd
from core_logic import SalaryCalculator
import os
import io
from datetime import datetime
import plotly.express as px
import calendar
import unicodedata

# Cấu hình file lưu trữ
STAFF_FILE = "employees.csv"

st.set_page_config(page_title="Koi Salary Tool", layout="wide")
st.title("🏯 Salary Calculation Tool - Dự án #Koi")

# --- FUNCTIONS ---
def get_staff_data():
    if os.path.exists(STAFF_FILE):
        df = pd.read_csv(STAFF_FILE)
        if "Group Order" not in df.columns:
            df["Group Order"] = 0
        if "Revenue" not in df.columns:
            df["Revenue"] = 0
        return df
    return pd.DataFrame(columns=['Name', 'Role', 'Base Salary', 'Salary Type', 'Group Order', 'Revenue'])

def save_staff_data(df):
    df.to_csv(STAFF_FILE, index=False)

def normalize_name(name):
    """Chuẩn hóa tên: strip, lower, NFC unicode, bỏ multi-space."""
    s = str(name).strip().lower()
    s = unicodedata.normalize('NFC', s)
    s = ' '.join(s.split())
    return s

def export_individual_salary(emp_name, df_results, df_details):
    """Tạo file Excel chi tiết cho một nhân viên."""
    target_name = str(emp_name).strip()
    target_norm = normalize_name(target_name)
    
    # Lấy dòng tổng hợp
    summary_row = df_results[df_results['Tên'].astype(str).str.strip() == target_name]
    
    # Clone và chuẩn hóa
    df_det = df_details.copy()
    df_det['Name'] = df_det['Name'].astype(str).str.strip()
    df_det['_name_norm'] = df_det['Name'].apply(normalize_name)
    
    # Chuẩn hóa Date: dùng smart_parse rồi normalize
    if pd.api.types.is_datetime64_any_dtype(df_det['Date']):
        df_det['Date'] = df_det['Date'].dt.normalize()
    else:
        df_det['Date'] = SalaryCalculator.smart_parse_dates(df_det['Date']).dt.normalize()
    
    # --- TÌM NHÂN VIÊN ---
    emp_data = df_det[df_det['_name_norm'] == target_norm].copy()
    
    if emp_data.empty:
        emp_data = df_det[df_det['_name_norm'].str.contains(target_norm, na=False, regex=False)].copy()
    
    if emp_data.empty:
        mask = df_det['_name_norm'].apply(lambda x: x in target_norm or target_norm in x)
        emp_data = df_det[mask].copy()
    
    # --- XÁC ĐỊNH THÁNG từ dữ liệu của NV ---
    if not emp_data.empty and not emp_data['Date'].dropna().empty:
        emp_months = emp_data['Date'].dropna().dt.to_period('M')
        ref_period = emp_months.mode().iloc[0]
    elif not df_det['Date'].dropna().empty:
        all_months = df_det['Date'].dropna().dt.to_period('M')
        ref_period = all_months.mode().iloc[0]
    else:
        return None
    
    min_date = ref_period.start_time.normalize()
    last_day = calendar.monthrange(min_date.year, min_date.month)[1]
    max_date = min_date.replace(day=last_day)
    
    # Tạo dải ngày đầy đủ
    all_days = pd.date_range(start=min_date, end=max_date).normalize()
    all_days_df = pd.DataFrame({'Date': all_days})
    
    # Merge
    full_log = pd.merge(all_days_df, emp_data, on='Date', how='left')
    
    # Đảm bảo tất cả cột tồn tại
    required_cols = ['Check-in', 'Check-out', 'IsSunday', 'Late_Min', 'Early_Min', 
                     'OT_Hours', 'Work_Day', 'Daily_Pay', 'Sunday_Bonus', 'Penalty_Amt', 'OT_Amt']
    for col in required_cols:
        if col not in full_log.columns:
            full_log[col] = 0
    
    export_columns = {
        'Date': 'Ngày', 'Check-in': 'Giờ Vào', 'Check-out': 'Giờ Ra',
        'IsSunday': 'Chủ Nhật?', 'Late_Min': 'Trễ (Phút)', 'Early_Min': 'Sớm (Phút)',
        'OT_Hours': 'OT (Giờ)', 'Work_Day': 'Công', 'Daily_Pay': 'Lương Ngày',
        'Sunday_Bonus': 'Thưởng CN', 'Penalty_Amt': 'Tiền Phạt', 'OT_Amt': 'Tiền OT'
    }
    
    display_log = full_log[list(export_columns.keys())].copy()
    display_log.rename(columns=export_columns, inplace=True)
    display_log['Ngày'] = display_log['Ngày'].dt.strftime('%d/%m/%Y')
    
    full_log['IsSunday'] = full_log['Date'].dt.weekday == 6
    display_log['Chủ Nhật?'] = full_log['IsSunday'].apply(lambda x: 'X' if x else '')
    
    num_cols = ['Trễ (Phút)', 'Sớm (Phút)', 'OT (Giờ)', 'Công', 'Lương Ngày', 'Thưởng CN', 'Tiền Phạt', 'Tiền OT']
    for col in num_cols:
        if col in display_log.columns:
            display_log[col] = display_log[col].fillna(0)
    
    display_log['Ghi chú'] = full_log.apply(lambda r: '' if pd.notna(r.get('Check-in')) else 'Vắng/Thiếu dữ liệu', axis=1)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        summary_row.to_excel(writer, index=False, sheet_name="Tổng hợp lương")
        display_log.to_excel(writer, index=False, sheet_name="Bảng chấm công chi tiết")
        
        workbook = writer.book
        worksheet = writer.sheets["Bảng chấm công chi tiết"]
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
        for col_num, value in enumerate(display_log.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
    return output.getvalue()

# --- TABS ---
tab1, tab2, tab3 = st.tabs(["🚀 Tính Lương", "👥 Cấu hình Nhân viên", "📊 Thống kê & Biểu đồ"])

with tab2:
    st.header("1. Cấu hình Danh sách Nhân viên")
    df_staff = get_staff_data()
    
    st.write("Nhập thông tin nhân viên tại đây. 'Group Order' dùng để sắp xếp, 'Revenue' để tính hoa hồng cho Saleman.")
    edited_staff = st.data_editor(
        df_staff, 
        num_rows="dynamic", 
        use_container_width=True, 
        key="staff_editor",
        column_config={
            "Group Order": st.column_config.NumberColumn("Group Order", help="Thứ tự hiển thị", min_value=0, step=1),
            "Revenue": st.column_config.NumberColumn("Revenue (Doanh thu)", help="Doanh thu bán hàng của Saleman trong tháng", min_value=0, format="%d")
        }
    )
    
    if st.button("💾 Lưu thay đổi"):
        save_staff_data(edited_staff)
        st.success("Đã cập nhật danh sách nhân viên!")

with tab1:
    st.header("2. Tính toán Lương từ Hanet")
    
    if get_staff_data().empty:
        st.error("⚠️ Vui lòng cấu hình danh sách nhân viên ở tab bên cạnh trước!")
    else:
        c1, c2 = st.columns([1, 2])
        
        with c1:
            st.write("👉 **Hướng dẫn:** Bấm vào ô đầu tiên (0,0), sau đó nhấn **Ctrl+V** để dán toàn bộ bảng từ Excel Hanet.")
            
            df_to_process = None

            if 'paste_buffer' not in st.session_state:
                cols = ['ID', 'Tên', 'Chức vụ', 'Phòng ban', 'Mã NV']
                for i in range(1, 96): 
                    cols.append(f"Cột {i}")
                final_cols = cols[:100]
                st.session_state['paste_buffer'] = pd.DataFrame(
                    [["" for _ in range(100)] for _ in range(200)],
                    columns=final_cols
                )
            
            df_pasted = st.data_editor(
                st.session_state['paste_buffer'],
                num_rows="dynamic",
                use_container_width=True,
                height=400,
                key="excel_paste_editor",
                column_config={
                    "ID": st.column_config.TextColumn("ID", help="Copy cột ID từ Excel vào đây", required=True),
                    "Tên": st.column_config.TextColumn("Tên", help="Tên nhân viên", required=True),
                    "Mã NV": st.column_config.TextColumn("Mã NV"),
                }
            )
            
            if st.button("📥 Xác nhận dữ liệu đã dán"):
                df_clean = df_pasted.replace('', pd.NA).dropna(how='all', axis=0)
                df_clean = df_clean.dropna(how='all', axis=1)
                
                if not df_clean.empty:
                    try:
                        df_to_process = SalaryCalculator.process_dataframe(df_clean)
                        st.session_state['temp_preview'] = df_to_process
                        st.success(f"✅ Đã nhận thành công {len(df_to_process)} dòng chấm công!")

                        anomalies = df_to_process[
                            (df_to_process['Check-in'].astype(str) == df_to_process['Check-out'].astype(str)) &
                            (df_to_process['Check-in'].astype(str).str.strip().isin(['', '-', '0:00', '00:00']) == False)
                        ]
                        
                        if not anomalies.empty:
                            st.warning("⚠️ **Phát hiện sai sót chấm công:** Một số ngày có giờ Vào và Giờ Ra trùng nhau.")
                            df_ano_display = anomalies[['Name', 'Date', 'Check-in', 'Check-out']].copy()
                            df_ano_display.columns = ['Nhân viên', 'Ngày', 'Vào', 'Ra']
                            st.dataframe(df_ano_display, use_container_width=True)
                            st.info("💡 Bạn nên kiểm tra lại các trường hợp trên trước khi tính lương.")
                        
                        df_check_hours = df_to_process.copy()
                        def quick_hour_check(row):
                            try:
                                t_in = pd.to_datetime(row['Check-in'], format='%H:%M').time()
                                t_out = pd.to_datetime(row['Check-out'], format='%H:%M').time()
                                dt_in = datetime.combine(datetime.min, t_in)
                                dt_out = datetime.combine(datetime.min, t_out)
                                hours = (dt_out - dt_in).total_seconds() / 3600
                                if hours > 5:
                                    hours -= 1
                                return hours
                            except:
                                return 0

                        df_check_hours['Hrs'] = df_check_hours.apply(quick_hour_check, axis=1)
                        
                        short_shifts = df_check_hours[
                            (df_check_hours['Hrs'] > 0) & (df_check_hours['Hrs'] < 8) &
                            (df_check_hours['Check-in'].astype(str) != df_check_hours['Check-out'].astype(str))
                        ]

                        if not short_shifts.empty:
                            st.warning("⚠️ **Cảnh báo làm thiếu giờ (< 8h):**")
                            df_short_display = short_shifts[['Name', 'Date', 'Check-in', 'Check-out', 'Hrs']].copy()
                            df_short_display.columns = ['Nhân viên', 'Ngày', 'Vào', 'Ra', 'Số giờ tính được']
                            st.dataframe(df_short_display, use_container_width=True)
                    
                    except Exception as e:
                        st.error(f"❌ Lỗi xử lý: {str(e)}")
                else:
                    st.warning("Bảng hiện đang trống.")

            std_days = st.number_input("Số ngày công chuẩn (VD: 26)", value=26, min_value=1)
            
            if st.button("🚀 Bắt đầu Tính toán", type="primary"):
                if 'temp_preview' in st.session_state:
                    try:
                        calc = SalaryCalculator(st.session_state['temp_preview'], get_staff_data(), std_days)
                        calc.process_timeintervals()
                        st.session_state['salary_results'] = calc.calculate_monthly_salary()
                        st.session_state['salary_details'] = calc.df_final
                        st.success("Tính toán hoàn tất!")
                    except Exception as e:
                        st.error(f"Lỗi tính toán: {str(e)}")
                else:
                    st.warning("Vui lòng nhập dữ liệu trước!")

        with c2:
            if 'temp_preview' in st.session_state:
                st.subheader("🔍 Xem trước Dữ liệu Chấm công (Raw)")
                st.write("Hệ thống đã trích xuất dữ liệu như sau. Hãy kiểm tra Cột 'Name' và 'Date' xem có đúng không.")
                st.dataframe(st.session_state['temp_preview'], use_container_width=True, height=400)
                
                if st.checkbox("Hiển thị thống kê nhanh"):
                    st.write(st.session_state['temp_preview'].groupby('Name').size().reset_index(name='Số ngày chấm công'))


    # Hiển thị kết quả
    if 'salary_results' in st.session_state:
        st.divider()
        st.header("3. Kết quả Bảng lương")
        
        df_staff_info = get_staff_data()[['Name', 'Group Order']].rename(columns={'Name': 'Tên'}).drop_duplicates()
        df_results = pd.merge(st.session_state['salary_results'], df_staff_info, on='Tên', how='left')
        df_results = df_results.sort_values(by=['Group Order', 'Tên']).reset_index(drop=True)
        
        df_display = df_results.copy()
        money_cols = ["Lương Cơ Bản", "Doanh Thu", "Hoa Hồng", "Lương Ngày Thường", "Lương Chủ Nhật", "Tiền OT", "Phạt (Trễ/Sớm)", "Tổng Thực Lãnh"]
        
        for col in money_cols:
            if col in df_display.columns:
                df_display[col] = df_display[col].apply(lambda x: f"{int(x):,d}".replace(",", ".") if pd.notna(x) else "0")
        
        metric_col = "Ngày công/Giờ công"
        if metric_col in df_display.columns:
            df_display[metric_col] = df_display[metric_col].apply(lambda x: f"{x:.1f}" if pd.notna(x) else "0.0")

        if not df_display.empty:
            groups = sorted(df_display['Group Order'].unique())
            
            for g in groups:
                g_label = f"Nhóm thứ tự: {int(g)}" if g > 0 else "Nhóm chưa phân loại"
                st.markdown(f"---")
                st.markdown(f"### 📍 {g_label}")
                
                df_group = df_display[df_display['Group Order'] == g].reset_index(drop=True)
                is_sales_group = any(df_group['Chức vụ'] == 'Saleman')

                if is_sales_group:
                    col_widths = [1.5, 1.2, 1.2, 1.2, 1.2, 1.2, 1.2, 1.2, 1.0, 1.2, 0.8]
                    fields = ["Tên", "Chức vụ", "Lương CB", "Doanh thu", "Hoa hồng", "Công/Giờ", "Lương", "Lương CN", "Tiền OT", "Phạt", "Chi tiết"]
                else:
                    col_widths = [1.5, 1.2, 1.5, 1.2, 1.5, 1.2, 1.2, 1.5, 0.8]
                    fields = ["Tên", "Chức vụ", "Lương CB", "Công/Giờ", "Lương", "Lương CN", "Tiền OT", "Phạt", "Chi tiết"]
                
                header_cols = st.columns(col_widths)
                for i, f in enumerate(fields):
                    header_cols[i].markdown(f"<p style='font-size: 0.75em; font-weight: bold; color: gray;'>{f}</p>", unsafe_allow_html=True)
                
                for idx, row in df_group.iterrows():
                    row_cols = st.columns(col_widths)
                    
                    if is_sales_group:
                        row_cols[0].write(row['Tên'])
                        row_cols[1].write(row['Chức vụ'])
                        row_cols[2].write(row['Lương Cơ Bản'])
                        row_cols[3].write(row['Doanh Thu'])
                        row_cols[4].write(row['Hoa Hồng'])
                        row_cols[5].write(row['Ngày công/Giờ công'])
                        row_cols[6].write(row['Lương Ngày Thường'])
                        row_cols[7].write(row['Lương Chủ Nhật'])
                        row_cols[8].write(row['Tiền OT'])
                        row_cols[9].write(row['Phạt (Trễ/Sớm)'])
                        btn_idx = 10
                    else:
                        row_cols[0].write(row['Tên'])
                        row_cols[1].write(row['Chức vụ'])
                        row_cols[2].write(row['Lương Cơ Bản'])
                        row_cols[3].write(row['Ngày công/Giờ công'])
                        row_cols[4].write(row['Lương Ngày Thường'])
                        row_cols[5].write(row['Lương Chủ Nhật'])
                        row_cols[6].write(row['Tiền OT'])
                        row_cols[7].write(row['Phạt (Trễ/Sớm)'])
                        btn_idx = 8
                    
                    with row_cols[btn_idx]:
                        excel_data = export_individual_salary(
                            row['Tên'], 
                            st.session_state['salary_results'], 
                            st.session_state['salary_details']
                        )
                        if excel_data:
                            st.download_button(
                                label="📄",
                                data=excel_data,
                                file_name=f"Chi_tiet_{row['Tên']}_{datetime.now().strftime('%Y%m')}.xlsx",
                                mime="application/vnd.ms-excel",
                                key=f"dl_{row['Tên']}_{idx}_{g}"
                            )
                        else:
                            st.write("⚠️")
                    st.markdown(f"<div style='text-align: right; font-weight: bold; color: #ff4b4b;'>Tổng Lãnh: {row['Tổng Thực Lãnh']}</div>", unsafe_allow_html=True)
                    st.divider()
            
            # --- BÁO CÁO TỔNG KẾT ---
            st.markdown("## 📈 Báo cáo Tổng kết Chi phí")
            
            grand_total_base = df_results['Lương Ngày Thường'].sum()
            grand_total_sunday = df_results['Lương Chủ Nhật'].sum()
            grand_total_ot = df_results['Tiền OT'].sum()
            grand_total_penalty = df_results['Phạt (Trễ/Sớm)'].sum()
            grand_total_income = df_results['Tổng Thực Lãnh'].sum()
            grand_total_commission = df_results['Hoa Hồng'].sum() if 'Hoa Hồng' in df_results.columns else 0

            m1, m2, m3 = st.columns(3)
            with m1:
                st.metric("TỔNG QUỸ LƯƠNG TRẢ", f"{grand_total_income:,.0f} đ")
            with m2:
                st.metric("Tổng chi OT", f"{grand_total_ot:,.0f} đ")
            with m3:
                st.metric("Tổng tiền phạt", f"{grand_total_penalty:,.0f} đ", delta_color="inverse")

            summary_data = {
                "Hạng mục chi phí": [
                    "1. Lương ngày công thường",
                    "2. Thưởng làm ngày Chủ Nhật (100% bonus)",
                    "3. Tiền tăng ca (OT)",
                    "4. Tiền hoa hồng doanh thu (Saleman)",
                    "5. Khấu trừ đi trễ / về sớm",
                    "TỔNG CỘNG THỰC CHI"
                ],
                "Số tiền (VNĐ)": [
                    f"{grand_total_base:,.0f}",
                    f"{grand_total_sunday:,.0f}",
                    f"{grand_total_ot:,.0f}",
                    f"{grand_total_commission:,.0f}",
                    f"-{grand_total_penalty:,.0f}",
                    f"**{grand_total_income:,.0f}**"
                ]
            }
            st.table(pd.DataFrame(summary_data))
            
            output_total = io.BytesIO()
            with pd.ExcelWriter(output_total, engine='xlsxwriter') as writer:
                df_results.to_excel(writer, index=False, sheet_name="Bang_Luong_Tong_Hop")
            
            st.download_button(
                label="📥 Tải xuống Bảng lương Tổng hợp (Tất cả nhân viên)",
                data=output_total.getvalue(),
                file_name=f"Bang_luong_tong_hop_{datetime.now().strftime('%Y%m')}.xlsx",
                mime="application/vnd.ms-excel",
                use_container_width=True
            )
        else:
            st.info("Chưa có kết quả tính toán.")
            
    with tab3:
        st.header("📊 Phân tích Chi phí Lương")
        
        if 'salary_results' in st.session_state and not st.session_state['salary_results'].empty:
            df_res = st.session_state['salary_results']
            df_det = st.session_state['salary_details']
            
            col_chart1, col_chart2 = st.columns(2)
            
            with col_chart1:
                st.subheader("💰 Phân bổ theo Chức vụ")
                fig_role = px.pie(
                    df_res, values='Tổng Thực Lãnh', names='Chức vụ',
                    hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel
                )
                fig_role.update_layout(margin=dict(t=0, b=0, l=0, r=0))
                st.plotly_chart(fig_role, use_container_width=True)
            
            with col_chart2:
                st.subheader("🧩 Cơ cấu Thu nhập")
                comp_data = []
                for _, row in df_res.iterrows():
                    comp_data.append({'Nhân viên': row['Tên'], 'Loại': 'Lương cứng', 'Số tiền': row['Lương Ngày Thường']})
                    comp_data.append({'Nhân viên': row['Tên'], 'Loại': 'Tiền OT', 'Số tiền': row['Tiền OT']})
                    comp_data.append({'Nhân viên': row['Tên'], 'Loại': 'Thưởng CN', 'Số tiền': row['Lương Chủ Nhật']})
                    if 'Hoa Hồng' in row:
                        comp_data.append({'Nhân viên': row['Tên'], 'Loại': 'Hoa hồng', 'Số tiền': row['Hoa Hồng']})
                
                df_comp = pd.DataFrame(comp_data)
                fig_comp = px.bar(
                    df_comp, x='Nhân viên', y='Số tiền', color='Loại',
                    barmode='stack', color_discrete_sequence=px.colors.qualitative.Safe
                )
                fig_comp.update_layout(xaxis_tickangle=-45, margin=dict(t=20, b=0, l=0, r=0))
                st.plotly_chart(fig_comp, use_container_width=True)
            
            st.subheader("📅 Biến động chi phí theo ngày")
            df_det_copy = df_det.copy()
            df_det_copy['Date'] = pd.to_datetime(df_det_copy['Date'])
            df_daily = df_det_copy.groupby('Date').agg({
                'Daily_Pay': 'sum', 'OT_Amt': 'sum', 'Sunday_Bonus': 'sum'
            }).reset_index()
            
            df_daily['Tổng chi phí'] = df_daily['Daily_Pay'] + df_daily['OT_Amt'] + df_daily['Sunday_Bonus']
            df_daily['Ngày'] = df_daily['Date'].dt.strftime('%d/%m')
            
            fig_daily = px.area(
                df_daily, x='Ngày', y='Tổng chi phí',
                labels={'Tổng chi phí': 'Số tiền (VNĐ)'},
                color_discrete_sequence=['#ff4b4b']
            )
            fig_daily.update_layout(hovermode="x unified")
            st.plotly_chart(fig_daily, use_container_width=True)
            
            st.divider()
            m1, m2, m3, m4 = st.columns(4)
            total_fund = df_res['Tổng Thực Lãnh'].sum()
            avg_salary = df_res['Tổng Thực Lãnh'].mean()
            total_ot = df_res['Tiền OT'].sum()
            total_staff = len(df_res)
            
            m1.metric("Tổng quỹ lương", f"{total_fund:,.0f} đ")
            m2.metric("Số lượng nhân sự", f"{total_staff} người")
            m3.metric("Trung bình/người", f"{avg_salary:,.0f} đ")
            m4.metric("Tổng chi phí OT", f"{total_ot:,.0f} đ", delta=f"{(total_ot/total_fund*100):.1f}%")

        else:
            st.info("Hãy thực hiện tính toán lương ở Tab 1 để xem các biểu đồ phân tích.")
        
        with st.expander("💡 Giải thích ký hiệu và công thức"):
            st.info("""
            - **Giờ OT**: Tính từ sau 18:30 (áp dụng cho Thợ thủ công). Hệ số x1.2.
            - **Chủ Nhật**: Tính 200% đơn giá ngày cho bộ phận Sản xuất.
            - **Lương Giờ**: Được quy đổi từ Lương tháng / 26 ngày / 8 giờ.
            - **Saleman**: Tính lương trực tiếp = (Giờ Out - Giờ In) * Đơn giá giờ.
            """)
        
        if 'salary_results' in st.session_state and not st.session_state['salary_results'].empty:
            df_staff_info = get_staff_data()[['Name', 'Group Order']].rename(columns={'Name': 'Tên'}).drop_duplicates()
            df_final_to_save = pd.merge(st.session_state['salary_results'], df_staff_info, on='Tên', how='left')
            df_final_to_save = df_final_to_save.sort_values(by=['Group Order', 'Tên'])

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final_to_save.to_excel(writer, index=False, sheet_name="Bảng Tổng Hợp")
                if 'salary_details' in st.session_state:
                    st.session_state['salary_details'].to_excel(writer, index=False, sheet_name="Chi Tiết Theo Ngày")
            
            st.download_button(
                label="📥 Tải xuống Bảng lương (Excel)",
                data=output.getvalue(),
                file_name=f"Bang_Luong_{datetime.now().strftime('%Y%m')}.xlsx",
                mime="application/vnd.ms-excel"
            )
