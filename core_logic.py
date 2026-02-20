import pandas as pd
import numpy as np
from datetime import datetime, time

class SalaryCalculator:
    """
    Logic xử lý lương cho dự án #koi.
    Sử dụng Tên nhân viên (Name) làm khóa chính để so khớp.
    """
    def __init__(self, checkin_df, staff_df, standard_days=26):
        self.df_checkin = checkin_df
        self.df_staff = staff_df
        self.standard_days = standard_days
        
        # Merge dữ liệu chấm công với danh sách nhân viên
        self.df = pd.merge(self.df_checkin, self.df_staff, on='Name', how='left')
        self.df_final = None

    @staticmethod
    def parse_pasted_text(text):
        """
        Xử lý dữ liệu văn bản được dán từ Hanet/Excel (thường là Tab-separated).
        """
        import io
        if not text or len(text.strip()) == 0:
            return pd.DataFrame()
        
        try:
            # Ưu tiên đọc dạng Tab (vì copy từ Excel/Web thường là Tab)
            df = pd.read_csv(io.StringIO(text), sep='\t', header=None)
            # Nếu chỉ có 1 cột, thử đọc bằng dấu phẩy
            if df.shape[1] == 1:
                df = pd.read_csv(io.StringIO(text), sep=',', header=None)
        except:
            return pd.DataFrame()
            
        return df

    @staticmethod
    def parse_hanet_wide(file_path):
        """Đọc từ file Excel."""
        if file_path is None: return pd.DataFrame()
        try:
            df_raw = pd.read_excel(file_path, header=None)
            return SalaryCalculator.process_dataframe(df_raw)
        except Exception as e:
            raise ValueError(f"Lỗi đọc file Excel: {e}")

    @staticmethod
    def process_dataframe(df_raw):
        """Hàm xử lý bảng dữ liệu (Pasted hoặc Excel)."""
        if df_raw is None or len(df_raw) == 0:
            return pd.DataFrame()

        # 1. Tìm dòng tiêu đề
        header_row_idx = -1
        header_col_idx = 0
        for i in range(min(50, len(df_raw))):
            for j in range(min(15, len(df_raw.columns))):
                val = str(df_raw.iloc[i, j]).strip().upper()
                if val in ["ID", "MÃ NV", "EMPLOYEE ID", "MÃ NHÂN VIÊN", "MÃ_NV", "STT", "TÊN"]:
                    header_row_idx = int(i)
                    header_col_idx = int(j)
                    break
            if header_row_idx != -1: break
        
        if header_row_idx == -1:
             header_row_idx = 0
             header_col_idx = 0

        # 2. Định vị cột ngày tháng và tạo bản đồ mapping
        start_time_col = -1
        for j in range(len(df_raw.columns)):
            val = str(df_raw.iloc[header_row_idx, j])
            if any(char.isdigit() for char in val) and ('-' in val or '/' in val):
                start_time_col = int(j)
                break
        
        if start_time_col == -1:
             start_time_col = int(header_col_idx) + 5 
             
        if start_time_col >= len(df_raw.columns):
             start_time_col = 5

        col_to_date = {}
        last_date = None
        
        # Thử lấy dòng chứa ngày
        row_temp = df_raw.iloc[header_row_idx].values
        if header_row_idx > 0:
            check_val = str(row_temp[min(int(start_time_col), len(row_temp)-1)])
            if not any(char.isdigit() for char in check_val):
                row_temp = df_raw.iloc[header_row_idx - 1].values

        for j in range(int(start_time_col), len(df_raw.columns)):
            val_d = row_temp[j]
            if pd.notna(val_d) and str(val_d).strip() not in ['', '-', 'nan', 'NaN']:
                if isinstance(val_d, datetime):
                    last_date = val_d.strftime("%Y-%m-%d")
                else:
                    last_date = str(val_d).strip().replace('/', '-')
            col_to_date[int(j)] = last_date

        # 3. Duyệt dữ liệu nhân viên
        long_data = []
        name_col = -1
        for j in range(int(header_col_idx), min(int(header_col_idx) + 5, len(df_raw.columns))):
            if "TÊN" in str(df_raw.iloc[header_row_idx, j]).upper():
                name_col = int(j)
                break
        if name_col == -1: name_col = int(header_col_idx) + 1

        for r in range(header_row_idx + 1, len(df_raw)):
            row_data = df_raw.iloc[r]
            if len(row_data) <= name_col: continue
            emp_name = str(row_data[name_col]).strip()
            if not emp_name or emp_name.lower() in ['nan', 'none', '']: continue
            
            # Duyệt các cặp cột (Vào - Ra)
            for c in range(int(start_time_col), len(df_raw.columns) - 1, 2):
                current_date_str = col_to_date.get(int(c))
                
                check_in = row_data[int(c)]
                check_out = row_data[int(c) + 1]
                
                # Hàm kiểm tra ô có dữ liệu thật sự không
                def is_valid(v):
                    if pd.isna(v): return False
                    s = str(v).strip().lower()
                    return s not in ['', '-', 'nan', 'none', '0:00', '00:00']

                if is_valid(check_in) or is_valid(check_out):
                    if not current_date_str: continue 
                    long_data.append({
                        'Name': emp_name,
                        'Date': current_date_str,
                        'Check-in': check_in,
                        'Check-out': check_out
                    })
        
        if not long_data:
            raise ValueError("⚠️ Dữ liệu rỗng hoặc không đúng định dạng. Hãy copy toàn bộ bảng bao gồm cả tiêu đề.")

        return pd.DataFrame(long_data)

    def process_timeintervals(self):
        """Bước tính toán chi tiết từng cặp giờ Vào-Ra."""
        # --- CẤU HÌNH THỜI GIAN ---
        START_TIME = time(9, 0)
        END_TIME = time(18, 0)
        LUNCH_START = time(12, 0)
        LUNCH_END = time(13, 0)
        OT_START = time(18, 30)

        def safe_time_parse(t):
            if pd.isna(t) or str(t).strip() in ['', '-', 'nan', 'NaN']: return None
            if isinstance(t, time): return t
            if isinstance(t, datetime): return t.time()
            
            t_str = str(t).strip()
            # Thử các định dạng phổ biến
            formats = ["%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M:%S %p", "%H:%M "]
            for fmt in formats:
                try:
                    return datetime.strptime(t_str, fmt).time()
                except:
                    continue
            
            # Trường hợp đặc biệt: 9:0 -> 09:00
            try:
                if ':' in t_str:
                    parts = t_str.split(':')
                    h = int(parts[0])
                    m = int(parts[1])
                    return time(h, m)
            except:
                pass
                
            return None

        # Pre-processing
        # Ép buộc dùng dayfirst=True để ưu tiên định dạng Ngày/Tháng/Năm (Việt Nam)
        self.df['Date'] = pd.to_datetime(self.df['Date'], dayfirst=True, errors='coerce')
        self.df['IsSunday'] = self.df['Date'].dt.weekday == 6
        self.df['InTime'] = self.df['Check-in'].apply(safe_time_parse)
        self.df['OutTime'] = self.df['Check-out'].apply(safe_time_parse)

        # Kiểm tra khớp nhân viên
        if self.df['Role'].isna().all():
            raise ValueError("⚠️ Không có nhân viên nào trong file khớp với danh sách quản lý!\nHành động: Hãy kiểm tra lại cột Tên (Name) ở cả 2 bên.")

        results = []
        for _, row in self.df.iterrows():
            role = row['Role']
            in_t = row['InTime']
            out_t = row['OutTime']
            base = row['Base Salary']
            s_type = row['Salary Type']
            
            # Mặc định kết quả
            res = {
                'Late_Min': 0, 'Early_Min': 0, 'OT_Hours': 0, 'Work_Hours': 0, 
                'Work_Day': 0, 'Penalty_Amt': 0, 'OT_Amt': 0, 'Daily_Pay': 0,
                'Sunday_Bonus': 0 # Cột mới để chứa phần tiền cộng thêm
            }

            if in_t is None or out_t is None:
                results.append(res)
                continue

            # 1. Tính Giờ làm việc
            t_in = datetime.combine(datetime.min, in_t)
            t_out = datetime.combine(datetime.min, out_t)
            duration = (t_out - t_in).total_seconds() / 3600
            
            # Chỉ trừ nghỉ trưa nếu KHÔNG PHẢI Saleman
            if role != 'Saleman':
                if in_t < LUNCH_START and out_t > LUNCH_END:
                    duration -= 1 # Trừ 1 tiếng nghỉ trưa
                    
            res['Work_Hours'] = max(0, duration)

            # 2. Tính Đi trễ / Về sớm (Áp dụng cho Thợ và Marketing)
            if role in ['Thợ thủ công', 'Thợ thủ công tập sự', 'Marketing']:
                # Đặc cách cho Tường Photo bắt đầu lúc 14h
                actual_start = time(14, 0) if "Tường Photo" in str(row['Name']) else START_TIME
                
                if in_t > actual_start:
                    res['Late_Min'] = (datetime.combine(datetime.min, in_t) - datetime.combine(datetime.min, actual_start)).total_seconds() / 60
                if out_t < END_TIME:
                    res['Early_Min'] = (datetime.combine(datetime.min, END_TIME) - datetime.combine(datetime.min, out_t)).total_seconds() / 60

            # 3. Tính OT (Chỉ tính nếu > 18:30 và không phải CN)
            if role in ['Thợ thủ công', 'Thợ thủ công tập sự'] and not row['IsSunday']:
                if out_t > OT_START:
                    res['OT_Hours'] = (datetime.combine(datetime.min, out_t) - datetime.combine(datetime.min, OT_START)).total_seconds() / 3600

            # 4. Quy đổi tiền
            hourly_rate = 0
            if s_type == 'Tháng': hourly_rate = base / self.standard_days / 8
            elif s_type == 'Ngày': hourly_rate = base / 8
            elif s_type == 'Giờ': hourly_rate = base
            
            res['Penalty_Amt'] = (res['Late_Min'] + res['Early_Min']) * (hourly_rate / 60)
            res['OT_Amt'] = res['OT_Hours'] * hourly_rate * 1.2
            
            # Tính lương ngày và thưởng Chủ Nhật
            res['Work_Day'] = 1 # Luôn đếm là 1 công hễ có đi làm
            
            # Lương gốc cơ bản cho 1 ngày công
            if s_type == 'Tháng':
                res['Daily_Pay'] = base / self.standard_days
            elif s_type == 'Ngày':
                res['Daily_Pay'] = base
            elif s_type == 'Giờ':
                res['Daily_Pay'] = res['Work_Hours'] * base

            # Nếu là Chủ Nhật, tính phần THƯỞNG THÊM (100% nữa để thành 200%)
            if row['IsSunday'] and role in ['Thợ thủ công', 'Thợ thủ công tập sự']:
                res['Sunday_Bonus'] = res['Daily_Pay'] # Cộng thêm đúng bằng 1 ngày lương
            else:
                res['Sunday_Bonus'] = 0
            
            results.append(res)

        self.df_final = pd.concat([self.df, pd.DataFrame(results)], axis=1)
        return self.df_final

    def calculate_monthly_salary(self):
        """Tổng hợp lương tháng."""
        if self.df_final is None: return None
        
        # Fill NaN để ko bị mất dữ liệu khi group
        df_safe = self.df_final.fillna({'Role': 'Unknown', 'Base Salary': 0, 'Salary Type': 'Unknown', 'Revenue': 0})
        
        grouped = df_safe.groupby(['Name', 'Role', 'Base Salary', 'Salary Type', 'Revenue'])
        summary = []

        # Lấy danh sách tất cả các ngày Thứ 7 có trong dữ liệu tháng này
        all_dates = pd.to_datetime(self.df_final['Date'].dropna().unique())
        all_saturdays = all_dates[all_dates.weekday == 5]

        for info, group in grouped:
            name, role, base, s_type, revenue = info
            
            total_penalty = group['Penalty_Amt'].sum()
            total_ot = group['OT_Amt'].sum()
            total_sunday_bonus = group['Sunday_Bonus'].sum()
            
            # Lương ngày thường
            base_earned = group['Daily_Pay'].sum()
            work_metric = group['Work_Day'].sum() if role not in ['Saleman', 'Intern'] else group['Work_Hours'].sum()

            # --- XỬ LÝ TRƯỜNG HỢP ĐẶC BIỆT: Tường Photo ---
            # Đặc cách nghỉ Thứ 7 vẫn tính lương (cộng bù công nếu vắng)
            if "Tường Photo" in str(name):
                worked_dates = pd.to_datetime(group['Date']).dt.date.unique()
                for sat in all_saturdays:
                    if sat.date() not in worked_dates:
                        # Cộng bù 1 ngày lương
                        day_val = (base / self.standard_days) if s_type == 'Tháng' else base
                        base_earned += day_val
                        if role not in ['Saleman', 'Intern']:
                            work_metric += 1
            
            # --- TÍNH HOA HỒNG (Chỉ áp dụng cho Saleman) ---
            commission = 0
            if role == 'Saleman' and revenue >= 80_000_000:
                if revenue <= 120_000_000:
                    commission = revenue * 0.03
                else:
                    commission = (120_000_000 * 0.03) + (revenue - 120_000_000) * 0.05
            
            total_income = base_earned + total_ot + total_sunday_bonus + commission - total_penalty
            
            summary.append({
                'Tên': name,
                'Chức vụ': role,
                'Lương Cơ Bản': base,
                'Doanh Thu': revenue,
                'Hoa Hồng': commission,
                'Ngày công/Giờ công': work_metric,
                'Lương Ngày Thường': base_earned,
                'Lương Chủ Nhật': total_sunday_bonus,
                'Tiền OT': total_ot,
                'Phạt (Trễ/Sớm)': total_penalty,
                'Tổng Thực Lãnh': total_income
            })
            
        return pd.DataFrame(summary)
