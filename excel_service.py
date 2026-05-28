import pandas as pd
import os
import datetime

class ExcelMatcher:
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.df = None
        self.load_data()

    def load_data(self):
        try:
            if self.excel_path and os.path.exists(self.excel_path):
                self.df = pd.read_excel(self.excel_path, dtype=str)
        except Exception as e:
            print(f"Lỗi khi đọc file Excel: {e}")
            self.df = None

    def match_passport(self, passport_num):
        """
        Searches for the passport number in the entire DataFrame.
        Returns True if found, False otherwise.
        """
        if self.df is None or self.df.empty or not passport_num:
            return False
            
        search_term = str(passport_num).strip().upper()
        mask = self.df.fillna('').astype(str).apply(lambda col: col.str.strip().str.upper() == search_term)
        return mask.any().any()

def format_date(date_str, output_format="DD/MM/YYYY"):
    import re
    date_str = str(date_str).strip()
    if not date_str:
        return ""
    
    # Try DD/MM/YYYY or DD-MM-YYYY
    m = re.match(r"(\d{2})[-/](\d{2})[-/](\d{4})", date_str)
    if m:
        d, m_m, y = m.groups()
        if output_format == "DD/MM/YYYY":
            return f"{d}/{m_m}/{y}"
        elif output_format == "YYYYMMDD":
            return f"{y}{m_m}{d}"
        else:
            return f"{y}/{m_m}/{d}"
            
    # Try YYYY/MM/DD or YYYY-MM-DD
    m = re.match(r"(\d{4})[-/](\d{2})[-/](\d{2})", date_str)
    if m:
        y, m_m, d = m.groups()
        if output_format == "DD/MM/YYYY":
            return f"{d}/{m_m}/{y}"
        elif output_format == "YYYYMMDD":
            return f"{y}{m_m}{d}"
        else:
            return f"{y}/{m_m}/{d}"
            
    # Try YYYYMMDD
    m = re.match(r"(\d{4})(\d{2})(\d{2})", date_str)
    if m:
        y, m_m, d = m.groups()
        if output_format == "DD/MM/YYYY":
            return f"{d}/{m_m}/{y}"
        elif output_format == "YYYYMMDD":
            return f"{y}{m_m}{d}"
        else:
            return f"{y}/{m_m}/{d}"

    return date_str

def save_results_to_excel(results_list, output_folder):
    """
    results_list: list of dicts with all passport fields.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = os.path.join(output_folder, f"Output_Results_{timestamp}.xlsx")
    
    output_data = []
    current_date = datetime.datetime.now().strftime("%Y/%m/%d")
    current_time = datetime.datetime.now().strftime("%H:%M:%S")
    
    for row in results_list:
        full_name = str(row.get("ho_ten", "")).strip()
        parts = full_name.split(" ", 1)
        if len(parts) == 2:
            last_name = parts[0]
            first_name = parts[1]
        else:
            last_name = full_name
            first_name = ""
            
        gender_raw = str(row.get("gioi_tinh", "")).strip().upper()
        if "NAM" in gender_raw or gender_raw == "M":
            gender = "M"
        elif "NỮ" in gender_raw or "NU" in gender_raw or gender_raw == "F":
            gender = "F"
        else:
            gender = gender_raw

        out_row = [
            row.get("stt", ""),
            current_date,
            current_time,
            "", # Time Out
            first_name.upper(),
            last_name.upper(),
            format_date(row.get("ngay_sinh", ""), "DD/MM/YYYY"),
            gender,
            row.get("so_passport", ""),
            row.get("quoc_tich", ""),
            format_date(row.get("ngay_het_han", ""), "DD/MM/YYYY"),
            format_date(row.get("ngay_cap", ""), "DD/MM/YYYY")
        ]
        output_data.append(out_row)
    
    cols = ["Index", "Date", "Time In", "Time Out", "First Name", "Last Name", "Date of Birth", "Gender", "Document Number", "Nationality", "Valid Until", "Date of Issue"]
    
    wb = Workbook()
    
    # --- STYLE CHUNG ---
    font_family = "Arial"
    font_regular = Font(name=font_family, size=10)
    font_header = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    
    fill_header = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid") # Xanh Navy đậm
    fill_zebra = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid") # Xám nhạt
    
    border_thin = Side(style='thin', color='D9D9D9')
    box_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

    # --- SHEET 1: ATTENDANT REPORT ---
    ws1 = wb.active
    ws1.title = "ATTENDANT REPORT"
    ws1.views.sheetView[0].showGridLines = True
    
    # Ghi headers
    ws1.append(cols)
    for col_idx in range(1, len(cols) + 1):
        cell = ws1.cell(row=1, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = box_border
    
    # Ghi dữ liệu
    for row_idx, row in enumerate(output_data, start=2):
        ws1.append(list(row))
        for col_idx in range(1, len(cols) + 1):
            cell = ws1.cell(row=row_idx, column=col_idx)
            cell.font = font_regular
            cell.border = box_border
            
            # Tô màu zebra
            if row_idx % 2 == 0:
                cell.fill = fill_zebra
                
            # Định dạng căn lề
            col_name = cols[col_idx - 1]
            if col_name in ['Index', 'Date', 'Time In', 'Time Out', 'Gender', 'Nationality', 'Valid Until', 'Date of Issue', 'Document Number', 'Date of Birth']:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Tự động căn chỉnh độ rộng cột
    for col in ws1.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    try:
        wb.save(output_filename)
        return output_filename, True
    except Exception as e:
        print(f"Error saving output Excel: {e}")
        return str(e), False
