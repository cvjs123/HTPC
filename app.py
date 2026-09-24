try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
    import ttkbootstrap as ttk
    from ttkbootstrap.constants import *
    GUI_AVAILABLE = True
except ImportError:
    tk = None
    filedialog = None
    messagebox = None
    ttk = None
    GUI_AVAILABLE = False
    VERTICAL = None
    PRIMARY = "primary"
    SUCCESS = "success"
    WARNING = "warning"
    DANGER = "danger"
    INFO = "info"
    SECONDARY = "secondary"

import openpyxl
import re
import os
import aggregate

def parse_phieu(filepath):
    wb = None
    try:
        # Some Excel files with merged cells / complex styles can crash the default
        # openpyxl loader in production. Try the safer read-only mode first to avoid
        # taking down the worker process on a single bad file.
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        except (ValueError, TypeError, OSError, RuntimeError, SystemExit):
            # Fallback to the standard loader for compatibility with files that need
            # the full workbook object, but keep the failure non-fatal.
            wb = openpyxl.load_workbook(filepath, data_only=True)

        visible_sheets = [ws for ws in wb.worksheets if ws.sheet_state == "visible"]
        if not visible_sheets:
            raise ValueError("File Excel không có trang tính đang hiển thị")
        sheet = wb.active if wb.active in visible_sheets else visible_sheets[0]
        if sheet.max_row < 5 or sheet.max_column < 2:
            raise ValueError("Trang tính không có đủ dữ liệu để đọc phiếu")
        
        # We can extract general info by searching through the first few rows
        # Số phiếu, Địa chỉ, Chủ hộ, Điện thoại, Diện cư trú
        general_info = {
            'so_phieu': '',
            'dia_chi': '',
            'chu_ho': '',
            'dien_thoai': '',
            'dien_cu_tru': ''
        }
        
        # Read the first 10 rows to find general info
        for r in range(1, 10):
            for c in range(1, 40):
                cell_value = sheet.cell(row=r, column=c).value
                if not cell_value or not isinstance(cell_value, str):
                    continue
                val = str(cell_value)
                
                if 'Số phiếu' in val:
                    match = re.search(r'Số phiếu.*:\s*(.*)', val)
                    if match: general_info['so_phieu'] = match.group(1).strip()
                if 'Địa chỉ' in val:
                    match = re.search(r'Địa chỉ.*:\s*(.*)', val)
                    if match: general_info['dia_chi'] = match.group(1).strip()
                if 'chủ hộ' in val.lower() and 'Họ và tên' in val:
                    match = re.search(r'chủ hộ.*:\s*(.*)', val)
                    if match: general_info['chu_ho'] = match.group(1).strip()
                if 'Điện thoại' in val:
                    match = re.search(r'Điện thoại.*:\s*([0-9\-\.\s]*)', val)
                    if match: 
                        phone = match.group(1).replace('_', '').strip()
                        general_info['dien_thoai'] = phone
                if 'Diện cư trú' in val:
                    general_info['dien_cu_tru'] = val
                    
        # Now find where the table starts
        start_row = -1
        for r in range(5, 20):
            cell = sheet.cell(row=r, column=1).value
            if str(cell).strip() == '1': # Finding TT = 1
                start_row = r
                break
                
        if start_row == -1:
            raise ValueError("Không tìm thấy bảng dữ liệu (cột TT = 1)")
            
        people = []
        current_row = start_row

        def is_checkbox_value(value):
            if value is None:
                return False
            text = str(value).strip().lower()
            if not text:
                return False
            if text in {'x', 'q', 'v', '✓', '☒', '☑', 'þ', 'ý'}:
                return True
            return 'x' in text or 'q' in text or 'v' in text

        def read_disability_flags(row_index):
            flags = {
                32: '',
                33: '',
                34: '',
                35: '',
                36: '',
                37: '',
                38: '',
                39: '',
                40: '',
                41: '',
            }

            try:
                disability_type = str(sheet.cell(row=row_index, column=29).value or '').strip().lower()
                type_columns = {
                    'vận động': 32,
                    'nghe nói': 33,
                    'nhìn': 34,
                    'thần kinh': 35,
                    'tâm thần': 35,
                    'trí tuệ': 36,
                    'học tập': 37,
                    'tự kỷ': 38,
                    'khác': 39,
                }
                for label, output_column in type_columns.items():
                    if label in disability_type:
                        flags[output_column] = 'x'
                        break

                if disability_type and not any(label in disability_type for label in type_columns):
                    flags[39] = disability_type

                if is_checkbox_value(sheet.cell(row=row_index, column=30).value):
                    flags[40] = 'x'
                if is_checkbox_value(sheet.cell(row=row_index, column=31).value):
                    flags[41] = 'x'
            except Exception:
                pass

            return flags
        
        while True:
            tt = sheet.cell(row=current_row, column=1).value
            if not tt or str(tt).strip() == '' or str(tt).strip().startswith('Họ, tên'):
                break # End of table
                
            ho_ten = sheet.cell(row=current_row, column=2).value or ""
            ho_ten = str(ho_ten).strip()
            if not ho_ten:
                break # Empty row, meaning no more people
                
            lop_hoc = sheet.cell(row=current_row, column=8).value or ""
            ma_truong = sheet.cell(row=current_row, column=11).value or ""
            bac_tn = sheet.cell(row=current_row, column=16).value or ""
            bo_tuc = sheet.cell(row=current_row, column=17).value or ""
            nam_tn = sheet.cell(row=current_row, column=18).value or ""
            bac_tn_nghe = sheet.cell(row=current_row, column=19).value or ""
            nam_tn_nghe = sheet.cell(row=current_row, column=20).value or ""
            # Học xong / Bỏ học (cột 22-25 của phiếu)
            hoc_xong_lop = sheet.cell(row=current_row, column=22).value or ""
            hoc_xong_nam = sheet.cell(row=current_row, column=23).value or ""
            bo_hoc_lop = sheet.cell(row=current_row, column=24).value or ""
            bo_hoc_nam = sheet.cell(row=current_row, column=25).value or ""
            
            qh_chu_ho = sheet.cell(row=current_row + 1, column=2).value or ""
            if "QH với chủ hộ:" in str(qh_chu_ho):
                qh_chu_ho = str(qh_chu_ho).split(":", 1)[-1].replace('_', '').strip()
                
            ngay_sinh_str = sheet.cell(row=current_row + 2, column=2).value or ""
            ngay, thang, nam = "", "", ""
            if "Ngày sinh:" in str(ngay_sinh_str):
                ns = str(ngay_sinh_str).split(":", 1)[-1].strip()
                parts = ns.split('/')
                if len(parts) == 3:
                    ngay, thang, nam = parts[0], parts[1], parts[2]
                elif len(parts) == 1 and len(ns) == 4:
                    nam = ns # Only year
                    
            # Tìm dấu check nữ trong các cột 2, 3, 4
            nu_str = ""
            for c in range(2, 5):
                val = str(sheet.cell(row=current_row + 3, column=c).value or "")
                if any(x in val.lower() for x in ['x', 'q', '☒', '☑', '\u2611', '\u2612', 'þ', 'ý', 'v']):
                    nu_str = val
                    break
            
            dt_str = str(sheet.cell(row=current_row + 3, column=4).value or "")
            nu_str_lower = nu_str.lower()
            has_x = any(c in nu_str_lower for c in ['x', 'q', '☒', '☑', '\u2611', '\u2612', 'þ', 'ý', 'v'])
            is_nu = "x" if has_x else ""
            
            dan_toc = ""
            if "DT:" in str(dt_str):
                dan_toc = str(dt_str).split(":", 1)[-1].replace('_', '').strip()
            for c in range(3, 8):
                val = str(sheet.cell(row=current_row + 3, column=c).value or "")
                if "DT:" in val:
                    dan_toc = val.split(":", 1)[-1].replace('_', '').strip()
                    break
                    
            cha_me_str = sheet.cell(row=current_row + 4, column=2).value or ""
            cha_me = ""
            if "Cha, mẹ" in str(cha_me_str):
                cha_me = str(cha_me_str).split(":", 1)[-1].replace('_', '').strip()
                
            ho_dem, ten = "", ""
            ho_ten = str(ho_ten).strip()
            if ho_ten:
                parts = ho_ten.split(' ')
                ten = parts[-1]
                ho_dem = " ".join(parts[:-1])
                
            if not cha_me and "con" in str(qh_chu_ho).lower():
                cha_me = general_info.get('chu_ho', '')

            disability_flags = read_disability_flags(current_row)
                
            people.append({
                'tt': tt,
                'ho_dem': ho_dem,
                'ten': ten,
                'qh_chu_ho': qh_chu_ho,
                'ngay': ngay,
                'thang': thang,
                'nam': nam,
                'is_nu': is_nu,
                'dan_toc': dan_toc,
                'cha_me': cha_me,
                'lop_hoc': lop_hoc,
                'ma_truong': ma_truong,
                'bac_tn': bac_tn,
                'bo_tuc': bo_tuc,
                'nam_tn': nam_tn,
                'bac_tn_nghe': bac_tn_nghe,
                'nam_tn_nghe': nam_tn_nghe,
                'hoc_xong_lop': hoc_xong_lop,
                'hoc_xong_nam': hoc_xong_nam,
                'bo_hoc_lop': bo_hoc_lop,
                'bo_hoc_nam': bo_hoc_nam,
                'disability_flags': disability_flags
            })
            
            current_row += 5
            
        return general_info, people, ""
    except (SystemExit, KeyboardInterrupt):
        error_message = (
            f"Không đọc được phiếu '{os.path.basename(filepath)}': file Excel có vùng hợp ô hoặc định dạng không hỗ trợ trên máy chủ."
        )
        return None, None, error_message
    except Exception as e:
        error_message = f"Không đọc được phiếu '{os.path.basename(filepath)}': {e}"
        return None, None, error_message
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass

def write_to_tong(general_info, people, tong_filepath, workbook=None):
    wb = workbook or openpyxl.load_workbook(tong_filepath)
    sheet = wb.active
    
    last_row = sheet.max_row
    start_write_row = 6
    
    for r in range(5, last_row + 2):
        if not sheet.cell(row=r, column=1).value:
            start_write_row = r
            break
            
    chu_ho = str(general_info.get('chu_ho') or '').strip()
    chu_ho_parts = chu_ho.split()
    chu_ho_ten = chu_ho_parts[-1] if chu_ho_parts else ""
    chu_ho_dem = " ".join(chu_ho_parts[:-1]) if len(chu_ho_parts) > 1 else ""
            
    for person in people:
        sheet.cell(row=start_write_row, column=1).value = person['tt']
        sheet.cell(row=start_write_row, column=2).value = person['ho_dem']
        sheet.cell(row=start_write_row, column=3).value = person['ten']
        sheet.cell(row=start_write_row, column=4).value = person['ngay']
        sheet.cell(row=start_write_row, column=5).value = person['thang']
        sheet.cell(row=start_write_row, column=6).value = person['nam']
        # Giới tính: ưu tiên dấu x từ phiếu, nếu không có thì suy luận từ họ đệm
        is_nu_val = person['is_nu']
        if not is_nu_val:
            ho_dem_lower = person['ho_dem'].lower()
            # Nữ nếu họ đệm chứa 'thị' hoặc tên chứa từ phổ biến của nữ dân tộc thiểu số
            if 'thị' in ho_dem_lower or ' thi ' in ho_dem_lower:
                is_nu_val = 'x'
        sheet.cell(row=start_write_row, column=7).value = is_nu_val
        sheet.cell(row=start_write_row, column=8).value = person['dan_toc']
        
        sheet.cell(row=start_write_row, column=11).value = chu_ho_dem
        sheet.cell(row=start_write_row, column=12).value = chu_ho_ten
        sheet.cell(row=start_write_row, column=13).value = general_info['dia_chi']
        sheet.cell(row=start_write_row, column=14).value = general_info['so_phieu']
        
        # Cột 17: Khối học - chỉ lấy phần số đầu (VD: 6A → 6, 10B → 10, MN3 giữ nguyên)
        lop_str = str(person['lop_hoc'])
        import re as _re
        m = _re.match(r'^(\d+)', lop_str)
        khoi_hoc = m.group(1) if m else lop_str
        sheet.cell(row=start_write_row, column=17).value = khoi_hoc
        # Cột 18: Lớp học - giữ nguyên giá trị gốc (VD: 6A, 10B, MN3...)
        sheet.cell(row=start_write_row, column=18).value = person['lop_hoc']
        sheet.cell(row=start_write_row, column=19).value = person['ma_truong']
        sheet.cell(row=start_write_row, column=20).value = person['bac_tn']
        sheet.cell(row=start_write_row, column=21).value = person['bo_tuc']
        sheet.cell(row=start_write_row, column=22).value = person['nam_tn']
        sheet.cell(row=start_write_row, column=23).value = person['bac_tn_nghe']
        sheet.cell(row=start_write_row, column=24).value = person['nam_tn_nghe']
        # Học xong: cột 25 (Lớp), 26 (Năm)
        sheet.cell(row=start_write_row, column=25).value = person['hoc_xong_lop']
        sheet.cell(row=start_write_row, column=26).value = person['hoc_xong_nam']
        # Bỏ học: cột 27 (Lớp), 28 (Năm)
        sheet.cell(row=start_write_row, column=27).value = person['bo_hoc_lop']
        sheet.cell(row=start_write_row, column=28).value = person['bo_hoc_nam']

        disability_flags = person.get('disability_flags', {})
        for col, value in disability_flags.items():
            if value:
                sheet.cell(row=start_write_row, column=col).value = value
        
        sheet.cell(row=start_write_row, column=44).value = person['qh_chu_ho']
        sheet.cell(row=start_write_row, column=45).value = person['cha_me']
        sheet.cell(row=start_write_row, column=46).value = general_info['dien_thoai']
        
        start_write_row += 1
        
    if workbook is None:
        wb.save(tong_filepath)

class App:
    def __init__(self, root):
        if not GUI_AVAILABLE:
            raise RuntimeError("Desktop GUI is not available in this environment. Use the web app instead.")
        self.root = root
        self.root.title("Tổng Hợp Phiếu Điều Tra")
        self.root.geometry("700x750")
        self.root.minsize(650, 600)
        
        self.phieu_files = []
        self.tong_file = ""
        self.m1_file = ""
        self.m2_file = ""
        self.spc_file = ""
        self.spc1_file = ""
        self.spcth_file = ""
        self.m1th_file = ""
        self.xmth_file = ""
        self.pl_file = os.path.join(os.getcwd(), "PL.xlsx") if os.path.exists(os.path.join(os.getcwd(), "PL.xlsx")) else ""
        self.kt_file = os.path.join(os.getcwd(), "KT.xlsx") if os.path.exists(os.path.join(os.getcwd(), "KT.xlsx")) else ""
        self.error_records = []
        self.review_records = []
        self.last_tong_result = ""
        
        self.create_widgets()
        
    def create_widgets(self):
        # Scrollable canvas
        canvas = tk.Canvas(self.root, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient=VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        
        frame = ttk.Frame(canvas, padding=30)
        frame_id = canvas.create_window((0, 0), window=frame, anchor=NW)
        
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox('all'))
        frame.bind('<Configure>', on_frame_configure)
        
        def on_canvas_configure(event):
            canvas.itemconfig(frame_id, width=event.width)
        canvas.bind('<Configure>', on_canvas_configure)
        
        # Mouse scroll: gom các tín hiệu nhỏ của touchpad rồi cuộn theo nhịp ổn định.
        wheel_state = {'delta': 0, 'job': None}

        def flush_mousewheel():
            wheel_state['job'] = None
            delta = wheel_state['delta']
            units = int(delta / 120)
            if units:
                wheel_state['delta'] -= units * 120
                canvas.yview_scroll(-units, 'units')

        def on_mousewheel(event):
            wheel_state['delta'] += event.delta
            if wheel_state['job'] is None:
                wheel_state['job'] = self.root.after(8, flush_mousewheel)

        canvas.bind_all('<MouseWheel>', on_mousewheel)
        
        # Title
        title_lbl = ttk.Label(frame, text="PHẦN MỀM PHỔ CẬP", font=("Segoe UI", 18, "bold"), bootstyle=PRIMARY)
        title_lbl.pack(pady=(0, 18))

        # Two-column layout for clearer grouping and easier usage
        main_grid = ttk.Frame(frame)
        main_grid.pack(fill=BOTH, expand=True)
        left_col = ttk.Frame(main_grid)
        right_col = ttk.Frame(main_grid)
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        right_col.grid(row=0, column=1, sticky="nsew")
        main_grid.columnconfigure(0, weight=3)
        main_grid.columnconfigure(1, weight=2)

        # Section 1: Phieu Files
        lf1 = ttk.Labelframe(left_col, text=" Bước 1: Chọn các file Phiếu Đầu Vào ", padding=15, bootstyle=INFO)
        lf1.pack(fill=X, pady=(0, 15))

        self.btn_select_phieu = ttk.Button(lf1, text="📂 Chọn Files Phiếu", command=self.select_phieu, bootstyle=(SUCCESS, "solid"), cursor="hand2")
        self.btn_select_phieu.pack(side=LEFT, padx=(0, 10), pady=4, ipadx=12, ipady=5)

        self.btn_select_phieu_dir = ttk.Button(lf1, text="📁 Chọn Thư Mục Phiếu", command=self.select_phieu_dir, bootstyle=(SUCCESS, "solid"), cursor="hand2")
        self.btn_select_phieu_dir.pack(side=LEFT, padx=(0, 10), pady=4, ipadx=10, ipady=5)

        self.lbl_phieu_count = ttk.Label(lf1, text="Chưa chọn file nào", font=("Segoe UI", 10))
        self.lbl_phieu_count.pack(side=LEFT, pady=4)

        # Section 2: Tong File
        lf2 = ttk.Labelframe(left_col, text=" Bước 2: Chọn file Tổng Hợp (Tong.xlsx) ", padding=15, bootstyle=INFO)
        lf2.pack(fill=X, pady=(0, 15))

        self.btn_select_tong = ttk.Button(lf2, text="📄 Chọn File Tổng", command=self.select_tong, bootstyle=(WARNING, "solid"), cursor="hand2")
        self.btn_select_tong.pack(side=LEFT, padx=(0, 10), pady=4, ipadx=10, ipady=5)

        self.lbl_tong_file = ttk.Label(lf2, text="Chưa chọn file", font=("Segoe UI", 10))
        self.lbl_tong_file.pack(side=LEFT, pady=4)

        export_tong_row = ttk.Frame(lf2)
        export_tong_row.pack(fill=X, pady=(10, 0))
        self.btn_export_tong = ttk.Button(
            export_tong_row,
            text="💾 Xuất file Tổng...",
            command=self.export_tong_file,
            bootstyle=(WARNING, "outline"),
            cursor="hand2"
        )
        self.btn_export_tong.pack(side=LEFT, padx=(0, 12), ipadx=10, ipady=4)
        self.lbl_export_tong = ttk.Label(export_tong_row, text="Chưa có file tổng hợp", font=("Segoe UI", 9))
        self.lbl_export_tong.pack(side=LEFT)

        # Progress and Action
        action_frame = ttk.Frame(left_col, padding=12)
        action_frame.pack(fill=X, pady=10)

        self.btn_run = ttk.Button(action_frame, text="🚀 BẮT ĐẦU TỔNG HỢP", command=self.run_extraction, bootstyle=(PRIMARY, "solid"), cursor="hand2")
        self.btn_run.pack(fill=X, pady=4, ipady=8)

        self.btn_export_errors = ttk.Button(
            action_frame,
            text="📋 Xuất danh sách phiếu lỗi...",
            command=self.export_error_report,
            bootstyle=(DANGER, "outline"),
            cursor="hand2"
        )
        self.btn_export_errors.pack(fill=X, pady=(6, 0), ipady=5)

        self.progress_bar = ttk.Progressbar(left_col, mode='determinate', bootstyle=SUCCESS)
        self.progress_bar.pack(fill=X, pady=(10, 5))

        self.progress_var = tk.StringVar()
        self.progress_var.set("Sẵn sàng.")
        self.lbl_progress = ttk.Label(left_col, textvariable=self.progress_var, font=("Segoe UI", 9))
        self.lbl_progress.pack(anchor=W)

        # Section 3: Tong hop M1, M2
        lf3 = ttk.Labelframe(left_col, text=" Bước 3: Tổng hợp báo cáo M1, M2 ", padding=15, bootstyle=SECONDARY)
        lf3.pack(fill=X, pady=(12, 10))

        m1_row = ttk.Frame(lf3)
        m1_row.pack(fill=X, pady=(0, 8))
        self.btn_select_m1 = ttk.Button(m1_row, text="📁 Chọn file mẫu M1.xlsx", command=self.select_m1, bootstyle=(SUCCESS, "outline"), cursor="hand2")
        self.btn_select_m1.pack(side=LEFT, padx=(0, 12), ipadx=8, ipady=4)
        self.lbl_m1_file = ttk.Label(m1_row, text="Chưa chọn", font=("Segoe UI", 10))
        self.lbl_m1_file.pack(side=LEFT)

        m2_row_frame = ttk.Frame(lf3)
        m2_row_frame.pack(fill=X, pady=(0, 12))
        self.btn_select_m2 = ttk.Button(m2_row_frame, text="📁 Chọn file mẫu M2.xlsx", command=self.select_m2, bootstyle=(SUCCESS, "outline"), cursor="hand2")
        self.btn_select_m2.pack(side=LEFT, padx=(0, 12), ipadx=8, ipady=4)
        self.lbl_m2_file = ttk.Label(m2_row_frame, text="Chưa chọn", font=("Segoe UI", 10))
        self.lbl_m2_file.pack(side=LEFT)

        kt_row = ttk.Frame(lf3)
        kt_row.pack(fill=X, pady=(0, 12))
        self.btn_select_kt = ttk.Button(kt_row, text="📁 Chọn mẫu KT.xlsx", command=self.select_kt, bootstyle=(INFO, "outline"), cursor="hand2")
        self.btn_select_kt.pack(side=LEFT, padx=(0, 10), ipadx=8, ipady=4)
        self.btn_export_kt = ttk.Button(kt_row, text="📝 Xuất file Khuyết tật", command=self.export_disability, bootstyle=(INFO, "solid"), cursor="hand2")
        self.btn_export_kt.pack(side=LEFT, padx=(0, 10), ipadx=8, ipady=4)
        self.lbl_kt_file = ttk.Label(kt_row, text=os.path.basename(self.kt_file) if self.kt_file else "Chưa chọn", font=("Segoe UI", 9))
        self.lbl_kt_file.pack(side=LEFT)

        self.btn_run_m1_m2 = ttk.Button(lf3, text="📊 Tổng hợp M1 & M2", command=self.run_aggregate, bootstyle=(DANGER, "solid"), cursor="hand2")
        self.btn_run_m1_m2.pack(fill=X, ipady=5)

        # Section 4: SPC1
        lf4 = ttk.Labelframe(left_col, text=" Bước 4: Tổng hợp danh sách SPC1 ", padding=15, bootstyle=SECONDARY)
        lf4.pack(fill=X, pady=(0, 10))

        spc1_row = ttk.Frame(lf4)
        spc1_row.pack(fill=X, pady=(0, 12))
        self.btn_select_spc1 = ttk.Button(spc1_row, text="📁 Chọn file mẫu SPC1.xlsx", command=self.select_spc1, bootstyle=(WARNING, "outline"), cursor="hand2")
        self.btn_select_spc1.pack(side=LEFT, padx=(0, 12), ipadx=8, ipady=4)
        self.lbl_spc1_file = ttk.Label(spc1_row, text="Chưa chọn", font=("Segoe UI", 10))
        self.lbl_spc1_file.pack(side=LEFT)

        self.btn_run_spc1 = ttk.Button(lf4, text="📝 Tổng hợp SPC1", command=self.run_aggregate_spc1, bootstyle=(WARNING, "solid"), cursor="hand2")
        self.btn_run_spc1.pack(fill=X, ipady=5)

        # Section 5: Review missing data
        lf5 = ttk.Labelframe(right_col, text=" Bước 5: Rà soát thông tin còn thiếu ", padding=15, bootstyle=SECONDARY)
        lf5.pack(fill=X, pady=(0, 10))

        self.btn_review_tong = ttk.Button(
            lf5,
            text="🔎 Rà soát thông tin thiếu",
            command=self.review_tong_data,
            bootstyle=(INFO, "outline"),
            cursor="hand2"
        )
        self.btn_review_tong.pack(side=LEFT, padx=(0, 12), ipadx=8, ipady=3)

        self.btn_export_review = ttk.Button(
            lf5,
            text="📤 Xuất danh sách cần bổ sung...",
            command=self.export_review_report,
            bootstyle=(WARNING, "outline"),
            cursor="hand2"
        )
        self.btn_export_review.pack(side=LEFT, ipadx=8, ipady=3)

        self.lbl_review_status = ttk.Label(lf5, text="Chưa rà soát", font=("Segoe UI", 9))
        self.lbl_review_status.pack(side=LEFT, padx=(12, 0))

        # Section 6: Primary education reports
        lf6 = ttk.Labelframe(right_col, text=" Bước 6: Tổng hợp 3 biểu phổ cập Tiểu học ", padding=15, bootstyle=SUCCESS)
        lf6.pack(fill=X, pady=(0, 10))

        primary_files = [
            ("SPCTH", "Chọn file mẫu SPCTH.xlsx", "spcth_file", "lbl_spcth_file", self.select_spcth),
            ("M1TH", "Chọn file mẫu M1TH.xlsx", "m1th_file", "lbl_m1th_file", self.select_m1th),
            ("XMTH", "Chọn file mẫu XMTH.xlsx", "xmth_file", "lbl_xmth_file", self.select_xmth),
        ]
        for label, button_text, file_attr, label_attr, command in primary_files:
            row = ttk.Frame(lf6)
            row.pack(fill=X, pady=(0, 8))
            ttk.Button(
                row, text=button_text, command=command,
                bootstyle=(SUCCESS, "outline"), cursor="hand2"
            ).pack(side=LEFT, padx=(0, 12), ipadx=8, ipady=4)
            setattr(self, label_attr, ttk.Label(row, text="Chưa chọn", font=("Segoe UI", 10)))
            getattr(self, label_attr).pack(side=LEFT)

        self.btn_run_primary = ttk.Button(
            lf6, text="📊 Tổng hợp 3 biểu Tiểu học",
            command=self.run_aggregate_primary, bootstyle=(SUCCESS, "solid"), cursor="hand2"
        )
        self.btn_run_primary.pack(fill=X, ipady=5, pady=(4, 0))

        # Section 7: Appendix export
        lf7 = ttk.Labelframe(right_col, text=" Bước 7: Xuất phụ lục theo mẫu PL.xlsx ", padding=15, bootstyle=INFO)
        lf7.pack(fill=X, pady=(0, 10))

        pl_row = ttk.Frame(lf7)
        pl_row.pack(fill=X, pady=(0, 8))
        self.btn_select_pl = ttk.Button(
            pl_row, text="📁 Chọn file mẫu PL.xlsx",
            command=self.select_pl, bootstyle=(INFO, "outline"), cursor="hand2"
        )
        self.btn_select_pl.pack(side=LEFT, padx=(0, 12), ipadx=8, ipady=4)
        self.lbl_pl_file = ttk.Label(pl_row, text=os.path.basename(self.pl_file) if self.pl_file else "Chưa chọn", font=("Segoe UI", 10))
        self.lbl_pl_file.pack(side=LEFT)

        self.btn_export_appendix = ttk.Button(
            lf7, text="📋 Xuất phụ lục",
            command=self.export_appendix, bootstyle=(INFO, "solid"), cursor="hand2"
        )
        self.btn_export_appendix.pack(fill=X, ipady=5)

    def select_phieu(self):
        files = filedialog.askopenfilenames(
            title="Chọn các file Phiếu điều tra",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if files:
            self.phieu_files = files
            self.lbl_phieu_count.config(text=f"Đã chọn: {len(self.phieu_files)} files")

    def select_phieu_dir(self):
        directory = filedialog.askdirectory(title="Chọn thư mục chứa các file Phiếu")
        if directory:
            files = [os.path.join(directory, f) for f in os.listdir(directory) if f.lower().endswith('.xlsx') and not f.startswith('~$')]
            if files:
                self.phieu_files = files
                self.lbl_phieu_count.config(text=f"Đã chọn: {len(self.phieu_files)} files")
            else:
                messagebox.showwarning("Cảnh báo", "Không tìm thấy file .xlsx nào trong thư mục này.")

    def select_tong(self):
        file = filedialog.askopenfilename(
            title="Chọn file Tong.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.tong_file = file
            self.lbl_tong_file.config(text=os.path.basename(file))

    def select_m1(self):
        file = filedialog.askopenfilename(
            title="Chọn file mẫu M1.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.m1_file = file
            self.lbl_m1_file.config(text=os.path.basename(file))

    def select_m2(self):
        file = filedialog.askopenfilename(
            title="Chọn file mẫu M2.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.m2_file = file
            self.lbl_m2_file.config(text=os.path.basename(file))

    def select_spc1(self):
        file = filedialog.askopenfilename(
            title="Chọn file mẫu SPC1.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.spc1_file = file
            self.lbl_spc1_file.config(text=os.path.basename(file))

    def select_spcth(self):
        file = filedialog.askopenfilename(
            title="Chọn file mẫu SPCTH.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.spcth_file = file
            self.lbl_spcth_file.config(text=os.path.basename(file))

    def select_m1th(self):
        file = filedialog.askopenfilename(
            title="Chọn file mẫu M1TH.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.m1th_file = file
            self.lbl_m1th_file.config(text=os.path.basename(file))

    def select_xmth(self):
        file = filedialog.askopenfilename(
            title="Chọn file mẫu XMTH.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.xmth_file = file
            self.lbl_xmth_file.config(text=os.path.basename(file))

    def select_pl(self):
        file = filedialog.askopenfilename(
            title="Chọn file mẫu PL.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.pl_file = file
            self.lbl_pl_file.config(text=os.path.basename(file))

    def select_kt(self):
        file = filedialog.askopenfilename(
            title="Chọn file mẫu KT.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.kt_file = file
            self.lbl_kt_file.config(text=os.path.basename(file))

    def get_source_tong_file(self):
        tong_dir = os.path.dirname(self.tong_file)
        ketqua_tong_file = os.path.join(tong_dir, "KetQua_Tong.xlsx")
        if os.path.exists(ketqua_tong_file):
            return ketqua_tong_file
        return self.tong_file

    def export_tong_file(self):
        source_file = self.last_tong_result or self.get_source_tong_file()
        if not source_file or not os.path.exists(source_file):
            messagebox.showwarning("Chưa có file", "Vui lòng tổng hợp phiếu trước khi xuất file Tổng.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Chọn nơi lưu file Tổng",
            defaultextension=".xlsx",
            initialfile="KetQua_Tong.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not save_path:
            return

        try:
            import shutil
            shutil.copy2(source_file, save_path)
            self.lbl_export_tong.config(text=os.path.basename(save_path))
            messagebox.showinfo("Xuất file thành công", f"File Tổng đã được lưu tại:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Lỗi xuất file", f"Không thể xuất file Tổng:\n{e}")

    def export_error_report(self):
        if not self.error_records:
            messagebox.showinfo("Danh sách phiếu lỗi", "Không có phiếu lỗi trong lần tổng hợp gần nhất.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Lưu danh sách phiếu lỗi",
            defaultextension=".xlsx",
            initialfile="DanhSachPhieuLoi.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not save_path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Phiếu lỗi"
            ws.append(["STT", "Tên file phiếu", "Đường dẫn", "Nguyên nhân"])
            for index, record in enumerate(self.error_records, start=1):
                ws.append([index, record["file_name"], record["file_path"], record["reason"]])
            ws.freeze_panes = "A2"
            ws.column_dimensions["A"].width = 8
            ws.column_dimensions["B"].width = 35
            ws.column_dimensions["C"].width = 70
            ws.column_dimensions["D"].width = 80
            wb.save(save_path)
            messagebox.showinfo("Xuất danh sách thành công", f"Danh sách phiếu lỗi đã được lưu tại:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Lỗi xuất danh sách", f"Không thể xuất danh sách phiếu lỗi:\n{e}")

    @staticmethod
    def _is_missing(value):
        if value is None:
            return True
        return str(value).strip().lower() in {"", "nan", "none", "null"}

    @staticmethod
    def _is_academic_year(value):
        if App._is_missing(value):
            return False
        match = re.fullmatch(r"\s*(\d{4})\s*-\s*(\d{4})\s*", str(value))
        if not match:
            return False
        return int(match.group(2)) == int(match.group(1)) + 1

    @staticmethod
    def _graduation_level(value):
        if App._is_missing(value):
            return ""
        normalized = re.sub(r"[\s()_\-]+", "", str(value).strip().upper())
        if "THPT" in normalized or "TRUNGHỌCPHỔTHÔNG" in normalized:
            return "THPT"
        if "THCS" in normalized or "TRUNGHỌCCƠSỞ" in normalized:
            return "THCS"
        if normalized == "TH" or "TIỂUHỌC" in normalized:
            return "TH"
        return ""

    @staticmethod
    def _class_grade(class_value, grade_value=None):
        value = class_value if not App._is_missing(class_value) else grade_value
        if App._is_missing(value):
            return None
        match = re.match(r"\s*(\d+)", str(value))
        return int(match.group(1)) if match else None

    @staticmethod
    def _graduation_class_mismatch(class_value, grade_value, graduation_value):
        graduation_level = App._graduation_level(graduation_value)
        class_grade = App._class_grade(class_value, grade_value)
        if not graduation_level or class_grade is None:
            return None

        valid_graduation = {
            "TH": range(6, 13),
            "THCS": range(10, 13),
            "THPT": range(13, 13),
        }
        if class_grade not in valid_graduation[graduation_level]:
            return graduation_level, class_grade
        return None

    def review_tong_data(self):
        # Rà soát file kết quả mới nhất để kiểm tra đúng dữ liệu vừa tổng hợp.
        source_file = self.last_tong_result or self.get_source_tong_file()
        if not source_file or not os.path.exists(source_file):
            messagebox.showwarning("Chưa có file Tổng", "Vui lòng chọn hoặc tạo file Tổng trước khi rà soát.")
            return

        try:
            wb = openpyxl.load_workbook(source_file, data_only=True)
            visible_sheets = [ws for ws in wb.worksheets if ws.sheet_state == "visible"]
            if not visible_sheets:
                raise ValueError("File Tổng không có trang tính đang hiển thị")
            sheet = wb.active if wb.active in visible_sheets else visible_sheets[0]

            self.review_records = []
            for row_number in range(6, sheet.max_row + 1):
                row_values = [sheet.cell(row=row_number, column=column).value for column in range(1, 48)]
                if all(self._is_missing(value) for value in row_values):
                    continue

                year_value = row_values[5]
                try:
                    birth_year = int(float(year_value))
                except (TypeError, ValueError):
                    birth_year = None

                person_name = " ".join(
                    str(row_values[index]).strip()
                    for index in (1, 2)
                    if not self._is_missing(row_values[index])
                )
                village = row_values[12]
                so_phieu = row_values[13]

                missing_fields = []
                if birth_year is None:
                    missing_fields.append(("M1, M2", "Năm sinh"))
                if self._is_missing(village):
                    missing_fields.append(("M2", "Địa chỉ/thôn bản"))
                if self._is_missing(row_values[7]):
                    missing_fields.append(("M1", "Dân tộc"))

                is_m2_age_group = bool(birth_year and 2008 <= birth_year <= 2014)
                dang_hoc = (
                    not self._is_missing(row_values[17])
                    and self._is_missing(row_values[24])
                    and self._is_missing(row_values[26])
                )
                if dang_hoc and self._is_missing(row_values[16]):
                    missing_fields.append(("M1", "Khối học"))
                if dang_hoc and is_m2_age_group and self._is_missing(row_values[16]):
                    missing_fields.append(("M2", "Khối học"))

                if birth_year and 2008 <= birth_year <= 2014:
                    if self._is_missing(row_values[19]):
                        missing_fields.append(("M2", "Bậc tốt nghiệp"))
                    if self._is_missing(row_values[21]):
                        missing_fields.append(("M2", "Năm tốt nghiệp"))

                if not self._is_missing(row_values[21]) and not self._is_academic_year(row_values[21]):
                    missing_fields.append(("M2", "Năm tốt nghiệp không đúng dạng YYYY-YYYY"))

                graduation_mismatch = self._graduation_class_mismatch(
                    row_values[17], row_values[16], row_values[19]
                )
                if graduation_mismatch:
                    graduation_level, class_grade = graduation_mismatch
                    missing_fields.append(
                        (
                            "M1, M2",
                            f"Cấp/lớp không phù hợp: lớp {class_grade} nhưng bậc tốt nghiệp {graduation_level}"
                        )
                    )

                completed_class = row_values[24]
                completed_year = row_values[25]
                if not self._is_missing(completed_class) and self._is_missing(completed_year):
                    missing_fields.append(("M1, M2", "Năm học xong"))
                if not self._is_missing(completed_year) and not self._is_academic_year(completed_year):
                    missing_fields.append(("M1, M2", "Năm học xong không đúng dạng YYYY-YYYY"))
                if not self._is_missing(completed_year) and self._is_missing(completed_class):
                    missing_fields.append(("M1, M2", "Lớp học xong"))

                dropout_class = row_values[26]
                dropout_year = row_values[27]
                if not self._is_missing(dropout_class) and self._is_missing(dropout_year):
                    missing_fields.append(("M1, M2", "Năm bỏ học"))
                if not self._is_missing(dropout_year) and not self._is_academic_year(dropout_year):
                    missing_fields.append(("M1, M2", "Năm bỏ học không đúng dạng YYYY-YYYY"))
                if not self._is_missing(dropout_year) and self._is_missing(dropout_class):
                    missing_fields.append(("M1, M2", "Lớp bỏ học"))

                if missing_fields:
                    for report, field in missing_fields:
                        self.review_records.append({
                            "row": row_number,
                            "so_phieu": "" if self._is_missing(so_phieu) else str(so_phieu),
                            "person": person_name,
                            "birth_year": "" if birth_year is None else birth_year,
                            "village": "" if self._is_missing(village) else str(village),
                            "report": report,
                            "field": field,
                            "reason": f"Thiếu {field}, có thể làm sai số liệu {report}"
                        })

            wb.close()
            total = len(self.review_records)
            self.lbl_review_status.config(text=f"Phát hiện {total} mục cần bổ sung")
            if total:
                preview = "\n".join(
                    f"- Dòng {record['row']}, {record['field']} ({record['report']})"
                    for record in self.review_records[:10]
                )
                if total > 10:
                    preview += f"\n... và {total - 10} mục khác"
                messagebox.showwarning(
                    "Có thông tin cần bổ sung",
                    f"Đã rà soát '{os.path.basename(source_file)}' và phát hiện {total} mục cần kiểm tra.\n\n{preview}\n\n"
                    "Bấm 'Xuất danh sách cần bổ sung...' để lưu đầy đủ chi tiết."
                )
            else:
                messagebox.showinfo(
                    "Rà soát hoàn tất",
                    f"Đã rà soát '{os.path.basename(source_file)}'. "
                    "Không phát hiện thông tin bắt buộc nào bị thiếu cho M1 và M2."
                )
        except Exception as e:
            self.review_records = []
            messagebox.showerror("Lỗi rà soát file Tổng", f"Không thể rà soát file Tổng:\n{e}")

    def export_review_report(self):
        if not self.review_records:
            messagebox.showinfo("Danh sách cần bổ sung", "Chưa có dữ liệu rà soát hoặc không phát hiện mục thiếu.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Lưu danh sách thông tin cần bổ sung",
            defaultextension=".xlsx",
            initialfile="DanhSachThongTinCanBoSung.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not save_path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Thông tin cần bổ sung"
            ws.append([
                "STT", "Dòng trong file Tổng", "Số phiếu", "Họ tên", "Năm sinh",
                "Thôn/bản", "Biểu", "Trường còn thiếu", "Nguyên nhân"
            ])
            for index, record in enumerate(self.review_records, start=1):
                ws.append([
                    index, record["row"], record["so_phieu"], record["person"],
                    record["birth_year"], record["village"], record["report"],
                    record["field"], record["reason"]
                ])
            ws.freeze_panes = "A2"
            widths = {"A": 8, "B": 18, "C": 16, "D": 30, "E": 12, "F": 25, "G": 12, "H": 25, "I": 60}
            for column, width in widths.items():
                ws.column_dimensions[column].width = width
            wb.save(save_path)
            messagebox.showinfo("Xuất danh sách thành công", f"Danh sách cần bổ sung đã được lưu tại:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Lỗi xuất danh sách", f"Không thể xuất danh sách cần bổ sung:\n{e}")

    def export_appendix(self):
        if not self.tong_file:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Tổng Hợp (Tong.xlsx) ở Bước 2 trước!")
            return

        template_path = self.pl_file or os.path.join(os.getcwd(), "PL.xlsx")
        if not os.path.exists(template_path):
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file mẫu PL.xlsx trước khi xuất phụ lục.")
            return

        aggregate.export_appendix(self.get_source_tong_file(), template_path)

    def export_disability(self):
        if not self.tong_file:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Tổng Hợp (Tong.xlsx) ở Bước 2 trước!")
            return

        template_path = self.kt_file or os.path.join(os.getcwd(), "KT.xlsx")
        if not os.path.exists(template_path):
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file mẫu KT.xlsx trước khi xuất danh sách khuyết tật.")
            return

        aggregate.export_disability(self.get_source_tong_file(), template_path)

    def run_aggregate_spc1(self):
        if not self.tong_file:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Tổng Hợp (Tong.xlsx) ở Bước 2 trước!")
            return
        if not self.spc1_file:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file mẫu SPC1.xlsx!")
            return
        aggregate.aggregate_spc1(self.get_source_tong_file(), self.spc1_file)

    def run_aggregate_primary(self):
        if not self.tong_file:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Tổng Hợp (Tong.xlsx) ở Bước 2 trước!")
            return
        if not self.spcth_file or not self.m1th_file or not self.xmth_file:
            messagebox.showwarning(
                "Cảnh báo",
                "Vui lòng chọn đủ 3 file mẫu SPCTH.xlsx, M1TH.xlsx và XMTH.xlsx!"
            )
            return
        aggregate.aggregate_primary(
            self.get_source_tong_file(),
            self.spcth_file,
            self.m1th_file,
            self.xmth_file
        )

    def run_aggregate(self):
        if not self.tong_file:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Tổng Hợp (Tong.xlsx) ở Bước 2 trước!")
            return
        if not self.m1_file or not self.m2_file:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn cả file mẫu M1.xlsx và M2.xlsx!")
            return
        aggregate.aggregate_m1_m2(self.get_source_tong_file(), self.m1_file, self.m2_file, "")

    def run_extraction(self):
        if not self.phieu_files:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất 1 file Phiếu!")
            return
        if not self.tong_file:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Tong.xlsx!")
            return
            
        success_count = 0
        error_count = 0
        self.error_records = []
        
        tong_dir = os.path.dirname(self.tong_file)
        ketqua_tong_file = os.path.join(tong_dir, "KetQua_Tong.xlsx")
        
        # Tự động xoá toàn bộ các biểu kết quả cũ để cập nhật kết quả mới nhất
        old_result_files = [
            os.path.join(tong_dir, "KetQua_Tong.xlsx"),
            os.path.join(tong_dir, "KetQua_M1.xlsx"),
            os.path.join(tong_dir, "KetQua_M2.xlsx"),
            os.path.join(tong_dir, "KetQua_SPC1.xlsx"),
            os.path.join(tong_dir, "KetQua_SPC.xlsx"),
            os.path.join(tong_dir, "KetQua_SPCTH.xlsx"),
            os.path.join(tong_dir, "KetQua_M1TH.xlsx"),
            os.path.join(tong_dir, "KetQua_XMTH.xlsx"),
        ]
        for f in old_result_files:
            if os.path.exists(f) and os.path.abspath(f) != os.path.abspath(self.tong_file):
                try:
                    os.remove(f)
                except PermissionError:
                    messagebox.showerror(
                        "Lỗi file đang mở",
                        f"File '{os.path.basename(f)}' đang được mở trong Excel hoặc chương trình khác.\n"
                        "Vui lòng đóng file đó lại rồi bấm 'Bắt đầu tổng hợp'!"
                    )
                    return
                except Exception as e:
                    messagebox.showerror("Lỗi xoá file cũ", f"Không thể xoá file cũ '{os.path.basename(f)}':\n{e}")
                    return

        import shutil
        try:
            shutil.copy2(self.tong_file, ketqua_tong_file)
            # Làm sạch dữ liệu từ dòng 5 trở đi nếu file mẫu có chứa dữ liệu cũ
            wb_init = openpyxl.load_workbook(ketqua_tong_file)
            ws_init = wb_init.active
            if ws_init.max_row > 4:
                ws_init.delete_rows(5, ws_init.max_row - 4)
                wb_init.save(ketqua_tong_file)
            wb_init.close()
            self.last_tong_result = ketqua_tong_file
        except Exception as e:
            messagebox.showerror("Lỗi tạo file Tổng", f"Không thể tạo file Tổng:\n{e}")
            return
            
        self.progress_bar['maximum'] = len(self.phieu_files)
        self.progress_bar['value'] = 0

        wb_result = openpyxl.load_workbook(ketqua_tong_file)
        
        with open("debug_log.txt", "w", encoding="utf-8") as log:
            for idx, file in enumerate(self.phieu_files):
                self.progress_var.set(f"Đang xử lý: {os.path.basename(file)} ({idx+1}/{len(self.phieu_files)})")
                self.progress_bar['value'] = idx
                self.root.update_idletasks()
                
                log.write(f"Processing {file}\n")
                gen_info, people, parse_error = parse_phieu(file)
                log.write(f"Parsed: gen_info={bool(gen_info)}, people count={len(people) if people else 0}\n")
                
                if gen_info and people:
                    try:
                        write_to_tong(gen_info, people, ketqua_tong_file, workbook=wb_result)
                        success_count += 1
                        log.write("Write success\n")
                    except Exception as e:
                        error_count += 1
                        reason = f"Không ghi được dữ liệu vào file Tổng: {e}"
                        self.error_records.append({
                            "file_name": os.path.basename(file),
                            "file_path": file,
                            "reason": reason
                        })
                        log.write(f"Write error: {reason}\n")
                else:
                    error_count += 1
                    reason = parse_error or "Không tìm thấy dữ liệu người trong phiếu"
                    self.error_records.append({
                        "file_name": os.path.basename(file),
                        "file_path": file,
                        "reason": reason
                    })
                    log.write(f"Failed: {reason}\n")

                wb_result.save(ketqua_tong_file)
                wb_result.close()
                
        self.progress_bar['value'] = len(self.phieu_files)
        self.progress_var.set(f"Hoàn thành! Thành công: {success_count}, Lỗi: {error_count}")
        if self.error_records:
            error_lines = "\n".join(
                f"- {record['file_name']}: {record['reason']}"
                for record in self.error_records[:10]
            )
            if len(self.error_records) > 10:
                error_lines += f"\n... và {len(self.error_records) - 10} phiếu lỗi khác"
            messagebox.showwarning(
                "Tổng hợp hoàn thành, có phiếu lỗi",
                f"Đã tổng hợp xong!\nThành công: {success_count}\nLỗi: {error_count}\n\n"
                f"Danh sách lỗi:\n{error_lines}\n\n"
                "Bấm 'Xuất danh sách phiếu lỗi...' để lưu đầy đủ nguyên nhân."
            )
        else:
            messagebox.showinfo("Hoàn thành", f"Đã tổng hợp xong!\nThành công: {success_count}\nLỗi: 0")

if __name__ == "__main__":
    if not GUI_AVAILABLE:
        raise SystemExit("Desktop GUI is unavailable in this environment. Use the web app via 'python web_app.py'.")
    app_root = ttk.Window(themename="cosmo") # Use a modern theme
    app = App(app_root)
    app_root.mainloop()
