import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import os
import re
from datetime import date

try:
    from tkinter import messagebox
except ImportError:
    class _HeadlessMessageBox:
        @staticmethod
        def showinfo(*args, **kwargs):
            return None

        @staticmethod
        def showwarning(*args, **kwargs):
            return None

        @staticmethod
        def showerror(*args, **kwargs):
            return None

    messagebox = _HeadlessMessageBox()

def aggregate_m1_m2(tong_file, m1_template_path, m2_template_path, output_dir):
    try:
        tong_dir = os.path.dirname(tong_file)
        out_m1 = os.path.join(tong_dir, "KetQua_M1.xlsx")
        out_m2 = os.path.join(tong_dir, "KetQua_M2.xlsx")

        # Tự động xoá các file kết quả M1, M2 cũ trước khi tổng hợp mới
        for f in [out_m1, out_m2]:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except PermissionError:
                    messagebox.showerror(
                        "Lỗi file đang mở",
                        f"File '{os.path.basename(f)}' đang được mở trong Excel hoặc phần mềm khác.\n"
                        "Vui lòng đóng file đó lại rồi bấm 'Tổng hợp M1 & M2'!"
                    )
                    return
                except Exception as e:
                    messagebox.showerror("Lỗi xoá file cũ", f"Không thể xoá file cũ '{os.path.basename(f)}':\n{e}")
                    return

        if not os.path.exists(m1_template_path) or not os.path.exists(m2_template_path):
            messagebox.showerror("Lỗi", "Không tìm thấy file mẫu M1.xlsx hoặc M2.xlsx")
            return
            
        # 1. READ TONG.XLSX DATA
        df_tong = pd.read_excel(tong_file, header=None, skiprows=4)
        df_tong[5] = pd.to_numeric(df_tong[5], errors='coerce')
        
        # M1 Aggregation
        wb_m1 = openpyxl.load_workbook(m1_template_path)
        sheet_name_m1 = 'Thống kê PCGD THCS' if 'Thống kê PCGD THCS' in wb_m1.sheetnames else wb_m1.sheetnames[0]
        sheet_m1 = wb_m1[sheet_name_m1]
        
        year_col_map = {}
        for c in range(5, 30):
            val = sheet_m1.cell(row=6, column=c).value
            try:
                year = int(val)
                year_col_map[year] = c
            except:
                pass

        disability_columns = [31, 32, 33, 34, 35, 36, 37, 38, 39, 40]

        def is_disabled(row):
            return row[disability_columns].notna().any()

        def is_thcs_graduate(row):
            bac = str(row[19]).strip().upper()
            return bac in ['THCS', 'THPT']

        def is_upper_study(row):
            khoi = str(row[16]).strip()
            bo_tuc = str(row[20]).strip().lower()
            nghe = str(row[22]).strip().lower()
            return (
                khoi in ['10', '11', '12'] or
                bo_tuc not in ['', 'nan', 'none', '0'] or
                nghe not in ['', 'nan', 'none', '0']
            )

        def is_local_study(row):
            school = str(row[18]).strip()
            return 'hùng lợi' in school.lower()

        def is_other_place_study(row):
            school = str(row[18]).strip()
            return bool(school) and school.lower() not in {'nan', 'none'} and not is_local_study(row)

        def has_school_name(row):
            school = str(row[18]).strip().lower()
            return school not in {'', 'nan', 'none'}

        def has_grade(row):
            grade = str(row[16]).strip().lower()
            return grade not in {'', 'nan', 'none'}

        def is_dropout(row):
            return any(str(row[index]).strip().lower() not in ['', 'nan', 'none', '0'] for index in (26, 27))

        def is_local_or_no_school(row):
            school = str(row[18]).strip().lower()
            return school in ['', 'nan', 'none'] or 'hùng lợi' in school

        def detail_counts(dataframe, age_group=None):
            eligible = dataframe[~dataframe.apply(is_disabled, axis=1)]

            if age_group == '11-14':
                dropout_rows = eligible[eligible.apply(is_dropout, axis=1)]
                if dropout_rows.empty:
                    return {34: None, 35: None, 36: None}

                local_count = int(dropout_rows.apply(is_local_study, axis=1).sum())
                other_place_count = int(
                    dropout_rows.apply(is_other_place_study, axis=1).sum()
                )
                return {
                    34: local_count or None,
                    35: other_place_count or None,
                    36: None,
                }

            # Rows 28-30 must reconcile exactly with G41, so TNTHCS is based
            # only on Bậc tốt nghiệp THCS/THPT and may overlap with row 31.
            graduate = eligible.apply(is_thcs_graduate, axis=1)
            graduate_rows = eligible[graduate]
            named_school_without_grade = graduate_rows.apply(
                lambda row: has_school_name(row) and not has_grade(row),
                axis=1,
            )
            studying_graduates = graduate_rows.apply(is_upper_study, axis=1) | named_school_without_grade
            counts = {}

            # All TNTHCS records are reported as "Tại chỗ".
            counts[28] = int(graduate.sum())
            counts[29] = 0
            counts[30] = 0

            # A studying record with a school name but no grade is reported
            # as "Đi học nơi khác"; the remaining studying records are local.
            studying_rows = graduate_rows[studying_graduates]
            elsewhere = studying_rows.apply(
                lambda row: has_school_name(row) and not has_grade(row),
                axis=1,
            )
            counts[31] = int(len(studying_rows) - elsewhere.sum())
            counts[32] = int(elsewhere.sum())
            counts[33] = 0

            # For ages 15-18, rows 34-36 contain only records explicitly
            # marked as dropout. THPT graduates are not counted as dropout.
            dropout_rows = eligible[eligible.apply(is_dropout, axis=1)]
            counts[34] = int(len(dropout_rows))
            counts[35] = 0
            counts[36] = 0
            return counts

        def grade_counts(dataframe, grade_value):
            eligible = dataframe[~dataframe.apply(is_disabled, axis=1)]
            grade_rows = eligible[eligible[16].astype(str).str.strip() == str(grade_value)]
            local_count = int(grade_rows.apply(is_local_study, axis=1).sum())
            other_place_count = int(grade_rows.apply(is_other_place_study, axis=1).sum())
            incoming_count = int(len(grade_rows) - local_count - other_place_count)
            return local_count, other_place_count, incoming_count

        def write_three_rows(first_row, counts, column):
            for offset, count in enumerate(counts):
                sheet_m1.cell(row=first_row + offset, column=column).value = count if count else None
                
        for year, col_idx in year_col_map.items():
            df_year = df_tong[df_tong[5] == year]
            if df_year.empty:
                for detail_row in list(range(8, 15)) + list(range(16, 37)):
                    sheet_m1.cell(row=detail_row, column=col_idx).value = None
                continue
            eligible_year = df_year[~df_year.apply(is_disabled, axis=1)]
            
            total = len(df_year)
            sheet_m1.cell(row=8, column=col_idx).value = total if total > 0 else None
            
            nu_count = df_year[df_year[6].astype(str).str.contains(r'[xqXQ]', na=False)].shape[0]
            sheet_m1.cell(row=9, column=col_idx).value = nu_count if nu_count > 0 else None
            
            dt_count = df_year[
                (df_year[7].notna()) & 
                (df_year[7].astype(str).str.strip() != "") & 
                (~df_year[7].astype(str).str.lower().str.contains("kinh"))
            ].shape[0]
            sheet_m1.cell(row=10, column=col_idx).value = dt_count if dt_count > 0 else None
            
            def get_hs_counts(khoi):
                return grade_counts(df_year, khoi)

            write_three_rows(16, get_hs_counts(6), col_idx)
            
            write_three_rows(19, get_hs_counts(7), col_idx)
            
            write_three_rows(22, get_hs_counts(8), col_idx)
            
            write_three_rows(25, get_hs_counts(9), col_idx)

            # Row 11: Khuyết tật - Tổng số (các cột loại, chứng nhận, khả năng học tập)
            kt_cols = disability_columns
            kt_count = df_year[df_year[kt_cols].notna().any(axis=1)].shape[0]
            sheet_m1.cell(row=11, column=col_idx).value = kt_count if kt_count > 0 else None

            # Row 12: Khuyết tật có khả năng học tập (cột 40 = 'Khả năng học tập' != rỗng/None)
            kt_ht_count = df_year[
                df_year[40].notna() & (df_year[40].astype(str).str.strip() != '')
            ].shape[0]
            sheet_m1.cell(row=12, column=col_idx).value = kt_ht_count if kt_ht_count > 0 else None

            # Row 13: Khuyết tật tiếp cận GD (cột 39 = 'Có chứng nhận khuyết tật')
            kt_cg_count = df_year[
                df_year[39].notna() & (df_year[39].astype(str).str.strip() != '')
            ].shape[0]
            sheet_m1.cell(row=13, column=col_idx).value = kt_cg_count if kt_cg_count > 0 else None

            # Row 14: Số phải phổ cập = tổng số trừ đối tượng khuyết tật
            ppc_total = total - kt_count
            sheet_m1.cell(row=14, column=col_idx).value = ppc_total if ppc_total > 0 else None

            # Rows 28-36: detail of the 15-18 age group, excluding disability.
            rows_15_18 = eligible_year
            if 2011 <= year <= 2014:
                for detail_row, count in detail_counts(df_year, '11-14').items():
                    sheet_m1.cell(row=detail_row, column=col_idx).value = count
            elif 2008 <= year <= 2015:
                is_studying = rows_15_18.apply(is_upper_study, axis=1)
                is_dropout_mask = rows_15_18.apply(is_dropout, axis=1)
                is_graduate = rows_15_18.apply(is_thcs_graduate, axis=1) & ~is_studying & ~is_dropout_mask
                is_other = ~(is_studying | is_graduate | is_dropout_mask)

                age_group = '15-18'
                for detail_row, count in detail_counts(df_year, age_group).items():
                    sheet_m1.cell(row=detail_row, column=col_idx).value = count if count else None
            else:
                for detail_row in range(28, 37):
                    sheet_m1.cell(row=detail_row, column=col_idx).value = None

            # Column Z (col 26): tổng 15-18 tuổi (năm sinh 2007-2010 tính theo năm hiện tại 2025)
            if year in [2007, 2008, 2009, 2010]:
                # Z column = SUM of all years 15-18 for this row
                pass  # will be set after loop with formula

        # M1: Cột Z (col 26) = Tổng 15-18 tuổi (sinh 2008-2011)
        # Tính trực tiếp từ dữ liệu, không dùng công thức ô
        df_15_18_all = df_tong[(df_tong[5] >= 2008) & (df_tong[5] <= 2011)]

        def z_total(field_fn):
            return field_fn(df_15_18_all)

        def z_nu(df): return df[df[6].astype(str).str.contains(r'[xqXQ]', na=False)].shape[0]
        def z_dt(df): return df[
            df[7].notna() & (df[7].astype(str).str.strip() != '') &
            (~df[7].astype(str).str.lower().str.contains('kinh'))
        ].shape[0]
        def z_kt(df): return df[df[[31,32,33,34,35,36,37,38,39,40]].notna().any(axis=1)].shape[0]
        def z_kt_ht(df): return df[df[40].notna() & (df[40].astype(str).str.strip() != '')].shape[0]
        def z_kt_cg(df): return df[df[39].notna() & (df[39].astype(str).str.strip() != '')].shape[0]
        def z_ppc(df):
            disabled = df[[31, 32, 33, 34, 35, 36, 37, 38, 39, 40]].notna().any(axis=1)
            return len(df[~disabled])

        z_vals = {
            8:  len(df_15_18_all),
            9:  z_nu(df_15_18_all),
            10: z_dt(df_15_18_all),
            11: z_kt(df_15_18_all),
            12: z_kt_ht(df_15_18_all),
            13: z_kt_cg(df_15_18_all),
            14: z_ppc(df_15_18_all),
        }
        # lớp 6-9 (đang học THCS ở địa phương)
        for khoi, rows in [(6, (16,17,18)), (7, (19,20,21)), (8, (22,23,24)), (9, (25,26,27))]:
            eligible_15_18 = df_15_18_all[~df_15_18_all.apply(is_disabled, axis=1)]
            df_k = eligible_15_18[eligible_15_18[16].astype(str).str.strip() == str(khoi)]
            tc = df_k[df_k[18].astype(str).str.contains('Hùng Lợi', case=False, na=False)]
            dh = df_k[~df_k[18].astype(str).str.contains('Hùng Lợi', case=False, na=False) & df_k[18].notna()]
            z_vals[rows[0]], z_vals[rows[1]], z_vals[rows[2]] = grade_counts(df_15_18_all, khoi)

        z_vals.update(detail_counts(df_15_18_all, '15-18'))

        for data_row, val in z_vals.items():
            try:
                sheet_m1.cell(row=data_row, column=26).value = val if val else None
            except Exception:
                pass

        # M1: G41:H44 - các tiêu chí cuối bảng.
        df_15_18 = df_tong[(df_tong[5] >= 2008) & (df_tong[5] <= 2011)]
        disabled_15_18 = df_15_18.apply(is_disabled, axis=1)
        eligible_15_18 = df_15_18[~disabled_15_18]
        ts_15_18 = len(eligible_15_18)
        upper_mask = eligible_15_18.apply(is_upper_study, axis=1)
        tn_thcs_mask = eligible_15_18.apply(is_thcs_graduate, axis=1)
        tn_thcs_15_18 = int(tn_thcs_mask.sum())
        study_detail = detail_counts(df_15_18, '15-18')
        upper_study_15_18 = int((study_detail.get(31) or 0) + (study_detail.get(32) or 0))
        kt_tiep_can_count = int(df_15_18[disability_columns].notna().any(axis=1).sum())
        df_11_18 = df_tong[(df_tong[5] >= 2008) & (df_tong[5] <= 2015)]
        eligible_11_18 = df_11_18[~df_11_18.apply(is_disabled, axis=1)]

        sheet_m1['A41'] = 'TTN 15-18 TNTHCS'
        sheet_m1['A42'] = 'TTN 15-18 đã, đang học CT GDPT hoặc GDTX cấp THPT hoặc GDNN'
        sheet_m1['A43'] = 'TTN KT được tiếp cận GD'
        sheet_m1['A44'] = 'Tổng sô học sinh'
        criteria_values = {
            41: (tn_thcs_15_18, ts_15_18),
            42: (upper_study_15_18, ts_15_18),
            43: (kt_tiep_can_count, ts_15_18),
            44: (len(eligible_11_18), len(eligible_11_18)),
        }
        for row, (count, denominator) in criteria_values.items():
            sheet_m1.cell(row=row, column=7).value = count if count > 0 else None
            sheet_m1.cell(row=row, column=8).value = round(count / denominator * 100, 1) if denominator > 0 else None

        out_m1 = os.path.join(tong_dir, "KetQua_M1.xlsx")
        wb_m1.save(out_m1)

        # M2 Aggregation
        wb_m2 = openpyxl.load_workbook(m2_template_path)
        sheet_name_m2 = 'Tiêu chuẩn PCGD THCS' if 'Tiêu chuẩn PCGD THCS' in wb_m2.sheetnames else wb_m2.sheetnames[0]
        sheet_m2 = wb_m2[sheet_name_m2]
        
        villages = df_tong[12].dropna().unique()
        villages = [str(v).strip() for v in villages if str(v).strip() != ""]
        
        start_row_m2 = 12
        
        # Năm học vừa qua: lấy năm học hiện tại trong biểu mẫu rồi lùi một năm
        nam_hoc_hien_tai = '2024-2025'
        nam_hoc_mau = str(sheet_m2['A3'].value or '')
        nam_hoc_match = re.search(r'(\d{4})\s*-\s*(\d{4})', nam_hoc_mau)
        if nam_hoc_match:
            nam_bat_dau = int(nam_hoc_match.group(1)) - 1
            nam_hoc_hien_tai = f'{nam_bat_dau}-{nam_bat_dau + 1}'
        
        for idx, v_name in enumerate(villages):
            m2_row = start_row_m2 + idx
            if m2_row > 19:
                break
            
            sheet_m2.cell(row=m2_row, column=1).value = idx + 1
            sheet_m2.cell(row=m2_row, column=2).value = v_name
            
            df_village = df_tong[df_tong[12].astype(str).str.contains(v_name, case=False, na=False)]
            
            def get_percentage(part, whole):
                return round((part / whole) * 100, 1) if whole > 0 else None
                
            # 1. Trẻ 6 tuổi (Sinh năm 2019)
            tre_6t = df_village[df_village[5] == 2019]
            ts_6t = len(tre_6t)
            # Đã vào học lớp 1: có thể đang lớp 1 (khoi=1) hoặc đã lên lớp cao hơn (khoi>=2)
            # Chỉ các em không có lớp học (khoi trống) mới được coi là chưa vào học
            chua_di_hoc = tre_6t[
                tre_6t[16].isna() |
                (tre_6t[16].astype(str).str.strip() == '') |
                (tre_6t[16].astype(str).str.strip().isin(['None', 'nan', '0']))
            ]
            vao_lop_1 = ts_6t - len(chua_di_hoc)
            sheet_m2.cell(row=m2_row, column=3).value = ts_6t or None
            sheet_m2.cell(row=m2_row, column=4).value = vao_lop_1 or None
            sheet_m2.cell(row=m2_row, column=5).value = get_percentage(vao_lop_1, ts_6t)
            
            # 2. HS Tốt nghiệp TH năm học vừa qua
            def bac_la_th(val):
                v = str(val).strip().upper()
                return v in ['TH', 'TIỂU HỌC', 'TIEU HOC']
            
            tn_th = df_village[
                df_village[19].apply(bac_la_th) &
                (df_village[21].astype(str).str.strip() == nam_hoc_hien_tai)
            ]
            ts_tn_th = len(tn_th)

            # Khối học có thể được ghi là "6", "6A", "6B"...
            khoi_6 = tn_th[16].astype(str).str.strip().str.extract(r'^(6)(?:\D|$)', expand=False).notna()
            bo_tuc_value = tn_th[20].fillna('').astype(str).str.strip().str.lower()
            la_gdtx = ~bo_tuc_value.isin(['', 'nan', 'none', '0'])
            vao_lop_6_pt = len(tn_th[khoi_6 & ~la_gdtx])
            vao_lop_6_gdtx = len(tn_th[khoi_6 & la_gdtx])
            sheet_m2.cell(row=m2_row, column=6).value = ts_tn_th or None
            sheet_m2.cell(row=m2_row, column=7).value = vao_lop_6_pt or None
            sheet_m2.cell(row=m2_row, column=8).value = vao_lop_6_gdtx or None
            sheet_m2.cell(row=m2_row, column=9).value = (vao_lop_6_pt + vao_lop_6_gdtx) or None
            sheet_m2.cell(row=m2_row, column=10).value = get_percentage(vao_lop_6_pt + vao_lop_6_gdtx, ts_tn_th) # Tỷ lệ
            
            # 3. Trẻ độ tuổi 11-14 (Sinh 2011 - 2014)
            tre_11_14 = df_village[(df_village[5] >= 2011) & (df_village[5] <= 2014)]
            ts_11_14 = len(tre_11_14)
            
            def da_tot_nghiep_th(row):
                khoi = str(row[16]).strip()
                bac = str(row[19]).strip().upper()
                if khoi in ['6','7','8','9','10','11','12'] or bac in ['TH', 'THCS', 'THPT']:
                    return True
                return False
                
            htct_th = len(tre_11_14[tre_11_14.apply(da_tot_nghiep_th, axis=1)])
            
            sheet_m2.cell(row=m2_row, column=11).value = ts_11_14 or None
            sheet_m2.cell(row=m2_row, column=12).value = htct_th or None
            sheet_m2.cell(row=m2_row, column=13).value = get_percentage(htct_th, ts_11_14)
            
            # 4. HS Lớp 9 năm qua TN THCS
            def bac_la_thcs(val):
                v = str(val).strip().upper()
                return v in ['THCS', 'TRUNG HỌC CƠ SỞ', 'THCS BỔ TÚC']
            
            tn_thcs = df_village[
                df_village[19].apply(bac_la_thcs) &
                (df_village[21].astype(str).str.strip() == nam_hoc_hien_tai)
            ]
            ts_tn_thcs = len(tn_thcs)
            sheet_m2.cell(row=m2_row, column=14).value = ts_tn_thcs or None
            sheet_m2.cell(row=m2_row, column=15).value = ts_tn_thcs or None
            sheet_m2.cell(row=m2_row, column=16).value = None # GDTX
            sheet_m2.cell(row=m2_row, column=17).value = get_percentage(ts_tn_thcs, ts_tn_thcs)
            
            # 5. Đối tượng 15-18 tuổi (Sinh 2008 - 2011)
            tre_15_18 = df_village[(df_village[5] >= 2008) & (df_village[5] <= 2011)]
            ts_15_18 = len(tre_15_18)
            
            def da_tot_nghiep_thcs(row):
                khoi = str(row[16]).strip()
                bac = str(row[19]).strip().upper()
                if khoi in ['10','11','12'] or bac in ['THCS', 'THPT']:
                    return True
                return False
                
            co_bang_thcs = len(tre_15_18[tre_15_18.apply(da_tot_nghiep_thcs, axis=1)])
            sheet_m2.cell(row=m2_row, column=18).value = ts_15_18 or None
            sheet_m2.cell(row=m2_row, column=19).value = co_bang_thcs or None # PT
            sheet_m2.cell(row=m2_row, column=20).value = None # GDTX
            sheet_m2.cell(row=m2_row, column=21).value = co_bang_thcs or None # Cộng
            sheet_m2.cell(row=m2_row, column=22).value = get_percentage(co_bang_thcs, ts_15_18)
            
        from openpyxl.styles import Alignment
        center_align = Alignment(horizontal='center', vertical='center')

        # 1. Đảm bảo các chỉ tiêu chuẩn ở cột L và N
        standards_L = {25: '90%', 26: '80%', 27: '95%', 28: '90%', 29: '80%'}
        standards_N = {25: '80%', 26: '70%', 27: '80%', 28: '75%', 29: '70%'}
        for r, val in standards_L.items():
            if not sheet_m2.cell(row=r, column=12).value:
                sheet_m2.cell(row=r, column=12).value = val
            sheet_m2.cell(row=r, column=12).alignment = center_align
        for r, val in standards_N.items():
            if not sheet_m2.cell(row=r, column=14).value:
                sheet_m2.cell(row=r, column=14).value = val
            sheet_m2.cell(row=r, column=14).alignment = center_align
        # Row 30 is merged across B:Q; write only to its top-left cell.
        if not sheet_m2['B30'].value:
            sheet_m2['B30'] = 'Vùng KK, ĐBKK'
        sheet_m2['B30'].alignment = center_align

        # 2. Đánh giá cột W (W12 đến W19 cho từng thôn): Đạt hay chưa đạt PCGD THCS
        for r in range(12, 20):
            cell_w = sheet_m2.cell(row=r, column=23)
            cell_w.value = (
                f'=IF(B{r}<>"", IF(AND('
                f'IF(C{r}>0, E{r}>=IF(ISNUMBER(N$25), IF(N$25<=1, N$25*100, N$25), VALUE(SUBSTITUTE(N$25, "%", ""))), TRUE), '
                f'IF(K{r}>0, M{r}>=IF(ISNUMBER(N$26), IF(N$26<=1, N$26*100, N$26), VALUE(SUBSTITUTE(N$26, "%", ""))), TRUE), '
                f'IF(F{r}>0, J{r}>=IF(ISNUMBER(N$27), IF(N$27<=1, N$27*100, N$27), VALUE(SUBSTITUTE(N$27, "%", ""))), TRUE), '
                f'IF(N{r}>0, Q{r}>=IF(ISNUMBER(N$28), IF(N$28<=1, N$28*100, N$28), VALUE(SUBSTITUTE(N$28, "%", ""))), TRUE), '
                f'IF(R{r}>0, V{r}>=IF(ISNUMBER(N$29), IF(N$29<=1, N$29*100, N$29), VALUE(SUBSTITUTE(N$29, "%", ""))), TRUE)), '
                f'"Đạt", "Chưa đạt"), "")'
            )
            cell_w.alignment = center_align

        # 3. Dòng 20: Công thức tổng cộng cho toàn xã
        sum_cols = ['C', 'D', 'F', 'G', 'H', 'I', 'K', 'L', 'N', 'O', 'P', 'R', 'S', 'T', 'U']
        for col in sum_cols:
            sheet_m2[f'{col}20'] = f'=IF(COUNT({col}12:{col}19)=0,"",SUM({col}12:{col}19))'
        sheet_m2['E20'] = '=IF(C20>0, ROUND(D20/C20*100, 1), "")'
        sheet_m2['J20'] = '=IF(F20>0, ROUND(I20/F20*100, 1), "")'
        sheet_m2['M20'] = '=IF(K20>0, ROUND(L20/K20*100, 1), "")'
        sheet_m2['Q20'] = '=IF(N20>0, ROUND((O20+P20)/N20*100, 1), "")'
        sheet_m2['V20'] = '=IF(R20>0, ROUND(U20/R20*100, 1), "")'
        sheet_m2['W20'] = '=IF(COUNTIF(P25:P29, "Đạt")=5, "Đạt", "Chưa đạt")'
        sheet_m2['W20'].alignment = center_align

        # 4. Tính tỉ số (cột I) và tỷ lệ % (cột K) từ dòng 25 đến 29
        # Hàng 25: Trẻ 6 tuổi vào học lớp 1
        sheet_m2['I25'] = '=IF(C20>0, D20&"/"&C20, "")'
        sheet_m2['K25'] = '=IF(C20>0, ROUND(D20/C20*100, 1), "")'

        # Hàng 26: Trẻ 11-14 tuổi hoàn thành chương trình tiểu học
        sheet_m2['I26'] = '=IF(K20>0, L20&"/"&K20, "")'
        sheet_m2['K26'] = '=IF(K20>0, ROUND(L20/K20*100, 1), "")'

        # Hàng 27: HS tốt nghiệp hoàn thành chương trình vào học lớp 6
        sheet_m2['I27'] = '=IF(F20>0, I20&"/"&F20, "")'
        sheet_m2['K27'] = '=IF(F20>0, ROUND(I20/F20*100, 1), "")'

        # Hàng 28: HS TN THCS năm học vừa qua
        sheet_m2['I28'] = '=IF(N20>0, (O20+P20)&"/"&N20, "")'
        sheet_m2['K28'] = '=IF(N20>0, ROUND((O20+P20)/N20*100, 1), "")'

        # Hàng 29: Thanh thiếu niên 15-18 tuổi TN THCS
        sheet_m2['I29'] = '=IF(R20>0, U20&"/"&R20, "")'
        sheet_m2['K29'] = '=IF(R20>0, ROUND(U20/R20*100, 1), "")'

        for r in range(25, 30):
            sheet_m2[f'I{r}'].alignment = center_align
            sheet_m2[f'K{r}'].alignment = center_align

        # 5. Căn cứ vào chỉ tiêu từ N25 đến N29 đánh giá P25 đến P29 là Đạt hay Chưa đạt
        for r in range(25, 30):
            sheet_m2[f'P{r}'] = (
                f'=IF(K{r}>=IF(ISNUMBER(N{r}), IF(N{r}<=1, N{r}*100, N{r}), VALUE(SUBSTITUTE(N{r}, "%", ""))), '
                f'"Đạt", "Chưa đạt")'
            )
            sheet_m2[f'P{r}'].alignment = center_align

        # Row 30 is a merged explanatory label in the template; do not write
        # the evaluation result into its merged child cells.

        out_m2 = os.path.join(tong_dir, "KetQua_M2.xlsx")
        wb_m2.save(out_m2)
        
        messagebox.showinfo("Thành công", f"Đã tổng hợp thành công!\nKết quả được lưu tại:\n1. {out_m1}\n2. {out_m2}")
        
    except Exception as e:
        messagebox.showerror("Lỗi tổng hợp M1/M2", f"Chi tiết lỗi: {e}")


def export_appendix(tong_file, pl_template_path):
    """Xuất phụ lục theo mẫu PL.xlsx dựa trên file Tổng đã tổng hợp."""
    try:
        if not os.path.exists(pl_template_path):
            raise FileNotFoundError(f"Không tìm thấy file mẫu PL.xlsx tại: {pl_template_path}")

        df_tong = pd.read_excel(tong_file, header=None, skiprows=4)
        if df_tong.empty:
            raise ValueError("File Tổng chưa có dữ liệu")

        grouped = {}
        for _, row in df_tong.iterrows():
            # Cấu trúc thực tế của file Tong.xlsx:
            # 10 = họ đệm chủ hộ, 11 = tên chủ hộ, 12 = tên thôn/bản/địa chỉ, 13 = số phiếu
            if pd.isna(row[10]) and pd.isna(row[11]) and pd.isna(row[12]) and pd.isna(row[13]):
                continue

            ho_dem = str(row[10]).strip() if pd.notna(row[10]) else ""
            ten = str(row[11]).strip() if pd.notna(row[11]) else ""
            chu_ho_name = f"{ho_dem} {ten}".strip()
            dia_chi = str(row[12]).strip() if pd.notna(row[12]) else ""

            so_phieu = row[13]
            if pd.isna(so_phieu):
                so_phieu = ""
            else:
                so_phieu = str(int(so_phieu)) if isinstance(so_phieu, float) and so_phieu.is_integer() else str(so_phieu).strip()

            # Ưu tiên số phiếu làm định danh hộ vì đó là khóa duy nhất và ổn định nhất.
            # Nếu thiếu số phiếu thì fallback về tên chủ hộ + địa chỉ để tránh tách hộ sai.
            household_key = so_phieu if so_phieu else (chu_ho_name, dia_chi)
            grouped.setdefault(household_key, {"name": chu_ho_name, "so_phieu": so_phieu, "dia_chi": dia_chi, "count": 0})
            grouped[household_key]["count"] += 1
            if not grouped[household_key]["name"] and chu_ho_name:
                grouped[household_key]["name"] = chu_ho_name
            if not grouped[household_key]["dia_chi"] and dia_chi:
                grouped[household_key]["dia_chi"] = dia_chi
            if not grouped[household_key]["so_phieu"] and so_phieu:
                grouped[household_key]["so_phieu"] = so_phieu

        # Chuẩn hóa lại để luôn có một tuple ổn định khi xuất file.
        normalized_rows = []
        for household_key, record in grouped.items():
            name = str(record["name"]).strip()
            so_phieu = str(record["so_phieu"]).strip()
            dia_chi = str(record["dia_chi"]).strip()
            if not name:
                continue
            normalized_rows.append(((name, so_phieu, dia_chi), int(record["count"])))

        def ticket_sort_key(item):
            ticket = item[0][1]
            match = re.search(r'\d+', str(ticket))
            if match:
                return (0, int(match.group()), str(ticket))
            return (1, float('inf'), str(ticket))

        rows = sorted(normalized_rows, key=ticket_sort_key)

        wb = openpyxl.load_workbook(pl_template_path)
        ws = wb.active

        for row in range(5, ws.max_row + 1):
            for col in range(1, 8):
                ws.cell(row=row, column=col).value = None

        ws['A2'] = "PHỤ LỤC XÓM NÀ MỘ, HÙNG LỢI, HUYỆN YÊN SƠN, TUYÊN QUANG"
        ws['A4'] = 'STT'
        ws['B4'] = 'Họ và tên chủ hộ'
        ws['C4'] = 'Số phiếu'
        ws['D4'] = 'Số đối tượng'
        ws['E4'] = 'Địa chỉ'

        for index, ((chu_ho_name, so_phieu, dia_chi), so_doi_tuong) in enumerate(rows, start=1):
            row = 4 + index
            ws.cell(row=row, column=1, value=index)
            ws.cell(row=row, column=2, value=chu_ho_name)
            ws.cell(row=row, column=3, value=so_phieu)
            ws.cell(row=row, column=4, value=so_doi_tuong)
            ws.cell(row=row, column=5, value=dia_chi)

        out_dir = os.path.dirname(tong_file)
        out_path = os.path.join(out_dir, "KetQua_PhụLục.xlsx")
        wb.save(out_path)
        messagebox.showinfo("Thành công", f"Đã xuất phụ lục theo mẫu PL.xlsx tại:\n{out_path}")
    except Exception as e:
        messagebox.showerror("Lỗi xuất phụ lục", f"Chi tiết lỗi: {e}")


def export_disability(tong_file, kt_template_path):
    """Xuất danh sách đối tượng khuyết tật theo mẫu KT.xlsx."""
    try:
        if not os.path.exists(kt_template_path):
            raise FileNotFoundError(f"Không tìm thấy file mẫu KT.xlsx tại: {kt_template_path}")

        df_tong = pd.read_excel(tong_file, header=None, skiprows=4)
        # Excel columns 32-41 (DataFrame index 31-40) are the disability flags.
        source_disability_cols = list(range(31, 41))
        disabled = df_tong[df_tong[source_disability_cols].notna().any(axis=1)].reset_index(drop=True)

        # Keep the same order as the template labels for notes.
        disability_columns = source_disability_cols

        wb = openpyxl.load_workbook(kt_template_path)
        ws = wb.active
        data_start_row = 9
        for row in range(data_start_row, ws.max_row + 1):
            for col in range(1, 22):
                cell = ws.cell(row=row, column=col)
                if cell.__class__.__name__ != 'MergedCell':
                    cell.value = None

        disability_labels = [
            'vận động', 'nghe nói', 'nhìn', 'thần kinh/tâm thần',
            'trí tuệ', 'học tập', 'tự kỷ', 'khác', 'có chứng nhận', 'khả năng học tập'
        ]

        def put(row_index, column_index, value):
            cell = ws.cell(row=row_index, column=column_index)
            if cell.__class__.__name__ != 'MergedCell':
                cell.value = value

        for index, record in disabled.iterrows():
            row = data_start_row + index
            full_name = f"{str(record[1]).strip() if pd.notna(record[1]) else ''} {str(record[2]).strip() if pd.notna(record[2]) else ''}".strip()
            dob = '/'.join(str(record[col]).strip() for col in (3, 4, 5) if pd.notna(record[col]))
            put(row, 1, index + 1)
            put(row, 2, full_name)
            put(row, 3, dob)
            put(row, 4, 'x' if str(record[6]).strip().lower() in {'x', 'q'} else None)
            put(row, 5, record[7] if pd.notna(record[7]) else None)
            put(row, 6, record[44] if pd.notna(record[44]) else None)
            put(row, 7, record[12] if pd.notna(record[12]) else None)

            grade = str(record[16]).strip() if pd.notna(record[16]) else ''
            grade_column = {'5': 8, '6': 9, '7': 10, '8': 11, '9': 12}.get(grade)
            if grade_column:
                put(row, grade_column, record[17] if pd.notna(record[17]) else 'x')

            put(row, 13, record[18] if pd.notna(record[18]) else None)
            graduation_level = str(record[19]).strip().upper()
            graduation_year = record[21] if pd.notna(record[21]) else None
            put(row, 14, graduation_year if graduation_level == 'TH' else None)
            put(row, 15, graduation_year if graduation_level in {'THCS', 'THPT'} else None)
            put(row, 16, record[24] if pd.notna(record[24]) else None)
            put(row, 17, record[25] if pd.notna(record[25]) else None)
            put(row, 18, record[26] if pd.notna(record[26]) else None)
            put(row, 19, record[27] if pd.notna(record[27]) else None)

            notes = [label for label, col in zip(disability_labels, disability_columns) if pd.notna(record[col])]
            circumstance = str(record[41]).strip() if pd.notna(record[41]) else ''
            circumstance_detail = str(record[42]).strip() if pd.notna(record[42]) else ''
            note_text = ', '.join(notes)
            if circumstance:
                note_text = f'{note_text}; {circumstance}' if note_text else circumstance
            if circumstance_detail:
                note_text = f'{note_text}; {circumstance_detail}' if note_text else circumstance_detail
            put(row, 20, note_text or None)
            put(row, 21, record[13] if pd.notna(record[13]) else None)

        out_path = os.path.join(os.path.dirname(tong_file), 'KetQua_KT.xlsx')
        wb.save(out_path)
        messagebox.showinfo('Thành công', f'Đã xuất danh sách khuyết tật tại:\n{out_path}')
    except Exception as e:
        messagebox.showerror('Lỗi xuất danh sách khuyết tật', f'Chi tiết lỗi: {e}')


def aggregate_spc1(tong_file, spc1_template_path):
    """Tổng hợp danh sách thanh thiếu niên 2008-2015 vào file SPC1."""
    try:
        tong_dir = os.path.dirname(tong_file)
        out_path = os.path.join(tong_dir, "KetQua_SPC1.xlsx")

        # Tự động xoá file kết quả SPC1 cũ trước khi tổng hợp mới
        if os.path.exists(out_path):
            try:
                os.remove(out_path)
            except PermissionError:
                messagebox.showerror(
                    "Lỗi file đang mở",
                    f"File '{os.path.basename(out_path)}' đang được mở trong Excel hoặc phần mềm khác.\n"
                    "Vui lòng đóng file đó lại rồi bấm 'Tổng hợp SPC1'!"
                )
                return
            except Exception as e:
                messagebox.showerror("Lỗi xoá file cũ", f"Không thể xoá file cũ '{os.path.basename(out_path)}':\n{e}")
                return

        if not os.path.exists(spc1_template_path):
            messagebox.showerror("Lỗi", f"Không tìm thấy file mẫu SPC1.xlsx")
            return
            
        # Read Tong.xlsx
        df_tong = pd.read_excel(tong_file, header=None, skiprows=4)
        df_tong[5] = pd.to_numeric(df_tong[5], errors='coerce')
        
        # Filter: sinh năm 2008 - 2015
        df_target = df_tong[(df_tong[5] >= 2008) & (df_tong[5] <= 2015)].reset_index(drop=True)
        
        # Open template
        wb = openpyxl.load_workbook(spc1_template_path)
        sheet = wb.active
        
        # Data starts at row 11 (after headers row 10)
        start_row = 11
        
        for i, row_data in df_target.iterrows():
            r = start_row + i
            
            # Col 1: STT
            sheet.cell(row=r, column=1).value = i + 1
            # Col 2: Họ và tên (index 2 in Tong)
            # Col 2: Họ và tên (họ tại index 1, tên tại index 2)
            surname = str(row_data[1]).strip() if pd.notna(row_data[1]) else ""
            given_name = str(row_data[2]).strip() if pd.notna(row_data[2]) else ""
            full_name = f"{surname} {given_name}".strip()
            sheet.cell(row=r, column=2).value = full_name
            # Col 3: Ngày tháng năm sinh (index 4 = năm sinh, index 3 = ngày tháng)
            # Col 3: Ngày tháng năm sinh (ngày ở index 3, tháng ở index 4, năm ở index 5)
            day = str(row_data[3]).strip() if pd.notna(row_data[3]) else ""
            month = str(row_data[4]).strip() if pd.notna(row_data[4]) else ""
            year = str(row_data[5]).strip() if pd.notna(row_data[5]) else ""
            dob_parts = [p for p in (day, month, year) if p]
            dob = "/".join(dob_parts)
            sheet.cell(row=r, column=3).value = dob
            # Col 4: Giới tính (Nữ đánh x, Nam để trống)
            sheet.cell(row=r, column=4).value = 'x' if str(row_data[6]).strip().lower() in ['x', 'q'] else None
            # Col 5: Dân tộc (index 7)
            sheet.cell(row=r, column=5).value = row_data[7]
            # Col 6: Họ tên cha/mẹ (dữ liệu cột 45 trong Tong.xlsx)
            sheet.cell(row=r, column=6).value = row_data[44] if pd.notna(row_data[44]) else ""
            # Col 7: Chỗ ở (thôn - index 12)
            sheet.cell(row=r, column=7).value = row_data[12]
            
            # Col 8-12: Lớp 5, 6, 7, 8, 9 (đánh dấu x nếu đang học)
            khoi = str(row_data[16]).strip() if pd.notna(row_data[16]) else ""
            if khoi == '5':
                sheet.cell(row=r, column=8).value = row_data[17] if pd.notna(row_data[17]) else 'x'
            elif khoi == '6':
                sheet.cell(row=r, column=9).value = row_data[17] if pd.notna(row_data[17]) else 'x'
            elif khoi == '7':
                sheet.cell(row=r, column=10).value = row_data[17] if pd.notna(row_data[17]) else 'x'
            elif khoi == '8':
                sheet.cell(row=r, column=11).value = row_data[17] if pd.notna(row_data[17]) else 'x'
            elif khoi == '9':
                sheet.cell(row=r, column=12).value = row_data[17] if pd.notna(row_data[17]) else 'x'
            
            # Col 13: Trường (index 18)
            sheet.cell(row=r, column=13).value = row_data[18] if pd.notna(row_data[18]) else ""
            
            # Helper to format clean number or string
            def clean_val(val):
                if pd.isna(val):
                    return ""
                if isinstance(val, float) and val.is_integer():
                    return int(val)
                s = str(val).strip()
                return s

            # Col 14: Năm hoàn thành CTTH / TN Tiểu học (Bậc TN index 19 + năm index 21)
            bac_tn = str(row_data[19]).strip().upper() if pd.notna(row_data[19]) else ""
            nam_tn = clean_val(row_data[21])
            if 'TH' in bac_tn and 'THCS' not in bac_tn:
                sheet.cell(row=r, column=14).value = nam_tn
            
            # Col 15: Năm TN THCS
            if 'THCS' in bac_tn:
                sheet.cell(row=r, column=15).value = nam_tn
            
            # Col 16-17: Đã học xong - Lớp, Năm (cột 25, 26 trong Tong.xlsx -> index 24, 25)
            hx_lop = clean_val(row_data[24])
            hx_nam = clean_val(row_data[25])
            if hx_lop:
                sheet.cell(row=r, column=16).value = hx_lop
            if hx_nam:
                sheet.cell(row=r, column=17).value = hx_nam
            
            # Col 18-19: Bỏ học - Lớp, Năm (cột 27, 28 trong Tong.xlsx -> index 26, 27)
            bh_lop = clean_val(row_data[26])
            bh_nam = clean_val(row_data[27])
            if bh_lop:
                sheet.cell(row=r, column=18).value = bh_lop
            if bh_nam:
                sheet.cell(row=r, column=19).value = bh_nam
            
            # Col 21: Số phiếu tham chiếu
            sheet.cell(row=r, column=21).value = clean_val(row_data[13])
        
        out_path = os.path.join(tong_dir, "KetQua_SPC1.xlsx")
        wb.save(out_path)
        messagebox.showinfo("Thành công", f"Đã xuất {len(df_target)} học sinh sinh 2008-2015 vào:\n{out_path}")
        
    except Exception as e:
        messagebox.showerror("Lỗi tổng hợp SPC1", f"Chi tiết lỗi: {e}")


def aggregate_spc(tong_file, spc_template_path):
    """Tổng hợp sổ phổ cập (SPC) theo mẫu SPC.xlsx - điền từng sheet năm sinh và sheet Tổng."""
    try:
        tong_dir = os.path.dirname(tong_file)
        out_path = os.path.join(tong_dir, "KetQua_SPC.xlsx")

        # Tự động xoá file kết quả SPC cũ trước khi tổng hợp mới
        if os.path.exists(out_path):
            try:
                os.remove(out_path)
            except PermissionError:
                messagebox.showerror(
                    "Lỗi file đang mở",
                    f"File '{os.path.basename(out_path)}' đang được mở trong Excel hoặc phần mềm khác.\n"
                    "Vui lòng đóng file đó lại rồi thử lại!"
                )
                return
            except Exception as e:
                messagebox.showerror("Lỗi xoá file cũ", f"Không thể xoá file cũ '{os.path.basename(out_path)}':\n{e}")
                return

        if not os.path.exists(spc_template_path):
            messagebox.showerror("Lỗi", "Không tìm thấy file mẫu SPC.xlsx")
            return

        df_tong = pd.read_excel(tong_file, header=None, skiprows=4)
        df_tong[5] = pd.to_numeric(df_tong[5], errors='coerce')

        wb = openpyxl.load_workbook(spc_template_path)

        def find_data_start(ws):
            """Tìm dòng đầu tiên có dữ liệu (sau header)."""
            for r in range(1, ws.max_row + 1):
                v = ws.cell(row=r, column=1).value
                if v == '(1)' or str(v).strip() == '(1)':
                    return r + 1
            return 12  # mặc định

        def write_people_to_sheet(ws, df_year, start_row):
            """Ghi danh sách người vào sheet."""
            df_year = df_year.reset_index(drop=True)
            for i, row_data in df_year.iterrows():
                r = start_row + i

                # Col 1: STT
                ws.cell(row=r, column=1).value = i + 1

                # Col 2: Họ và tên
                ho = str(row_data[1]).strip() if pd.notna(row_data[1]) else ''
                ten = str(row_data[2]).strip() if pd.notna(row_data[2]) else ''
                ws.cell(row=r, column=2).value = f"{ho} {ten}".strip()

                # Col 3: Ngày tháng năm sinh
                day = str(row_data[3]).strip() if pd.notna(row_data[3]) else ''
                month = str(row_data[4]).strip() if pd.notna(row_data[4]) else ''
                year = str(row_data[5]).strip() if pd.notna(row_data[5]) else ''
                parts = [p for p in (day, month, year) if p]
                ws.cell(row=r, column=3).value = '/'.join(parts)

                # Col 4: Giới tính
                is_nu = str(row_data[6]).strip() if pd.notna(row_data[6]) else ''
                if is_nu.lower() in ['x', 'q']:
                    ws.cell(row=r, column=4).value = 'Nữ'
                else:
                    ws.cell(row=r, column=4).value = 'Nam'

                # Col 5: Dân tộc
                ws.cell(row=r, column=5).value = row_data[7] if pd.notna(row_data[7]) else ''

                # Col 6: Họ tên cha/mẹ (index 44)
                ws.cell(row=r, column=6).value = row_data[44] if pd.notna(row_data[44]) else ''

                # Col 7: Chỗ ở (thôn - index 12)
                ws.cell(row=r, column=7).value = row_data[12] if pd.notna(row_data[12]) else ''

                # Col 8-12: Lớp đang học (5-9)
                khoi = str(row_data[16]).strip() if pd.notna(row_data[16]) else ''
                lop = str(row_data[17]).strip() if pd.notna(row_data[17]) else ''
                lop_val = lop if lop else khoi
                if khoi == '5':
                    ws.cell(row=r, column=8).value = lop_val
                elif khoi == '6':
                    ws.cell(row=r, column=9).value = lop_val
                elif khoi == '7':
                    ws.cell(row=r, column=10).value = lop_val
                elif khoi == '8':
                    ws.cell(row=r, column=11).value = lop_val
                elif khoi == '9':
                    ws.cell(row=r, column=12).value = lop_val

                # Col 13: Trường (index 18)
                ws.cell(row=r, column=13).value = row_data[18] if pd.notna(row_data[18]) else ''

                # Col 14: Năm TN Tiểu học
                bac_tn = str(row_data[19]).strip().upper() if pd.notna(row_data[19]) else ''
                nam_tn = clean_val(row_data[21])
                if 'TH' in bac_tn and 'THCS' not in bac_tn:
                    ws.cell(row=r, column=14).value = nam_tn

                # Col 15: Năm TN THCS
                if 'THCS' in bac_tn:
                    ws.cell(row=r, column=15).value = nam_tn

                # Col 16-17: Đã học xong - Lớp, Năm (index 24, 25)
                hx_lop = clean_val(row_data[24])
                hx_nam = clean_val(row_data[25])
                if hx_lop:
                    ws.cell(row=r, column=16).value = hx_lop
                if hx_nam:
                    ws.cell(row=r, column=17).value = hx_nam

                # Col 18-19: Bỏ học - Lớp, Năm (index 26, 27)
                bh_lop = clean_val(row_data[26])
                bh_nam = clean_val(row_data[27])
                if bh_lop:
                    ws.cell(row=r, column=18).value = bh_lop
                if bh_nam:
                    ws.cell(row=r, column=19).value = bh_nam

                # Col 21: Số phiếu (index 13)
                ws.cell(row=r, column=21).value = clean_val(row_data[13])

        # Ghi từng sheet năm sinh (2007-2014)
        for year in range(2007, 2015):
            sheet_name = str(year)
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            df_year = df_tong[df_tong[5] == year]
            if df_year.empty:
                continue
            start_row = find_data_start(ws)
            write_people_to_sheet(ws, df_year, start_row)

        # Ghi sheet Tổng (2007-2014)
        if 'Tổng' in wb.sheetnames:
            ws_tong = wb['Tổng']
            df_all = df_tong[(df_tong[5] >= 2007) & (df_tong[5] <= 2014)].sort_values(by=5)
            start_row = find_data_start(ws_tong)
            write_people_to_sheet(ws_tong, df_all, start_row)

        out_path = os.path.join(tong_dir, "KetQua_SPC.xlsx")
        wb.save(out_path)
        total = len(df_tong[(df_tong[5] >= 2007) & (df_tong[5] <= 2014)])
        messagebox.showinfo("Thành công", f"Đã xuất {total} đối tượng sinh 2007-2014 vào:\n{out_path}")

    except Exception as e:
        messagebox.showerror("Lỗi tổng hợp SPC", f"Chi tiết lỗi: {e}")


def aggregate_primary(tong_file, spcth_template_path, m1th_template_path, xmth_template_path):
    """Tổng hợp ba biểu phổ cập giáo dục tiểu học từ file Tổng."""
    try:
        tong_dir = os.path.dirname(tong_file)
        df_tong = pd.read_excel(tong_file, header=None, skiprows=4)
        df_tong[5] = pd.to_numeric(df_tong[5], errors='coerce')

        def clean(value):
            if pd.isna(value):
                return ""
            if isinstance(value, float) and value.is_integer():
                return int(value)
            return str(value).strip()

        def text(value):
            return clean(value).lower()

        def is_female(row):
            return text(row[6]) in {'x', 'q', 'nữ'}

        def is_ethnic(row):
            return text(row[7]) not in {'', 'kinh', 'nan', 'none'}

        def grade(row):
            match = re.match(r'\s*(\d+)', clean(row[16]))
            return int(match.group(1)) if match else None

        def scope_for_sheet(sheet_name, include_all=False):
            normalized_sheet_name = sheet_name.strip().lower()
            if (
                include_all
                or normalized_sheet_name == 'tổng'
                or normalized_sheet_name == 'xoá mù'
                or normalized_sheet_name.startswith('thống kê trẻ')
            ):
                return df_tong
            name = re.sub(r'\s+', ' ', sheet_name).strip().lower()
            name = re.sub(r'\s+pcgd$', '', name)
            scoped = df_tong[df_tong[12].map(lambda value: name in text(value))]
            return scoped if not scoped.empty else df_tong.iloc[0:0]

        def set_value(ws, row, column, value):
            ws.cell(row=row, column=column).value = value if value else None

        def write_spcth():
            wb = openpyxl.load_workbook(spcth_template_path)
            ws = wb.active

            def put(row, column, value):
                cell = ws.cell(row=row, column=column)
                if cell.__class__.__name__ != 'MergedCell':
                    cell.value = value

            for row in range(11, ws.max_row + 1):
                for column in range(1, 18):
                    cell = ws.cell(row=row, column=column)
                    if cell.__class__.__name__ != 'MergedCell':
                        cell.value = None
            rows = df_tong[df_tong[5].between(2011, 2020, inclusive='both')].reset_index(drop=True)
            for index, row_data in rows.iterrows():
                row = 11 + index
                put(row, 1, index + 1)
                put(row, 2, f'{clean(row_data[1])} {clean(row_data[2])}'.strip())
                dob = [clean(row_data[column]) for column in (3, 4, 5) if clean(row_data[column])]
                put(row, 3, '/'.join(map(str, dob)))
                put(row, 4, 'x' if is_female(row_data) else None)
                put(row, 5, clean(row_data[7]))
                put(row, 6, clean(row_data[44]))
                put(row, 7, clean(row_data[12]))
                current_grade = grade(row_data)
                if current_grade in range(1, 6):
                    put(row, 7 + current_grade, clean(row_data[17]) or 'x')
                put(row, 13, clean(row_data[18]))
                if AppGraduationLevel(row_data[19]) == 'TH':
                    put(row, 14, clean(row_data[21]))
                put(row, 15, clean(row_data[26]))
                put(row, 16, clean(row_data[27]))
                put(row, 17, clean(row_data[13]))
            wb.save(os.path.join(tong_dir, 'KetQua_SPCTH.xlsx'))

        def write_m1th():
            wb = openpyxl.load_workbook(m1th_template_path)
            year_columns = {int(wb[wb.sheetnames[0]].cell(4, column).value): column
                            for column in range(1, 51)
                            if isinstance(wb[wb.sheetnames[0]].cell(4, column).value, int)
                            and wb[wb.sheetnames[0]].cell(4, column).value >= 1900}
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                scoped = scope_for_sheet(sheet_name)
                for year, column in year_columns.items():
                    rows = scoped[scoped[5] == year]
                    set_value(ws, 6, column, len(rows))
                    set_value(ws, 7, column, sum(is_female(row) for _, row in rows.iterrows()))
                    set_value(ws, 8, column, sum(is_ethnic(row) for _, row in rows.iterrows()))
                    disabled = rows[rows[[31, 32, 33, 34, 35, 36, 37, 38, 39, 40]].notna().any(axis=1)]
                    set_value(ws, 9, column, len(disabled))
                    set_value(ws, 10, column, int(disabled[40].notna().sum()))
                    set_value(ws, 11, column, int(disabled[39].notna().sum()))
                    set_value(ws, 12, column, len(rows))
                    for class_grade, first_row in zip(range(1, 6), range(13, 26, 3)):
                        class_rows = rows[rows.apply(lambda row: grade(row) == class_grade, axis=1)]
                        set_value(ws, first_row, column, len(class_rows))
                        set_value(ws, first_row + 1, column, len(class_rows))
                        set_value(ws, first_row + 2, column, 0)
                    completed = rows[rows[19].map(AppGraduationLevel) == 'TH']
                    set_value(ws, 28, column, len(completed))
                    set_value(ws, 29, column, len(completed))
                    set_value(ws, 30, column, 0)
                    set_value(ws, 31, column, 0)
                    dropped = rows[rows[26].notna() & (rows[26].astype(str).str.strip() != '')]
                    set_value(ws, 35, column, len(dropped))
                    set_value(ws, 36, column, len(dropped))
                    set_value(ws, 37, column, 0)
                for column in (11, 16):
                    for row in range(6, 38):
                        if row not in (9, 10, 11):
                            ws.cell(row=row, column=column).value = f'=SUM(F{row}:J{row})' if column == 11 else f'=SUM(L{row}:O{row})'

                six_years = scoped[scoped[5] == 2019]
                eleven_years = scoped[scoped[5] == 2014]
                eleven_to_fourteen = scoped[scoped[5].between(2011, 2014, inclusive='both')]
                primary_graduates = scoped[scoped[19].map(AppGraduationLevel) == 'TH']
                disabled_with_learning = scoped[
                    scoped[[31, 32, 33, 34, 35, 36, 37, 38, 39, 40]].notna().any(axis=1)
                    & scoped[40].notna()
                    & (scoped[40].astype(str).str.strip() != '')
                ]

                criteria = {
                    40: (len(six_years[six_years.apply(lambda row: grade(row) == 1, axis=1)]), len(six_years)),
                    41: (len(eleven_years[eleven_years[19].map(AppGraduationLevel) == 'TH']), len(eleven_years)),
                    42: (len(eleven_years[eleven_years.apply(lambda row: grade(row) in range(1, 6), axis=1)]), len(eleven_years)),
                    43: (len(eleven_to_fourteen[eleven_to_fourteen[19].map(AppGraduationLevel) == 'TH']), len(eleven_to_fourteen)),
                    44: (len(disabled_with_learning), len(scoped[scoped[[31, 32, 33, 34, 35, 36, 37, 38, 39, 40]].notna().any(axis=1)])),
                }
                for row, (count, total) in criteria.items():
                    ws.cell(row=row, column=6).value = count
                    ws.cell(row=row, column=7).value = round(count / total * 100, 1) if total else 0
            wb.save(os.path.join(tong_dir, 'KetQua_M1TH.xlsx'))

        def write_xmth():
            wb = openpyxl.load_workbook(xmth_template_path)
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                scoped = scope_for_sheet(sheet_name)
                for row in range(7, 55):
                    year = ws.cell(row=row, column=2).value
                    if not isinstance(year, int):
                        continue
                    rows = scoped[scoped[5] == year]
                    female = rows.apply(is_female, axis=1) if not rows.empty else pd.Series(dtype=bool)
                    ethnic = rows.apply(is_ethnic, axis=1) if not rows.empty else pd.Series(dtype=bool)
                    relapse = rows[30].map(text).isin({'1', 'mức 1', '1.0'}) if not rows.empty else pd.Series(dtype=bool)
                    relapse2 = rows[30].map(text).isin({'2', 'mức 2', '2.0'}) if not rows.empty else pd.Series(dtype=bool)
                    values = {
                        3: len(rows), 4: int(female.sum()), 5: int(ethnic.sum()),
                        6: int((female & ethnic).sum()), 7: int(relapse.sum()),
                        8: int((relapse & female).sum()), 9: int((relapse & ethnic).sum()),
                        10: int((relapse & female & ethnic).sum()), 11: int(relapse2.sum()),
                        12: int((relapse2 & female).sum()), 13: int((relapse2 & ethnic).sum()),
                        14: int((relapse2 & female & ethnic).sum())
                    }
                    literate = ~relapse & ~relapse2
                    values.update({15: int(literate.sum()), 16: int((literate & female).sum()),
                                   17: int((literate & ethnic).sum()), 18: int((literate & female & ethnic).sum())})
                    for column, value in values.items():
                        set_value(ws, row, column, value)
                    ws.cell(row=row, column=19).value = f'=IF(C{row}>0,O{row}/C{row}%,0)'
            wb.save(os.path.join(tong_dir, 'KetQua_XMTH.xlsx'))

        def AppGraduationLevel(value):
            normalized = re.sub(r'[\s()_\-]+', '', clean(value).upper())
            if 'THCS' in normalized or 'TRUNGHỌCCƠSỞ' in normalized:
                return 'THCS'
            if 'THPT' in normalized or 'TRUNGHỌCPHỔTHÔNG' in normalized:
                return 'THPT'
            if normalized == 'TH' or 'TIỂUHỌC' in normalized:
                return 'TH'
            return ''

        write_spcth()
        write_m1th()
        write_xmth()
        messagebox.showinfo(
            'Thành công',
            'Đã tổng hợp 3 biểu Tiểu học:\n'
            f'1. {os.path.join(tong_dir, "KetQua_SPCTH.xlsx")}\n'
            f'2. {os.path.join(tong_dir, "KetQua_M1TH.xlsx")}\n'
            f'3. {os.path.join(tong_dir, "KetQua_XMTH.xlsx")}'
        )
    except Exception as e:
        messagebox.showerror('Lỗi tổng hợp Tiểu học', f'Chi tiết lỗi: {e}')


if __name__ == "__main__":
    pass
