"""
Module Xử Lý Template BCC (Bảng Chấm Công)
Đọc file template BCC Trần Phú và điền dữ liệu mới
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy
from datetime import datetime, date
import calendar
import io


def lay_thong_tin_thang_nam(ws):
    """
    Lấy tháng và năm từ template (ô V2 = THÁNG, X2 = NĂM)
    """
    thang_cell = ws.cell(row=2, column=22).value  # V2 = THÁNG
    nam_cell = ws.cell(row=2, column=24).value    # X2 = NĂM
    
    thang = int(thang_cell) if thang_cell else datetime.now().month
    nam = int(nam_cell) if nam_cell else datetime.now().year
    
    return thang, nam


def lay_danh_sach_ngay_trong_thang(ws, thang, nam):
    """
    Lấy danh sách các ngày trong tháng từ dòng 5 của template
    Trả về list các ngày (datetime.date)
    """
    ngay_list = []
    for col in range(7, 38):  # G (7) đến AK (37) = 31 cột
        cell = ws.cell(row=5, column=col)
        val = cell.value
        if val and val != 'nan':
            if isinstance(val, datetime):
                ngay_list.append(val.date())
            elif isinstance(val, date):
                ngay_list.append(val)
    return ngay_list


def dem_so_dong_du_lieu(ws, col_ten=3):
    """
    Đếm số dòng có dữ liệu nhân viên trong template
    """
    last_row = 6
    for row in range(7, 100):
        val = ws.cell(row=row, column=col_ten).value
        if val and str(val).strip() and val != 'nan':
            last_row = row
    return last_row


def xoa_du_lieu_cu(ws, start_row, end_row):
    """
    Xóa giá trị các ô từ start_row đến end_row (giữ nguyên format)
    """
    for row in range(start_row, end_row + 1):
        for col in range(1, 54):  # Clear tất cả các cột
            cell = ws.cell(row=row, column=col)
            cell.value = None


def copy_row_format(ws, source_row, target_row):
    """
    Copy format từ dòng nguồn sang dòng đích
    """
    for col in range(1, 54):
        source_cell = ws.cell(row=source_row, column=col)
        target_cell = ws.cell(row=target_row, column=col)
        
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.fill = copy(source_cell.fill)
            target_cell.border = copy(source_cell.border)
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.number_format = source_cell.number_format
            target_cell.protection = copy(source_cell.protection)


def them_dong_moi(ws, template_row, insert_position, count=1):
    """
    Chèn thêm dòng mới sau insert_position, copy format từ template_row
    Trả về số dòng đã chèn
    """
    ws.insert_rows(insert_position + 1, amount=count)
    for i in range(count):
        target_row = insert_position + 1 + i
        copy_row_format(ws, template_row, target_row)
    return count


def dien_du_lieu_nhan_vien(ws, row, stt, ma_nv, ho_ten, cham_cong_dict, nghi_dict, ngay_list):
    """
    Điền dữ liệu cho 1 nhân viên tại dòng row
    
    Args:
        ws: worksheet
        row: số dòng
        stt: số thứ tự
        ma_nv: mã nhân viên
        ho_ten: họ và tên
        cham_cong_dict: dict {date: True/False} - có chấm công ngày nào
        nghi_dict: dict {date: 'P' or 'P/2'} - nghỉ ngày nào
        ngay_list: list các ngày trong tháng
    """
    # Cột A: STT
    ws.cell(row=row, column=1, value=stt)
    # Cột B: Mã NV
    ws.cell(row=row, column=2, value=ma_nv)
    # Cột C: Họ tên
    ws.cell(row=row, column=3, value=ho_ten)
    
    # Cột G-AK (7-37): 31 cột ngày
    for col_idx, ngay in enumerate(ngay_list):
        target_col = 7 + col_idx
        ngay_obj = ngay if isinstance(ngay, date) else ngay.date()
        
        if ngay_obj in nghi_dict:
            ws.cell(row=row, column=target_col, value=nghi_dict[ngay_obj])
        elif cham_cong_dict.get(ngay_obj, False):
            ws.cell(row=row, column=target_col, value='8')


def tinh_tong_cac_cot(ws, row, cong_chuan=26):
    """
    Tính các cột tổng hợp cho nhân viên
    """
    # Cột AL (38): Ngày công thực tế = SUM(G:AK)/8 + COUNTIF(P/2)*0.5
    ws.cell(row=row, column=38, value=f'=SUM(G{row}:AK{row})/8+COUNTIF(G{row}:AK{row},"P/2")*0.5')
    
    # Cột AM (39): Ngày phép, nghỉ HL
    ws.cell(row=row, column=39, value=f'=COUNTIF(G{row}:AK{row},"P")+COUNTIF(G{row}:AK{row},"P/2")*0.5')
    
    # Cột AN (40): Ngày nghỉ Lễ/Tết
    ws.cell(row=row, column=40, value=f'=COUNTIF(G{row}:AK{row},"HL")')
    
    # Cột AO (41): Ngày nghỉ Ốm/Thai sản
    ws.cell(row=row, column=41, value=0)
    
    # Cột AP (42): Nghỉ không lương
    ws.cell(row=row, column=42, value=0)
    
    # Cột AQ (43): Tổng ngày nghỉ trong tháng
    ws.cell(row=row, column=43, value=f'=AM{row}+AN{row}+AO{row}+AP{row}')
    
    # Cột AS (45): Ngày phép tồn
    ws.cell(row=row, column=45, value=0)
    
    # Cột AT (46): Ngày phép tồn cuối tháng
    ws.cell(row=row, column=46, value=f'=AS{row}-AM{row}')
    
    # Cột AU (47): Trừ BH - giữ nguyên nếu có, hoặc set = 0
    # Cột AV (48): Phí giữ ngoài giờ
    # Cột AW (49): Tiền ăn trưa
    # Cột AX (50): Ứng lương
    # Cột AY (51): Sinh nhật
    # Cột AZ (52): Aerobic
    # Cột BA (53): Ghi chú


def xuat_bcc_theo_template(template_path, ds_nhan_vien, cham_cong_data, nghi_data, 
                            thang=None, nam=None, cong_chuan=26):
    """
    Xuất file BCC theo template
    
    Args:
        template_path: đường dẫn file template
        ds_nhan_vien: list of dict [{'ma_nv': ..., 'ho_ten': ...}, ...]
        cham_cong_data: dict {(ma_nv, date): True}
        nghi_data: dict {(ma_nv, date): 'P'/'P/2'}
        thang: tháng (None = tự lấy từ template)
        nam: năm (None = tự lấy từ template)
        cong_chuan: công chuẩn (mặc định 26)
    
    Returns:
        BytesIO buffer
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    # Lấy tháng/năm
    thang_template, nam_template = lay_thong_tin_thang_nam(ws)
    if thang is None:
        thang = thang_template
    if nam is None:
        nam = nam_template
    
    # Cập nhật tháng/năm trong header
    ws.cell(row=2, column=22, value=thang)
    ws.cell(row=2, column=24, value=nam)
    
    # Lấy danh sách ngày trong tháng
    ngay_list = lay_danh_sach_ngay_trong_thang(ws, thang, nam)
    
    # Tìm dòng dữ liệu cuối cùng
    last_data_row = dem_so_dong_du_lieu(ws)
    
    # Xóa dữ liệu cũ (giữ format)
    xoa_du_lieu_cu(ws, 7, last_data_row)
    
    # Đếm số NV cần điền
    so_nv_can = len(ds_nhan_vien)
    so_dong_co_sẵn = last_data_row - 6
    
    # Nếu cần thêm dòng, chèn thêm
    if so_nv_can > so_dong_co_sẵn:
        them_dong_moi(ws, 7, 6, so_nv_can - so_dong_co_sẵn)
    
    # Điền dữ liệu cho từng nhân viên
    for idx, nv in enumerate(ds_nhan_vien):
        row = 7 + idx
        ma_nv = str(nv.get('ma_nv', '')).strip()
        ho_ten = nv.get('ho_ten', '').strip()
        chuc_vu = nv.get('chuc_vu', '').strip()
        ngay_nhan_viec = nv.get('ngay_nhan_viec')
        
        # Build dict cho nhân viên này
        cham_cong_dict = {}
        for (nv_id, ngay), _ in cham_cong_data.items():
            if str(nv_id).strip() == ma_nv:
                if isinstance(ngay, datetime):
                    cham_cong_dict[ngay.date()] = True
                elif isinstance(ngay, date):
                    cham_cong_dict[ngay] = True
        
        nghi_dict = {}
        for (nv_id, ngay), loai in nghi_data.items():
            if str(nv_id).strip() == ma_nv:
                if isinstance(ngay, datetime):
                    nghi_dict[ngay.date()] = loai
                elif isinstance(ngay, date):
                    nghi_dict[ngay] = loai
        
        # Điền thông tin cơ bản
        ws.cell(row=row, column=1, value=idx + 1)  # STT
        ws.cell(row=row, column=2, value=ma_nv)
        ws.cell(row=row, column=3, value=ho_ten)
        ws.cell(row=row, column=4, value=chuc_vu)
        if ngay_nhan_viec:
            ws.cell(row=row, column=5, value=ngay_nhan_viec)
        
        # Điền các cột ngày
        for col_idx, ngay in enumerate(ngay_list):
            target_col = 7 + col_idx
            ngay_obj = ngay if isinstance(ngay, date) else ngay.date()
            
            if ngay_obj in nghi_dict:
                ws.cell(row=row, column=target_col, value=nghi_dict[ngay_obj])
            elif cham_cong_dict.get(ngay_obj, False):
                ws.cell(row=row, column=target_col, value='8')
        
        # Tính các cột tổng hợp
        tinh_tong_cac_cot(ws, row, cong_chuan)
    
    # Lưu vào buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def doc_bcc_template(template_path):
    """
    Đọc thông tin từ template BCC
    Trả về dict: {'thang': X, 'nam': Y, 'ngay_list': [...], 'last_row': N}
    """
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    thang, nam = lay_thong_tin_thang_nam(ws)
    ngay_list = lay_danh_sach_ngay_trong_thang(ws, thang, nam)
    last_row = dem_so_dong_du_lieu(ws)
    
    return {
        'thang': thang,
        'nam': nam,
        'ngay_list': ngay_list,
        'last_row': last_row,
        'sheet_name': ws.title
    }
