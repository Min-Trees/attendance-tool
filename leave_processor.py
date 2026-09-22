"""
Module Xử Lý Đăng Ký Nghỉ
Đọc file đăng ký nghỉ và trả về số ngày nghỉ theo nhân viên
"""

import pandas as pd
from datetime import datetime


def doc_file_dang_ky_nghi(file_bytes):
    """
    Đọc file đăng ký nghỉ với nhiều engine
    """
    engines = ['calamine', 'openpyxl', 'xlrd']
    for engine in engines:
        try:
            df = pd.read_excel(file_bytes, engine=engine, header=0)
            return df
        except:
            continue
    raise Exception("Không thể đọc file đăng ký nghỉ")


def xu_ly_dang_ky_nghi(df, cot_ma_nv=None, cot_ngay=None):
    """
    Xử lý đăng ký nghỉ, trả về dict {mã_nv: số_ngày_nghỉ}
    
    Args:
        df: DataFrame đã đọc
        cot_ma_nv: tên cột mã nhân viên (auto detect nếu None)
        cot_ngay: tên cột ngày nghỉ (auto detect nếu None)
    
    Returns:
        dict: {mã_nv: số_ngày_nghỉ}
    """
    all_cols = df.columns.tolist()
    
    # Auto detect cột
    if cot_ma_nv is None:
        cot_ma_nv = next((c for c in all_cols if any(k in str(c).lower() for k in ['mã', 'ma', 'staff', 'nv', 'id'])), all_cols[0])
    
    if cot_ngay is None:
        cot_ngay = next((c for c in all_cols if any(k in str(c).lower() for k in ['ngày', 'ngay', 'date', 'nghi'])), all_cols[1])
    
    # Chuyển ngày về date
    df_work = df[[cot_ma_nv, cot_ngay]].copy()
    df_work[cot_ma_nv] = df_work[cot_ma_nv].astype(str).str.strip()
    df_work[cot_ngay] = pd.to_datetime(df_work[cot_ngay], dayfirst=True, errors='coerce').dt.date
    
    # Loại bỏ ngày null
    df_work = df_work.dropna(subset=[cot_ngay])
    
    # Loại bỏ trùng lặp (cùng NV cùng ngày chỉ tính 1)
    df_work = df_work.drop_duplicates(subset=[cot_ma_nv, cot_ngay])
    
    # Đếm số ngày nghỉ theo NV
    so_ngay_nghi = df_work.groupby(cot_ma_nv).size().to_dict()
    
    return so_ngay_nghi, cot_ma_nv, cot_ngay


def get_cot_names(df):
    """
    Lấy danh sách tên cột để auto-detect
    """
    return df.columns.tolist()


def tinh_ngay_cong_chuan(so_ngay_cham_cong, so_ngay_nghi, cong_chuan=26):
    """
    Tính ngày công chuẩn = công chuẩn - số ngày nghỉ đã đăng ký
    
    Args:
        so_ngay_cham_cong: int - số ngày đã chấm công
        so_ngay_nghi: int - số ngày nghỉ đã đăng ký
        cong_chuan: int - công chuẩn tháng
    
    Returns:
        int - số ngày công chuẩn còn lại
    """
    return max(0, cong_chuan - so_ngay_nghi)
