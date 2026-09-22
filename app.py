"""
Tool Xuất Dữ Liệu Chấm Công - Phiên bản BCC
Xuất file BCC (Bảng Chấm Công) theo template Trần Phú
"""

import streamlit as st
import pandas as pd
import io
from datetime import datetime, date
from openpyxl import Workbook
from leave_processor import doc_file_dang_ky_nghi, xu_ly_dang_ky_nghi, get_cot_names
from template_processor import xuat_bcc_theo_template, doc_bcc_template

st.set_page_config(
    page_title="Tool Xuất BCC - Trần Phú",
    layout="wide"
)

# ====== CSS ======
st.markdown("""
<style>
    .main-title {
        font-size: 2rem;
        font-weight: 700;
        color: #1E40AF;
        text-align: center;
        padding: 1.5rem 0 0.5rem 0;
        border-bottom: 3px solid #1E40AF;
        margin-bottom: 1.5rem;
    }
    
    .subtitle {
        text-align: center;
        color: #374151;
        font-size: 0.95rem;
        margin-bottom: 1rem;
    }
    
    .metric-box {
        background: white;
        border: 1px solid #D1D5DB;
        border-left: 4px solid #1E40AF;
        border-radius: 8px;
        padding: 1rem 1.25rem;
        margin: 0.5rem 0;
    }
    
    .metric-box .label {
        color: #374151;
        font-size: 0.85rem;
        font-weight: 500;
        margin: 0;
    }
    
    .metric-box .value {
        color: #111827;
        font-size: 1.75rem;
        font-weight: 700;
        margin: 0.25rem 0 0 0;
    }
    
    .stButton > button {
        background: #1E40AF;
        color: white;
        font-weight: 600;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1.5rem;
    }
    
    .stButton > button:hover {
        background: #1D4ED8;
        color: white;
    }
    
    .stDownloadButton > button {
        background: #059669;
        color: white;
        font-weight: 600;
        border: none;
        border-radius: 8px;
    }
    
    .info-box {
        background: white;
        border: 1px solid #D1D5DB;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
        color: #111827;
    }
    
    .info-box h4 {
        margin: 0 0 0.5rem 0;
        color: #111827;
    }
    
    .info-box p {
        margin: 0;
        color: #374151;
    }
    
    section[data-testid="stSidebar"] h3 {
        color: #111827;
        font-weight: 600;
    }
    
    section[data-testid="stSidebar"] p {
        color: #374151;
    }
    
    .stTextInput label, .stSelectbox label, .stNumberInput label {
        color: #374151 !important;
    }
    
    .file-section {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    
    .file-section h4 {
        margin: 0 0 0.75rem 0;
        color: #1E40AF;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">Tool Xuất BCC - Trần Phú</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Xuất bảng chấm công theo định dạng mẫu Trần Phú 2026</div>', unsafe_allow_html=True)


# ====== HÀM XỬ LÝ ======
def doc_file_excel_cham_cong(file_bytes):
    """Đọc file chấm công với header=2 (như cũ)"""
    engines = ['calamine', 'openpyxl', 'xlrd']
    for engine in engines:
        try:
            df = pd.read_excel(file_bytes, engine=engine, header=2)
            df = df.iloc[1:].reset_index(drop=True)
            return df
        except:
            continue
    raise Exception("Không thể đọc file Excel chấm công")


def build_cham_cong_dict(df_cham_cong, cot_ma_nv, cot_ngay):
    """
    Tạo dict {(ma_nv, date): True} từ file chấm công
    """
    result = {}
    df_work = df_cham_cong[[cot_ma_nv, cot_ngay]].copy()
    df_work = df_work.dropna(subset=[cot_ngay])
    df_work[cot_ngay] = pd.to_datetime(df_work[cot_ngay], dayfirst=True, errors='coerce')
    df_work = df_work.dropna(subset=[cot_ngay])
    df_work = df_work.drop_duplicates(subset=[cot_ma_nv, cot_ngay])
    
    for _, row in df_work.iterrows():
        ma_nv = str(row[cot_ma_nv]).strip()
        ngay = row[cot_ngay].date()
        result[(ma_nv, ngay)] = True
    
    return result


def build_nghi_dict(df_nghi):
    """
    Tạo dict {(ma_nv, date): 'P'} từ file đăng ký nghỉ
    """
    result = {}
    nghi_dict, cot_ma, cot_ngay = xu_ly_dang_ky_nghi(df_nghi)
    
    for _, row in df_nghi.iterrows():
        ma_nv = str(row[cot_ma]).strip()
        ngay_val = row[cot_ngay]
        if pd.isna(ngay_val):
            continue
        ngay = pd.to_datetime(ngay_val, dayfirst=True, errors='coerce')
        if pd.isna(ngay):
            continue
        result[(ma_nv, ngay.date())] = 'P'
    
    return result


def build_ds_nhan_vien(df_cham_cong, cot_ma_nv, cot_ten_nv, cot_chuc_vu=None, cot_ngay_nv=None):
    """
    Tạo danh sách nhân viên từ file chấm công
    """
    df_work = df_cham_cong[[cot_ma_nv, cot_ten_nv]].copy()
    if cot_chuc_vu and cot_chuc_vu in df_cham_cong.columns:
        df_work[cot_chuc_vu] = df_cham_cong[cot_chuc_vu]
    if cot_ngay_nv and cot_ngay_nv in df_cham_cong.columns:
        df_work[cot_ngay_nv] = df_cham_cong[cot_ngay_nv]
    
    df_work = df_work.drop_duplicates(subset=[cot_ma_nv]).reset_index(drop=True)
    
    ds_nv = []
    for _, row in df_work.iterrows():
        nv = {
            'ma_nv': str(row[cot_ma_nv]).strip(),
            'ho_ten': str(row[cot_ten_nv]).strip() if pd.notna(row[cot_ten_nv]) else '',
            'chuc_vu': str(row[cot_chuc_vu]).strip() if cot_chuc_vu and cot_chuc_vu in df_work.columns and pd.notna(row.get(cot_chuc_vu)) else '',
        }
        if cot_ngay_nv and cot_ngay_nv in df_work.columns and pd.notna(row.get(cot_ngay_nv)):
            ngay_nv_val = row[cot_ngay_nv]
            try:
                nv['ngay_nhan_viec'] = pd.to_datetime(ngay_nv_val, dayfirst=True, errors='coerce').date()
            except:
                pass
        ds_nv.append(nv)
    
    return ds_nv


# ====== GIAO DIỆN ======
with st.sidebar:
    st.markdown("### Upload Files")
    
    st.markdown("---")
    st.markdown("#### 1. File Template BCC")
    uploaded_template = st.file_uploader("Template BCC Trần Phú (bắt buộc)", type=['xlsx'], key="template")
    
    st.markdown("---")
    st.markdown("#### 2. File Chấm Công")
    uploaded_cham_cong = st.file_uploader("File chấm công (bắt buộc)", type=['xlsx', 'xls'], key="cham_cong")
    
    st.markdown("---")
    st.markdown("#### 3. File Đăng Ký Nghỉ")
    uploaded_nghi = st.file_uploader("File đăng ký nghỉ (tùy chọn)", type=['xlsx', 'xls'], key="nghi")


if uploaded_template is None or uploaded_cham_cong is None:
    st.markdown("""
    <div class="info-box">
        <h4>Vui lòng upload đầy đủ file để bắt đầu</h4>
        <p>Cần upload:</p>
        <p>- File template BCC Trần Phú (bắt buộc)</p>
        <p>- File chấm công (bắt buộc)</p>
        <p>- File đăng ký nghỉ (tùy chọn)</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

try:
    # Đọc template
    template_info = doc_bcc_template(uploaded_template)
    st.markdown(f'<div class="info-box"><h4>Đã đọc template: {template_info["sheet_name"]} - Tháng {template_info["thang"]}/{template_info["nam"]}</h4><p>Template có sẵn {template_info["last_row"] - 6} dòng nhân viên</p></div>', unsafe_allow_html=True)
    
    # Đọc file chấm công
    df_raw = doc_file_excel_cham_cong(uploaded_cham_cong)
    st.markdown(f'<div class="info-box"><h4>Đã đọc file chấm công: {len(df_raw)} dòng</h4></div>', unsafe_allow_html=True)
    
    # Đọc file đăng ký nghỉ
    nghi_data = {}
    if uploaded_nghi is not None:
        try:
            df_nghi = doc_file_dang_ky_nghi(uploaded_nghi)
            nghi_data = build_nghi_dict(df_nghi)
            st.markdown(f'<div class="info-box"><h4>Đã đọc file đăng ký nghỉ: {len(nghi_data)} ngày nghỉ</h4></div>', unsafe_allow_html=True)
        except Exception as e:
            st.warning(f"Không thể đọc file đăng ký nghỉ: {str(e)}")
    
    # Cấu hình cột
    st.markdown("### Cấu hình cột dữ liệu")
    all_cols = df_raw.columns.tolist()
    
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        ma_nv_def = next((c for c in all_cols if any(k in str(c).lower() for k in ['mã', 'ma', 'staff'])), all_cols[1])
        cot_ma_nv = st.selectbox("Cột Mã NV:", all_cols, index=all_cols.index(ma_nv_def))
    with col2:
        ten_nv_def = next((c for c in all_cols if any(k in str(c).lower() for k in ['tên', 'ten', 'name'])), all_cols[2])
        cot_ten_nv = st.selectbox("Cột Họ tên:", all_cols, index=all_cols.index(ten_nv_def))
    with col3:
        ngay_def = next((c for c in all_cols if any(k in str(c).lower() for k in ['ngày', 'ngay', 'date'])), all_cols[5])
        cot_ngay = st.selectbox("Cột Ngày chấm công:", all_cols, index=all_cols.index(ngay_def))
    with col4:
        chuc_vu_options = ['(Không có)'] + all_cols
        chuc_vu_def = next((c for c in all_cols if any(k in str(c).lower() for k in ['chức', 'chuc', 'vị trí', 'position'])), '(Không có)')
        chuc_vu_idx = chuc_vu_options.index(chuc_vu_def) if chuc_vu_def in chuc_vu_options else 0
        cot_chuc_vu = st.selectbox("Cột Chức vụ:", chuc_vu_options, index=chuc_vu_idx)
        cot_chuc_vu = None if cot_chuc_vu == '(Không có)' else cot_chuc_vu
    with col5:
        ngay_nv_options = ['(Không có)'] + all_cols
        ngay_nv_def = next((c for c in all_cols if any(k in str(c).lower() for k in ['nhận việc', 'nhận', 'start', 'vào làm'])), '(Không có)')
        ngay_nv_idx = ngay_nv_options.index(ngay_nv_def) if ngay_nv_def in ngay_nv_options else 0
        cot_ngay_nv = st.selectbox("Cột Ngày nhận việc:", ngay_nv_options, index=ngay_nv_idx)
        cot_ngay_nv = None if cot_ngay_nv == '(Không có)' else cot_ngay_nv
    
    col1, col2 = st.columns(2)
    with col1:
        thang = st.number_input("Tháng:", min_value=1, max_value=12, value=template_info['thang'])
    with col2:
        nam = st.number_input("Năm:", min_value=2020, max_value=2099, value=template_info['nam'])
    
    st.markdown("---")
    
    if st.button("Xử lý và xuất BCC", use_container_width=True):
        with st.spinner("Đang xử lý..."):
            # Build dữ liệu
            cham_cong_data = build_cham_cong_dict(df_raw, cot_ma_nv, cot_ngay)
            ds_nv = build_ds_nhan_vien(df_raw, cot_ma_nv, cot_ten_nv, cot_chuc_vu, cot_ngay_nv)
            
            st.session_state['cham_cong_data'] = cham_cong_data
            st.session_state['nghi_data'] = nghi_data
            st.session_state['ds_nv'] = ds_nv
    
    if 'ds_nv' in st.session_state:
        ds_nv = st.session_state['ds_nv']
        cham_cong_data = st.session_state['cham_cong_data']
        nghi_data = st.session_state['nghi_data']
        
        st.markdown("### Thống kê")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f'<div class="metric-box"><p class="label">Tổng nhân viên</p><p class="value">{len(ds_nv)}</p></div>', unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="metric-box"><p class="label">Tổng ngày công</p><p class="value">{len(cham_cong_data):,}</p></div>', unsafe_allow_html=True)
        with col3:
            st.markdown(f'<div class="metric-box"><p class="label">Tổng ngày nghỉ</p><p class="value">{len(nghi_data):,}</p></div>', unsafe_allow_html=True)
        with col4:
            st.markdown(f'<div class="metric-box"><p class="label">Tháng/Năm</p><p class="value">{thang}/{nam}</p></div>', unsafe_allow_html=True)
        
        st.markdown("### Danh sách nhân viên")
        df_preview = pd.DataFrame(ds_nv)
        st.dataframe(df_preview, use_container_width=True, hide_index=True, height=300)
        
        st.markdown("### Tải xuống")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Tạo file BCC", use_container_width=True):
                with st.spinner("Đang tạo file BCC..."):
                    excel_buf = xuat_bcc_theo_template(
                        template_path=uploaded_template,
                        ds_nhan_vien=ds_nv,
                        cham_cong_data=cham_cong_data,
                        nghi_data=nghi_data,
                        thang=int(thang),
                        nam=int(nam),
                        cong_chuan=26
                    )
                    st.session_state['excel_buf'] = excel_buf
        
        if 'excel_buf' in st.session_state:
            st.download_button(
                "Tải xuống BCC (.xlsx)",
                data=st.session_state['excel_buf'],
                file_name=f"BCC_Thang_{int(thang):02d}_{int(nam)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

except Exception as e:
    st.error(f"Lỗi: {str(e)}")
    import traceback
    st.code(traceback.format_exc())
