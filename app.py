"""
Tool Xuất Dữ Liệu Chấm Công
Xuất: Mã nhân viên, Họ và tên, Số ngày chấm công, Số ngày nghỉ, Công chuẩn
"""

import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from leave_processor import doc_file_dang_ky_nghi, xu_ly_dang_ky_nghi, get_cot_names, tinh_ngay_cong_chuan

st.set_page_config(
    page_title="Tool Xuất Chấm Công",
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

st.markdown('<div class="main-title">Tool Xuất Dữ Liệu Chấm Công</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Xuất ra 5 cột: Mã nhân viên | Họ và tên | Số ngày chấm công | Số ngày nghỉ | Công chuẩn</div>', unsafe_allow_html=True)


# ====== THUẬT TOÁN ======
def xu_ly_du_lieu_cham_cong(df, cot_ma_nv, cot_ten_nv, cot_ngay):
    df_work = df[[cot_ma_nv, cot_ten_nv, cot_ngay]].copy()
    df_work = df_work.drop_duplicates(subset=[cot_ma_nv, cot_ngay])
    df_work[cot_ngay] = pd.to_datetime(df_work[cot_ngay], dayfirst=True, errors='coerce').dt.date
    
    df_ket_qua = df_work.groupby([cot_ma_nv, cot_ten_nv], as_index=False).agg(
        so_ngay_cham_cong=(cot_ngay, 'nunique')
    )
    
    df_ket_qua.columns = ['Ma_nv', 'Ho_ten', 'So_ngay_cham_cong']
    
    df_ket_qua['_sort'] = pd.to_numeric(df_ket_qua['Ma_nv'], errors='coerce')
    df_ket_qua = df_ket_qua.sort_values('_sort', na_position='first').drop('_sort', axis=1).reset_index(drop=True)
    
    return df_ket_qua


def doc_file_excel(file_bytes):
    engines = ['calamine', 'openpyxl', 'xlrd']
    for engine in engines:
        try:
            df = pd.read_excel(file_bytes, engine=engine, header=2)
            df = df.iloc[1:].reset_index(drop=True)
            return df
        except:
            continue
    raise Exception("Không thể đọc file Excel")


def xuat_excel_dep(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Chấm Công"
    
    ws.merge_cells('A1:E1')
    cell = ws['A1']
    cell.value = "BÁO CÁO CHẤM CÔNG"
    cell.font = Font(name='Arial', size=16, bold=True, color='FFFFFF')
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.fill = PatternFill(start_color='1E40AF', end_color='1E40AF', fill_type='solid')
    ws.row_dimensions[1].height = 30
    
    headers = ['Mã nhân viên', 'Họ và tên', 'Số ngày chấm công', 'Số ngày nghỉ', 'Công chuẩn']
    border = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )
    
    for col_idx, header in enumerate(headers, start=1):
        c = ws.cell(row=2, column=col_idx, value=header)
        c.font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
        c.fill = PatternFill(start_color='1E40AF', end_color='1E40AF', fill_type='solid')
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = border
    ws.row_dimensions[2].height = 25
    
    for row_idx, row in enumerate(df.itertuples(index=False), start=3):
        ws.cell(row=row_idx, column=1, value=int(row[0]) if str(row[0]).isdigit() else row[0])
        ws.cell(row=row_idx, column=2, value=row[1])
        ws.cell(row=row_idx, column=3, value=int(row[2]))
        ws.cell(row=row_idx, column=4, value=int(row[3]) if pd.notna(row[3]) else 0)
        ws.cell(row=row_idx, column=5, value=int(row[4]))
        
        if row_idx % 2 == 1:
            fill = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')
            for col_idx in range(1, 6):
                ws.cell(row=row_idx, column=col_idx).fill = fill
        
        for col_idx in range(1, 6):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = Font(name='Arial', size=11, color='111827')
            cell.border = border
            if col_idx in [1, 3, 4, 5]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')
    
    last_row = len(df) + 3
    ws.cell(row=last_row, column=1, value='TỔNG CỘNG')
    ws.cell(row=last_row, column=2, value=f"{len(df)} nhân viên")
    ws.cell(row=last_row, column=3, value=int(df['So_ngay_cham_cong'].sum()))
    ws.cell(row=last_row, column=4, value=int(df['So_ngay_nghi'].sum()))
    ws.cell(row=last_row, column=5, value=int(df['Cong_chuan'].sum()))
    
    for col_idx in range(1, 6):
        c = ws.cell(row=last_row, column=col_idx)
        c.font = Font(name='Arial', size=12, bold=True, color='FFFFFF')
        c.fill = PatternFill(start_color='1E40AF', end_color='1E40AF', fill_type='solid')
        c.border = border
        c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[last_row].height = 25
    
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 16
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ====== GIAO DIỆN ======
with st.sidebar:
    st.markdown("### Upload Files")
    
    st.markdown("---")
    st.markdown("#### 1. File Chấm Công")
    uploaded_file_cham_cong = st.file_uploader("Chọn file chấm công", type=['xlsx', 'xls'], key="cham_cong")
    
    st.markdown("---")
    st.markdown("#### 2. File Đăng Ký Nghỉ")
    uploaded_file_nghi = st.file_uploader("Chọn file đăng ký nghỉ (tùy chọn)", type=['xlsx', 'xls'], key="nghi")
    
    st.markdown("---")
    st.markdown("### Kết quả xuất")
    st.markdown("""
    - Mã nhân viên
    - Họ và tên
    - Số ngày chấm công
    - Số ngày nghỉ
    - Công chuẩn
    """)


# Kiểm tra đã upload đủ file chưa
if uploaded_file_cham_cong is None:
    st.markdown("""
    <div class="info-box">
        <h4>Vui lòng upload file để bắt đầu</h4>
        <p>Cần upload:</p>
        <p>- File chấm công (bắt buộc)</p>
        <p>- File đăng ký nghỉ (tùy chọn)</p>
    </div>
    """, unsafe_allow_html=True)
else:
    try:
        df_raw = doc_file_excel(uploaded_file_cham_cong)
        st.markdown('<div class="info-box"><h4>Đã đọc file chấm công thành công</h4></div>', unsafe_allow_html=True)
        
        all_cols = df_raw.columns.tolist()
        ma_nv_def = next((c for c in all_cols if any(k in str(c).lower() for k in ['mã', 'ma', 'staff'])), all_cols[1])
        ten_nv_def = next((c for c in all_cols if any(k in str(c).lower() for k in ['tên', 'ten', 'name'])), all_cols[2])
        ngay_def = next((c for c in all_cols if any(k in str(c).lower() for k in ['ngày', 'ngay', 'date'])), all_cols[5])
        
        # Xử lý file đăng ký nghỉ
        nghi_dict = {}
        cot_ma_nghi = None
        cot_ngay_nghi = None
        
        if uploaded_file_nghi is not None:
            try:
                df_nghi = doc_file_dang_ky_nghi(uploaded_file_nghi)
                cols_nghi = get_cot_names(df_nghi)
                nghi_dict, cot_ma_nghi, cot_ngay_nghi = xu_ly_dang_ky_nghi(df_nghi)
                st.markdown(f'<div class="info-box"><h4>Đã đọc file đăng ký nghỉ thành công</h4><p>Tìm thấy {len(nghi_dict)} nhân viên có đăng ký nghỉ</p></div>', unsafe_allow_html=True)
            except Exception as e:
                st.warning(f"Không thể đọc file đăng ký nghỉ: {str(e)}")
        
        st.markdown("### Cấu hình")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            cot_ma_nv = st.selectbox("Cột Mã NV:", all_cols, index=all_cols.index(ma_nv_def))
        with col2:
            cot_ten_nv = st.selectbox("Cột Họ tên:", all_cols, index=all_cols.index(ten_nv_def))
        with col3:
            cot_ngay = st.selectbox("Cột Ngày:", all_cols, index=all_cols.index(ngay_def))
        with col4:
            cong_chuan = st.number_input("Công chuẩn:", min_value=1, max_value=31, value=26)
        
        st.markdown("---")
        
        if st.button("Xử lý dữ liệu", use_container_width=True):
            with st.spinner("Đang xử lý..."):
                df_kq = xu_ly_du_lieu_cham_cong(df_raw, cot_ma_nv, cot_ten_nv, cot_ngay)
                
                # Thêm cột số ngày nghỉ
                df_kq['Ma_nv_str'] = df_kq['Ma_nv'].astype(str).str.strip()
                df_kq['So_ngay_nghi'] = df_kq['Ma_nv_str'].map(nghi_dict).fillna(0).astype(int)
                
                # Tính công chuẩn = công chuẩn - số ngày nghỉ
                df_kq['Cong_chuan'] = df_kq['So_ngay_nghi'].apply(lambda x: max(0, cong_chuan - x))
                
                # Đổi tên cột cuối
                df_kq = df_kq.rename(columns={'Cong_chuan': 'Công chuẩn', 'So_ngay_nghi': 'Số ngày nghỉ'})
                
                st.session_state['df_kq'] = df_kq
        
        if 'df_kq' in st.session_state:
            df_kq = st.session_state['df_kq']
            
            st.markdown("### Thống kê")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f'<div class="metric-box"><p class="label">Tổng nhân viên</p><p class="value">{len(df_kq)}</p></div>', unsafe_allow_html=True)
            with col2:
                st.markdown(f'<div class="metric-box"><p class="label">Tổng ngày công</p><p class="value">{int(df_kq["So_ngay_cham_cong"].sum()):,}</p></div>', unsafe_allow_html=True)
            with col3:
                st.markdown(f'<div class="metric-box"><p class="label">Tổng ngày nghỉ</p><p class="value">{int(df_kq["Số ngày nghỉ"].sum()):,}</p></div>', unsafe_allow_html=True)
            with col4:
                st.markdown(f'<div class="metric-box"><p class="label">Trung bình / NV</p><p class="value">{df_kq["So_ngay_cham_cong"].mean():.1f}</p></div>', unsafe_allow_html=True)
            
            # Hiển thị dataframe với tên cột đẹp
            df_hien_thi = df_kq[['Ma_nv', 'Ho_ten', 'So_ngay_cham_cong', 'Số ngày nghỉ', 'Công chuẩn']].copy()
            df_hien_thi.columns = ['Mã nhân viên', 'Họ và tên', 'Số ngày chấm công', 'Số ngày nghỉ', 'Công chuẩn']
            
            st.markdown("### Kết quả")
            st.dataframe(df_hien_thi, use_container_width=True, hide_index=True, height=450)
            
            st.markdown("### Tải xuống")
            col1, col2 = st.columns(2)
            with col1:
                excel_buf = xuat_excel_dep(df_kq)
                st.download_button(
                    "Tải Excel (.xlsx)",
                    data=excel_buf,
                    file_name=f"cham_cong_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col2:
                csv_buf = df_hien_thi.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button(
                    "Tải CSV (.csv)",
                    data=csv_buf,
                    file_name=f"cham_cong_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
    
    except Exception as e:
        st.error(f"Lỗi: {str(e)}")
