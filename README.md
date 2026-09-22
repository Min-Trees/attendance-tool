# attendance-tool

Tool Python + Streamlit để định dạng và xuất dữ liệu chấm công từ file Excel.

## 🚀 Deploy lên Internet (Miễn phí!)

### Cách 1: Streamlit Cloud (Khuyến nghị)

1. Truy cập [share.streamlit.io](https://share.streamlit.io)
2. Đăng nhập với GitHub
3. Click **New app**
4. Chọn repository `attendance-tool`
5. Chọn branch `main`
6. Set main file path: `app.py`
7. Click **Deploy!**

## 📖 Cách sử dụng

1. Upload file Excel từ thanh bên trái
2. Tool tự động phát hiện các cột: Mã nhân viên, Họ và tên, Số ngày chấm công, Công chuẩn
3. Chọn cột thủ công nếu tên cột không khớp
4. Xem trước dữ liệu và kiểm tra
5. Tải xuống file đã định dạng (Excel hoặc CSV)

## 💻 Chạy Local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 📁 Cấu trúc Files

```
attendance-tool/
├── app.py              # File chính chạy Streamlit
├── requirements.txt    # Thư viện cần thiết
├── run.bat             # Script chạy trên Windows
├── run.ps1             # Script PowerShell
└── README.md           # File này
```

---

Made with ❤️ using Streamlit
