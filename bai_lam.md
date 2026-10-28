# BÀI LÀM – HACKATHON 01

- **Mã đề:** 13
- **Họ tên:** Phan Huy Hung
- **Lớp:** CNTT1
- **Link repository:** [https://github.com/huyyhung05/CNTT1_Nhap-mon-CNTT_Hackathon01]
- **Hệ điều hành:** Windows (PowerShell)

## 4.2. Thao tác máy tính cơ bản bằng CLI

### Các lệnh thực hiện

```powershell
Get-Location
ls -Force
mv .\thoi_khoa_bieu.txt .\docs\
mv .\logo_trung_tam_ngoai_ngu.png .\media\
cp .\media\logo_trung_tam_ngoai_ngu.png .\backup\
rm .\temp_213.tmp
cd .\logs
ls
cd ..
Get-ChildItem -Recurse -Force
```

Các file Excel và log đã được chuyển bằng File Explorer ở mục 4.1.

### Kết quả liệt kê toàn bộ dự án (gồm file ẩn)

```text
ngoai-ngu-website/
├── .env
├── .gitignore
├── README.md
├── bai_lam.md
├── auto_report.ps1
├── docs/
│   └── thoi_khoa_bieu.txt
├── media/
│   └── logo_trung_tam_ngoai_ngu.png
├── data/
│   └── dang_ky_hoc_thang_3.xlsx
├── logs/
│   ├── course.log
│   ├── daily_report.txt
│   ├── auto_report.txt
│   └── tong_hop.txt
└── backup/
    └── logo_trung_tam_ngoai_ngu.png
```

## 4.3. Lọc log thủ công

### Lệnh

```powershell
Get-Content .\logs\course.log | Select-String "ERROR" > .\logs\daily_report.txt
Get-Content .\logs\daily_report.txt
```

### Hai dòng lỗi đầu tiên

```text
2026-03-28 08:24:28 [ERROR] Het thoi gian cho (timeout)
2026-03-28 08:30:28 [ERROR] Het thoi gian cho (timeout)
```

## 4.4. Thao tác với tài liệu Markdown (README.md)

- [x] Heading 1: tên dự án.
- [x] Có chữ **in đậm** và _in nghiêng_.
- [x] Có bảng thư mục và nội dung.
- [x] Có code block lệnh PowerShell.
- [x] Có task list.
- [x] Có ghi chú không upload `.env`.

## 4.5. Chạy script auto_report.ps1

### Lệnh chạy

```powershell
.\auto_report.ps1
```

Nếu PowerShell chặn script, chạy trong cửa sổ hiện tại:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\auto_report.ps1
```

### Kết quả cần hiển thị

```text
Dang loc log...
Hoan thanh!
```

File `logs/auto_report.txt` có tiêu đề và 6 dòng ERROR.

## 4.6. Đếm số dòng lỗi

### Lệnh

```powershell
Get-Content .\logs\course.log | Select-String "ERROR" | Measure-Object
```

### Kết quả Count

```text
Count : 6
```

## 4.7. Báo cáo tổng hợp

### Các lệnh

```powershell
"=== TONG HOP LOG ===" > .\logs\tong_hop.txt
Get-Content .\logs\course.log | Select-String "WARNING" >> .\logs\tong_hop.txt
(Get-Content .\logs\course.log | Select-String "ERROR" | Measure-Object).Count >> .\logs\tong_hop.txt
Get-Content .\logs\tong_hop.txt
```

### Nội dung logs/tong_hop.txt

```text
=== TONG HOP LOG ===
2026-03-28 08:51:28 [WARNING] Thoi gian tai trang cham
2026-03-28 08:54:28 [WARNING] Video bai giang dung luong lon
2026-03-28 08:57:28 [WARNING] Thoi gian tai trang cham
6
```

## 4.8. Excel

Cong thuc sheet `DangKyHoc`:

- `E2:E16`: `=VLOOKUP(C2,KhoaHoc!$A$2:$C$8,2,FALSE)`
- `F2:F16`: `=VLOOKUP(C2,KhoaHoc!$A$2:$C$8,3,FALSE)`
- `G2:G16`: `=D2*F2`

Đã bật nút lọc cho bảng. Điều kiện cần chọn ở cột `tong_hoc_phi`: **lớn hơn hoặc bằng 9000000**. Có 8 phiếu đáp ứng.
