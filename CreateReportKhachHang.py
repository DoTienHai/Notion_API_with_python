from Utils import *
import os
from openpyxl import Workbook

def create_report_khach_hang(location=""):
    data = get_danh_sach_khach_hang(location, columns=["ALL"])
    data = data.drop(columns=["notion id"])
    # Kiểm tra xem file Excel đã tồn tại hay chưa
    excel_file_path = os.path.join(report_khach_hang, f"Danh sách khách hàng {location}.xlsx")
    if os.path.exists(excel_file_path):
        # Nếu đã tồn tại, xóa file cũ đi
        try:
            os.remove(excel_file_path)
            print(f"Đã xóa file Excel cũ '{excel_file_path}'")
        except Exception as e:
            print(f"Lỗi khi xóa file Excel cũ: {e}")
            # Tạo workbook mới
    wb = Workbook()
    # Tạo sheet Đơn sale chính
    ws1 = wb.active
    ws1.title = 'Danh sách khách hàng'
    writeDataframeToSheet(ws1, data)

    # Lưu workbook vào file Excel
    try:
        wb.save(excel_file_path)
        print(f"Đã tạo file Excel mới '{excel_file_path}' thành công")
    except Exception as e:
        print(f"Lỗi khi tạo file Excel mới: {e}")

def create_all_report_khach_hang():
    create_report_khach_hang()
    for location in vn_locations:
        create_report_khach_hang(location)