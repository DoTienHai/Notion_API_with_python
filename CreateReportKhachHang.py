from Config import *
import os
from openpyxl import Workbook

def create_report_khach_hang(path, location):
    data = get_danh_sach_khach_hang(location, columns=["ALL"])
    data = data.drop(columns=["notion id"])
    # Kiểm tra xem file Excel đã tồn tại hay chưa
    folder_path = os.path.join(path, "Danh sách khách hàng") 
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # Tạo report toàn bộ danh sách khách hàng 
    excel_file_path = os.path.join(folder_path,f"Danh sách khách hàng tại {location}.xlsx")
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

    # Tạo report danh sách khách hàng có dư nợ
    excel_file_path = os.path.join(folder_path,f"Danh sách khách hàng còn dư nợ tại {location}.xlsx")
    if os.path.exists(excel_file_path):
        # Nếu đã tồn tại, xóa file cũ đi
        try:
            os.remove(excel_file_path)
            print(f"Đã xóa file Excel cũ '{excel_file_path}'")
        except Exception as e:
            print(f"Lỗi khi xóa file Excel cũ: {e}")
            # Tạo workbook mới
    wb2 = Workbook()
    # Tạo sheet Đơn sale chính
    ws1 = wb2.active
    ws1.title = 'Danh sách khách hàng'
    writeDataframeToSheet(ws1, data[data["Dư nợ"] > 0])

    # Lưu workbook vào file Excel
    try:
        wb2.save(excel_file_path)
        print(f"Đã tạo file Excel mới '{excel_file_path}' thành công")
    except Exception as e:
        print(f"Lỗi khi tạo file Excel mới: {e}")