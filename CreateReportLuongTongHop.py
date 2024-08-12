from Config import *
from openpyxl import Workbook
import pandas as pd

def create_report_luong():
    location_HT = "HỆ THỐNG"
    index = location_list.index(location_HT)
    location_report_folder_name = str(index + 1)+"_"+location_HT
    location_report_folder_path = os.path.join(report_folder,location_report_folder_name)

    excel_file_path = os.path.join(location_report_folder_path, f"Tổng hợp lương {month} - {year}.xlsx")
    # Tạo workbook mới
    wb = Workbook()
    # Tạo report về Doanh số
    ws1 = wb.active
    ws1.title = 'sheet 1'

    for location in location_list:
        if location != "HỆ THỐNG":
            index = location_list.index(location)
            location_report_folder_name = str(index + 1)+"_"+location
            location_report_folder_path = os.path.join(report_folder,location_report_folder_name,"Báo cáo cá nhân")
            report_ca_nhan_file_paths = []
            # Duyệt từng tệp và thư mục trong thư mục hiện tại
            for root, dirs, files in os.walk(os.path.join(location_report_folder_path)):
                for file in files:
                    # Kiểm tra nếu tệp có đuôi là suffix
                    if file.endswith('.xlsx'):
                        report_ca_nhan_file_paths.append(os.path.join(root, file))
            for file in report_ca_nhan_file_paths:
                parts = os.path.basename(file).split()
                sheet_name = " ".join(parts[0:-1])
                ws = wb.create_sheet(title=sheet_name)
                luong_sheet = pd.read_excel(file, sheet_name="Lương")
                writeDataframeToSheet(ws, luong_sheet)
    # Lưu workbook vào file Excel
    try:
        wb.save(excel_file_path)
        print(f"Đã tạo file Excel mới '{excel_file_path}' thành công")
    except Exception as e:
        print(f"Lỗi khi tạo file Excel mới: {e}")

