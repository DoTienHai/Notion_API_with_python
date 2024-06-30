import os
from openpyxl import Workbook
import pandas as pd
from datetime import datetime
from Utils import *

today = datetime.today()

def filter_date(data, column_name):
    data[column_name] = pd.to_datetime(data[column_name])
    data = data[(data[column_name].dt.year == today.year) & (data[column_name].dt.month == today.month)]
    data = data.rename(columns={column_name:f"{column_name}_temp"})
    data[column_name] = data[f"{column_name}_temp"].dt.strftime('%m-%d-%Y')
    return data

def get_don_sale_chinh(notion_id):
    data = get_data_doanh_thu("",["ALL"])
    data = data[data["id sale chính"] == notion_id]
    data = filter_date(data, "Ngày thực hiện")
    data = data[["Tiền tố", "Mã dịch vụ", "Ngày thực hiện",
                "Cơ sở", "Khách hàng", "Nguồn khách", "Tên dịch vụ", "Sale chính", "Đơn giá gốc", 
                "Sale phụ", "Upsale", "Đơn giá", "Thanh toán lần đầu", "Trả sau",
                "Đã thanh toán", "Dư nợ", "Bác sĩ 1", "Bác sĩ 2", "Phụ phẫu 1", "Phụ phẫu 2" ]]
    return data

def get_don_sale_phu(notion_id):
    data = get_data_doanh_thu("",["ALL"])
    data = data[data["id sale phụ"] == notion_id]
    data = filter_date(data, "Ngày thực hiện")
    data = data[["Tiền tố", "Mã dịch vụ", "Ngày thực hiện",
                "Cơ sở", "Khách hàng", "Nguồn khách", "Tên dịch vụ", "Sale chính", "Đơn giá gốc", 
                "Sale phụ", "Upsale", "Đơn giá", "Thanh toán lần đầu", "Trả sau",
                "Đã thanh toán", "Dư nợ", "Bác sĩ 1", "Bác sĩ 2", "Phụ phẫu 1", "Phụ phẫu 2" ]]
    return data

def get_don_1_bac_si(notion_id):
    data = get_data_doanh_thu("",["ALL"])
    data = data[(data["id bác sĩ 1"] == notion_id) & (data["id bác sĩ 2"].isnull())]
    data = filter_date(data, "Ngày thực hiện")
    data = data[["Tiền tố", "Mã dịch vụ", "Ngày thực hiện",
                "Cơ sở", "Khách hàng", "Nguồn khách", "Tên dịch vụ", "Sale chính", "Đơn giá gốc", 
                "Sale phụ", "Upsale", "Đơn giá", "Thanh toán lần đầu", "Trả sau",
                "Đã thanh toán", "Dư nợ", "Bác sĩ 1", "Bác sĩ 2", "Phụ phẫu 1", "Phụ phẫu 2" ]]
    return data

def get_don_2_bac_si(notion_id):
    data = get_data_doanh_thu("",["ALL"])
    data = data[(data["id bác sĩ 1"] == notion_id)& ~(data["id bác sĩ 2"].isnull()) | (data["id bác sĩ 2"] == notion_id)]
    data = filter_date(data, "Ngày thực hiện")
    data = data[["Tiền tố", "Mã dịch vụ", "Ngày thực hiện",
                "Cơ sở", "Khách hàng", "Nguồn khách", "Tên dịch vụ", "Sale chính", "Đơn giá gốc", 
                "Sale phụ", "Upsale", "Đơn giá", "Thanh toán lần đầu", "Trả sau",
                "Đã thanh toán", "Dư nợ", "Bác sĩ 1", "Bác sĩ 2", "Phụ phẫu 1", "Phụ phẫu 2" ]]
    return data

def get_don_phụ_phau_1(notion_id):
    data = get_data_doanh_thu("",["ALL"])
    data = data[data["id phụ phẫu 1"] == notion_id]
    data = filter_date(data, "Ngày thực hiện")
    data = data[["Tiền tố", "Mã dịch vụ", "Ngày thực hiện",
                "Cơ sở", "Khách hàng", "Nguồn khách", "Tên dịch vụ", "Sale chính", "Đơn giá gốc", 
                "Sale phụ", "Upsale", "Đơn giá", "Thanh toán lần đầu", "Trả sau",
                "Đã thanh toán", "Dư nợ", "Bác sĩ 1", "Bác sĩ 2", "Phụ phẫu 1", "Phụ phẫu 2" ]]
    return data

def get_don_phụ_phau_2(notion_id):
    data = get_data_doanh_thu("",["ALL"])
    data = data[data["id phụ phẫu 2"] == notion_id]
    data = filter_date(data, "Ngày thực hiện")
    data = data[["Tiền tố", "Mã dịch vụ", "Ngày thực hiện",
                "Cơ sở", "Khách hàng", "Nguồn khách", "Tên dịch vụ", "Sale chính", "Đơn giá gốc", 
                "Sale phụ", "Upsale", "Đơn giá", "Thanh toán lần đầu", "Trả sau",
                "Đã thanh toán", "Dư nợ", "Bác sĩ 1", "Bác sĩ 2", "Phụ phẫu 1", "Phụ phẫu 2" ]]
    return data

def get_don_thu_no(notion_id):
    data = get_data_thu_no("", ["ALL"])
    data = data[data["id sale"] == notion_id]
    data = filter_date(data, "Ngày thu")
    data = data[["Tiền tố", "Mã đơn thu nợ", "Đơn nợ",
                  "Cơ sở", "Lượng thu", "Sale",  "Ngày thu"]]
    return data
    

def create_doanh_so_ca_nhan():
    danh_sach_nhan_su = get_ho_so_nhan_su("",["notion id", "Tiền tố", "Mã nhân viên", "Họ và tên"])
    for index_row in range(len(danh_sach_nhan_su)):
        row = danh_sach_nhan_su.iloc[index_row]
        notion_id = row["notion id"]
        ho_va_ten = row["Họ và tên"]
        ma_nhan_vien = f"{row["Tiền tố"]}-{row["Mã nhân viên"]}"
        # Kiểm tra xem file Excel đã tồn tại hay chưa
        excelFilePath = f"report_ca_nhan\\{ma_nhan_vien} {ho_va_ten} {today.month}-{today.year}.xlsx"
        if os.path.exists(excelFilePath):
            # Nếu đã tồn tại, xóa file cũ đi
            try:
                os.remove(excelFilePath)
                print(f"Đã xóa file Excel cũ '{excelFilePath}'")
            except Exception as e:
                print(f"Lỗi khi xóa file Excel cũ: {e}")

        # Tạo workbook mới
        wb = Workbook()
        # Tạo sheet Đơn sale chính
        ws1 = wb.active
        ws1.title = 'Đơn sale chính'
        data_sale_chinh = get_don_sale_chinh(notion_id)
        writeDataframeToSheet(ws1, data_sale_chinh)
        # Tạo sheet Đơn sale phụ
        ws2 = wb.create_sheet(title='Đơn sale phụ')
        data_sale_phu = get_don_sale_phu(notion_id)
        writeDataframeToSheet(ws2, data_sale_phu)
        # Tạo sheet Đơn 1 bác sĩ 
        ws3 = wb.create_sheet(title="Đơn 1 bác sĩ")
        data_don_1_bac_si = get_don_1_bac_si(notion_id)
        writeDataframeToSheet(ws3, data_don_1_bac_si)
        # Tạo sheet Đơn 2 bác sĩ
        ws4 = wb.create_sheet(title="Đơn 2 bác sĩ")
        data_don_2_bac_si = get_don_2_bac_si(notion_id)
        writeDataframeToSheet(ws4, data_don_2_bac_si)
        # Tạo sheet Đơn phụ phẫu 1
        ws5 = wb.create_sheet("Đơn phụ phẫu 1")
        data_phu_phau_1 = get_don_phụ_phau_1(notion_id)
        writeDataframeToSheet(ws5, data_phu_phau_1)
        # Tạo sheet Đơn phụ phẫu 2
        ws6 = wb.create_sheet("Đơn phụ phẫu 2")
        data_phu_phau_2 = get_don_phụ_phau_2(notion_id)
        writeDataframeToSheet(ws6, data_phu_phau_2)
        # Tạo sheet Đơn thu nợ
        ws7 = wb.create_sheet("Đơn thu nợ")
        data_don_thu_no = get_don_thu_no(notion_id)
        writeDataframeToSheet(ws7, data_don_thu_no)
        # Tạo sheet Tổng hợp
        # ws8 = wb.create_sheet("Tổng hợp")
        # data_tong_hop = pd.DataFrame(columns=["Loại", "Đơn giá gốc", "Upsale", "Thanh toán lần đầu",  "Đơn giá", "Đã thanh toán"])
        # writeDataframeToSheet(ws8, data_tong_hop)


        # Lưu workbook vào file Excel
        try:
            wb.save(excelFilePath)
            print(f"Đã tạo file Excel mới '{excelFilePath}' thành công")
        except Exception as e:
            print(f"Lỗi khi tạo file Excel mới: {e}")

    