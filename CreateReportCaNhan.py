import os
from openpyxl import Workbook
import pandas as pd
from datetime import datetime
from Utils import *

month = datetime.today().month - 1
year = datetime.today().year

columns = ["Tiền tố", "Mã dịch vụ", "Ngày thực hiện",
                "Cơ sở", "Khách hàng", "Nguồn khách", "Tên dịch vụ", "Sale chính", "Đơn giá gốc", 
                "Sale phụ", "Upsale", "Đơn giá", "Thanh toán lần đầu", "Trả sau",
                "Đã thanh toán", "Dư nợ", "Bác sĩ 1", "Bác sĩ 2", "Phụ phẫu 1", "Phụ phẫu 2", "Công phụ phẫu 1", "Công phụ phẫu 2" ]

def filter_date(data, column_name):
    data[column_name] = pd.to_datetime(data[column_name])
    data = data[(data[column_name].dt.year == year) & (data[column_name].dt.month == month)]
    data = data.rename(columns={column_name:f"{column_name}_temp"})
    data[column_name] = data[f"{column_name}_temp"].dt.strftime('%m-%d-%Y')
    data = data.drop(columns=[f"{column_name}_temp"])
    columns = ['Ngày'] + [col for col in data.columns if col != 'Ngày']
    data = data[columns]
    return data

def add_total_row(data):
    sum_data = data.select_dtypes(include=['number']).sum()
    sum_data["Mã dịch vụ"] = data["Mã dịch vụ"].count()  
    total_df = pd.DataFrame(sum_data).T
    # Thêm các cột không phải là số vào dòng tổng
    for col in data.columns:
        if col not in total_df.columns:
            total_df[col] = ''  
    # Đặt lại thứ tự các cột để khớp với DataFrame gốc
    total_df = total_df[data.columns]
    total_df["Tiền tố"] = "Tổng"
    # Nối dòng tổng với DataFrame gốc
    data = pd.concat([data, total_df])
    return data 

def get_don_sale_chinh(notion_id):
    data = get_data_doanh_thu("",["ALL"])
    data = data[data["id sale chính"] == notion_id]
    data = filter_date(data, "Ngày thực hiện")
    data = data[columns]
        
    data = add_total_row(data) 

    return data

def get_don_sale_phu(notion_id):
    data = get_data_doanh_thu("",["ALL"])
    data = data[data["id sale phụ"] == notion_id]
    data = filter_date(data, "Ngày thực hiện")
    data = data[columns]
    data = add_total_row(data) 
    return data

def get_don_1_bac_si(notion_id):
    data = get_data_doanh_thu("",["ALL"])
    data = data[(data["id bác sĩ 1"] == notion_id) & (data["id bác sĩ 2"].isnull())]
    data = filter_date(data, "Ngày thực hiện")
    data = data[columns]
    data = add_total_row(data) 
    return data

def get_don_2_bac_si(notion_id):
    data = get_data_doanh_thu("",["ALL"])
    data = data[(data["id bác sĩ 1"] == notion_id)& ~(data["id bác sĩ 2"].isnull()) | (data["id bác sĩ 2"] == notion_id)]
    data = filter_date(data, "Ngày thực hiện")
    data = data[columns]
    data = add_total_row(data) 
    return data

def get_don_phụ_phau_1(notion_id):
    data = get_data_doanh_thu("",["ALL"])
    data = data[data["id phụ phẫu 1"] == notion_id]
    data = filter_date(data, "Ngày thực hiện")
    data = data[columns]
    data = add_total_row(data) 
    return data

def get_don_phụ_phau_2(notion_id):
    data = get_data_doanh_thu("",["ALL"])
    data = data[data["id phụ phẫu 2"] == notion_id]
    data = filter_date(data, "Ngày thực hiện")
    data = data[columns]
    data = add_total_row(data) 

    return data

def get_don_thu_no(notion_id):
    data = get_data_thu_no("", ["ALL"])
    data = data[data["id sale"] == notion_id]
    data = filter_date(data, "Ngày thu")
    data = data[["Tiền tố", "Mã đơn thu nợ", "Đơn nợ",
                  "Cơ sở", "Lượng thu", "Sale",  "Ngày thu"]]
    
    sum_data = data.select_dtypes(include=['number']).sum()
    sum_data["Mã đơn thu nợ"] = data["Mã đơn thu nợ"].count()  
    total_df = pd.DataFrame(sum_data).T
    # Thêm các cột không phải là số vào dòng tổng
    for col in data.columns:
        if col not in total_df.columns:
            total_df[col] = ''  
    # Đặt lại thứ tự các cột để khớp với DataFrame gốc
    total_df = total_df[data.columns]
    total_df["Tiền tố"] = "Tổng"
    # Nối dòng tổng với DataFrame gốc
    data = pd.concat([data, total_df])

    return data
    

def create_doanh_so_ca_nhan():
    danh_sach_nhan_su = get_ho_so_nhan_su("",["notion id", "Tiền tố", "Mã nhân viên", "Họ và tên"])
    for index_row in range(len(danh_sach_nhan_su)):
        row = danh_sach_nhan_su.iloc[index_row]
        notion_id = row["notion id"]
        ho_va_ten = row["Họ và tên"]
        ma_nhan_vien = f"{row["Tiền tố"]}-{row["Mã nhân viên"]}"
        # Kiểm tra xem file Excel đã tồn tại hay chưa
        excel_file_path = os.path.join("report_ca_nhan", f"{ma_nhan_vien} {ho_va_ten} {month}-{year}.xlsx")
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
        ws1.title = 'Đơn sale chính'
        data_sale_chinh = get_don_sale_chinh(notion_id)
        if len(data_sale_chinh) > 1:
            writeDataframeToSheet(ws1, data_sale_chinh)
        # Tạo sheet Đơn sale phụ
        data_sale_phu = get_don_sale_phu(notion_id)
        if len(data_sale_phu) > 1:
            ws2 = wb.create_sheet(title='Đơn sale phụ')
            writeDataframeToSheet(ws2, data_sale_phu)
        # Tạo sheet Đơn 1 bác sĩ 
        data_don_1_bac_si = get_don_1_bac_si(notion_id)
        if len(data_don_1_bac_si) > 1:
            ws3 = wb.create_sheet(title="Đơn 1 bác sĩ")
            writeDataframeToSheet(ws3, data_don_1_bac_si)
        # Tạo sheet Đơn 2 bác sĩ
        data_don_2_bac_si = get_don_2_bac_si(notion_id)
        if len(data_don_2_bac_si) > 1:
            ws4 = wb.create_sheet(title="Đơn 2 bác sĩ")
            writeDataframeToSheet(ws4, data_don_2_bac_si)
        # Tạo sheet Đơn phụ phẫu 1
        data_phu_phau_1 = get_don_phụ_phau_1(notion_id)
        if len(data_phu_phau_1) > 1:
            ws5 = wb.create_sheet("Đơn phụ phẫu 1")
            writeDataframeToSheet(ws5, data_phu_phau_1)
        # Tạo sheet Đơn phụ phẫu 2
        data_phu_phau_2 = get_don_phụ_phau_2(notion_id)
        if len(data_phu_phau_2) > 1:
            ws6 = wb.create_sheet("Đơn phụ phẫu 2")
            writeDataframeToSheet(ws6, data_phu_phau_2)
        # Tạo sheet Đơn thu nợ
        data_don_thu_no = get_don_thu_no(notion_id)
        if len(data_don_thu_no) > 1:
            ws7 = wb.create_sheet("Đơn thu nợ")
            writeDataframeToSheet(ws7, data_don_thu_no)
        # Tạo sheet Tổng hợp
        # ws8 = wb.create_sheet("Tổng hợp")
        # data_tong_hop = pd.DataFrame(columns=["Loại", "Đơn giá gốc", "Upsale", "Thanh toán lần đầu",  "Đơn giá", "Đã thanh toán"])
        # writeDataframeToSheet(ws8, data_tong_hop)


        # Lưu workbook vào file Excel
        try:
            wb.save(excel_file_path)
            print(f"Đã tạo file Excel mới '{excel_file_path}' thành công")
        except Exception as e:
            print(f"Lỗi khi tạo file Excel mới: {e}")

# create_doanh_so_ca_nhan()