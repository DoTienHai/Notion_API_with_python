import pandas as pd
import requests
import json
from pandas import DataFrame
import os
from datetime import datetime

notion_data_folder = "Notion data"
file_all_notion_data = os.path.join(notion_data_folder, "ALL.xlsx")
report_folder = "Báo cáo"

notion_api_token = "secret_o7gsjeVNFo5Wpg1bUJ7eHo8VpkF7riKIEAIcB8P0HyR"
dataBaseDict = {
    "Hồ sơ nhân sự" : "3213b32ae23044e2afdd04abcc992e96",
    "Thông tin khách hàng" : "0079739c1dde4bb481129cb500ff0df6",
    "Doanh thu HỆ THỐNG" : "30eeb9f5c324499387973eba8662f1d0",
    "Chi tiêu" : "a31162209bca40f486878c8549b81c3f",
    "Danh sách thu nợ" : '34df6c27358c48729e6af38f304a4a8f',
    "Danh mục dịch vụ" : "a305ae47994942ffa7e4c2f249b32723",
    "Lũy kế tháng HỆ THỐNG" : "1b6d19049c284fa3bbef3e6d432c9260",
    "Lũy kế tháng CẦN THƠ" : "f6449c902be44ed9b1fc1d02d6aa6b1c",
    "Lũy kế tháng LONG XUYÊN" : "497028c136b04c74a230586750c5fa5f",
    "Lũy kế tháng SÓC TRĂNG" : "b8a9201150ec43b5b3e5f54ce27757ed",
    "Lũy kế ngày HỆ THỐNG" : "88e410cca4104367aad555987f9467f7",
    "Lũy kế ngày CẦN THƠ" : "f209a4184aa24b659cfc4a94b6af86b0",
    "Lũy kế ngày LONG XUYÊN" : "95908de9fba942478bcd04b57e56bd1b",
    "Lũy kế ngày SÓC TRĂNG" : "b881c341e6ba4fa5acba1d0d0c17746d",
    "Chấm công HỆ THỐNG" : "0c70fa057bda4ceea9f23ccf963208ff",
    "Chấm công CẦN THƠ" : "81b741581c2e4c3aa685ce6602d70cd7",
    "Chấm công LONG XUYÊN" : "7a9c30037d164eafa09da55aafacde5e",
    "Chấm công SÓC TRĂNG" : "c3e4dde0278f416b9ee0edd617b07a1e" 
}

cham_cong_ref = {
    "Nghỉ không phép" : 0,
    "Nghỉ có phép" : 1,
    "Đầy đủ" : 1,
    "Nửa ngày" : 0.5
}

location_list = ["CẦN THƠ", "LONG XUYÊN", "SÓC TRĂNG", "HỆ THỐNG"]

date_format = "%Y-%m-%d"
month = datetime.today().month
year = datetime.today().year

###-------------------------- LIÊN QUAN ĐẾN DATAFRAME ---------------------------###
def filter_column(data, location, columns):
    if ("Cơ sở" in data.columns.tolist()):
        if (location):
            if (location in location_list) and (location != "HỆ THỐNG"):
                data = data[data["Cơ sở"] == location]

    if (columns[0] == "ALL"):
        return data   
         
    valid_field = []
    for field in columns:
        if (field in data.columns.to_list()):
            valid_field.append(field)
    if valid_field:
        return data[valid_field]
    else:
        return None

def get_data_doanh_thu(location, columns):
    data = pd.read_excel(file_all_notion_data, sheet_name="Doanh thu HỆ THỐNG", parse_dates=['Ngày thực hiện'], date_format=date_format)
    return filter_column(data, location, columns)

def get_data_thu_no(location, columns):
    data = pd.read_excel(file_all_notion_data, sheet_name="Thu nợ", parse_dates=['Ngày thu'], date_format=date_format)
    return filter_column(data, location, columns)

def get_data_chi_tieu(location, columns):
    data =  pd.read_excel(file_all_notion_data, sheet_name="Chi tiêu", parse_dates=['Ngày chi'], date_format=date_format)
    return filter_column(data, location, columns)

def get_data_danh_muc_dich_vu(location, columns):
    data =  pd.read_excel(file_all_notion_data, sheet_name="danh mục dịch vụ")
    return filter_column(data, location, columns)   

def get_ho_so_nhan_su(location, columns):
    data =  pd.read_excel(file_all_notion_data, sheet_name="Hồ sơ nhân sự")
    return filter_column(data, location, columns)  

def get_danh_sach_khach_hang(location, columns):
    data = pd.read_excel(file_all_notion_data, sheet_name="Danh sách khách hàng") 
    return filter_column(data, location, columns) 

def get_data_cham_cong(location, columns):
    sheet_name = f"Chấm công {location}"
    data = pd.read_excel(file_all_notion_data, sheet_name)
    location = ""
    return filter_column(data, location, columns)

def get_data_cham_cong_tong_hop():
    data = pd.read_excel(file_all_notion_data, "Chấm công HỆ THỐNG")
    return data
###------------------------------------------- LIÊN QUAN ĐẾN NOTION API ---------------------------------------###
# Headers notion api token
headers = {
    "Authorization": f"Bearer {notion_api_token}",
    "Content-Type": "application/json",
    "Notion-Version": "2022-06-28"
}
def create_page(json_data):
    url = f"https://api.notion.com/v1/pages/"
    response = requests.post(url, headers=headers, data=json.dumps(json_data))

    # Kiểm tra kết quả
    if response.status_code == 200:
        # print(f"Lũy kế đã được tạo mới thành công!")
        pass
    else:
        print(f"Tạo mới lũy kế đã xảy ra lỗi: {response.text}")
        print(response.text)
        
def update_page(page_id, json_data):
    url = f"https://api.notion.com/v1/pages/{page_id}"
    response = requests.patch(url, headers=headers, data=json.dumps(json_data))

    # Kiểm tra kết quả
    if response.status_code == 200:
        # print(f"Lũy kế đã được update thành công!")
        pass
    else:
        print(f"Update lũy kế đã xảy ra lỗi: {response.text}")

###-------------------------------- LIÊN QUAN ĐẾN FILE EXCEL ------------------------------------###
def writeDataframeToSheet(ws, dataframe: pd.DataFrame):
    if(dataframe is not None):
        # Ghi tên cột vào hàng đầu tiên
        for col_num, column_title in enumerate(dataframe.columns, 1):
            ws.cell(row=1, column=col_num, value=column_title)

        # Ghi từng hàng dữ liệu vào sheet
        for row_num, row in enumerate(dataframe.itertuples(index=False), 2):
            for col_num, value in enumerate(row, 1):
                ws.cell(row=row_num, column=col_num, value=value)
    else:
        print(f"Dataframe không có data")

def readSheetFromExcel(excel_file_path, sheet_name) -> DataFrame:
    try:
        # Đọc dữ liệu từ sheet cụ thể trong file Excel
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        # print(f"Đã đọc dữ liệu từ sheet '{sheet_name}' trong file '{excel_file_path}' thành công")
        return df
    except Exception as e:
        print(f"Lỗi khi đọc dữ liệu từ file Excel: {e}")
        return None
    

def moveRowToEnd(dataframe, columnName, rowName):
    # Tách hàng 
    row = dataframe[dataframe[columnName] == rowName] 
    # Loại bỏ hàng  từ DataFrame ban đầu
    dataframe = dataframe[dataframe[columnName] != rowName]
    # Thêm hàng vào cuối DataFrame
    dataframe = pd.concat([dataframe, row], ignore_index=True)
    
    return dataframe

def format_percent(ws, col):
    pass