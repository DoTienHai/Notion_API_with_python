from Utils import *
from datetime import datetime
import pandas as pd
import json

# Lấy ngày hôm nay
today = datetime.today()
start_date = f"{today.year}-01-01"
# start_date = f"{today.year}-{today.month}-01"
update_from_date = datetime.strptime(start_date, date_format)

# Dữ liệu của trang mới
def get_data_cho_luy_ke(location = ""):
    # lấy data đơn giá và đã thanh toán
    data = get_data_doanh_thu(location=location, columns=["Ngày thực hiện", "Đơn giá", "Thanh toán lần đầu"]).groupby("Ngày thực hiện").sum().reset_index()
    data = data.rename(columns={"Ngày thực hiện":"Ngày"})
    # lấy data số lượng đơn
    data_so_don = get_data_doanh_thu(location=location, columns=["Ngày thực hiện"]).groupby('Ngày thực hiện').size().reset_index(name='Số lượng đơn')
    data_so_don = data_so_don.rename(columns={"Ngày thực hiện":"Ngày"})
    data = pd.merge(data, data_so_don, on="Ngày", how='outer')
    # lấy data thu nợ
    data_thu_no = get_data_thu_no(location=location, columns=["Ngày thu", "Lượng thu"]).groupby("Ngày thu").sum().reset_index()
    data_thu_no = data_thu_no.rename(columns={"Ngày thu":"Ngày", "Lượng thu":"Thu nợ"})
    data = pd.merge(data, data_thu_no, on="Ngày", how='outer')
    # lấy data chi tiêu
    data_chi_tieu = get_data_chi_tieu(location=location, columns=["Ngày chi", "Lượng chi"]).groupby("Ngày chi").sum().reset_index()
    data_chi_tieu = data_chi_tieu.rename(columns={"Ngày chi":"Ngày"})
    data = pd.merge(data, data_chi_tieu, on="Ngày", how='outer')
    # Loại bỏ các giá trị NAN
    data = data.fillna(0)
    # Lọc data theo ngày thực hiện
    query_string = f"'{start_date}' <= `Ngày`"
    data = data.query(query_string)
    return data

def update_luy_ke_theo_ngay(location = ""):
    print(f"Update luy kế theo ngày {location}")
    data = get_data_cho_luy_ke(location=location)
    number_of_row = len(data)

    # Đường dẫn tới tệp JSON của bạn
    # API key và database ID của bạn
    database_id = dataBaseDict["LUY_KE_NGAY_HE_THONG"]
    json_file_path = os.path.join(notion_data_folder, 'LUY_KE_NGAY_HE_THONG.json')
    if (location in vn_locations):
        database_id = dataBaseDict[f"LUY_KE_NGAY_{e_locations[vn_locations.index(location)]}"]
        json_file_path = os.path.join(notion_data_folder, f'LUY_KE_NGAY_{e_locations[vn_locations.index(location)]}.json')
    
    # Đọc tệp JSON và chuyển đổi nội dung thành chuỗi JSON
    with open(json_file_path, 'r', encoding='utf-8') as file:
        json_data = json.load(file)

    # Lọc dữ liệu của những ngày cần cập nhật 
    number_of_page = len(json_data)    
    json_data = [item for item in json_data if datetime.strptime(item["properties"]["Ngày"]["date"]["start"], date_format) >= update_from_date]
    bias = number_of_page - len(json_data)
    number_of_page = len(json_data)

    # Update and Create new
    for index in range(number_of_row):
        row = data.iloc[index]
        date_row = row["Ngày"].strftime(date_format)
        if(index < number_of_page):
            # update page in notion
            page_id = json_data[index]["id"]
            # URL endpoint của Notion API
            template_json = {"properties": {
                                "Chi tiêu": {
                                    "number": row["Lượng chi"]
                                },
                                "Đã thanh toán": {
                                    "number": row["Thanh toán lần đầu"]
                                },
                                "Số lượng đơn": {
                                    "number": row["Số lượng đơn"]
                                },
                                "Ngày": {
                                    "date": {
                                        "start": date_row,
                                    }
                                },
                                "Thu nợ": {
                                    "number": row["Thu nợ"]
                                },
                                "Đơn giá": {
                                    "number": row["Đơn giá"]
                                },
                                "STT": {
                                    "title": [
                                        {
                                            "text": {
                                                "content": str(index+1+bias),
                                            },
                                            "plain_text": str(index+1+bias),
                                        }
                                    ]
                                }
                            }
                        }
            update_page(page_id, template_json)
        else:
            # create new data
            template_json = {"parent": {"database_id": database_id},
                            "properties": {
                                "Chi tiêu": {
                                    "number": row["Lượng chi"]
                                },
                                "Đã thanh toán": {
                                    "number": row["Thanh toán lần đầu"]
                                },
                                "Số lượng đơn": {
                                    "number": row["Số lượng đơn"]
                                },
                                "Ngày": {
                                    "date": {
                                        "start": date_row,
                                    }
                                },
                                "Thu nợ": {
                                    "number": row["Thu nợ"]
                                },
                                "Đơn giá": {
                                    "number": row["Đơn giá"]
                                },
                                "STT": {
                                    "title": [
                                        {
                                            "text": {
                                                "content": str(index+1+bias),
                                            },
                                            "plain_text": str(index+1+bias),
                                        }
                                    ]
                                }
                            }
                        }
            
            create_page(template_json)
        print(f"{location} {index+1}/{number_of_row}")
    print(f"Đã update luy kế theo ngày {location}")


def update_luy_ke_theo_thang(location = ""):
    print(f"Update luy kế theo tháng {location}")

    data = get_data_cho_luy_ke(location=location)
    data["Ngày"] = pd.to_datetime(data["Ngày"]).dt.month
    data = data.rename(columns={"Ngày" : "Tháng"})
    data = data.groupby("Tháng").sum().reset_index()
    def add_thang(item):
        return "Tháng " + str(item)
    data["Tháng"] = data["Tháng"].apply(add_thang)
    number_of_row = len(data)
    
    # Đường dẫn tới tệp JSON của bạn
    # API key và database ID của bạn
    database_id = dataBaseDict["LUY_KE_THANG_HE_THONG"]
    json_file_path = os.path.join(notion_data_folder, 'LUY_KE_THANG_HE_THONG.json')
    if (location in vn_locations):
        database_id = dataBaseDict[f"LUY_KE_THANG_{e_locations[vn_locations.index(location)]}"]
        json_file_path = os.path.join(notion_data_folder, f'LUY_KE_THANG_{e_locations[vn_locations.index(location)]}.json')
        
    # Đọc tệp JSON và chuyển đổi nội dung thành chuỗi JSON
    with open(json_file_path, 'r', encoding='utf-8') as file:
        json_data = json.load(file)

    number_of_page = len(json_data)    

    for index_row in range(number_of_row):
        row = data.iloc[index_row]
        for index_page in range(number_of_page):
            if(row["Tháng"] == json_data[index_page]["properties"]["Tháng"]["title"][0]["plain_text"]):
                page_id = json_data[index_page]["id"]
                template_json = {
                                "parent": {
                                    "database_id": database_id
                                },
                                "properties": {
                                    "Chi tiêu": {
                                        "number": row["Lượng chi"]
                                    },
                                    "Đã thanh toán": {
                                        "number": row["Thanh toán lần đầu"]
                                    },
                                    "Số lượng đơn": {
                                        "number": row["Số lượng đơn"]
                                    },
                                    "Thu nợ": {
                                        "number": row["Thu nợ"]
                                    },
                                    "Đơn giá": {
                                        "number": row["Đơn giá"]
                                    },
                                },
                            }
                update_page(page_id, template_json)
    print(f"Đã update luy kế theo tháng {location}")
