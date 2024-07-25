from Config import *
from CreateReportCaNhan import *

def update_KPI():
    data_nhan_su = get_ho_so_nhan_su("HỆ THỐNG", ["notion id"])
    data_danh_thu = get_data_doanh_thu("HỆ THỐNG", ["ALL"])
    data_danh_thu = filter_date(data_danh_thu, "Ngày thực hiện")

    for index in range(len(data_nhan_su)):
        notion_id_nhan_su = data_nhan_su.iloc[index].values[0]
        tong_doanh_so_sale_chinh = data_danh_thu[data_danh_thu["id sale chính"] == notion_id_nhan_su]["Đơn giá gốc"].sum()
        tong_doanh_so_sale_phu = data_danh_thu[data_danh_thu["id sale phụ"] == notion_id_nhan_su]["Upsale"].sum()
        tong_doanh_so = tong_doanh_so_sale_chinh + tong_doanh_so_sale_phu

        template_json = {
                "properties": {
                "Doanh số": {
                    "id": "jn%3BH",
                    "type": "number",
                    "number": tong_doanh_so
                },
                }
            }
        update_page(notion_id_nhan_su, template_json)
    print("Đã update KPI nhân sự!")
    
