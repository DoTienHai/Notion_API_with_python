from Config import *
from CreateReportCaNhan import *
from CreateReportCoSo import *
from CreateReportKhachHang import *
from CreateReportLuongTongHop import *


def create_all_report_ca_nhan(location):
    index = location_list.index(location)
    location_report_folder_name = str(index + 1)+"_"+location
    location_report_folder_path = os.path.join(report_folder,location_report_folder_name)
    if not os.path.exists(location_report_folder_path):
        os.makedirs(location_report_folder_path)

    if location != "HỆ THỐNG":
        danh_sach_nhan_su = get_ho_so_nhan_su("",["ALL"])
        danh_sach_nhan_su = danh_sach_nhan_su[danh_sach_nhan_su["Cơ sở"] == location]
        for index_row in range(len(danh_sach_nhan_su)):
            info_nhan_su = danh_sach_nhan_su.iloc[index_row]
            create_report_ca_nhan(location_report_folder_path, info_nhan_su)

def create_all_report_co_so(location):
    index = location_list.index(location)
    location_report_folder_name = str(index + 1)+"_"+location
    location_report_folder_path = os.path.join(report_folder,location_report_folder_name)
    if not os.path.exists(location_report_folder_path):
        os.makedirs(location_report_folder_path)

    create_report_co_so(location_report_folder_path, location)
    create_report_khach_hang(location_report_folder_path, location)
    if location == "HỆ THỐNG":
        create_report_luong()