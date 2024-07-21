import multiprocessing.process
from Config import *
from GetNotionDataToJson import get_all_data_to_json
from CollectJsonToExcel import collect_data
from CreateReportCaNhan import create_doanh_so_ca_nhan
from CreateReportCoSo import create_report_co_so
from CreateReportKhachHang import create_all_report_khach_hang
from UpdateLuyKe import update_luy_ke_theo_ngay, update_luy_ke_theo_thang
from UpdateChamCong import update_cham_cong_tong_hop
import time
from datetime import datetime
import multiprocessing

def fetch_data():
    get_all_data_to_json()
    collect_data()

if __name__ == "__main__":
    while(1):
        start_time = time.time()
        process1 = multiprocessing.Process(target=fetch_data)
        process1.start()
        process1.join()
        print(f"Cập nhật toàn bộ data {(time.time() - start_time):.6f} giây\n")
        
        processes = []
        process2 = multiprocessing.Process(target=update_luy_ke_theo_ngay)
        process3 = multiprocessing.Process(target=update_luy_ke_theo_thang)
        process4 = multiprocessing.Process(target=update_cham_cong_tong_hop)
        processes.append(process2)
        processes.append(process3)
        processes.append(process4)
        
        min = datetime.now().minute
        if min < 5:
            process5 = multiprocessing.Process(target=create_doanh_so_ca_nhan)
            process6 = multiprocessing.Process(target=create_report_co_so)
            process7 = multiprocessing.Process(target=create_all_report_khach_hang)            
            processes.append(process5)
            processes.append(process6)
            processes.append(process7)

        # process5 = multiprocessing.Process(target=create_doanh_so_ca_nhan)
        # process6 = multiprocessing.Process(target=create_report_co_so)
        # process7 = multiprocessing.Process(target=create_all_report_khach_hang)            
        # processes.append(process5)
        # processes.append(process6)
        # processes.append(process7)
        for i in range(len(vn_locations)):
            process_ngay = multiprocessing.Process(target=update_luy_ke_theo_ngay, args=(vn_locations[i],))
            process_thang = multiprocessing.Process(target=update_luy_ke_theo_thang, args=(vn_locations[i],))
            processes.append(process_ngay)
            processes.append(process_thang)

        for process in processes:
            process.start()

        for process in processes:
            process.join()

        print("All processes have finished.")
        print(f"Tổng thời gian một vòng lặp {(time.time() - start_time):.6f} giây\n")