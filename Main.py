import multiprocessing.process
from Config import *
from GetNotionDataToJson import get_all_data_to_json
from CollectJsonToExcel import collect_data
from UpdateLuyKe import update_luy_ke_theo_ngay, update_luy_ke_theo_thang
from UpdateChamCong import update_cham_cong_tong_hop
from CreateReport import create_all_report
from UpdateKPI import update_KPI
import time
from datetime import datetime
import multiprocessing

def update_notion():
    update_notion_processes = []
    update_notion_processes.append(multiprocessing.Process(target=update_cham_cong_tong_hop))
    update_notion_processes.append(multiprocessing.Process(target=update_KPI))
    for location in location_list:
        update_notion_processes.append(multiprocessing.Process(target=update_luy_ke_theo_ngay, args=(location,)))
        update_notion_processes.append(multiprocessing.Process(target=update_luy_ke_theo_thang, args=(location,)))

    for process in update_notion_processes:
        process.start()
    for process in update_notion_processes:
        process.join()

def create_report():
    create_report_process = []
    for location in location_list:
        if location != "HỆ THỐNG":
            create_report_process.append(multiprocessing.Process(target=create_all_report, args=(location,)))

    for process in create_report_process:
        process.start()
    for process in create_report_process:
        process.join() 

    # Cần tạo các report tại các cơ sở trước rồi mới tạo report hệ thống
    process_report_he_thong = multiprocessing.Process(target=create_all_report, args=("HỆ THỐNG",))
    process_report_he_thong.start()
    process_report_he_thong.join()
    

if __name__ == "__main__":
    while(1):
        start_time = time.time()

        get_all_data_to_json()
        collect_data()
        print(f"Cập nhật toàn bộ data {(time.time() - start_time):.6f} giây\n")

        processes = []
        processes.append(multiprocessing.Process(target=update_notion))
        min = datetime.now().minute
        if (min < 7):
            processes.append(multiprocessing.Process(target=create_report))

        for process in processes:
            process.start()
        for process in processes:
            process.join() 

        print("All processes have finished.")
        print(f"Tổng thời gian một vòng lặp {(time.time() - start_time):.6f} giây\n")