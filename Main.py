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
        process2 = multiprocessing.Process(target=update_cham_cong_tong_hop)
        process3 = multiprocessing.Process(target=update_KPI)
        processes.append(process2)
        processes.append(process3)
        
        min = datetime.now().minute

        for location in location_list:
            process_ngay = multiprocessing.Process(target=update_luy_ke_theo_ngay, args=(location,))
            process_thang = multiprocessing.Process(target=update_luy_ke_theo_thang, args=(location,))
            processes.append(process_ngay)
            processes.append(process_thang)
            # if (min < 5):
            #     process_create_report = multiprocessing.Process(target=create_all_report, args=(location,))
            #     processes.append(process_create_report)
            process_create_report = multiprocessing.Process(target=create_all_report, args=(location,))
            processes.append(process_create_report)

        for process in processes:
            process.start()

        for process in processes:
            process.join()

        print("All processes have finished.")
        print(f"Tổng thời gian một vòng lặp {(time.time() - start_time):.6f} giây\n")