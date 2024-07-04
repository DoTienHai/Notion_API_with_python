import multiprocessing.process
from Utils import *
from GetNotionDataToJson import get_all_data_to_json
from CollectJsonToExcel import collect_data
from CreateLuyKe import update_luy_ke_theo_ngay, update_luy_ke_theo_thang
from CreateReportCaNhan import create_doanh_so_ca_nhan
from CreateReportCoSo import create_report_co_so
from CreateReportKhachHang import create_all_report_khach_hang
import time
import multiprocessing

def fetch_data():
    get_all_data_to_json()
    collect_data()

if __name__ == "__main__":
    # while(1):
        start_time = time.time()
        process1 = multiprocessing.Process(target=fetch_data)
        process1.start()
        process1.join()
        print(f"Cập nhật toàn bộ data {(time.time() - start_time):.6f} giây\n")
        
        processes = []
        # process2 = multiprocessing.Process(target=update_luy_ke_theo_ngay)
        # process3 = multiprocessing.Process(target=update_luy_ke_theo_thang)
        # process4 = multiprocessing.Process(target=create_doanh_so_ca_nhan)
        process5 = multiprocessing.Process(target=create_report_co_so)
        process6 = multiprocessing.Process(target=create_all_report_khach_hang)            

        # processes.append(process2)
        # processes.append(process3)
        # processes.append(process4)
        processes.append(process5)
        processes.append(process6)

        # for i in range(len(vn_locations)):
        #     process_ngay = multiprocessing.Process(target=update_luy_ke_theo_ngay, args=(vn_locations[i],))
        #     process_thang = multiprocessing.Process(target=update_luy_ke_theo_thang, args=(vn_locations[i],))
        #     processes.append(process_ngay)
        #     processes.append(process_thang)

        for process in processes:
            process.start()

        for process in processes:
            process.join()

        print("All processes have finished.")
        print(f"Tổng thời gian một vòng lặp {(time.time() - start_time):.6f} giây\n")