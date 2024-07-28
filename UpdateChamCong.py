from Config import *        
    
def update_cham_cong_tong_hop():
    data_cham_cong_tong_hop = get_data_cham_cong("HỆ THỐNG", ["ALL"])
    data_cham_cong_tong_hop = data_cham_cong_tong_hop.fillna(0)
    # Tính Tổng công
    for col in cham_cong_ref.keys():
        data_cham_cong_tong_hop[col] = 0
    for id_nhan_su in data_cham_cong_tong_hop["id nhân sự"]:
        tong_cong = {
            "CẦN THƠ" : 0,
            "LONG XUYÊN" : 0,
            "SÓC TRĂNG" : 0
        }
        for location in location_list:
            if location != "HỆ THỐNG":
                data = get_data_cham_cong(location, ["ALL"])
                row = data.loc[data["id nhân sự"] == id_nhan_su]
                for col in row.columns:
                    value = row[col].values[0]
                    if isinstance(value, str):
                        if value in cham_cong_ref.keys():
                            data_cham_cong_tong_hop.loc[data_cham_cong_tong_hop["id nhân sự"] == id_nhan_su, value] += 1
                            tong_cong[location] += cham_cong_ref[value]
        
                data_cham_cong_tong_hop.loc[data_cham_cong_tong_hop["id nhân sự"] == id_nhan_su, f"Tổng công tại {location}"] = tong_cong[location]
        

    json_file_path = os.path.join(notion_data_folder, 'Chấm công HỆ THỐNG.json')

    # Đọc tệp JSON và chuyển đổi nội dung thành chuỗi JSON
    with open(json_file_path, 'r', encoding='utf-8') as file:
        json_data = json.load(file)   
    
    for item in json_data:
        page_id = item["id"]
        id_nhan_su_update = item["properties"]["Nhân sự"]["relation"][0]["id"]
        row = data_cham_cong_tong_hop[data_cham_cong_tong_hop["id nhân sự"] == id_nhan_su_update]
        # tong_cong = 0
        # for key, value in cham_cong_ref.items():
        #     tong_cong = tong_cong + row.iloc[0][key]*value
        template_json ={
                        "properties": {
                            "Nửa ngày": {
                                "number": int(row.iloc[0]["Nửa ngày"])
                            },
                            "Nghỉ không phép": {
                                "number": int(row.iloc[0]["Nghỉ không phép"])
                            },
                            "Đầy đủ": {
                                "number": int(row.iloc[0]["Đầy đủ"])
                            },
                            "Nghỉ có phép": {
                                "number": int(row.iloc[0]["Nghỉ có phép"])
                            },
                            "Tổng công tại CẦN THƠ": {
                                "number": float(row.iloc[0]["Tổng công tại CẦN THƠ"])
                            },
                            "Tổng công tại LONG XUYÊN": {
                                "number": float(row.iloc[0]["Tổng công tại LONG XUYÊN"])
                            },
                            "Tổng công tại SÓC TRĂNG": {
                                "number": float(row.iloc[0]["Tổng công tại SÓC TRĂNG"])
                            },
                        },
                    }
        update_page(page_id, template_json)
    print("Đã update bảng chấm công HỆ THỐNG!")

# update_cham_cong_tong_hop()