from Config import *        
    
def update_cham_cong_tong_hop():
    data_cham_cong_tong_hop = get_data_cham_cong("HỆ THỐNG", ["ALL"])
    data_cham_cong_tong_hop = data_cham_cong_tong_hop.fillna(0)
    for col in cham_cong_ref.keys():
        data_cham_cong_tong_hop[col] = 0
    for id_nhan_su in data_cham_cong_tong_hop["id nhân sự"]:
        for location in location_list:
            if location != "HỆ THỐNG":
                data = get_data_cham_cong(location, ["ALL"])
                row = data.loc[data["id nhân sự"] == id_nhan_su]
                for col in row.columns:
                    value = row[col].values[0]
                    if isinstance(value, str):
                        if value in cham_cong_ref.keys():
                            data_cham_cong_tong_hop.loc[data_cham_cong_tong_hop["id nhân sự"] == id_nhan_su, value] += 1

    json_file_path = os.path.join(notion_data_folder, 'Chấm công HỆ THỐNG.json')

    # Đọc tệp JSON và chuyển đổi nội dung thành chuỗi JSON
    with open(json_file_path, 'r', encoding='utf-8') as file:
        json_data = json.load(file)   
    
    for item in json_data:
        page_id = item["id"]
        id_nhan_su_update = item["properties"]["Nhân sự"]["relation"][0]["id"]
        row = data_cham_cong_tong_hop[data_cham_cong_tong_hop["id nhân sự"] == id_nhan_su_update]
        tong_cong = 0
        for key, value in cham_cong_ref.items():
            tong_cong = tong_cong + row.iloc[0][key]*value
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
                            "Tổng công": {
                                "number": float(tong_cong)
                            },
                            "Nghỉ có phép": {
                                "number": int(row.iloc[0]["Nghỉ có phép"])
                            },
                        },
                    }
        update_page(page_id, template_json)
    print("Đã update bảng chấm công HỆ THỐNG!")

# update_cham_cong_tong_hop()