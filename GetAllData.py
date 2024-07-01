from Utils import *
import requests
import json
import os
import multiprocessing

def fetch_all_notion_pages(database_name,database_id, notion_token):
    url = f"https://api.notion.com/v1/databases/{database_id}/query"
    headers = {
        "Authorization": f"Bearer {notion_token}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json"
    }
    
    all_results = []
    has_more = True
    next_cursor = None

    while has_more:
        payload = {"start_cursor": next_cursor} if next_cursor else {}
        
        response = requests.post(url, headers=headers, json=payload)
        data = response.json()
        
        if response.status_code != 200:
            raise Exception(f"Error: {response.status_code}, {response.text}")

        all_results.extend(data.get("results", []))
        has_more = data.get("has_more", False)
        next_cursor = data.get("next_cursor")

    save_to_json(f"{database_name}.json", all_results)
    return all_results

def save_to_json(file_name, json_data):
    if file_name:
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            
        file_path = os.path.join(output_folder,file_name)

        # Ghi dữ liệu JSON ra file
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, ensure_ascii=False, indent=4)

def get_all_data_to_json():
    processes = []
    for key, value in dataBaseDict.items():
        process = multiprocessing.Process(target=fetch_all_notion_pages, args=[key, value, notion_api_token])
        processes.append(process)
        process.start()
    for process in processes:
        process.join()
        print("Đã get all data in notion database")


