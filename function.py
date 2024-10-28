import os
import re
from datetime import datetime

class get_API_key:
    def __init__(self, filename, n):
        self.filename = filename
        self.n = n
        
    def get_api_key(self, line_number):
        with open(self.filename, 'r') as file:
            lines = file.readlines()
            if 1 <= line_number <= len(lines):
                api_key = lines[line_number - 1].strip()
                return api_key
            else:
                print(f"Error: Line {line_number} does not exist in the file.")
                return None

class find_file:
    def extract_date_from_filename(filename):
    # 파일 이름에서 날짜 패턴을 추출 (예: 'ebay_products_2024-10-15.xlsx')
        match = re.search(r'Scotty Products_(\d{4}-\d{2}-\d{2})\.xlsx', filename)
        if match:
            return datetime.strptime(match.group(1), '%Y-%m-%d')
        return None

# ebay_products_yyyy-mm-dd.xlsx 파일 중 가장 최근 파일 찾기
    def find_latest_file():
        output_dir = './output'
        files = os.listdir(output_dir)
        excel_files = [f for f in files if f.startswith('Scotty Products_') and f.endswith('.xlsx')]
        
        latest_file = None
        latest_date = None
        
        for file in excel_files:
            file_date = find_file.extract_date_from_filename(file)
            if file_date:
                if latest_date is None or file_date > latest_date:
                    latest_date = file_date
                    latest_file = file

        return latest_file
