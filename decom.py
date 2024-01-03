import openpyxl
from datetime import datetime
import os

folder_path = r'C:\Folder path'

search_text = 'Text for search'
search_column = 'A'
current_date = datetime.now().strftime('%m/%d/%Y')

for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(folder_path,filename)
        
        wb=openpyxl.load_workbook(file_path)
        sheet = wb.active

        for row in range(1, sheet.max_row +1):
            cel_value = sheet[f'{search_column}{row}'].value
            existing_data = sheet[f'D{row}'].value
            
            if cel_value == search_text and not existing_data:
                sheet[f'D{row}'] = 'TEXT'
                sheet[f'E{row}'] = 'TEXT'
                sheet[f'F{row}'] = current_date
                wb.save(file_path)
                print(f"Updated file: {filename}")
print("Completed")
