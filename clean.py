import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

# Đọc file Excel
file_path = r"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report\2025050813.xlsx"
wb = openpyxl.load_workbook(file_path)
ws = wb.active

# Chuyển file Excel sang DataFrame
df = pd.read_excel(file_path)

# Loại bỏ các ô NaN, giữ nguyên các ô có giá trị
df_cleaned = df.dropna(how='all', axis=1)  # Loại bỏ các cột không có dữ liệu
df_cleaned = df_cleaned.dropna(how='all', axis=0)  # Loại bỏ các hàng không có dữ liệu

# Chuyển DataFrame đã làm sạch thành HTML hoặc định dạng bạn muốn
html_content = df_cleaned.to_html(index=False, header=True)

# Cập nhật lại sheet Excel với dữ liệu đã làm sạch
for row in dataframe_to_rows(df_cleaned, index=False, header=True):
    ws.append(row)

# Lưu lại file Excel
wb.save('cleaned_excel_output.xlsx')

print("Đã hoàn thành việc làm sạch dữ liệu và lưu lại vào file mới!")
