import win32com.client
import os
from datetime import datetime
from bs4 import BeautifulSoup

def convert_excel_to_html_with_format(source_path, target_path, sheet_name="汇总"):
    # Mở ứng dụng Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Ẩn Excel trong quá trình chạy

    # Mở file Excel
    wb = excel.Workbooks.Open(source_path)

    try:
        # Lấy sheet theo tên
        ws = wb.Sheets(sheet_name)

        # Ẩn tất cả các sheet khác
        for sheet in wb.Sheets:
            if sheet.Name != sheet_name:
                sheet.Visible = False

        # Nếu thư mục chưa tồn tại thì tạo
        folder = os.path.dirname(target_path)
        if not os.path.exists(folder):
            os.makedirs(folder)

        # Lưu sheet thành file HTML
        ws.SaveAs(target_path, FileFormat=44)  # 44 = xlHtml

        print(f"✅ Đã chuyển đổi sheet '{ws.Name}' thành HTML tại: {target_path}")

    finally:
        # Hiển thị lại các sheet và đóng workbook
        for sheet in wb.Sheets:
            sheet.Visible = True
        wb.Close(False)
        excel.Quit()
def get_cell_colors_from_excel(excel_file, sheet_name):
    # Đọc màu sắc từ các ô trong Excel
    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Open(excel_file)
    ws = wb.Sheets(sheet_name)
    
    colors = {}
    for row in ws.UsedRange.Rows:
        for cell in row.Cells:
            # Lấy mã màu của ô
            color = cell.Interior.Color
            if isinstance(color, (int, float)):  # Kiểm tra nếu color là số (integer hoặc float)
                hex_color = f"#{int(color):06x}"  # Chuyển đổi giá trị màu thành hex
                colors[cell.Address] = hex_color

    wb.Close(False)
    excel.Quit()

    return colors

def add_colors_to_html(html_path, colors):
    # Đọc nội dung HTML và xử lý lại màu sắc nếu cần
    with open(html_path, 'r', encoding='utf-8') as f:
        html_content = f.read()

    # Dùng BeautifulSoup để phân tích và chỉnh sửa HTML
    soup = BeautifulSoup(html_content, 'html.parser')

    # Lặp qua tất cả các thẻ <td> để kiểm tra và thêm màu nền nếu có
    for td in soup.find_all('td'):
        # Tìm tọa độ ô (có thể bạn sẽ cần điều chỉnh cách lấy tọa độ này)
        cell_address = td.get('data-coordinate')  # Nếu có data-coordinate, bạn cần lưu thêm thông tin này vào HTML
        if cell_address and cell_address in colors:
            td['style'] = f'background-color: {colors[cell_address]};'

    # Lưu lại nội dung HTML sau khi chỉnh sửa
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(str(soup))

# === Phần cấu hình đường dẫn ===
timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
excel_file = rf"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report\FATP_WIP_Report_Detail_20250509103600.xlsx"
html_output_file = rf"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report\汇总_{timestamp}.html"
sheet_name = "汇总"

# === Chuyển đổi Excel thành HTML ===
convert_excel_to_html_with_format(excel_file, html_output_file, sheet_name)

# === Lấy màu sắc từ Excel ===
colors = get_cell_colors_from_excel(excel_file, sheet_name)

# === Thêm màu sắc vào HTML ===
add_colors_to_html(html_output_file, colors)

print("✅ Đã hoàn thành quá trình chuyển đổi và thêm màu sắc vào HTML.")
