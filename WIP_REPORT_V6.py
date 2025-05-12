# đã ok, cần chuyển sang mail của công ti 

import os
import win32com.client
import time
from datetime import datetime
timeStart = datetime.now()


# ======================
# 🛠 CẤU HÌNH ĐƯỜNG DẪN
# ======================
EXCEL_FOLDER = r"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report"
SHEET_NAME = "汇总"
HTML_EXPORT = r"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report\HTML_EXPORT"  # Thư mục xuất HTML
sheet_description = "以下为 PWYU & PWZJ WIP 状况."


# 🔎 TÌM FILE EXCEL MỚI NHẤT
# ======================
def get_latest_excel_file(folder_path):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
    if not excel_files:
        raise FileNotFoundError("❌ Không tìm thấy file .xlsx nào trong thư mục.")

    excel_files = sorted(excel_files, key=lambda f: os.path.getmtime(os.path.join(folder_path, f)), reverse=True)
    return os.path.join(folder_path, excel_files[0])

# ======================
# 🔎 ĐỌC EMAIL TỪ FILE
# ======================
def read_emails_from_file(filename):
    with open(filename, 'r') as file:
        emails = file.readlines()
    return [email.strip() for email in emails if email.strip()]

# ======================
# 📤 CHUYỂN FILE EXCEL SANG HTML
# ======================
def convert_excel_to_html_with_format(source_path, target_path, sheet_name="汇总"):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(source_path)

    try:
        ws = wb.Sheets(sheet_name)
        for sheet in wb.Sheets:
            if sheet.Name != sheet_name:
                sheet.Visible = False

        folder = os.path.dirname(target_path)
        if not os.path.exists(folder):
            os.makedirs(folder)

        ws.SaveAs(target_path, FileFormat=44)  # xlHtml
        print(f"✅ Đã chuyển đổi sheet '{ws.Name}' thành HTML tại: {target_path}")
    finally:
        for sheet in wb.Sheets:
            sheet.Visible = True
        wb.Close(False)
        excel.Quit()

# ======================
# 📧 GỬI EMAIL CÓ NỘI DUNG HTML
# ======================
def send_email_with_html_content(html_main_path, recipients, cc_recipients):
    html_dir = html_main_path.replace(".html", "_files")
    sheet_file = os.path.join(html_dir, "sheet001.html")
    css_file = os.path.join(html_dir, "stylesheet.css")

    if not os.path.exists(sheet_file):
        print(f"❌ Không tìm thấy file: {sheet_file}")
        return

    # Đọc CSS và HTML chính
    css_content = ""
    if os.path.exists(css_file):
        with open(css_file, 'r', encoding='utf-8', errors='replace') as css:
            css_content = f"<style>{css.read()}</style>"

    with open(sheet_file, 'r', encoding='utf-8', errors='replace') as file:
        html_body = file.read()

    # Gửi email cho danh sách người nhận
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = "FATP WIP Report"
    
    # Tạo chuỗi email người nhận
    mail.To = "; ".join(recipients)  # Nhiều người nhận được phân cách bởi dấu chấm phẩy
    mail.CC = "; ".join(cc_recipients)  # CC người nhận
    cutoff_time = timeStart.strftime("%Y-%m-%d %H:%M")
    mail.HTMLBody = f"""
        <html>
            <head>{css_content}</head>
            <body>
                <p><b><span style="font-size:16pt;">Dear all:</span></b></p>
                <p>{sheet_description} Cut off time: <b style="color:red;">{cutoff_time}</b>, Thanks</p>
                <p style="color:blue;"><b>PWYU&PWZJ WIP 状况：</b></p>
                {html_body}
            </body>
        </html>
        """
    mail.Send()
    print(f"✅ Đã gửi email đến {', '.join(recipients)} và CC tới {', '.join(cc_recipients)}.")

# ======================
# 🚀 THỰC THI TOÀN BỘ
# ======================
try:
    latest_excel_file = get_latest_excel_file(EXCEL_FOLDER)
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    
    # Tạo thư mục xuất HTML nếu chưa có
    if not os.path.exists(HTML_EXPORT):
        os.makedirs(HTML_EXPORT)
    
    html_output_file = os.path.join(HTML_EXPORT, f"{SHEET_NAME}_{timestamp}.html")

    print(f"📄 Đang xử lý file Excel mới nhất: {latest_excel_file}")
    convert_excel_to_html_with_format(latest_excel_file, html_output_file, SHEET_NAME)
    
    # Đọc danh sách email từ file
    recipients = read_emails_from_file("recipients.txt")
    cc_recipients = read_emails_from_file("cc.txt")
    
    send_email_with_html_content(html_output_file, recipients, cc_recipients)

    print("✅ Hoàn tất quá trình gửi email.")
except Exception as e:
    print(f"❌ Lỗi: {e}")
