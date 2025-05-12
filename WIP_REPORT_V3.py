#rất ok nhưng chưa lấy file mới nhất trong list thay vì file excel cụ thể, cấu hình đường dẫn file cần đưa lên đầu 

import win32com.client
import os
from datetime import datetime

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

def send_email_with_html_content(html_main_path, recipient):
    html_dir = html_main_path.replace(".html", "_files")
    sheet_file = os.path.join(html_dir, "sheet001.html")
    css_file = os.path.join(html_dir, "stylesheet.css")

    if not os.path.exists(sheet_file):
        print(f"❌ Không tìm thấy file: {sheet_file}")
        return

    # Đọc CSS và nội dung HTML chính
    css_content = ""
    if os.path.exists(css_file):
        with open(css_file, 'r', encoding='utf-8', errors='replace') as css:
            css_content = f"<style>{css.read()}</style>"

    with open(sheet_file, 'r', encoding='utf-8', errors='replace') as file:
        html_body = file.read()

    # Gửi email
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = "FATP WIP Report"
    mail.To = recipient
    mail.HTMLBody = f"""
    <html>
        <head>
            {css_content}
        </head>
        <body>
            <p>Dear {recipient},</p>
            <p>Please find the WIP report below:</p>
            {html_body}
        </body>
    </html>
    """
    mail.Send()
    print(f"✅ Đã gửi email đến {recipient}.")

# === Cấu hình ===
timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
excel_file = r"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report\FATP_WIP_Report_Detail_20250509103600.xlsx"
html_output_file = fr"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report\汇总_{timestamp}.html"
sheet_name = "汇总"
recipient_email = "lucas.vu@quantacn.com"

# === Thực thi ===
convert_excel_to_html_with_format(excel_file, html_output_file, sheet_name)
send_email_with_html_content(html_output_file, recipient_email)

print("✅ Hoàn tất quá trình gửi email.")
