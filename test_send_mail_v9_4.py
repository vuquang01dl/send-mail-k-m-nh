import win32com.client
import datetime
import os
import logging
import time
import win32clipboard

# --- Cấu hình log ---
log_file = r"C:\Users\V5030587\Downloads\send-mail-k-m-nh\log_wip_report.txt"
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logging.info("=== Start executing the email sending script ===")

# --- Hàm đọc danh sách email từ file ---
def read_emails_from_file(filename):
    with open(filename, 'r') as file:
        return [email.strip() for email in file if email.strip()]

# --- Hàm lấy file Excel .xlsx mới nhất ---
def get_latest_excel(directory):
    files = [
        f for f in os.listdir(directory)
        if f.endswith('.xlsx') and not f.startswith('~$')
    ]
    if not files:
        raise FileNotFoundError("Không tìm thấy file Excel hợp lệ.")
    full_paths = [os.path.join(directory, f) for f in files]
    latest = max(full_paths, key=os.path.getmtime)
    logging.info(f"Found the latest Excel file: {latest}")
    return latest

# --- Đọc danh sách email ---
try:
    to_addrs = read_emails_from_file(r"C:\Users\V5030587\Downloads\send-mail-k-m-nh\recipients.txt")
    cc_addrs = read_emails_from_file(r"C:\Users\V5030587\Downloads\send-mail-k-m-nh\cc.txt")
    logging.info("Read recipient list and CC.")
except Exception as e:
    logging.error(f"Error reading email file: {e}")
    raise

# --- Xác định file Excel mới nhất ---
excel_dir = r"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report"
excel_path = get_latest_excel(excel_dir)
timeStart = datetime.datetime.now()
today_str = timeStart.strftime("%Y-%m-%d")
today_for_subject = timeStart.strftime("%Y%m%d")

# --- Mở Excel, copy nội dung sheet dưới dạng HTML ---
try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    workbook = excel.Workbooks.Open(excel_path, ReadOnly=True)
    sheet = workbook.Sheets(1)

    # Lấy tiêu đề sheet
    sheet_names = [sh.Name for sh in workbook.Sheets]
    sheet_titles = "&".join(sheet_names[:2]) if len(sheet_names) >= 2 else sheet_names[0]
    sheet_description = "以下为 PWYU & PWZJ WIP 状况."
    sheet_summary = f"{sheet_titles} WIP 状况："

    # Copy vùng dữ liệu đang dùng
    sheet.UsedRange.Copy()
    time.sleep(1)  # Chờ clipboard sẵn sàng

    # Lấy dữ liệu HTML từ clipboard
    win32clipboard.OpenClipboard()
    cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
    if win32clipboard.IsClipboardFormatAvailable(cf_html):
        html_content = win32clipboard.GetClipboardData(cf_html)
    else:
        raise ValueError("Clipboard không chứa dữ liệu HTML.")
    win32clipboard.CloseClipboard()

    workbook.Close(False)
    excel.Quit()
    logging.info("Copied Excel content as HTML successfully.")
except Exception as e:
    logging.error(f"Excel processing error: {e}")
    raise

# --- Soạn email sử dụng Outlook ---
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = Mail item

    mail.Subject = f"FATP WIP 分布状况 {today_for_subject}"
    mail.To = ";".join(to_addrs)
    mail.CC = ";".join(cc_addrs)

    cutoff_time = timeStart.strftime("%Y-%m-%d %H:%M")
    html_body = f"""
    <html>
    <body style="font-family:Calibri,sans-serif;font-size:12pt;">
        <p><b><span style="font-size:16pt;">Dear all:</span></b></p>
        <p>{sheet_description} Cut off time: <b style="color:red;">{cutoff_time}</b>, Thanks</p>
        <p style="color:blue;"><b>{sheet_summary}</b></p>
        {html_content}
        <p>Thanks & Best regards.</p>
    </body>
    </html>
    """

    mail.HTMLBody = html_body

    # Gửi mail
    mail.Send()
    logging.info("Email sent successfully using Outlook.")
except Exception as e:
    logging.error(f"Error sending email using Outlook: {e}")
    raise

logging.info("=== End of email sending script ===")
