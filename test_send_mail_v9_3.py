import win32com.client
import datetime
import os
import time
import logging

# Cấu hình log
log_file = os.path.join(os.getcwd(), "email_log.txt")
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='[%(asctime)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def read_emails_from_file(filename):
    with open(filename, 'r') as file:
        emails = file.readlines()
    return [email.strip() for email in emails if email.strip()]

def get_latest_excel_file(folder, prefix="FATP_WIP_Report_Detail", ext=".xlsx"):
    files = [f for f in os.listdir(folder) if f.startswith(prefix) and f.endswith(ext)]
    if not files:
        return None
    files = sorted(files, key=lambda f: os.path.getmtime(os.path.join(folder, f)), reverse=True)
    return os.path.join(folder, files[0])

def count_reports_sent_today(log_path, date_str):
    if not os.path.exists(log_path):
        return 0
    count = 0
    with open(log_path, "r") as f:
        for line in f:
            if f"[{date_str}" in line and "Đã gửi email thành công." in line:
                count += 1
    return count

timeStart = datetime.datetime.now()
today_str = timeStart.strftime("%Y-%m-%d")
today_for_subject = timeStart.strftime("%Y%m%d")

# Thư mục chứa Excel
excel_dir = r"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report"
excel_path = get_latest_excel_file(excel_dir)

if not excel_path or not os.path.exists(excel_path):
    logging.error("❌ Không tìm thấy file Excel mới nhất.")
    exit()
else:
    logging.info(f"✅ Đã tìm thấy file Excel: {excel_path}")

# Tạo tên ảnh từ thời gian hiện tại
image_path = os.path.join(
    excel_dir,
    timeStart.strftime('screenshot_%Y%m%d%H%M%S') + ".png"
)

# Mở Excel và tạo ảnh
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
workbook = excel.Workbooks.Open(excel_path, ReadOnly=True)
sheet = workbook.Sheets(1)

sheet_names = [sh.Name for sh in workbook.Sheets]
sheet_titles = "&".join(sheet_names[:2]) if len(sheet_names) >= 2 else sheet_names[0]
sheet_description = "以下为 PWYU & PWZJ WIP 状况."
sheet_summary = f"{sheet_titles} WIP 状况："

used_range = sheet.UsedRange

# Lấy số dòng và cột cuối
last_row = used_range.Row + used_range.Rows.Count - 1
last_col = used_range.Column + used_range.Columns.Count - 1

# Vùng cần chụp: từ A1 đến ô cuối cùng được dùng
range_with_border = sheet.Range("A1", sheet.Cells(last_row, last_col))
range_with_border.CopyPicture(Format=2)

left = range_with_border.Left
top = range_with_border.Top
width = range_with_border.Width
height = range_with_border.Height
scale_factor = 1.3

chart_object = sheet.ChartObjects().Add(left, top, width * scale_factor, height * scale_factor)
chart = chart_object.Chart
chart.Paste()
time.sleep(1)

chart.Export(image_path)
chart_object.Delete()
workbook.Close(False)
excel.Quit()

logging.info("✅ Đã tạo ảnh từ Excel.")

# Gửi email
try:
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = "; ".join(read_emails_from_file("recipients.txt"))
    mail.CC = "; ".join(read_emails_from_file("cc.txt"))

    report_count = count_reports_sent_today(log_file, today_str) + 1
    mail.Subject = f"FATP WIP 分布状况{today_for_subject}"

    cutoff_time = timeStart.strftime("%Y-%m-%d %H:%M")

    html_body = f"""
      <html>
        <body style="font-family:Calibri,sans-serif;font-size:12pt;">
          <p><b><span style="font-size:16pt;">Dear all:</span></b></p>
          <p>{sheet_description} Cut off time: <b style="color:red;">{cutoff_time}</b>, Thanks</p>
          <p style="color:blue;"><b>PWYU&PWZJ WIP 状况：</b></p>
          <img src="cid:excel_img" style="display:block; margin-left:auto; margin-right:auto;">
          <p>Thanks & Best regards.</p>
        </body>
      </html>
      """
    mail.HTMLBody = html_body

    if os.path.exists(excel_path):
        mail.Attachments.Add(excel_path)
    else:
        logging.warning(f"⚠️ Không tìm thấy file Excel để đính kèm: {excel_path}")

    if os.path.exists(image_path):
        attachment = mail.Attachments.Add(image_path)
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "excel_img"
        )
    else:
        logging.warning(f"⚠️ Không tìm thấy file ảnh để đính kèm: {image_path}")

    mail.Send()
    logging.info("✅ Đã gửi email thành công.")

    # Xóa ảnh sau khi gửi
    if os.path.exists(image_path):
        os.remove(image_path)
        logging.info("🧹 Đã xóa ảnh sau khi gửi mail.")

except Exception as e:
    logging.error(f"❌ Lỗi: {e}")
