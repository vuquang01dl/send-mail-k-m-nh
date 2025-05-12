import smtplib
import win32com.client
import datetime
import os
import io
import logging
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

# --- Cấu hình log ---
log_file = r"E:\Report\WIP_ReportDetail\log_wip_report.txt"
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logging.info("=== Start executing the email sending script ===")

def read_emails_from_file(filename):
    with open(filename, 'r') as file:
        return [email.strip() for email in file if email.strip()]

def get_latest_excel(directory):
    files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    full_paths = [os.path.join(directory, f) for f in files]
    latest = max(full_paths, key=os.path.getmtime)
    logging.info(f"Found the latest Excel file: {latest}")
    return latest

# --- Cấu hình SMTP ---
smtp_server = "172.20.179.74"
smtp_port = 25
smtp_user = "QMS_FA1_Notice"
smtp_password = "Quanta123"
from_addr = "QMS_FA1_Notice@quantacn.com"

# --- Đọc danh sách email ---
try:
    to_addrs = read_emails_from_file(r"E:\Report\WIP_ReportDetail\recipients.txt")
    cc_addrs = read_emails_from_file(r"E:\Report\WIP_ReportDetail\cc.txt")
    logging.info("Read recipient list and CC.")
except Exception as e:
    logging.error(f"Error reading email file: {e}")
    raise

# --- Lấy file Excel mới nhất ---
excel_dir = r"E:\Report\WIP_ReportDetail\Report"
excel_path = get_latest_excel(excel_dir)
timeStart = datetime.datetime.now()
today_str = timeStart.strftime("%Y-%m-%d")
today_for_subject = timeStart.strftime("%Y%m%d")

# --- Mở Excel và xuất ảnh ---
image_stream = io.BytesIO()
image_path_temp = os.path.join(excel_dir, "temp_screenshot.png")

try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    workbook = excel.Workbooks.Open(excel_path, ReadOnly=True)
    sheet = workbook.Sheets(1)

    # Lấy tên sheet để viết tiêu đề
    sheet_names = [sh.Name for sh in workbook.Sheets]
    sheet_titles = "&".join(sheet_names[:2]) if len(sheet_names) >= 2 else sheet_names[0]
    sheet_description = "以下为 PWYU & PWZJ WIP 状况."
    sheet_summary = f"{sheet_titles} WIP 状况："

    # Chụp ảnh vùng dữ liệu
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

    # Export ảnh
    if chart.Export(image_path_temp):
        logging.info(f"Chart image exported successfully: {image_path_temp}")
        if not os.path.exists(image_path_temp):
            raise FileNotFoundError(f"Temporary image file not found: {image_path_temp}")
        with open(image_path_temp, "rb") as f:
            image_stream.write(f.read())
        if image_stream.tell() == 0:
            raise ValueError("The image in image_stream is empty.")
        os.remove(image_path_temp)
    else:
        raise Exception("Unable to export chart to image.")

    chart_object.Delete()
    workbook.Close(False)
    excel.Quit()
    logging.info("Created and processed images from Excel files.")
except Exception as e:
    logging.error(f"Excel processing error: {e}")
    raise

# --- Soạn email ---
msg = MIMEMultipart("related")
msg["From"] = from_addr
msg["To"] = ", ".join(to_addrs)
msg["Cc"] = ", ".join(cc_addrs)
msg["Subject"] = f"FATP WIP 分布状况{today_for_subject}"

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

msg_alternative = MIMEMultipart("alternative")
msg.attach(msg_alternative)
msg_alternative.attach(MIMEText(html_body, "html", "utf-8"))

image_stream.seek(0)
img = MIMEImage(image_stream.read())
img.add_header("Content-ID", "<excel_img>")
msg.attach(img)

# Đính kèm file Excel
with open(excel_path, "rb") as f:
    part = MIMEApplication(f.read(), Name=os.path.basename(excel_path))
    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(excel_path)}"'
    msg.attach(part)

# --- Gửi email ---
try:
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.ehlo()
        server.login(smtp_user, smtp_password)
        server.sendmail(from_addr, to_addrs + cc_addrs, msg.as_string())
    logging.info("Send mail successfull")
except smtplib.SMTPException as e:
    logging.error(f"Error send email (SMTP): {e}")
    raise
except Exception as e:
    logging.error(f"Unspecified error while sending email: {e}")
    raise

logging.info("=== End of email sending script ===")
