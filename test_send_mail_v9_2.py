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

timeStart = datetime.datetime.now()

# Đường dẫn file Excel và ảnh
excel_path = os.path.join(
    r"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report",
    f"FATP_WIP_Report_Detail_{timeStart.strftime('%Y%m%d%H')}.xlsx"
)

image_path = os.path.join(
    r"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report",
    timeStart.strftime('screenshot_%Y%m%d%H') + ".png"
)

if not os.path.exists(excel_path):
    logging.error("❌ File Excel không tồn tại.")
    exit()
else:
    logging.info("✅ Đã tìm thấy file Excel.")

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
used_range.CopyPicture(Format=2)

left = used_range.Left
top = used_range.Top
width = used_range.Width
height = used_range.Height
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
    mail.Subject =  "FATP WIP 分布状况20250509-1"

    cutoff_time = timeStart.strftime("%Y-%m-%d %H:%M")

    html_body = f"""
      <html>
        <body style="font-family:Calibri,sans-serif;font-size:12pt;">
          <p><b><span style="font-size:16pt;">Dear all:</span></b></p>
          <p>{sheet_description} Cut off time: <b style="color:red;">{cutoff_time}, Thanks</b></p>
          <p style="color:blue;"><b>PWYU&PWZJ WIP 状况：</b></p>
          <img src="cid:excel_img">
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

    # Xóa ảnh sau khi gửi để không lưu
    if os.path.exists(image_path):
        os.remove(image_path)
        logging.info("🧹 Đã xóa ảnh sau khi gửi mail.")

except Exception as e:
    logging.error(f"❌ Lỗi: {e}")
