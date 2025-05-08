import smtplib
import win32com.client
import datetime
import os
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

def read_emails_from_file(filename):
    with open(filename, 'r') as file:
        return [email.strip() for email in file if email.strip()]

def get_latest_excel(directory):
    files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    full_paths = [os.path.join(directory, f) for f in files]
    return max(full_paths, key=os.path.getmtime)

# Cấu hình thư
smtp_server = "172.20.179.74"
smtp_port = 25
from_addr = "QMS_FA1_Notice@quantacn.com"

to_addrs = read_emails_from_file(r"E:\Report\WIP_ReportDetail\Report\recipients.txt")
cc_addrs = read_emails_from_file(r"E:\Report\WIP_ReportDetail\Report\cc.txt")

# Tìm file Excel mới nhất
excel_dir = r"E:\Report\WIP_ReportDetail\Report"
excel_path = get_latest_excel(excel_dir)
timeStart = datetime.datetime.now()

# Mở Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
workbook = excel.Workbooks.Open(excel_path, ReadOnly=True)
sheet = workbook.Sheets(1)

sheet_names = [sh.Name for sh in workbook.Sheets]
sheet_titles = "&".join(sheet_names[:2]) if len(sheet_names) >= 2 else sheet_names[0]
sheet_description = f"以下为 {sheet_titles} WIP 状况."
sheet_summary = f"{sheet_titles} WIP 状况："

# Copy vùng dữ liệu thành hình ảnh trong bộ nhớ
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

# Lưu ảnh vào bộ nhớ (stream)
image_stream = io.BytesIO()
image_path_temp = os.path.join(excel_dir, "temp_screenshot.png")
chart.Export(image_path_temp)
with open(image_path_temp, "rb") as f:
    image_stream.write(f.read())
os.remove(image_path_temp)
chart_object.Delete()
workbook.Close(False)
excel.Quit()

# Soạn email
msg = MIMEMultipart("related")
msg["From"] = from_addr
msg["To"] = ", ".join(to_addrs)
msg["Cc"] = ", ".join(cc_addrs)
msg["Subject"] = "QMH FATP OCT Performance_" + datetime.datetime.now().strftime('%Y-%m-%d')

cutoff_time = timeStart.strftime("%Y-%m-%d %H:%M")
html_body = f"""
<html>
  <body style="font-family:Calibri,sans-serif;font-size:12pt;">
    <p><b><span style="font-size:16pt;">Dear all:</span></b></p>
    <p>{sheet_description}</p>
    <p>Cut off time: <b style="color:red;">{cutoff_time}</b></p>
    <p style="color:blue;"><b>({sheet_titles}) {sheet_summary}</b></p>
    <img src="cid:excel_img">
    <p>Thanks & Best regards.</p>
  </body>
</html>
"""

# Gắn HTML và hình ảnh
msg_alternative = MIMEMultipart("alternative")
msg.attach(msg_alternative)
msg_alternative.attach(MIMEText(html_body, "html", "utf-8"))

image_stream.seek(0)
img = MIMEImage(image_stream.read())
img.add_header("Content-ID", "<excel_img>")
msg.attach(img)

# Gắn file Excel
with open(excel_path, "rb") as f:
    part = MIMEApplication(f.read(), Name=os.path.basename(excel_path))
    part['Content-Disposition'] = f'attachment; filename="{os.path.basename(excel_path)}"'
    msg.attach(part)

# Gửi email
with smtplib.SMTP(smtp_server, smtp_port) as server:
    server.sendmail(from_addr, to_addrs + cc_addrs, msg.as_string())

print("Gửi email thành công.")
