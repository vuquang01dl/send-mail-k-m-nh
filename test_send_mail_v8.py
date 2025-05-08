#send mail và ảnh ok, nhưng giao diện chưa được đẹp lắm cần sửa lại
import win32com.client
import datetime
import os
import time

def read_emails_from_file(filename):
    with open(filename, 'r') as file:
        emails = file.readlines()
    return [email.strip() for email in emails if email.strip()]

timeStart = datetime.datetime.now()

# Đường dẫn file Excel và ảnh
excel_path = os.path.join(
    r"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report",
    timeStart.strftime('%Y%m%d%H') + ".xlsx"
)

image_path = os.path.join(
    r"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report",
    timeStart.strftime('screenshot_%Y%m%d%H') + ".png"
)

if not os.path.exists(excel_path):
    print("❌ File Excel không tồn tại.")
    exit()

# Mở Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
workbook = excel.Workbooks.Open(excel_path, ReadOnly=True)
sheet = workbook.Sheets(1)

# Lấy tên các sheet trước khi workbook bị đóng
sheet_names = [sh.Name for sh in workbook.Sheets]
sheet_titles = "&".join(sheet_names[:2]) if len(sheet_names) >= 2 else sheet_names[0]
sheet_description = f"以下为 {sheet_titles} WIP 状况."
sheet_summary = f"{sheet_titles} WIP 状况："

# Copy vùng dữ liệu thành hình ảnh
used_range = sheet.UsedRange
used_range.CopyPicture(Format=2)

# Tính kích thước và dán vào biểu đồ
left = used_range.Left
top = used_range.Top
width = used_range.Width
height = used_range.Height
scale_factor = 1.3

chart_object = sheet.ChartObjects().Add(
    left, top,
    width * scale_factor,
    height * scale_factor
)
chart = chart_object.Chart
chart.Paste()
time.sleep(1)  # Chờ dán

# Xuất ảnh và xóa biểu đồ tạm
chart.Export(image_path)
chart_object.Delete()

# Đóng Excel
workbook.Close(False)
excel.Quit()

# Gửi email
outlook = win32com.client.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)
mail.To = "; ".join(read_emails_from_file("recipients.txt"))
mail.CC = "; ".join(read_emails_from_file("cc.txt"))
mail.Subject = f"📸 Báo cáo hình ảnh lúc {timeStart.strftime('%H:00 %d/%m/%Y')}"

cutoff_time = timeStart.strftime("%Y-%m-%d %H:%M")

html_body = f"""
<html>
  <body>
    <p style="font-size:10pt;font-family:Calibri,sans-serif;margin:0;">
      <b><span style="font-size:16pt;font-family:Times;" lang="en-US">Dear all:</span></b>
    </p>

    <p style="font-size:10pt;font-family:Calibri,sans-serif;margin:0;">&nbsp;</p>

    <p style="font-size:10pt;font-family:Calibri,sans-serif;margin:0;">
      <span style="font-size:14pt;font-family:Times;" lang="en-US">
        {sheet_description} Cut off time: 
        <span style="color:black;">{cutoff_time.split()[0]}</span> 
        <b style="color:red;font-size:16pt;">{cutoff_time.split()[1]}</b> 
        <span style="font-size:14pt;font-family:Times;color:black;">Thanks!</span>
      </span>
    </p>

    <p style="font-size:10pt;font-family:Calibri,sans-serif;margin:0;">&nbsp;</p>

    <p style="margin:0cm; margin-bottom:.0001pt">
      <b><span style="font-size: 18pt; font-family: Arial, sans-serif, serif, EmojiFont; color: blue;" lang="VI">
        ({sheet_titles}) {sheet_summary}
      </span></b>
    </p>

    <img src="cid:excel_img">
  </body>
</html>
"""

mail.HTMLBody = html_body
mail.Attachments.Add(excel_path)
attachment = mail.Attachments.Add(image_path)
attachment.PropertyAccessor.SetProperty(
    "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "excel_img"
)

mail.Send()
print("✅ Email đã được gửi kèm ảnh chụp nội dung Excel và file Excel.")
