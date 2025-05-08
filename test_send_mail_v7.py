#gửi mail với kích thước ảnh ok nhưng chưa được tối ưu với mọi loại thư
import win32com.client
import datetime
import os
import time

def read_emails_from_file(filename):
    with open(filename, 'r') as file:
        emails = file.readlines()
    return [email.strip() for email in emails if email.strip()]

timeStart = datetime.datetime.now()

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
workbook = excel.Workbooks.Open(excel_path)
sheet = workbook.Sheets(1)

# Copy vùng dữ liệu thành hình ảnh
used_range = sheet.UsedRange
used_range.CopyPicture(Format=2)  # 2 = Hình ảnh bitmap

# Tạo một biểu đồ tạm để dán hình vào và export
chart_object = sheet.ChartObjects().Add(100, 30, 1000, 800)  # Tăng kích thước biểu đồ
chart = chart_object.Chart
chart.Paste()  # Dán hình vào biểu đồ
time.sleep(1)  # Chờ dán hoàn tất

# Xuất hình ảnh
chart.Export(image_path)

# Xóa biểu đồ tạm
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

html_body = f"""
<html>
  <body>
    <p>Chào anh/chị,</p>
    <p>Dưới đây là ảnh chụp bảng Excel:</p>
    <img src="cid:excel_img">
    <p>File Excel cũng được đính kèm.</p>
  </body>
</html>
"""

mail.HTMLBody = html_body
mail.Attachments.Add(excel_path)
attachment = mail.Attachments.Add(image_path)
attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "excel_img")

mail.Send()
print("✅ Email đã được gửi kèm ảnh chụp nội dung Excel và file Excel.")
