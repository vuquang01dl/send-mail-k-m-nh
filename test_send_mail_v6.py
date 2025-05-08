#nó chỉ gửi sheet 2 không gửi sheet 1
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

if not os.path.exists(excel_path):
    print("❌ File Excel không tồn tại.")
    exit()

# Mở Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
workbook = excel.Workbooks.Open(excel_path)

# Tạo danh sách các hình ảnh để đính kèm
image_paths = []

# Lặp qua tất cả các sheet và chụp hình
for sheet_index in range(1, workbook.Sheets.Count + 1):
    sheet = workbook.Sheets(sheet_index)

    # Kiểm tra nếu sheet có dữ liệu
    used_range = sheet.UsedRange
    if used_range.Rows.Count > 1 or used_range.Columns.Count > 1:  # Kiểm tra nếu có dữ liệu trong vùng

        # Nếu có dữ liệu, copy vùng dữ liệu thành hình ảnh
        used_range.CopyPicture(Format=2)  # 2 = Hình ảnh bitmap

        # Tạo một biểu đồ tạm để dán hình vào và export
        chart_object = sheet.ChartObjects().Add(100, 30, 500, 300)
        chart = chart_object.Chart
        chart.Paste()  # Dán hình vào biểu đồ
        time.sleep(3)  # Chờ dán hoàn tất

        # Lưu hình ảnh cho từng sheet
        image_path = os.path.join(
            r"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report",
            f"screenshot_{timeStart.strftime('%Y%m%d%H')}_sheet{sheet_index}.png"
        )
        # Sau chart.Export(image_path), thêm sleep + gọi DoEvents để đảm bảo dán hoàn tất
        chart.Export(image_path)
        time.sleep(1)
        excel.Wait(datetime.datetime.now() + datetime.timedelta(seconds=3))  # Cho Excel xử lý xong hoàn toàn
        # Xóa biểu đồ tạm
        chart_object.Delete()
    else:
        print(f"Sheet {sheet_index} không có dữ liệu hoặc không nhận diện được.")

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
    <p>Dưới đây là ảnh chụp các sheet trong file Excel:</p>
"""

# Đính kèm hình ảnh của từng sheet
for i, image_path in enumerate(image_paths, 1):
    attachment = mail.Attachments.Add(image_path)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", f"excel_img_{i}")
    html_body += f'<p><img src="cid:excel_img_{i}"></p>'

html_body += """
    <p>File Excel cũng được đính kèm.</p>
  </body>
</html>
"""

mail.HTMLBody = html_body
mail.Attachments.Add(excel_path)

mail.Send()
print("✅ Email đã được gửi kèm ảnh chụp từng sheet và file Excel.")
