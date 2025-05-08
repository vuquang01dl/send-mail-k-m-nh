import win32com.client
import datetime
import os

# Hàm đọc danh sách email từ tệp
def read_emails_from_file(filename):
    with open(filename, 'r') as file:
        emails = file.readlines()
    return [email.strip() for email in emails if email.strip()]

# Thời gian dùng để tạo tên file
timeStart = datetime.datetime.now()

# Tạo đường dẫn tới file Excel
file_path = os.path.join(
    r"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report",
    timeStart.strftime('%Y%m%d%H') + ".xlsx"
)

# Kiểm tra file có tồn tại không
if not os.path.exists(file_path):
    print(f"File không tồn tại: {file_path}")
else:
    # Khởi tạo đối tượng Outlook
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)

    # Đọc danh sách người nhận và CC
    to_emails = read_emails_from_file('recipients.txt')
    cc_emails = read_emails_from_file('cc.txt')

    # Cấu hình email
    mail.Subject = 'Thử nghiệm gửi email kèm file Excel'
    mail.Body = 'Chào anh/chị,\n\nĐây là file Excel được gửi tự động từ Python.\n\nTrân trọng.'
    mail.To = "; ".join(to_emails)
    mail.CC = "; ".join(cc_emails)

    # Đính kèm file Excel
    mail.Attachments.Add(file_path)

    # Gửi email
    mail.Send()
    print("Email đã được gửi thành công với file đính kèm!")
