import win32com.client

# Hàm đọc danh sách email từ tệp
def read_emails_from_file(filename):
    with open(filename, 'r') as file:
        emails = file.readlines()
    # Loại bỏ khoảng trắng và dòng trống
    return [email.strip() for email in emails if email.strip()]

# Khởi tạo đối tượng Outlook
outlook = win32com.client.Dispatch('Outlook.Application')

# Tạo một thư mới
mail = outlook.CreateItem(0)

# Đọc danh sách người nhận và CC từ tệp
to_emails = read_emails_from_file('recipients.txt')  # Đọc người nhận chính
cc_emails = read_emails_from_file('cc.txt')  # Đọc người nhận CC

# Cấu hình các thuộc tính của email
mail.Subject = 'Thử nghiệm gửi email qua Outlook 2'  # Tiêu đề email
mail.Body = 'đây là thông báo từ danh sách email va cc'  # Nội dung email

# Thêm người nhận chính (To)
mail.To = "; ".join(to_emails)  # Chuyển danh sách email thành chuỗi với dấu phân cách ";"

# Thêm người nhận CC
mail.CC = "; ".join(cc_emails)  # Chuyển danh sách email CC thành chuỗi với dấu phân cách ";"

# Gửi email
mail.Send()

print("Email đã được gửi thành công!")
