import win32com.client

# Khởi tạo đối tượng Outlook
outlook = win32com.client.Dispatch('Outlook.Application')

# Tạo một thư mới
mail = outlook.CreateItem(0)

# Cấu hình các thuộc tính của email
mail.To = 'River.Do@quantaqmh.com'  # Địa chỉ người nhận
mail.Subject = 'Thử nghiệm gửi email qua Outlook 2'  # Tiêu đề email
mail.Body = 'Nội dung email từ Python'  # Nội dung email

# Gửi email
mail.Send()

print("Email đã được gửi thành công!")
