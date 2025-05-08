#gửi mail ok, chuyển nội dung file excel sang HTML
import win32com.client
import datetime
import os
import pandas as pd

# Hàm đọc danh sách email từ file txt
def read_emails_from_file(filename):
    with open(filename, 'r') as file:
        emails = file.readlines()
    return [email.strip() for email in emails if email.strip()]

# Lấy thời gian hiện tại
timeStart = datetime.datetime.now()

# Tạo đường dẫn tới file Excel
file_path = os.path.join(
    r"C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report",
    timeStart.strftime('%Y%m%d%H') + ".xlsx"
)

# Kiểm tra file có tồn tại
if not os.path.exists(file_path):
    print(f"File không tồn tại: {file_path}")
else:
    # Đọc dữ liệu từ file Excel
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print("Lỗi đọc file Excel:", e)
        exit()

    # Chuyển DataFrame thành HTML table (thêm CSS nhẹ cho đẹp)
    table_html = df.to_html(index=False, border=1, classes='excel-table', justify='center')

    # HTML body của email
    html_content = f"""
    <html>
        <head>
            <style>
                .excel-table {{
                    border-collapse: collapse;
                    width: 100%;
                }}
                .excel-table th, .excel-table td {{
                    border: 1px solid #ddd;
                    padding: 8px;
                    text-align: center;
                }}
                .excel-table th {{
                    background-color: #4CAF50;
                    color: white;
                }}
            </style>
        </head>
        <body>
            <p>Chào anh/chị,</p>
            <p>Dưới đây là nội dung bảng báo cáo:</p>
            {table_html}
            <p>File Excel cũng đã được đính kèm theo email này.</p>
            <p>Trân trọng,<br>Hệ thống Python</p>
        </body>
    </html>
    """

    # Khởi tạo Outlook
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)

    # Đọc email người nhận và CC
    mail.To = "; ".join(read_emails_from_file("recipients.txt"))
    mail.CC = "; ".join(read_emails_from_file("cc.txt"))

    # Thiết lập tiêu đề và nội dung HTML
    mail.Subject = f"Báo cáo lúc {timeStart.strftime('%H:00 %d/%m/%Y')}"
    mail.HTMLBody = html_content

    # Đính kèm file Excel
    mail.Attachments.Add(file_path)

    # Gửi email
    mail.Send()
    print("Đã gửi email HTML có bảng Excel và file đính kèm!")
