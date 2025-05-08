from xlsx2html import xlsx2html
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime
import win32com.client
import re

def read_emails_from_txt(file_path):
    """Đọc danh sách email từ file txt, mỗi dòng là 1 email"""
    with open(file_path, 'r', encoding='utf-8') as f:
        emails = [line.strip() for line in f if line.strip()]
    return ','.join(emails)

if __name__ == "__main__":
    # Định dạng thời gian để tạo tên file
    timeStart = datetime.datetime.now()
    file_path = '//QMHFS01/Digital_Worforce_RPA/QMSTemp/PP15WO_Detail_' + timeStart.strftime('%Y%m%d%H') + '.xlsx'

    # Mở và lưu lại file Excel để đảm bảo định dạng
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(file_path)
    excel.Visible = False
    workbook.Save()
    workbook.Close(SaveChanges=True)
    excel.Quit()

    # Chuyển Excel sang HTML
    html_path_1 = r'\\Qmhrpa\Victor\LCM\AutoSentWODetail\.code\Python\table1.html'
    html_path_2 = r'\\Qmhrpa\Victor\LCM\AutoSentWODetail\.code\Python\table2.html'
    html_path_3 = r'\\Qmhrpa\Victor\LCM\AutoSentWODetail\.code\Python\table3.html'

    xlsx2html(file_path, html_path_1, sheet=3)
    xlsx2html(file_path, html_path_2, sheet=1)
    xlsx2html(file_path, html_path_3, sheet=2)

    # Soạn nội dung email
    body = """
    <html><body>
    <p><b>Dear All,</b></p>
    <p><b>QMH PU6 LCM/FATP/SMT WO Status Report """ + timeStart.strftime('%Y-%m-%d %H:%M') + """</b></p>
    <p><b>Please let me know if you have any questions.</b></p>
    <p><b>超周工单明细：</b></p>
    """

    for html_file in [html_path_1, html_path_2, html_path_3]:
        with open(html_file, 'r', encoding='utf-8') as file:
            up_content = re.sub(r'font-size:\s*\d+(\.\d+)?px', 'font-size: 14.0px', file.read())
            body += up_content

    body += """
    <p>Time start: """ + timeStart.strftime('%Y-%m-%d %H:%M:%S') + """</p>
    <p>Time finish: """ + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + """</p>
    <p><b style='color:red'>MIS负责RPA开发，使用单位需依附件确认RPA执行结果，并处理异常</b></p>
    <p><b style='color:red'>MIS chịu trách nhiệm phát triển RPA, đơn vị người dùng phải xác nhận kết quả thực hiện RPA theo tệp đính kèm và xử lý bất thường.</b></p>
    </body></html>
    """

    # Lấy email người nhận và CC từ file
    to_emails = read_emails_from_txt("recipients.txt")
    cc_emails = read_emails_from_txt("cc.txt")

    # Gửi email
    sender_email = "QMH_SDSNotice@quantacn.com"
    sender_password = "D+1234567890"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_emails
    msg['Cc'] = cc_emails
    msg['Subject'] = "[AutoMail] QMH WO Detail Report - " + timeStart.strftime('%Y-%m-%d %H:%M')

    msg.attach(MIMEText(body, 'html'))

    # Tạo kết nối SMTP và gửi email
    try:
        with smtplib.SMTP('10.121.11.219', 25) as server:
            server.sendmail(sender_email, to_emails.split(',') + cc_emails.split(','), msg.as_string())
        print("✔ Email sent successfully!")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")
