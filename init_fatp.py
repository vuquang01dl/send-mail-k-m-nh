import json

from xlsx2html import xlsx2html
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import datetime
import requests
import pandas as pd
from openpyxl import load_workbook
import re

def clear_data_excel(ws, start_row: int, end_row: int):
    for row in range(start_row, end_row):
        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col, value="")

def df_to_excel(df, ws, start_row):
    for r_idx, row in enumerate(df.values, start=start_row):
        for c_idx, value in enumerate(row, start=1):
            if isinstance(value, str) and '%' in value:
                try:
                    value = float(value.replace('%', '').strip()) / 100
                except ValueError:
                    pass
            else:
                try:
                    value = float(value)
                except ValueError:
                    pass
            ws.cell(row=r_idx, column=c_idx + 1, value=value)

if __name__ == "__main__":
    timeStart = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    time_now = datetime.datetime.now()

    url = "http://10.36.6.76/Web_CTO_HitRate_Report/CTOHitRate/GetAllData"
    response = requests.request("POST", url)
    data = json.loads(response.json()["ShowTablelist"])
    data_J614 = [row for row in data["values"] if row[1] == 'J614']
    data_J616 = [row for row in data["values"] if row[1] == 'J616']
    df_J614 = pd.DataFrame(data_J614)
    df_J616 = pd.DataFrame(data_J616)

    file_path = r"\\Qmhrpa\Victor\FATP\QMH OCT Per formence report.xlsx"
    html_path_J614 = r"\\Qmhrpa\Victor\FATP\table_J614.html"
    html_path_J616 = r"\\Qmhrpa\Victor\FATP\table_J616.html"
    email_file_path = r"\\Qmhrpa\Victor\FATP\sent_email_list.json"

    wb = load_workbook(file_path)
    ws_J614 = wb["J614"]
    ws_J616 = wb["J616"]
    start_row = 5
    end_row = 11
    clear_data_excel(ws_J614, start_row, end_row)
    clear_data_excel(ws_J616, start_row, end_row)

    df_to_excel(df_J614, ws_J614, start_row)
    df_to_excel(df_J616, ws_J616, start_row)

    wb.save(file_path)

    xlsx2html(file_path, html_path_J614, locale='en', sheet=0)
    xlsx2html(file_path, html_path_J616, locale='en', sheet=1)

    body = ('<p style="font-size:10pt;font-family:Calibri,sans-serif;margin:0;"><b><span style="font-size:16pt;font-family:Times;" lang="en-US">Dear all:</span></b><span style="font-size:10.5pt;" lang="en-US"></span></p>'
            '<p style="font-size:10pt;font-family:Calibri,sans-serif;margin:0;"><span lang="en-US">&nbsp;</span></p>'
            '<p style="font-size:10pt;font-family:Calibri,sans-serif;margin:0;"> <span style="font-size:14pt;font-family:Times;" lang="en-US">Update FATP CTO output status. Cut off time: <span style="color:black;">' + time_now.strftime("%Y-%m-%d") + '</span> <b style="color:red;font-size:16pt;">' + time_now.strftime("%H:%M") + '</b> </span> <span style="font-size:14pt;font-family:Times;color:black;">Thanks!</span> </p>'
            '<p style="font-size:10pt;font-family:Calibri,sans-serif;margin:0;"><span lang="en-US">&nbsp;</span></p>'
            '<p style="margin:0cm; margin-bottom:.0001pt"><b><span style="font-size: 18pt; font-family: Arial, sans-serif, serif, EmojiFont; color: blue;" lang="VI">(PWYU) J614:</span></b></p>')

    with open(html_path_J614, 'r', encoding='utf-8') as file:
        up_content = re.sub('cellpadding="0"', 'cellpadding="10"', file.read())
        body += up_content

    body += ('<p style="font-size:10pt;font-family:Calibri,sans-serif;margin:0;"><span lang="en-US">&nbsp;</span></p>'
             '<p style="margin:0cm; margin-bottom:.0001pt"><b><span style="font-size: 18pt; font-family: Arial, sans-serif, serif, EmojiFont; color: blue;" lang="VI">(PWZJ) J616:</span></b></p>')

    with open(html_path_J616, 'r', encoding='utf-8') as file:
        up_content = re.sub('cellpadding="0"', 'cellpadding="10"', file.read())
        body += up_content
    with open(email_file_path, 'r') as file:
        email = json.load(file)

    body += ('<p style="font-size:10pt;font-family:Calibri,sans-serif;margin:0;"><span lang="en-US">&nbsp;</span></p>'
             "<div class=WordSection1><p class=MsoNormal align=left style='text-align:left'><span style='font-size:14.0pt;font-family:\"Times New Roman\",serif'>Time start: " + timeStart + " <o:p></o:p></span></p><p class=MsoNormal align=left style='text-align:left'><span style='font-size:14.0pt;font-family:\"Times New Roman\",serif'>Time finish: " + datetime.datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S") + "</span><span lang=VI style='font-size:14.0pt;font-family:\"Times New Roman\",serif'><o:p></o:p></span></p><p class=MsoNormal style='margin-bottom:12.0pt'><b><span lang=VI style='font-size:14.0pt;font-family:\"Times New Roman\",serif;color:red'>MIS负责RPA开发，使用单位需依附件确认RPA执行结果，并处理异常<o:p></o:p></span></b></p><p class=MsoNormal style='margin-bottom:12.0pt'><b><span lang=VI style='font-size:14.0pt;font-family:\"Times New Roman\",serif;color:red'>MIS chịu trách nhiệm phát triển RPA, đơn vị người dùng phải xác nhận kết quả thực hiện RPA theo tệp đính kèm và xử lý bất thường.<o:p></o:p></span></b></p></div></body></html>")
    sender_email = "QMH_SDSNotice@quantacn.com"
    sender_password = "D+1234567890"
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = ";".join(email["receiver_email"])
    message['Cc'] = ";".join(email["cc_email"])
    message['Subject'] = "QMH FATP OCT Performance_" + datetime.datetime.now().strftime('%Y-%m-%d')

    message.attach(MIMEText(body, 'html'))

    attachment = open(file_path, "rb")

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)

    part.add_header('Content-Disposition',
                    "attachment; filename=QMH OCT Per formence report.xlsx")

    message.attach(part)

    server = smtplib.SMTP("qsmcfe.quantacn.com", 25)
    server.login(sender_email, sender_password)
    server.sendmail(message['From'], email["receiver_email"] + email["cc_email"], message.as_string())
