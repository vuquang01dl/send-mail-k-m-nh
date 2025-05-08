from xlsx2html import xlsx2html
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import datetime
import win32com.client
import re

if __name__ == "__main__":
    timeStart = datetime.datetime.now()
    file_path = '//QMHFS01/Digital_Worforce_RPA/QMSTemp/PP15WO_Detail_' + timeStart.strftime('%Y%m%d%H') + '.xlsx'
    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(file_path)
    excel.Visible = True
    workbook.Save()
    workbook.Close(SaveChanges=True)
    excel.Quit()

    html_path_1 = r'\\Qmhrpa\Victor\LCM\AutoSentWODetail\.code\Python\table1.html'
    html_path_2 = r'\\Qmhrpa\Victor\LCM\AutoSentWODetail\.code\Python\table2.html'
    html_path_3 = r'\\Qmhrpa\Victor\LCM\AutoSentWODetail\.code\Python\table3.html'

    xlsx2html(file_path, html_path_1, sheet=3)
    xlsx2html(file_path, html_path_2, sheet=1)
    xlsx2html(file_path, html_path_3, sheet=2)

    body = "<div><p style='font-size:12pt;font-family:宋体;margin:0'><span lang='en-US'>&nbsp;</span></p><p style='font-size:12pt;font-family:宋体;margin:0'><b><span style='color:#0432ff;font-size:16pt;font-family:Verdana,sans-serif' lang='en-US'>Dear All: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></b><span style='color:#0432ff;font-size:16pt;font-family:Verdana,sans-serif' lang='en-US'>&nbsp;</span><span lang='en-US'></span></p><p style='font-size:12pt;font-family:宋体;text-align:justify;margin:0'><b><span style='color:#0432ff;font-size:16pt;font-family:Verdana,sans-serif' lang='en-US'>&nbsp;</span></b></p><p style='font-size:12pt;font-family:宋体;text-align:justify;margin:0'><b><span style='color:#0432ff;font-size:16pt;font-family:Verdana,sans-serif' lang='en-US'>QMH PU6 LCM/FATP/SMT WO Status Report " + timeStart.strftime('%Y-%m-%d %H:%M%p') + ".</span></b></p><p style='font-size:12pt;font-family:宋体;text-align:justify;margin:0'><b><span style='color:#0432ff;font-size:16pt;font-family:Verdana,sans-serif' lang='en-US'>Please let me know if you have any questions.</span></b></p><p style='font-size:12pt;font-family:宋体;margin:0'><b><span style='color:#0432ff;font-size:16pt;font-family:Verdana,sans-serif' lang='en-US'>&nbsp;</span></b></p>"

    body += "<span style='color:#0432ff;font-size:16pt' lang='zh-TW'>超周工单明细：</span><br>"

    with open(html_path_1, 'r', encoding='utf-8') as file:
        up_content = re.sub(r'font-size:\s*\d+(\.\d+)?px', 'font-size: 14.0px', file.read())
        body += up_content
    body += "<span style='font-size:10.5pt;font-family:Aptos,sans-serif' lang='vi'></span><div><div><p style='font-size:12pt;font-family:宋体;margin:0'><b><span style='color:#0432ff;font-size:16pt;font-family:Verdana,sans-serif' lang='en-US'>&nbsp;</span></b></p><p style='font-size:12pt;font-family:宋体;margin:0'><b><span style='color:#0432ff;font-size:16pt;font-family:Verdana,sans-serif' lang='en-US'>PP10/PP11 :</span></b><b><span style='color:#0432ff;font-size:16pt;font-family:Verdana,sans-serif' lang='vi'></span></b></p><p style='font-size:12pt;font-family:宋体;margin:0'><b><span style='font-family:Aptos,sans-serif' lang='vi'>&nbsp;</span></b></p><p style='font-size:12pt;font-family:宋体;margin:0'><span lang='en-US'>"
    with open(html_path_2, 'r', encoding='utf-8') as file:
        up_content = re.sub(r'font-size:\s*\d+(\.\d+)?px', 'font-size: 14.0px', file.read())
        body += up_content
    body += "<span style='color:#0432ff;font-size:16pt;font-family:Verdana,sans-serif' lang='en-US'></span></b></p><p style='font-size:12pt;font-family:宋体;margin:0'><b><span style='color:#0432ff;font-size:16pt;font-family:Verdana,sans-serif' lang='en-US'>&nbsp;</span></b></p><p style='font-size:12pt;font-family:宋体;margin:0'><b><span style='color:#0432ff;font-size:16pt;font-family:Verdana,sans-serif' lang='en-US'>PP13:</span></b><b><span style='font-size:16pt;font-family:Verdana,sans-serif' lang='en-US'></span></b></p><p style='font-size:12pt;font-family:宋体;margin:0'><b><span style='font-family:Aptos,sans-serif' lang='en-US'>&nbsp;</span></b></p><p style='font-size:12pt;font-family:宋体;margin:0'><span lang='en-US'>"
    with open(html_path_3, 'r', encoding='utf-8') as file:
        up_content = re.sub(r'font-size:\s*\d+(\.\d+)?px', 'font-size: 14.0px', file.read())
        body += up_content
    body += "<span style='font-size:10.5pt;font-family:等线' lang='en-US'></span></b></p><p style='font-size:12pt;font-family:宋体;margin:0'><b><span style='color:#0432ff;font-size:16pt;font-family:Verdana,sans-serif' lang='en-US'>&nbsp;</span></b></p></div></div><p style='font-size:12pt;font-family:宋体;margin:0'><span style='font-size:10.5pt;font-family:等线' lang='en-US'>&nbsp;</span></p><p style='font-size:12pt;font-family:宋体;margin:0'><span style='font-family:Aptos,sans-serif' lang='en-US'>&nbsp;</span></p>"
    body += "<div class=WordSection1><p class=MsoNormal align=left style='text-align:left'><span style='font-size:14.0pt;font-family:\"Times New Roman\",serif'>Time start: " + timeStart.strftime(
        '%Y-%m-%d %H:%M:%S') + " <o:p></o:p></span></p><p class=MsoNormal align=left style='text-align:left'><span style='font-size:14.0pt;font-family:\"Times New Roman\",serif'>Time finish: " + datetime.datetime.now().strftime(
        '%Y-%m-%d %H:%M:%S') + "</span><span lang=VI style='font-size:14.0pt;font-family:\"Times New Roman\",serif'><o:p></o:p></span></p><p class=MsoNormal style='margin-bottom:12.0pt'><b><span lang=VI style='font-size:14.0pt;font-family:\"Times New Roman\",serif;color:red'>MIS负责RPA开发，使用单位需依附件确认RPA执行结果，并处理异常<o:p></o:p></span></b></p><p class=MsoNormal style='margin-bottom:12.0pt'><b><span lang=VI style='font-size:14.0pt;font-family:\"Times New Roman\",serif;color:red'>MIS chịu trách nhiệm phát triển RPA, đơn vị người dùng phải xác nhận kết quả thực hiện RPA theo tệp đính kèm và xử lý bất thường.<o:p></o:p></span></b></p></div></body></html>"
    sender_email = "QMH_SDSNotice@quantacn.com"
    sender_password = "D+1234567890"
    receiver_email = "James.Wang@quantacn.com,Rook.Ding@quantacn.com,W-D.Wu@quantacn.com,Allen.Wang@quantacn.com,Fangli.Yang@quantacn.com,Yunlong.Xing@quantacn.com,Justin.Lv@quantacn.com,Saiping.Zheng@quantacn.com,Apple.Sun@quantacn.com,Peter.Mao@quantacn.com,Shi.Ningning@quantacn.com,shunting.zhao@quantacn.com,Sandy.Liu@quantacn.com,Park.Zhang@quantacn.com,Davy.Wang@quantacn.com,W.Z.Chen@quantacn.com,Ming.Fang@quantacn.com,B.Z.Zhang@quantacn.com,Peter.Zhou@quantacn.com,Bo.Wang@quantacn.com,Guo-Tao.Li@quantacn.com,G.Q.Cui@quantacn.com,Xiaowen.Xu@quantacn.com,Xiaoxi.Chen@quantacn.com,Wei.Bian@quantacn.com,Ou.Yang@quantacn.com,King.Dong@quantacn.com,Xigui.Quan@quantacn.com,Fan.Fan@quantacn.com,Xingbo.Yang@quantacn.com,Ting.Xu@quantacn.com,alina.lan@quantacn.com,bella.wei@quantacn.com,Zhengde.Gan@quantacn.com,Robert.Wang1@quantacn.com,Neil.tran@quantaqmh.com,Kevin.Bui@quantaqmh.com,doris.chan@quantaqmh.com,Qingxuan.Chenshi@quantaqmh.com,Helen.Fanshi@quantaqmh.com,Edric.Tran@quantaqmh.com,Alice.Ruan@quantaqmh.com,Micheal.Chen@quantaqmh.com,Cris.Le@quantaqmh.com,Yulong.Zhao@quantacn.com,Conan.Zhong@quantacn.com"
    cc_email = "Ken.Hsu@quantacn.com,yang.yuan@quantacn.com,Bruce.Wu@quantacn.com,David.Lee@quantacn.com,Jacky.Mu@quantacn.com,bobo.zhou@quantacn.com,michael.liu1@quantacn.com,otto.chen@quantacn.com"
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Cc'] = cc_email
    message['Subject'] = "QMH LCM/FATP/SMT PP10/PP11/PP13 WO Detail Report-" + timeStart.strftime('%Y%m%d')

    message.attach(MIMEText(body, 'html'))

    attachment = open(file_path, "rb")

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)

    part.add_header('Content-Disposition', "attachment; filename=PP15WO_Detail_" + timeStart.strftime('%Y%m%d%H') + ".xlsx")

    message.attach(part)

    try:
        server = smtplib.SMTP("qsmcfe.quantacn.com", 25)

        server.login(sender_email, sender_password)

        text = message.as_string()
        server.sendmail(sender_email, receiver_email.split(",") + cc_email.split(","), text)
    except Exception as e:
        print("Error: " + e)
