import os
import win32com.client
import time
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging

# ======================
# 📝 CẤU HÌNH LOGGING
# ======================
log_file = r"E:\Report\WIP_ReportDetail\log_wip_report.txt"
logging.basicConfig(
    filename=log_file,
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO,
    encoding='utf-8'
)

logging.info("=== Start executing the email sending script ===")
timeStart = datetime.now()

# --- Cấu hình SMTP ---
smtp_server = "172.20.179.74"
smtp_port = 25
smtp_user = "QMS_FA1_Notice"
smtp_password = "Quanta123"
from_addr = "QMS_FA1_Notice@quantacn.com"

# Cấu hình đường dẫn
EXCEL_FOLDER = r"E:\Report\WIP_ReportDetail\Report"
SHEET_NAME = "汇总"
HTML_EXPORT = r"E:\Report\WIP_ReportDetail\Report\HTML_EXPORT"
sheet_description = "以下为 PWYU & PWZJ WIP 状况."


def get_latest_excel_file(folder_path):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
    if not excel_files:
        raise FileNotFoundError("No .xlsx files found in the directory.")

    excel_files = sorted(excel_files, key=lambda f: os.path.getmtime(os.path.join(folder_path, f)), reverse=True)
    return os.path.join(folder_path, excel_files[0])


def read_emails_from_file(filename):
    with open(filename, 'r') as file:
        emails = file.readlines()
    return [email.strip() for email in emails if email.strip()]


def convert_excel_to_html_with_format(source_path, target_path, sheet_name="汇总"):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(source_path)

    try:
        ws = wb.Sheets(sheet_name)
        for sheet in wb.Sheets:
            if sheet.Name != sheet_name:
                sheet.Visible = False

        folder = os.path.dirname(target_path)
        if not os.path.exists(folder):
            os.makedirs(folder)

        ws.SaveAs(target_path, FileFormat=44)  # xlHtml
        logging.info("Copied Excel content as HTML successfully.")
    finally:
        for sheet in wb.Sheets:
            sheet.Visible = True
        wb.Close(False)
        excel.Quit()


def send_email_with_html_content(html_main_path, recipients, cc_recipients):
    html_dir = html_main_path.replace(".html", "_files")
    sheet_file = os.path.join(html_dir, "sheet001.html")
    css_file = os.path.join(html_dir, "stylesheet.css")

    if not os.path.exists(sheet_file):
        raise FileNotFoundError(f"Missing file: {sheet_file}")

    css_content = ""
    if os.path.exists(css_file):
        with open(css_file, 'r', encoding='utf-8', errors='replace') as css:
            css_content = f"<style>{css.read()}</style>"

    with open(sheet_file, 'r', encoding='utf-8', errors='replace') as file:
        html_body = file.read()

    cutoff_time = timeStart.strftime("%Y-%m-%d %H:%M")

    full_html = f"""
        <html>
            <head>{css_content}</head>
            <body>
                <p><b><span style="font-size:16pt;">Dear all:</span></b></p>
                <p>{sheet_description} Cut off time: <b style="color:red;">{cutoff_time}</b>, Thanks</p>
                <p style="color:blue;"><b>PWYU&PWZJ WIP 状况：</b></p>
                {html_body}
            </body>
        </html>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "FATP WIP Report"
    msg["From"] = from_addr
    msg["To"] = ", ".join(recipients)
    msg["Cc"] = ", ".join(cc_recipients)
    msg.attach(MIMEText(full_html, "html", "utf-8"))

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.login(smtp_user, smtp_password)
            all_recipients = recipients + cc_recipients
            server.sendmail(from_addr, all_recipients, msg.as_string())
            logging.info("Email sent successfully using Outlook.")
    except Exception as e:
        logging.error(f"Failed to send email: {e}")
        raise


# ======================
# 🚀 MAIN EXECUTION
# ======================
try:
    latest_excel_file = get_latest_excel_file(EXCEL_FOLDER)
    logging.info(f"Found the latest Excel file: {latest_excel_file}")
    
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    if not os.path.exists(HTML_EXPORT):
        os.makedirs(HTML_EXPORT)

    html_output_file = os.path.join(HTML_EXPORT, f"{SHEET_NAME}_{timestamp}.html")

    convert_excel_to_html_with_format(latest_excel_file, html_output_file, SHEET_NAME)

    recipients = read_emails_from_file(r"E:\Report\WIP_ReportDetail\recipients.txt")
    cc_recipients = read_emails_from_file(r"E:\Report\WIP_ReportDetail\cc.txt")
    logging.info("Read recipient list and CC.")

    send_email_with_html_content(html_output_file, recipients, cc_recipients)
    
    logging.info("=== End of email sending script ===")

except Exception as e:
    logging.error(f"Script terminated with error: {e}")
