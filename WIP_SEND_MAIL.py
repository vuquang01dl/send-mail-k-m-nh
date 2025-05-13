import os
import win32com.client
import time
from datetime import datetime
import logging
# --- Cấu hình SMTP ---
smtp_server = "172.20.179.74"
smtp_port = 25
smtp_user = "QMS_FA1_Notice"
smtp_password = "Quanta123"
from_addr = "QMS_FA1_Notice@quantacn.com"

# ======================
# 📂 CẤU HÌNH ĐƯỜNG DẪN
# ======================
EXCEL_FOLDER = r"E:\Report\WIP_ReportDetail\Report"
HTML_EXPORT = r"E:\Report\WIP_ReportDetail\HTML_EXPORT"
LOG_FILE = r"E:\Report\WIP_ReportDetail\log_wip_report.txt"
RECIPIENT_FILE = r"E:\Report\WIP_ReportDetail\recipients.txt"
CC_FILE = r"E:\Report\WIP_ReportDetail\cc.txt"
SHEET_NAME = "汇总"
SHEET_DESCRIPTION = "以下为 PWYU & PWZJ WIP 状况."

# ======================
# 📝 CẤU HÌNH LOGGING
# ======================
logging.basicConfig(
    filename=LOG_FILE,
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO,
    encoding='utf-8'
)

logging.info("=== Start executing the email sending script ===")
timeStart = datetime.now()
today_for_subject = timeStart.strftime("%Y%m%d")

# ======================
# 🔧 HÀM XỬ LÝ FILE
# ======================
def get_latest_excel_file(folder_path):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
    if not excel_files:
        raise FileNotFoundError("No .xlsx files found in the directory.")
    excel_files = sorted(excel_files, key=lambda f: os.path.getmtime(os.path.join(folder_path, f)), reverse=True)
    return os.path.join(folder_path, excel_files[0])

def read_emails_from_file(filename):
    with open(filename, 'r', encoding='utf-8') as file:
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
        logging.info(f"Saved HTML to {target_path}")
    finally:
        for sheet in wb.Sheets:
            sheet.Visible = True
        wb.Close(False)
        excel.Quit()

def clear_html_export_folder():
    if os.path.exists(HTML_EXPORT):
        for filename in os.listdir(HTML_EXPORT):
            file_path = os.path.join(HTML_EXPORT, filename)
            try:
                if os.path.isfile(file_path):
                    os.remove(file_path)
                elif os.path.isdir(file_path):
                    import shutil
                    shutil.rmtree(file_path)
            except Exception as e:
                logging.warning(f"Unable to delete {file_path}. Reason: {e}")
        logging.info("Cleared old HTML export files.")


# ======================
# 📧 GỬI MAIL QUA OUTLOOK
# ======================
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
from email.mime.base import MIMEBase
from email import encoders

def send_email_with_html_content(html_main_path, recipients, cc_recipients, attachment_path=None):
    # Thử với cả hai loại thư mục
    possible_dirs = [
        html_main_path.replace(".html", "_files"),
        html_main_path.replace(".html", ".files")
    ]

    html_dir = None
    for d in possible_dirs:
        if os.path.exists(d):
            html_dir = d
            break

    if html_dir is None:
        raise FileNotFoundError(f"Missing HTML export directory for: {html_main_path}")

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
                <p>{SHEET_DESCRIPTION} Cut off time: <b style="color:red;">{cutoff_time}</b>, Thanks</p>
                <p style="color:blue;"><b>PWYU&PWZJ WIP 状况：</b></p>
                {html_body}
            </body>
        </html>
    """

    msg = EmailMessage()
    msg['Subject'] = f"FATP WIP 分布状况{today_for_subject}"
    msg['From'] = formataddr(("QMS FA1 Notice", from_addr))
    msg['To'] = ", ".join(recipients)
    msg['Cc'] = ", ".join(cc_recipients)
    msg.set_content("This is an HTML email. Please view it in HTML-compatible mail client.")
    msg.add_alternative(full_html, subtype='html')

    # Đính kèm file nếu có
    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(attachment_path)}"')
            msg.attach(part)
        logging.info(f"Attached file: {attachment_path}")

    # Gửi mail
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.login(smtp_user, smtp_password)
            server.send_message(msg)
        logging.info("Email sent successfully via SMTP.")
    except Exception as e:
        logging.error(f"Failed to send email via SMTP: {e}")
        raise

# ======================
# 🚀 MAIN EXECUTION
# ======================
try:
    if not os.path.exists(HTML_EXPORT):
        os.makedirs(HTML_EXPORT)
        logging.info(f"Created HTML export directory: {HTML_EXPORT}")

    clear_html_export_folder()

    latest_excel_file = get_latest_excel_file(EXCEL_FOLDER)
    logging.info(f"Found the latest Excel file: {latest_excel_file}")

    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    html_output_file = os.path.join(HTML_EXPORT, f"summary_{timestamp}.html")

    convert_excel_to_html_with_format(latest_excel_file, html_output_file, SHEET_NAME)

    recipients = read_emails_from_file(RECIPIENT_FILE)
    cc_recipients = read_emails_from_file(CC_FILE)
    logging.info("Read recipient list and CC list successfully.")

    send_email_with_html_content(html_output_file, recipients, cc_recipients, attachment_path=latest_excel_file)

    logging.info("=== End of email sending script ===")

except Exception as e:
    logging.error(f"Script terminated with error: {e}")
