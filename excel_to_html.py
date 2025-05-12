import win32com.client as win32
 
file_path = r'C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report\FATP_WIP_Report_Detail_20250509093600.xlsx'
html_path = r'C:\Users\V5030587\OneDrive - quantacn.com\Desktop\excel_report\output.html'
 
if __name__ == "__main__":
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    workbook = excel.Workbooks.Open(file_path)
    workbook.SaveAs(html_path, FileFormat=44)
    workbook.Close()
    excel.Quit()