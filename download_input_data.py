import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, Color
from progress.bar import ChargingBar
import time

scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("My First Project-d3babb25faaa.json", scope)
client = gspread.authorize(creds)

spr = client.open_by_url(
    'https://docs.google.com/spreadsheets/d/1rmnPuoNXHWGy8AmFbhrr3QwJh2jj5AHso9bMmOYyEJc/edit?usp=sharing')
wks = spr.worksheet('input')
# INPUT GOOGLE SHEET CONNECTION ESTABLISHED

inpt_workbook = Workbook()
inpt_worksheet = inpt_workbook.active
inpt_worksheet.title = "Input"
# Excel file with inputs made and a sheet with Input name created


header_row = wks.row_values(1)
inpt_worksheet.append(header_row)
fmt_bold = Font(bold=True)
inpt_worksheet['A1'].font = fmt_bold
inpt_worksheet['B1'].font = fmt_bold
inpt_worksheet['C1'].font = fmt_bold
inpt_worksheet['D1'].font = fmt_bold
inpt_worksheet['E1'].font = fmt_bold
inpt_worksheet['F1'].font = fmt_bold
inpt_worksheet['G1'].font = fmt_bold
inpt_worksheet['H1'].font = fmt_bold
inpt_worksheet['I1'].font = fmt_bold
inpt_worksheet['J1'].font = fmt_bold
inpt_worksheet['K1'].font = fmt_bold
# header_row inputted and made bold for 8 input samples

num_inputs = len(wks.col_values(1))
bar = ChargingBar('\rDownloading data ', max=num_inputs - 2)
for i in range(2, num_inputs):
    time.sleep(1)
    inpt_worksheet.append(wks.row_values(i))
    bar.next()
# All the data inputted
bar.finish()

try:
    inpt_workbook.save('Input.xlsx')
except:
    print('There already exists an Input.xlsx file in the folder. Kindly move/delete the existing file before proceeding.')
