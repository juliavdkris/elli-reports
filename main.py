import openpyxl
from docxtpl import DocxTemplate

SPREADSHEET = 'original/PersRep_Data Pseud.xlsx'
TEMPLATE = 'original/PersonalizedReport_DraftV6.docx'


wb = openpyxl.load_workbook(SPREADSHEET)
sheet = wb[wb.sheetnames[0]]
student_identification = wb['student_identification']
