import openpyxl
from docxtpl import DocxTemplate
from dataclasses import dataclass
from typing import Optional

SPREADSHEET = 'original/PersRep_Data Pseud.xlsx'
TEMPLATE = 'original/PersonalizedReport_DraftV6.docx'


@dataclass
class Entry:
	new_id: str
	course: str
	eb1: Optional[float]
	mean_eb1_total: Optional[float]
	eb2: Optional[float]
	mean_eb2_total: Optional[float]
	vb1: Optional[float]
	mean_vb1_total: Optional[float]
	vb2: Optional[float]
	mean_vb2_total: Optional[float]
	ae1: Optional[float]
	mean_ae1_total: Optional[float]
	ae2: Optional[float]
	mean_ae2_total: Optional[float]
	ae3: Optional[float]
	mean_ae3_total: Optional[float]
	ae4: Optional[float]
	mean_ae4_total: Optional[float]
	be1: Optional[float]
	mean_be1_total: Optional[float]
	be2: Optional[float]
	mean_be2_total: Optional[float]
	be3: Optional[float]
	mean_be3_total: Optional[float]
	be4: Optional[float]
	mean_be4_total: Optional[float]
	ce1: Optional[float]
	mean_ce1_total: Optional[float]
	ce2: Optional[float]
	mean_ce2_total: Optional[float]
	ce3: Optional[float]
	mean_ce3_total: Optional[float]
	ce4: Optional[float]
	mean_ce4_total: Optional[float]


# Parse every row in the sheet as an Entry object
def parse_sheet_entries(filename) -> list[Entry]:
	wb = openpyxl.load_workbook(filename)
	sheet = wb[wb.sheetnames[0]]
	# student_identification = wb['student_identification']
	return [Entry(*row[:-1]) for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=True)] # type: ignore


print(parse_sheet_entries(SPREADSHEET)[0])
