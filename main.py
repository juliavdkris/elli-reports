import openpyxl
from docx import Document
from docxtpl import DocxTemplate

from dataclasses import dataclass
from typing import Optional

SPREADSHEET = 'original/PersRep_Data Pseud.xlsx'
TEMPLATE = 'original/template.docx'
OUTPUT_DIR = 'out'


@dataclass
class Entry:
	new_id: str
	student_name: str
	student_email: str
	student_id: str
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
def parse_sheet_entries(filename: str) -> list[Entry]:
	wb = openpyxl.load_workbook(filename)
	sheet = wb[wb.sheetnames[0]]
	student_identification = wb['student_identification']

	entries: list[Entry] = []

	# Join the two sheets together on the student_id column
	# A lil sketchy but ¯\_(ツ)_/¯
	for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=True):
		for sid in student_identification.iter_rows(min_row=2, max_row=student_identification.max_row, values_only=True):
			if row[0] == sid[3]:
				new_id = sid[3]
				student_name = sid[0]
				student_email = sid[1]
				student_id = sid[2]
				entries.append(Entry(*[new_id, student_name, student_email, student_id, *row[1:-1]])) # type: ignore
		else:
			print(f'[!] No identification found for {row[0]}')
	return entries


def generate_report(entry: Entry, template: str, output_dir: str) -> None:
	doc = DocxTemplate(template)
	context = {
		'new_id': entry.new_id,
		'student_name': entry.student_name,
		'student_email': entry.student_email,
		'student_id': entry.student_id,
		'course': entry.course,
	}
	doc.render(context)
	doc.save(f'{output_dir}/{entry.new_id}.docx')

	# doc = Document(f'{output_dir}/{entry.new_id}.docx')

	print(f'[*] Generated report for {entry.new_id} ({entry.student_name})')


if __name__ == '__main__':
	for entry in parse_sheet_entries(SPREADSHEET):
		generate_report(entry, TEMPLATE, OUTPUT_DIR)
