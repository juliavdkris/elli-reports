import openpyxl
from docxtpl import DocxTemplate
from docx_charts import Document

SPREADSHEET = 'original/PersRep_Data Pseud.xlsx'
TEMPLATE = 'original/template.docx'
OUTPUT_DIR = 'out'


Entry = dict[str, str|float]


def parse_sheet_entries(filename: str) -> list[Entry]:
	wb = openpyxl.load_workbook(filename)
	sheet = wb[wb.sheetnames[0]]
	student_identification = wb['student_identification']

	entries: list[Entry] = []

	cols = [str(cell.value) for cell in sheet[1]] + [str(cell.value) for cell in student_identification[1]]
	rows = list(sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=True))
	sids = list(student_identification.iter_rows(min_row=2, max_row=student_identification.max_row, values_only=True))
	for i, (row, sid) in enumerate(zip(rows, sids)):
		entry: Entry = {}
		for col, val in zip(cols, row+sid):
			if col is None or val is None:
				continue
			if entry.get(col) and entry[col] != str(val):
				print(f'[!] Duplicate column name with different values not allowed: {col} on row {i+1} ({entry[col]} != {val})')
				exit(1)
			entry[col.lower()] = float(str(val)) if col.startswith('num_') else str(val)
		entries.append(entry)

	if len(entries) < len(rows):
		print(f'[!] No identification was found for {len(rows) - len(sids)}/{len(rows)} students')

	return entries


def generate_report(entry: Entry, template: str, output_dir: str) -> None:
	doc = DocxTemplate(template)
	context = {k: round(v, 2) if isinstance(v, float) else v for k, v in entry.items()}
	doc.render(context)
	doc.save(f'{output_dir}/{entry.get("new_id")}.docx')

	print(f'[*] Generated report for {entry.get("new_id")} ({entry.get("student_name")})')


if __name__ == '__main__':
	for entry in parse_sheet_entries(SPREADSHEET):
		generate_report(entry, TEMPLATE, OUTPUT_DIR)
