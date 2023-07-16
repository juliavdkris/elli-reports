import openpyxl
from docxtpl import DocxTemplate
from docx_charts import Document

SPREADSHEET = 'original/PersRep_Data Pseud.xlsx'
TEMPLATE = 'original/template.docx'
OUTPUT_DIR = 'out'


Entry = dict[str, str|float|None]


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
			if col is None:
				continue
			if val is None:
				entry[col.lower()] = None
				continue
			if entry.get(col) and entry[col] != str(val):
				print(f'[!] Duplicate column name with different values not allowed: {col} on row {i+1} ({entry[col]} != {val})')
				exit(1)
			entry[col.lower()] = float(str(val)) if col.startswith('num_') else str(val)
		entries.append(entry)

	if len(entries) < len(rows):
		print(f'[!] No identification was found for {len(rows) - len(sids)}/{len(rows)} students')

	return entries


def underscores_to_dict(strings: dict[str, float]) -> dict[str, dict[str, dict[int, float]]]:
	'''
	Example input: {'num_a_student_1': 1, 'num_a_student_2': 2, 'num_a_mean_1': 3, 'num_a_mean_2': 4}
	Example output: {'a': {'student': {1: 1, 2: 2}, 'mean': {1: 3, 2: 4}}}
	'''
	output = {}
	for key, value in strings.items():
		split_key = key.split('_')
		if split_key[1] not in output:
			output[split_key[1]] = {}
		if split_key[2] not in output[split_key[1]]:
			output[split_key[1]][split_key[2]] = {}
		output[split_key[1]][split_key[2]][int(split_key[3])] = value
	return output


def generate_report(entry: Entry, template: str, output_dir: str) -> None:
	# Stage 1: replace text in template
	filename = f'{output_dir}/{entry.get("new_id")}.docx'
	doc = DocxTemplate(template)
	context = {k: round(v, 2) if isinstance(v, float) else v for k, v in entry.items()}
	doc.render(context)
	doc.save(filename)
	print(f'[*] Generated report for {entry.get("new_id")} ({entry.get("student_name")})')

	# Stage 2: update charts in generated report
	doc = Document(filename)
	nums = underscores_to_dict({k: v for k, v in entry.items() if k.startswith('num_') and isinstance(v, float)})
	for key, value in nums.items():
		chart = doc.find_charts_by_name(key)[0]
		data = chart.data()

		# TODO: series names instead of just indices? dict go brrr
		for i, (series, new_series) in enumerate(zip(data, value.values())):
			data[i] = {k1: v2 for (k1, v1), (k2, v2) in list(zip(series.items(), new_series.items()))}

		chart.write_data(data)
	doc.save(filename)
	print(f'    [*] Updated charts for {entry.get("new_id")} ({entry.get("student_name")})')


if __name__ == '__main__':
	for entry in parse_sheet_entries(SPREADSHEET):
		generate_report(entry, TEMPLATE, OUTPUT_DIR)
