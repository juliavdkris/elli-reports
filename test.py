from odf.opendocument import OpenDocument, load
from odf.element import Element
from odf.chart import Chart
from odf.table import Table, TableRows, TableRow, TableCell
from dataclasses import dataclass
from pprint import pprint


@dataclass
class CellValue:
	value: str | None
	type: str | None

	def __str__(self) -> str:
		if self.value is None:
			return 'NONE'
		if self.type == 'float':
			return f'({float(self.value):.2f}, {self.type})'
		else:
			return f'(\'{self.value}\', {self.type})'
	__repr__ = __str__

ChartData = list[list[CellValue]]


def get_charts(doc: OpenDocument) -> list[Element]:
	return [obj for obj in doc.childobjects if obj.getMediaType() == 'application/vnd.oasis.opendocument.chart']


def parse_cell(cell: Element) -> CellValue:
	if cell.getAttribute('value'):
		return CellValue(cell.getAttribute('value'), cell.attributes.get(('urn:oasis:names:tc:opendocument:xmlns:office:1.0', 'value-type')))
	elif len(cell.attributes) >= 1  and len(cell.childNodes[0].childNodes) >= 1:
		return CellValue(cell.childNodes[0].childNodes[0].data, 'string')
	else:
		return CellValue(None, None)


def dump_chart(doc: OpenDocument, index: int) -> ChartData:
	chart = get_charts(doc)[index]
	return [[parse_cell(cell) for cell in row.getElementsByType(TableCell)] for row in chart.getElementsByType(TableRow)]


def dump_all_charts(doc: OpenDocument) -> list[ChartData]:
	charts = get_charts(doc)
	return [dump_chart(doc, i) for i in range(len(charts))]


def write_chart(doc: OpenDocument, index: int, data: ChartData) -> None:
	chart = get_charts(doc)[index]
	for row, new_row in zip(chart.getElementsByType(TableRow), data):
		for cell, new_cell in zip(row.getElementsByType(TableCell), new_row):
			if new_cell.type == 'float' and len(cell.attributes) >= 2 and cell.getAttribute('value') != new_cell.value:
				# TODO: also edit the text value?
				old_value = cell.getAttribute('value')
				cell.setAttribute('value', new_cell.value)
				cell.childNodes[0].childNodes[0].data = new_cell.value
				print(f'Changed {old_value} to {new_cell.value}')
	doc.save('original/template2.odt')


if __name__ == '__main__':
	doc = load('original/template2.odt')

	# pprint(dump_all_charts(doc))

	data = dump_chart(doc, 0)
	pprint(data)
	if data[1][2].value:
		data[1][2].value = str(float(data[1][2].value) + 1)
	write_chart(doc, 0, data)
