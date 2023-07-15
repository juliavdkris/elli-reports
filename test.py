from odf.opendocument import OpenDocument, load
from odf.element import Element
from odf.chart import Chart
from odf.table import Table, TableRows, TableRow, TableCell
from pprint import pprint


# Type aliases
CellValue = tuple[str|None, str|None] | None
ChartData = list[list[CellValue]]


def get_charts(doc: OpenDocument) -> list[Element]:
	return [obj for obj in doc.childobjects if obj.getMediaType() == 'application/vnd.oasis.opendocument.chart']


def parse_cell(cell: Element) -> CellValue:
	if cell.getAttribute('value'):
		return (cell.getAttribute('value'), cell.attributes.get(('urn:oasis:names:tc:opendocument:xmlns:office:1.0', 'value-type')))
	elif len(cell.attributes) >= 1  and len(cell.childNodes[0].childNodes) >= 1:
		return (cell.childNodes[0].childNodes[0].data, 'string')
	else:
		return None


def dump_all_charts(doc: OpenDocument) -> list[ChartData]:
	charts = get_charts(doc)
	return [dump_chart(doc, i) for i in range(len(charts))]


def dump_chart(doc: OpenDocument, index: int) -> ChartData:
	chart = get_charts(doc)[index]
	return [[parse_cell(cell) for cell in row.getElementsByType(TableCell)] for row in chart.getElementsByType(TableRow)]


def write_chart(doc: OpenDocument, index: int, data: ChartData) -> None:
	chart = get_charts(doc)[index]
	for row, new_row in zip(chart.getElementsByType(TableRow), data):
		for cell, new_value in zip(row.getElementsByType(TableCell), new_row):
			if new_value and new_value[1] == 'float' and len(cell.attributes) >= 2 and cell.getAttribute('value') != new_value[0]:
				# TODO: also edit the text value?
				cell.setAttribute('value', new_value[0])
	doc.save('original/template2.odt')


if __name__ == '__main__':
	doc = load('original/template2.odt')

	pprint(dump_all_charts(doc))

	# data = dump_chart(doc, 0)
	# pprint(data)
	# data[1][2] = '999'
	# write_chart(doc, 0, data)
