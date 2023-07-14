from odf.opendocument import OpenDocument, load
from odf.chart import Chart
from odf.table import Table, TableRows, TableRow, TableCell
from pprint import pprint


# Maybe TODO: deal with cells that don't have a direct value, but e.g. a child element with text
def dump_all_charts(doc: OpenDocument) -> list[list[list[str]]]:
	charts = [obj for obj in doc.childobjects if obj.getMediaType() == 'application/vnd.oasis.opendocument.chart']
	return [[[cell.getAttribute('value') for cell in row.getElementsByType(TableCell)] for row in chart.getElementsByType(TableRow)] for chart in charts]

def dump_chart(doc: OpenDocument, index: int) -> list[list[str]]:
	chart = [obj for obj in doc.childobjects if obj.getMediaType() == 'application/vnd.oasis.opendocument.chart'][index]
	return [[str(cell) for cell in row.getElementsByType(TableCell)] for row in chart.getElementsByType(TableRow)]

def write_chart(doc: OpenDocument, index: int, data: list[list[str]]) -> None:
	chart = [obj for obj in doc.childobjects if obj.getMediaType() == 'application/vnd.oasis.opendocument.chart'][index]
	for row, new_row in zip(chart.getElementsByType(TableRow), data):
		for cell, new_value in zip(row.getElementsByType(TableCell), new_row):
			if len(cell.attributes) >= 2 and cell.getAttribute('value') != new_value:
				# TODO: also edit the text value
				cell.setAttribute('value', new_value)
	doc.save('original/template2.odt')


if __name__ == '__main__':
	doc = load('original/template2.odt')

	pprint(dump_all_charts(doc))

	# data = dump_chart(doc, 0)
	# pprint(data)
	# data[0][0] = 'Hello'
	# write_chart(doc, 0, data)
