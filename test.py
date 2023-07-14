from odf.opendocument import OpenDocument, load
from odf.chart import Chart
from odf.table import Table, TableRows, TableRow, TableCell
from pprint import pprint


if __name__ == '__main__':
	doc = load('original/template.odt')

	charts = [obj for obj in doc.childobjects if obj.getMediaType() == 'application/vnd.oasis.opendocument.chart']
	pprint(charts)

	print('###')

	for chart in charts:
		rows = chart.getElementsByType(TableRow)
		for row in rows:
			for cell in row.getElementsByType(TableCell):
				print(cell)
		print('---')

	# rows = c.element_dict[('urn:oasis:names:tc:opendocument:xmlns:table:1.0', 'table-rows')]

	pass
