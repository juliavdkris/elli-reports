from odf.opendocument import OpenDocument, load
from odf.chart import Chart
from odf.table import Table
from pprint import pprint


if __name__ == '__main__':
	doc = load('original/demo.odt')

	charts = [obj for obj in doc.childobjects if obj.getMediaType() == 'application/vnd.oasis.opendocument.chart']
	pprint(charts)

	pass
