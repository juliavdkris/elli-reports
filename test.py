from odf.opendocument import OpenDocument, load
from odf.chart import Chart
from odf.table import Table
from pprint import pprint


if __name__ == '__main__':
	doc = load('original/template.odt')
	tables = doc.getElementsByType(Table)

	pprint([list(map(lambda c: c.attributes, t.childNodes)) for t in tables])
