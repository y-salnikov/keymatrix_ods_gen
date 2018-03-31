#!/usr/bin/env python
# -*- coding: utf-8 -*-

from odf.opendocument import OpenDocumentSpreadsheet
from odf.style import Style, TextProperties, TableColumnProperties, Map, ParagraphProperties, TableCellProperties
from odf.number import NumberStyle, CurrencyStyle, CurrencySymbol,  Number,  Text
from odf.text import P
from odf.table import Table, TableColumn, TableRow, TableCell


MSIZE=26

def lit(n):
	return chr(ord('A')-1+n)


textdoc = OpenDocumentSpreadsheet()
# Create a style for the table content. One we can modify
# later in the spreadsheet.
tablecontents = Style(name="Large number", family="table-cell")
tablecontents.addElement(TextProperties(fontfamily="Arial", fontsize="10pt"))
textdoc.styles.addElement(tablecontents)

# Create automatic styles for the column widths.
widewidth = Style(name="co1", family="table-column")
widewidth.addElement(TableColumnProperties(columnwidth="1cm", breakbefore="auto"))
textdoc.automaticstyles.addElement(widewidth)

cs1=Style(name="cs1",family="table-cell")
cs1.addElement(ParagraphProperties(textalign='center'))
cs1.addElement(TextProperties(attributes={'fontsize':"10pt",'fontweight':"bold", 'color':"#CCCCCC" }))
cs1.addElement(TableCellProperties(backgroundcolor="#ff0000",border="0.74pt solid #000000"))
textdoc.automaticstyles.addElement(cs1)


cs2=Style(name="cs2",family="table-cell")
cs2.addElement(ParagraphProperties(textalign='center'))
cs2.addElement(TextProperties(attributes={'fontsize':"10pt", 'color':"#000000" }))
cs2.addElement(TableCellProperties(backgroundcolor="#ffff80",border="0.74pt solid #000000"))
textdoc.automaticstyles.addElement(cs2)


cs3=Style(name="cs3",family="table-cell")
cs3.addElement(ParagraphProperties(textalign='center'))
cs3.addElement(TextProperties(attributes={'fontsize':"10pt", 'color':"#000000" }))
cs3.addElement(TableCellProperties(border="0.74pt solid #000000"))
textdoc.automaticstyles.addElement(cs3)


cs4=Style(name="cs4",family="table-cell")
cs4.addElement(ParagraphProperties(textalign='center'))
cs4.addElement(TextProperties(attributes={'fontsize':"10pt",'fontweight':"bold", 'color':"#000000" }))
cs4.addElement(TableCellProperties(backgroundcolor="#C0C0C0",border="0.74pt solid #000000"))
textdoc.automaticstyles.addElement(cs4)



# Start the table, and describe the columns
table = Table(name="KeyMatrix")


for col in xrange(1,MSIZE+1):
	# Create a column (same as <col> in HTML) Make all cells in column default to currency
	table.addElement(TableColumn(stylename=widewidth, defaultcellstylename="ce1"))

for row in xrange(1,MSIZE+1):
	# Create a row (same as <tr> in HTML)
	tr = TableRow()
	table.addElement(tr)

	for col in xrange(1,MSIZE+1):
		if(col==row):
			cell = TableCell(valuetype="string",stylename="cs1")
			cell.addElement(P(text=u"X")) # The current displayed value
		elif (col<row):
			cell = TableCell(formula=u"=%s%s" %(lit(row),col),valuetype="string",stylename="cs2")
		else:
			cell = TableCell(valuetype="string", stylename="cs3")
			cell.addElement(P(text=u"")) # The current displayed value
		tr.addElement(cell)

tr = TableRow()
table.addElement(tr)
for col in xrange(1,MSIZE+1):
	cell = TableCell(valuetype="string",stylename="cs4")
	cell.addElement(P(text=u"%s" %(col))) # The current displayed value
	tr.addElement(cell)

textdoc.spreadsheet.addElement(table)
textdoc.save("keymatrix.ods")
