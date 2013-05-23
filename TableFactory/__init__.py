# Copyright (c) 2013 Kirk Strauser

# Permission is hereby granted, free of charge, to any person
# obtaining a copy of this software and associated documentation files
# (the "Software"), to deal in the Software without restriction,
# including without limitation the rights to use, copy, modify, merge,
# publish, distribute, sublicense, and/or sell copies of the Software,
# and to permit persons to whom the Software is furnished to do so,
# subject to the following conditions:

# The above copyright notice and this permission notice shall be
# included in all copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT.  IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
# BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
# ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
# CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

"""
TableFactory is a high-level frontend to several table generators

It provides a common API for creating HTML, PDF, or spreadsheet tables
from common Python data sources. For example, this is a working
example on my development system:

# This creates a row with two columns:
rowmaker = RowSpec(ColumnSpec('customer', 'Customer'),
                   ColumnSpec('invamt', 'Invoice Amount'))

# Fetch 10 invoices from our database and convert them to TableRow
# objects
lines = rowmaker.makeall(session.query(Invoice).limit(10))

# Make a PDF out of those lines:
table1 = PDFTable('Invoice amounts by customer', headers=rowmaker)
open('invoicetable.pdf', 'wb').write(table1.render(lines))

# Want to make a spreadsheet from the same data? The API is identical:
table2 = SpreadsheetTable('Invoice amounts by customer', headers=rowmaker)
open('invoicetable.xls', 'wb').write(table2.render(lines))

# Inside a Pyramid view callable and want to create an HTML table that
# can be rendered in a template? It's exactly like the first two
# examples:
table3 = HTMLTable('Invoice amounts by customer', headers=rowmaker)
return {'tablecontents': table3.render(lines)}
"""

__author__ = "Kirk Strauser"
__copyright__ = "Copyright 2013, Daycos and Kirk Strauser"
__credits__ = ["Kirk Strauser"]
__license__ = "MIT License"
__email__ = "kirk@strauser.com"
__version__ = '0.2'


from TableFactory.layout import (
    Cell, ColumnSpec, RowSpec, StyleAttributes, TableRow)
from TableFactory.htmltable import HTMLTable
from TableFactory.pdftable import PDFTable
from TableFactory.spreadsheettable import SpreadsheetTable
