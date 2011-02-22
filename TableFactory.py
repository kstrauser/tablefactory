#!/usr/bin/env python

# Copyright (c) 2011 Kirk Strauser

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

"""TableFactory is a high-level frontend to several table generators

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
__copyright__ = "Copyright 2011, Daycos"
__credits__ = ["Kirk Strauser"]
__license__ = "GPLv3"
__maintainer__ = "Kirk Strauser"
__email__ = "kirk@strauser.com"
__status__ = "Production"

import cgi
import copy
import datetime
import StringIO

import xlwt
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer
from reportlab.platypus.tables import TableStyle, Table

class StyleAttributes(object):
    """StyleAttribute objects represent the formatting that will be
    applied to a cell. Current properties are:

          bold: bool, display a cell in bold

          money: bool, display the cell right-aligned
    
          width: float, the width of a column in inches

          span: integer, the number of columns the cell should span

    By acting as a thin wrapper around a dict and deferring
    calculations until they're needed, we don't do any unnecessary
    work or have to worry about values being updated after they're
    calculated."""
    
    def __init__(self, **properties):
        """Save the value of keyword/dict properties"""
        self.properties = properties

    def __getattr__(self, key):
        """Return the requested property after applying appropriate
        processing to it"""
        value = self.properties.get(key, None)
        if key == 'width' and value is not None:
            return value * inch
        if key == 'span' and value is None:
            return 1
        return value
    
class Cell(object):
    """Cell objects represent a single table cell"""

    def __init__(self, value, style=None):
        """'value' is the displayed value of the cell. 'properties' is
        a dict of cell styles that each table generator may interpret
        as appropriate."""
        
        self.value = value
        if style is None:
            self.style = StyleAttributes()
        else:
            self.style = style

    def __repr__(self):
        """Human-readable Cell representation"""
        return '<Cell(%s)>' % self.value

class TableRow(object):
    """A TableRow is a list of cells"""

    def __init__(self, *cells):
        """Store the given list of cells"""
        self.cells = cells

    def __repr__(self):
        """Human-readable TableRow representation"""
        return '<TableRow(%s)>' % str(self.cells)

    def __iter__(self):
        """Return each of the row's cells in turn"""
        for cell in self.cells:
            yield cell
    
class ColumnSpec(object):
    """A ColumnSpec describes the source of values for a particular
    column, as well as the properties of each of its cells"""
    
    def __init__(self, attribute, title=None, **properties):
        """'attribute' is the name of the attribute or dictionary key
        that will be pulled from a row object to find a cell's
        value. If this ColumnSpec is printed as part of a table header
        it will be captioned with 'title', which defaults to the value
        of 'attribute'. Any properties will be applied to cells
        created by this ColumnSpec."""
        
        self.attribute = attribute
        if title:
            self.title = title
        else:
            self.title = attribute
        self.style = StyleAttributes(**properties)

    def __repr__(self):
        """Human-readable ColumnSpec representation"""
        return '<ColumnSpec(%s)>' % self.title

class RowSpec(object):
    """A RowSpec is a list of ColumnSpecs. It has two main uses:

    1) When passed to a table generator as the 'headers' argument
    (possibly in a list of other RowSpecs), its ColumnSpecs form the
    title row of a table.

    2) As a callable, it creates TableRow objects from various Python
    objects that are passed into it, saving you the trouble of
    building them manually. This is the recommended method of creating
    TableRows as it's easy and it also guarantees that your column
    titles (see #1 above) will match their contents.
    """
    
    def __init__(self, *columnspecs):
        """Store the given list of ColumnSpecs"""
        self.columnspecs = columnspecs

    def __repr__(self):
        """Human-readable RowSpec representation"""
        return '<RowSpec(%s)>' % str(self.columnspecs)

    def __call__(self, rowobject):
        """A RowSpec can be used as a factory that can take an object
        like a dict or SQLAlchemy row, apply each of the ColumnSpecs
        to that object in turn, and return a corresponding TableRow
        object."""
        output = []
        for column in self.columnspecs:
            try:
                value = rowobject[column.attribute]
            except (KeyError, TypeError):
                value = getattr(rowobject, column.attribute)
            output.append(Cell(value, column.style))
        return TableRow(*output)

    def __iter__(self):
        """Return each of the row's ColumnSpecs in turn"""
        for columnspec in self.columnspecs:
            yield columnspec

    def makeall(self, rowobjects):
        """Create a list of TableRows from a list of source objects"""
        return [self(rowobject) for rowobject in rowobjects]

class TableBase(object):
    """Base class implementing common functionality for all table
    classes."""

    def __init__(self, title=None, explanation=None, headers=None):
        """A rowset is either a TableRow or a collection of
        TableRows. 'rowsets' is a collection of rowsets. Passing
        multiple rows as a single rowset has two main advantages:

        1) The HTMLTable and PDFTable classes use alternating row
        colors, and each row in a rowset gets the same color. For
        example, suppose the first row in each rowset contains a list
        of detailed columns, and the second row is a note explaining
        the first row. By passing them together as single rowsets,
        both rows will be colored alike and the colors will alternate
        after every other row in the table.

        2) The PDFTable class will do its best not to break up rowsets
        across page boundaries.

        'title' is the table's optional title.

        'explanation', if given, will usually be displayed below the
        title.

        'headers' is a RowSpec or a collection of RowSpecs used to
        generate the table's header. If more than one RowSpec is
        given, each will be rendered in order as a header row.
        """
        self.title = title
        self.explanation = explanation
        if isinstance(headers, RowSpec):
            self.headers = [headers]
        else:
            self.headers = headers

    def _cast(self, cell):
        """This doesn't do a lot right now, but this is where we'd
        implement code to convert various datatypes to their desired
        output format"""
        if cell.value is None:
            return ''
        else:
            return cgi.escape(str(cell.value))

class PDFTable(TableBase):
    """Table generator that yields a PDF representation of the data"""

    rowoddcolor = colors.Color(.92, .92, .92)
    gridcolor = colors.Color(.8, .8, .8)
    rowevencolor = colors.Color(.98, .98, .98)
    headerbackgroundcolor = colors.Color(.004, 0, .5)

    # Every table starts off with this style
    tablebasestyle = TableStyle([
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('INNERGRID', (0,0), (-1,-1), 1, gridcolor),
            ])

    # The parent table is the outside wrapper around everything
    tableparentstyle = TableStyle([
            ('ROWBACKGROUNDS', (0,0), (-1,-1), [rowoddcolor, rowevencolor]),
            ('LINEABOVE', (0,1), (-1,-2), 1, colors.black),
            ('LINEBELOW', (0,1), (-1,-2), 1, colors.black),
            ('BOX', (0,0), (-1,-1), 1, colors.black),
            ])

    # Give content rows a little bit of side padding
    tablerowstyle = TableStyle([
            ('LEFTPADDING', (0,0), (-1,-1), 3),
            ('RIGHTPADDING', (0,0), (-1,-1), 3),
            ])

    tableheaderstyle = TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), headerbackgroundcolor),
            ])

    titlestyle = ParagraphStyle(name='Title Style', fontName='Helvetica-Bold', fontSize=16)
    explanationstyle = ParagraphStyle(name='Explanation Style', fontName='Helvetica', fontSize=12)
    headercellstyle = ParagraphStyle(name='Table Header Style', fontName='Helvetica-Bold', textColor=colors.white)
    contentcellstyle = ParagraphStyle(name='Table Cell Style', fontName='Helvetica', fontSize=8)
    contentmoneycellstyle = ParagraphStyle(name='Table Cell Style', fontName='Helvetica', fontSize=8, alignment=TA_RIGHT)

    def _rendercell(self, cell):
        """Render data as a Paragraph"""

        value = self._cast(cell)
        
        # Wrap the cell's contents in onion-tag goodness
        if cell.style.bold:
            value = '<b>%s</b>' % value

        if cell.style.money:
            style = self.contentmoneycellstyle
        else:
            style = self.contentcellstyle
        return Paragraph(value, style)
    
    def render(self, rowsets):
        """Return the data as a binary string holding a PDF"""

        # Start by creating the table headers
        rowtables = []
        if self.headers:
            for headerrow in self.headers:
                widths = [headercolumn.style.width for headercolumn in headerrow]
                # Let ReportLab calculate the width of the last column
                # so that it occupies the total remaining open space
                widths[-1] = None
                headertable = Table([[Paragraph(headercolumn.title, self.headercellstyle)
                                      for headercolumn in headerrow]],
                                    style=self.tablebasestyle,
                                    colWidths=widths)
                headertable.setStyle(self.tablerowstyle)
                headertable.setStyle(self.tableheaderstyle)
                rowtables.append([headertable])

        # Then create a table to hold the contents of each line
        for rowset in rowsets:
            subrowtables = []
            if isinstance(rowset, TableRow):
                rowset = [rowset]
            for subrow in rowset:
                subrowtable = Table([[self._rendercell(cell) for cell in subrow]],
                                    style=self.tablebasestyle,
                                    colWidths=[cell.style.width for cell in subrow])
                subrowtable.setStyle(self.tablerowstyle)
                subrowtables.append([subrowtable])

            rowtable = Table(subrowtables, style=self.tablebasestyle)
            rowtables.append([rowtable])

        # Wrap all of those rows into an outer table
        parenttable = Table(rowtables, style=self.tablebasestyle, repeatRows=1)
        parenttable.setStyle(self.tableparentstyle)

        # Finally, build the list of elements that the table will
        # comprise
        components = []
        if self.title:
            components.append(Paragraph(self.title, self.titlestyle))
        if self.explanation:
            components.extend([Spacer(1, .2 * inch),
                               Paragraph(self.explanation, self.explanationstyle)])
        components.extend([Spacer(1, .3 * inch), parenttable])

        # Compile the whole thing and return the results
        stringbuf = StringIO.StringIO()
        doc = SimpleDocTemplate(stringbuf,
                                bottomMargin=.5 * inch, topMargin=.5 * inch,
                                rightMargin=.5 * inch, leftMargin=.5 * inch)
        doc.build(components)
        return stringbuf.getvalue()
            
class SpreadsheetTable(TableBase):
    """Table generator that yields an Excel spreadsheet representation
    of the data. It will have one worksheet named with the given
    title"""

    headerstyle = xlwt.easyxf('pattern: pattern solid, fore_colour blue;'
                              'font: colour white, bold True;')
    explanationstyle = xlwt.easyxf('font: bold True;')

    # Styles to apply to given data types. The style for None is the
    # default when no other type is applicable.
    styletypemap = {
        None             : {None: xlwt.easyxf()},
        datetime.date    : {None: xlwt.easyxf(num_format_str='YYYY-MM-DD')},
        datetime.datetime: {None: xlwt.easyxf(num_format_str='YYYY-MM-DD HH:MM:SS')},
        }

    def _getstyle(self, cell):
        """Return the appropriate style for a cell, generating it
        first if it doesn't already exist"""
        try:
            cellstyles = self.styletypemap[type(cell.value)]
        except KeyError:
            cellstyles = self.styletypemap[None]
        # Build the list of desired attributes
        attrs = set()
        if cell.style.bold:
            attrs.add('bold')
        if cell.style.money:
            attrs.add('money')

        # Use the cached value if we've already seen this style
        try:
            return cellstyles[tuple(sorted(attrs))]
        except KeyError:
            pass

        # Build the style by copying the default and applying each of
        # the requested attributes
        cellstyle = copy.deepcopy(cellstyles[None])
        if 'bold' in attrs:
            cellstyle.font.bold = 1
        if 'money' in attrs:
            cellstyle.alignment.horz = xlwt.Alignment.HORZ_RIGHT
            cellstyle.num_format_str = '0.00'

        # Cache the results for next time
        cellstyles[tuple(sorted(attrs))] = cellstyle
        return cellstyle
    
    def render(self, rowsets):
        """Return the data as a binary string holding an Excel spreadsheet"""
        book = xlwt.Workbook()
        mainsheet = book.add_sheet(self.title or 'Sheet 1')
        rownum = 0

        if self.explanation:
            mainsheet.write(rownum, 0, self.explanation, self.explanationstyle)
            # Clear the first row's color, or else the headerstyle
            # will take over. I have no idea why.
            mainsheet.row(0).set_style(self.styletypemap[None][None])
            rownum += 2
        
        # Generate any header rows
        if self.headers:
            for headerrow in self.headers:
                colnum = 0
                for headercolumn in headerrow:
                    mainsheet.write(rownum, colnum, headercolumn.title, self.headerstyle)
                    colnum += headercolumn.style.span
                mainsheet.row(rownum).set_style(self.headerstyle)
                rownum += 1

        # Write every line
        for rowset in rowsets:
            if isinstance(rowset, TableRow):
                rowset = [rowset]
            for subrow in rowset:
                colnum = 0
                row = mainsheet.row(rownum)
                for cell in subrow:
                    row.write(colnum, cell.value, self._getstyle(cell))
                    colnum += cell.style.span
                rownum += 1

        stringbuf = StringIO.StringIO()
        book.save(stringbuf)
        return stringbuf.getvalue()
            
class HTMLTable(TableBase):
    """Table generator that yields an HTML representation of the
    data. Note that this class yields *only* the table itself and not
    an entire HTML document."""

    # These are the CSS classes emitted by the generator
    cssdefs = {
        'bold'      : 'cell_bold',
        'money'     : 'cell_money',
        'tablestyle': 'reporttable',
        }
    
    def _rendercell(self, cell):
        """Render data as a td"""

        cssclasses = []
        if cell.style.bold:
            cssclasses.append(self.cssdefs['bold'])
        if cell.style.money:
            cssclasses.append(self.cssdefs['money'])
        if cssclasses:
            cssstring = ' class="%s"' % ' '.join(cssclasses)
        else:
            cssstring = ''
        colspan = cell.style.span
        if colspan > 1:
            colspanstring = ' colspan="%d"' % colspan
        else:
            colspanstring = ''
        return '<td%s%s>%s</td>' % (cssstring, colspanstring, self._cast(cell).replace('\r', '<br />'))

    def render(self, rowsets):
        """Return the data as a string of HTML"""
        lines = []

        # Display the title, if given
        if self.title:
            lines.append('<h2>%s</h2>' % self.title)

        # Display the explanation, if given
        if self.explanation:
            lines.append('<p>%s</p>' % self.explanation)

        # Create the table
        if self.title:
            lines.append('<table summary="%s" class="%s">' % (self.title, self.cssdefs['tablestyle']))
        else:
            lines.append('<table style="%s">' % self.cssdefs['tablestyle'])

        # Generate any header rows
        if self.headers:
            lines.append('  <thead>')
            for headerrow in self.headers:
                lines.append('    <tr>')
                for headercolumn in headerrow:
                    span = headercolumn.style.span
                    if span > 1:
                        lines.append('      <th colspan="%d">%s</th>' % (span, headercolumn.title))
                    else:
                        lines.append('      <th>%s</th>' % headercolumn.title)
                lines.append('    </tr>')
                
            lines.append('  </thead>')
        lines.append('  <tbody>')

        # Write every line
        for rowsetindex, rowset in enumerate(rowsets):
            if isinstance(rowset, TableRow):
                rowset = [rowset]
            for subrow in rowset:
                lines.append('    <tr class="tr_%s">' % ('odd', 'even')[rowsetindex % 2])
                for cell in subrow:
                    lines.append('      %s' % self._rendercell(cell))
                lines.append('    </tr>')

        # Finish up
        lines.append('  </tbody>')
        lines.append('</table>')
        return '\n'.join(lines)
    
def example():
    """Create a set of sample tables"""

    # In practice, you'd most likely be embedding your HTML tables in
    # a web page template. For demonstration purposes, we'll create a
    # simple page with a few default styles.
    htmlheader = """\
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>
<head>
<title>Sample table</title>
<style type="text/css">
body { font-family: Helvetica,Arial,FreeSans; }
table.reporttable { border-style: solid; border-width: 1px; }
table.reporttable tr.tr_odd { background-color: #eee; }
table.reporttable tr.tr_even { background-color: #bbb; }
table.reporttable th { background-color: blue; color: white; }
table.reporttable td.cell_bold { font-weight: bold; }
table.reporttable td.cell_money { text-align: right; font-family: monospace; }
</style>
</head>
<body>
"""
    htmlfooter = """\
</body>
</html>"""

    exampletypes = ((PDFTable, 'pdf'), (HTMLTable, 'html'), (SpreadsheetTable, 'xls'))
    
    #### Example with several row types
    
    mainrs = RowSpec(
        ColumnSpec('foo', 'Column 1', width=1),
        ColumnSpec('bar', 'Column 2', width=1),
        ColumnSpec('baz', 'Column 3', width=1),
        ColumnSpec('qux', 'Column 4', width=4))

    subrow1 = RowSpec(
        ColumnSpec('subfoo1', 'First wide column', bold=True, span=2, width=2),
        ColumnSpec('subbar1', 'Second wide column', span=2))

    subrow2 = RowSpec(
        ColumnSpec('subfoo2', 'A table-wide column', span=4))

    summaryrow = RowSpec(
        ColumnSpec('junk1', span=2, width=2),
        ColumnSpec('baz', bold=True, width=1),
        ColumnSpec('junk2'))

    lines = []
    lines.append([mainrs({'foo': 1, 'bar': 2, 'baz': 3, 'qux': 'Bar.  ' * 30}),
                  subrow1({'subfoo1': 'And', 'subbar1': 'another'}),
                  subrow2({'subfoo2': 'This is a test.  ' * 15}),
                  subrow2({'subfoo2': 'And another test'})])
    for i in range(5):
        lines.append(mainrs({'foo': i, 'bar': 14, 'baz': 15, 'qux': 'extra'}))
    lines.append(summaryrow({'junk1': None, 'baz': 'Summary!', 'junk2': None}))

    for tableclass, extension in exampletypes:
        outfile = open('showcase.%s' % extension, 'wb')
        if tableclass is HTMLTable:
            outfile.write(htmlheader)
        outfile.write(tableclass('Sample Table',
                                 '%s test' % extension.upper(),
                                 headers=[mainrs, subrow1, subrow2]).render(lines))
        if tableclass is HTMLTable:
            outfile.write(htmlfooter)

            
    #### Example of a typical "invoices" table

    import decimal
    import random

    # Most common names in the US, according to the census
    names = ['Smith', 'Johnson', 'Williams', 'Jones', 'Brown', 'Davis', 'Miller', 'Wilson',
             'Moore', 'Taylor', 'Anderson', 'Thomas', 'Jackson', 'White', 'Harris', 'Martin',
             'Thompson', 'Garcia']
    random.shuffle(names)
    rows = [{'invoiceid': invoiceid,
             'name': name,
             'amount': decimal.Decimal('%.2f' % (random.randrange(500000) / 100.0))}
            for invoiceid, name in enumerate(names)]

    invoicerow = RowSpec(ColumnSpec('invoiceid', 'Invoice #'),
                         ColumnSpec('name', 'Customer Name'),
                         ColumnSpec('amount', 'Total', money=True))
    lines = invoicerow.makeall(rows)
    
    for tableclass, extension in exampletypes:
        outfile = open('invoice.%s' % extension, 'wb')
        if tableclass is HTMLTable:
            outfile.write(htmlheader)
        outfile.write(tableclass('Invoices by Customer',
                                 'Amount of each invoice, sorted by invoiceid',
                                 headers=invoicerow).render(lines))
        if tableclass is HTMLTable:
            outfile.write(htmlfooter)

if __name__ == '__main__':
    for _ in range(1):
        example()
