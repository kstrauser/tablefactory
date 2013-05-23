#!/usr/bin/env python


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
        None: {None: xlwt.easyxf()},
        datetime.date: {None: xlwt.easyxf(num_format_str='YYYY-MM-DD')},
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
