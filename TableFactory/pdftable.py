"""Implementation of TableBase that generates PDF tables"""

import StringIO

from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer
from reportlab.platypus.tables import TableStyle, Table

from TableFactory import layout
from TableFactory.base import TableBase


class PDFTable(TableBase):  # pylint: disable=R0903
    """Table generator that yields a PDF representation of the data"""

    rowoddcolor = colors.Color(.92, .92, .92)
    gridcolor = colors.Color(.8, .8, .8)
    rowevencolor = colors.Color(.98, .98, .98)
    headerbackgroundcolor = colors.Color(.004, 0, .5)

    # Every table starts off with this style
    tablebasestyle = TableStyle([
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('INNERGRID', (0, 0), (-1, -1), 1, gridcolor),
    ])

    # The parent table is the outside wrapper around everything
    tableparentstyle = TableStyle([
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), [rowoddcolor, rowevencolor]),
        ('LINEABOVE', (0, 1), (-1, -2), 1, colors.black),
        ('LINEBELOW', (0, 1), (-1, -2), 1, colors.black),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
    ])

    # Give content rows a little bit of side padding
    tablerowstyle = TableStyle([
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
    ])

    tableheaderstyle = TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), headerbackgroundcolor),
    ])

    titlestyle = ParagraphStyle(
        name='Title Style', fontName='Helvetica-Bold', fontSize=16)
    explanationstyle = ParagraphStyle(
        name='Explanation Style', fontName='Helvetica', fontSize=12)
    headercellstyle = ParagraphStyle(
        name='Table Header Style', fontName='Helvetica-Bold',
        textColor=colors.white)
    contentcellstyle = ParagraphStyle(
        name='Table Cell Style', fontName='Helvetica', fontSize=8)
    contentmoneycellstyle = ParagraphStyle(
        name='Table Cell Style', fontName='Helvetica', fontSize=8,
        alignment=TA_RIGHT)

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

    def render(self, rowsets):  # pylint: disable=R0914
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
            if isinstance(rowset, layout.TableRow):
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
