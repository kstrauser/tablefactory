#!/usr/bin/env python

"""Implementation of TableBase that generates HTML tables"""

from TableFactory import layout
from TableFactory.base import TableBase


class HTMLTable(TableBase):  # pylint: disable=R0903
    """Table generator that yields an HTML representation of the
    data. Note that this class yields *only* the table itself and not
    an entire HTML document.

    The CSS classes are compatible with jQuery's tablesorter plugin
    <http://tablesorter.com/docs/>. With this combination, all
    generated tables can be sorted in a client's browser just by
    clicking on the column headers.

    When a rowset is made of multiple TableRow objects, all rows after
    the first are additionally assigned the 'childrow' CSS class. This
    adds compatibility with the "Children Rows" mod to tablesorter
    <http://www.pengoworks.com/workshop/jquery/tablesorter/tablesorter.htm>,
    which groups child rows with their parent rows when sorting.

    For example, the following lines in a page's <head> section will
    enable all of those client-side options:

        <script type="text/javascript"
            src="/javascript/jquery-1.5.min.js"></script>
        <script
            type="text/javascript"
            src="/javascript/jquery.tablesorter.min.js"></script>
        <script
            type="text/javascript"
            src="/javascript/jquery.tablesorter.mod.js"></script>
        <script type="text/javascript">
        $(document).ready(function()
            {
                $(".reporttable").tablesorter({widgets: ['zebra']});
            }
        );
    """

    # These are the CSS classes emitted by the renderer
    cssdefs = {
        'bold': 'cell_bold',
        'money': 'cell_money',
        'table': 'reporttable',
        'childrow': 'expand-child',
        'zebra': ('odd', 'even'),
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
        return '<td%s%s>%s</td>' % (
            cssstring, colspanstring, self._cast(cell).replace('\r', '<br />'))

    def render(self, rowsets):  # pylint: disable=R0912
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
            lines.append('<table summary="%s" class="%s">' % (
                self.title, self.cssdefs['table']))
        else:
            lines.append('<table class="%s">' % self.cssdefs['table'])

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
            if isinstance(rowset, layout.TableRow):
                rowset = [rowset]
            for subrowindex, subrow in enumerate(rowset):
                trclasses = [self.cssdefs['zebra'][rowsetindex % 2]]
                if subrowindex:
                    trclasses.append(self.cssdefs['childrow'])
                lines.append('    <tr class="%s">' % ' '.join(trclasses))
                for cell in subrow:
                    lines.append('      %s' % self._rendercell(cell))
                lines.append('    </tr>')

        # Finish up
        lines.append('  </tbody>')
        lines.append('</table>')
        return '\n'.join(lines)
