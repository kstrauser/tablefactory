#!/usr/bin/env python

from TableFactory import *

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

    def makelines(linecount):
        lines = []
        lines.append([mainrs({'foo': 1, 'bar': 2, 'baz': 3, 'qux': 'Bar.  ' * 30}),
                      subrow1({'subfoo1': 'And', 'subbar1': 'another'}),
                      subrow2({'subfoo2': 'This is a test.  ' * 15}),
                      subrow2({'subfoo2': 'And another test'})])
        for i in range(linecount):
            lines.append(mainrs({'foo': i, 'bar': 14, 'baz': 15, 'qux': 'extra'}))
        lines.append(summaryrow({'junk1': None, 'baz': 'Summary!', 'junk2': None}))
        return lines

    firsttable = PDFTable(
        'Sample Table 1',
        'Multi-Table PDF test (first table)',
        headers=[mainrs, subrow1, subrow2]).prepare(makelines(20))
    report = PDFReport(firsttable)

    for i in range(5):
        subtable = PDFTable(
            'Sample Table %d' % (i + 2),
            'Multi-Table PDF test (following table #%d' % (i + 1),
            headers=[mainrs, subrow1, subrow2]).prepare(
                makelines(((i + 1) * 73) % 40))
        report.add(subtable)

    with open('multitable.pdf', 'wb') as outfile:
        outfile.write(report.render())


if __name__ == '__main__':
    for _ in range(1):
        example()
