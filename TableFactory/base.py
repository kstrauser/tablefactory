#!/usr/bin/env python


class TableBase(object):
    """Base class implementing common functionality for all table
    classes."""

    castfunctions = {}

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
        value = cell.value
        if cell.style.raw:
            return value
        if value is None:
            return ''
        castfunction = self.castfunctions.get(type(value), unicode)
        return cgi.escape(castfunction(value))
