#!/usr/bin/env python

"""Classes that describe report layouts"""

from reportlab.lib.units import inch


class StyleAttributes(object):  # pylint: disable=R0903
    """StyleAttribute objects represent the formatting that will be
    applied to a cell. Current properties are:

          bold: bool, display a cell in bold

          money: bool, display the cell right-aligned

          width: float, the width of a column in inches

          span: integer, the number of columns the cell should span

          raw: bool, use the cell's contents as-is without escaping
          them

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


class Cell(object):  # pylint: disable=R0903
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


class TableRow(object):  # pylint: disable=R0903
    """A TableRow is a list of cells"""

    def __init__(self, *cells):
        """Store the given list of cells"""
        self.cells = cells

    def __repr__(self):
        """Human-readable TableRow representation"""
        return '<TableRow(%s)>' % unicode(self.cells)

    def __iter__(self):
        """Return each of the row's cells in turn"""
        for cell in self.cells:
            yield cell


class ColumnSpec(object):  # pylint: disable=R0903
    """A ColumnSpec describes the source of values for a particular
    column, as well as the properties of each of its cells"""

    def __init__(self, attribute, title=None, **properties):
        """'attribute' is the name of the attribute or dictionary key
        that will be pulled from a row object to find a cell's
        value. If 'attribute' is a tuple, each of its elements will be
        resolved in turn, recursively. For example, an attribute tuple
        of ('foo', 'bar', 'baz') might resolve to:

        >>> rowobject['foo'].bar['baz']

        If this ColumnSpec is printed as part of a table header it
        will be captioned with 'title', which defaults to the value of
        'attribute'. Any properties will be applied to cells created
        by this ColumnSpec."""

        if isinstance(attribute, tuple):
            self.attributes = attribute
        else:
            self.attributes = (attribute,)
        if title:
            self.title = title
        else:
            self.title = attribute
        self.style = StyleAttributes(**properties)

    def __repr__(self):
        """Human-readable ColumnSpec representation"""
        return '<ColumnSpec(%s)>' % self.title


class RowSpec(object):  # pylint: disable=R0903
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
        return '<RowSpec(%s)>' % unicode(self.columnspecs)

    def __call__(self, rowobject):
        """A RowSpec can be used as a factory that can take an object
        like a dict or SQLAlchemy row, apply each of the ColumnSpecs
        to that object in turn, and return a corresponding TableRow
        object."""
        output = []
        for column in self.columnspecs:
            value = rowobject
            for attribute in column.attributes:
                try:
                    value = value[attribute]
                except (KeyError, TypeError):
                    value = getattr(value, attribute)
            output.append(Cell(value, column.style))
        return TableRow(*output)  # pylint: disable=W0142

    def __iter__(self):
        """Return each of the row's ColumnSpecs in turn"""
        for columnspec in self.columnspecs:
            yield columnspec

    def makeall(self, rowobjects):
        """Create a list of TableRows from a list of source objects"""
        return [self(rowobject) for rowobject in rowobjects]
