from __future__ import absolute_import

from openpyxl.compat.strings import safe_repr

from openpyxl.utils import (
    range_boundaries,
    range_to_tuple,
    get_column_letter,
    quote_sheetname,
)


### Notes
# focus on range functions


class CellRange(object):
    """
    Represents a range in a sheet: title and coordinates.

    This object is used to perform calculations on ranges, like:

    - shifting to a direction,
    - union/intersection with another sheet range,
    - collapsing to a 1 x 1 range,
    - expanding to a given size.

    We can check whether a range is:

    - equal or not equal to another,
    - disjoint of another,
    - contained in another.

    We can get:

    - the string representation,
    - the coordinates,
    - the size of a range.
    """


    def __init__(self, range_string=None, min_col=None, min_row=None,
                 max_col=None, max_row=None, title=None):
        if range_string is not None:
            try:
                title, (min_col, min_row, max_col, max_row) = range_to_tuple(range_string)
            except ValueError:
                min_col, min_row, max_col, max_row = range_boundaries(range_string)
        # None > 0 is False
        if not all(idx > 0 for idx in (min_col, min_row, max_col, max_row)):
            msg = "Values for 'min_col', 'min_row', 'max_col' *and* 'max_row_' " \
                  "must be strictly positive"
            raise ValueError(msg)
        # Intervals are inclusive
        if not min_col <= max_col:
            fmt = "{max_col} must be greater than {min_col}"
            raise ValueError(fmt.format(min_col=min_col, max_col=max_col))
        if not min_row <= max_row:
            fmt = "{max_row} must be greater than {min_row}"
            raise ValueError(fmt.format(min_row=min_row, max_row=max_row))
        self.min_col = min_col
        self.min_row = min_row
        self.max_col = max_col
        self.max_row = max_row
        self.title = title


    @property
    def bounds(self):
        """
        Vertices of the range as a tuple
        """
        return self.min_col, self.min_row, self.max_col, self.max_row


    @property
    def coord(self):
        """
        Excel style representation of the range
        """
        fmt = "{min_col}{min_row}:{max_col}{max_row}"
        if (self.min_col == self.max_col
            and self.min_row == self.max_row):
            fmt = "{min_col}{min_row}"

        return fmt.format(
            min_col=get_column_letter(self.min_col),
            min_row=self.min_row,
            max_col=get_column_letter(self.max_col),
            max_row=self.max_row
        )


    def __repr__(self):
        fmt = u"{coord}"
        if self.title:
            fmt = u"{title!r}!{coord}"
        return safe_repr(fmt.format(title=self.title, coord=self.coord))


    def _get_range_string(self):
        fmt = "{coord}"
        title = self.title
        if self.title:
            fmt = u"{title}!{coord}"
            title = quote_sheetname(self.title)
        return fmt.format(title=title, coord=self.coord)

    __unicode__ = _get_range_string


    def __str__(self):
        coord = self._get_range_string()
        return safe_repr(coord)


    def __copy__(self):
        return self.__class__(min_col=self.min_col, min_row=self.min_row,
                              max_col=self.max_col, max_row=self.max_row,
                              title=self.title)


    def shift(self, col_shift=0, row_shift=0):
        """
        Shift the range according to the shift values (*col_shift*, *row_shift*).

        :type col_shift: int
        :param col_shift: number of columns to be moved by, can be negative
        :type row_shift: int
        :param row_shift: number of rows to be moved by, can be negative
        :raise: :class:`ValueError` if any row or column index < 1
        """

        if (self.min_col + col_shift <= 0
            or self.min_row + row_shift <= 0):
            raise ValueError("Invalid shift value: col_shift={0}, row_shift={1}".format(col_shift, row_shift))
        self.min_col += col_shift
        self.min_row += row_shift
        self.max_col += col_shift
        self.max_row += row_shift


    def __iadd__(self, value):
        col_shift, row_shift = value
        self.shift(col_shift, row_shift)
        return self


    def __add__(self, other):
        return self.__copy__().__iadd__(other)


    def __ne__(self, other):
        """
        Test whether the ranges are not equal.

        :type other: CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* != *other*.
        """
        if isinstance(other, CellRange):
            # Test whether sheet titles are different and not empty.
            this_title = self.title
            that_title = other.title
            ne_sheet_title = this_title and that_title and this_title.upper() != that_title.upper()
            return (ne_sheet_title or
                    other.min_row != self.min_row or self.max_row != other.max_row or
                    other.min_col != self.min_col or self.max_col != other.max_col)
        raise TypeError(repr(type(other)))


    def __eq__(self, other):
        """
        Test whether the ranges are equal.

        :type other: CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* == *other*.
        """
        return not self.__ne__(other)


    def issubset(self, other):
        """
        Test whether every element in the range is in *other*.

        :type other: SheetRange
        :param other: Other sheet range
        :return: ``True`` if *range* <= *other*.
        """
        if not isinstance(other, CellRange):
            raise TypeError(repr(type(other)))

        # Test whether sheet titles are equals (or if one of them is empty).
        this_title = self.title
        that_title = other.title
        eq_sheet_title = not this_title or not that_title or this_title.upper() == that_title.upper()
        return (eq_sheet_title and
                (other.min_row <= self.min_row <= self.max_row <= other.max_row) and
                (other.min_col <= self.min_col <= self.max_col <= other.max_col))


    __le__ = issubset


    def __lt__(self, other):
        """
        Test whether every element in the range is in *other*, but not all.

        :type other: CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* < *other*.
        """
        return self.__le__(other) and self.__ne__(other)


    def issuperset(self, other):
        """
        Test whether every element in *other* is in the range.

        :type other: CellRange or tuple[int, int]
        :param other: Other sheet range or cell index (*row_idx*, *col_idx*).
        :return: ``True`` if *range* >= *other* (or *other* in *range*).
        """
        return not self.issubset(other)

    __contains__ = __ge__ = issuperset


    def __gt__(self, other):
        """
        Test whether every element in *other* is in the range, but not all.

        :type other: CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* > *other*.
        """
        return self.__ge__(other) and self.__ne__(other)


    def isdisjoint(self, other):
        """
        Return ``True`` if the range has no elements in common with other.
        Ranges are disjoint if and only if their intersection is the empty range.

        :type other: CellRange
        :param other: Other sheet range.
        :return: `True`` if the range has no elements in common with other.
        """
        if not isinstance(other, CellRange):
            raise TypeError(repr(type(other)))

        if self.title != other.title:
            raise ValueError("Cannot compare ranges from different worksheets")

        # sort by top-left vertex
        if self.bounds > other.bounds:
            i = self
            self = other
            other = i

        return (self.max_col, self.max_row) < (other.min_col, other.max_row)


    def intersection(self, other):
        """
        Return a new range with elements common to the range and another

        :type others: tuple[CellRange]
        :param others: Other sheet ranges.
        :return: the current sheet range.
        :raise: :class:`ValueError` if an *other* range don't intersect
            with the current range.
        """
        if not isinstance(other, CellRange):
            raise TypeError(repr(type(other)))

        if self.title != other.title:
            raise ValueError("Cannot compare ranges from different worksheets")

        if self.isdisjoint(other):
            raise ValueError("Range {0} don't intersect {0}".format(self, other))

        min_row = max(self.min_row, other.min_row)
        max_row = min(self.max_row, other.max_row)
        min_col = max(self.min_col, other.min_col)
        max_col = min(self.max_col, other.max_col)

        return CellRange(min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)

    __iand__ = intersection


    def __and__(self, other):
        return self.__copy__().__iand__(other)


    def union(self, other):
        """
        Return a new range with elements from the range and all *others*.

        :type others: tuple[CellRange]
        :param others: Other sheet ranges.
        :return: the current sheet range.
        """
        if not isinstance(other, CellRange):
            raise TypeError(repr(type(other)))

        if self.title != other.title:
            raise ValueError("Cannot merge ranges from different worksheets")

        min_row = min(self.min_row, other.min_row)
        max_row = max(self.max_row, other.max_row)
        min_col = min(self.min_col, other.min_col)
        max_col = max(self.max_col, other.max_col)
        return CellRange(min_col=min_col, min_row=min_row, max_col=max_col,
                         max_row=max_row, title=self.title)


    __ior__ = union


    def __or__(self, other):
        return self.__copy__().__ior__(other)


    def expand(self, right=0, down=0, left=0, up=0):
        """
        Expand the range by the dimensions provided.

        :type right: int
        :param right: expand range to the right by this number of cells
        :type down: int
        :param down: expand range down by this number of cells
        :type left: int
        :param left: expand range to the left by this number of cells
        :type up: int
        :param up: expand range up by this number of cells
        """
        self.min_col -= left
        self.min_row -= up
        self.max_col += right
        self.max_row += down


    def shrink(self, right=0, bottom=0, left=0, top=0):
        """
        Shrink the range by the dimensions provided.

        :type right: int
        :param right: shrink range from the right by this number of cells
        :type down: int
        :param down: shrink range from the top by this number of cells
        :type left: int
        :param left: shrink range from the left by this number of cells
        :type up: int
        :param up: shrink range from the bottown by this number of cells
        """
        self.min_col += left
        self.min_row += top
        self.max_col -= right
        self.max_row -= bottom


    @property
    def size(self):
        """ Return the size of the range as a dicitionary of rows and columns. """
        cols = self.max_col + 1 - self.min_col
        rows = self.max_row + 1 - self.min_row
        return {'columns':cols, 'rows':rows}
