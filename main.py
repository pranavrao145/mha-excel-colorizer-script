import xlsxwriter


def colorize_columns(workbook: xlsxwriter.Workbook, sheet_name: str, column_formatting: dict[str, str],
                     margins: tuple[int], colour_options: tuple[str]):
    """
    Applies the given column_formatting, colour_options, and margins to the
    sheet with sheet_name in the given workbook.

    EXAMPLE USAGE:
    Assume there is a workbook which contains a sheet named 'sheet_1'.

    If the following inputs were given to the function:
    - sheet_name = sheet_1
    - column_formatting = 'mean': 'colour_upper', 'missing': 'colour_lower'}
    - margins = (5, 10)
    - colour_options = ('green', 'red')

    Then the following would happen:
    - For the column in sheet_1 called 'mean', the upper 5 percent in the Excel document would be coloured green.
    - For the column in sheet_1 called 'missing', the lower 10 percent in the Excel document would be coloured red.

    Args:
    - workbook: the workbook to edit
    - sheet_name: the name of the sheet to search for the columns in
    - column_formatting: a dictionary of the form {column_name: format_option} where
      column_name is a valid column name in the given sheet and format_option is an
      item from the set { 'colour_both', 'colour_upper', 'colour_lower', 'colour_none'}
    - margins: a tuple of the form (X, Y) where X and Y are integers from 0 to 100 (inclusive).
      X is the percent of the data to consider as the UPPER margin. Y is the percent of the data
      to consider as the LOWER margin.
    - colour_options: a tuple of the form (A, B) where A and B are strings
      from the set {'red', 'green'}. A is the colour to apply to the UPPER margin of
      the data in a given column. B is the colour to apply to the LOWER margin of the data
      in a given column.
    """
    pass


def colorize_all_sheets(workbook: xlsxwriter.Workbook, column_formatting: list[str],
                        margins: tuple[int], colour_options: tuple[str]):
    """
    Applies the given column_formatting, colour_options, and margins to the
    all sheets in the given workbook. Arguments are the same as for
    the function colorize_columns. Read about column_formatting, margins, and
    colour_options in by calling help(colorize_columns).
    """
    pass
