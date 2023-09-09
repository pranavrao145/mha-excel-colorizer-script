import pandas as pd
import xlsxwriter

RANGE = 'A:Z'  # this is an arbitrary range over which this script will format cells, and can be changed


def calculate_bounds(series: pd.Series, upper_bound: float, lower_bound: float,
                     include_zero: bool = False) -> tuple[float, float]:
    """
    Returns a tuple of the form (X, Y) where  X is the value at the upper_bound
    percentile and Y is the value of the lower_bound percentile of the data.
    In other words, X is the number such that all numbers in the column > X are
    in the top upper_bound percentage of data, and Y is the number such that
    all numbers in the column < Y are in the bottom lower_bound percentage of
    data. In other words,

    Args:
    - series: the pandas Series for which to calculate X and Y
    - upper_bound: the desired upper bound. For example, if this is 5.0, X
      will be the smallest number in the column such that the next largest
      number after X is part of the top 5% of data in the column.
    - lower_bound: the desired lower bound. For example, if this is 5.0, Y
      will be the largest number in the column such that the next largest
      number after Y is part of the bottom 5% of data in the column.
    - lower_bound: the desired lower bound. For example, if this is 5.0, Y
      will be the largest number in the column such that the next largest
      number after Y is part of the bottom 5% of data in the column.
    - include_zero_in_bounds: whether or not to include 0 when calculating the
      upper or lower X% of a column. Default is False (does not include 0)
    """
    # if we don't want zeros, we have to filter them out
    filtered_series = series if include_zero else series[series != 0]
    sorted_series = filtered_series.sort_values()

    # adjusting user inputs for the series.quantile function
    upper_bound_quantile = 1.0 - (upper_bound / 100.0)
    lower_bound_quantile = 1.0 - (lower_bound / 100.0)

    # calculating X and Y using the series.quantile function
    x = sorted_series.quantile(upper_bound_quantile)
    y = sorted_series.quantile(lower_bound_quantile)

    return (x, y)


pass


def colorize_columns(workbook: xlsxwriter.Workbook, workbook_name: str, sheet_name: str, column_formatting: dict[str, str],
                     margin_options: tuple[float], colour_options: tuple[str], include_zero_in_bounds: bool = False) -> None:
    """
    Applies the given column_formatting, colour_options, and margin_options to the
    sheet with sheet_name in the given workbook.

    EXAMPLE USAGE:
    Assume there is a workbook called 'workbook_1' which contains a sheet named 'sheet_1'.

    If the following inputs were given to the function:
    - workbook_name = 'workbook_1'
    - sheet_name = 'sheet_1'
    - column_formatting = {'mean': 'colour_upper', 'missing': 'colour_lower'}
    - margin_options = (5, 10)
    - colour_options = ('green', 'red')
    - include_zero_in_bounds = False

    Then the following would happen:
    - For the column in sheet_1 called 'mean', the upper 5 percent of the data
      would be coloured green. Note that 0 was NOT taken into account when
      calculating upper bound.
    - For the column in sheet_1 called 'missing', the lower 10 percent of the
      data would be coloured red. Note that 0 was NOT taken into account when
      calculating lower bound.

    Args:
    - workbook: the xlsxwriter workbook to edit
    - workbook_name: the name of the workbook to edit
    - sheet_name: the name of the sheet to search for the columns in
    - column_formatting: a dictionary of the form {column_name: format_option} where
      column_name is a valid column name in the given sheet and format_option is an
      item from the set { 'colour_both', 'colour_upper', 'colour_lower', 'colour_none'}
    - margin_options: a tuple of the form (X, Y) where X and Y are floats from 0.0 to 100.0 (inclusive).
      X is the percent of the data to consider as the UPPER margin. Y is the percent of the data
      to consider as the LOWER margin.
    - colour_options: a tuple of the form (A, B) where A and B are strings
      from the set {'red', 'green'}. A is the colour to apply to the UPPER margin of
      the data in a given column. B is the colour to apply to the LOWER margin of the data
      in a given column.
    - include_zero_in_bounds: whether or not to include 0 when calculating the upper or lower
      X% of a column. Default is False (does not include 0)
    """
    # PART 1: calculate bounds for the needed columns

    # reading the data and putting it in a dataframe so we can analyze it
    data_df = pd.read_excel(workbook_name, sheet_name=sheet_name)

    # extracting margins
    upper_bound, lower_bound = margin_options

    # using the helper function to calculate bounds for each column
    bounds = {column_name: calculate_bounds(
        data_df[column_name]) for column_name in column_formatting}

    # PART 2: applying conditional formatting based on the bounds and column
    # formatting

    # exctracting the colours
    upper_colour, lower_colour = colour_options

    # creating reusable formats
    upper_colour_format = workbook.add_format({'bg_color': upper_colour})
    lower_colour_format = workbook.add_format({'bg_color': lower_colour})

    worksheet = workbook.sheets[sheet_name]

    # TODO: refactor, and should it be greater than EQUAL to the upper bound or strictly greater than? Same for lower.
    for column_name, operation in column_formatting:
        if operation == "colour_upper":
            # format the upper X%
            worksheet.conditional_format(RANGE, {
                                         'type': 'cell', 'criteria': '>=', 'value': bounds[column_name][0], 'format': upper_colour_format})
        elif operation == "colour_lower":
            # format the lower Y%
            worksheet.conditional_format(RANGE, {
                                         'type': 'cell', 'criteria': '<=', 'value': bounds[column_name][1], 'format': lower_colour_format})
        elif operation == "colour_both":
            # format the upper X%
            worksheet.conditional_format(RANGE, {
                                         'type': 'cell', 'criteria': '>=', 'value': bounds[column_name][0], 'format': upper_colour_format})

            # format the lower Y%
            worksheet.conditional_format(RANGE, {
                                         'type': 'cell', 'criteria': '<=', 'value': bounds[column_name][1], 'format': lower_colour_format})
