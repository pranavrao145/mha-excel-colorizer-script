import pandas as pd


class ExcelModifier:
    """
    Template for an object that can modify an Excel workbook.

    Attributes:
    - workbook_writer: the XlsxWriter of the workbook you want to modify
    - sheets_to_modify: a dictionary which has key-value pairs of format {<sheet
      name>: <sheet data>}, where sheet_name is the name of a sheet you want to
      modify and sheet_data is the pandas DataFrame that contains the data in
      that sheet
    """

    def __init__(self, writer):
        self.workbook_writer = writer
        self.sheets_to_modify = {}

    def close(self):
        """
        Closes the writer associated with this ExcelModifier.
        """
        self.workbook_writer.close()

    def set_sheets_to_modify(self, sheet_names: list[str]):
        """
        Setter for self.sheets_to_modify. Sets the sheet names that the any
        modifying function in this class will use.

        Args:
        - sheet_names: A list of strings that contain the names of the sheets
          to modify.
        """
        self.sheets_to_modify = sheet_names

    def _calculate_bounds(self, series: pd.Series, upper_bound: float,
                          lower_bound: float) -> tuple[float, float]:
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
        """
        sorted_series = series.sort_values()

        # adjusting user inputs for the series.quantile function
        upper_bound_quantile = 1.0 - (upper_bound / 100.0)
        lower_bound_quantile = lower_bound / 100.0

        # calculating X and Y using the series.quantile function
        x = sorted_series.quantile(upper_bound_quantile)
        y = sorted_series.quantile(lower_bound_quantile)

        return (x, y)

    def colorize_columns(self, column_formatting: dict[str, str],
                         margin_options: tuple[float, float], colour_options:
                         tuple[str, str] = ('green', 'red'), majority_percentage: int = 10) -> None:
        """
        Applies the given column_formatting, colour_options, and margin_options to the
        each sheet in self.sheets_to_modify with self.workbook_writer.

        Args:
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
        - majority_percentage: an integer which signifies the percentage that a
          specific data point must take up in a column to be considered a
          "majority". If the smallest or largest element in a column takes
          up at least majority_percentage% of the column, then it will not
          be included in the bound.

        EXAMPLE USAGE:
        Assume this ExcelModifier has a workbook_writer that edits the file
        'workbook_1.xlsx' which contains a sheet named 'sheet_1'.

        If the following inputs were given to the function:
        - self.sheets_to_modify contained 'sheet_1' and the associated DataFrame
        - column_formatting = {'mean': 'colour_upper', 'missing': 'colour_lower'}
        - margin_options = (5, 10)
        - colour_options = ('green', 'red')
        - majority_percentage = 10

        Then the following would happen:
        - For the column in sheet_1 called 'mean', the upper 5 percent of the data
          would be coloured green.
            - If the largest element takes up at least majority_percentage% of
              the column, it will not be included in the colorization.
        - For the column in sheet_1 called 'missing', the lower 10 percent of the
          data would be coloured red.
            - If the smallest element takes up at least majority_percentage% of
              the column, it will not be included in the colorization.
        """
        # extracting margins
        upper_bound, lower_bound = margin_options

        # exctracting the colours
        upper_colour, lower_colour = colour_options

        # do it for each sheet specified
        for sheet_name, sheet_data in self.sheets_to_modify.items():
            # PART 1: calculate bounds for the needed columns

            # using the helper function to calculate bounds for each column
            bounds = {column_name: self._calculate_bounds(
                sheet_data[column_name], upper_bound, lower_bound) for column_name in column_formatting}

            # PART 2: applying conditional formatting based on the bounds and column
            # formatting

            # creating reusable formats
            # TODO: custom colours
            upper_colour_format = self.workbook_writer.book.add_format(
                {'bg_color': ('#FFC7CE' if upper_colour == 'red' else '#C6EFCE')})
            lower_colour_format = self.workbook_writer.book.add_format(
                {'bg_color': ('#FFC7CE' if lower_colour == 'red' else '#C6EFCE')})

            worksheet = self.workbook_writer.sheets[sheet_name]
            num_rows = sheet_data.shape[0]

            for column_name, operation in column_formatting.items():
                # which column are we in
                column_pos = sheet_data.columns.get_loc(column_name)
                current_column = sheet_data[column_name]

                # if the biggest and lowest data points are within this
                # percentage, the bound will not be included in the
                # colorization
                ignore_bound_percentage = majority_percentage / 100

                # find the largest item in the column
                largest = current_column.max()
                # this will contain whether or not the largest element appeas
                # in more than majority_percentage% of this column
                upper_bound_majority_exists = (
                    current_column == largest).mean() >= ignore_bound_percentage
                # find the smallest item in the column
                smallest = current_column.min()
                # this will contain whether or not the smallest element appears
                # in more than majority_percentage% of this column
                lower_bound_majority_exists = (
                    current_column == smallest).mean() >= ignore_bound_percentage

                # for each row in the column
                for row_pos in range(1, num_rows):
                    current_data = sheet_data.iat[row_pos - 1, column_pos]

                    if current_data != 'nan':
                        # if the upper bound must be coloured
                        if operation in {'colour_upper', 'colour_both'}:
                            if upper_bound_majority_exists:  # if this is true, we don't want to include the bound
                                if current_data > bounds[column_name][0]:
                                    # overwriting the cell with the original data and new formatting
                                    worksheet.write(row_pos, column_pos,
                                                    current_data, upper_colour_format)
                            else:
                                if current_data >= bounds[column_name][0]:
                                    # overwriting the cell with the original data and new formatting
                                    worksheet.write(row_pos, column_pos,
                                                    current_data, upper_colour_format)

                        if operation in {'colour_lower', 'colour_both'}:
                            if lower_bound_majority_exists:  # if this is true, we don't want to include the bound
                                if current_data < bounds[column_name][1]:
                                    # overwriting the cell with the original data and new formatting
                                    worksheet.write(row_pos, column_pos,
                                                    current_data, lower_colour_format)
                            else:
                                if current_data <= bounds[column_name][1]:
                                    # overwriting the cell with the original data and new formatting
                                    worksheet.write(row_pos, column_pos,
                                                    current_data, lower_colour_format)
