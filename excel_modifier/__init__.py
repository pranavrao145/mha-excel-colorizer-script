import pandas as pd
import xlsxwriter


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

    def _parse_instructions(self, columns: list[str] | dict[str, str], instructions: dict) -> dict[str, str]:
        """
        Given an instruction string of the format described in the docstring of
        self.colourize_columns, returned a dictionary of options to pass into
        self._colourize_columns.
        """
        instruction_split = instructions.split()
        POSSIBLE_INSTRUCTIONS = 'MmCcps'

        params_to_pass = {}

        # going through each element in the split instructions
        for i in range(len(instruction_split)):
            current_elem = instruction_split[i]
            # first, if columns is already a dict, pre-emptively get ready to pass it
            if isinstance(columns, dict):
                params_to_pass['column_formatting'] = columns

            # if this is indeed a new instruction
            if current_elem in POSSIBLE_INSTRUCTIONS:
                next_item = instruction_split[i + 1]
                match current_elem:
                    case 'M':
                        params_to_pass['margin_upper'] = float(next_item)
                    case 'm':
                        params_to_pass['margin_lower'] = float(next_item)
                    case 'C':
                        params_to_pass['colour_upper'] = 'red' if next_item == 'r' else 'green'
                    case 'c':
                        params_to_pass['colour_lower'] = 'red' if next_item == 'r' else 'green'
                    case 'p':
                        params_to_pass['majority_percentage'] = float(
                            next_item)
                    case 's':
                        if isinstance(columns, list):
                            OPTION_EXPANSIONS = {'u': 'colour_upper', 'l':
                                                 'colour_lower', 'b': 'colour_both'}
                            current_option = OPTION_EXPANSIONS[next_item]
                            params_to_pass['column_formatting'] = {column:
                                                                   current_option
                                                                   for column
                                                                   in columns}
        return params_to_pass

    def colourize_columns(self, columns: list[str] | dict[str, str], instructions: dict) -> None:
        """
        Given a string of instructions (specification below), colorize columns
        of all sheets in self.sheets_to_modify based on the instructions.

        Args:
        - columns: One of the following:
            - a list of strings which signify the columns to which the
              instructions passed to this function should be applied (for every
              sheet in self.sheets_to_modify). If this is a list, then the
              "s" argument from the instruction string will be used
            - a dictionary of the form {column_name: format_option} where
              column_name is a valid column name in the given sheet and
              format_option is an item from the set { 'u', 'l', 'b' }. 'u'
              stands for upper, meaning elements within the upper margin
              will be colorized, and similarly for lower and both.

              Note that, if this argument is a dict, then the "s" option from
              the instruction string will be ignored.
        - instructions: A string of instructions that specify how each column
        should be modified. The specification for this instructions string is
        found below:

        Specification for Instructions String (note: order in which these options are specified
        does not matter but all of them except 's' must be present in an instructions string):
        - M is used to specify the upper margin percentage, which is the percentage
        of the data in a column to consider as an UPPER margin. It is followed by
        a space and an float to 0.0 to 100.0 (e.g. M 35.0).
        - m is used to specify the upper margin percentage, which is the percentage
        of the data in a column to consider as an LOWER margin. It is followed by
        a space and an float from 0.0 to 100.0 (e.g. m 35.0).
        - C is used to specify the upper margin colour, which is the colour that
        any data within the upper margin will have. It is followed by a space
        and one of {'r', 'g'} (e.g. C g).
        - c is used to specify the lower margin colour, which is the colour that
        any data within the lower margin will have. It is followed by a space
        and one of {'r', 'g'} (e.g. c r).
        - p is used to specify the majority percentage, which is the percentage
        either the smallest or largest element in a column will have to take
        up to be a "majority", which will exclude it from being colorized. It
        is followed by a space and a float from 0.0 to 100.0 (e.g. p 10.0).
        - s is used to specify the sections of the columns that must be colorized.
        It is followed by a space and one of {'u', 'l', 'b'}. 'u' stands for upper,
        meaning elements within the upper margin will be colorized, and similarly
        for lower and both. Note that this argument will be ignored if the type
        of columns is a dictionary.

        EXAMPLE USAGE:
        Here are 2 example calls:

        1. colourize_columns(['mean', 'missing'], 'M 5.0 m 10.0 C g c r p 10.0 s b')

        This will colourize the 'mean' and 'missing' columns in every sheet in
        self.set_sheets_to_modify, considering the upper margin as 5% and colouring
        it green, the lower margin as 10% and colouring it red, considering the "majority
        percentage" as 10%, and colourizing both the upper and lower margins. Note that
        even if integers are used in place of the floats, this function still works.

        2. colourize_columns({'mean': 'u', 'missing': 'l'}, 'M 5.0 m 10.0 C g c r p 10.0')

        This does the same thing as 1, except the 'mean' column only has the upper
        margin colourized and the 'missing' column only has the lower margin
        colourized.

        For any further clarification on options, view the docstring of
        self._colourize_columns.
        """
        # parse the instruction string
        parsed_instructions = self._parse_instructions(columns, instructions)

        # extract the parameters from the parsed instructions
        column_formatting = parsed_instructions['column_formatting']
        margin_options = (
            parsed_instructions['margin_upper'], parsed_instructions['margin_lower'])
        colour_options = (
            parsed_instructions['colour_upper'], parsed_instructions['colour_lower'])
        majority_percentage = parsed_instructions['majority_percentage']

        # call the helper function
        self._colourize_columns(column_formatting, margin_options,
                                colour_options, majority_percentage)

    def _colourize_columns(self, column_formatting: dict[str, str],
                           margin_options: tuple[float, float], colour_options:
                           tuple[str, str] = ('green', 'red'), majority_percentage: float = 10.0) -> None:
        """
        Applies the given column_formatting, colour_options, and margin_options to the
        each sheet in self.sheets_to_modify with self.workbook_writer.

        Args:
        - column_formatting: a dictionary of the form {column_name: format_option} where
          column_name is a valid column name in the given sheet and format_option is an
          item from the set { 'colour_both', 'colour_upper', 'colour_lower' }
        - margin_options: a tuple of the form (X, Y) where X and Y are floats from
          0.0 to 100.0 (inclusive). X is the percent of the data to consider as the UPPER margin.
          Y is the percent of the data to consider as the LOWER margin.
        - colour_options: a tuple of the form (A, B) where A and B are strings
          from the set {'red', 'green'}. A is the colour to apply to the UPPER margin of
          the data in a given column. B is the colour to apply to the LOWER margin of the data
          in a given column.
        - majority_percentage: a float from 0.0 to 100.0 which signifies the
          percentage that a specific data point must take up in a column to be
          considered a "majority". If the smallest or largest element in a
          column takes up at least majority_percentage% of the column, then it
          will not be included in the bound.

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

                # if the biggest and lowest data points take up this much percentage
                # of the upper and lower bounds, respectively, they will not be
                # colorized
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
                            if upper_bound_majority_exists:  # if this is true, we don't want to include the largest element
                                if largest > current_data >= bounds[column_name][0]:
                                    # overwriting the cell with the original data and new formatting
                                    worksheet.write(row_pos, column_pos,
                                                    current_data, upper_colour_format)
                            else:
                                if largest >= current_data >= bounds[column_name][0]:
                                    # overwriting the cell with the original data and new formatting
                                    worksheet.write(row_pos, column_pos,
                                                    current_data, upper_colour_format)

                        if operation in {'colour_lower', 'colour_both'}:
                            if lower_bound_majority_exists:  # if this is true, we don't want to include the smallest element
                                if smallest < current_data <= bounds[column_name][1]:
                                    # overwriting the cell with the original data and new formatting
                                    worksheet.write(row_pos, column_pos,
                                                    current_data, lower_colour_format)
                            else:
                                if smallest <= current_data <= bounds[column_name][1]:
                                    # overwriting the cell with the original data and new formatting
                                    worksheet.write(row_pos, column_pos,
                                                    current_data, lower_colour_format)

    def autofit_sheets(self) -> None:
        """
        Autofits all sheets in self.set_sheets_to_modify.
        """
        for sheet_name in self.sheets_to_modify:
            worksheet = self.workbook_writer.sheets[sheet_name]
            worksheet.autofit()


if __name__ == "__main__":
    writer = pd.ExcelWriter('test.xlsx')

    df = pd.DataFrame({
        'test1': [3, 4, 5, 6, 7, 8, 9],
        'test2': [3, 4, 5, 6, 7, 8, 9],
        'test3': [3, 4, 5, 6, 7, 8, 9],
        'test4': [3, 4, 5, 6, 7, 8, 9],
        'test5': [3, 4, 5, 6, 7, 8, 9],
    })

    df.to_excel(writer, sheet_name='Sheet1',
                startrow=3, startcol=1, index=False)

    writer.close()
