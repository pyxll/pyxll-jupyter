"""
Magic functions for use when running IPython inside Excel.
"""
from IPython.core.magic import Magics, magics_class, line_magic
from IPython.core.magic_arguments import argument, magic_arguments, parse_argstring
from pyxll import xl_app, plot, XLCell
import logging

_log = logging.getLogger(__name__)


@magics_class
class ExcelMagics(Magics):
    """Magic functions for interacting with Excel."""

    xlDown = -4121
    xlToRight = -4161

    @line_magic
    @magic_arguments()
    @argument("-c", "--cell", help="Address of cell to get value of.")
    @argument("-t", "--type", help="Datatype to convert the value to.")
    @argument("-f", "--formatter", help="PyXLL Formatter to use when setting the value.")
    @argument("-x", "--no-auto-resize", action="store_true", help="Don't auto-resize the range.")
    @argument("value", type=str, help="Value to set in Excel.")
    def xl_set(self, line):
        """Set a value to the current selection in Excel."""
        argv = self._split_args(line)
        args = self.xl_set.parser.parse_args(argv)
        value = eval(args.value, self.shell.user_ns, self.shell.user_ns)

        xl = xl_app(com_package="win32com")

        # Get the specified range, or use the current selection
        if args.cell:
            selection = xl.Range(args.cell.strip("\"' "))
        else:
            selection = xl.Selection
            if not selection:
                raise Exception("Nothing selected")

        # Get an XLCell object from the range to make it easier to set the value
        cell = XLCell.from_range(selection)

        if not args.no_auto_resize:
            cell = cell.options(auto_resize=True)

        if args.formatter:
            formatter = eval(args.formatter, self.shell.user_ns, self.shell.user_ns)
            cell = cell.options(formatter=formatter)

        if args.type:
            cell = cell.options(type=args.type)
        else:
            try:
                import pandas as pd
            except ImportError:
                pd = None

            if pd is not None and isinstance(value, pd.DataFrame):
                type_kwargs = {}
                if value.index.name:
                    type_kwargs["index"] = True
                cell = cell.options(type="dataframe", type_kwargs=type_kwargs)

        # Finally set the value in Excel
        cell.value = value

    @line_magic
    @magic_arguments()
    @argument("-c", "--cell", help="Address of cell to get value of.")
    @argument("-t", "--type", help="Datatype to convert the value to.")
    @argument("-x", "--no-auto-resize", action="store_true", help="Don't auto-resize the range.")
    def xl_get(self, line):
        """Get the current selection in Excel into Python."""
        argv = self._split_args(line)
        args = self.xl_get.parser.parse_args(argv)
        xl = xl_app(com_package="win32com")

        # Get the specified range, or use the current selection
        if args.cell:
            selection = xl.Range(args.cell.strip("\"' "))
        else:
            selection = xl.Selection
            if not selection:
                raise Exception("Nothing selected")

        # Expand the range if possible
        if not args.no_auto_resize:
            top_left = selection.Cells.Item(1)
            bottom_left = top_left.GetOffset(selection.Rows.Count-1, 0)
            bottom_right = top_left.GetOffset(selection.Rows.Count-1, selection.Columns.Count-1)

            # Check to see if there is a value below this range
            below = bottom_left.GetOffset(1, 0)
            if below.Value is not None:
                new_bottom_left = bottom_left.End(self.xlDown)

                # If the initial value is None then navigating down from it will only
                # go to the next non-empty cell, not right to be bottom of the range.
                # This happens if the top left cell is empty (above an index) and only
                # that single cell is selected.
                if selection.Count == 1 and bottom_left.Value is None:
                    new_below = new_bottom_left.GetOffset(1, 0)
                    if new_below.Value is not None:
                        new_bottom_left = new_bottom_left.End(self.xlDown)

                new_bottom_right = new_bottom_left.GetOffset(0, selection.Columns.Count-1)

                # Check to see if we can expand to the right as well
                right = new_bottom_right.GetOffset(0, 1)
                if right.Value is not None:
                    new_bottom_right = new_bottom_right.End(self.xlToRight)

                    if new_bottom_right.Row < bottom_right.Row:
                        new_bottom_right = new_bottom_left.GetOffset(0, selection.Columns.Count-1)

                selection = xl.Range(top_left, new_bottom_right)

        # Get an XLCell object from the range to make it easier to get the value
        cell = XLCell.from_range(selection)

        # If a type was passed use that
        if args.type:
            return cell.options(type=args.type).value

        # Otherwise check to see if it looks like a pandas DataFrame
        try_dataframe = False
        has_index = False
        if selection.Rows.Count > 1 and selection.Columns.Count > 1:
            try:
                import pandas as pd
            except ImportError:
                pd = None

            if pd is not None:
                # If the top left corner is empty assume the first column is an index.
                try_dataframe = True
                top_left = selection.Cells.Item(1).Value
                if top_left is None:
                    has_index = True

        # Get the value using PyXLL's dataframe converter, or as a plain value.
        value = None
        if try_dataframe:
            try:
                type_kwargs = {"index": 1 if has_index else 0}
                value = cell.options(type="dataframe", type_kwargs=type_kwargs).value
            except:
                _log.warning("Error converting selection to DataFrame", exc_info=True)

        if value is None:
            value = cell.value

        return value

    @line_magic
    @magic_arguments()
    @argument("-n", "--name", help="Name of the picture object in Excel to use.")
    @argument("-c", "--cell", help="Address of cell to use when creating the Picture in Excel.")
    @argument("-w", "--width", type=float, help="Width in points to use when creating the Picture in Excel.")
    @argument("-h", "--height", type=float, help="Height in points to use when creating the Picture in Excel.")
    @argument("figure", type=str, help="Figure to plot.")
    def xl_plot(self, line):
        """Plot a figure to Excel in the same way as pyxll.plot.

        The figure is exported as an image and inserted into Excel as a Picture object.

        If the --name argument is used and the picture already exists then it will not
        be resized or moved.
        """
        argv = self._split_args(line)
        args = self.xl_plot.parser.parse_args(argv)
        figure = eval(args.figure, self.shell.user_ns, self.shell.user_ns)

        kwargs = {}
        if args.cell is not None:
            xl = xl_app(com_package="win32com")
            cell = xl.Range(args.cell.strip("\"' "))
            kwargs["top"] = cell.Top
            kwargs["left"] = cell.Left

        plot(figure,
             name=args.name,
             width=args.width,
             height=args.height,
             **kwargs)

    @staticmethod
    def _split_args(line):
        """This is used instead of the standard arg_split to allow full Python
        expressions to be used as arguments.

        For example, %xl_plot df.plot(x="x", y="y") should be kept as a single
        argument and not split by spaces.
        """
        line = line.strip()
        if not line:
            return []

        open_close = {
            "(": ")",
            "{": "}",
            "[": "]",
            "<": ">",
            '"': '"',
            "'": "'"
        }

        separators = {
            " ",
            "\t"
        }

        args = [""]
        stack = []
        for char in line:
            # If we're in something already check if it's closed
            if stack:
                if char == open_close[stack[-1]]:
                    stack.pop(-1)
                    args[-1] += char
                    continue

            # Check if entering parentheses or a quoted string
            if char in open_close:
                stack.append(char)
                args[-1] += char
                continue

            # If we're not in the middle of something check for a separator
            if not stack and char in separators:
                if args[-1]:
                    args.append("")
                continue

            args[-1] += char
            continue

        return list(filter(None, args))
