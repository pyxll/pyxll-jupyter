"""
Magic functions for use when running IPython inside Excel.
"""
from IPython.core.magic import Magics, magics_class, line_magic
from IPython.core.magic_arguments import argument, magic_arguments, parse_argstring
from pyxll import xl_app, XLCell
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
        args = parse_argstring(self.xl_set, line)
        value = eval(args.value, self.shell.user_ns, self.shell.user_ns)

        xl = xl_app(com_package="win32com")

        # Get the specified range, or use the current selection
        if args.cell:
            selection = xl.Range(args.cell)
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
        args = parse_argstring(self.xl_get, line)
        xl = xl_app(com_package="win32com")

        # Get the specified range, or use the current selection
        if args.cell:
            selection = xl.Range(args.cell)
        else:
            selection = xl.Selection
            if not selection:
                raise Exception("Nothing selected")

        # Expand the range if possible
        if not args.no_auto_resize:
            top_left = selection.Cells[1]
            bottom_left = top_left.GetOffset(selection.Rows.Count-1, 0)
            bottom_right = top_left.GetOffset(selection.Rows.Count-1, selection.Columns.Count-1)

            # Check to see if there is a value below this range
            below = bottom_left.GetOffset(1, 0)
            if below.Value is not None:
                new_bottom_left = bottom_left.End(self.xlDown)
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
                top_left = selection.Cells[1].Value
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
