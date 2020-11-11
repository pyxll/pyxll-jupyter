# PyXLL-Jupyter

Integration for Jupyter notebooks and Microsoft Excel.

Requires:

- PyXLL >= 5.0.0
- Jupyter
- PySide2

To install this package use:

    pip install pyxll-jupyter

Once installed a "Jupyter Notebook" button will be added to the PyXLL ribbon tab in Excel, so
long as you have PyXLL 5 or above already installed.

When using Jupyter in Excel the Python kernel runs inside the Excel process using PyXLL. You
can interact with Excel from code in the Jupyter notebook, and write Excel functions
using the PyXLL decorators @xl_menu and @xl_macro etc.

As the kernel runs in the existing Python interpreter in the Excel process it is not possible
to restart the kernel or to use other Python versions or other languages.

To configure add the following to your pyxll.cfg file (default values shown):

    [JUPYTER]
    use_workbook_dir = 0
    notebook_dir = Documents

If *use_workbook_dir* is set and the current workbook is saved then Jupyter will open in the same folder
as the current workbook.

The following magic functions are available in addition to the standard Jupyter magic functions:

```
%xl_get [-c CELL] [-t TYPE] [-x]

Get the current selection in Excel into Python.

optional arguments:
  -c CELL, --cell CELL  Address of cell to get value of.
  -t TYPE, --type TYPE  Datatype to convert the value to.
  -x, --no-auto-resize  Don't auto-resize the range.
```

```
%xl_set [-c CELL] [-t TYPE] [-f FORMATTER] [-x] value

Set a value to the current selection in Excel.

positional arguments:
  value                 Value to set in Excel.

optional arguments:
  -c CELL, --cell CELL  Address of cell to get value of.
  -t TYPE, --type TYPE  Datatype to convert the value to.
  -f FORMATTER, --formatter FORMATTER
                        PyXLL Formatter to use when setting the value.
  -x, --no-auto-resize  Don't auto-resize the range.
```

```
%xl_plot [-n NAME] [-c CELL] [-w WIDTH] [-h HEIGHT] figure

Plot a figure to Excel in the same way as pyxll.plot.

The figure is exported as an image and inserted into Excel as a Picture object.

If the --name argument is used and the picture already exists then it will not
be resized or moved.

positional arguments:
  figure                Figure to plot.

optional arguments:
  -n NAME, --name NAME  Name of the picture object in Excel to use.
  -c CELL, --cell CELL  Address of cell to use when creating the Picture in
                        Excel.
  -w WIDTH, --width WIDTH
                        Width in points to use when creating the Picture in
                        Excel.
  -h HEIGHT, --height HEIGHT
                        Height in points to use when creating the Picture in
                        Excel.
```

For more information about installing and using PyXLL see https://www.pyxll.com.

Copyright (c) 2020 PyXLL Ltd
