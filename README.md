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

For more information about installing and using PyXLL see https://www.pyxll.com.

Copyright (c) 2020 PyXLL Ltd
