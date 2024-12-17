# PyXLL-Jupyter

Integration for Jupyter notebooks and Microsoft Excel.

See the [Python Jupyter Notebooks in Excel](https://www.pyxll.com/blog/python-jupyter-notebooks-in-excel/) blog post for more details.

## Requirements

- PyXLL >= 5.1.0
- Jupyter >= 1.0.0
- notebook >= 6.0.0
- PySide2, or PySide6 for Python >= 3.10
  
### Optional

- jupyterlab >= 4.0.0

## Installation

To install this package use:

    pip install pyxll-jupyter

Once installed a "Jupyter Notebook" button will be added to the PyXLL ribbon tab in Excel, so
long as you have PyXLL 5 or above already installed.

When using Jupyter in Excel the Python kernel runs inside the Excel process using PyXLL. You
can interact with Excel from code in the Jupyter notebook, and write Excel functions
using the PyXLL decorators @xl_menu and @xl_macro etc.

As the kernel runs in the existing Python interpreter in the Excel process it is not possible
to restart the kernel or to use other Python versions or other languages.

## Configuration

To configure, add any of the following settings to your pyxll.cfg file. You do not need
to set all of these, only the ones you wish to change::

    [JUPYTER]
    ; Workbook settings
    use_workbook_dir = 0
    notebook_dir = C:\Path\To\Your\Documents
    subcommand = notebook

    ; Browser settings
    qt =
    allow_cookies = 1
    private_browser = 0
    cache_path =
    storage_path =

    ; Other settings
    timeout = 60
    disable_ribbon = 0

If *use_workbook_dir* is set and the current workbook is saved then Jupyter will open in the same folder
as the current workbook.

*notebook_dir* can be set to an existing folder that will be used as the root documents folder the Jupyter
opens in.

The *subcommand* option can be used to switch the Jupyter subcommand used to launch the Jupyter web server.
It can be set to either `notebook` for the default Jupyter notebook interface, or `lab` if using Jupyterlab
*(experimental)*.

*qt* can be used to switch which Qt implementation is used. Possible values are 'PySide6', 'PyQt6', 'PySide2',
and 'PyQt5'. By default, whichever Qt implementation is installed will be used.

*allow_cookies* will prevent the Qt browser from saving cookies if set to 0.

*private_browser* will prevent the Qt browser from using any previously stored data or saving any data from
the browser session.

*cache_path* can be set to an existing folder for the browser to save cached data. By default this will be
the Qt browser's default cached data path.

*storage_path* can be set to an existing folder for the browser to save persistent storage data. By default this will be
the Qt browser's default persistent storage path.

*timeout* is the maximum number of seconds of inactivity to wait for when starting the Jupyter server process. If you
are getting timeout errors then increasing this may help.

If *disable_ribbon* is set then the ribbon button to start Jupyter will not be shown, however Jupyter
may still be opened using the "OpenJupyterNotebook" macro.

## Experimental JupyterLab Support

Jupyterlab can be used instead of the default Jupyter Notebook interface by specifying
`subcommand = lab` in the ``[JUPYTER]`` section of the pyxll.cfg file.

This requires Jupyterlab >= 4.0.0 to be installed. At the time of writing, version 4 of Jupyterlab is in
pre-release and can be installed using:

    pip install --pre jupyterlab

### Qt

The pyxll-jupyter package uses the Qt [QWebEngineView](https://doc.qt.io/qt-5/qwebengineview.html) widget, and by
default will use the [PySide2](https://pypi.org/project/PySide2/) package for Python <= 3.9 or
the [PySide6](https://pypi.org/project/PySide6/) package for Python >= 3.10.

This can be changed to use [PyQt5](https://www.riverbankcomputing.com/software/pyqt/) by setting `qt = PyQt5` in
the `JUPYTER` section of the config. You will need to have both the `pyqt5` and `pyqtwebengine` packages installed
if using this option. Both can be installed using pip as follows:

    pip install pyqt5 pyqtwebengine

## Magic Functions

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

## Opening from VBA

You can open the Jupyter notebook from VBA using the ``OpenJupyterNotebook`` macro, called
via VBA's ``Run`` method. For example::

  Run "OpenJupyterNotebook"

The macro takes two arguments, the initial path and a boolean to open in a browser rather
than in Excel task pane if True.

The initial path can either be a valid path, or an empty string.

For example, to open Jupyter in a browser with the default path you would run the macro as follows::

    Run "OpenJupyterNotebook", "", True


For more information about installing and using PyXLL see https://www.pyxll.com.

Copyright (c) PyXLL Ltd
