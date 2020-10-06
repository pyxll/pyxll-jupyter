"""
Entry points for PyXLL integration.

These are picked up automatically when PyXLL starts to add Jupyter
functionality to Excel as long as this package is installed.

To install this package use::

    pip install pyxll_jupyter

"""
from .widget import JupyterQtWidget
from .qtimports import QApplication
from pyxll import get_config, xl_app
import ctypes.wintypes
import pkg_resources
import win32api
import logging
import sys
import os

_log = logging.getLogger(__name__)


def get_notebook_path():
    """Return the path to open the Jupyter notebook in."""
    cfg = get_config()

    # Use the path of the active workbook if use_workbook_dir is set
    use_workbook_dir = False
    if cfg.has_option("JUPYTER", "use_workbook_dir"):
        try:
            use_workbook_dir = bool(int(cfg.get("JUPYTER", "use_workbook_dir")))
        except (ValueError, TypeError):
            _log.error("Unexpected value for JUPYTER.use_workbook_dir.")

    if use_workbook_dir:
        xl = xl_app()
        wb = xl.ActiveWorkbook
        if wb is not None and wb.FullName and os.path.exists(wb.FullName):
            return os.path.dirname(wb.FullName)

    # Otherwise use the path option
    if cfg.has_option("JUPYTER", "notebook_dir"):
        path = cfg.get("JUPYTER", "notebook_dir").strip("\"\' ")
        if os.path.exists(path):
            return os.path.normpath(path)
        _log.warning("Notebook path '%s' does not exist" % path)

    # And if that's not set use My Documents
    CSIDL_PERSONAL = 5  # My Documents
    SHGFP_TYPE_CURRENT = 0  # Get current, not default value

    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
    return buf.value


def open_jupyter_notebook(*args):
    """Ribbon action function for opening the Jupyter notebook
    browser control.
    """
    from pyxll import create_ctp

    # Create the Qt Application
    app = QApplication.instance()
    if app is None:
        app = QApplication([])

    # The create the widget and add it as an Excel CTP
    path = get_notebook_path()
    print("OPENING JUPTER IN %s" % path)
    widget = JupyterQtWidget(initial_path=path)

    create_ctp(widget, width=800)


def set_selection_in_ipython(*args):
    """Gets the value of the selected cell and copies it to
    the globals dict in the IPython kernel.
    """
    from pyxll import xl_app, XLCell

    try:
        if not getattr(sys, "_ipython_app", None) or not sys._ipython_kernel_running:
            raise Exception("IPython kernel not running")

        xl = xl_app(com_package="win32com")
        selection = xl.Selection
        if not selection:
            raise Exception("Nothing selected")

        cell = XLCell.from_range(selection)
        value = cell.value

        # set the value in the shell's locals
        sys._ipython_app.shell.user_ns["_"] = value
        print("\n\n>>> Selected value set as _")
    except:
        if win32api:
            win32api.MessageBox(None, "Error setting selection in Excel")
        _log.error("Error setting selection in Excel", exc_info=True)


def modules():
    """Entry point for getting the pyxll modules.
    Returns a list of module names."""
    return [
        __name__
    ]


def ribbon():
    """Entry point for getting the pyxll ribbon file.
    Returns a list of (filename, data) tuples.
    """
    ribbon = pkg_resources.resource_string(__name__, "resources/ribbon.xml")
    return [
        (None, ribbon)
    ]
