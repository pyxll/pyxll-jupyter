"""
Entry points for PyXLL integration.

These are picked up automatically when PyXLL starts to add Jupyter
functionality to Excel as long as this package is installed.

To install this package use::

    pip install pyxll_jupyter

"""
from .widgets import JupyterQtWidget, QApplication, QMessageBox
from .kernel import launch_jupyter
from pyxll import xlcAlert, get_config, xl_app, xl_macro, schedule_call
from functools import partial
import ctypes.wintypes
import pkg_resources
import logging
import sys
import os

_log = logging.getLogger(__name__)


def _get_qt_app():
    """Get or create the Qt application"""
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


def _get_notebook_path(cfg):
    """Return the path to open the Jupyter notebook in."""
    # Use the path of the active workbook if use_workbook_dir is set
    use_workbook_dir = False
    if cfg.has_option("JUPYTER", "use_workbook_dir"):
        try:
            use_workbook_dir = bool(int(cfg.get("JUPYTER", "use_workbook_dir")))
        except (ValueError, TypeError):
            _log.error("Unexpected value for JUPYTER.use_workbook_dir.")

    if use_workbook_dir:
        xl = xl_app(com_package="win32com")
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


def _get_jupyter_timeout(cfg):
    """Return the timeout in seconds to use when starting Jupyter."""
    timeout = 30.0
    if cfg.has_option("JUPYTER", "timeout"):
        try:
            timeout = float(cfg.get("JUPYTER", "timeout"))
            _log.debug("Using a timeout of %.1fs for starting the Jupyter notebook." % timeout)
        except (ValueError, TypeError):
            _log.error("Unexpected value for JUPYTER.timeout.")
    return max(timeout, 1.0)


def _get_notebook_kwargs(initial_path=None, notebook_path=None):
    """Get the kwargs for calling launch_jupyter.

    :param initial_path: Path to open Jupyter in.
    :param notebook_path: Path of Jupyter notebook to open.
    """
    if initial_path is not None and notebook_path is not None:
        raise RuntimeError("'initial_path' and 'notebook_path' cannot both be set.")

    if notebook_path is not None:
        if not os.path.exists(notebook_path):
            raise RuntimeError("Notebook path '%s' not found." % notebook_path)
        if not os.path.isfile(notebook_path):
            raise RuntimeError("Notebook path '%s' is not a file." % notebook_path)
        notebook_path = os.path.abspath(notebook_path)

    cfg = get_config()
    timeout = _get_jupyter_timeout(cfg)

    if notebook_path is None and initial_path is None:
        initial_path = _get_notebook_path(cfg)
    if initial_path and not os.path.exists(initial_path):
        raise RuntimeError("Directory '%s' does not exist.")
    if initial_path and not os.path.isdir(initial_path):
        raise RuntimeError("Path '%s' is not a directory.")

    return {
        "initial_path": initial_path,
        "notebook_path": notebook_path,
        "timeout": timeout
    }


def open_jupyter_notebook(*args, initial_path=None, notebook_path=None):
    """Ribbon action function for opening the Jupyter notebook
    browser control in a custom task pane.

    :param initial_path: Path to open Jupyter in.
    :param notebook_path: Path of Jupyter notebook to open.
    """
    from pyxll import create_ctp

    # Get the Qt Application
    app = _get_qt_app()

    # Create the Jupyter web browser widget
    kwargs = _get_notebook_kwargs(initial_path=initial_path, notebook_path=notebook_path)
    widget = JupyterQtWidget(**kwargs)

    # Show it in a CTP
    create_ctp(widget, width=800)


def open_jupyter_notebook_in_browser(*args, initial_path=None, notebook_path=None):
    """Ribbon action function for opening the Jupyter notebook in a web browser.

    :param initial_path: Path to open Jupyter in.
    :param notebook_path: Path of Jupyter notebook to open.
    """
    kwargs = _get_notebook_kwargs(initial_path=initial_path, notebook_path=notebook_path)
    launch_jupyter(no_browser=False, **kwargs)


def set_selection_in_ipython(*args):
    """Gets the value of the selected cell and copies it to
    the globals dict in the IPython kernel.
    """
    from pyxll import xl_app, XLCell

    try:
        if not getattr(sys, "_ipython_app", None) or not sys._ipython_kernel_running:
            raise Exception("IPython kernel not running")

        # Get the current selected range
        xl = xl_app(com_package="win32com")
        selection = xl.Selection
        if not selection:
            raise Exception("Nothing selected")

        # Check to see if it looks like a pandas DataFrame
        try_dataframe = False
        has_index = False
        if selection.Rows.Count > 1 and selection.Columns.Count > 1:
            try:
                import pandas as pd
            except ImportError:
                pd = None
                pass

            if pd is not None:
                # If the top left corner is empty assume the first column is an index.
                try_dataframe = True
                top_left = selection.Cells[1].Value
                if top_left is None:
                    has_index = True

        # Get an XLCell object from the range to make it easier to get the value
        cell = XLCell.from_range(selection)

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

        # set the value in the shell's locals
        sys._ipython_app.shell.user_ns["_"] = value
        print("\n\n>>> Selected value set as _")
    except:
        app = _get_qt_app()
        QMessageBox.warning(None, "Error", "Error setting selection in Excel")
        _log.error("Error setting selection in Excel", exc_info=True)


@xl_macro
def OpenJupyterNotebook(path=None):
    """
    Open a Jupyter notebook in a new task pane.

    :param path: Path to Jupyter notebook file or directory.
    :return: True on success
    """
    try:
        if path is not None:
            if not os.path.isabs(path):
                # Try and get the absolute path relative to the active workbook
                xl = xl_app(com_package="win32com")
                wb = xl.ActiveWorkbook
                if wb is not None and wb.FullName and os.path.exists(wb.FullName):
                    abs_path = os.path.join(os.path.dirname(wb.FullName), path)
                    if os.path.exists(abs_path):
                        path = abs_path
            if not os.path.exists(path):
                raise RuntimeError(f"Path '{path}' not found.")

        initial_path = None
        notebook_path = None
        if path is not None:
            if os.path.isdir(path):
                initial_path = path
            elif os.path.isfile(path):
                notebook_path = path
            else:
                raise RuntimeError(f"Something is wrong with {path}")

        # Use schedule_call to actually open the notebook since if this was called
        # from a Workbook.Open macro Excel may not yet be ready to open a CTP.
        schedule_call(partial(open_jupyter_notebook,
                                initial_path=initial_path,
                                notebook_path=notebook_path))

        return True
    except Exception as e:
        xlcAlert(f"Error opening Jupyter notebook: {e}")
        raise


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
    cfg = get_config()

    disable_ribbon = False
    if cfg.has_option("JUPYTER", "disable_ribbon"):
        try:
            disable_ribbon = bool(int(cfg.get("JUPYTER", "disable_ribbon")))
        except (ValueError, TypeError):
            _log.error("Unexpected value for JUPYTER.disable_ribbon.")

    if disable_ribbon:
        return []

    ribbon = pkg_resources.resource_string(__name__, "resources/ribbon.xml")
    return [
        (None, ribbon)
    ]
