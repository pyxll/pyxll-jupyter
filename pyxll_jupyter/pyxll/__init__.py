"""
Entry points for PyXLL integration.

These are picked up automatically when PyXLL starts to add Jupyter
functionality to Excel as long as this package is installed.

This module is kept intentionally light and all the actual functions
are implemented in the impl module. This is to minimize import time
to avoid slowing down Excel when opening.

To install this package use::

    pip install pyxll_jupyter

"""
from pyxll import get_config, xl_macro
import logging
import sys

_log = logging.getLogger(__name__)


if sys.version_info[:2] >= (3, 7):
    import importlib.resources

    def _resource_bytes(package, resource_name):
        return importlib.resources.read_binary(package, resource_name)

else:
    import pkg_resources

    def _resource_bytes(package, resource_name):
        return pkg_resources.resource_stream(package, resource_name).read()


def open_jupyter_notebook(*args, initial_path=None, notebook_path=None):
    """Ribbon action function for opening the Jupyter notebook
    browser control in a custom task pane.

    :param initial_path: Path to open Jupyter in.
    :param notebook_path: Path of Jupyter notebook to open.
    """
    from .impl import open_jupyter_notebook
    open_jupyter_notebook(*args, initial_path=initial_path, notebook_path=notebook_path)


def open_jupyter_notebook_in_browser(*args, initial_path=None, notebook_path=None):
    """Ribbon action function for opening the Jupyter notebook in a web browser.

    :param initial_path: Path to open Jupyter in.
    :param notebook_path: Path of Jupyter notebook to open.
    """
    from .impl import open_jupyter_notebook_in_browser
    open_jupyter_notebook_in_browser(*args, initial_path=initial_path, notebook_path=notebook_path)


def set_selection_in_ipython(*args):
    """Gets the value of the selected cell and copies it to
    the globals dict in the IPython kernel.
    """
    from .impl import set_selection_in_ipython
    set_selection_in_ipython(*args)


@xl_macro
def OpenJupyterNotebook(path=None, browser=False):
    """
    Open a Jupyter notebook in a new task pane.

    :param path: Path to Jupyter notebook file or directory.
    :param browser: Set to true to open in a browser instead of a task pane.
    :return: True on success
    """
    from .impl import OpenJupyterNotebook
    OpenJupyterNotebook(path=path, browser=browser)


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

    ribbon = _resource_bytes("pyxll_jupyter.resources", "ribbon.xml").decode("utf-8")
    return [
        (None, ribbon)
    ]
