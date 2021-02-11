"""
JupyterQtWidget is the widget that gets embedded in Excel and hosts
a tabbed browser widget containing the Jupyter notebook.
"""
from .kernel import start_kernel, launch_jupyter
from .browser import Browser
from .qtimports import QWidget, QVBoxLayout
import subprocess
import logging
import ctypes
import os

_log = logging.getLogger(__name__)


class JupyterQtWidget(QWidget):

    def __init__(self, parent=None, scale=None, initial_path=None, timeout=30, notebook_path=None):
        super().__init__(parent)

        # proc gets set to the subprocess when the jupyter is started
        self.proc = None

        notebook = None
        if notebook_path is not None:
            if initial_path is not None:
                raise RuntimeError("'notebook_path' and 'initial_path' cannot be set together.")
            initial_path = os.path.dirname(notebook_path)
            notebook = os.path.basename(notebook_path)

        # Get the scale from the window DPI
        if scale is None:
            LOGPIXELSX = 88
            hwnd = self.winId()
            if isinstance(hwnd, str):
                hwnd = int(hwnd, 16 if hwnd.startswith("0x") else 10)
            hwnd = ctypes.c_size_t(hwnd)
            screen = ctypes.windll.user32.GetDC(hwnd)
            try:
                scale = ctypes.windll.gdi32.GetDeviceCaps(screen, LOGPIXELSX) / 96.0
            finally:
                ctypes.windll.user32.ReleaseDC(hwnd, screen)

        # Create the browser widget
        self.browser = Browser(self, scale=scale)
        self.browser.closed.connect(self.close)

        # Add the browser to the widgets layout
        layout = QVBoxLayout()
        layout.addWidget(self.browser)
        self.setLayout(layout)

        # Start the kernel and open Jupyter in a new tab
        app = start_kernel()
        connection_file = os.path.abspath(app.abs_connection_file)
        _log.debug(f"Kernel started with connection file '{connection_file}'")

        self.proc, url = launch_jupyter(connection_file,
                                        cwd=initial_path,
                                        timeout=timeout)

        # Update the URL to point to the notebook
        if notebook is not None:
            root, params = url.split("?", 1)
            url = root.rstrip("/") + "/notebooks/" + notebook + "?" + params

        self.browser.create_tab(url)

    def closeEvent(self, event):
        # Kill the Jupyter subprocess using taskkill (just killing the process using POpen.kill
        # doesn't terminate any child processes)
        if self.proc is not None:
            while self.proc.poll() is None:
                si = subprocess.STARTUPINFO()
                si.wShowWindow = subprocess.SW_HIDE
                subprocess.check_call(['taskkill', '/F', '/T', '/PID', str(self.proc.pid)],
                                      startupinfo=si,
                                      shell=True)
