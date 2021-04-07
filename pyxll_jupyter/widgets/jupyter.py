"""
JupyterQtWidget is the widget that gets embedded in Excel and hosts
a tabbed browser widget containing the Jupyter notebook.
"""
from ..kernel import start_kernel, launch_jupyter, kill_process
from .browser import Browser
from .qtimports import QWidget, QVBoxLayout
import subprocess
import logging
import ctypes
import os

_log = logging.getLogger(__name__)


class JupyterQtWidget(QWidget):

    def __init__(self, parent=None, scale=None, **kwargs):
        super().__init__(parent)

        # proc gets set to the subprocess when the jupyter is started
        self.proc = None

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
        self.proc, url = launch_jupyter(no_browser=True, **kwargs)
        self.browser.create_tab(url)

    def closeEvent(self, event):
        # Kill the Jupyter subprocess
        if self.proc is not None:
            kill_process(self.proc)
