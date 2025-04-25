"""
JupyterQtWidget is the widget that gets embedded in Excel and hosts
a tabbed browser widget containing the Jupyter notebook.
"""
from ..kernel import launch_jupyter, release_kernel, pause_kernel, resume_kernel
from .browser import Browser
from .qtimports import Qt, QApplication, QEvent, QWidget, QVBoxLayout, qVersion
import logging
import ctypes

_log = logging.getLogger(__name__)


class JupyterQtWidget(QWidget):

    def __init__(self,
                 parent=None,
                 scale=None,
                 private_browser=False,
                 allow_cookies=True,
                 cache_path=None,
                 storage_path=None,
                 pause_on_focus_lost=True,
                 **kwargs):
        super().__init__(parent)
        self.__closed = False
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus)

        # proc gets set to the subprocess when the jupyter is started
        self.proc = None
        self.token = None

        # Get the scale from the window DPI if using Qt5
        if scale is None:
            qt_version = int(qVersion().split(".", 1)[0])
            if qt_version < 6:
                LOGPIXELSX = 88
                hwnd = self.winId()
                if isinstance(hwnd, str):
                    hwnd = int(hwnd, 16 if hwnd.startswith("0x") else 10)
                hwnd = ctypes.c_size_t(int(hwnd))
                screen = ctypes.windll.user32.GetDC(hwnd)
                try:
                    scale = ctypes.windll.gdi32.GetDeviceCaps(screen, LOGPIXELSX) / 96.0
                finally:
                    ctypes.windll.user32.ReleaseDC(hwnd, screen)

        self.__paused = False
        self.__pause_on_focus_lost = pause_on_focus_lost
        app = QApplication.instance()
        app.installEventFilter(self)

        # Create the browser widget
        self.browser = Browser(self,
                               scale=scale,
                               private_browser=private_browser,
                               allow_cookies=allow_cookies,
                               cache_path=cache_path,
                               storage_path=storage_path)

        self.browser.closed.connect(self.close)

        # Add the browser to the widgets layout
        layout = QVBoxLayout()
        layout.addWidget(self.browser)
        self.setLayout(layout)

        # Start the kernel and open Jupyter in a new tab
        self.token, url = launch_jupyter(no_browser=True, **kwargs)

        # Pause the kernel until we get the focus
        if pause_on_focus_lost and not self.hasFocus():
            self.pauseKernel()

        self.browser.create_tab(url)

    def closeEvent(self, event):
        self.__closed = True

        # Pause the kernel and kill the Jupyter subprocess
        release_kernel(self.token)

    def pauseKernel(self):
        """Maybe pause the kernel, if nothing else requires it."""
        if not self.__paused:
            _log.debug(f"Pausing kernel session {self.token}")
            pause_kernel(self.token)
            self.__paused = True

    def resumeKernel(self):
        """Resume the kernel, if not already running."""
        if self.__paused:
            _log.debug(f"Resuming kernel session {self.token}")
            resume_kernel(self.token)
            self.__paused = False

    def eventFilter(self, source, event):
        # The source can be a QWindow wrapper of the native CTP window, but
        # the underlying HWND will be the same.
        if (
            not self.__closed
            and self.__pause_on_focus_lost
            and source.isWindowType()
            # winId and effectiveWinId are of type sip.voidptr which are not comparible in PyQt
            and int(source.winId()) == int(self.effectiveWinId())
        ):
            if event.type() == QEvent.FocusIn:
                self.resumeKernel()

            elif event.type() == QEvent.FocusOut:
                self.pauseKernel()

        return super().eventFilter(source, event)
