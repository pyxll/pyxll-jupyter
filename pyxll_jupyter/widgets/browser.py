"""Minimal browser implementation for running Jupyter.

Requires:

pip install PyQt5
pip install PyQtWebEngine

or

pip install PySide2
"""
from .qtimports import *
import logging
import warnings
_log = logging.getLogger(__name__)


class Browser(QWidget):
    """Tabbed browser widget."""

    closed = Signal()

    def __init__(self,
                 parent=None,
                 scale=None,
                 private_browser=False,
                 allow_cookies=True,
                 cache_path=None,
                 storage_path=None):
        super().__init__(parent)
        self.tabs = []

        storage_name = "pyxll-jupyter"
        if private_browser:
            _log.debug("Using private browser mode.")
            storage_name = None

        self.profile = QWebEngineProfile(storage_name)

        if allow_cookies:
            _log.debug(f"Setting browser cookies policy to 'AllowPersistentCookies'.")
            self.profile.setPersistentCookiesPolicy(QWebEngineProfile.PersistentCookiesPolicy.AllowPersistentCookies)

        if cache_path is not None:
            self.profile.setCachePath(cache_path)

        if storage_path is not None:
            self.profile.setPersistentStoragePath(storage_path)

        _log.debug(f"Browser cache path is {self.profile.cachePath()}.")
        _log.debug(f"Browser persistent storage path is {self.profile.persistentStoragePath()}.")

        layout = QVBoxLayout()
        layout.setSpacing(0)
        layout.setContentsMargins(0, 0, 0, 0)

        self.tab_widget = TabWidget(self, self.profile, scale=scale)
        self.tab_widget.closed.connect(self.close)

        layout.addWidget(self.tab_widget)
        self.setLayout(layout)

    def create_tab(self, url):
        return self.tab_widget.create_tab(url)

    def closeEvent(self, event):
        self.closed.emit()


class TabWidget(QTabWidget):
    """Tabbed widget containing multiple QWebEngineView
    as a basic tabbed browser.

    :param parent: Parent widget.
    :param profile: QWebEngineProfile to use for all tabs.
    :param scale: Scale for scaling the content for HDI displays.
    """

    closed = Signal()

    def __init__(self, parent, profile, scale=None):
        super().__init__(parent)
        self.profile = profile
        self.scale = scale
        self.tabs = []

        tab_bar = self.tabBar()
        tab_bar.setTabsClosable(True)
        tab_bar.setSelectionBehaviorOnRemove(QTabBar.SelectPreviousTab)
        tab_bar.setMovable(True)
        tab_bar.setContextMenuPolicy(Qt.CustomContextMenu)
        tab_bar.tabCloseRequested.connect(self.close_tab)

        self.setDocumentMode(True)
        self.setElideMode(Qt.ElideRight)

        # Set up some shortcuts
        next_tab_sc = QShortcut(QKeySequence.fromString("Ctrl+PgUp"), self)
        next_tab_sc.activated.connect(self.next_tab)

        prev_tab_sc = QShortcut(QKeySequence.fromString("Ctrl+PgDown"), self)
        prev_tab_sc.activated.connect(self.prev_tab)

    def next_tab(self):
        i = self.currentIndex()
        self.setCurrentIndex((i + 1) % self.count())

    def prev_tab(self):
        i = self.currentIndex()
        self.setCurrentIndex((i - 1) % self.count())

    def close_tab(self, index):
        view = self.widget(index)
        if view:
            self.tabs.remove(view)
            self.removeTab(index)
            view.deleteLater()

        if len(self.tabs) == 0:
            self.close()

    def closeEvent(self, event):
        self.closed.emit()

    def create_tab(self, url=None):
        view = WebView(self)
        self.tabs.append(view)
        page = WebEnginePage(self.profile, view)
        view.setPage(page)

        if self.scale:
            view.setZoomFactor(self.scale)

        self.addTab(view, "Loading...")
        view.resize(self.currentWidget().size())
        self.__setup_view(view)
        self.setCurrentWidget(view)
        view.show()

        if url is not None:
            page.setUrl(QUrl(url))

        def url_changed(url):
            _log.debug("Loading page '%s'" % url)

        page.urlChanged.connect(url_changed)

        return view

    def __setup_view(self, view):
        page = view.page()

        def title_changed(title):
            index = self.indexOf(view)
            if index != -1:
                self.setTabText(index, title)
                self.setTabToolTip(index, title)

        view.titleChanged.connect(title_changed)

        def url_changed(url):
            index = self.indexOf(view)
            if index != -1:
                self.tabBar().setTabData(index, url)

        view.urlChanged.connect(url_changed)

        def icon_changed(icon):
            index = self.indexOf(view)
            if index != -1:
                self.setTabIcon(index, icon)

        view.iconChanged.connect(icon_changed)

        def close_requested():
            index = self.indexOf(view)
            if index != -1:
                self.closeTab(index)

        page.windowCloseRequested.connect(close_requested)



class WebEnginePage(QWebEnginePage):

    def javaScriptConsoleMessage(self, level, message, lineNumber, sourceID):
        """Log JavaScript errors to the Python log."""
        pylevel = logging.ERROR
        try:
            if level == QWebEnginePage.JavaScriptConsoleMessageLevel.InfoMessageLevel:
                pylevel = logging.INFO
            elif level == QWebEnginePage.JavaScriptConsoleMessageLevel.WarningMessageLevel:
                pylevel = logging.WARNING
        except AttributeError:
            warnings.warn("QWebEnginePage.JavaScriptConsoleMessageLevel attribute not found")

        if pylevel >= logging.WARNING:
            message = f"JavaScript Error (line {lineNumber}, source {sourceID}): {message}"
        _log.log(pylevel, message)


class WebView(QWebEngineView):
    """QWebEngineView that implements 'createWindow' so that
    all new windows are opened in a new tab.

    :param owner: QTabWidget containing this view.
    """

    def __init__(self, owner):
        super().__init__()
        self.__owner = owner

    def createWindow(self, type):
        if type == QWebEnginePage.WebBrowserTab \
        or type == QWebEnginePage.WebBrowserBackgroundTab \
        or type == QWebEnginePage.WebBrowserBackgroundTab:
            return self.__owner.create_tab()
