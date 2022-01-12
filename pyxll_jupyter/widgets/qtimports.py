"""Common imports used by other modules in this package.

They are collected here so we can switch between PySide6, PyQt6, PySide2 and PyQt5
depending on what's installed.

A specific Qt package can be specified in the pyxll.cfg file by setting
'qt' in the JUPYTER section, eg::

    [JUPYTER]
    qt = PyQt5
"""
import logging
import pyxll

_log = logging.getLogger(__name__)


def _get_qt_packages():
    """Return a tuple of Qt packages to try importing."""
    qt_pkgs = "pyside6", "pyqt6", "pyside2", "pyqt5"

    cfg = pyxll.get_config()
    if cfg.has_option("JUPYTER", "qt"):
        qt_pkg = cfg.get("JUPYTER", "qt").strip().lower()
        _log.debug("pyxll_jupyter:Qt package specified as '%s'." % qt_pkg)
        if qt_pkg not in qt_pkgs:
            raise RuntimeError("Unsupported Qt package specified: '%s'" % qt_pkg)
        return qt_pkg,

    return qt_pkgs


_qt_packages = _get_qt_packages()

for pkg in _qt_packages:
    try:
        if pkg == "pyside6":
            # Requires PySide6
            from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QTabWidget, QTabBar, QMessageBox
            from PySide6.QtWebEngineCore import QWebEngineProfile, QWebEnginePage
            from PySide6.QtWebEngineWidgets import QWebEngineView
            from PySide6.QtGui import QKeySequence, QShortcut
            from PySide6.QtCore import QUrl, Qt, Signal, qVersion
            _log.debug("pyxll_jupyter:Using PySide6")
        elif pkg == "pyqt6":
            # Requires PyQt6 and PyQt6-WebEngine
            from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QTabWidget, QTabBar, QMessageBox
            from PyQt6.QtWebEngineCore import QWebEngineProfile, QWebEnginePage
            from PyQt6.QtWebEngineWidgets import QWebEngineView
            from PyQt6.QtGui import QKeySequence, QShortcut
            from PyQt6.QtCore import QUrl, Qt, qVersion
            from PyQt6.QtCore import pyqtSignal as Signal
            _log.debug("pyxll_jupyter:Using PyQt6")
        elif pkg == "pyside2":
            # Requires PySide2
            from PySide2.QtWidgets import QApplication, QWidget, QVBoxLayout, QTabWidget, QTabBar, QShortcut, QMessageBox
            from PySide2.QtWebEngineWidgets import QWebEngineView, QWebEngineProfile, QWebEnginePage
            from PySide2.QtGui import QKeySequence
            from PySide2.QtCore import QUrl, Qt, Signal, qVersion
            _log.debug("pyxll_jupyter:Using PySide2")
        elif pkg == "pyqt5":
            # Requires PyQt5 and PyQt5-WebEngine
            from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTabWidget, QTabBar, QShortcut, QMessageBox
            from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineProfile, QWebEnginePage
            from PyQt5.QtGui import QKeySequence
            from PyQt5.QtCore import QUrl, Qt, qVersion
            from PyQt5.QtCore import pyqtSignal as Signal
            _log.debug("pyxll_jupyter:Using PyQt5")
        else:
            raise RuntimeError("Unexpected Qt package '%s'" % pkg)
        break
    except ImportError:
        if pkg == _qt_packages[-1]:
            raise
        continue
else:
    raise RuntimeError("No suitable Qt package found.")
