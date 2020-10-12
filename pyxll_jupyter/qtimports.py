"""Common imports used by other modules in this package.

They are collected here so we can switch between PySide2 and PyQt5
depending on what's installed.
"""
try:
    from PySide2.QtWidgets import QApplication, QWidget, QVBoxLayout, QTabWidget, QTabBar, QShortcut
    from PySide2.QtWebEngineWidgets import QWebEngineView, QWebEngineProfile, QWebEnginePage
    from PySide2.QtGui import QKeySequence
    from PySide2.QtCore import QUrl, Qt, Signal
except ImportError:
    from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTabWidget, QTabBar, QShortcut
    from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineProfile, QWebEnginePage
    from PyQt5.QtGui import QKeySequence
    from PyQt5.QtCore import QUrl, Qt
    from PyQt5.QtCore import pyqtSignal as Signal
