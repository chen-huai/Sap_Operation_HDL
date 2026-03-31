from PyQt5.QtWidgets import QApplication
from qt_material import apply_stylesheet, list_themes
import random

class ThemeManager:
    def __init__(self, app):
        self.app = app
        self.current_theme = "light_blue.xml"
        self.themes = list_themes()

    def set_theme(self, theme):
        if theme in self.themes:
            self.current_theme = theme
            apply_stylesheet(self.app, theme=theme)
            self._adjust_button_style()
        else:
            print(f"Theme {theme} not found. Using default theme.")
            self.set_default_theme()

    def set_default_theme(self):
        self.current_theme = "light_blue.xml"
        apply_stylesheet(self.app, theme="light_blue.xml")
        self._adjust_button_style()

    def toggle_theme(self):
        current_index = self.themes.index(self.current_theme)
        next_index = (current_index + 1) % len(self.themes)
        self.set_theme(self.themes[next_index])

    def get_available_themes(self):
        return self.themes

    def set_random_theme(self):
        """随机设置一个主题"""
        random_theme = random.choice(self.themes)
        self.set_theme(random_theme)

    def _adjust_button_style(self):
        if "light" in self.current_theme:
            self.app.setStyleSheet(self.app.styleSheet() + """
                QPushButton:disabled {
                    color: #808080;
                }
                QPushButton[enabled="false"] {
                    color: #808080;
                }
                QDoubleSpinBox {
                    color: black;
                }
                QSpinBox {
                    color: black;
                }
                QLineEdit {
                    color: black;
                }
                QComboBox {
                    color: black;
                }
                QComboBox QAbstractItemView {
                    color: black;
                }
                QTextBrowser {
                    color: black;
                }
            """)
        else:
            self.app.setStyleSheet(self.app.styleSheet() + """
                QPushButton:disabled {
                    color: #808080;
                }
                QPushButton[enabled="false"] {
                    color: #808080;
                }
                QDoubleSpinBox {
                    color: white;
                }
                QSpinBox {
                    color: white;
                }
                QLineEdit {
                    color: white;
                }
                QComboBox {
                    color: white;
                }
                QComboBox QAbstractItemView {
                    color: white;
                }
                QTextBrowser {
                    color: white;
                }
            """)
