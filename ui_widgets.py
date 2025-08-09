from PyQt5.QtWidgets import QComboBox, QListView

class HoverComboBox(QComboBox):
    def __init__(self):
        super().__init__()
        view = QListView()
        view.setStyleSheet("""
            QListView::item:hover {
                background-color: #cce5ff;
                font-weight: bold;
            }
        """)
        self.setView(view)
        self.setMaximumWidth(200)