import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


class LauncherWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MicroCT Analysis Suite")
        self.setFixedSize(320, 190)

        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(36, 28, 36, 28)

        title = QLabel("MicroCT Analysis Suite")
        title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        fkbp5_btn = QPushButton("FKBP5 Analysis")
        fkbp5_btn.setFixedHeight(42)
        fkbp5_btn.clicked.connect(self._open_fkbp5)
        layout.addWidget(fkbp5_btn)

        normal_btn = QPushButton("Normal Study")
        normal_btn.setFixedHeight(42)
        normal_btn.clicked.connect(self._open_normal)
        layout.addWidget(normal_btn)

    def _open_fkbp5(self):
        from FKBP5Heat_MicroCT_Analysis import MainWindow as FKBP5Window
        self._child = FKBP5Window()
        self._child.show()
        self.close()

    def _open_normal(self):
        from MicroCT_Analysis import MainWindow as NormalWindow
        self._child = NormalWindow()
        self._child.show()
        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = LauncherWindow()
    window.show()
    sys.exit(app.exec_())
