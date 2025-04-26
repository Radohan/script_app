import sys
import os

# Add the current directory to Python's path if not already there
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from PyQt5.QtWidgets import QApplication, QStyleFactory
from PyQt5.QtGui import QFont

# Now import the main window class
from ui.main_window import MXLIFFParser  # Use direct import instead of package import

def main():
    app = QApplication(sys.argv)

    # Set app-wide font
    app.setFont(QFont("Segoe UI", 10))

    # Use Fusion style as base
    app.setStyle(QStyleFactory.create('Fusion'))

    window = MXLIFFParser()
    window.show()

    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
