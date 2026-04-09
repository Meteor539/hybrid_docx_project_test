import sys
from PyQt6.QtWidgets import QApplication
from gui.main_window import MainWindow

def main():
    """
    程序入口函数
    """
    app = QApplication(sys.argv)
    app.setStyle("WindowsVista")  
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
