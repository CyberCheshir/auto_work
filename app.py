from PyQt5.QtWidgets import QApplication

from Menu.menu import MainWindow


def main():
    app = QApplication([])
    menu = MainWindow()
    menu.show()
    app.exec_()


if __name__ == '__main__':
    main()
