from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMainWindow

from Menu.interface.MenuUI import Ui_MainWindow
from Menu.pivot_sheets import DialogWindow


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.__event_connection()

    def __event_connection(self):
        self.pushButton.clicked.connect(self.open_1)
        self.menu_help.triggered.connect(lambda x: print(6))
        action_show_poem = QtWidgets.QAction('По сводке', self)
        action_show_poem.triggered.connect(self.show_pivot_dialog)
        self.menu_help.addAction(action_show_poem)



    def show_pivot_dialog(self):
        poem_text = """
           Сведение замеров
           
           1. Получение списка листов из excel документа
           2. Вычленение важной информации из каждого листа
           3. Составление сводной таблицы и сохранение в одной директории с текущим файлом
           """
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle('По сведению замеров')
        layout = QtWidgets.QVBoxLayout()
        label = QtWidgets.QLabel(poem_text)
        layout.addWidget(label)
        dialog.setLayout(layout)
        dialog.exec_()

    def open_1(self):
        self.window = DialogWindow()
        self.window.show()