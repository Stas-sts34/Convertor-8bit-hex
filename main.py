import sys  # sys нужен для передачи argv в QApplication
from PyQt5 import QtWidgets, QtCore, QtGui
import design  # Это наш конвертированный файл дизайна
import os
from convertor import Converter


class ExampleApp(QtWidgets.QMainWindow, design.Ui_Dialog):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна

        self.BtnBrows.clicked.connect(self.browse_file)
        self.Convertation.clicked.connect(self.conv_file)
        self.file = ''
        self.label.setText(self.file)
        self.out_file.textEdited.connect(self.get_out_file)
        self.file_out = 'out.xlsx'

    def get_out_file(self, text):           # имя выходного файла
        self.file_out = text + '.xlsx'
        return text + '.xlsx'

    def browse_file(self):
        file = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File',
                        './',
                        'xls Files (*.xlsx)')

        if file:
            self.file = file[0]             # выбирается пусть к файлу и название
            self.label.setText(self.file)
            return file[0]
        else:
            return

    def conv_file(self):        # конвертация файла
        elem = Converter()
        if self.file:
            try:
                elem.open_xls(self.file, self.file_out)
                self.label_message.setStyleSheet('color: rgb(0, 255, 0);')
                self.label_message.setText('Convertation success ')
            except:
                self.label_message.setStyleSheet('color: rgb(255, 0, 0);')
                self.label_message.setText('Error convertation')
        else:
            self.browse_file()
            elem.open_xls(self.file, self.file_out)


def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение


if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()
