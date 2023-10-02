import multiprocessing

from PyQt5.QtWidgets import QGridLayout, QWidget, QLabel,\
                            QLineEdit, QPushButton, QFileDialog, \
                                QErrorMessage, QTextEdit, QApplication, QMessageBox
from PyQt5 import QtGui
from functions import make_items_images, get_unique_items, get_trans_group_product_manuals, to_excel
import time
import win32com.client as win32

from PIL import ExifTags, Image, ImageOps
import seaborn as sn

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        # Установка размеров окна
        self.setGeometry(100, 50, 1100, 500)
        # Установка названия приложения
        self.setWindowTitle('Распознавание веса и текста')
        self.show()

        # параметры шрифта
        font = QtGui.QFont()
        font1 = QtGui.QFont()
        font.setBold(True)
        font.setPointSize(12)  # установите нужный размер шрифта
        font1.setPointSize(12)  # установите нужный размер шрифта

        self.directory_path = ""

        # ДИРЕКТОРИЯ С ФОТО
        self.dir_photo_text = QLineEdit(self)
        self.dir_photo_text.setReadOnly(True)
        self.dir_photo_text_button = QPushButton('Выбрать', self)
        self.dir_photo_text_button.setFont(font1)
        # Обработка клика по кнопке dir_photo_text_button
        self.dir_photo_text_button.clicked.connect(self.dir_photo_text_button_clicked)
        # Создание метки QLabel
        self.dir_photo_label = QLabel('Директория с фото:', self)
        self.dir_photo_label.setFont(font)

        # ИСХОДНЫЙ ЭКСЕЛЬ
        self.excel_text = QLineEdit(self)
        self.excel_text.setReadOnly(True)
        self.excel_text_button = QPushButton('Выбрать', self)
        self.excel_text_button.setFont(font1)
        # Обработка клика по кнопке excel_text_button
        self.excel_text_button.clicked.connect(self.excel_text_button_clicked)
        # Создание метки QLabel
        self.excel_label = QLabel('Исходный Excel:', self)
        self.excel_label.setFont(font)

        # ДИРЕКТОРИЯ ДЛЯ ФИНАЛЬНОГО ЭКСЕЛЯ
        self.dir_excel_text = QLineEdit(self)
        self.dir_excel_text.setReadOnly(True)
        self.dir_excel_text_button = QPushButton('Выбрать', self)
        self.dir_excel_text_button.setFont(font1)
        # Обработка клика по кнопке excel_text_button
        self.dir_excel_text_button.clicked.connect(self.dir_excel_text_button_clicked)
        # Создание метки QLabel
        self.dir_excel_label = QLabel('Директория для финального Excel:', self)
        self.dir_excel_label.setFont(font)

        # Поле статуса
        # self.status_text = QLineEdit(self)
        self.status_text = QTextEdit(self)
        self.status_text.setReadOnly(True)
        self.status_text.setFixedHeight(50)
        self.status_label = QLabel('Статус:', self)
        self.status_label.setFont(font)
        # self.status_label.setFixedHeight(200)
        # self.status_label.setFixedWidth(400)


        # Добавление виджетов на форму
        self.layout = QGridLayout(self)
        # ДИРЕКТОРИЯ С ФОТО:
        self.layout.addWidget(self.dir_photo_label,0,0)
        self.layout.addWidget(self.dir_photo_text_button, 0, 1)
        self.layout.addWidget(self.dir_photo_text,0,2)
        # ИСХОДНЫЙ ЭКСЕЛЬ:
        self.layout.addWidget(self.excel_label, 1, 0)
        self.layout.addWidget(self.excel_text_button, 1, 1)
        self.layout.addWidget(self.excel_text, 1, 2)
        # ДИРЕКТОРИЯ ДЛЯ ФИНАЛЬНОГО ЭКСЕЛЯ:
        self.layout.addWidget(self.dir_excel_label, 2, 0)
        self.layout.addWidget(self.dir_excel_text_button, 2, 1)
        self.layout.addWidget(self.dir_excel_text, 2, 2)


        # создаем кнопку "Распознавание и сохранение в новом Excel"
        self.button = QPushButton('Распознавать и сохранить новый Excel')
        self.button.clicked.connect(self.run)
        self.button.setFont(font)
        self.layout.addWidget(self.button, 3, 0)
        # Поле статуса
        self.layout.addWidget(self.status_label, 3, 1)
        self.layout.addWidget(self.status_text, 3, 2)

    def dir_photo_text_button_clicked(self):
        # Вызов QFileDialog для выбора директории
        directory_path = QFileDialog.getExistingDirectory(self, 'Выбрать директорию', '', QFileDialog.ShowDirsOnly)

        if directory_path:
            self.directory_path = directory_path
            self.dir_photo_text.setText(self.directory_path)
            self.status_text.setText('')


    def excel_text_button_clicked(self):
        # Вызов QFileDialog для выбора директории
        file_path, _ = QFileDialog.getOpenFileName(self, 'Выбрать файл', '',
                                                   'Таблицы (*.xlsx);;') # Все файлы (*.*);;
        if file_path:
            if file_path.endswith('.xls'):
                message_box = QMessageBox()
                message_box.setWindowTitle('Предупреждение')
                message_box.setText('Исходный файл Excel в недопустимом формате .xls\n'
                                    'Пожалуйста, пересохраните файл в формате .xlsx')
                message_box.setStandardButtons(QMessageBox.Ok)
                message_box.exec_()
                excel = win32.Dispatch('Excel.Application')
                excel.Visible = True
                excel.Workbooks.Open(file_path, CorruptLoad=True)
            else:
                self.file_path = file_path
                self.excel_text.setText(self.file_path)
                self.status_text.setText('')


    def dir_excel_text_button_clicked(self):
        # Вызов QFileDialog для выбора директории
        directory_path = QFileDialog.getExistingDirectory(self, 'Выбрать директорию', '', QFileDialog.ShowDirsOnly)
        if directory_path:
            self.directory_path = directory_path
            self.dir_excel_text.setText(self.directory_path)
            self.status_text.setText('')



    def run(self):

        # записываем значения полей в переменные и выполняем необходимое действие
        # print(self.dir_photo_text.text())
        # print(self.excel_text.text())
        # print(self.dir_excel_text.text())
        if self.dir_photo_text.text() and self.excel_text.text() and self.dir_excel_text.text():
            time_start = time.time()
            BATCH_PATH = self.dir_photo_text.text()  # путь к папке с партией
            PATH_TO_EXCEL_ORIGIN = self.excel_text.text()
            PATH_TO_EXCEL_RESULT = self.dir_excel_text.text()

            unique_artickles = get_unique_items(PATH_TO_EXCEL_ORIGIN)  # список артикулов из экселя

            # распознавание полей и весов (выход - словари с картинками-вырезками)
            items_images = make_items_images(BATCH_PATH)

            self.status_text.setText(f'Артикулы: {unique_artickles}')
            # формирование словарей по item
            trans, group, product, manuals, n = get_trans_group_product_manuals(items_images, unique_artickles)

            # запись в эксель
            status = to_excel(PATH_TO_EXCEL_ORIGIN,
                              PATH_TO_EXCEL_RESULT,
                              unique_artickles,
                              trans,
                              group,
                              product,
                              manuals)

            duration = int(time.time() - time_start)

            self.status_text.setText(status + '\n' + 'Время выполнения:' + str(duration) + ' c')
            self.status_text.setStyleSheet('color: green')
        else:
            error = QErrorMessage()
            error.setWindowTitle('Ошибка')
            error.showMessage('Заполните все поля')
            error.exec_()


if __name__ == '__main__':
    multiprocessing.freeze_support()
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()