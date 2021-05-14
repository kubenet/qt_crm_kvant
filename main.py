import sys  # sys нужен для передачи argv в QApplication
import os  # Отсюда нам понадобятся методы для отображения содержимого директорий

from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QAction

import gui  # Это наш конвертированный файл дизайна
import shutil
from pathlib import Path
from docx2pdf import convert
import openpyxl
import openpyxl.utils
from docxtpl import DocxTemplate

templates = os.listdir('learn_templates')
user_list = os.listdir('user_lists')


def view_diplomas(self):
    os.system('explorer.exe "diplomas"')


class ExampleApp(QtWidgets.QMainWindow, gui.Ui_MainWindow):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам и т.д. в файле design.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.pushButton_2.clicked.connect(self.start_generation)    # Выполнить функцию start_generation
        self.pushButton.clicked.connect(self.browse_folder)         # Выполнить функцию browse_folder
        self.pushButton_3.clicked.connect(view_diplomas)    # связь кнопки открыть папку с результатом (дипломами)
        self.comboBox.addItems(templates)       # заполнение comboBox списком файлов шаблонов
        self.comboBox_2.addItems(user_list)     # заполнение comboBox списком файлов групп
        self.dateEdit.setCalendarPopup(True)    # настройка даты опция выпадающего календаря
        self.dateEdit.setDateTime(QtCore.QDateTime.currentDateTime())       # устанавливаем текущую дату
        self.dateEdit_2.setDateTime(QtCore.QDateTime.currentDateTime())     # устанавливаем текущую дату
        self.dateEdit_2.setCalendarPopup(True)  # настройка даты опция выпадающего календаря
        self.progressBar.setValue(0)            # настройка первоначального значения progressBar
        self.menu.setEnabled(False)             # неактивный пункт меню "Справка"
        self.action = QAction("Выход")          # создание нового пункта меню
        self.menu_2.addAction(self.action)      # добавление в меню пункт "Выход"
        self.action.triggered.connect(self.close)   # кнопка закрытия приложения
        self.checkBox.setChecked(True)              # по умолчанию конвертация в pdf включена

    def browse_folder(self):
        file_name = QtWidgets.QFileDialog.getOpenFileName(self, "Выбор шаблона", None, "word (*.doc *.docx)")[0]
        self.label_5.setText(file_name)

    def start_generation(self):
        context = {}
        template = str(self.comboBox.currentText())
        group_list = str(self.comboBox_2.currentText())
        # --- radio button --- #
        if self.radioButton.isChecked():
            print('Базовая группа')
        elif self.radioButton_2.isChecked():
            print('Проектная группа')
        elif self.radioButton_3.isChecked():
            print('Свой шаблон')
        # --- checkBox --- #
        if self.checkBox_2.isChecked():
            print('Сохранить в один файл')
        if self.checkBox_3.isChecked():
            print('Дата начала и окончания')

        self.label_5.setText("Готовим Docx...")
        lists_path = Path('user_lists')
        pattern_path = Path('learn_templates')
        shutil.rmtree("diplomas", ignore_errors=True)
        os.mkdir("diplomas")
        i = 0   # счетчик для именования файлов
        loading = 0     # значение для индикации загрузки
        wb = openpyxl.load_workbook(lists_path / group_list)
        sheet = wb.active
        rows = sheet.max_row
        step = 100 / rows   # вычисление шага преодразования одного документа для индикации
        # (первая строка из документа Excel)
        pattern_name = template     # название шаблона
        doc = DocxTemplate(pattern_path / pattern_name)
        date1 = self.dateEdit.date().toString('dd.MM.yyyy')
        date2 = self.dateEdit_2.date().toString('dd.MM.yyyy')
        for row_num in range(2, rows + 1):
            line = sheet.cell(row=row_num, column=1).value + ' ' + \
                   sheet.cell(row=row_num, column=2).value + ' ' + \
                   sheet.cell(row=row_num, column=3).value
            context['fio'] = line
            context['kvant'] = str(sheet.cell(row=1, column=1).value)  # название учебной программы
            context['date1'] = date1    # дата начала обучения по программе
            context['date2'] = date2    # дата окончания обучения по программе
            context['duration'] = str(72)   # в объеме 72 часов
            loading += step
            doc.render(context)
            name_document = str(i) + "_" + str(sheet.cell(row=row_num, column=1).value) + ".docx"
            doc.save(name_document)
            shutil.move(name_document, "diplomas")
            i += 1
            self.progressBar.setValue(int(loading))
        self.progressBar.setValue(100)
        if self.checkBox.isChecked():   # если выбран пункт сохранить в формате pdf,
                                        # то все документы docx из папки diplomas переконвертируются в pdf
            self.label_5.setText("Готовим PDF...")
            convert("diplomas/")
            self.label_5.setText("Свежие PDF готовы!)")


def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение


if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()
