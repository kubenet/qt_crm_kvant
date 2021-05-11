import sys  # sys нужен для передачи argv в QApplication
import os  # Отсюда нам понадобятся методы для отображения содержимого директорий

from PyQt5 import QtWidgets
import gui  # Это наш конвертированный файл дизайна

templates = os.listdir('learn_templates')
user_list = os.listdir('user_lists')


class ExampleApp(QtWidgets.QMainWindow, gui.Ui_MainWindow):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.pushButton_2.clicked.connect(self.start_generation)  # Выполнить функцию start_generation
        self.pushButton.clicked.connect(self.browse_folder)  # Выполнить функцию browse_folder
        self.comboBox.addItems(templates)
        self.comboBox_2.addItems(user_list)

    def browse_folder(self):
        # self.listWidget.clear()  # На случай, если в списке уже есть элементы
        directory = QtWidgets.QFileDialog.getExistingDirectory(self, "Выберите папку")
        # открыть диалог выбора директории и установить значение переменной
        # равной пути к выбранной директории

        if directory:  # не продолжать выполнение, если пользователь не выбрал директорию
            for file_name in os.listdir(directory):  # для каждого файла в директории
                self.listWidget.addItem(file_name)  # добавить файл в listWidget

    def start_generation(self):
        # --- radio button --- #
        if self.radioButton.isChecked():
            print('RadioButton')
        if self.radioButton_2.isChecked():
            print('RadioButton_2')
        if self.radioButton_3.isChecked():
            print('RadioButton_3')
        # --- checkBox --- #
        if self.checkBox.isChecked():
            print('checkBox')
        elif self.checkBox_2.isChecked():
            print('checkBox_2')
        elif self.checkBox_3.isChecked():
            print('checkBox_3')
        # --- date --- #
        date1 = self.dateEdit.date()
        date2 = self.dateEdit_2.date()
        print('c {} по {}'.format(date1.toString('dd.MM.yyyy'), date2.toString('dd.MM.yyyy')))


def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение


if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()
