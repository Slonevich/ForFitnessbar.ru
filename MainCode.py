# -*- coding: utf-8 -*-
import email
import imaplib
import pandas as pd
import os
import os.path
import xlsxwriter
import configparser
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QTableWidgetItem, QMessageBox, QStyledItemDelegate
import tkinter as tk
from tkinter import Tk, filedialog, simpledialog
from tkinter.filedialog import askopenfilename
from functools import partial
from datetime import datetime


class ReadOnlyDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        return


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 700)
        MainWindow.setLayoutDirection(QtCore.Qt.LeftToRight)
        MainWindow.setStyleSheet("background-color: rgb(85, 255, 0);")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(0, 0, 800, 700))
        font = QtGui.QFont()
        font.setPointSize(44)
        self.label.setFont(font)
        self.label.setLineWidth(1)
        self.label.setText("")
        self.label.setTextFormat(QtCore.Qt.AutoText)
        self.label.setScaledContents(False)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(10, 10, 300, 120))
        font.setPointSize(15)
        self.pushButton.setFont(font)
        self.pushButton.setStyleSheet("background-color: rgb(0, 85, 255);")
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(10, 140, 300, 120))
        self.pushButton_2.setFont(font)
        self.pushButton_2.setStyleSheet("background-color: rgb(0, 85, 255);")
        self.pushButton_2.setObjectName("pushButton_2")
        self.Orders_list_to_processing = QtWidgets.QTableWidget(self.centralwidget)
        self.Orders_list_to_processing.setGeometry(QtCore.QRect(320, 10, 460, 630))
        self.Orders_list_to_processing.setStyleSheet("background-color: rgb(158, 251, 255);")
        self.Orders_list_to_processing.setObjectName("Orders_list_to_processing")
        self.Order_to_processing = QtWidgets.QTableWidget(self.centralwidget)
        self.Order_to_processing.setGeometry(QtCore.QRect(10, 270, 300, 120))
        self.Order_to_processing.setStyleSheet("background-color: rgb(158, 251, 255);")
        self.Order_to_processing.setObjectName("Order_to_processing")
        self.Order_to_processing.setRowCount(1)
        self.Order_to_processing.setColumnCount(2)
        self.Order_to_processing.setHorizontalHeaderLabels(['ID', 'Order number'])
        self.Order_to_processing.setColumnWidth(0, 140)
        self.Order_to_processing.setColumnWidth(1, 140)
        delegate = ReadOnlyDelegate()
        self.Order_to_processing.setItemDelegateForColumn(0, delegate)
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(10, 400, 300, 120))
        self.pushButton_3.setFont(font)
        self.pushButton_3.setStyleSheet("background-color: rgb(0, 85, 255);")
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(100, 340, 130, 40))
        self.pushButton_4.setStyleSheet("background-color: rgb(85, 85, 127);")
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_5 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_5.setGeometry(QtCore.QRect(10, 530, 300, 120))
        self.pushButton_5.setFont(font)
        self.pushButton_5.setStyleSheet("background-color: rgb(0, 85, 255);")
        self.pushButton_5.setObjectName("pushButton_5")


        self.checkBox = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox.setEnabled(True)
        self.checkBox.setChecked(True)
        self.checkBox.setGeometry(QtCore.QRect(20, 650, 170, 15))
        self.checkBox.setStyleSheet("alternate-background-color: rgb(85, 255, 255);")
        self.checkBox.setTristate(False)
        self.checkBox.setObjectName("checkBox")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(210, 650, 570, 15))
        self.label_2.setStyleSheet("background-color: rgb(0, 255, 255);")
        self.label_2.setObjectName("label_2")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.display_table()  # отображаем файл со списком заказов, если он есть - либо новый, пустой
        self.add_function()  # добавляем функционал кнопкам

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pushButton.setText(_translate("MainWindow", "Выгрузить список оплаченных \n"
                                                         "заказов с почты"))
        self.pushButton_2.setText(_translate("MainWindow", "Загрузить файл"))
        self.pushButton_3.setText(_translate("MainWindow", "Обработать в 1C"))
        self.pushButton_4.setText(_translate("MainWindow", "добавить заказ"))
        self.pushButton_5.setText(_translate("MainWindow", "Обработать отчёты \n"
                                                           "Solid Express"))
        self.checkBox.setText(_translate("MainWindow", "Путь к файлу по умолчанию"))
        self.label_2.setText(_translate("MainWindow", config['DEFAULT']['work_dir']))

    def add_function(self):  # добавляем функции кнопкам
        self.pushButton.clicked.connect(lambda: order_processing())
        self.pushButton.clicked.connect(lambda: self.display_table())
        self.pushButton_2.clicked.connect(lambda: self.file_to_open())
        self.pushButton_3.clicked.connect(lambda: self.processing_in_1c())
        self.pushButton_4.clicked.connect(lambda: self.add_order())
        self.pushButton_5.clicked.connect(lambda: self.report_process())
        enable_slot = partial(self.enable_mod)
        disable_slot = partial(self.disable_mod)
        self.checkBox.stateChanged.connect(lambda x: enable_slot() if x else disable_slot())

    def report_process(self):
        folder_name = 'Reports'
        mail, id_list = email_data(folder_name)

        for num in range(1, len(id_list) + 1):
            latest_email_id = id_list[-num]  # Задаем переменную, значением которой будет номер последнего письма

            result, data = mail.fetch(latest_email_id,
                                      "(RFC822)")  # Получаем последнее письмо).
            raw_email = data[0][1]  # В переменную raw_email заносим необработанное письмо
            try:
                email_message = email.message_from_string(raw_email)
            except TypeError:
                email_message = email.message_from_bytes(raw_email)
            for part in email_message.walk():
                if "application" in part.get_content_type():
                    current_datetime = datetime.now()
                    now = str(datetime.now().date()) + '-' + str(current_datetime.hour) + '-' + \
                          str(current_datetime.minute) + '-' + str(current_datetime.second) + '-' + str(num)
                    filename = now + '.xls'
                    fp = open(os.path.join(config['DEFAULT']['work_dir'], filename), 'wb')
                    fp.write(part.get_payload(decode=1))
            # copy_and_delete_mail(latest_email_id)  # переносим в папку и удаляем обработанное письмо
        mail.close()
        mail.logout()

    def enable_mod(self):
        self.checkBox.setText("Путь к файлу по умолчанию")
        self.label_2.setText(os.path.abspath(os.curdir))
        config['DEFAULT']['work_dir'] = os.path.abspath(os.curdir)
        with open('config.ini', 'w') as configfile:
            config.write(configfile)

    def disable_mod(self):
        Tk().withdraw()
        new_dir_name = filedialog.askdirectory()
        self.label_2.setText("Новый путь к файлу: " + new_dir_name)
        config['DEFAULT']['work_dir'] = new_dir_name
        with open('config.ini', 'w') as configfile:
            config.write(configfile)

    def add_order(self):  # добавление файла к обработке ручным вводом
        df2 = pd.DataFrame(columns=['ID', 'Order number', 'Processed'])  # создаём датафрейм с названным столбцом
        df = pd.read_excel(config['DEFAULT']['work_dir'] + '\\' + config['DEFAULT']['order_list_f_name'])
        self.Order_to_processing.setItem(0, 0, QtWidgets.QTableWidgetItem(str(len(df)+1)))  # ID по таблице
        df2.loc[0] = [int(self.Order_to_processing.item(0, 0).text()), int(self.Order_to_processing.item(0, 1).text()), 'NO']
        frames = [df, df2]
        result = pd.concat(frames)  # объединение данных файла и обработки
        result = result.sort_values(by=['ID'])  # сортируем по колонке ID
        result.to_excel(config['DEFAULT']['work_dir'] + '\\' + config['DEFAULT']['order_list_f_name'], index=False)  # экспорт в Excel
        self.display_table()

    def processing_in_1c(self):  # обработка в 1С + отметка
        df = pd.read_excel(config['DEFAULT']['work_dir'] + '\\' + config['DEFAULT']['order_list_f_name'])
        for i in range(len(df)):
            df.at[i, 'Processed'] = 'YES'
        df.to_excel(config['DEFAULT']['work_dir'] + '\\' + config['DEFAULT']['order_list_f_name'], index=False)  # экспорт в Excel без столбца индексов
        self.display_table()

    def file_to_open(self):  # функция открытия Excel файла с проверкой расширения
        Tk().withdraw()
        file_path = askopenfilename()
        directory = os.path.split(file_path)[0]
        fname = os.path.split(file_path)[1]
        filename, file_extension = os.path.splitext(file_path)
        if file_extension == '.xlsx':
            # global work_dir
            config['DEFAULT']['work_dir'] = directory
            config['DEFAULT']['order_list_f_name'] = fname
            with open('config.ini', 'w') as configfile:
                config.write(configfile)
            self.display_table()
        else:
            error = QMessageBox()
            error.setWindowTitle('Ошибка')
            error.setText('Выбран файл неверного формата. Выберите файл .xlsx')
            error.setIcon(QMessageBox.Warning)
            error.setStandardButtons(QMessageBox.Ok)
            error.buttonClicked.connect(self.error_popup_action)
            error.exec_()

    def display_table(self):  # добавляет данные в виджет программы из выбранной таблицы
        df = pd.read_excel(config['DEFAULT']['work_dir'] + '\\' + config['DEFAULT']['order_list_f_name'])
        df.fillna('', inplace=True)  # заменяем\смещаем пустые ячейки
        self.Orders_list_to_processing.setRowCount(df.shape[0])  # задаём количество строк из файла
        self.Orders_list_to_processing.setColumnCount(df.shape[1])  # задаём количество столбцов из файла
        self.Orders_list_to_processing.setHorizontalHeaderLabels(df.columns)  # задаём названия столбцов из файла
        # возвращает pandas array object
        for row in df.iterrows():
            values = row[1]
            for col_index, value in enumerate(values):
                tableitem = QTableWidgetItem(str(value))
                self.Orders_list_to_processing.setItem(row[0], col_index, tableitem)

    def error_popup_action(self, btn):  # кнопка всплывающего окна ошибки
        if btn.text() == "OK":
            self.file_to_open()


def email_data(folder_name):  # функция возвращает письма из выбранной папки
    mail = imaplib.IMAP4_SSL(config['MAIL']['mail_server_name'])
    mail.login(config['MAIL']['login'], config['MAIL']['password'])  # заходим на сервер
    mail.list()  # получаем список папок
    mail.select(folder_name, readonly=False)  # выбираем папку
    result, data = mail.search(None, "ALL")  # Получаем массив со списком найденных почтовых сообщений
    ids = data[0]  # Сохраняем в переменную ids строку с номерами писем
    id_list = ids.split()  # Получаем массив номеров писем
    return mail, id_list


def order_processing():  # основная функция - обработка заказов с почты
    folder_name = 'inbox'
    mail, id_list = email_data(folder_name)
    df = pd.DataFrame(columns=['ID', 'Order number', 'Processed'])  # создаём датафрейм с названным столбцом

    for num in range(1, len(id_list)+1):
        latest_email_id = id_list[-num]  # Задаем переменную, значением которой будет номер последнего письма

        result, data = mail.fetch(latest_email_id,
                                  "(RFC822)")  # Получаем последнее письмо).
        raw_email = data[0][1]  # В переменную raw_email заносим необработанное письмо
        raw_email_string = raw_email.decode(
            'utf-8')  # Переводим текст письма в кодировку UTF-8 и сохраняем в переменную raw_email_string
        df.loc[len(df)] = [num, order_number(raw_email_string), 'NO']
    # copy_and_delete_mail(latest_email_id)  # переносим в папку и удаляем обработанное письмо
    df2 = pd.read_excel(config['DEFAULT']['work_dir'] + '\\' + config['DEFAULT']['order_list_f_name'])
    df.loc[df['ID'] <= len(df2)+1, 'ID'] = df['ID']+len(df2)
    frames = [df, df2]
    result = pd.concat(frames)  # объединение данных файла и обработки
    result = result.sort_values(by=['ID'])  # сортируем по колонке ID
    result.to_excel(config['DEFAULT']['work_dir'] + '\\' + config['DEFAULT']['order_list_f_name'],
                    index=False)  # экспорт в Excel
    mail.close()  # закрываем соединение


def input_mail_settings():  # функция предлагает пользователю ввести настройки почты для обработки
    application_window = tk.Tk()
    application_window.withdraw()
    answer = simpledialog.askstring("Input", "Введите название почтового сервера: ",
                                    parent=application_window)
    if answer:
        config['MAIL']['mail_server_name'] = answer
    answer = simpledialog.askstring("Input", "Введите логин почты: ",
                                    parent=application_window)
    if answer:
        config['MAIL']['login'] = answer
    answer = simpledialog.askstring("Input", "Введите пароль: ",
                                    parent=application_window)
    if answer:
        config['MAIL']['password'] = answer


def create_default_config():  # создаём файл со стандартными настройками
    config['DEFAULT'] = {'order_list_f_name': 'OrderNumbers.xlsx', 'work_dir': os.path.abspath(os.curdir)}
    config['MAIL'] = {}
    input_mail_settings()
    with open('config.ini', 'w') as configfile:
        config.write(configfile)


def order_number(raw_email_string):  # функция возвращает номер заказа из текста письма
    email_message = email.message_from_string(raw_email_string)  # Получаем заголовки и тело письма в перем email
    # _message. Записываем в ту же переменную, где было необработанное письмо

    from email.header import decode_header
    tema = decode_header(email_message['Subject'])
    el = str(tema).split()
    processed_order_number= ''
    for i in el[2]:  # перебираем часть строки из темы письма, в которой содержится номер до лишних символов
        if i == "_" or i == "-":
            break
        processed_order_number = processed_order_number + i
    return processed_order_number  # обработанный номер заказа


def copy_and_delete_mail(mail_id):  # Функция копирует письмо с переданным айди в указанную папку и удаляет из текущей
    copy_res = mail.copy(mail_id, 'Processed')  # копируем письмо в указанную папку
    if copy_res[0] == 'OK':  # удаляем письмо если перемещение успешно
        mail.store(mail_id, '+FLAGS', '\\Deleted')  # помечает флагом письма к удалению
        mail.expunge()  # удаляет отмеченные письма


def start_app_files_check():  # проверка наличия файлов конфига и таблицы заказов
    check_config_file = os.path.exists('config.ini')
    if not check_config_file:
        create_default_config()
    config.read('config.ini')
    check_order_list_file = os.path.exists(config['DEFAULT']['work_dir'] + '\\'
                                           + config['DEFAULT']['order_list_f_name'])
    if not check_order_list_file:  # проверка существует ли уже файл выгрузки
        workbook = xlsxwriter.Workbook(config['DEFAULT']['work_dir'] + '\\' + config['DEFAULT']['order_list_f_name'])
        workbook.add_worksheet()
        workbook.close()
        empty_df = pd.DataFrame(columns=['ID', 'Order number', 'Processed'])  # создаём датафрейм с названными столбцами
        empty_df.to_excel(config['DEFAULT']['work_dir'] + '\\' + config['DEFAULT']['order_list_f_name'],
                          index=False)  # экспорт в Excel


if __name__ == '__main__':
    import sys
    config = configparser.ConfigParser()
    start_app_files_check()
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

