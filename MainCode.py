# -*- coding: utf-8 -*-
import email
import imaplib
import numpy as np
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog



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
        self.pushButton.setGeometry(QtCore.QRect(10, 10, 300, 150))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.pushButton.setFont(font)
        self.pushButton.setStyleSheet("background-color: rgb(0, 85, 255);")
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(10, 170, 300, 150))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setStyleSheet("background-color: rgb(0, 85, 255);")
        self.pushButton_2.setObjectName("pushButton_2")
        self.Orders_list_to_processing = QtWidgets.QTableWidget(self.centralwidget)
        self.Orders_list_to_processing.setGeometry(QtCore.QRect(320, 10, 470, 680))
        self.Orders_list_to_processing.setStyleSheet("background-color: rgb(158, 251, 255);")
        self.Orders_list_to_processing.setObjectName("Orders_list_to_processing")
        self.Orders_list_to_processing.setColumnCount(2)
        self.Orders_list_to_processing.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        self.Orders_list_to_processing.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.Orders_list_to_processing.setHorizontalHeaderItem(1, item)
        self.Orders_list_to_processing.horizontalHeader().setDefaultSectionSize(235)
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(10, 330, 300, 150))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setStyleSheet("background-color: rgb(0, 85, 255);")
        self.pushButton_3.setObjectName("pushButton_3")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.add_function()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pushButton.setText(_translate("MainWindow", "Take list of paid orders \n"
                                                         " from mail to processing"))
        self.pushButton_2.setText(_translate("MainWindow", "Load file from Excel"))
        item = self.Orders_list_to_processing.horizontalHeaderItem(0)
        item.setText(_translate("MainWindow", "ID"))
        item = self.Orders_list_to_processing.horizontalHeaderItem(1)
        item.setText(_translate("MainWindow", "Order number"))
        self.pushButton_3.setText(_translate("MainWindow", "process in 1C"))

    def add_function(self):
        self.pushButton.clicked.connect(lambda: self.order_processing())  # кнопка обработки писем
        self.pushButton_2.clicked.connect(lambda: self.action_cliked())
        self.pushButton_3.clicked.connect(lambda: self.action_cliked())

    # def action_cliked(self):  # активирует клик согласно надписи на кнопке
    #     action = self.sender()
    #     if action.text() == "Take list of paid orders \n from mail to processing":
    #         self.order_processing()
    #         f = open('Y:\Виталий\Новая папка\Программирование\Phyton\MyProjectForWork/OrderNumbers.xlsx', 'r')
    #     elif action.text() == "Load file from Excel":
    #         fname = QFileDialog.getOpenFileName(self)[0]
    #
    #         f = open(fname, 'r')
    #         with f:
    #             data = f.read()

    def order_processing(self):  # основная функция - обработка заказов с почты
        mail = imaplib.IMAP4_SSL('mail.fitnessbar.ru')
        mail.login('region-zakaz@fitnessbar.ru', '222QQq222')  # заходим на сервер
        mail.list()  # получаем список папок
        mail.select('inbox', readonly=False)  # выбираем папку
        result, data = mail.search(None, "ALL")  # Получаем массив со списком найденных почтовых сообщений
        ids = data[0]  # Сохраняем в переменную ids строку с номерами писем
        id_list = ids.split()  # Получаем массив номеров писем
        df = pd.DataFrame(columns=['Order number'])  # создаём датафрейм с названным столбцом

        for num in range(len(id_list)):
            latest_email_id = id_list[
                -1 - num]  # Задаем переменную latest_email_id, значением будет номер последнего письма

            result, data = mail.fetch(latest_email_id,
                                      "(RFC822)")  # Получаем письмо с идентификатором latest_email_id (последнее письмо).
            raw_email = data[0][1]  # В переменную raw_email заносим необработанное письмо
            raw_email_string = raw_email.decode(
                'utf-8')  # Переводим текст письма в кодировку UTF-8 и сохраняем в переменную raw_email_string
            df.loc[len(df)] = order_number(raw_email_string)
            #  copy_and_delete_mail(latest_email_id)  # переносим в папку и удаляем обработанное письмо

        df.index = np.arange(1, len(df) + 1)  # начинаем номерацию индексов с единицы
        df.index.name = 'ID'  # задаём название столбца индексов
        df.to_excel(
            'Y:\Виталий\Новая папка\Программирование\Phyton\MyProjectForWork/OrderNumbers.xlsx')  # экспорт в Excel
        mail.close()  # закрываем соединение


def order_number(raw_email_string):  # функция возвращает номер заказа из текста письма
    email_message = email.message_from_string(raw_email_string)  # Получаем заголовки и тело письма и в перем email
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


if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

