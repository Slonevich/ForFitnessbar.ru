# -*- coding: utf-8 -*-
import email
import imaplib
import pandas as pd


def order_number(raw_email_string): # функция возвращает номер заказа из текста письма
    email_message = email.message_from_string(raw_email_string)  # Получаем заголовки и тело письма и заносим результат в переменную email
# _message. Записываем в ту же переменную, где было необработанное письмо

    from email.header import decode_header
    theme = decode_header(email_message['Subject'])
    el = str(theme).split()
    letter_subject = ''
    for i in el[2]: # перебираем часть строки из темы письма, в которой содержится номер до знака _, который добавляет банк
# разобраться с байтовыми строкаи 0 и 1 элементы
        if i == "_" or i == "-":
            break
        letter_subject = letter_subject + i
    return letter_subject  # тема письма
# переназвать выводящуюся переменную


def delete_mail(mail_id): # Функция копирует письмо с переданным айди в указанную папку и удаляет из текущей
    copy_res = mail.copy(mail_id, 'Processed') # копируем письмо в указанную папку
    if copy_res[0] == 'OK': # удаляем письмо если перемещение успешно
        mail.store(mail_id, '+FLAGS', '\\Deleted')
        mail.expunge()


# if __init__ == '__main__':

mail = imaplib.IMAP4_SSL('mail.ru') # mail.fitnessbar.ru
mail.login('slone123@mail.ru', '') # заходим на сервер раб почта region-zakaz@fitnessbar.ru', '222QQq222
mail.list() # получаем список папок
mail.select('test', readonly = False) # выбираем папку inbox
result, data = mail.search(None, "ALL") # Получаем массив со списком найденных почтовых сообщений

ids = data[0] # Сохраняем в переменную ids строку с номерами писем
id_list = ids.split() # Получаем массив номеров писем

df = pd.DataFrame({}) # создаём пустой словарь под DataFrame для Excel

for num in range(len(id_list)):
    latest_email_id = id_list[-1-num]  # Задаем переменную latest_email_id, значением которой будет номер последнего письма
# можно не брать последнее письмо
    result, data = mail.fetch(latest_email_id,
                              "(RFC822)")  # Получаем письмо с идентификатором latest_email_id (последнее письмо).
    raw_email = data[0][1]  # В переменную raw_email заносим необработанное письмо
    raw_email_string = raw_email.decode(
        'utf-8')  # Переводим текст письма в кодировку UTF-8 и сохраняем в переменную raw_email_string
    df[num+1] = [order_number(raw_email_string)] # формируем массив обработанных заказов
# нужно записать в столбец
    delete_mail(latest_email_id) # переносим в папку и удаляем обработанное письмо

df.to_excel('Y:\Виталий\Новая папка\Программирование\Phyton\MyProjectForWork/OrderNumbers.xlsx')

mail.close() #закрываем соединение