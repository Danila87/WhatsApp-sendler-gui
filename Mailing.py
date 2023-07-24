import json
import time
import requests
import Formatting_numbers
import Support_function

from Excel_work import excel_table
from colorama import init
from docxtpl import DocxTemplate
from GreenAPI import gr_api


init(autoreset=True)


def main_mailing(progress_bar, context_col):

    """
    Главная функция рассылки, которая запускает вспомогательные функции.
    Проверяет каждый номер на его регистрацию в WhatsApp. Неудачные номера или клиентов записывает на новый лист
    """

    Formatting_numbers.update_number()  # Запускаем форматирование номеров
    excel_table.create_new_sheet()  # Создаём\очищаем лист с результатами

    max_row = excel_table.get_sheet_row()  # Длина строк
    columns = excel_table.selected_column  # Колонки с номерами

    doc = DocxTemplate("content.docx")  # Открываем шаблон

    error = None

    successful_attempt, failed_attempt, num_row = 0, 0, 0  # Счётчики (Успешные попытки, неудачные попытки, текущая
    # строка)

    path_to_word = 'content.docx'

    step_pb = 100//max_row

    for row in excel_table.sheet_main.iter_rows(min_row=2, max_row=max_row, values_only=True):  # Цикл в длину строк

        num_row += 1
        list_failed_numbers = []  # Список для проверки отправки сообщений
        list_full = []  # Список с номерами из всех колонок

        for col in columns:  # Формируем список всех номеров по всем выбранным колонкам
            if row[col] is None:
                continue

            list_numbers = str(row[col]).split(';')
            list_full += list_numbers

        if not list_full:  # Если список пустой, пишем на доп лист и пропускаем итерацию
            failed_attempt += 1
            excel_table.write_additional_sheet(row=row, num_row=num_row, status=False)
            progress_bar.update_value(step=step_pb, current_value=num_row)
            continue

        if context_col:  # Если ключи в docx есть, то формируем словарь с выбранными значениями и пишем в docx_2
            context = {i: row[context_col[i]] for i in context_col}
            doc.render(context)
            doc.save('content_2.docx')
            path_to_word = 'content_2.docx'

        list_full = list(set(list_full))

        for number in list_full:  # Проходимся по каждому номеру и отправляем сообщение, результат false\true пишем в
            # список

            if check_number(number) == 466:
                error = "Вы исчерпали лимит проверок. Проверьте свой тарифный план"
                break

            elif check_number(number) == 200:
                list_failed_numbers.append(False)
                continue

            send_message(number, path_to_word)
            list_failed_numbers.append(True)
            time.sleep(3)

        if all(x is False for x in list_failed_numbers):  # Если по всем номерам не получилось отправить - пишем на
            # доп лист

            failed_attempt += 1
            excel_table.write_additional_sheet(row=row, num_row=num_row, status=False)
            progress_bar.update_value(step=step_pb, current_value=num_row)
            continue

        successful_attempt += 1
        excel_table.write_additional_sheet(row=row, num_row=num_row)  # То же самое только пишем успех
        progress_bar.update_value(step=step_pb, current_value=num_row)

    return successful_attempt, failed_attempt, error


def check_number(phone_number: str) -> bool or int:
    """
    Функция проверяет Зарегистрирован ли номер в WhatsApp или нет

    :param phone_number: Проверяемый номер телефона

    :return: Возвращает True или False в зависимости от результата. Если номер зарегистрирован, то True,
     в обратном случае False
    """

    url = f"https://api.green-api.com/waInstance{gr_api.id_instance}/checkWhatsapp/{gr_api.api_token}"

    payload = json.dumps({"phoneNumber": phone_number})
    headers = {
        'Content-Type': 'application/json'
    }

    response = requests.request("POST", url, headers=headers, data=payload)

    if response.status_code == 466:
        return 466

    if response.status_code == 200 and response.json()['existsWhatsapp'] is False:
        return 200

    return True


def send_message(phone_number: str, path_to_word: str = "content.docx") -> bool:
    """
    Функция рассылки сообщения

    :param path_to_word: Путь к файлу, откуда брать данные
    :param phone_number: Номер телефона
    :return: Ничего не возвращает
    """

    text = Support_function.get_text_from_word(path_to_word)

    gr_api.green_api.sending.sendMessage(f'{phone_number}@c.us', f'{text}')

    return True
