import time
import Mailing
import Formatting_numbers
import Support_function
import os
import webbrowser

from tkinter import messagebox, filedialog
from GreenAPI import gr_api
from Excel_work import excel_table
from GUI.Gui_classes import *
from ctypes import *


class MainWindow:

    def __init__(self):

        self.window = Window(title_text="Main", geometry=f"420x260+{windll.user32.GetSystemMetrics(0)//2}+{windll.user32.GetSystemMetrics(1)//2}")

        menu = MainMenu(parent=self.window)
        menu.open_file()

        self.window.config(menu=menu)

        self.frame_button = Frame(parent=self.window, pad_y=10, pad_x=10, row=0, col=0)
        self.frame_inform = Frame(parent=self.window, pad_y=10, pad_x=50, row=0, col=1)

        self.btn_mailing = Button(parent=self.frame_button, path_img="img/button_img/button_open_mailing.png", command=self.open_mailing)
        self.btn_format = Button(parent=self.frame_button, path_img="img/button_img/button_format.png", command=self.open_work_columns)
        self.btn_settings = Button(parent=self.frame_button, path_img="img/button_img/button_settings.png", command=self.open_settings)
        self.btn_exit = Button(parent=self.frame_button, path_img="img/button_img/button_exit.png", command=lambda: exit())

        self.btn_information = Button(parent=self.frame_inform, path_img="img/button_img/button_inform.png", command=self.open_information)

        if not gr_api.check_authorization():
            self.green_api_error()

        self.window.start()

    def open_settings(self):

        window_settings = Settings(self.window)
        self.window.withdraw()

    def green_api_error(self):

        result = messagebox.askyesno(title="Ошибка GreenAPI", message=f"Вы не авторизованы в GreenApi.\nОШИБКА:{gr_api.error}\n\nВвести данные сейчас?")
        if result:
            window_green_api = InputGreenApi(parent=self.window)

    def open_work_columns(self):
        self.window.withdraw()
        window_work_columns = WorkColumns(parent=self.window)

    def open_mailing(self):
        self.window.withdraw()
        window_mailing = WindowMailing(self.window)

    @staticmethod
    def open_information():
        message_info = f"Статус авторизации: {gr_api.status_green_api}\n" \
                        f"Id_instance: {gr_api.id_instance}\n" \
                        f"Api_token: {gr_api.api_token}\n" \
                        f"Текущий лист: {excel_table.sheet_main.title}\n"\
                        f"Выбранные колонки: {excel_table.selected_column_print}"

        messagebox.showinfo(title="GreenApi", message=message_info)

    @property
    def main_window(self):
        return self.window


class MainMenu(tk.Menu):

    def __init__(self, parent):
        super().__init__()

        self.parent = parent

        self.add_command(label="Выбрать файл", command=self.open_file)
        self.add_command(label="Открыть документацию", command=self.open_documentation)
        self.add_cascade(label=f"Выбранный файл: {excel_table.file_path}", state="disabled")

    def open_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

        if not file_path:
            print('Возникла ошибка! Возможно вы используете не тот файл!')

        excel_table.file_path = file_path
        excel_table.initialization_excel()

        if len(excel_table.sheets) > 1:
            messagebox.showinfo(title="Информация", message="В данном Excel файле найдено несколько листов. Выберите пожалуйста лист с которым будете работать")
            window_select_sheet = SelectSheet(self.parent)
        else:
            excel_table.sheet_main = excel_table.sheets[0]

        self.entryconfigure(3, label=os.path.basename(excel_table.file_path))

    @staticmethod
    def open_documentation():
        webbrowser.open("documentation\\Документация.pdf")


class Settings:

    def __init__(self, parent):
        self.window = ChildWindow(title="Настройки", parent=parent, geometry=f"275x250+{windll.user32.GetSystemMetrics(0)//2}+{windll.user32.GetSystemMetrics(1)//2}")
        self.frame_buttons = Frame(parent=self.window, pad_x=10, pad_y=10)
        btn_change_list = Button(parent=self.frame_buttons, path_img="img/button_img/button_change_sheet.png", command=self.open_change_sheet)
        btn_select_column = Button(parent=self.frame_buttons, path_img="img/button_img/button_select_columns.png", command=self.open_select_columns)
        btn_update_green_api = Button(parent=self.frame_buttons, path_img="img/button_img/button_update_greenapi.png", command=self.open_update_green_api)
        btn_back = Button(parent=self.frame_buttons, path_img="img/button_img/button_settings_back.png", command=lambda: self.window.back_to_main_window(parent=parent))

    def open_change_sheet(self):
        self.window.withdraw()
        window_change_list = SelectSheet(parent=self.window)

    def open_select_columns(self):
        self.window.withdraw()
        window_select_columns = WorkColumns(self.window)

    def open_update_green_api(self):
        self.window.withdraw()
        window_update_green_api = InputGreenApi(self.window)


class WorkColumns:

    def __init__(self, parent):

        if not excel_table.file_path:
            MainMenu.open_file()

        self.window = ChildWindow(title="Ввод колонок", parent=parent, geometry=f"430x210+{windll.user32.GetSystemMetrics(0)//2}+{windll.user32.GetSystemMetrics(1)//2}")

        self.frame_label = Frame(parent=self.window, col=0, row=1, sticky=tk.E)
        self.frame_input = Frame(parent=self.window, col=1, row=1)
        self.frame_buttons = Frame(parent=self.window, col=1, row=2)

        self.label_columns_img = Label(parent=self.frame_label, path_img="img/label_img/label_columns.png")
        self.label_selected_columns_img = Label(parent=self.frame_label, path_img="img/label_img/label_selected_columns.png")

        self.entry_columns = Entry(parent=self.frame_input)
        self.label_selected_columns = Label(parent=self.frame_input, pad_y=5, text=excel_table.selected_column_print)

        self.btn_confirm = Button(parent=self.frame_buttons, path_img="img/button_img/button_confirm.png", command=self.confirm_column)
        self.btn_format = Button(parent=self.frame_buttons, path_img="img/button_img/button_formating.png", command=self.format_numbers)
        self.btn_cancel = Button(parent=self.frame_buttons, path_img="img/button_img/button_close.png", command=lambda: self.window.back_to_main_window(parent=parent))

        self.window.grab_set()

    def confirm_column(self):

        columns = self.entry_columns.get()
        excel_table.select_column(columns=columns, clear=True)
        self.label_selected_columns.config(text=excel_table.selected_column_print)
        print(excel_table.selected_column_print)

    def format_numbers(self):

        if self.entry_columns.get():
            self.confirm_column()

        if not excel_table.selected_column:
            messagebox.showerror(title="Ошибка", message="Вы не выбрали колонки с номерами. Введите колонки и "
                                                         "попробуйте еще раз")
            return

        Formatting_numbers.update_number()

        messagebox.showinfo(title="Успех", message="Все номера отформатированы!")


class MailingProcess:

    def __init__(self, parent):
        self.window = ChildWindow(parent=parent, title="Рассылка", geometry=f"700x500+{windll.user32.GetSystemMetrics(0)//2}+{windll.user32.GetSystemMetrics(1)//2}")

        self.progress_bar = ProgressBar(self.window)

        self.max_row = excel_table.get_sheet_row()

        self.progress_label = Label(parent=self.window, text=f"0 из {self.max_row}", anchor=tk.CENTER)
        self.progress_label.config(font=("Arial", 10))

        self.btn_close = Button(parent=self.window, path_img="img/button_img/button_close.png", command=lambda: self.window.back_to_main_window(parent=parent))
        self.btn_close.config(state="disabled")

    def update_value(self, step, current_value):
        self.progress_bar['value'] += step
        self.progress_label.config(text=f"{current_value} из {self.max_row}")
        self.window.update()
        time.sleep(1)


class WindowMailing(WorkColumns):

    def __init__(self, parent):

        super(WindowMailing, self).__init__(parent)
        self.window.geometry(f"1150x500+{windll.user32.GetSystemMetrics(0)//2}+{windll.user32.GetSystemMetrics(1)//2}")

        self.btn_mailing = Button(parent=self.frame_buttons, path_img="img/button_img/button_mailing.png", command=self.check)
        self.send_message = Text(parent=self.window, text=f"ОТПРАВЛЯЕМОЕ СООБЩЕНИЕ:\n{Support_function.get_text_from_word('content.docx')}", x=450, y=5)
        self.btn_format.destroy()

        self.keys_list = Support_function.get_keys_from_word()
        self.entries = {}

        if self.keys_list:
            for i in self.keys_list:
                label_key = Label(parent=self.frame_label, text=i)
                entry_key = Entry(parent=self.frame_input)
                self.entries[i] = entry_key

    def check(self):

        empty_values = []
        values_outside_index = []

        # Проверяем авторизацию
        if not gr_api.check_authorization():
            result = messagebox.askyesno(title="Ошибка GreenAPI", message="Вы не авторизованы в GreenApi. Ввести данные сейчас?")
            if result:
                window_green_api = InputGreenApi(parent=self.window)
                return

        # Проверяем колонки с номерами
        if not self.entry_columns.get() and not excel_table.selected_column:
            messagebox.showerror(title="Error", message="Не выбраны колонки с номерами!")
            return

        elif self.entry_columns.get():
            self.confirm_column()

        # Проверяем заполненность полей ключей
        for key, value in self.entries.items():
            entry_value = value.get()

            if not entry_value:
                empty_values.append(key)
                continue

            if not (1 <= int(entry_value) <= excel_table.max_column):
                values_outside_index.append(key)
                continue

        if empty_values or values_outside_index:
            empty_values = "Следующие ключи не заполнены:\n" + "\n".join(empty_values)
            values_outside_index = "\nСледующие ключи находятся вне границ массива:\n" + "\n".join(values_outside_index)
            messagebox.showerror(title="Error", message=empty_values + values_outside_index)
            return

        self.start_mailing()

    def start_mailing(self):

        context_col = {}

        for key, value in self.entries.items():
            context_col[key] = int(value.get()) - 1

        progress_bar = MailingProcess(parent=self.window)

        self.window.withdraw()

        successful_attempt, failed_attempt, error = Mailing.main_mailing(progress_bar, context_col)  # Запускаем функцию рассылки и получаем количество удачных/неудачных попыток

        progress_bar.btn_close.config(state="normal")

        if error:
            messagebox.showerror(title="Неудача", message=f"Возникла ошибка {error}")
        else:
            messagebox.showinfo(title="Успех!", message=f"Рассылка успешно завершена!\nОтправлено: {successful_attempt}\nНе отправлено: {failed_attempt}")


class InputGreenApi:

    def __init__(self, parent):
        self.window = ChildWindow(title="Ввод данных GreenApi", parent=parent, geometry=f"340x165+{windll.user32.GetSystemMetrics(0)//2}+{windll.user32.GetSystemMetrics(1)//2}")

        self.frame_label = Frame(parent=self.window, col=0, row=0)
        self.frame_input = Frame(parent=self.window, col=1, row=0)
        self.frame_buttons = Frame(parent=self.window, col=1, row=1)

        self.label_id_instance_img = Label(parent=self.frame_label, path_img="img/label_img/label_id-instance.png")
        self.label_api_token_img = Label(parent=self.frame_label, path_img="img/label_img/label_api-token.png")

        self.entry_id_instance = Entry(parent=self.frame_input)
        self.entry_api_token = Entry(parent=self.frame_input)

        self.btn_confirm = Button(parent=self.frame_buttons, path_img="img/button_img/button_confirm.png", command=self.check_green_api)
        self.btn_cancel = Button(parent=self.frame_buttons, path_img="img/button_img/button_close.png", command=lambda: self.window.back_to_main_window(parent=parent))

        self.window.grab_set()

    def check_green_api(self):

        id_instance = self.entry_id_instance.get().replace(" ", "")
        api_token = self.entry_api_token.get().replace(" ", "")

        if len(id_instance) != 10 or not id_instance.isdigit():
            messagebox.showerror(title="Id Instance", message="Ошибка ввода Id Instance.\nId instance должен состоять из 10 цифр.")
            return

        if len(api_token) != 50:
            messagebox.showerror(title="Api token", message="Ошибка ввода Api token.\nApi token должен состоять из 50 символов.")
            return

        if not gr_api.check_authorization(id_instance=id_instance, api_token=api_token):
            messagebox.showerror(title="Ошибка авторизации", message=gr_api.status_green_api)
            return

        gr_api.write_green_api_data(id_instance=id_instance, api_token=api_token)

        messagebox.showinfo(title="Данные записаны!", message="Данные успешно записаны. Авторизация пройдена")


class SelectSheet:

    def __init__(self, parent):

        self.window = ChildWindow(title="Выбор листа", parent=parent, geometry=f"250x180+{windll.user32.GetSystemMetrics(0)//2}+{windll.user32.GetSystemMetrics(1)//2}")
        parent.withdraw()

        self.label_text = Label(parent=self.window, text="Листы", pad_x=5, anchor=tk.W)

        self.radio_list = []
        self.value_variable = tk.StringVar()
        self.value_variable.set(excel_table.sheets[0])

        for i in excel_table.sheets:
            self.radio = RadioButton(parent=self.window, text=i, value=i, value_variable= self.value_variable, padx=5)
            self.radio_list.append(self.radio)

        self.button_confirm = Button(parent=self.window, path_img="img/button_img/button_confirm.png", command=lambda: self.set_main_list(parent=parent), anchor=tk.W, pady=5)

    def set_main_list(self, parent):
        excel_table.sheet_main = self.value_variable.get()
        self.window.back_to_main_window(parent=parent)


main = MainWindow()
