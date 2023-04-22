import tkinter as tk
from tkinter import messagebox
import docx
import os
import tempfile
import shutil


def replace_text_in_word_document(doc_file, old_text1, new_text1, old_text2=None, new_text2=None,
 old_text3=None, new_text3=None, old_text4=None, new_text4=None, old_text5=None, new_text5=None, 
 old_text6=None, new_text6=None, old_text7=None, new_text7=None, old_text8=None, new_text8=None, 
 old_text9=None, new_text9=None, old_text10=None, new_text10=None, old_text11=None, new_text11=None, 
 old_text12=None, new_text12=None, old_text13=None, new_text13=None, old_text14=None, new_text14=None):
    try:
        # Открыть документ Word
        doc = docx.Document(doc_file)

        # Создать временный файл
        temp_file = tempfile.NamedTemporaryFile(delete=False)

        # Пройтись по всем параграфам в документе
        for para in doc.paragraphs:
            # Пройтись по всем объектам Run в параграфе
            for run in para.runs:
                # Заменить старый текст на новый, если он найден в объекте Run
                if old_text1 in run.text:
                    new_text = run.text.replace(old_text1, new_text1)
                    run.text = new_text
                # Заменить второй старый текст на новый, если он найден в объекте Run
                if old_text2 is not None and old_text2 in run.text:
                    new_text = run.text.replace(old_text2, new_text2)
                    run.text = new_text
                # Заменить третий старый текст на новый, если он найден в объекте Run
                if old_text3 is not None and old_text3 in run.text:
                    new_text = run.text.replace(old_text3, new_text3)
                    run.text = new_text
                # Заменить четвертый старый текст на новый, если он найден в объекте Run
                if old_text4 is not None and old_text4 in run.text:
                    new_text = run.text.replace(old_text4, new_text4)
                    run.text = new_text
                # Заменить пятый старый текст на новый, если он найден в объекте Run
                if old_text5 is not None and old_text5 in run.text:
                    new_text = run.text.replace(old_text5, new_text5)
                    run.text = new_text
                # Заменить шестой старый текст на новый, если он найден в объекте Run
                if old_text6 is not None and old_text6 in run.text:
                    new_text = run.text.replace(old_text6, new_text6)
                    run.text = new_text
                # Заменить седьмой старый текст на новый, если он найден в объекте Run
                if old_text7 is not None and old_text7 in run.text:
                    new_text = run.text.replace(old_text7, new_text7)
                    run.text = new_text
                # Заменить восьмой старый текст на новый, если он найден в объекте Run
                if old_text8 is not None and old_text8 in run.text:
                    new_text = run.text.replace(old_text8, new_text8)
                    run.text = new_text
                # Заменить девятый старый текст на новый, если он найден в объекте Run
                if old_text9 is not None and old_text9 in run.text:
                    new_text = run.text.replace(old_text9, new_text9)
                    run.text = new_text
                # Заменить десятый старый текст на новый, если он найден в объекте Run
                if old_text10 is not None and old_text10 in run.text:
                    new_text = run.text.replace(old_text10, new_text10)
                    run.text = new_text
                # Заменить одинадцатый старый текст на новый, если он найден в объекте Run
                if old_text11 is not None and old_text11 in run.text:
                    new_text = run.text.replace(old_text11, new_text11)
                    run.text = new_text
                # Заменить двенадцатый старый текст на новый, если он найден в объекте Run
                if old_text12 is not None and old_text12 in run.text:
                    new_text = run.text.replace(old_text12, new_text12)
                    run.text = new_text
                # Заменить тренадцатый старый текст на новый, если он найден в объекте Run
                if old_text13 is not None and old_text13 in run.text:
                    new_text = run.text.replace(old_text13, new_text13)
                    run.text = new_text
                # Заменить четырнадцатый старый текст на новый, если он найден в объекте Run
                if old_text12 is not None and old_text14 in run.text:
                    new_text = run.text.replace(old_text14, new_text14)
                    run.text = new_text

        # Сохранить изменения во временном файле
        doc.save(temp_file.name)

        # Открыть временный файл в Word
        os.startfile(temp_file.name)

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка при обработке документа: {e}")


def on_submit():
    try:
        # Получить данные из полей ввода
        employee_data = employee_data_entry.get()
        weapon_data = weapon_data_entry.get()
        passenger_data = passenger_data_entry.get()
        act_number = act_number_entry.get()
        day_number = day_number_entry.get()
        month_number = month_number_entry.get()
        year_number = year_number_entry.get()
        passport_data = passport_data_entry.get()
        flight_data = flight_data_entry.get()
        city_data = city_data_entry.get()
        bort_data = bort_data_entry.get()
        baggageweap_data = baggageweap_data_entry.get()
        ammo_data = ammo_data_entry.get()
        baggageammo_data = baggageammo_data_entry.get()

        # Заменить текст в документе Word
        replace_text_in_word_document('D:\\123.docx', 'SAB', employee_data, 'WEAP', weapon_data, 'NAM', passenger_data, 
            'ACT', act_number, 'DAY', day_number, 'MON', month_number, 'YEAR', year_number, 'PASS', passport_data, 'NUMB', flight_data, 
            'CYT', city_data, 'BOR', bort_data, 'BGV', baggageweap_data, 'AMMO', ammo_data, 'BAG', baggageammo_data)
    except Exception as e:
       messagebox.showerror("Ошибка", f"Произошла ошибка при обработке данных: {e}")

def on_exit():
    # Закрыть окно
    root.destroy()
    # Удалить временный файл
    os.remove(temp_file.name)

# Создать главное окно
root = tk.Tk()
root.geometry("1024x768")

# Добавить поле для ввода номера акта
act_number_label = tk.Label(root, text="Номер акта: ")
act_number_label.grid(row=1, column=0, padx=5, pady=5)
act_number_entry = tk.Entry(root, width=70)
act_number_entry.grid(row=1, column=1, padx=5, pady=5)
def show_help():
    messagebox.showinfo("Подсказка", "Введите номер акта. Не забудьте проверить номер на предыдущем акте.")
help_button = tk.Button(root, text="?", command=show_help)
help_button.grid(row=1, column=2, padx=5, pady=5)

# Добавить поле для ввода дня
day_number_label = tk.Label(root, text="День: ")
day_number_label.grid(row=2, column=0, padx=5, pady=5)
day_number_entry = tk.Entry(root, width=70)
day_number_entry.grid(row=2, column=1, padx=5, pady=5)
def show_help():
    messagebox.showinfo("Подсказка", "Введите текущий день.")
help_button = tk.Button(root, text="?", command=show_help)
help_button.grid(row=2, column=2, padx=5, pady=5)

# Добавить поле для ввода месяца
month_number_label = tk.Label(root, text="Месяц: ")
month_number_label.grid(row=3, column=0, padx=5, pady=5)
month_number_entry = tk.Entry(root, width=70)
month_number_entry.grid(row=3, column=1, padx=5, pady=5)
def show_help():
    messagebox.showinfo("Подсказка", "Введите Текущий месяц текстом а не цифрами.")
help_button = tk.Button(root, text="?", command=show_help)
help_button.grid(row=3, column=2, padx=5, pady=5)

# Добавить поле для ввода года
year_number_label = tk.Label(root, text="Год: ")
year_number_label.grid(row=4, column=0, padx=5, pady=5)
year_number_entry = tk.Entry(root, width=70)
year_number_entry.grid(row=4, column=1, padx=5, pady=5)
def show_help():
    messagebox.showinfo("Подсказка", "Введите текущий год полностью. Пример: 2023")
help_button = tk.Button(root, text="?", command=show_help)
help_button.grid(row=4, column=2, padx=5, pady=5)

#Добавить поле для ввода данных сотрудника
employee_data_label = tk.Label(root, text="Сотрудник САБ: ")
employee_data_label.grid(row=5, column=0, padx=5, pady=5)
employee_data_entry = tk.Entry(root, width=70)
employee_data_entry.grid(row=5, column=1, padx=5, pady=5)
def show_help():
    messagebox.showinfo("Подсказка", "Введите должность, фамилию, имя и отчество сотрудника, оформляющего акт.")
help_button = tk.Button(root, text="?", command=show_help)
help_button.grid(row=5, column=2, padx=5, pady=5)

# Добавить поле для ввода данных сотрудника
# employee_data_label = tk.Label(root, text="Данные сотрудника: ")
# employee_data_label.grid(row=0, column=0, padx=5, pady=5)

# определение списка для выпадающего меню
# employee_data_values = ['Старший инспектор Серов Андрей Владимирович', 'Jane Doe', 'Alex Johnson', 'Maria Garcia', 'Chris Lee']
# employee_data_value = tk.StringVar(root)
# employee_data_value.set(employee_data_values[0])

#Создать выпадающее меню для выбора данных сотрудника
#employee_data_dropdown = tk.OptionMenu(root, employee_data_value, *employee_data_values)
#employee_data_dropdown.grid(row=0, column=1, padx=5, pady=5)

#Добавить поле для ввода данных пассажира
passenger_data_label = tk.Label(root, text="Фамилия, Имя, Отчество Пассажира: ")
passenger_data_label.grid(row=6, column=0, padx=5, pady=5)
passenger_data_entry = tk.Entry(root, width=70)
passenger_data_entry.grid(row=6, column=1, padx=5, pady=5)

#Добавить поле для ввода паспортных данных пассажира
passport_data_label = tk.Label(root, text="Паспортные данные пассажира: ")
passport_data_label.grid(row=7, column=0, padx=5, pady=5)
passport_data_entry = tk.Entry(root, width=70)
passport_data_entry.grid(row=7, column=1, padx=5, pady=5)

#Добавить поле для ввода номера рейса
flight_data_label = tk.Label(root, text="Номер рейса: ")
flight_data_label.grid(row=8, column=0, padx=5, pady=5)
flight_data_entry = tk.Entry(root, width=70)
flight_data_entry.grid(row=8, column=1, padx=5, pady=5)

#Добавить поле для ввода аэропорта назначения
city_data_label = tk.Label(root, text="Аэропорт назначения: ")
city_data_label.grid(row=9, column=0, padx=5, pady=5)
city_data_entry = tk.Entry(root, width=70)
city_data_entry.grid(row=9, column=1, padx=5, pady=5)

#Добавить поле для ввода бортового номера воздушного судна
bort_data_label = tk.Label(root, text="Бортовой номер ВС: ")
bort_data_label.grid(row=10, column=0, padx=5, pady=5)
bort_data_entry = tk.Entry(root, width=70)
bort_data_entry.grid(row=10, column=1, padx=5, pady=5)

#Добавить поле для ввода данных об оружии
weapon_data_label = tk.Label(root, text="Тип оружия: ")
weapon_data_label.grid(row=11, column=0, padx=5, pady=5)
weapon_data_entry = tk.Entry(root, width=70)
weapon_data_entry.grid(row=11, column=1, padx=5, pady=5)
def show_help():
    messagebox.showinfo("Подсказка", "Введите тип оружия и его полные данные (тип, модель оружия,  его регистрационный номер).")
help_button = tk.Button(root, text="?", command=show_help)
help_button.grid(row=11, column=2, padx=5, pady=5)

#Добавить поле для ввода багажной бирки оружия
baggageweap_data_label = tk.Label(root, text="Номера багажных бирок оружия: ")
baggageweap_data_label.grid(row=12, column=0, padx=5, pady=5)
baggageweap_data_entry = tk.Entry(root, width=70)
baggageweap_data_entry.grid(row=12, column=1, padx=5, pady=5)

#Добавить поле для ввода количества боеприпасов
ammo_data_label = tk.Label(root, text="Количество боепирпасов: ")
ammo_data_label.grid(row=13, column=0, padx=5, pady=5)
ammo_data_entry = tk.Entry(root, width=70)
ammo_data_entry.grid(row=13, column=1, padx=5, pady=5)

#Добавить поле для ввода богажной бирки боеприпасов
baggageammo_data_label = tk.Label(root, text="Номера багажных бирок боеприпасов: ")
baggageammo_data_label.grid(row=14, column=0, padx=5, pady=5)
baggageammo_data_entry = tk.Entry(root, width=70)
baggageammo_data_entry.grid(row=14, column=1, padx=5, pady=5)

#Добавить кнопку для отправки данных
submit_button = tk.Button(root, text="Отправить", command=on_submit)
submit_button.grid(row=15, column=0, columnspan=2, pady=10, sticky='n')
submit_button.configure(bg='#98FB98')

#Добавить кнопку для выхода из программы
exit_button = tk.Button(root, text="Выход", command=on_exit)
exit_button.grid(row=16, column=1, padx=5, pady=5, sticky='se')
exit_button.configure(bg='#CD5C5C')

# Создать виджет Label с копирайтом
copyright_label = tk.Label(root, text="© 2023 El.Psy.Congroo")
copyright_label.grid(row=16, column=0, padx=5, pady=5, sticky='sw')
# Разместить его внизу окна, слева
# copyright_label.pack(side="left", padx=5, pady=5)
# Запустить главный цикл программы
root.mainloop()
