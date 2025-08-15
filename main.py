import os
import customtkinter as ctk
from datetime import datetime, timedelta
import locale
from tkinter import filedialog
import sys

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

from docx2pdf import convert

import os
import sys

def resource_path(relative_path):
    """ Получает абсолютный путь к ресурсу """
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# Используйте так:
TEMPLATES_DIR = resource_path("templates")

def set_russian_locale():
    try:
        # Для Linux/macOS
        locale.setlocale(locale.LC_TIME, 'ru_RU.UTF-8')
    except locale.Error:
        try:
            # Для Windows
            locale.setlocale(locale.LC_ALL, 'rus')
        except locale.Error:
            print("Не удалось установить русскую локаль. Месяц будет на английском.")

# Устанавливаем локаль
set_russian_locale()

# Настройка внешнего вида
ctk.set_appearance_mode("dark")  # Режим: "light", "dark", "system"
ctk.set_default_color_theme("blue")  # Темы: "blue", "green", "dark-blue"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Генератор нормированных заданий")
        self.geometry("1000x700")  # Увеличенный размер
        
        # Главный контейнер для страниц
        self.container = ctk.CTkFrame(self)
        self.container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Явное создание страниц
        self.pages = {
            "MainPage": MainPage(self.container, self),
            "Novoros": Novoros(self.container, self),
            'Astrakhan': Astrakhan(self.container, self),
            'Arhangelsk': Arhangelsk(self.container, self),
            'Bor': Bor(self.container, self),
            'Ekaterinburg': Ekaterinburg(self.container, self),
            'Habarovsk': Habarovsk(self.container, self),
            'Kaliningrad': Kaliningrad(self.container, self),
            'NN': NN(self.container, self),
            'Murmansk': Murmansk(self.container, self),
            'Nahodka': Nahodka(self.container, self),
            'PK': PK(self.container, self),
            'Rostov': Rostov(self.container, self),
            'Samara': Samara(self.container, self),
            'Sevastopol': Sevastopol(self.container, self),
            'Spb': Spb(self.container, self),
            'US': US(self.container, self),
            'Vladivostok': Vladivostok(self.container, self)
        }
        
        # Размещение всех страниц
        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")
        
        # Конфигурация grid для контейнера
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)
        
        # Показываем главную страницу
        self.show_page("MainPage")
    
    def show_page(self, page_name):
        """Показать выбранную страницу"""
        page = self.pages.get(page_name)
        if page:
            page.tkraise()
            if hasattr(page, "on_show"):
                page.on_show()

class MultiSelectDropdown(ctk.CTkFrame):
    def __init__(self, master, values, **kwargs):
        super().__init__(master, **kwargs)
        self.values = values
        self.selected = []
        self.check_vars = {}
        
        self.toggle_btn = ctk.CTkButton(
            self,
            text="Выбрать сотрудников ▼",
            command=self.toggle_dropdown,
            anchor="w"
        )
        self.toggle_btn.pack(fill="x")
        
        self.dropdown_frame = ctk.CTkFrame(self)
        self.dropdown_shown = False
        
        for value in values:
            var = ctk.BooleanVar(value=False)
            self.check_vars[value] = var
            cb = ctk.CTkCheckBox(
                self.dropdown_frame,
                text=value,
                variable=var,
                command=lambda v=value: self.update_selection(v)
            )
            cb.pack(anchor="w", pady=2)

    def toggle_dropdown(self):
        if self.dropdown_shown:
            self.dropdown_frame.pack_forget()
            self.toggle_btn.configure(text="Выбрать сотрудников ▼")
        else:
            self.dropdown_frame.pack(fill="x", pady=5)
            self.toggle_btn.configure(text="Скрыть список ▲")
        self.dropdown_shown = not self.dropdown_shown

    def update_selection(self, value):
        if self.check_vars[value].get():
            if value not in self.selected:
                self.selected.append(value)
        else:
            if value in self.selected:
                self.selected.remove(value)
        self.update_button_text()

    def update_button_text(self):
        if len(self.selected) == 0:
            self.toggle_btn.configure(text="Выбрать сотрудников ▼")
        else:
            self.toggle_btn.configure(text=f"Выбрано: {len(self.selected)} ▼")

    def get_selected(self):
        if len(self.selected) == 1:
            return self.selected[0]
        elif len(self.selected) == 2:
            return "\n" + '\n'.join(self.selected) + "\n"
        else:
            return "\n".join(self.selected)

class BasePage(ctk.CTkFrame):
    """Базовый класс для всех страниц"""
    def __init__(self, parent, controller, worker_list=[], template='', name='', east=False, template_fizo_path = ''):
        super().__init__(parent)
        self.controller = controller
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1) 
        self.font = ('TimesNewRoman', 13)
        self.template = template
        self.template_fizo_path = template_fizo_path
        self.name = name
        self.east = east
        self.results = {}
        


        self.annotation_frame = ctk.CTkFrame(master=self)
        self.annotation_frame.grid(row=0, column=0, padx=20, pady=10, sticky="nsew")

        ctk.CTkLabel(self.annotation_frame, text='Выберите необходимые параметры', font=('TimesNewRoman', 20)).grid(row=1, column=0, columnspan=4, sticky='nsew')

        self.info_frame = ctk.CTkFrame(master=self)
        self.info_frame.grid(row=1, column=0, padx=20, pady=10, sticky='nsew')

        # Номер заявки
        ctk.CTkLabel(self.info_frame, text='Номер заявки:', font=self.font).grid(row=0, column=0, padx=5, pady=5, sticky='w')

        self.num_var = ctk.StringVar()
        self.num_entry = ctk.CTkEntry(self.info_frame, width=300, textvariable=self.num_var, placeholder_text='Номер заявки по нормированному заданию')
        self.num_entry.grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky='ew')

        # Дата формирования
        ctk.CTkLabel(self.info_frame, text='Дата формирования:', font=self.font).grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.ISSUE_DATE = ctk.StringVar()
        self.issue_date_entry = ctk.CTkEntry(self.info_frame, width=300, textvariable=self.ISSUE_DATE)
        self.issue_date_entry.insert(0, datetime.now().strftime('«%d» %B %Y г.'))
        self.issue_date_entry.grid(row=1, column=1, columnspan=3, padx=5, pady=5, sticky='ew')

        # Дата проверки
        ctk.CTkLabel(self.info_frame, text='Дата проведения проверки:', font=self.font).grid(row=2, column=0, padx=5, pady=5, sticky='w')

        self.WORK_DATE = ctk.StringVar()
        self.work_date_entry = ctk.CTkEntry(self.info_frame, width=300, textvariable=self.WORK_DATE)
        self.work_date_entry.insert(0, self.work_date(self.east).strftime('%d %B %Y г.'))
        self.work_date_entry.grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky='ew')

        # Время проверки
        ctk.CTkLabel(self.info_frame, text='Время проведения проверки:', font=self.font).grid(row=3, column=0, padx=5, pady=5, sticky='w')
        
        self.TIME_START = ctk.StringVar()
        self.entry_start_time = ctk.CTkEntry(self.info_frame, width=100, textvariable=self.TIME_START)
        self.entry_start_time.insert(0, '10:00')
        self.entry_start_time.grid(row=3, column=1, padx=5, pady=5, sticky='w')
        
        ctk.CTkLabel(self.info_frame, text='-').grid(row=3, column=2, padx=5, pady=5)
        
        self.TIME_END = ctk.StringVar()
        self.entry_end_time = ctk.CTkEntry(self.info_frame, width=100,textvariable=self.TIME_END)
        self.entry_end_time.insert(0, '16:00')
        self.entry_end_time.grid(row=3, column=3, padx=5, pady=5, sticky='w')

        #Типы проверок и УИНы
        self.pfo_var = ctk.BooleanVar(value=False)
        self.fizo_var = ctk.BooleanVar(value=False)
        self.zun_var = ctk.BooleanVar(value=False)

        self.pfo_value = ctk.StringVar()
        self.fizo_value = ctk.StringVar()
        self.zun_value = ctk.StringVar()

        self.cb_pfo = ctk.CTkCheckBox(self.info_frame, text='ПФО', variable=self.pfo_var, command=lambda: self.toggle_entry('pfo'))
        self.cb_pfo.grid(row=4, column=0, padx=5, pady=5, sticky='w')
        
        self.cb_fizo = ctk.CTkCheckBox(self.info_frame, text='ФИЗО', variable=self.fizo_var, command=lambda: self.toggle_entry('fizo'))
        self.cb_fizo.grid(row=5, column=0, padx=5, pady=5, sticky='w')
        

        self.cb_zun = ctk.CTkCheckBox(self.info_frame, text='ЗУН', variable=self.zun_var, command=lambda: self.toggle_entry('zun'))
        self.cb_zun.grid(row=6, column=0, padx=5, pady=5, sticky='w')

        self.entry_pfo = ctk.CTkEntry(self.info_frame, textvariable=self.pfo_value)
        self.entry_fizo = ctk.CTkEntry(self.info_frame, textvariable=self.fizo_value)
        self.entry_zun = ctk.CTkEntry(self.info_frame, textvariable=self.zun_value)

        

        # Выбор сотрудников
        ctk.CTkLabel(self.info_frame, text='Выберите сотрудников:', font=self.font).grid(row=7, column=0, padx=5, pady=5, sticky='w')
        
        self.worker_list = worker_list
        self.workers_dropdown = MultiSelectDropdown(self.info_frame, self.worker_list)
        self.workers_dropdown.grid(row=7, column=1, columnspan=4, padx=5, pady=5, sticky='ew')

        #Выбор должностного лица
        ctk.CTkLabel(self.info_frame, text='Выберите должностное лицо:', font=self.font).grid(row=8, column=0, pady=5, sticky='w')

        self.chief_var = ctk.StringVar(value='')
        self.cb_bezzubcev = ctk.CTkCheckBox(self.info_frame, text='Беззубцев А.А.', font=self.font, variable=self.chief_var, onvalue='Беззубцев', offvalue='')
        self.cb_bezzubcev.grid(row=8, column=1, sticky='w')

        self.cb_aliabiev = ctk.CTkCheckBox(self.info_frame, text='Алябьев А.Б.', font=self.font, variable=self.chief_var, onvalue='Алябьев', offvalue='')
        self.cb_aliabiev.grid(row=8, column=2, sticky='w')

        #Исполнитель
        ctk.CTkLabel(self.info_frame, text='Укажите исполнителя:', font=self.font).grid(row=9, column=0, sticky='w')
        self.issuer = ctk.StringVar(value='')

        self.cb_marusya = ctk.CTkCheckBox(self.info_frame, text='Разгуляева М.А.', font=self.font, variable=self.issuer, onvalue='Маруся', offvalue='')
        self.cb_marusya.grid(row=9, column=1, sticky='w')
        self.cb_vantuz = ctk.CTkCheckBox(self.info_frame, text='Воротилов И.И.', font=self.font, variable=self.issuer, onvalue='Вантуз', offvalue='')
        self.cb_vantuz.grid(row=9, column=2, sticky='w')
        self.cb_creator = ctk.CTkCheckBox(self.info_frame, text='Желудков А.В.', font=self.font, variable=self.issuer, onvalue='Создатель!!!', offvalue='')
        self.cb_creator.grid(row=9, column=3, sticky='w')
        

        # Настройка колонок
        for i in range(4):
            self.info_frame.grid_columnconfigure(i, weight=1 if i > 0 else 0)

        #Сохранение
        self.btn_save = ctk.CTkButton(self, text='Сформировать документ', font=self.font, command=self.save_data_and_form_doc)
        self.btn_save.grid(row=2, column=0, sticky='nsew')

    #Функция отображения поля ввода по чекбоксам ПФО, ФИЗО, ЗУН
    def toggle_entry(self, checkbox):
        if checkbox == 'pfo':
            if self.pfo_var.get():
                self.entry_pfo.grid(row=4, column=1, columnspan=3, padx=5, pady=5, sticky='ew')
            else:
                self.entry_pfo.grid_forget()
        elif checkbox == 'fizo':
            if self.fizo_var.get() and self.template_fizo_path:
                self.entry_fizo.grid(row=5, column=1, padx=5, pady=5, sticky='ew')
                self.number_fiz = ctk.StringVar()
                ctk.CTkLabel(self.info_frame, text='Номер заявки для ФИЗО', font=self.font).grid(row=5, column=2, sticky='ew')
                ctk.CTkEntry(self.info_frame, textvariable=self.number_fiz).grid(row=5, column=3, sticky='ew')
            elif self.fizo_var.get():
                self.entry_fizo.grid(row=5, column=1, columnspan=3, padx=5, pady=5, sticky='ew')
            else:
                self.entry_fizo.grid_forget()
        elif checkbox == 'zun':
            if self.zun_var.get():
                self.entry_zun.grid(row=6, column=1, columnspan=3, padx=5, pady=5, sticky='ew')
            else:
                self.entry_zun.grid_forget()
    
    #Экзамены без физры первая страница
    def get_exams(self):
        res = []

        if self.pfo_var.get():
            res.append('соответствия личностных (психофизиологических) качеств')
        if self.zun_var.get():
            res.append('знаний, умений и навыков')
        
        if len(res) == 1:
            return res[0]
        elif len(res) == 2:
            return "\n" + '\n'.join(res) + "\n"
        else:
            return "\n".join(res)


    #Фукнция сбора УИНОВ
    def get_uins(self):
        total_uins = set()
        exams = {}
        type_of_exams = []
        if self.pfo_var.get():
            exams['личностных (психофизиологических) качеств'] = self.formate_uins(self.pfo_value.get())
            type_of_exams.append('личностных (психофизиологических) качеств')
        if self.fizo_var.get():
            exams['уровня физической подготовки'] = self.formate_uins(self.fizo_value.get())
            type_of_exams.append('уровня физической подготовки')
        if self.zun_var.get():
            exams['знаний, умений и навыков'] = self.formate_uins(self.zun_value.get())
            type_of_exams.append('знаний, умений и навыков')
        
        result_string = ''
        exams_string = ''
        for index, (key,value) in enumerate(exams.items(), 1):
            for item in value:
                total_uins.add(item)
            if len(exams) == 1:
                result_string += f'{key} аттестуемых лиц (УИН: {"; ".join(value)}).'
                break
            if index == len(exams):
                result_string += f'а также {key} аттестуемых лиц (УИН: {"; ".join(value)}).'
            else:
                result_string += f'{key} аттестуемых лиц (УИН: {"; ".join(value)}), '

        if len(type_of_exams) == 1:
            exams_string = type_of_exams[0]
        elif len(type_of_exams) == 2:
            exams_string = f"{', а также '.join(type_of_exams)}"
        else:
            exams_string = f'{type_of_exams[0]}, {type_of_exams[1]}, а также {type_of_exams[2]}'
                
        total_uins = sorted(list(total_uins))
        return total_uins, result_string, exams_string


        
    #Функция сбора и формирования уинов
    def formate_uins(self, string):
        cur_year = datetime.now().year - 2000
        start_uin = f'11{cur_year}20020'
        string = string.split(', ')
        result = []
        for item in string:
            if '-' in item:
                start, end = [int(num) for num in item.split('-')]
                for i in range(start, end + 1):
                    result.append(start_uin + str(i))
            else:
                result.append(start_uin + item)
        return result
    
    #Чекбоксы подписанты
    def get_result_chief(self):
        selected = []

        if self.chief_var.get() == 'Беззубцев':
            selected.append('Начальник отдела проверок сил ОТБ')
            selected.append('А.А. Беззубцев')

        if self.chief_var.get() == 'Алябьев':
            selected.append('Заместитель начальника Службы')
            selected.append('А.Б. Алябьев')
        
        return selected
    
    #Функция чекбоксы исполнители
    def get_result_issuer(self):
        if self.issuer.get() == 'Вантуз':
            return 'Воротилов Иван Иванович', 'oa13@msecurity.ru'
        elif self.issuer.get() == 'Маруся':
            return 'Разгуляева Мария Андреевна', 'oa6@msecurity.ru'
        elif self.issuer.get() == 'Создатель!!!':
            return 'Желудков Андрей Викторович', 'oa15@msecurity.ru'
        return '', ''


    #Функция определения следующей даты
    def work_date(self, east=False):
        is_friday = datetime.now().isoweekday() == 5
        if is_friday:
            return datetime.now() + timedelta(days=3)
        elif is_friday and east:
            return datetime.now() + timedelta(days=4)
        elif east and datetime.now().isoweekday() == 4:
            return datetime.now() + timedelta(days=4)
        elif east:
            return datetime.now() + timedelta(days=2)
        return datetime.now() + timedelta(days=1)
    
    #Функция расчета затраченого времени
    def est_time(self):
        start = int(self.TIME_START.get().split(':')[0])
        end = int(self.TIME_END.get().split(':')[0])
        res = end - start
        if self.entry_fizo and self.template_fizo_path:
            res += 1
        return f'{res} часов' if res >= 5 else f'{res} часа'

    #Функция замены меток в документах
    def formate_docx(self, replacements_dict, template):
        edited_doc = Document(template)
        
        # Обработка обычных параграфов
        for paragraph in edited_doc.paragraphs:
            for key, value in replacements_dict.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, value)   
                    if key in ['{{NUMBER}}', '{{DATE_OF_ISSUE}}', '{{NUMBER_FIZO}}']:
                        for run in paragraph.runs:
                            run.font.bold = True
                    for run in paragraph.runs:
                        run.font.name = "Times New Roman"
                        run.font.size = Pt(13)
                                        
        # Обработка таблиц
        for table in edited_doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for key, value in replacements_dict.items():
                            if key in paragraph.text:
                                paragraph.text = paragraph.text.replace(key, value)
                                # Выравнивание по центру для замененного текста
                                if key not in ('{{JOB_TITLE}}', '{{FULLNAME}}'):
                                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                
                        for run in paragraph.runs:
                            run.font.name = "Times New Roman"
                            run.font.size = Pt(13)
                    
                    # Выравнивание содержимого ячейки по вертикали по центру
                    # cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    
        # Обработка колонтитулов
        for sect in edited_doc.sections:
            footer = sect.footer
            for paragraph in footer.paragraphs:
                for key, value in replacements_dict.items():
                    if key in paragraph.text:
                        paragraph.text = paragraph.text.replace(key, value)
                    for run in paragraph.runs:
                        run.font.name = "Times New Roman"
                        run.font.size = Pt(9)
        
        return edited_doc
            
        

    #Функция для сохранения результата и формирования документа
    def save_data_and_form_doc(self):
        time = f'с {self.TIME_START.get()} час. до {self.TIME_END.get()} час.'
        total_uins, exams_str, types_str = self.get_uins()
        job_title, fullname = self.get_result_chief()
        issuer, email = self.get_result_issuer()
        self.results = {
            '{{NUMBER}}': self.num_var.get(),
            '{{NUMBER_FIZO}}': self.number_fiz.get() if hasattr(self, 'number_fiz') and self.number_fiz.get() else '',
            '{{DATE_OF_ISSUE}}': self.ISSUE_DATE.get(),
            '{{DATE_TO_WORK}}': self.WORK_DATE.get(),
            '{{TYPE_OF_EXAMS}}': types_str,
            '{{IDENTIFICATORS}}': '; '.join(total_uins) + '.',
            '{{EXAMS_P2}}': exams_str,
            '{{TIME}}': time,
            '{{EST_TIME}}': self.est_time(),
            '{{PEOPLE}}': self.workers_dropdown.get_selected(),
            '{{JOB_TITLE}}': job_title,
            '{{FULLNAME}}': fullname,
            '{{ISSUER}}': issuer,
            '{{EMAIL}}': email,
            '{{EXAMS}}': self.get_exams(),
            '{{FIZO_UIN}}':'; '.join(self.formate_uins(self.fizo_value.get())) + '.'
        }
        edited_doc = self.formate_docx(self.results, self.template)
        default_filename = f'Нормированное задание на {self.work_date(self.east).strftime("%d.%m")} {self.name}'
        file_path = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        initialfile=default_filename,
        title="Сохранить документ как"
    )
        
        
        if not file_path:
            return
        
        edited_doc.save(file_path)
        
        pdf_path = os.path.splitext(file_path)[0] + '.pdf'
        convert(file_path, pdf_path)

        if self.template_fizo_path:
            fizo_doc = self.formate_docx(self.results, self.template_fizo_path)
            default_filename = f'Заявка на {self.work_date(self.east).strftime("%d.%m")} {self.name}'
            file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
            initialfile=default_filename,
            title="Сохранить документ как")

            if not file_path:
                return
            
            fizo_doc.save(file_path)

            pdf_path = os.path.splitext(file_path)[0] + '.pdf'
            convert(file_path, pdf_path)
        
        
        

        popup = ctk.CTkToplevel(self.info_frame)
        popup.title("Уведомление")
        popup.geometry("300x100")
        popup.resizable(False, False)
        
        
        ctk.CTkLabel(popup, text="Документ сформирован!").pack(pady=20)
        
        
        ctk.CTkButton(popup, text="OK", command=popup.destroy).pack(pady=5)

class MainPage(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        # Заголовок
        self.label = ctk.CTkLabel(self, text="Главная страница", font=("TimesNewRoman", 24))
        self.label.grid(row=0, column=0, columnspan=3, pady=20, padx=20)
        
        # Кнопки для перехода на другие страницы
        buttons = [
            ("Новороссийск", "Novoros"),
            ('Астрахань', 'Astrakhan'),
            ('Архангельск', 'Arhangelsk'),
            ('Бор', 'Bor'),
            ('Екатеринбург', 'Ekaterinburg'),
            ('Хабаровск', 'Habarovsk'),
            ('Калининград', 'Kaliningrad'),
            ('Нижний Новгород', 'NN'),
            ('Мурманск', 'Murmansk'),
            ('Находка', 'Nahodka'),
            ('Петропавловск-Камчатский', 'PK'),
            ('Ростов-на-Дону', 'Rostov'),
            ('Самара', 'Samara'),
            ('Севастополь', 'Sevastopol'),
            ('Санкт-Петербург', 'Spb'),
            ('Южно-Сахалинск', 'US'),
            ('Владивосток', 'Vladivostok')
        ]
        
        
        for i, (text, page) in enumerate(buttons):
            row = i // 3 + 1  
            col = i % 3
            btn = ctk.CTkButton(
                self, 
                text=text,
                command=lambda p=page: controller.show_page(p),
                height=40,
                width=200
            )
            btn.grid(row=row, column=col, pady=10, padx=10, sticky='nsew')

class Novoros(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Novoros.docx')
        super().__init__(parent, controller, 
                        ['Лагутин Ростислав Борисович', 'Соловьев Сергей Викторович', 'Соловьев Артем Андреевич'], 
                        template_path, 'Новороссийск')

        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Vladivostok(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Vladivostok.docx')
        template_fizo_path = os.path.join(TEMPLATES_DIR, 'Vladivostok_FIZO.docx')
        super().__init__(parent, controller, 
                        ['Проценко Игорь Владимирович', 'Лабонин Илья Валентинович', 'Родионова Елена Ивановна'], 
                        template_path, 'Владивосток', east=True, template_fizo_path=template_fizo_path)

        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class US(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'US.docx')
        template_fizo_path = os.path.join(TEMPLATES_DIR, 'US_FIZO.docx')
        super().__init__(parent, controller, 
                        ['Мельчарик Владимир Юрьевич', 'Шадрин Владимир Валерьевич'], 
                        template_path, 'Сахалин', east=True, template_fizo_path=template_fizo_path)

        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Spb(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Spb.docx')
        super().__init__(parent, controller, 
                        ['Савельев Дмитрий Генрихович', 'Семенова Мария Сергеевна'], 
                        template_path, 'Спб')

        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Sevastopol(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Sevastopol.docx')
        super().__init__(parent, controller, 
                        ['Кошелев Тимур Сергеевич', 'Кошелева Джеваире Айдеровна'], 
                        template_path, 'Севастополь')

        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Samara(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Samara.docx')
        super().__init__(parent, controller, 
                        ['Башмакова Виктория Владимировна'], 
                        template_path, 'Самара')

        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Rostov(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Rostov.docx')
        super().__init__(parent, controller, 
                        ['Евстратова Наталья Павловна', 'Язьков Анатолий Сергеевич'], 
                        template_path, 'Ростов')

        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class PK(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'PK.docx')
        template_fizo_path = os.path.join(TEMPLATES_DIR, 'PK_FIZO.docx')
        super().__init__(parent, controller, 
                        ['Каушанов Александр Викторович', 'Еремеев Геннадий Гайсович', 'Зотов Станислав Викторович'], 
                        template_path, 'Петропавловск-Камчатский', east=True, template_fizo_path=template_fizo_path)

        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Nahodka(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Nahodka.docx')
        super().__init__(parent, controller, 
                        ['Равнянский Константин Витальевич', 'Зубакин Эдуард Николаевич'], 
                        template_path, 'Находка', east=True)

        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Murmansk(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Murmansk.docx')
        template_fizo_path = os.path.join(TEMPLATES_DIR, 'Murmansk_FIZO.docx')
        super().__init__(parent, controller, 
                        ['Савельева Карина Дмитриевна', 'Пахомов Андрей Юрьевич'], 
                        template_path, 'Мурманск', template_fizo_path=template_fizo_path)

        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class NN(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'NN.docx')
        super().__init__(parent, controller, 
                        ['Еркович Наталья Евгеньевна'], 
                        template_path, 'Нижний Новгород')

        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Astrakhan(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Astrakhan.docx')
        template_fizo_path = os.path.join(TEMPLATES_DIR, 'Astrakhan_FIZO.docx')
        super().__init__(parent, controller, 
                        ['Шматова Раиса Анатольевна', 'Сивоконь Светлана Викторовна'], 
                        template_path, 'Астрахань', template_fizo_path=template_fizo_path)
        
        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Arhangelsk(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Arhangelsk.docx')
        super().__init__(parent, controller, 
                        ['Шубина Ксения Александровна'], 
                        template_path, 'Архангельск')
        
        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Bor(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Bor.docx')
        template_fizo_path = os.path.join(TEMPLATES_DIR, 'Bor_FIZO.docx')
        super().__init__(parent, controller, 
                        ['Селюнин Александр Николаевич', 'Аладьин Алексей Николаевич'], 
                        template_path, 'Бор', template_fizo_path=template_fizo_path)
        
        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Ekaterinburg(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Ekaterinburg.docx')
        template_fizo_path = os.path.join(TEMPLATES_DIR, 'Ekaterinburg_FIZO.docx')
        super().__init__(parent, controller, 
                        ['Рослякова Надежда Викторовна'], 
                        template_path, 'Екатеринбург', east=True, template_fizo_path=template_fizo_path)
        
        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Habarovsk(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Habarovsk.docx')
        template_fizo_path = os.path.join(TEMPLATES_DIR, 'Habarovsk_FIZO.docx')
        super().__init__(parent, controller, 
                        ['Смирных Евгений Михайлович', 'Смирных Надежда Евгеньевна'], 
                        template_path, 'Хабаровск', east=True, template_fizo_path=template_fizo_path)
        
        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Kaliningrad(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Kaliningrad.docx')
        template_fizo_path = os.path.join(TEMPLATES_DIR, 'Kaliningrad_FIZO.docx')
        super().__init__(parent, controller, 
                        ['Сироткин Сергей Николаевич'], 
                        template_path, 'Калининград', template_fizo_path=template_fizo_path)
        
        ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: controller.show_page('MainPage')
        ).grid(row=3, column=0, pady=20, padx=20, sticky="ew")

if __name__ == "__main__":
    app = App()
    app.mainloop()
