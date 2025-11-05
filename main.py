import customtkinter as ctk
from datetime import datetime, timedelta
import locale
from tkinter import filedialog
import sys
import os

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

# Получаем абсолютный путь к директории с шаблонами
TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), 'templates')

class LocaleManager:
    """Класс для управления локализацией"""
    
    @staticmethod
    def set_russian_locale():
        """Установка русской локали"""
        try:
            locale_options = ['ru_RU', 'Russian', 'Russian_Russia.1251', 'rus']
            for loc in locale_options:
                try:
                    locale.setlocale(locale.LC_TIME, loc)
                    print(f"Установлена локаль: {loc}")
                    return
                except locale.Error:
                    continue
            print("Не удалось установить русскую локаль, используем системную")
        except Exception as e:
            print(f"Ошибка установки локали: {e}")

class DateFormatter:
    """Класс для форматирования дат"""
    
    MONTH_NAMES = {
        1: 'января', 2: 'февраля', 3: 'марта', 4: 'апреля',
        5: 'мая', 6: 'июня', 7: 'июля', 8: 'августа',
        9: 'сентября', 10: 'октября', 11: 'ноября', 12: 'декабря'
    }
    
    @classmethod
    def get_formatted_issue_date(cls):
        """Получить отформатированную дату формирования"""
        now = datetime.now()
        return f'«{now.day}» {cls.MONTH_NAMES[now.month]} {now.year} г.'
    
    @classmethod
    def get_formatted_work_date(cls, east=False):
        """Получить отформатированную дату проверки"""
        next_day = cls._calculate_work_date(east)
        return f'{next_day.day} {cls.MONTH_NAMES[next_day.month]} {next_day.year} г.'
    
    @staticmethod
    def _calculate_work_date(east=False):
        """Рассчитать дату проверки"""
        today = datetime.now()
        weekday = today.isoweekday()
        
        if weekday == 5:  # Пятница
            return today + timedelta(days=4 if east else 3)
        elif east and weekday == 4:  # Четверг для восточных регионов
            return today + timedelta(days=4)
        elif east:  # Восточные регионы
            return today + timedelta(days=2)
        else:  # Остальные регионы
            return today + timedelta(days=1)

class ClipboardManager:
    """Класс для управления буфером обмена с обработкой ошибок Windows 7"""
    
    @staticmethod
    def copy_text(event):
        """Копирование текста с обработкой ошибок"""
        try:
            if hasattr(event.widget, 'get') and hasattr(event.widget, 'selection_get'):
                selected_text = event.widget.selection_get()
                event.widget.clipboard_clear()
                event.widget.clipboard_append(selected_text)
        except Exception:
            # Резервный метод для Windows 7
            try:
                if hasattr(event.widget, 'get'):
                    event.widget.clipboard_clear()
                    # Просто очищаем буфер, так как selection_get может не работать
            except Exception:
                pass
        return "break"
    
    @staticmethod
    def cut_text(event):
        """Вырезание текста с обработкой ошибок"""
        try:
            if (hasattr(event.widget, 'get') and hasattr(event.widget, 'delete') 
                and hasattr(event.widget, 'selection_get')):
                selected_text = event.widget.selection_get()
                event.widget.clipboard_clear()
                event.widget.clipboard_append(selected_text)
                event.widget.delete("sel.first", "sel.last")
        except Exception:
            # Резервный метод для Windows 7
            try:
                if hasattr(event.widget, 'delete') and hasattr(event.widget, 'get'):
                    # Удаляем выделенный текст без использования selection_get
                    event.widget.delete("sel.first", "sel.last")
            except Exception:
                pass
        return "break"
    
    @staticmethod
    def paste_text(event):
        """Вставка текста с обработкой ошибок"""
        try:
            if hasattr(event.widget, 'insert'):
                text = event.widget.clipboard_get()
                event.widget.insert("insert", text)
        except Exception:
            # Резервный метод для Windows 7
            try:
                # Пытаемся вставить из системного буфера
                import tkinter as tk
                root = tk.Tk()
                root.withdraw()  # Скрываем окно
                text = root.clipboard_get()
                event.widget.insert("insert", text)
                root.destroy()
            except Exception:
                pass
        return "break"

class MultiSelectDropdown(ctk.CTkFrame):
    """Виджет для множественного выбора сотрудников"""
    
    def __init__(self, master, values, **kwargs):
        super().__init__(master, **kwargs)
        self.values = values
        self.selected_workers = []
        self.check_vars = {}
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Создание виджетов dropdown"""
        self.toggle_button = ctk.CTkButton(
            self,
            text="Выбрать сотрудников ▼",
            command=self.toggle_dropdown,
            anchor="w"
        )
        self.toggle_button.pack(fill="x")
        
        self.dropdown_frame = ctk.CTkFrame(self)
        self.is_dropdown_shown = False
        
        for value in self.values:
            var = ctk.BooleanVar(value=False)
            self.check_vars[value] = var
            checkbox = ctk.CTkCheckBox(
                self.dropdown_frame,
                text=value,
                variable=var,
                command=lambda v=value: self._update_selection(v)
            )
            checkbox.pack(anchor="w", pady=2)
    
    def toggle_dropdown(self):
        """Переключение видимости dropdown"""
        if self.is_dropdown_shown:
            self.dropdown_frame.pack_forget()
            self.toggle_button.configure(text="Выбрать сотрудников ▼")
        else:
            self.dropdown_frame.pack(fill="x", pady=5)
            self.toggle_button.configure(text="Скрыть список ▲")
        self.is_dropdown_shown = not self.is_dropdown_shown
    
    def _update_selection(self, value):
        """Обновление выбранных значений"""
        if self.check_vars[value].get():
            if value not in self.selected_workers:
                self.selected_workers.append(value)
        else:
            if value in self.selected_workers:
                self.selected_workers.remove(value)
        self._update_button_text()
    
    def _update_button_text(self):
        """Обновление текста кнопки"""
        if not self.selected_workers:
            self.toggle_button.configure(text="Выбрать сотрудников ▼")
        else:
            self.toggle_button.configure(text=f"Выбрано: {len(self.selected_workers)} ▼")
    
    def get_selected(self):
        """Получить выбранных сотрудников в формате строки"""
        if not self.selected_workers:
            return ""
        elif len(self.selected_workers) == 1:
            return self.selected_workers[0]
        elif len(self.selected_workers) == 2:
            return f"\n{chr(10).join(self.selected_workers)}\n"
        else:
            return chr(10).join(self.selected_workers)
    
    def reset_selection(self):
        """Сброс выбора сотрудников"""
        for var in self.check_vars.values():
            var.set(False)
        self.selected_workers = []
        self._update_button_text()

class DocumentGenerator:
    """Класс для генерации документов Word"""
    
    @staticmethod
    def format_document(replacements_dict, template_path):
        """Форматирование документа Word с заменой плейсхолдеров"""
        try:
            document = Document(template_path)
            DocumentGenerator._process_tables(document, replacements_dict)
            DocumentGenerator._process_paragraphs(document, replacements_dict)
            DocumentGenerator._process_footers(document, replacements_dict)
            return document
        except Exception as e:
            raise Exception(f"Ошибка при форматировании документа: {str(e)}")
    
    @staticmethod
    def _process_tables(document, replacements_dict):
        """Обработка таблиц в документе"""
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        DocumentGenerator._process_paragraph_text(paragraph, replacements_dict)
    
    @staticmethod
    def _process_paragraphs(document, replacements_dict):
        """Обработка обычных параграфов"""
        for paragraph in document.paragraphs:
            DocumentGenerator._process_paragraph_text(paragraph, replacements_dict)
    
    @staticmethod
    def _process_footers(document, replacements_dict):
        """Обработка колонтитулов"""
        for section in document.sections:
            footer = section.footer
            for paragraph in footer.paragraphs:
                DocumentGenerator._process_paragraph_text(paragraph, replacements_dict, font_size=9)
    
    @staticmethod
    def _process_paragraph_text(paragraph, replacements_dict, font_size=13):
        """Обработка текста параграфа с заменой плейсхолдеров"""
        original_alignment = paragraph.alignment
        
        for key, value in replacements_dict.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
                
                # Специальная обработка для идентификаторов
                if '{{IDENTIFICATORS}}' in key:
                    paragraph.paragraph_format.left_indent = Pt(0)
                    paragraph.paragraph_format.first_line_indent = Pt(0)
                
                # Центрирование для большинства полей, кроме должности и ФИО
                if key not in ('{{JOB_TITLE}}', '{{FULLNAME}}'):
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Настройка шрифта
                for run in paragraph.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(font_size)
        
        paragraph.alignment = original_alignment
    
    @staticmethod
    def convert_to_pdf(docx_path, pdf_path):
        """Конвертация DOCX в PDF"""
        try:
            # Попытка использовать comtypes (Windows)
            import comtypes.client
            return DocumentGenerator._convert_with_comtypes(docx_path, pdf_path)
        except ImportError:
            try:
                # Попытка использовать docx2pdf
                from docx2pdf import convert
                convert(docx_path, pdf_path)
                return True
            except Exception as e:
                print(f"Ошибка конвертации через docx2pdf: {e}")
                return False
        except Exception as e:
            print(f"Общая ошибка конвертации: {e}")
            return False
    
    @staticmethod
    def _convert_with_comtypes(docx_path, pdf_path):
        """Конвертация через comtypes (Windows)"""
        try:
            word = comtypes.client.CreateObject('Word.Application')
            word.Visible = False
            
            doc = word.Documents.Open(docx_path)
            doc.SaveAs(pdf_path, FileFormat=17)
            doc.Close()
            word.Quit()
            return True
        except Exception as e:
            print(f"Ошибка конвертации через comtypes: {e}")
            return False

class UINProcessor:
    """Класс для обработки УИНов"""
    
    @staticmethod
    def format_uins(input_string):
        """Форматирование УИНов из строки ввода"""
        if not input_string:
            return []
        
        current_year = datetime.now().year - 2000
        uin_prefix = f'11{current_year}20020'
        parts = input_string.split(', ')
        result_uins = []
        
        for part in parts:
            if '-' in part:
                # Обработка диапазона УИНов
                start_num, end_num = [int(num) for num in part.split('-')]
                for i in range(start_num, end_num + 1):
                    result_uins.append(uin_prefix + str(i))
            else:
                # Одиночный УИН
                result_uins.append(uin_prefix + part)
        
        return result_uins
    
    @staticmethod
    def process_exam_uins(pfo_var, pfo_value, fizo_var, fizo_value, zun_var, zun_value):
        """Обработка УИНов для всех типов проверок"""
        exam_types = []
        exam_data = {}
        all_uins = set()
        
        # Обработка ПФО
        if pfo_var.get():
            uins = UINProcessor.format_uins(pfo_value.get())
            exam_data['личностных (психофизиологических) качеств'] = uins
            exam_types.append('личностных (психофизиологических) качеств')
            all_uins.update(uins)
        
        # Обработка ФИЗО
        if fizo_var.get():
            uins = UINProcessor.format_uins(fizo_value.get())
            exam_data['уровня физической подготовки'] = uins
            exam_types.append('уровня физической подготовки')
            all_uins.update(uins)
        
        # Обработка ЗУН
        if zun_var.get():
            uins = UINProcessor.format_uins(zun_value.get())
            exam_data['знаний, умений и навыков'] = uins
            exam_types.append('знаний, умений и навыков')
            all_uins.update(uins)
        
        # Формирование результирующих строк
        uins_string = UINProcessor._build_uins_string(exam_data)
        exams_string = UINProcessor._build_exams_string(exam_types)
        sorted_uins = sorted(list(all_uins))
        
        return sorted_uins, uins_string, exams_string
    
    @staticmethod
    def _build_uins_string(exam_data):
        """Построение строки с УИНами"""
        if not exam_data:
            return ""
        
        result_parts = []
        for index, (exam_type, uins) in enumerate(exam_data.items(), 1):
            uins_text = "; ".join(uins)
            
            if len(exam_data) == 1:
                result_parts.append(f'{exam_type} аттестуемых лиц (УИН: {uins_text}).')
            elif index == len(exam_data):
                result_parts.append(f'а также {exam_type} аттестуемых лиц (УИН: {uins_text}).')
            else:
                result_parts.append(f'{exam_type} аттестуемых лиц (УИН: {uins_text}), ')
        
        return "".join(result_parts)
    
    @staticmethod
    def _build_exams_string(exam_types):
        """Построение строки с типами экзаменов"""
        if not exam_types:
            return ""
        
        if len(exam_types) == 1:
            return exam_types[0]
        elif len(exam_types) == 2:
            return f"{', а также '.join(exam_types)}"
        else:
            return f'{exam_types[0]}, {exam_types[1]}, а также {exam_types[2]}'

class BasePage(ctk.CTkFrame):
    """Базовый класс для всех страниц приложения"""
    
    def __init__(self, parent, controller, worker_list=None, template_path='', location_name='', 
                 is_eastern_region=False, fizo_template_path=''):
        super().__init__(parent)
        
        self.controller = controller
        self.template_path = template_path
        self.fizo_template_path = fizo_template_path
        self.location_name = location_name
        self.is_eastern_region = is_eastern_region
        self.worker_list = worker_list or []
        self.results_data = {}
        
        self._setup_grid_configuration()
        self._create_variables()
        self._create_interface()
        self._setup_navigation()
    
    def _setup_grid_configuration(self):
        """Настройка конфигурации grid"""
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
    
    def _create_variables(self):
        """Создание переменных для хранения данных"""
        # Основные переменные
        self.application_number = ctk.StringVar()
        self.issue_date = ctk.StringVar()
        self.work_date = ctk.StringVar()
        self.start_time = ctk.StringVar(value='10:00')
        self.end_time = ctk.StringVar(value='16:00')
        
        # Переменные типов проверок
        self.is_pfo_selected = ctk.BooleanVar(value=False)
        self.is_fizo_selected = ctk.BooleanVar(value=False)
        self.is_zun_selected = ctk.BooleanVar(value=False)
        
        # Переменные УИНов
        self.pfo_uins = ctk.StringVar()
        self.fizo_uins = ctk.StringVar()
        self.zun_uins = ctk.StringVar()
        self.fizo_application_number = ctk.StringVar()
        
        # Переменные подписантов
        self.approving_person = ctk.StringVar(value='')
        self.executor_person = ctk.StringVar(value='')
    
    def _create_interface(self):
        """Создание интерфейса страницы"""
        self._create_annotation_frame()
        self._create_info_frame()
        self._create_action_buttons()
    
    def _create_annotation_frame(self):
        """Создание рамки с аннотацией"""
        self.annotation_frame = ctk.CTkFrame(master=self)
        self.annotation_frame.grid(row=0, column=0, padx=20, pady=10, sticky="nsew")
        
        title_label = ctk.CTkLabel(
            self.annotation_frame, 
            text='Выберите необходимые параметры', 
            font=('TimesNewRoman', 20)
        )
        title_label.grid(row=1, column=0, columnspan=4, sticky='nsew')
    
    def _create_info_frame(self):
        """Создание основной информационной рамки"""
        self.info_frame = ctk.CTkFrame(master=self)
        self.info_frame.grid(row=1, column=0, padx=20, pady=10, sticky='nsew')
        
        self._create_basic_info_section()
        self._create_exam_type_section()
        self._create_personnel_section()
        self._create_approval_section()
        
        # Настройка колонок
        for i in range(4):
            self.info_frame.grid_columnconfigure(i, weight=1 if i > 0 else 0)
    
    def _create_basic_info_section(self):
        """Создание секции основной информации"""
        font = ('TimesNewRoman', 13)
        row = 0
        
        # Номер заявки
        ctk.CTkLabel(self.info_frame, text='Номер заявки:', font=font).grid(
            row=row, column=0, padx=5, pady=5, sticky='w')
        
        self.application_number_entry = ctk.CTkEntry(
            self.info_frame, 
            width=300, 
            textvariable=self.application_number,
            placeholder_text='Номер заявки по нормированному заданию'
        )
        self.application_number_entry.grid(row=row, column=1, columnspan=3, padx=5, pady=5, sticky='ew')
        row += 1
        
        # Дата формирования
        ctk.CTkLabel(self.info_frame, text='Дата формирования:', font=font).grid(
            row=row, column=0, padx=5, pady=5, sticky='w')
        
        self.issue_date_entry = ctk.CTkEntry(self.info_frame, width=300, textvariable=self.issue_date)
        self.issue_date_entry.insert(0, DateFormatter.get_formatted_issue_date())
        self.issue_date_entry.grid(row=row, column=1, columnspan=3, padx=5, pady=5, sticky='ew')
        row += 1
        
        # Дата проверки
        ctk.CTkLabel(self.info_frame, text='Дата проведения проверки:', font=font).grid(
            row=row, column=0, padx=5, pady=5, sticky='w')
        
        self.work_date_entry = ctk.CTkEntry(self.info_frame, width=300, textvariable=self.work_date)
        self.work_date_entry.insert(0, DateFormatter.get_formatted_work_date(self.is_eastern_region))
        self.work_date_entry.grid(row=row, column=1, columnspan=3, padx=5, pady=5, sticky='ew')
        row += 1
        
        # Время проверки
        ctk.CTkLabel(self.info_frame, text='Время проведения проверки:', font=font).grid(
            row=row, column=0, padx=5, pady=5, sticky='w')
        
        self.start_time_entry = ctk.CTkEntry(self.info_frame, width=100, textvariable=self.start_time)
        self.start_time_entry.grid(row=row, column=1, padx=5, pady=5, sticky='w')
        
        ctk.CTkLabel(self.info_frame, text='-').grid(row=row, column=2, padx=5, pady=5)
        
        self.end_time_entry = ctk.CTkEntry(self.info_frame, width=100, textvariable=self.end_time)
        self.end_time_entry.grid(row=row, column=3, padx=5, pady=5, sticky='w')
        row += 1
    
    def _create_exam_type_section(self):
        """Создание секции типов проверок"""
        font = ('TimesNewRoman', 13)
        row = 4
        
        # Чекбоксы типов проверок
        self.pfo_checkbox = ctk.CTkCheckBox(
            self.info_frame, 
            text='ПФО', 
            variable=self.is_pfo_selected, 
            command=lambda: self._toggle_exam_entry('pfo')
        )
        self.pfo_checkbox.grid(row=row, column=0, padx=5, pady=5, sticky='w')
        
        self.fizo_checkbox = ctk.CTkCheckBox(
            self.info_frame, 
            text='ФИЗО', 
            variable=self.is_fizo_selected, 
            command=lambda: self._toggle_exam_entry('fizo')
        )
        self.fizo_checkbox.grid(row=row+1, column=0, padx=5, pady=5, sticky='w')
        
        self.zun_checkbox = ctk.CTkCheckBox(
            self.info_frame, 
            text='ЗУН', 
            variable=self.is_zun_selected, 
            command=lambda: self._toggle_exam_entry('zun')
        )
        self.zun_checkbox.grid(row=row+2, column=0, padx=5, pady=5, sticky='w')
        
        # Поля ввода УИНов
        self.pfo_uins_entry = ctk.CTkEntry(self.info_frame, textvariable=self.pfo_uins)
        self.fizo_uins_entry = ctk.CTkEntry(self.info_frame, textvariable=self.fizo_uins)
        self.zun_uins_entry = ctk.CTkEntry(self.info_frame, textvariable=self.zun_uins)
        
        # Поле для номера заявки ФИЗО
        self.fizo_number_label = ctk.CTkLabel(
            self.info_frame, 
            text='Номер заявки для ФИЗО', 
            font=font
        )
        self.fizo_number_entry = ctk.CTkEntry(self.info_frame, textvariable=self.fizo_application_number)
    
    def _create_personnel_section(self):
        """Создание секции выбора персонала"""
        font = ('TimesNewRoman', 13)
        row = 7
        
        # Выбор сотрудников
        ctk.CTkLabel(self.info_frame, text='Выберите сотрудников:', font=font).grid(
            row=row, column=0, padx=5, pady=5, sticky='w')
        
        self.workers_dropdown = MultiSelectDropdown(self.info_frame, self.worker_list)
        self.workers_dropdown.grid(row=row, column=1, columnspan=4, padx=5, pady=5, sticky='ew')
        row += 1
        
        # Выбор должностного лица
        ctk.CTkLabel(self.info_frame, text='Выберите должностное лицо:', font=font).grid(
            row=row, column=0, pady=5, sticky='w')
        
        self.bezzubcev_checkbox = ctk.CTkCheckBox(
            self.info_frame, 
            text='Беззубцев А.А.', 
            font=font, 
            variable=self.approving_person, 
            onvalue='Беззубцев', 
            offvalue=''
        )
        self.bezzubcev_checkbox.grid(row=row, column=1, sticky='w')
        
        self.aliabiev_checkbox = ctk.CTkCheckBox(
            self.info_frame, 
            text='Алябьев А.Б.', 
            font=font, 
            variable=self.approving_person, 
            onvalue='Алябьев', 
            offvalue=''
        )
        self.aliabiev_checkbox.grid(row=row, column=2, sticky='w')
        
        self.popirina_checkbox = ctk.CTkCheckBox(
            self.info_frame, 
            text='Попырина Е.М.', 
            font=font, 
            variable=self.approving_person, 
            onvalue='Попырина', 
            offvalue=''
        )
        self.popirina_checkbox.grid(row=row, column=3, sticky='w')
        row += 1
        
        # Исполнитель
        ctk.CTkLabel(self.info_frame, text='Укажите исполнителя:', font=font).grid(
            row=row, column=0, sticky='w')
        
        self.marusya_checkbox = ctk.CTkCheckBox(
            self.info_frame, 
            text='Разгуляева М.А.', 
            font=font, 
            variable=self.executor_person, 
            onvalue='Маруся', 
            offvalue=''
        )
        self.marusya_checkbox.grid(row=row, column=1, sticky='w')
        
        self.vantuz_checkbox = ctk.CTkCheckBox(
            self.info_frame, 
            text='Воротилов И.И.', 
            font=font, 
            variable=self.executor_person, 
            onvalue='Вантуз', 
            offvalue=''
        )
        self.vantuz_checkbox.grid(row=row, column=2, sticky='w')
        
        self.creator_checkbox = ctk.CTkCheckBox(
            self.info_frame, 
            text='Желудков А.В.', 
            font=font, 
            variable=self.executor_person, 
            onvalue='Создатель!!!', 
            offvalue=''
        )
        self.creator_checkbox.grid(row=row, column=3, sticky='w')
    
    def _create_approval_section(self):
        """Создание секции утверждения"""
        self.save_button = ctk.CTkButton(
            self, 
            text='Сформировать документ', 
            font=('TimesNewRoman', 13), 
            command=self._save_data_and_generate_document
        )
        self.save_button.grid(row=2, column=0, sticky='nsew')
    
    def _create_action_buttons(self):
        """Создание кнопок действий (переопределяется в дочерних классах)"""
        pass
    
    def _toggle_exam_entry(self, exam_type):
        """Переключение видимости полей ввода для типов проверок"""
        if exam_type == 'pfo':
            if self.is_pfo_selected.get():
                self.pfo_uins_entry.grid(row=4, column=1, columnspan=3, padx=5, pady=5, sticky='ew')
            else:
                self.pfo_uins_entry.grid_forget()
        
        elif exam_type == 'fizo':
            if self.is_fizo_selected.get():
                if self.fizo_template_path:
                    self.fizo_uins_entry.grid(row=5, column=1, padx=5, pady=5, sticky='ew')
                    self.fizo_number_label.grid(row=5, column=2, sticky='ew')
                    self.fizo_number_entry.grid(row=5, column=3, sticky='ew')
                else:
                    self.fizo_uins_entry.grid(row=5, column=1, columnspan=3, padx=5, pady=5, sticky='ew')
                    self.fizo_number_label.grid_forget()
                    self.fizo_number_entry.grid_forget()
            else:
                self.fizo_uins_entry.grid_forget()
                self.fizo_number_label.grid_forget()
                self.fizo_number_entry.grid_forget()
        
        elif exam_type == 'zun':
            if self.is_zun_selected.get():
                self.zun_uins_entry.grid(row=6, column=1, columnspan=3, padx=5, pady=5, sticky='ew')
            else:
                self.zun_uins_entry.grid_forget()
    
    def _setup_navigation(self):
        """Настройка навигации по полям ввода"""
        widgets = [
            self.application_number_entry, self.issue_date_entry, self.work_date_entry,
            self.start_time_entry, self.end_time_entry,
            self.pfo_checkbox, self.fizo_checkbox, self.zun_checkbox,
            self.pfo_uins_entry, self.fizo_uins_entry, self.zun_uins_entry,
            self.fizo_number_entry,
            self.workers_dropdown.toggle_button,
            self.bezzubcev_checkbox, self.aliabiev_checkbox, self.popirina_checkbox,
            self.marusya_checkbox, self.vantuz_checkbox, self.creator_checkbox,
            self.save_button
        ]
        
        for widget in widgets:
            if hasattr(widget, 'bind'):
                widget.bind('<Up>', lambda e: self._navigate_fields(e, -1))
                widget.bind('<Down>', lambda e: self._navigate_fields(e, 1))
                widget.bind('<Control-c>', ClipboardManager.copy_text)
                widget.bind('<Control-x>', ClipboardManager.cut_text)
                widget.bind('<Control-v>', ClipboardManager.paste_text)
    
    def _navigate_fields(self, event, direction):
        """Навигация по полям ввода с помощью стрелок"""
        widgets = [
            self.application_number_entry, self.issue_date_entry, self.work_date_entry,
            self.start_time_entry, self.end_time_entry,
            self.pfo_uins_entry, self.fizo_uins_entry, self.zun_uins_entry,
            self.fizo_number_entry,
            self.workers_dropdown.toggle_button,
            self.bezzubcev_checkbox, self.aliabiev_checkbox, self.popirina_checkbox,
            self.marusya_checkbox, self.vantuz_checkbox, self.creator_checkbox,
            self.save_button
        ]
        
        current_widget = event.widget
        try:
            current_index = widgets.index(current_widget)
            new_index = (current_index + direction) % len(widgets)
            widgets[new_index].focus_set()
        except ValueError:
            pass
        
        return "break"
    
    def on_show(self):
        """Действия при показе страницы"""
        self._reset_fields()
    
    def _reset_fields(self):
        """Сброс всех полей ввода к значениям по умолчанию"""
        # Сброс основных полей
        self.application_number.set('')
        
        self.issue_date_entry.delete(0, 'end')
        self.issue_date_entry.insert(0, DateFormatter.get_formatted_issue_date())
        
        self.work_date_entry.delete(0, 'end')
        self.work_date_entry.insert(0, DateFormatter.get_formatted_work_date(self.is_eastern_region))
        
        self.start_time.set('10:00')
        self.end_time.set('16:00')
        
        # Сброс типов проверок
        self.is_pfo_selected.set(False)
        self.is_fizo_selected.set(False)
        self.is_zun_selected.set(False)
        
        self.pfo_uins.set('')
        self.fizo_uins.set('')
        self.zun_uins.set('')
        self.fizo_application_number.set('')
        
        # Скрытие полей ввода
        self.pfo_uins_entry.grid_forget()
        self.fizo_uins_entry.grid_forget()
        self.zun_uins_entry.grid_forget()
        self.fizo_number_label.grid_forget()
        self.fizo_number_entry.grid_forget()
        
        # Сброс выбора сотрудников
        self.workers_dropdown.reset_selection()
        
        # Сброс выбора лиц
        self.approving_person.set('')
        self.executor_person.set('')
    
    def _get_exams_without_fizo(self):
        """Получить типы экзаменов без ФИЗО"""
        exam_types = []
        
        if self.is_pfo_selected.get():
            exam_types.append('личностных (психофизиологических) качеств')
        if self.is_zun_selected.get():
            exam_types.append('знаний, умений и навыков')
        
        if len(exam_types) == 1:
            return exam_types[0]
        elif len(exam_types) == 2:
            return ', а также '.join(exam_types)
        else:
            return ""
    
    def _get_approving_person_info(self):
        """Получить информацию об утверждающем лице"""
        selected_person = self.approving_person.get()
        
        if selected_person == 'Беззубцев':
            return ['Начальник отдела проверок сил ОТБ', 'А.А. Беззубцев']
        elif selected_person == 'Алябьев':
            return ['Заместитель начальника Службы', 'А.Б. Алябьев']
        elif selected_person == 'Попырина':
            return ['Заместитель начальника отдела проверок сил ОТБ', 'Е.М. Попырина']
        else:
            return ['', '']
    
    def _get_executor_info(self):
        """Получить информацию об исполнителе"""
        selected_executor = self.executor_person.get()
        
        if selected_executor == 'Вантуз':
            return 'Воротилов Иван Иванович', 'oa13@msecurity.ru'
        elif selected_executor == 'Маруся':
            return 'Разгуляева Мария Андреевна', 'oa6@msecurity.ru'
        elif selected_executor == 'Создатель!!!':
            return 'Желудков Андрей Викторович', 'oa15@msecurity.ru'
        else:
            return '', ''
    
    def _calculate_estimated_time(self):
        """Рассчитать затраченное время"""
        try:
            start_hour = int(self.start_time.get().split(':')[0])
            end_hour = int(self.end_time.get().split(':')[0])
            total_hours = end_hour - start_hour
            
            # Добавление часа для ФИЗО, если есть отдельный шаблон
            if self.is_fizo_selected.get() and self.fizo_template_path:
                total_hours += 1
            
            if total_hours >= 5:
                return f'{total_hours} часов'
            else:
                return f'{total_hours} часа'
        except:
            return '6 часов'
    
    def _prepare_document_data(self):
        """Подготовка данных для заполнения документа"""
        # Обработка УИНов
        all_uins, uins_text, exams_text = UINProcessor.process_exam_uins(
            self.is_pfo_selected, self.pfo_uins,
            self.is_fizo_selected, self.fizo_uins,
            self.is_zun_selected, self.zun_uins
        )
        
        # Форматирование идентификаторов
        identifiers_text = ('; '.join(all_uins) + '.').strip() if len(all_uins) > 1 else all_uins[0].strip()
        
        # Форматирование УИНов ФИЗО
        fizo_uins_list = UINProcessor.format_uins(self.fizo_uins.get())
        fizo_ids_text = '; '.join(fizo_uins_list) + '.' if len(fizo_uins_list) > 1 else fizo_uins_list[0]
        
        # Получение информации о лицах
        job_title, full_name = self._get_approving_person_info()
        executor_name, executor_email = self._get_executor_info()
        
        # Подготовка данных для замены
        time_range = f'с {self.start_time.get()} час. до {self.end_time.get()} час.'
        
        return {
            '{{NUMBER}}': self.application_number.get(),
            '{{NUMBER_FIZO}}': self.fizo_application_number.get(),
            '{{DATE_OF_ISSUE}}': self.issue_date.get().lower(),
            '{{DATE_TO_WORK}}': self.work_date.get().lower(),
            '{{TYPE_OF_EXAMS}}': exams_text,
            '{{IDENTIFICATORS}}': identifiers_text,
            '{{EXAMS_P2}}': uins_text,
            '{{TIME}}': time_range,
            '{{EST_TIME}}': self._calculate_estimated_time(),
            '{{PEOPLE}}': self.workers_dropdown.get_selected(),
            '{{JOB_TITLE}}': job_title,
            '{{FULLNAME}}': full_name,
            '{{ISSUER}}': executor_name,
            '{{EMAIL}}': executor_email,
            '{{EXAMS}}': self._get_exams_without_fizo(),
            '{{FIZO_UIN}}': fizo_ids_text
        }
    
    def _save_data_and_generate_document(self):
        """Сохранение данных и генерация документа"""
        try:
            # Подготовка данных
            document_data = self._prepare_document_data()
            
            # Генерация основного документа
            self._generate_main_document(document_data)
            
            # Генерация документа ФИЗО при необходимости
            if self.is_fizo_selected.get() and self.fizo_template_path:
                self._generate_fizo_document(document_data)
            
            self._show_success_message()
            
        except Exception as e:
            self._show_error_message(str(e))
    
    def _generate_main_document(self, document_data):
        """Генерация основного документа"""
        edited_doc = DocumentGenerator.format_document(document_data, self.template_path)
        
        default_filename = f'Нормированное задание на {DateFormatter._calculate_work_date(self.is_eastern_region).strftime("%d.%m")} {self.location_name}'
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
            initialfile=default_filename,
            title="Сохранить документ как"
        )
        
        if file_path:
            edited_doc.save(file_path)
            # Конвертация в PDF
            pdf_path = file_path.replace('.docx', '.pdf')
            DocumentGenerator.convert_to_pdf(file_path, pdf_path)
    
    def _generate_fizo_document(self, document_data):
        """Генерация документа ФИЗО"""
        fizo_doc = DocumentGenerator.format_document(document_data, self.fizo_template_path)
        
        default_filename = f'Заявка на {DateFormatter._calculate_work_date(self.is_eastern_region).strftime("%d.%m")} {self.location_name}'
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
            initialfile=default_filename,
            title="Сохранить документ ФИЗО как"
        )
        
        if file_path:
            fizo_doc.save(file_path)
            # Конвертация в PDF
            pdf_path = file_path.replace('.docx', '.pdf')
            DocumentGenerator.convert_to_pdf(file_path, pdf_path)
    
    def _show_success_message(self):
        """Показать сообщение об успехе"""
        popup = ctk.CTkToplevel(self.info_frame)
        popup.title("Уведомление")
        popup.geometry("300x100")
        popup.resizable(False, False)
        
        ctk.CTkLabel(popup, text="Документ сформирован!").pack(pady=20)
        ctk.CTkButton(popup, text="OK", command=popup.destroy).pack(pady=5)
    
    def _show_error_message(self, error_text):
        """Показать сообщение об ошибке"""
        popup = ctk.CTkToplevel(self.info_frame)
        popup.title("Ошибка")
        popup.geometry("400x100")
        popup.resizable(False, False)
        
        ctk.CTkLabel(popup, text=f"Ошибка при формировании документа: {error_text}").pack(pady=20)
        ctk.CTkButton(popup, text="OK", command=popup.destroy).pack(pady=5)

# Классы для конкретных локаций (сокращено для примера)
class MainPage(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self._create_interface()
    
    def _create_interface(self):
        """Создание интерфейса главной страницы"""
        # Заголовок
        title_label = ctk.CTkLabel(self, text="Главная страница", font=("TimesNewRoman", 24))
        title_label.grid(row=0, column=0, columnspan=3, pady=20, padx=20)
        
        # Кнопки локаций
        location_buttons = [
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
        
        for i, (button_text, page_name) in enumerate(location_buttons):
            row = i // 3 + 1
            col = i % 3
            button = ctk.CTkButton(
                self, 
                text=button_text,
                command=lambda p=page_name: self.controller.show_page(p),
                height=40,
                width=200
            )
            button.grid(row=row, column=col, pady=10, padx=10, sticky='nsew')
        
        # Кнопка выхода
        exit_button = ctk.CTkButton(
            self,
            text="Выход",
            command=self.controller.quit,
            fg_color="red",
            hover_color="darkred",
            height=40,
            width=200
        )
        exit_button.grid(row=len(location_buttons)//3 + 2, column=1, pady=20, padx=10, sticky='nsew')

class Novoros(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Novoros.docx')
        super().__init__(
            parent, controller, 
            worker_list=['Лагутин Ростислав Борисович', 'Соловьев Сергей Викторович', 'Соловьев Артем Андреевич'], 
            template_path=template_path, 
            location_name='Новороссийск'
        )
        self._create_back_button()
    
    def _create_back_button(self):
        """Создание кнопки возврата"""
        back_button = ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: self.controller.show_page('MainPage')
        )
        back_button.grid(row=3, column=0, pady=20, padx=20, sticky="ew")

class Vladivostok(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Vladivostok.docx')
        fizo_template_path = os.path.join(TEMPLATES_DIR, 'Vladivostok_FIZO.docx')
        super().__init__(
            parent, controller, 
            worker_list=['Проценко Игорь Владимирович', 'Лабонин Илья Валентинович', 'Родионова Елена Ивановна'], 
            template_path=template_path, 
            location_name='Владивосток', 
            is_eastern_region=True,
            fizo_template_path=fizo_template_path
        )
        self._create_back_button()
    
    def _create_back_button(self):
        """Создание кнопки возврата"""
        back_button = ctk.CTkButton(
            self, 
            text='Вернуться назад', 
            command=lambda: self.controller.show_page('MainPage')
        )
        back_button.grid(row=3, column=0, pady=20, padx=20, sticky="ew")

# Аналогично для других классов локаций (US, Spb, Sevastopol, и т.д.)
# Сокращено для экономии места...

class App(ctk.CTk):
    """Главное приложение"""
    
    def __init__(self):
        super().__init__()
        self._setup_appearance()
        self._create_interface()
    
    def _setup_appearance(self):
        """Настройка внешнего вида приложения"""
        self.title("Генератор нормированных заданий")
        self.geometry("1000x700")
        
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
    
    def _create_interface(self):
        """Создание интерфейса приложения"""
        # Главный контейнер
        self.main_container = ctk.CTkFrame(self)
        self.main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Создание страниц
        self.pages = self._create_pages()
        self._arrange_pages()
        
        # Показ главной страницы
        self.show_page("MainPage")
    
    def _create_pages(self):
        """Создание всех страниц приложения"""
        return {
            "MainPage": MainPage(self.main_container, self),
            "Novoros": Novoros(self.main_container, self),
            'Astrakhan': Astrakhan(self.main_container, self),
            'Arhangelsk': Arhangelsk(self.main_container, self),
            'Bor': Bor(self.main_container, self),
            'Ekaterinburg': Ekaterinburg(self.main_container, self),
            'Habarovsk': Habarovsk(self.main_container, self),
            'Kaliningrad': Kaliningrad(self.main_container, self),
            'NN': NN(self.main_container, self),
            'Murmansk': Murmansk(self.main_container, self),
            'Nahodka': Nahodka(self.main_container, self),
            'PK': PK(self.main_container, self),
            'Rostov': Rostov(self.main_container, self),
            'Samara': Samara(self.main_container, self),
            'Sevastopol': Sevastopol(self.main_container, self),
            'Spb': Spb(self.main_container, self),
            'US': US(self.main_container, self),
            'Vladivostok': Vladivostok(self.main_container, self)
        }
    
    def _arrange_pages(self):
        """Размещение всех страниц в контейнере"""
        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")
        
        # Конфигурация grid
        self.main_container.grid_rowconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)
    
    def show_page(self, page_name):
        """Показать выбранную страницу"""
        page = self.pages.get(page_name)
        if page:
            page.tkraise()
            if hasattr(page, "on_show"):
                page.on_show()

if __name__ == "__main__":
    # Установка локали
    LocaleManager.set_russian_locale()
    
    # Запуск приложения
    app = App()
    app.mainloop()
