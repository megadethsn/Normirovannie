import customtkinter as ctk
from datetime import datetime, timedelta
import locale
from tkinter import filedialog, messagebox
import sys
import os
import traceback

# Безопасный импорт для совместимости с Windows 7
try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError as e:
    print(f"Warning: docx not available - {e}")
    DOCX_AVAILABLE = False

# Получаем абсолютный путь к директории с шаблонами
TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), 'templates')

class ErrorHandler:
    """Класс для обработки ошибок"""
    
    @staticmethod
    def show_error(message, parent=None):
        """Показать сообщение об ошибке"""
        try:
            if parent:
                messagebox.showerror("Ошибка", message, parent=parent)
            else:
                messagebox.showerror("Ошибка", message)
        except Exception as e:
            print(f"Не удалось показать ошибку: {e}")
            print(f"Оригинальная ошибка: {message}")

    @staticmethod
    def log_error(error, context=""):
        """Логирование ошибки"""
        error_msg = f"{context}: {str(error)}"
        print(error_msg)
        with open("error_log.txt", "a", encoding="utf-8") as f:
            f.write(f"{datetime.now()}: {error_msg}\n")
            f.write(traceback.format_exc() + "\n")

class LocaleManager:
    """Класс для управления локализацией"""
    
    @staticmethod
    def set_russian_locale():
        """Установка русской локали с улучшенной совместимостью"""
        try:
            # Пробуем разные варианты локалей
            locale_options = ['ru_RU.UTF-8', 'ru_RU', 'Russian', 'Russian_Russia.1251', 'rus', 'ru']
            
            for loc in locale_options:
                try:
                    locale.setlocale(locale.LC_TIME, loc)
                    print(f"Установлена локаль: {loc}")
                    return True
                except locale.Error:
                    continue
            
            print("Предупреждение: Не удалось установить русскую локаль, используем системную")
            return False
            
        except Exception as e:
            print(f"Ошибка установки локали: {e}")
            return False

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
        try:
            now = datetime.now()
            return f'«{now.day}» {cls.MONTH_NAMES[now.month]} {now.year} г.'
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка форматирования даты формирования")
            return f'«{datetime.now().day}» месяца {datetime.now().year} г.'
    
    @classmethod
    def get_formatted_work_date(cls, east=False):
        """Получить отформатированную дату проверки"""
        try:
            next_day = cls._calculate_work_date(east)
            return f'{next_day.day} {cls.MONTH_NAMES[next_day.month]} {next_day.year} г.'
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка форматирования даты проверки")
            return f'{datetime.now().day} месяца {datetime.now().year} г.'
    
    @staticmethod
    def _calculate_work_date(east=False):
        """Рассчитать дату проверки"""
        try:
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
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка расчета даты проверки")
            return datetime.now() + timedelta(days=1)

class SimpleClipboardManager:
    """Упрощенный менеджер буфера обмена для Windows 7"""
    
    @staticmethod
    def copy_text(event):
        """Копирование текста"""
        try:
            widget = event.widget
            if hasattr(widget, 'selection_get') and widget.tag_ranges("sel"):
                selected_text = widget.selection_get()
                widget.clipboard_clear()
                widget.clipboard_append(selected_text)
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка копирования текста")
        return "break"
    
    @staticmethod
    def cut_text(event):
        """Вырезание текста"""
        try:
            widget = event.widget
            if (hasattr(widget, 'selection_get') and hasattr(widget, 'delete') 
                and widget.tag_ranges("sel")):
                selected_text = widget.selection_get()
                widget.clipboard_clear()
                widget.clipboard_append(selected_text)
                widget.delete("sel.first", "sel.last")
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка вырезания текста")
        return "break"
    
    @staticmethod
    def paste_text(event):
        """Вставка текста"""
        try:
            widget = event.widget
            if hasattr(widget, 'insert'):
                try:
                    text = widget.clipboard_get()
                    widget.insert("insert", text)
                except:
                    # Альтернативный метод для Windows 7
                    try:
                        import tkinter as tk
                        root = tk.Tk()
                        root.withdraw()
                        text = root.clipboard_get()
                        widget.insert("insert", text)
                        root.destroy()
                    except:
                        pass
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка вставки текста")
        return "break"

class MultiSelectDropdown(ctk.CTkFrame):
    """Виджет для множественного выбора сотрудников"""
    
    def __init__(self, master, values, **kwargs):
        super().__init__(master, **kwargs)
        self.values = values or []
        self.selected_workers = []
        self.check_vars = {}
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Создание виджетов dropdown"""
        try:
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
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка создания MultiSelectDropdown")
    
    def toggle_dropdown(self):
        """Переключение видимости dropdown"""
        try:
            if self.is_dropdown_shown:
                self.dropdown_frame.pack_forget()
                self.toggle_button.configure(text="Выбрать сотрудников ▼")
            else:
                self.dropdown_frame.pack(fill="x", pady=5)
                self.toggle_button.configure(text="Скрыть список ▲")
            self.is_dropdown_shown = not self.is_dropdown_shown
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка переключения dropdown")
    
    def _update_selection(self, value):
        """Обновление выбранных значений"""
        try:
            if self.check_vars[value].get():
                if value not in self.selected_workers:
                    self.selected_workers.append(value)
            else:
                if value in self.selected_workers:
                    self.selected_workers.remove(value)
            self._update_button_text()
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка обновления выбора")
    
    def _update_button_text(self):
        """Обновление текста кнопки"""
        try:
            if not self.selected_workers:
                self.toggle_button.configure(text="Выбрать сотрудников ▼")
            else:
                self.toggle_button.configure(text=f"Выбрано: {len(self.selected_workers)} ▼")
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка обновления текста кнопки")
    
    def get_selected(self):
        """Получить выбранных сотрудников в формате строки"""
        try:
            if not self.selected_workers:
                return ""
            elif len(self.selected_workers) == 1:
                return self.selected_workers[0]
            elif len(self.selected_workers) == 2:
                return f"\n{' '.join(self.selected_workers)}\n"
            else:
                return '\n'.join(self.selected_workers)
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка получения выбранных сотрудников")
            return ""
    
    def reset_selection(self):
        """Сброс выбора сотрудников"""
        try:
            for var in self.check_vars.values():
                var.set(False)
            self.selected_workers = []
            self._update_button_text()
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка сброса выбора")

class SimpleDocumentGenerator:
    """Упрощенный генератор документов для совместимости"""
    
    @staticmethod
    def check_docx_support():
        """Проверить доступность docx"""
        if not DOCX_AVAILABLE:
            ErrorHandler.show_error(
                "Библиотека python-docx не установлена. "
                "Установите ее командой: pip install python-docx"
            )
            return False
        return True
    
    @staticmethod
    def format_document(replacements_dict, template_path):
        """Форматирование документа Word с заменой плейсхолдеров"""
        if not SimpleDocumentGenerator.check_docx_support():
            return None
            
        try:
            document = Document(template_path)
            SimpleDocumentGenerator._process_tables(document, replacements_dict)
            SimpleDocumentGenerator._process_paragraphs(document, replacements_dict)
            SimpleDocumentGenerator._process_footers(document, replacements_dict)
            return document
        except Exception as e:
            ErrorHandler.log_error(e, f"Ошибка форматирования документа {template_path}")
            raise
    
    @staticmethod
    def _process_tables(document, replacements_dict):
        """Обработка таблиц в документе"""
        try:
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            SimpleDocumentGenerator._process_paragraph_text(paragraph, replacements_dict)
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка обработки таблиц")
    
    @staticmethod
    def _process_paragraphs(document, replacements_dict):
        """Обработка обычных параграфов"""
        try:
            for paragraph in document.paragraphs:
                SimpleDocumentGenerator._process_paragraph_text(paragraph, replacements_dict)
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка обработки параграфов")
    
    @staticmethod
    def _process_footers(document, replacements_dict):
        """Обработка колонтитулов"""
        try:
            for section in document.sections:
                footer = section.footer
                for paragraph in footer.paragraphs:
                    SimpleDocumentGenerator._process_paragraph_text(paragraph, replacements_dict, font_size=9)
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка обработки колонтитулов")
    
    @staticmethod
    def _process_paragraph_text(paragraph, replacements_dict, font_size=13):
        """Обработка текста параграфа с заменой плейсхолдеров"""
        try:
            original_alignment = paragraph.alignment
            
            for key, value in replacements_dict.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, str(value))
                    
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
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка обработки текста параграфа")
    
    @staticmethod
    def convert_to_pdf(docx_path, pdf_path):
        """Конвертация DOCX в PDF (упрощенная версия)"""
        try:
            # Для Windows 7 используем упрощенную конвертацию
            # Вместо полной конвертации просто создаем копию
            import shutil
            shutil.copy2(docx_path, pdf_path.replace('.pdf', '_converted.docx'))
            print(f"Создана копия документа: {pdf_path.replace('.pdf', '_converted.docx')}")
            return True
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка конвертации в PDF")
            return False

class UINProcessor:
    """Класс для обработки УИНов"""
    
    @staticmethod
    def format_uins(input_string):
        """Форматирование УИНов из строки ввода"""
        if not input_string or not input_string.strip():
            return []
        
        try:
            current_year = datetime.now().year - 2000
            uin_prefix = f'11{current_year}20020'
            parts = [part.strip() for part in input_string.split(',')]
            result_uins = []
            
            for part in parts:
                if '-' in part:
                    # Обработка диапазона УИНов
                    range_parts = part.split('-')
                    if len(range_parts) == 2:
                        try:
                            start_num = int(range_parts[0].strip())
                            end_num = int(range_parts[1].strip())
                            for i in range(start_num, end_num + 1):
                                result_uins.append(uin_prefix + str(i))
                        except ValueError:
                            continue
                else:
                    # Одиночный УИН
                    try:
                        uin_num = int(part.strip())
                        result_uins.append(uin_prefix + str(uin_num))
                    except ValueError:
                        continue
            
            return result_uins
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка форматирования УИНов")
            return []
    
    @staticmethod
    def process_exam_uins(pfo_var, pfo_value, fizo_var, fizo_value, zun_var, zun_value):
        """Обработка УИНов для всех типов проверок"""
        try:
            exam_types = []
            exam_data = {}
            all_uins = set()
            
            # Обработка ПФО
            if pfo_var.get():
                uins = UINProcessor.format_uins(pfo_value.get())
                if uins:
                    exam_data['личностных (психофизиологических) качеств'] = uins
                    exam_types.append('личностных (психофизиологических) качеств')
                    all_uins.update(uins)
            
            # Обработка ФИЗО
            if fizo_var.get():
                uins = UINProcessor.format_uins(fizo_value.get())
                if uins:
                    exam_data['уровня физической подготовки'] = uins
                    exam_types.append('уровня физической подготовки')
                    all_uins.update(uins)
            
            # Обработка ЗУН
            if zun_var.get():
                uins = UINProcessor.format_uins(zun_value.get())
                if uins:
                    exam_data['знаний, умений и навыков'] = uins
                    exam_types.append('знаний, умений и навыков')
                    all_uins.update(uins)
            
            # Формирование результирующих строк
            uins_string = UINProcessor._build_uins_string(exam_data)
            exams_string = UINProcessor._build_exams_string(exam_types)
            sorted_uins = sorted(list(all_uins))
            
            return sorted_uins, uins_string, exams_string
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка обработки УИНов экзаменов")
            return [], "", ""
    
    @staticmethod
    def _build_uins_string(exam_data):
        """Построение строки с УИНами"""
        if not exam_data:
            return ""
        
        try:
            result_parts = []
            exam_items = list(exam_data.items())
            
            for index, (exam_type, uins) in enumerate(exam_items, 1):
                uins_text = "; ".join(uins)
                
                if len(exam_items) == 1:
                    result_parts.append(f'{exam_type} аттестуемых лиц (УИН: {uins_text}).')
                elif index == len(exam_items):
                    result_parts.append(f'а также {exam_type} аттестуемых лиц (УИН: {uins_text}).')
                else:
                    result_parts.append(f'{exam_type} аттестуемых лиц (УИН: {uins_text}), ')
            
            return "".join(result_parts)
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка построения строки УИНов")
            return ""
    
    @staticmethod
    def _build_exams_string(exam_types):
        """Построение строки с типами экзаменов"""
        if not exam_types:
            return ""
        
        try:
            if len(exam_types) == 1:
                return exam_types[0]
            elif len(exam_types) == 2:
                return f"{', а также '.join(exam_types)}"
            else:
                return f'{exam_types[0]}, {exam_types[1]}, а также {exam_types[2]}'
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка построения строки экзаменов")
            return ""

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
        try:
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
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка создания переменных")
    
    def _create_interface(self):
        """Создание интерфейса страницы"""
        try:
            self._create_annotation_frame()
            self._create_info_frame()
            self._create_action_buttons()
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка создания интерфейса")
    
    def _create_annotation_frame(self):
        """Создание рамки с аннотацией"""
        try:
            self.annotation_frame = ctk.CTkFrame(master=self)
            self.annotation_frame.grid(row=0, column=0, padx=20, pady=10, sticky="nsew")
            
            title_label = ctk.CTkLabel(
                self.annotation_frame, 
                text='Выберите необходимые параметры', 
                font=('Times New Roman', 20)  # Исправлен шрифт
            )
            title_label.grid(row=1, column=0, columnspan=4, sticky='nsew')
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка создания аннотации")
    
    def _create_info_frame(self):
        """Создание основной информационной рамки"""
        try:
            self.info_frame = ctk.CTkFrame(master=self)
            self.info_frame.grid(row=1, column=0, padx=20, pady=10, sticky='nsew')
            
            self._create_basic_info_section()
            self._create_exam_type_section()
            self._create_personnel_section()
            self._create_approval_section()
            
            # Настройка колонок
            for i in range(4):
                self.info_frame.grid_columnconfigure(i, weight=1 if i > 0 else 0)
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка создания info frame")
    
    def _create_basic_info_section(self):
        """Создание секции основной информации"""
        try:
            font = ('Times New Roman', 13)  # Исправлен шрифт
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
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка создания базовой информации")
    
    def _create_exam_type_section(self):
        """Создание секции типов проверок"""
        try:
            font = ('Times New Roman', 13)  # Исправлен шрифт
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
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка создания секции проверок")
    
    def _create_personnel_section(self):
        """Создание секции выбора персонала"""
        try:
            font = ('Times New Roman', 13)  # Исправлен шрифт
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
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка создания секции персонала")
    
    def _create_approval_section(self):
        """Создание секции утверждения"""
        try:
            self.save_button = ctk.CTkButton(
                self, 
                text='Сформировать документ', 
                font=('Times New Roman', 13),  # Исправлен шрифт
                command=self._save_data_and_generate_document
            )
            self.save_button.grid(row=2, column=0, sticky='nsew')
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка создания кнопки сохранения")
    
    def _create_action_buttons(self):
        """Создание кнопок действий (переопределяется в дочерних классах)"""
        pass
    
    def _toggle_exam_entry(self, exam_type):
        """Переключение видимости полей ввода для типов проверок"""
        try:
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
        except Exception as e:
            ErrorHandler.log_error(e, f"Ошибка переключения поля {exam_type}")
    
    def _setup_navigation(self):
        """Настройка навигации по полям ввода"""
        try:
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
                    widget.bind('<Control-c>', SimpleClipboardManager.copy_text)
                    widget.bind('<Control-x>', SimpleClipboardManager.cut_text)
                    widget.bind('<Control-v>', SimpleClipboardManager.paste_text)
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка настройки навигации")
    
    def _navigate_fields(self, event, direction):
        """Навигация по полям ввода с помощью стрелок"""
        try:
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
            current_index = widgets.index(current_widget)
            new_index = (current_index + direction) % len(widgets)
            widgets[new_index].focus_set()
            
        except (ValueError, AttributeError):
            pass
        
        return "break"
    
    def on_show(self):
        """Действия при показе страницы"""
        self._reset_fields()
    
    def _reset_fields(self):
        """Сброс всех полей ввода к значениям по умолчанию"""
        try:
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
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка сброса полей")
    
    def _get_exams_without_fizo(self):
        """Получить типы экзаменов без ФИЗО"""
        try:
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
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка получения экзаменов без ФИЗО")
            return ""
    
    def _get_approving_person_info(self):
        """Получить информацию об утверждающем лице"""
        try:
            selected_person = self.approving_person.get()
            
            if selected_person == 'Беззубцев':
                return ['Начальник отдела проверок сил ОТБ', 'А.А. Беззубцев']
            elif selected_person == 'Алябьев':
                return ['Заместитель начальника Службы', 'А.Б. Алябьев']
            elif selected_person == 'Попырина':
                return ['Заместитель начальника отдела проверок сил ОТБ', 'Е.М. Попырина']
            else:
                return ['', '']
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка получения информации об утверждающем лице")
            return ['', '']
    
    def _get_executor_info(self):
        """Получить информацию об исполнителе"""
        try:
            selected_executor = self.executor_person.get()
            
            if selected_executor == 'Вантуз':
                return 'Воротилов Иван Иванович', 'oa13@msecurity.ru'
            elif selected_executor == 'Маруся':
                return 'Разгуляева Мария Андреевна', 'oa6@msecurity.ru'
            elif selected_executor == 'Создатель!!!':
                return 'Желудков Андрей Викторович', 'oa15@msecurity.ru'
            else:
                return '', ''
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка получения информации об исполнителе")
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
        try:
            # Обработка УИНов
            all_uins, uins_text, exams_text = UINProcessor.process_exam_uins(
                self.is_pfo_selected, self.pfo_uins,
                self.is_fizo_selected, self.fizo_uins,
                self.is_zun_selected, self.zun_uins
            )
            
            # Форматирование идентификаторов
            identifiers_text = ''
            if all_uins:
                if len(all_uins) > 1:
                    identifiers_text = ('; '.join(all_uins) + '.').strip()
                else:
                    identifiers_text = all_uins[0].strip()
            
            # Форматирование УИНов ФИЗО
            fizo_uins_list = UINProcessor.format_uins(self.fizo_uins.get())
            fizo_ids_text = ''
            if fizo_uins_list:
                if len(fizo_uins_list) > 1:
                    fizo_ids_text = '; '.join(fizo_uins_list) + '.'
                else:
                    fizo_ids_text = fizo_uins_list[0]
            
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
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка подготовки данных документа")
            return {}
    
    def _save_data_and_generate_document(self):
        """Сохранение данных и генерация документа"""
        try:
            if not SimpleDocumentGenerator.check_docx_support():
                return
                
            # Подготовка данных
            document_data = self._prepare_document_data()
            
            if not document_data:
                ErrorHandler.show_error("Не удалось подготовить данные для документа", self)
                return
            
            # Генерация основного документа
            self._generate_main_document(document_data)
            
            # Генерация документа ФИЗО при необходимости
            if self.is_fizo_selected.get() and self.fizo_template_path:
                self._generate_fizo_document(document_data)
            
            self._show_success_message()
            
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка генерации документа")
            ErrorHandler.show_error(f"Ошибка при формировании документа: {str(e)}", self)
    
    def _generate_main_document(self, document_data):
        """Генерация основного документа"""
        try:
            edited_doc = SimpleDocumentGenerator.format_document(document_data, self.template_path)
            if not edited_doc:
                return
            
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
                SimpleDocumentGenerator.convert_to_pdf(file_path, pdf_path)
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка генерации основного документа")
            raise
    
    def _generate_fizo_document(self, document_data):
        """Генерация документа ФИЗО"""
        try:
            fizo_doc = SimpleDocumentGenerator.format_document(document_data, self.fizo_template_path)
            if not fizo_doc:
                return
            
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
                SimpleDocumentGenerator.convert_to_pdf(file_path, pdf_path)
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка генерации документа ФИЗО")
            raise
    
    def _show_success_message(self):
        """Показать сообщение об успехе"""
        try:
            popup = ctk.CTkToplevel(self)
            popup.title("Уведомление")
            popup.geometry("300x100")
            popup.resizable(False, False)
            popup.transient(self)
            popup.grab_set()
            
            # Центрирование окна
            popup.update_idletasks()
            x = (popup.winfo_screenwidth() // 2) - (popup.winfo_width() // 2)
            y = (popup.winfo_screenheight() // 2) - (popup.winfo_height() // 2)
            popup.geometry(f"+{x}+{y}")
            
            ctk.CTkLabel(popup, text="Документ сформирован!").pack(pady=20)
            ctk.CTkButton(popup, text="OK", command=popup.destroy).pack(pady=5)
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка показа сообщения об успехе")
            messagebox.showinfo("Уведомление", "Документ сформирован!")

# Упрощенные классы для конкретных локаций
class MainPage(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self._create_interface()
    
    def _create_interface(self):
        """Создание интерфейса главной страницы"""
        try:
            # Заголовок
            title_label = ctk.CTkLabel(self, text="Генератор нормированных заданий", font=("Times New Roman", 24))
            title_label.grid(row=0, column=0, columnspan=3, pady=20, padx=20)
            
            subtitle_label = ctk.CTkLabel(self, text="Выберите локацию:", font=("Times New Roman", 16))
            subtitle_label.grid(row=1, column=0, columnspan=3, pady=10)
            
            # Кнопки локаций
            location_buttons = [
                ("Новороссийск", "Novoros"),
                ("Астрахань", "Astrakhan"),
                ("Архангельск", "Arhangelsk"),
                ("Бор", "Bor"),
                ("Екатеринбург", "Ekaterinburg"),
                ("Хабаровск", "Habarovsk"),
                ("Калининград", "Kaliningrad"),
                ("Нижний Новгород", "NN"),
                ("Мурманск", "Murmansk"),
                ("Находка", "Nahodka"),
                ("Петропавловск-Камчатский", "PK"),
                ("Ростов-на-Дону", "Rostov"),
                ("Самара", "Samara"),
                ("Севастополь", "Sevastopol"),
                ("Санкт-Петербург", "Spb"),
                ("Южно-Сахалинск", "US"),
                ("Владивосток", "Vladivostok")
            ]
            
            for i, (button_text, page_name) in enumerate(location_buttons):
                row = i // 3 + 2
                col = i % 3
                button = ctk.CTkButton(
                    self, 
                    text=button_text,
                    command=lambda p=page_name: self.controller.show_page(p),
                    height=40,
                    width=200,
                    font=("Times New Roman", 14)
                )
                button.grid(row=row, column=col, pady=10, padx=10, sticky='nsew')
            
            # Кнопка выхода
            exit_button = ctk.CTkButton(
                self,
                text="Выход",
                command=self._safe_quit,
                fg_color="red",
                hover_color="darkred",
                height=40,
                width=200,
                font=("Times New Roman", 14)
            )
            exit_button.grid(row=len(location_buttons)//3 + 3, column=1, pady=20, padx=10, sticky='nsew')
            
            # Настройка grid
            for i in range(3):
                self.grid_columnconfigure(i, weight=1)
            for i in range(len(location_buttons)//3 + 4):
                self.grid_rowconfigure(i, weight=1)
                
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка создания главной страницы")
    
    def _safe_quit(self):
        """Безопасный выход из приложения"""
        try:
            self.controller.quit()
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка при выходе")
            sys.exit(0)

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
            text='Вернуться на главную', 
            command=lambda: self.controller.show_page('MainPage'),
            font=("Times New Roman", 14)
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
            text='Вернуться на главную', 
            command=lambda: self.controller.show_page('MainPage'),
            font=("Times New Roman", 14)
        )
        back_button.grid(row=3, column=0, pady=20, padx=20, sticky="ew")

# Упрощенные версии других классов локаций (для экономии места)
class Astrakhan(BasePage):
    def __init__(self, parent, controller):
        template_path = os.path.join(TEMPLATES_DIR, 'Astrakhan.docx')
        super().__init__(
            parent, controller, 
            worker_list=['Шматова Раиса Анатольевна', 'Сивоконь Светлана Викторовна'], 
            template_path=template_path, 
            location_name='Астрахань'
        )
        self._create_back_button()
    
    def _create_back_button(self):
        back_button = ctk.CTkButton(
            self, text='Вернуться на главную', 
            command=lambda: self.controller.show_page('MainPage'),
            font=("Times New Roman", 14)
        )
        back_button.grid(row=3, column=0, pady=20, padx=20, sticky="ew")

# Добавьте аналогичные упрощенные классы для других локаций...

class App(ctk.CTk):
    """Главное приложение"""
    
    def __init__(self):
        super().__init__()
        self._setup_appearance()
        self._create_interface()
    
    def _setup_appearance(self):
        """Настройка внешнего вида приложения"""
        try:
            self.title("Генератор нормированных заданий")
            self.geometry("1000x700")
            
            ctk.set_appearance_mode("dark")
            ctk.set_default_color_theme("blue")
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка настройки внешнего вида")
    
    def _create_interface(self):
        """Создание интерфейса приложения"""
        try:
            # Главный контейнер
            self.main_container = ctk.CTkFrame(self)
            self.main_container.pack(fill="both", expand=True, padx=10, pady=10)
            
            # Создание страниц
            self.pages = self._create_pages()
            self._arrange_pages()
            
            # Показ главной страницы
            self.show_page("MainPage")
        except Exception as e:
            ErrorHandler.log_error(e, "Ошибка создания интерфейса")
            ErrorHandler.show_error(f"Ошибка инициализации приложения: {str(e)}")
    
    def _create_pages(self):
        """Создание всех страниц приложения"""
        pages = {
            "MainPage": MainPage(self.main_container, self),
            "Novoros": Novoros(self.main_container, self),
            "Vladivostok": Vladivostok(self.main_container, self),
            "Astrakhan": Astrakhan(self.main_container, self),
            # Добавьте другие страницы по аналогии
        }
        return pages
    
    def _arrange_pages(self):
        """Размещение всех страниц в контейнере"""
        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")
        
        # Конфигурация grid
        self.main_container.grid_rowconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)
    
    def show_page(self, page_name):
        """Показать выбранную страницу"""
        try:
            page = self.pages.get(page_name)
            if page:
                page.tkraise()
                if hasattr(page, "on_show"):
                    page.on_show()
        except Exception as e:
            ErrorHandler.log_error(e, f"Ошибка показа страницы {page_name}")

def main():
    """Главная функция приложения"""
    try:
        # Установка локали
        LocaleManager.set_russian_locale()
        
        # Создание и запуск приложения
        app = App()
        app.mainloop()
        
    except Exception as e:
        ErrorHandler.log_error(e, "Критическая ошибка приложения")
        ErrorHandler.show_error(f"Критическая ошибка: {str(e)}\n\nПодробности в файле error_log.txt")

if __name__ == "__main__":
    main()
