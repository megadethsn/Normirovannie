import customtkinter as ctk
from datetime import datetime
import locale
from tkinter import PhotoImage, filedialog, messagebox
import os

from app_logging import get_logger, setup_logging
from app_paths import resolve_asset, templates_dir
from centers import load_centers, normalized_center, save_centers
from date_utils import MONTH_NAMES, format_ru_date, next_work_date
from document_service import cleanup_word_apps, convert_to_pdf, format_docx
from people import load_people, save_people
from uins import format_uins

setup_logging()
logger = get_logger(__name__)
TEMPLATES_DIR = templates_dir()

def set_russian_locale():
    try:
        # Пробуем разные варианты для Windows
        locale_options = ['ru_RU', 'Russian', 'Russian_Russia.1251', 'rus']
        for loc in locale_options:
            try:
                locale.setlocale(locale.LC_TIME, loc)
                logger.info("Установлена локаль: %s", loc)
                return
            except locale.Error:
                continue
        logger.warning("Не удалось установить русскую локаль, используем системную")
    except Exception as e:
        logger.exception("Ошибка установки локали: %s", e)

# Устанавливаем локаль
set_russian_locale()

# Настройка внешнего вида
ctk.set_appearance_mode("dark")  # Режим: "light", "dark", "system"
ctk.set_default_color_theme("blue")  # Темы: "blue", "green", "dark-blue"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Генератор нормированных заданий")
        self.geometry("1000x700")  
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.configure_app_icon()
        self.centers = load_centers()
        self.people = load_people()
        
        # Главный контейнер для страниц
        self.container = ctk.CTkFrame(self)
        self.container.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.rebuild_pages()
        
        # Конфигурация grid для контейнера
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)
        
        # Показываем главную страницу
        self.show_page("MainPage")

    def rebuild_pages(self):
        if hasattr(self, "pages"):
            for page in self.pages.values():
                page.destroy()

        self.pages = {
            "MainPage": MainPage(self.container, self),
            "ManageCenters": ManageCentersPage(self.container, self),
            "Settings": SettingsPage(self.container, self),
        }
        for center in self.centers:
            page_name = center.get("id")
            if page_name:
                self.pages[page_name] = BasePage(self.container, self, **normalized_center(center))

        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")

    def save_centers(self, centers):
        self.centers = centers
        save_centers(centers)
        self.rebuild_pages()

    def save_people(self, people):
        self.people = people
        save_people(people)
        self.rebuild_pages()

    def configure_app_icon(self):
        icon_path = resolve_asset("msecurity_logo.ico")
        png_path = resolve_asset("msecurity_logo.png")
        try:
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
                return
        except Exception:
            logger.exception("Не удалось установить ICO приложения")
        try:
            if os.path.exists(png_path):
                self._app_icon = PhotoImage(file=png_path)
                self.iconphoto(True, self._app_icon)
        except Exception:
            logger.exception("Не удалось установить PNG-иконку приложения")
    
    def show_page(self, page_name):
        """Показать выбранную страницу"""
        page = self.pages.get(page_name)
        if page:
            page.tkraise()
            if hasattr(page, "on_show"):
                page.on_show()

    def on_close(self):
        logger.info("Закрытие приложения")
        try:
            cleanup_word_apps()
        except Exception as error:
            logger.exception("Ошибка при закрытии Word COM-объектов: %s", error)
        try:
            self.quit()
            self.destroy()
        finally:
            os._exit(0)

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
            joined_string = '\n'.join(self.selected)
            return f"\n{joined_string}\n"
        else:
            return "\n".join(self.selected)

class BasePage(ctk.CTkFrame):
    """Базовый класс для всех страниц"""
    def __init__(
        self,
        parent,
        controller,
        id="",
        button="",
        name="",
        template="",
        template_fizo="",
        template_path="",
        template_fizo_path="",
        east=False,
        workers=None,
    ):
        super().__init__(parent)
        self.controller = controller
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1) 
        self.font = ('TimesNewRoman', 13)
        self.center_id = id
        self.template = template_path or template
        self.template_fizo_path = template_fizo_path
        self.name = name
        self.east = east
        self.results = {}
        self.month_names = MONTH_NAMES
        
        self.header_frame = ctk.CTkFrame(master=self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, padx=22, pady=(18, 8), sticky="ew")
        self.header_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            self.header_frame,
            text=self.name,
            font=('TimesNewRoman', 24)
        ).grid(row=0, column=0, sticky='w')
        ctk.CTkLabel(
            self.header_frame,
            text='Формирование нормированного задания и заявки ФИЗО',
            font=('TimesNewRoman', 14),
            text_color='gray70'
        ).grid(row=1, column=0, sticky='w')

        self.info_frame = ctk.CTkScrollableFrame(master=self)
        self.info_frame.grid(row=1, column=0, padx=22, pady=8, sticky='nsew')
        self.info_frame.grid_columnconfigure(0, weight=1)

        request_frame = self.create_section(self.info_frame, 'Заявка и даты', 0)
        request_frame.grid_columnconfigure(1, weight=1)
        request_frame.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(request_frame, text='Номер заявки:', font=self.font).grid(row=0, column=0, padx=10, pady=8, sticky='w')

        self.num_var = ctk.StringVar()
        self.num_entry = ctk.CTkEntry(request_frame, width=260, textvariable=self.num_var, placeholder_text='Номер заявки')
        self.num_entry.grid(row=0, column=1, columnspan=3, padx=10, pady=8, sticky='ew')

        ctk.CTkLabel(request_frame, text='Дата формирования:', font=self.font).grid(row=1, column=0, padx=10, pady=8, sticky='w')
        self.ISSUE_DATE = ctk.StringVar()
        self.issue_date_entry = ctk.CTkEntry(request_frame, width=260, textvariable=self.ISSUE_DATE)
        self.issue_date_entry.insert(0, format_ru_date(datetime.now(), quoted=True))
        self.issue_date_entry.grid(row=1, column=1, padx=10, pady=8, sticky='ew')

        ctk.CTkLabel(request_frame, text='Дата проверки:', font=self.font).grid(row=1, column=2, padx=10, pady=8, sticky='w')
        self.WORK_DATE = ctk.StringVar()
        self.work_date_entry = ctk.CTkEntry(request_frame, width=260, textvariable=self.WORK_DATE)
        self.work_date_entry.insert(0, format_ru_date(self.work_date(east=self.east)))
        self.work_date_entry.grid(row=1, column=3, padx=10, pady=8, sticky='ew')

        ctk.CTkLabel(request_frame, text='Время проверки:', font=self.font).grid(row=2, column=0, padx=10, pady=8, sticky='w')
        time_frame = ctk.CTkFrame(request_frame, fg_color="transparent")
        time_frame.grid(row=2, column=1, padx=10, pady=8, sticky='w')
        self.TIME_START = ctk.StringVar()
        self.entry_start_time = ctk.CTkEntry(time_frame, width=90, textvariable=self.TIME_START)
        self.entry_start_time.insert(0, '10:00')
        self.entry_start_time.grid(row=0, column=0, padx=(0, 8), sticky='w')
        ctk.CTkLabel(time_frame, text='-').grid(row=0, column=1, padx=(0, 8))
        self.TIME_END = ctk.StringVar()
        self.entry_end_time = ctk.CTkEntry(time_frame, width=90, textvariable=self.TIME_END)
        self.entry_end_time.insert(0, '16:00')
        self.entry_end_time.grid(row=0, column=2, sticky='w')

        exams_frame = self.create_section(self.info_frame, 'Проверки и УИН', 1)
        exams_frame.grid_columnconfigure(1, weight=1)
        exams_frame.grid_columnconfigure(3, weight=1)

        self.pfo_var = ctk.BooleanVar(value=False)
        self.fizo_var = ctk.BooleanVar(value=False)
        self.zun_var = ctk.BooleanVar(value=False)

        self.pfo_value = ctk.StringVar()
        self.fizo_value = ctk.StringVar()
        self.zun_value = ctk.StringVar()

        self.cb_pfo = ctk.CTkCheckBox(exams_frame, text='ПФО', variable=self.pfo_var, command=lambda: self.toggle_entry('pfo'))
        self.cb_pfo.grid(row=0, column=0, padx=10, pady=8, sticky='w')
        
        self.cb_fizo = ctk.CTkCheckBox(exams_frame, text='ФИЗО', variable=self.fizo_var, command=lambda: self.toggle_entry('fizo'))
        self.cb_fizo.grid(row=1, column=0, padx=10, pady=8, sticky='w')
        

        self.cb_zun = ctk.CTkCheckBox(exams_frame, text='ЗУН', variable=self.zun_var, command=lambda: self.toggle_entry('zun'))
        self.cb_zun.grid(row=2, column=0, padx=10, pady=8, sticky='w')

        self.entry_pfo = ctk.CTkEntry(exams_frame, textvariable=self.pfo_value, placeholder_text='Например: 3580-3585')
        self.entry_fizo = ctk.CTkEntry(exams_frame, textvariable=self.fizo_value, placeholder_text='Например: п3580-п3585')
        self.entry_zun = ctk.CTkEntry(exams_frame, textvariable=self.zun_value, placeholder_text='Например: 3580, 3582-3585')

        self.fizo_number_label = ctk.CTkLabel(exams_frame, text='Номер заявки ФИЗО:', font=self.font)
        self.fizo_number_var = ctk.StringVar()
        self.fizo_number_entry = ctk.CTkEntry(exams_frame, textvariable=self.fizo_number_var)

        self.uins_preview_var = ctk.StringVar(value='УИНы появятся здесь после выбора проверки')
        self.uins_preview_label = ctk.CTkLabel(
            exams_frame,
            textvariable=self.uins_preview_var,
            font=('TimesNewRoman', 12),
            text_color='gray75',
            justify='left',
            anchor='w',
            wraplength=760,
        )
        self.uins_preview_label.grid(row=3, column=0, columnspan=4, padx=10, pady=(8, 4), sticky='ew')

        people_frame = self.create_section(self.info_frame, 'Участники', 2)
        people_frame.grid_columnconfigure(1, weight=1)
        people_frame.grid_columnconfigure(2, weight=1)
        people_frame.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(people_frame, text='Сотрудники:', font=self.font).grid(row=0, column=0, padx=10, pady=8, sticky='w')
        
        self.worker_list = workers or []
        self.workers_dropdown = MultiSelectDropdown(people_frame, self.worker_list)
        self.workers_dropdown.grid(row=0, column=1, columnspan=3, padx=10, pady=8, sticky='ew')

        ctk.CTkLabel(people_frame, text='Должностное лицо:', font=self.font).grid(row=1, column=0, padx=10, pady=8, sticky='w')

        self.chief_var = ctk.StringVar(value='')
        self.chief_buttons = []
        for index, chief in enumerate(self.controller.people.get("chiefs", [])):
            button = ctk.CTkRadioButton(
                people_frame,
                text=chief.get("label", chief.get("fullname", "")),
                font=self.font,
                variable=self.chief_var,
                value=chief.get("id", ""),
            )
            button.grid(row=1 + index // 3, column=1 + index % 3, padx=8, pady=6, sticky='w')
            self.chief_buttons.append(button)

        issuer_start_row = 2 + max(0, (len(self.chief_buttons) - 1) // 3)
        ctk.CTkLabel(people_frame, text='Исполнитель:', font=self.font).grid(row=issuer_start_row, column=0, padx=10, pady=8, sticky='w')
        self.issuer = ctk.StringVar(value='')
        self.issuer_buttons = []
        for index, issuer in enumerate(self.controller.people.get("issuers", [])):
            button = ctk.CTkRadioButton(
                people_frame,
                text=issuer.get("label", issuer.get("fullname", "")),
                font=self.font,
                variable=self.issuer,
                value=issuer.get("id", ""),
            )
            button.grid(row=issuer_start_row + index // 3, column=1 + index % 3, padx=8, pady=6, sticky='w')
            self.issuer_buttons.append(button)

        actions_frame = ctk.CTkFrame(master=self, fg_color="transparent")
        actions_frame.grid(row=2, column=0, padx=22, pady=(8, 18), sticky='ew')
        actions_frame.grid_columnconfigure(0, weight=1)
        actions_frame.grid_columnconfigure(1, weight=1)

        self.btn_save = ctk.CTkButton(actions_frame, text='Сформировать документ', font=self.font, command=self.save_data_and_form_doc, height=42)
        self.btn_save.grid(row=0, column=0, padx=(0, 8), sticky='ew')

        ctk.CTkButton(
            actions_frame,
            text='Вернуться назад',
            command=lambda: controller.show_page('MainPage'),
            height=42,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        for entry in (self.entry_pfo, self.entry_fizo, self.entry_zun):
            entry.bind('<KeyRelease>', self.update_uins_preview)
            inner = getattr(entry, "_entry", None)
            if inner:
                inner.bind('<KeyRelease>', self.update_uins_preview)

        # Настройка навигации по Tab и стрелкам
        self.setup_navigation()

    def create_section(self, parent, title, row):
        section = ctk.CTkFrame(parent)
        section.grid(row=row, column=0, padx=4, pady=8, sticky='ew')
        section.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            section,
            text=title,
            font=('TimesNewRoman', 18)
        ).grid(row=0, column=0, columnspan=4, padx=10, pady=(10, 6), sticky='w')
        content = ctk.CTkFrame(section, fg_color="transparent")
        content.grid(row=1, column=0, padx=0, pady=(0, 8), sticky='ew')
        content.grid_columnconfigure(0, weight=0)
        return content

    def setup_navigation(self):
        """Настройка навигации - упрощенная версия"""
        # Создаем список всех виджетов в правильном порядке
        self.navigation_widgets = [
            self.num_entry,
            self.issue_date_entry, 
            self.work_date_entry,
            self.entry_start_time,
            self.entry_end_time,
            self.cb_pfo,
            self.cb_fizo,
            self.cb_zun,
            self.entry_pfo,
            self.entry_fizo,
            self.fizo_number_entry,
            self.entry_zun,
            self.workers_dropdown.toggle_btn,
            *self.chief_buttons,
            *self.issuer_buttons,
            self.btn_save
        ]

        # Привязываем события к каждому виджету
        for widget in self.navigation_widgets:
            if widget:  # Проверяем что виджет существует
                widget.bind('<Up>', self.on_arrow_up)
                widget.bind('<Down>', self.on_arrow_down)
                inner = getattr(widget, "_entry", None)
                if inner:
                    inner.bind('<Up>', self.on_arrow_up)
                    inner.bind('<Down>', self.on_arrow_down)
        
        # Отдельно привязываем буфер обмена к полям ввода
        entries = [
            self.num_entry, self.issue_date_entry, self.work_date_entry,
            self.entry_start_time, self.entry_end_time, self.entry_pfo,
            self.entry_fizo, self.entry_zun, self.fizo_number_entry,
        ]
        for entry in entries:
            entry.bind('<Control-KeyPress>', self.on_control_key)
            inner = getattr(entry, "_entry", None)
            if inner:
                inner.bind('<Control-KeyPress>', self.on_control_key)

    def on_arrow_up(self, event):
        """Обработка стрелки вверх"""
        self.navigate(-1)
        return "break"

    def on_arrow_down(self, event):
        """Обработка стрелки вниз"""
        self.navigate(1)
        return "break"

    def navigate(self, direction):
        """Навигация между виджетами"""
        try:
            current_focus = self.focus_get()
            current_widget = current_focus

            for widget in self.navigation_widgets:
                if current_focus in (widget, getattr(widget, "_entry", None)):
                    current_widget = widget
                    break
            
            if current_widget in self.navigation_widgets:
                current_index = self.navigation_widgets.index(current_widget)
                next_widget = self.next_focusable_widget(current_index, direction)
                next_widget.focus_set()
                self.scroll_to_widget(next_widget)
                
            else:
                next_widget = self.next_focusable_widget(-1, 1)
                next_widget.focus_set()
                self.scroll_to_widget(next_widget)
                
        except Exception as e:
            logger.exception("Ошибка навигации: %s", e)

    def next_focusable_widget(self, current_index, direction):
        total = len(self.navigation_widgets)
        for step in range(1, total + 1):
            index = (current_index + direction * step) % total
            widget = self.navigation_widgets[index]
            if self.is_focusable(widget):
                return widget
        return self.num_entry

    def is_focusable(self, widget):
        if widget == self.entry_pfo:
            return self.pfo_var.get()
        if widget == self.entry_fizo:
            return self.fizo_var.get()
        if widget == self.fizo_number_entry:
            return self.fizo_var.get() and bool(self.template_fizo_path)
        if widget == self.entry_zun:
            return self.zun_var.get()
        try:
            return bool(widget.winfo_exists())
        except Exception:
            return False

    def scroll_to_widget(self, widget):
        try:
            self.after(10, lambda: widget._entry.focus_set() if hasattr(widget, "_entry") else widget.focus_set())
        except Exception:
            pass

    def on_control_key(self, event):
        key = (getattr(event, "keysym", "") or "").lower()
        keycode = getattr(event, "keycode", None)
        if key in ("c", "с", "cyrillic_es") or keycode == 67:
            return self.on_copy(event)
        if key in ("x", "ч", "cyrillic_che") or keycode == 88:
            return self.on_cut(event)
        if key in ("v", "м", "cyrillic_em") or keycode == 86:
            return self.on_paste(event)
        if key in ("a", "ф", "cyrillic_ef") or keycode == 65:
            return self.on_select_all(event)

    def on_copy(self, event):
        """Копирование"""
        try:
            widget = self._entry_widget(event.widget)
            try:
                selected_text = widget.selection_get()
            except Exception:
                selected_text = ''
            if selected_text:
                self.clipboard_clear()
                self.clipboard_append(selected_text)
        except Exception as e:
            logger.exception("Ошибка копирования: %s", e)
        return "break"

    def on_cut(self, event):
        """Вырезание"""
        try:
            widget = self._entry_widget(event.widget)
            try:
                selected_text = widget.selection_get()
            except Exception:
                selected_text = ''
            if selected_text:
                self.clipboard_clear()
                self.clipboard_append(selected_text)
                widget.delete("sel.first", "sel.last")
        except Exception as e:
            logger.exception("Ошибка вырезания: %s", e)
        return "break"

    def on_paste(self, event):
        """Вставка"""
        try:
            widget = self._entry_widget(event.widget)
            text = self.clipboard_get()
            try:
                widget.delete("sel.first", "sel.last")
            except Exception:
                pass
            widget.insert("insert", text)
        except Exception as e:
            logger.exception("Ошибка вставки: %s", e)
        return "break"

    def on_select_all(self, event):
        try:
            widget = self._entry_widget(event.widget)
            widget.selection_range(0, "end")
            widget.icursor("end")
        except Exception as e:
            logger.exception("Ошибка выделения текста: %s", e)
        return "break"

    def _entry_widget(self, widget):
        if hasattr(widget, "_entry"):
            return widget._entry
        return widget

    def on_show(self):
        """Сброс полей при показе страницы"""
        self.reset_fields()

    def reset_fields(self):
        """Сброс всех полей ввода"""
        self.num_var.set('')
        self.ISSUE_DATE.set('')
        now = datetime.now()
        formatted_date = format_ru_date(now, quoted=True)
        self.issue_date_entry.delete(0, 'end')
        self.issue_date_entry.insert(0, formatted_date)
        
        self.WORK_DATE.set('')
        next_day = self.work_date(east=self.east)
        formatted_date_next = format_ru_date(next_day)
        self.work_date_entry.delete(0, 'end')
        self.work_date_entry.insert(0, formatted_date_next)
        
        self.TIME_START.set('10:00')
        self.TIME_END.set('16:00')
        
        self.pfo_var.set(False)
        self.fizo_var.set(False)
        self.zun_var.set(False)
        self.pfo_value.set('')
        self.fizo_value.set('')
        self.zun_value.set('')
        self.fizo_number_var.set('')
        self.update_uins_preview()
        
        self.entry_pfo.grid_forget()
        self.entry_fizo.grid_forget()
        self.entry_zun.grid_forget()
        self.fizo_number_label.grid_forget()
        self.fizo_number_entry.grid_forget()
        
        # Сброс выбора сотрудников
        for value in self.worker_list:
            if value in self.workers_dropdown.check_vars:
                self.workers_dropdown.check_vars[value].set(False)
        self.workers_dropdown.selected = []
        self.workers_dropdown.update_button_text()
        
        self.chief_var.set('')
        self.issuer.set('')

    #Функция отображения поля ввода по чекбоксам ПФО, ФИЗО, ЗУН
    def toggle_entry(self, checkbox):
        if checkbox == 'pfo':
            if self.pfo_var.get():
                self.entry_pfo.grid(row=0, column=1, columnspan=3, padx=10, pady=8, sticky='ew')
            else:
                self.entry_pfo.grid_forget()
        elif checkbox == 'fizo':
            if self.fizo_var.get():
                if self.template_fizo_path:
                    self.entry_fizo.grid(row=1, column=1, padx=10, pady=8, sticky='ew')
                    self.fizo_number_label.grid(row=1, column=2, padx=10, pady=8, sticky='e')
                    self.fizo_number_entry.grid(row=1, column=3, padx=10, pady=8, sticky='ew')
                else:
                    self.entry_fizo.grid(row=1, column=1, columnspan=3, padx=10, pady=8, sticky='ew')
                    self.fizo_number_label.grid_forget()
                    self.fizo_number_entry.grid_forget()
            else:
                self.entry_fizo.grid_forget()
                self.fizo_number_label.grid_forget()
                self.fizo_number_entry.grid_forget()
        elif checkbox == 'zun':
            if self.zun_var.get():
                self.entry_zun.grid(row=2, column=1, columnspan=3, padx=10, pady=8, sticky='ew')
            else:
                self.entry_zun.grid_forget()
        self.update_uins_preview()

    def update_uins_preview(self, event=None):
        try:
            preview_parts = []
            if self.pfo_var.get():
                preview_parts.extend(self.formate_uins(self.pfo_value.get()))
            if self.fizo_var.get():
                preview_parts.extend(self.formate_uins(self.fizo_value.get()))
            if self.zun_var.get():
                preview_parts.extend(self.formate_uins(self.zun_value.get()))
            if preview_parts:
                preview = '; '.join(preview_parts)
                if len(preview) > 420:
                    preview = preview[:420] + '...'
                self.uins_preview_var.set(f'Предпросмотр УИН: {preview}')
            else:
                self.uins_preview_var.set('УИНы появятся здесь после выбора проверки')
        except ValueError as error:
            self.uins_preview_var.set(f'Ошибка УИН: {error}')
    
    #Экзамены без физры первая страница
    def get_exams(self):
        res = []

        if self.pfo_var.get():
            res.append('личностных (психофизиологических) качеств')
        if self.zun_var.get():
            res.append('знаний, умений и навыков')
        
        if len(res) == 1:
                return res[0]
        if len(res) == 2:
                return ', а также '.join(res)


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
                exams_string =  f"{', а также '.join(type_of_exams)}"
        elif len(type_of_exams) == 3:
                exams_string = f'{type_of_exams[0]}, {type_of_exams[1]}, а также {type_of_exams[2]}'
        
        total_uins = sorted(list(total_uins))
        return total_uins, result_string, exams_string


        
    #Функция сбора и формирования уинов
    def formate_uins(self, string):
        return format_uins(string)
    
    #Чекбоксы подписанты
    def get_result_chief(self):
        selected_id = self.chief_var.get()
        for chief in self.controller.people.get("chiefs", []):
            if chief.get("id") == selected_id:
                return [chief.get("job_title", ""), chief.get("fullname", "")]
        return ["", ""]
    
    #Функция чекбоксы исполнители
    def get_result_issuer(self):
        selected_id = self.issuer.get()
        for issuer in self.controller.people.get("issuers", []):
            if issuer.get("id") == selected_id:
                return issuer.get("fullname", ""), issuer.get("email", "")
        return '', ''


    #Функция определения следующей даты
    def work_date(self, east=False):
        return next_work_date(east=east)
    
    #Функция расчета затраченого времени
    def est_time(self):
        try:
            start = int(self.TIME_START.get().split(':')[0])
            end = int(self.TIME_END.get().split(':')[0])
            res = end - start
            if self.fizo_var.get() and self.template_fizo_path:
                res += 1
            return f'{res} часов' if res >= 5 else f'{res} часа'
        except:
            return '6 часов'
    
    def convert_to_pdf(self, docx_path, pdf_path):
        return convert_to_pdf(docx_path, pdf_path)

    #Функция замены меток в документах
    def formate_docx(self, replacements_dict, template):
        return format_docx(replacements_dict, template)

    def validate_form(self):
        errors = []
        if not self.num_var.get().strip():
            errors.append("Укажите номер заявки")
        if not (self.pfo_var.get() or self.fizo_var.get() or self.zun_var.get()):
            errors.append("Выберите хотя бы один тип проверки")
        if self.pfo_var.get() and not self.pfo_value.get().strip():
            errors.append("Укажите УИН для ПФО")
        if self.fizo_var.get() and not self.fizo_value.get().strip():
            errors.append("Укажите УИН для ФИЗО")
        if self.zun_var.get() and not self.zun_value.get().strip():
            errors.append("Укажите УИН для ЗУН")
        if self.fizo_var.get() and self.template_fizo_path and not self.fizo_number_var.get().strip():
            errors.append("Укажите номер заявки ФИЗО")
        if not self.workers_dropdown.selected:
            errors.append("Выберите сотрудников центра")
        if not self.chief_var.get():
            errors.append("Выберите должностное лицо")
        if not self.issuer.get():
            errors.append("Выберите исполнителя")
        if not os.path.exists(self.template):
            errors.append(f"Основной шаблон не найден: {self.template}")
        if self.fizo_var.get() and self.template_fizo_path and not os.path.exists(self.template_fizo_path):
            errors.append(f"Шаблон ФИЗО не найден: {self.template_fizo_path}")
        return errors
            
        

    def show_simple_popup(self, title, message):
        """Простое всплывающее окно с авто-закрытием через 2 секунды"""
        try:
            popup = ctk.CTkToplevel(self)
            popup.title(title)
            popup.geometry("400x120")
            popup.resizable(False, False)
            popup.transient(self)
            popup.grab_set()
            
            # Центрирование
            popup.update_idletasks()
            x = (popup.winfo_screenwidth() // 2) - (400 // 2)
            y = (popup.winfo_screenheight() // 2) - (120 // 2)
            popup.geometry(f"400x120+{x}+{y}")
            
            label = ctk.CTkLabel(popup, text=message, wraplength=380)
            label.pack(pady=30, padx=10)
            
            # Автоматическое закрытие через 2 секунды
            popup.after(2000, popup.destroy)
            
        except Exception as e:
            logger.exception("Ошибка создания popup: %s", e)

    def save_data_and_form_doc(self):
        time = f'с {self.TIME_START.get()} час. до {self.TIME_END.get()} час.'
        validation_errors = self.validate_form()
        if validation_errors:
            self.show_simple_popup("Проверьте поля", "\n".join(validation_errors[:4]))
            logger.info("Формирование остановлено: %s", "; ".join(validation_errors))
            return

        try:
            total_uins, exams_str, types_str = self.get_uins()
            fizo_uins = self.formate_uins(self.fizo_value.get())
        except ValueError as error:
            self.show_simple_popup("Ошибка", str(error))
            return

        if not total_uins:
            self.show_simple_popup("Ошибка", "Выберите тип проверки и укажите УИН")
            return

        job_title, fullname = self.get_result_chief()
        issuer, email = self.get_result_issuer()
        ids = ('; '.join(total_uins) + '.').strip() if len(total_uins) > 1 else total_uins[0].strip()
        fizo_ids = ('; '.join(fizo_uins) + '.') if len(fizo_uins) > 1 else (fizo_uins[0] if fizo_uins else '')
        
        # Получаем номер заявки ФИЗО, если есть
        fizo_number = self.fizo_number_var.get() if hasattr(self, 'fizo_number_var') and self.fizo_var.get() else ''

        self.results = {
            '{{NUMBER}}': self.num_var.get(),
            '{{NUMBER_FIZO}}': fizo_number,
            '{{DATE_OF_ISSUE}}': self.ISSUE_DATE.get().lower(),
            '{{DATE_TO_WORK}}': self.WORK_DATE.get().lower(),
            '{{TYPE_OF_EXAMS}}': types_str,
            '{{IDENTIFICATORS}}': ids,
            '{{EXAMS_P2}}': exams_str,
            '{{TIME}}': time,
            '{{EST_TIME}}': self.est_time(),
            '{{PEOPLE}}': self.workers_dropdown.get_selected(),
            '{{JOB_TITLE}}': job_title,
            '{{FULLNAME}}': fullname,
            '{{ISSUER}}': issuer,
            '{{EMAIL}}': email,
            '{{EXAMS}}': self.get_exams(),
            '{{FIZO_UIN}}':fizo_ids
        }
        
        try:
            # Проверяем существование основного шаблона
            if not os.path.exists(self.template):
                self.show_simple_popup("Ошибка", f"Основной шаблон не найден: {self.template}")
                return
            
            # Создаем основной документ
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
            
            # Сохраняем основной DOCX
            edited_doc.save(file_path)
            
            # Конвертируем основной в PDF
            pdf_file_path = file_path.replace('.docx', '.pdf')
            pdf_success = self.convert_to_pdf(file_path, pdf_file_path)
            
            # Показываем результат для основного документа
            if pdf_success:
                self.show_simple_popup("Успешно", "Документ и PDF сформированы")
            else:
                self.show_simple_popup("Успешно", "Документ сформирован")
            
            if self.fizo_var.get():
                logger.info("Создание документа ФИЗО для центра %s", self.name)
                self.create_fizo_document()
                
            # Финальное уведомление
            self.show_simple_popup("Готово", "Документы сформированы")
                
        except Exception as e:
            logger.exception("Ошибка при формировании документа: %s", e)
            self.show_simple_popup("Ошибка", "Ошибка при формировании документа")

    def create_fizo_document(self):
        """Создание документа ФИЗО с проверкой данных"""
        try:
            if not self.template_fizo_path:
                logger.info("ФИЗО выбран, но путь к шаблону не указан")
                return
                
            if not os.path.exists(self.template_fizo_path):
                logger.error("Шаблон ФИЗО не найден: %s", self.template_fizo_path)
                self.show_simple_popup("Ошибка", "Шаблон ФИЗО не найден")
                return
            
            clean_results = {}
            for key, value in self.results.items():
                if isinstance(value, list):
                    clean_value = ', '.join(str(item) for item in value)
                elif value is None:
                    clean_value = ''
                else:
                    clean_value = str(value)
                clean_results[key] = clean_value
            
            fizo_doc = self.formate_docx(clean_results, self.template_fizo_path)
            fizo_default_filename = f'Заявка ФИЗО на {self.work_date(east=self.east).strftime("%d.%m")} {self.name}'
            
            fizo_file_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
                initialfile=fizo_default_filename,
                title="Сохранить заявку ФИЗО как"
            )

            if not fizo_file_path:
                logger.info("Пользователь отменил сохранение документа ФИЗО")
                return
            
            fizo_doc.save(fizo_file_path)
            logger.info("Документ ФИЗО сохранен: %s", fizo_file_path)
            
            try:
                fizo_pdf_path = fizo_file_path.replace('.docx', '.pdf')
                fizo_pdf_success = self.convert_to_pdf(fizo_file_path, fizo_pdf_path)
                
                if fizo_pdf_success:
                    logger.info("ФИЗО PDF создан: %s", fizo_pdf_path)
                else:
                    logger.warning("ФИЗО PDF не создан, DOCX сохранен: %s", fizo_file_path)
            except Exception as pdf_error:
                logger.exception("Ошибка конвертации ФИЗО в PDF: %s", pdf_error)
                    
        except Exception as e:
            logger.exception("Ошибка при создании ФИЗО: %s", e)
            self.show_simple_popup("Ошибка ФИЗО", "Ошибка при создании заявки ФИЗО")
        
class MainPage(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.configure(fg_color="#101416")

        watermark_path = resolve_asset("msecurity_logo_watermark.png")
        if os.path.exists(watermark_path):
            try:
                self.watermark_image = PhotoImage(file=watermark_path)
                self.watermark = ctk.CTkLabel(self, text="", image=self.watermark_image, fg_color="transparent")
                self.watermark.place(relx=0.58, rely=0.52, anchor="center")
            except Exception:
                logger.exception("Не удалось загрузить фоновый логотип")

        sidebar = ctk.CTkFrame(self, width=300, corner_radius=0, fg_color="#162023")
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        sidebar.grid_rowconfigure(6, weight=1)

        logo_path = resolve_asset("msecurity_logo.png")
        if os.path.exists(logo_path):
            try:
                self.logo_image = PhotoImage(file=logo_path).subsample(5, 5)
                ctk.CTkLabel(sidebar, text="", image=self.logo_image, fg_color="transparent").grid(
                    row=0, column=0, padx=26, pady=(26, 8), sticky="w"
                )
            except Exception:
                logger.exception("Не удалось загрузить логотип главной страницы")

        ctk.CTkLabel(
            sidebar,
            text="Генератор\nнормированных заданий",
            font=("TimesNewRoman", 24),
            justify="left",
            anchor="w",
        ).grid(row=1, column=0, padx=26, pady=(6, 8), sticky="ew")
        ctk.CTkLabel(
            sidebar,
            text="Выберите центр или откройте настройки приложения.",
            font=("TimesNewRoman", 13),
            text_color="gray72",
            wraplength=230,
            justify="left",
        ).grid(row=2, column=0, padx=26, pady=(0, 26), sticky="ew")

        ctk.CTkButton(
            sidebar,
            text="Центры",
            command=lambda: controller.show_page("ManageCenters"),
            height=38,
        ).grid(row=3, column=0, padx=26, pady=6, sticky="ew")

        ctk.CTkButton(
            sidebar,
            text="Настройки",
            command=lambda: controller.show_page("Settings"),
            height=38,
        ).grid(row=4, column=0, padx=26, pady=6, sticky="ew")

        ctk.CTkButton(
            sidebar,
            text="Выход",
            command=self.controller.on_close,
            fg_color="#8b1e24",
            hover_color="#6f171c",
            height=38,
        ).grid(row=7, column=0, padx=26, pady=(6, 26), sticky="ew")

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=0, column=1, sticky="nsew", padx=24, pady=24)
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            content,
            text="Центры",
            font=("TimesNewRoman", 24),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", pady=(0, 12))

        centers_frame = ctk.CTkScrollableFrame(content, fg_color="#151b1e")
        centers_frame.grid(row=1, column=0, sticky="nsew")
        for column in range(3):
            centers_frame.grid_columnconfigure(column, weight=1)

        buttons = [
            (center.get("button") or center.get("name"), center.get("id"))
            for center in self.controller.centers
            if center.get("id")
        ]

        for i, (text, page) in enumerate(buttons):
            row = i // 3
            col = i % 3
            btn = ctk.CTkButton(
                centers_frame,
                text=text,
                command=lambda p=page: controller.show_page(p),
                height=44,
                anchor="w",
            )
            btn.grid(row=row, column=col, pady=8, padx=8, sticky='ew')


class ManageCentersPage(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.centers = [dict(center) for center in controller.centers]
        self.current_index = 0

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self, text="Центры", font=("TimesNewRoman", 24)).grid(
            row=0, column=0, columnspan=2, pady=20, padx=20
        )

        self.list_frame = ctk.CTkScrollableFrame(self, width=260)
        self.list_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        self.form_frame = ctk.CTkFrame(self)
        self.form_frame.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)
        self.form_frame.grid_columnconfigure(1, weight=1)
        self.form_frame.grid_columnconfigure(2, weight=0)
        self.form_frame.grid_rowconfigure(6, weight=1)

        self.id_var = ctk.StringVar()
        self.button_var = ctk.StringVar()
        self.name_var = ctk.StringVar()
        self.template_var = ctk.StringVar()
        self.template_fizo_var = ctk.StringVar()
        self.east_var = ctk.BooleanVar()

        fields = [
            ("ID", self.id_var, False),
            ("Кнопка", self.button_var, False),
            ("Название в файле", self.name_var, False),
            ("Шаблон", self.template_var, True),
            ("Шаблон ФИЗО", self.template_fizo_var, True),
        ]

        for row, (label, var, browse) in enumerate(fields):
            ctk.CTkLabel(self.form_frame, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=6)
            entry = ctk.CTkEntry(self.form_frame, textvariable=var)
            entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
            self.bind_text_shortcuts(entry)
            if browse:
                command = self.choose_template if var == self.template_var else self.choose_fizo_template
                ctk.CTkButton(self.form_frame, text="Выбрать", width=110, command=command).grid(
                    row=row, column=2, sticky="ew", padx=8, pady=6
                )

        ctk.CTkCheckBox(self.form_frame, text="Дальний восток", variable=self.east_var).grid(
            row=5, column=1, sticky="w", padx=8, pady=6
        )

        ctk.CTkLabel(self.form_frame, text="Сотрудники").grid(row=6, column=0, sticky="nw", padx=8, pady=6)
        self.workers_text = ctk.CTkTextbox(self.form_frame, height=170)
        self.workers_text.grid(row=6, column=1, sticky="nsew", padx=8, pady=6)
        self.bind_text_shortcuts(self.workers_text)

        buttons_frame = ctk.CTkFrame(self.form_frame)
        buttons_frame.grid(row=7, column=0, columnspan=2, sticky="ew", padx=8, pady=10)
        buttons_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        ctk.CTkButton(buttons_frame, text="Добавить", command=self.add_center).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(buttons_frame, text="Удалить", command=self.delete_center, fg_color="red", hover_color="darkred").grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(buttons_frame, text="Сохранить", command=self.save_current).grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(buttons_frame, text="Назад", command=lambda: controller.show_page("MainPage")).grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        self.refresh_list()
        self.load_current()

    def bind_text_shortcuts(self, widget):
        widget.bind('<Control-KeyPress>', self.on_control_key)
        inner = getattr(widget, "_entry", None) or getattr(widget, "_textbox", None)
        if inner:
            inner.bind('<Control-KeyPress>', self.on_control_key)

    def on_control_key(self, event):
        key = (getattr(event, "keysym", "") or "").lower()
        keycode = getattr(event, "keycode", None)
        if key in ("c", "с", "cyrillic_es") or keycode == 67:
            return
        if key in ("x", "ч", "cyrillic_che") or keycode == 88:
            return
        if key in ("v", "м", "cyrillic_em") or keycode == 86:
            return
        if key in ("a", "ф", "cyrillic_ef") or keycode == 65:
            return self.on_select_all(event)

    def on_select_all(self, event):
        widget = getattr(event.widget, "_entry", None) or getattr(event.widget, "_textbox", None) or event.widget
        try:
            widget.selection_range(0, "end")
            widget.icursor("end")
        except Exception:
            widget.tag_add("sel", "1.0", "end")
            widget.mark_set("insert", "end")
        return "break"

    def refresh_list(self):
        for widget in self.list_frame.winfo_children():
            widget.destroy()

        for index, center in enumerate(self.centers):
            text = center.get("button") or center.get("name") or center.get("id") or "Новый центр"
            button = ctk.CTkButton(self.list_frame, text=text, command=lambda i=index: self.select_center(i))
            button.pack(fill="x", padx=5, pady=4)

    def select_center(self, index):
        self.current_index = index
        self.load_current()

    def load_current(self):
        if not self.centers:
            self.add_center()
            return

        center = self.centers[self.current_index]
        self.id_var.set(center.get("id", ""))
        self.button_var.set(center.get("button", ""))
        self.name_var.set(center.get("name", ""))
        self.template_var.set(center.get("template", ""))
        self.template_fizo_var.set(center.get("template_fizo", ""))
        self.east_var.set(bool(center.get("east", False)))
        self.workers_text.delete("1.0", "end")
        self.workers_text.insert("1.0", "\n".join(center.get("workers", [])))

    def collect_current(self):
        center_id = self.id_var.get().strip()
        workers = [
            line.strip()
            for line in self.workers_text.get("1.0", "end").splitlines()
            if line.strip()
        ]
        return {
            "id": center_id,
            "button": self.button_var.get().strip(),
            "name": self.name_var.get().strip(),
            "template": os.path.basename(self.template_var.get().strip()),
            "template_fizo": os.path.basename(self.template_fizo_var.get().strip()),
            "east": bool(self.east_var.get()),
            "workers": workers,
        }

    def save_current(self):
        center = self.collect_current()
        errors = self.validate_center(center)
        if errors:
            self.show_simple_popup("Ошибка", "\n".join(errors[:4]))
            return
        self.centers[self.current_index] = center
        self.controller.save_centers(self.centers)
        self.controller.show_page("ManageCenters")

    def validate_center(self, center):
        errors = []
        if not center["id"] or not center["name"] or not center["template"]:
            errors.append("Заполните ID, название и шаблон")
        ids = [
            item.get("id", "").strip()
            for index, item in enumerate(self.centers)
            if index != self.current_index
        ]
        if center["id"] and center["id"] in ids:
            errors.append("ID центра должен быть уникальным")
        if center["template"] and not os.path.exists(os.path.join(TEMPLATES_DIR, center["template"])):
            errors.append("Основной шаблон не найден в папке templates")
        if center["template_fizo"] and not os.path.exists(os.path.join(TEMPLATES_DIR, center["template_fizo"])):
            errors.append("Шаблон ФИЗО не найден в папке templates")
        if not center["workers"]:
            errors.append("Добавьте хотя бы одного сотрудника")
        return errors

    def choose_template(self):
        self.choose_template_file(self.template_var)

    def choose_fizo_template(self):
        self.choose_template_file(self.template_fizo_var)

    def choose_template_file(self, variable):
        path = filedialog.askopenfilename(
            initialdir=TEMPLATES_DIR,
            title="Выберите шаблон Word",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        if path:
            variable.set(os.path.basename(path))

    def add_center(self):
        new_number = len(self.centers) + 1
        self.centers.append({
            "id": f"Center{new_number}",
            "button": "Новый центр",
            "name": "Новый центр",
            "template": "",
            "template_fizo": "",
            "east": False,
            "workers": [],
        })
        self.current_index = len(self.centers) - 1
        self.refresh_list()
        self.load_current()

    def delete_center(self):
        if not self.centers:
            return
        if not messagebox.askyesno("Удалить центр", "Удалить выбранный центр?"):
            return
        self.centers.pop(self.current_index)
        self.current_index = max(0, self.current_index - 1)
        self.controller.save_centers(self.centers)
        self.controller.show_page("ManageCenters")

    def show_simple_popup(self, title, message):
        popup = ctk.CTkToplevel(self)
        popup.title(title)
        popup.geometry("400x120")
        popup.transient(self)
        ctk.CTkLabel(popup, text=message, wraplength=380).pack(pady=30, padx=10)
        popup.after(2000, popup.destroy)


class SettingsPage(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.people = {
            "chiefs": [dict(item) for item in controller.people.get("chiefs", [])],
            "issuers": [dict(item) for item in controller.people.get("issuers", [])],
        }
        self.current_group = "chiefs"
        self.current_index = 0

        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self, text="Настройки", font=("TimesNewRoman", 24)).grid(
            row=0, column=0, columnspan=2, pady=20, padx=20, sticky="w"
        )

        self.list_frame = ctk.CTkScrollableFrame(self, width=300)
        self.list_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        self.form_frame = ctk.CTkFrame(self)
        self.form_frame.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)
        self.form_frame.grid_columnconfigure(1, weight=1)

        self.group_var = ctk.StringVar(value="chiefs")
        group_frame = ctk.CTkFrame(self.form_frame, fg_color="transparent")
        group_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
        ctk.CTkRadioButton(group_frame, text="Должностные лица", variable=self.group_var, value="chiefs", command=self.change_group).grid(row=0, column=0, padx=8, pady=4, sticky="w")
        ctk.CTkRadioButton(group_frame, text="Исполнители", variable=self.group_var, value="issuers", command=self.change_group).grid(row=0, column=1, padx=8, pady=4, sticky="w")

        self.id_var = ctk.StringVar()
        self.label_var = ctk.StringVar()
        self.job_title_var = ctk.StringVar()
        self.fullname_var = ctk.StringVar()
        self.email_var = ctk.StringVar()

        self.fields = [
            ("ID", self.id_var),
            ("Отображение", self.label_var),
            ("Должность", self.job_title_var),
            ("ФИО для документа", self.fullname_var),
            ("Email", self.email_var),
        ]

        self.field_entries = {}
        for row, (label, variable) in enumerate(self.fields, start=1):
            ctk.CTkLabel(self.form_frame, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=6)
            entry = ctk.CTkEntry(self.form_frame, textvariable=variable)
            entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
            self.field_entries[label] = entry
            self.bind_text_shortcuts(entry)

        buttons_frame = ctk.CTkFrame(self.form_frame)
        buttons_frame.grid(row=6, column=0, columnspan=2, sticky="ew", padx=8, pady=10)
        buttons_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        ctk.CTkButton(buttons_frame, text="Добавить", command=self.add_item).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(buttons_frame, text="Удалить", command=self.delete_item, fg_color="red", hover_color="darkred").grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(buttons_frame, text="Сохранить", command=self.save_current).grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(buttons_frame, text="Назад", command=lambda: controller.show_page("MainPage")).grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        self.refresh_list()
        self.load_current()

    def bind_text_shortcuts(self, widget):
        widget.bind('<Control-KeyPress>', self.on_control_key)
        inner = getattr(widget, "_entry", None)
        if inner:
            inner.bind('<Control-KeyPress>', self.on_control_key)

    def on_control_key(self, event):
        key = (getattr(event, "keysym", "") or "").lower()
        keycode = getattr(event, "keycode", None)
        if key in ("a", "ф", "cyrillic_ef") or keycode == 65:
            widget = getattr(event.widget, "_entry", None) or event.widget
            widget.selection_range(0, "end")
            widget.icursor("end")
            return "break"

    def change_group(self):
        self.current_group = self.group_var.get()
        self.current_index = 0
        self.refresh_list()
        self.load_current()

    def refresh_list(self):
        for widget in self.list_frame.winfo_children():
            widget.destroy()

        for group, title in (("chiefs", "Должностные лица"), ("issuers", "Исполнители")):
            ctk.CTkLabel(self.list_frame, text=title, font=("TimesNewRoman", 16)).pack(anchor="w", padx=5, pady=(10, 4))
            for index, item in enumerate(self.people.get(group, [])):
                text = item.get("label") or item.get("fullname") or item.get("id") or "Новая запись"
                button = ctk.CTkButton(
                    self.list_frame,
                    text=text,
                    command=lambda g=group, i=index: self.select_item(g, i),
                )
                button.pack(fill="x", padx=5, pady=3)

    def select_item(self, group, index):
        self.current_group = group
        self.group_var.set(group)
        self.current_index = index
        self.load_current()

    def load_current(self):
        items = self.people.get(self.current_group, [])
        if not items:
            self.add_item()
            return

        item = items[self.current_index]
        self.id_var.set(item.get("id", ""))
        self.label_var.set(item.get("label", ""))
        self.job_title_var.set(item.get("job_title", ""))
        self.fullname_var.set(item.get("fullname", ""))
        self.email_var.set(item.get("email", ""))

        self.field_entries["Должность"].configure(state="normal" if self.current_group == "chiefs" else "disabled")
        self.field_entries["Email"].configure(state="normal" if self.current_group == "issuers" else "disabled")

    def collect_current(self):
        item = {
            "id": self.id_var.get().strip(),
            "label": self.label_var.get().strip(),
            "fullname": self.fullname_var.get().strip(),
        }
        if self.current_group == "chiefs":
            item["job_title"] = self.job_title_var.get().strip()
        else:
            item["email"] = self.email_var.get().strip()
        return item

    def save_current(self):
        item = self.collect_current()
        errors = self.validate_item(item)
        if errors:
            self.show_simple_popup("Ошибка", "\n".join(errors))
            return

        self.people[self.current_group][self.current_index] = item
        self.controller.save_people(self.people)
        self.controller.show_page("Settings")

    def validate_item(self, item):
        errors = []
        if not item["id"] or not item["label"] or not item["fullname"]:
            errors.append("Заполните ID, отображение и ФИО")
        if self.current_group == "chiefs" and not item.get("job_title"):
            errors.append("Укажите должность")
        if self.current_group == "issuers" and not item.get("email"):
            errors.append("Укажите email")

        ids = [
            value.get("id", "").strip()
            for index, value in enumerate(self.people.get(self.current_group, []))
            if index != self.current_index
        ]
        if item["id"] in ids:
            errors.append("ID должен быть уникальным")
        return errors

    def add_item(self):
        group = self.current_group
        number = len(self.people.get(group, [])) + 1
        if group == "chiefs":
            item = {"id": f"chief_{number}", "label": "Новое лицо", "job_title": "", "fullname": ""}
        else:
            item = {"id": f"issuer_{number}", "label": "Новый исполнитель", "fullname": "", "email": ""}
        self.people.setdefault(group, []).append(item)
        self.current_index = len(self.people[group]) - 1
        self.refresh_list()
        self.load_current()

    def delete_item(self):
        items = self.people.get(self.current_group, [])
        if not items:
            return
        if not messagebox.askyesno("Удалить запись", "Удалить выбранную запись?"):
            return
        items.pop(self.current_index)
        self.current_index = max(0, self.current_index - 1)
        self.controller.save_people(self.people)
        self.controller.show_page("Settings")

    def show_simple_popup(self, title, message):
        popup = ctk.CTkToplevel(self)
        popup.title(title)
        popup.geometry("420x140")
        popup.transient(self)
        ctk.CTkLabel(popup, text=message, wraplength=400).pack(pady=30, padx=10)
        popup.after(2500, popup.destroy)


if __name__ == "__main__":
    app = App()
    app.mainloop()
