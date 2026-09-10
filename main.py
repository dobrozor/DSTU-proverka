import os
import sys
import json
import time
import difflib
import threading
import queue
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from docx import Document
from docx.shared import RGBColor
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

# ==============================================================================
#  КОНФИГУРАЦИЯ
# ==============================================================================

def _app_dir():
    """Папка, где лежит exe (или скрипт при запуске из исходников)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(_app_dir(), "config.json")
LOG_DIR = os.path.join(_app_dir(), "logs")

DEFAULT_CONFIG = {
    "login": "eselezneva",
    "password": "RZiPbrQA",
    "base_url": "https://edu.donstu.ru/WebApp/#",
    "browser_channel": "msedge",     # msedge / chrome / chromium
    "headless": False,
    "auto_open_result": True,
    "fio_similarity_threshold": 0.55  # защита от ложных совпадений похожих ФИО
}


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            cfg.update(saved)
        except Exception:
            pass
    else:
        save_config(cfg)
    return cfg


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ==============================================================================
#  СЛОВАРИ / ЭВРИСТИКИ
# ==============================================================================

# Полный словарь замен для имен (для перебора вариантов написания при поиске)
NAME_REPLACEMENTS = {
    "Артем": "Артём",
    "Aлена": "Алёна",
    "Федор": "Фёдор",
    "Эдуарт": "Эдуард",
    "Семен": "Семён",
    "Петр": "Пётр",
    "Наталья": "Наталия",
    "Софья": "София",
    "Дарья": "Дария",
    "Марья": "Мария",
    "Данил": "Даниил",
    "Данила": "Даниил",
    "Кирил": "Кирилл",
    "Филип": "Филипп",
    "Генадий": "Геннадий",
    "Ала": "Алла",
    "Темур": "Тимур",
    "Артём": "Артем",
    "Алёна": "Aлена",
    "Фёдор": "Федор",
    "Эдуард": "Эдуарт",
    "Семён": "Семен",
    "Пётр": "Петр",
    "Наталия": "Наталья",
    "София": "Софья",
    "Дария": "Дарья",
    "Мария": "Марья",
    "Даниил": "Данил",
    "Кирилл": "Кирил",
    "Филипп": "Филип",
    "Геннадий": "Генадий",
    "Алла": "Ала",
    "Тимур": "Темур"
}

# Словарь для поиска колонок в Word и соответствующих им лейблов на сайте
SITE_LABELS_MAPPING = {
    'фио': 'ФИО',
    'ф.и.о': 'ФИО',
    'групп': 'Группа',
    'курс': 'Курс',
    'зачетн': 'Номер зачетной книжки',
    'зачётн': 'Номер зачетной книжки',
    'факультет': 'Факультет',
    'кафедр': 'Кафедра',
    'датарожд': 'Дата рождения',
    'гражданств': 'Гражданство',
    'годпоступл': 'Год поступления'
}


def get_name_variations(fio):
    variations = [fio]
    for key, val in NAME_REPLACEMENTS.items():
        if key in fio:
            variations.append(fio.replace(key, val))
        elif val in fio:
            variations.append(fio.replace(val, key))
    return list(dict.fromkeys(variations))


def normalize_group(group):
    """Нормализует запись группы для сравнения: убирает пробелы/дефисы, верхний регистр."""
    if not group:
        return ""
    return "".join(ch for ch in group.upper() if ch.isalnum())


def fio_similarity(a, b):
    """Коэффициент схожести двух ФИО (0..1), защита от ложных совпадений в поиске."""
    a = (a or "").strip().lower()
    b = (b or "").strip().lower()
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def retry(times=2, delay=0.6, exceptions=(PWTimeoutError, Exception)):
    """Простой декоратор ретраев для нестабильных сетевых операций."""
    def decorator(fn):
        def wrapper(*args, **kwargs):
            last_exc = None
            for attempt in range(times + 1):
                try:
                    return fn(*args, **kwargs)
                except exceptions as e:
                    last_exc = e
                    if attempt < times:
                        time.sleep(delay)
            raise last_exc
        return wrapper
    return decorator


# ==============================================================================
#  ОСНОВНОЕ ПРИЛОЖЕНИЕ
# ==============================================================================

class DocCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Автоматизация проверки студентов ДОНСТУ")
        self.root.geometry("700x760")
        self.root.minsize(620, 620)
        self.root.configure(padx=25, pady=20, bg="#FFFFFF")

        self.config_data = load_config()

        self.file_path = None
        self.msg_queue = queue.Queue()
        self.last_changes = None
        self.last_file_path = None
        self.cancel_event = threading.Event()
        self.is_running = False
        self.run_start_time = None

        self.stats = {"processed": 0, "found": 0, "updated": 0, "errors": 0, "not_found": 0}

        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("Orange.Horizontal.TProgressbar",
                              troughcolor='#F5F5F5', background='#FF6F00',
                              bordercolor='#FFFFFF', lightcolor='#FF6F00', darkcolor='#FF6F00')

        self._build_menu()
        self._build_ui()

        self.check_queue()

    # ------------------------------------------------------------------ UI --

    def _build_menu(self):
        menubar = tk.Menu(self.root)
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="Настройки подключения...", command=self.open_settings)
        settings_menu.add_command(label="Открыть папку логов", command=self.open_logs_folder)
        menubar.add_cascade(label="Настройки", menu=settings_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="О программе", command=self.show_about)
        menubar.add_cascade(label="Справка", menu=help_menu)
        self.root.config(menu=menubar)

    def _build_ui(self):
        self.lbl_title = tk.Label(self.root, text="Проверка списков ДГТУ", font=("Segoe UI", 18, "bold"),
                                   bg="#FFFFFF", fg="#212121")
        self.lbl_title.pack(anchor=tk.W, pady=(0, 2))

        self.lbl_subtitle = tk.Label(self.root, text="Выберите документ .docx для запуска сверки с базой данных",
                                      font=("Segoe UI", 10), bg="#FFFFFF", fg="#757575")
        self.lbl_subtitle.pack(anchor=tk.W, pady=(0, 20))

        # --- Выбор файла ---
        self.frame_file = tk.Frame(self.root, bg="#FFFFFF")
        self.frame_file.pack(fill=tk.X, pady=5)

        self.btn_select = tk.Button(self.frame_file, text="Выбрать файл", command=self.select_file, width=15,
                                     font=("Segoe UI", 10, "bold"), bg="#333333", fg="#FFFFFF",
                                     activebackground="#444444", activeforeground="#FFFFFF", relief="flat", bd=0,
                                     pady=5, cursor="hand2")
        self.btn_select.pack(side=tk.LEFT, padx=(0, 15))

        self.lbl_file_name = tk.Label(self.frame_file, text="Файл не выбран", font=("Segoe UI", 10, "italic"),
                                       bg="#FFFFFF", fg="#9E9E9E")
        self.lbl_file_name.pack(side=tk.LEFT, fill=tk.X)

        # --- Опции ---
        self.frame_options = tk.Frame(self.root, bg="#FFFFFF")
        self.frame_options.pack(fill=tk.X, pady=(12, 0))

        self.smart_group_var = tk.BooleanVar(value=False)
        self.chk_smart_group = tk.Checkbutton(
            self.frame_options, text="Умная замена группы (проверять совпадение первой буквы)",
            variable=self.smart_group_var, font=("Segoe UI", 10), bg="#FFFFFF", fg="#333333",
            activebackground="#FFFFFF", activeforeground="#333333", selectcolor="#FFFFFF"
        )
        self.chk_smart_group.pack(anchor=tk.W)

        self.headless_var = tk.BooleanVar(value=self.config_data.get("headless", False))
        self.chk_headless = tk.Checkbutton(
            self.frame_options, text="Фоновый режим (без окна браузера — быстрее)",
            variable=self.headless_var, font=("Segoe UI", 10), bg="#FFFFFF", fg="#333333",
            activebackground="#FFFFFF", activeforeground="#333333", selectcolor="#FFFFFF"
        )
        self.chk_headless.pack(anchor=tk.W)

        self.auto_open_var = tk.BooleanVar(value=self.config_data.get("auto_open_result", True))
        self.chk_auto_open = tk.Checkbutton(
            self.frame_options, text="Открыть файл автоматически после завершения",
            variable=self.auto_open_var, font=("Segoe UI", 10), bg="#FFFFFF", fg="#333333",
            activebackground="#FFFFFF", activeforeground="#333333", selectcolor="#FFFFFF"
        )
        self.chk_auto_open.pack(anchor=tk.W)

        # --- Кнопки управления ---
        self.frame_buttons = tk.Frame(self.root, bg="#FFFFFF")
        self.frame_buttons.pack(fill=tk.X, pady=15)

        self.btn_start = tk.Button(self.frame_buttons, text="Начать сверку", bg="#FF6F00", fg="#FFFFFF",
                                    font=("Segoe UI", 12, "bold"), command=self.start_processing, state=tk.DISABLED,
                                    activebackground="#E65C00", activeforeground="#FFFFFF", relief="flat", bd=0,
                                    pady=7, cursor="hand2")
        self.btn_start.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        self.btn_stop = tk.Button(self.frame_buttons, text="Стоп", bg="#B00020", fg="#FFFFFF",
                                   font=("Segoe UI", 12, "bold"), command=self.stop_processing, state=tk.DISABLED,
                                   activebackground="#8E001A", activeforeground="#FFFFFF", relief="flat", bd=0,
                                   pady=7, cursor="hand2")
        self.btn_stop.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        self.btn_show_last = tk.Button(self.frame_buttons, text="Показать прошлый отчет", bg="#333333", fg="#FFFFFF",
                                        font=("Segoe UI", 12, "bold"), command=self.reopen_summary, state=tk.DISABLED,
                                        activebackground="#444444", activeforeground="#FFFFFF", relief="flat", bd=0,
                                        pady=7, cursor="hand2")
        self.btn_show_last.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # --- Прогресс ---
        self.progress = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, mode='determinate',
                                         style="Orange.Horizontal.TProgressbar")
        self.progress.pack(fill=tk.X, pady=(0, 8))

        self.lbl_status = tk.Label(self.root, text="Ожидание запуска...", font=("Segoe UI", 10, "bold"),
                                    bg="#FFFFFF", fg="#FF6F00")
        self.lbl_status.pack(anchor=tk.W)

        self.lbl_stats = tk.Label(self.root, text="", font=("Segoe UI", 9), bg="#FFFFFF", fg="#757575")
        self.lbl_stats.pack(anchor=tk.W, pady=(0, 12))

        # --- Лог ---
        self.log_text = tk.Text(self.root, height=12, bg="#F9F9F9", fg="#333333", font=("Courier New", 10),
                                 relief="flat", bd=1, highlightbackground="#E0E0E0", highlightcolor="#FF6F00",
                                 highlightthickness=1, padx=10, pady=10, state=tk.DISABLED, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_configure("error", foreground="#B00020")
        self.log_text.tag_configure("warn", foreground="#E65C00")
        self.log_text.tag_configure("success", foreground="#1B8A2F")
        self.log_text.tag_configure("muted", foreground="#9E9E9E")

        info_text = (
            "• Скрипт автоматически найдет колонки (Зачетная книжка, Курс и др.) и заполнит пустые поля.\n"
            "• Если у студента несколько профилей (бакалавр/магистр), данные впишутся через запятую.\n"
            "• Не найденные студенты выделяются красным, обновлённые поля — зелёным."
        )
        self.lbl_info = tk.Label(self.root, text=info_text, justify=tk.LEFT, bg="#FFFFFF", fg="#757575",
                                  font=("Segoe UI", 9), anchor=tk.W)
        self.lbl_info.pack(fill=tk.X, pady=(15, 0))

    # ---------------------------------------------------------- Настройки --

    def open_settings(self):
        win = tk.Toplevel(self.root)
        win.title("Настройки подключения")
        win.geometry("420x360")
        win.configure(padx=20, pady=20, bg="#FFFFFF")
        win.transient(self.root)
        win.grab_set()

        def add_field(label_text, value, show=None):
            tk.Label(win, text=label_text, font=("Segoe UI", 10, "bold"), bg="#FFFFFF", fg="#333333") \
                .pack(anchor=tk.W, pady=(8, 2))
            entry = tk.Entry(win, font=("Segoe UI", 10), show=show, relief="solid", bd=1)
            entry.insert(0, value)
            entry.pack(fill=tk.X, ipady=4)
            return entry

        e_login = add_field("Логин", self.config_data.get("login", ""))
        e_password = add_field("Пароль", self.config_data.get("password", ""), show="*")
        e_url = add_field("URL сайта", self.config_data.get("base_url", ""))

        tk.Label(win, text="Канал браузера", font=("Segoe UI", 10, "bold"), bg="#FFFFFF", fg="#333333") \
            .pack(anchor=tk.W, pady=(8, 2))
        channel_var = tk.StringVar(value=self.config_data.get("browser_channel", "msedge"))
        channel_combo = ttk.Combobox(win, textvariable=channel_var, state="readonly",
                                      values=["msedge", "chrome", "chromium"])
        channel_combo.pack(fill=tk.X)

        def do_save():
            self.config_data["login"] = e_login.get().strip()
            self.config_data["password"] = e_password.get()
            self.config_data["base_url"] = e_url.get().strip()
            self.config_data["browser_channel"] = channel_var.get()
            save_config(self.config_data)
            messagebox.showinfo("Готово", "Настройки сохранены.")
            win.destroy()

        btn_save = tk.Button(win, text="Сохранить", command=do_save, bg="#FF6F00", fg="#FFFFFF",
                              font=("Segoe UI", 10, "bold"), relief="flat", bd=0, pady=6, cursor="hand2")
        btn_save.pack(fill=tk.X, pady=(20, 0))

    def open_logs_folder(self):
        os.makedirs(LOG_DIR, exist_ok=True)
        try:
            if sys.platform == "win32":
                os.startfile(LOG_DIR)
            elif sys.platform == "darwin":
                os.system(f'open "{LOG_DIR}"')
            else:
                os.system(f'xdg-open "{LOG_DIR}"')
        except Exception:
            pass

    def show_about(self):
        messagebox.showinfo(
            "О программе",
            "Автоматизация проверки студентов ДОНСТУ\n\n"
            "Сверяет данные из Word-документа с внутренней базой edu.donstu.ru "
            "и дозаполняет отсутствующие поля."
        )

    # ------------------------------------------------------------ Работа --

    def select_file(self):
        initial_dir = self.config_data.get("last_dir", os.path.expanduser("~"))
        file_path = filedialog.askopenfilename(title="Выберите документ", initialdir=initial_dir,
                                                 filetypes=[("Word Documents", "*.docx")])
        if file_path:
            self.file_path = file_path
            self.config_data["last_dir"] = os.path.dirname(file_path)
            save_config(self.config_data)
            self.lbl_file_name.config(text=os.path.basename(file_path), fg="#333333", font=("Segoe UI", 10, "bold"))
            self.btn_start.config(state=tk.NORMAL)

    def log(self, message, tag=None):
        print(message)
        self.msg_queue.put({"type": "log", "text": message, "tag": tag})

    def update_status(self, text, percent=None):
        self.msg_queue.put({"type": "progress", "text": text, "value": percent})

    def update_stats(self):
        s = self.stats
        elapsed = ""
        if self.run_start_time:
            secs = int(time.time() - self.run_start_time)
            elapsed = f" · Время: {secs // 60:02d}:{secs % 60:02d}"
        text = (f"Обработано: {s['processed']} · Найдено: {s['found']} · "
                f"Обновлено: {s['updated']} · Не найдено: {s['not_found']} · "
                f"Ошибок: {s['errors']}{elapsed}")
        self.msg_queue.put({"type": "stats", "text": text})

    def check_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                if msg["type"] == "log":
                    self.log_text.config(state=tk.NORMAL)
                    tag = msg.get("tag")
                    self.log_text.insert(tk.END, msg["text"] + "\n", tag if tag else ())
                    self.log_text.see(tk.END)
                    self.log_text.config(state=tk.DISABLED)
                elif msg["type"] == "progress":
                    if msg["value"] is not None:
                        self.progress["value"] = msg["value"]
                    self.lbl_status.config(text=msg["text"])
                elif msg["type"] == "stats":
                    self.lbl_stats.config(text=msg["text"])
                elif msg["type"] == "done":
                    self.is_running = False
                    self.progress["value"] = 100
                    status_text = "Проверка остановлена пользователем" if msg.get("cancelled") \
                        else "Проверка успешно завершена!"
                    self.lbl_status.config(text=status_text)
                    self.last_changes = msg["changes"]
                    self.last_file_path = msg["path"]
                    self.btn_show_last.config(state=tk.NORMAL)
                    self.show_summary_window(msg["changes"], msg["path"], cancelled=msg.get("cancelled", False))
                    self.btn_select.config(state=tk.NORMAL)
                    self.btn_start.config(state=tk.NORMAL)
                    self.btn_stop.config(state=tk.DISABLED)
                    if self.auto_open_var.get() and not msg.get("cancelled"):
                        self._open_file(msg["path"])
                elif msg["type"] == "error":
                    self.is_running = False
                    self.lbl_status.config(text="Ошибка!")
                    messagebox.showerror("Ошибка", msg["text"])
                    self.btn_select.config(state=tk.NORMAL)
                    self.btn_start.config(state=tk.NORMAL)
                    self.btn_stop.config(state=tk.DISABLED)
        except queue.Empty:
            pass
        self.root.after(100, self.check_queue)

    def _open_file(self, path):
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except Exception:
            pass

    def show_summary_window(self, changes, file_path, cancelled=False):
        summary_win = tk.Toplevel(self.root)
        summary_win.title("Отчет об изменениях")
        summary_win.geometry("580x460")
        summary_win.configure(padx=20, pady=20, bg="#FFFFFF")

        title = "Проверка остановлена — внесённые до этого изменения:" if cancelled else "Внесенные изменения:"
        lbl = tk.Label(summary_win, text=title, font=("Segoe UI", 12, "bold"), bg="#FFFFFF", fg="#212121")
        lbl.pack(anchor=tk.W, pady=(0, 10))

        text_frame = tk.Frame(summary_win, bg="#FFFFFF")
        text_frame.pack(fill=tk.BOTH, expand=True)

        text_area = tk.Text(text_frame, wrap=tk.WORD, font=("Courier New", 10), bg="#F9F9F9", fg="#333333",
                             relief="flat", bd=1, highlightbackground="#E0E0E0", highlightthickness=1, padx=8, pady=8)
        scrollbar = tk.Scrollbar(text_frame, command=text_area.yview)
        text_area.config(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        if not changes:
            text_area.insert(tk.END, "Данные остались без изменений.\n")
        else:
            for change in changes:
                text_area.insert(tk.END, f"• {change}\n")
        text_area.config(state=tk.DISABLED)

        lbl_file = tk.Label(summary_win, text=f"Файл успешно сохранен как:\n{os.path.basename(file_path)}",
                             fg="#FF6F00", justify=tk.LEFT, bg="#FFFFFF", font=("Segoe UI", 10, "bold"))
        lbl_file.pack(anchor=tk.W, pady=(15, 0))

        btn_frame = tk.Frame(summary_win, bg="#FFFFFF")
        btn_frame.pack(fill=tk.X, pady=(10, 0))

        def export_report():
            export_path = filedialog.asksaveasfilename(
                title="Сохранить отчёт", defaultextension=".txt",
                initialfile="отчет_об_изменениях.txt", filetypes=[("Текстовый файл", "*.txt")])
            if export_path:
                with open(export_path, "w", encoding="utf-8") as f:
                    f.write(f"Отчёт по файлу: {os.path.basename(file_path)}\n")
                    f.write(f"Дата: {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}\n\n")
                    if not changes:
                        f.write("Данные остались без изменений.\n")
                    else:
                        for change in changes:
                            f.write(f"• {change}\n")
                messagebox.showinfo("Готово", "Отчёт сохранён.")

        btn_export = tk.Button(btn_frame, text="Экспорт отчёта", command=export_report, width=15,
                                font=("Segoe UI", 10, "bold"), bg="#333333", fg="#FFFFFF",
                                activebackground="#444444", activeforeground="#FFFFFF", relief="flat", bd=0, pady=5,
                                cursor="hand2")
        btn_export.pack(side=tk.LEFT)

        btn_open = tk.Button(btn_frame, text="Открыть файл", command=lambda: self._open_file(file_path), width=15,
                              font=("Segoe UI", 10, "bold"), bg="#FF6F00", fg="#FFFFFF",
                              activebackground="#E65C00", activeforeground="#FFFFFF", relief="flat", bd=0, pady=5,
                              cursor="hand2")
        btn_open.pack(side=tk.LEFT, padx=(10, 0))

        btn_close = tk.Button(btn_frame, text="Закрыть", command=summary_win.destroy, width=15,
                               font=("Segoe UI", 10, "bold"), bg="#333333", fg="#FFFFFF",
                               activebackground="#444444", activeforeground="#FFFFFF", relief="flat", bd=0, pady=5,
                               cursor="hand2")
        btn_close.pack(side=tk.RIGHT)

    def reopen_summary(self):
        if self.last_changes is not None and self.last_file_path is not None:
            self.show_summary_window(self.last_changes, self.last_file_path)

    def start_processing(self):
        if not self.file_path:
            return

        self.btn_select.config(state=tk.DISABLED)
        self.btn_start.config(state=tk.DISABLED)
        self.btn_show_last.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)

        use_smart_group = self.smart_group_var.get()
        headless = self.headless_var.get()
        self.config_data["headless"] = headless
        self.config_data["auto_open_result"] = self.auto_open_var.get()
        save_config(self.config_data)

        self.last_changes = None
        self.last_file_path = None
        self.cancel_event.clear()
        self.is_running = True
        self.run_start_time = time.time()
        self.stats = {"processed": 0, "found": 0, "updated": 0, "errors": 0, "not_found": 0}

        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        self.progress["value"] = 0
        self.update_status("Инициализация браузера...", 0)

        threading.Thread(target=self.process_file_worker,
                          args=(self.file_path, use_smart_group, headless), daemon=True).start()

    def stop_processing(self):
        if self.is_running:
            self.cancel_event.set()
            self.btn_stop.config(state=tk.DISABLED)
            self.log("⏹ Получен сигнал остановки — завершаем текущего студента и сохраняем файл...", "warn")

    # -------------------------------------------------------- Обработка --

    def process_file_worker(self, file_path, use_smart_group, headless):
        os.makedirs(LOG_DIR, exist_ok=True)
        log_file_path = os.path.join(LOG_DIR, f"log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        log_lines = []

        def log_both(message, tag=None):
            self.log(message, tag)
            log_lines.append(message)

        cancelled = False
        try:
            log_both("Открытие документа...")
            doc = Document(file_path)
            if not doc.tables:
                self.msg_queue.put({"type": "error", "text": "В документе не найдены таблицы."})
                return

            target_table = None
            headers = []
            for tbl in doc.tables:
                if len(tbl.rows) == 0:
                    continue
                current_headers = [cell.text.strip().lower() for cell in tbl.rows[0].cells]
                if any('фио' in h.replace(" ", "") or 'ф.и.о' in h.replace(" ", "") for h in current_headers):
                    target_table = tbl
                    headers = current_headers
                    break

            if not target_table:
                self.msg_queue.put(
                    {"type": "error", "text": "В документе не найдена таблица со списком студентов (нет колонки ФИО)."})
                return

            # Динамическое определение колонок
            doc_columns = {}
            for i, h in enumerate(headers):
                h_clean = h.replace(" ", "").replace(".", "")
                for key, label in SITE_LABELS_MAPPING.items():
                    if key in h_clean:
                        doc_columns[label] = i
                        break

            if 'ФИО' not in doc_columns:
                self.msg_queue.put({"type": "error", "text": "Колонка ФИО не распознана."})
                return

            log_both(f"Найдены колонки: {', '.join(doc_columns.keys())}")

            total_rows = len(target_table.rows) - 1
            changes_summary = []

            cfg = self.config_data
            base_url = cfg.get("base_url", DEFAULT_CONFIG["base_url"])
            browser_channel = cfg.get("browser_channel", "msedge")
            similarity_threshold = cfg.get("fio_similarity_threshold", 0.55)

            with sync_playwright() as p:
                log_both("Запуск браузера...")
                self.update_status("Запуск браузера...", 0)

                try:
                    browser = p.chromium.launch(headless=headless, channel=browser_channel,
                                                 args=["--no-proxy-server"])
                except Exception as e:
                    log_both(f"⚠️ Не удалось запустить браузер '{browser_channel}' ({e}). Пробуем встроенный Chromium.",
                              "warn")
                    browser = p.chromium.launch(headless=headless, args=["--no-proxy-server"])

                context = browser.new_context()
                page = context.new_page()

                @retry(times=2, delay=0.6)
                def goto_search():
                    page.goto(base_url)
                    search_input = page.locator("input[placeholder='Поиск']")
                    search_input.wait_for(state="visible", timeout=15000)
                    return search_input

                try:
                    page.goto(base_url)
                    page.locator('input[name="login"]').fill(cfg.get("login", ""))
                    page.locator('input[name="password"]').fill(cfg.get("password", ""))
                    page.locator('input[name="password"]').press("Enter")

                    search_input = page.locator("input[placeholder='Поиск']")
                    search_input.wait_for(state="visible", timeout=15000)
                    log_both("Авторизация успешна. Начинаем сверку.", "success")

                    for idx, row in enumerate(target_table.rows[1:]):
                        if self.cancel_event.is_set():
                            cancelled = True
                            log_both("⏹ Остановлено пользователем перед следующим студентом.", "warn")
                            break

                        row_data = {}
                        for label, col_idx in doc_columns.items():
                            if col_idx < len(row.cells):
                                row_data[label] = row.cells[col_idx].text.strip()
                            else:
                                row_data[label] = ""

                        fio = row_data.get('ФИО', "")
                        doc_group = row_data.get('Группа', "")

                        if not fio or fio == doc_group:
                            continue

                        self.stats["processed"] += 1
                        percent = int(((idx + 1) / total_rows) * 100)
                        self.update_status(f"Обработка: {idx + 1} из {total_rows} ({fio})", percent)
                        log_both("-" * 40)
                        log_both(f"[{idx + 1}/{total_rows}] Обработка: {fio}")

                        variations = get_name_variations(fio)
                        student_found = False
                        aggregated_site_data = {label: [] for label in doc_columns.keys()}

                        try:
                            for current_fio in variations:
                                if student_found or self.cancel_event.is_set():
                                    break

                                try:
                                    search_input = goto_search()
                                except Exception as e:
                                    log_both(f"  [!] Не удалось открыть страницу поиска: {e}", "error")
                                    continue

                                search_input.click()
                                search_input.fill(current_fio)

                                listbox_options = page.locator("div[role='listbox'] div[role='option']")

                                try:
                                    listbox_options.first.wait_for(state="visible", timeout=6000)
                                    options_count = listbox_options.count()
                                except Exception:
                                    options_count = 0

                                if options_count == 0:
                                    continue

                                for opt_idx in range(options_count - 1, -1, -1):
                                    if self.cancel_event.is_set():
                                        break

                                    if opt_idx < options_count - 1:
                                        try:
                                            search_input = goto_search()
                                        except Exception as e:
                                            log_both(f"  [!] Не удалось перезагрузить поиск: {e}", "error")
                                            break
                                        search_input.click()
                                        search_input.fill(current_fio)
                                        listbox_options = page.locator("div[role='listbox'] div[role='option']")
                                        try:
                                            listbox_options.first.wait_for(state="visible", timeout=6000)
                                        except Exception:
                                            break

                                    if opt_idx >= listbox_options.count():
                                        continue

                                    target_option = listbox_options.nth(opt_idx)
                                    raw_text = target_option.inner_text().strip().split('\n')[0]
                                    current_site_fio = raw_text.split('(')[0].strip()

                                    # Защита от ложного совпадения похожих ФИО в выдаче поиска
                                    sim = fio_similarity(fio, current_site_fio)
                                    if sim < similarity_threshold:
                                        log_both(
                                            f"  [!] Пропуск профиля: {current_site_fio}. "
                                            f"Схожесть ФИО слишком низкая ({sim:.2f}).", "warn")
                                        continue

                                    target_option.click()

                                    try:
                                        page.locator(
                                            "xpath=//label[contains(text(), 'ФИО')]/following-sibling::input"
                                        ).wait_for(state="attached", timeout=6000)
                                        # ждём, пока значение реально подтянется, а не просто появится поле
                                        fio_input = page.locator(
                                            "xpath=//label[contains(text(), 'ФИО')]/following-sibling::input").first
                                        for _ in range(20):
                                            if fio_input.input_value().strip():
                                                break
                                            page.wait_for_timeout(100)
                                    except Exception:
                                        pass

                                    site_data_current = {}
                                    for label in doc_columns.keys():
                                        xpath = f"xpath=//label[contains(text(), '{label}')]/following-sibling::input"
                                        loc = page.locator(xpath)
                                        if loc.count() > 0:
                                            site_data_current[label] = loc.first.input_value().strip()
                                        else:
                                            site_data_current[label] = ""

                                    if not site_data_current.get('ФИО') and current_site_fio:
                                        site_data_current['ФИО'] = current_site_fio

                                    site_group = site_data_current.get('Группа', "")

                                    is_match = False
                                    if not doc_group:
                                        is_match = True
                                    elif not use_smart_group:
                                        if site_group:
                                            is_match = True
                                    elif site_group and normalize_group(doc_group)[:1] == normalize_group(site_group)[:1]:
                                        is_match = True

                                    if is_match and site_group:
                                        student_found = True
                                        self.stats["found"] += 1
                                        for label in doc_columns.keys():
                                            val = site_data_current.get(label)
                                            if val and val not in aggregated_site_data[label]:
                                                aggregated_site_data[label].append(val)
                                    else:
                                        disp_group = site_group if site_group else "Отсутствует (Сотрудник)"
                                        log_both(
                                            f"  [!] Пропуск профиля: {current_site_fio}. "
                                            f"Группа [{disp_group}] не подошла.", "muted")
                        except Exception as e:
                            self.stats["errors"] += 1
                            log_both(f"⚠️ Непредвиденная ошибка при обработке '{fio}': {e}. Идём дальше.", "error")

                        if not student_found:
                            self.stats["not_found"] += 1
                            log_both(f"❌ Подходящие варианты для '{fio}' не найдены. Выделяем красным.", "error")
                            fio_cell = row.cells[doc_columns['ФИО']]
                            for paragraph in fio_cell.paragraphs:
                                for run in paragraph.runs:
                                    run.font.color.rgb = RGBColor(255, 0, 0)
                            self.update_stats()
                            continue

                        try:
                            row_updated = False
                            for label, col_idx in doc_columns.items():
                                if not aggregated_site_data[label]:
                                    continue

                                if label == 'ФИО':
                                    final_val = aggregated_site_data['ФИО'][0]
                                else:
                                    final_val = ", ".join(aggregated_site_data[label])

                                doc_val = row_data.get(label, "")

                                if doc_val != final_val and final_val:
                                    display_doc_val = doc_val if doc_val else "пусто"
                                    log_both(f"🔄 Обновляем [{label}]: {display_doc_val} -> {final_val}", "success")
                                    changes_summary.append(
                                        f"{aggregated_site_data.get('ФИО', [fio])[0]}: {label} "
                                        f"[{display_doc_val}] ➔ [{final_val}]")
                                    row.cells[col_idx].text = final_val
                                    row_updated = True
                                    # подсветка обновлённой ячейки зелёным
                                    for paragraph in row.cells[col_idx].paragraphs:
                                        for run in paragraph.runs:
                                            run.font.color.rgb = RGBColor(0, 140, 40)

                            if row_updated:
                                self.stats["updated"] += 1

                        except Exception as e:
                            self.stats["errors"] += 1
                            log_both(f"⚠️ Ошибка при записи данных {fio}: {e}", "error")

                        self.update_stats()

                finally:
                    browser.close()

            name, ext = os.path.splitext(file_path)
            updated_file_path = f"{name}_ОБНОВЛЕННЫЙ{ext}"
            doc.save(updated_file_path)

            if cancelled:
                log_both("⏹ Проверка остановлена. Файл сохранён с уже внесёнными изменениями.", "warn")
            else:
                log_both("✅ Проверка завершена! Файл сохранен.", "success")

            self.msg_queue.put({"type": "done", "path": updated_file_path, "changes": changes_summary,
                                 "cancelled": cancelled})

        except Exception as e:
            self.msg_queue.put({"type": "error", "text": str(e)})
        finally:
            try:
                with open(log_file_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(log_lines))
            except Exception:
                pass


if __name__ == '__main__':
    root = tk.Tk()
    app = DocCheckerApp(root)
    root.mainloop()
