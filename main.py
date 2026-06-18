import os
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document
from docx.shared import RGBColor
from playwright.sync_api import sync_playwright

# Константы для авторизации
LOGIN = "eselezneva"
PASSWORD = "RZiPbrQA"
BASE_URL = "https://edu.donstu.ru/WebApp/#"

# Полный словарь замен для имен
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


def get_name_variations(fio):
    variations = [fio]
    for key, val in NAME_REPLACEMENTS.items():
        if key in fio:
            variations.append(fio.replace(key, val))
        elif val in fio:
            variations.append(fio.replace(val, key))
    return list(dict.fromkeys(variations))


class DocCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Автоматизация проверки студентов ДОНСТУ")
        self.root.geometry("650x670")
        self.root.configure(padx=25, pady=25, bg="#FFFFFF")

        self.file_path = None
        self.msg_queue = queue.Queue()

        # Переменные для хранения данных последнего отчета
        self.last_changes = None
        self.last_file_path = None

        # --- Настройка стилей для элементов TTK (Прогресс-бар) ---
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("Orange.Horizontal.TProgressbar",
                             troughcolor='#F5F5F5',
                             background='#FF6F00',
                             bordercolor='#FFFFFF',
                             lightcolor='#FF6F00',
                             darkcolor='#FF6F00')

        # --- UI Элементы ---

        # Заголовок
        self.lbl_title = tk.Label(root, text="Проверка списков ДГТУ", font=("Segoe UI", 18, "bold"), bg="#FFFFFF",
                                  fg="#212121")
        self.lbl_title.pack(anchor=tk.W, pady=(0, 2))

        self.lbl_subtitle = tk.Label(root, text="Выберите документ .docx для запуска сверки с базой данных",
                                     font=("Segoe UI", 10), bg="#FFFFFF", fg="#757575")
        self.lbl_subtitle.pack(anchor=tk.W, pady=(0, 20))

        # Файловый блок
        self.frame_file = tk.Frame(root, bg="#FFFFFF")
        self.frame_file.pack(fill=tk.X, pady=5)

        self.btn_select = tk.Button(self.frame_file, text="Выбрать файл", command=self.select_file, width=15,
                                    font=("Segoe UI", 10, "bold"), bg="#333333", fg="#FFFFFF",
                                    activebackground="#444444", activeforeground="#FFFFFF", relief="flat", bd=0, pady=5)
        self.btn_select.pack(side=tk.LEFT, padx=(0, 15))

        self.lbl_file_name = tk.Label(self.frame_file, text="Файл не выбран", font=("Segoe UI", 10, "italic"),
                                      bg="#FFFFFF", fg="#9E9E9E")
        self.lbl_file_name.pack(side=tk.LEFT, fill=tk.X)

        # Опции (Галочки)
        self.frame_options = tk.Frame(root, bg="#FFFFFF")
        self.frame_options.pack(fill=tk.X, pady=(10, 0))

        self.smart_group_var = tk.BooleanVar(value=False)
        self.chk_smart_group = tk.Checkbutton(
            self.frame_options,
            text="Умная замена группы (проверять совпадение первой буквы)",
            variable=self.smart_group_var,
            font=("Segoe UI", 10),
            bg="#FFFFFF",
            fg="#333333",
            activebackground="#FFFFFF",
            activeforeground="#333333",
            selectcolor="#FFFFFF"
        )
        self.chk_smart_group.pack(anchor=tk.W)

        # Контейнер для кнопок управления
        self.frame_buttons = tk.Frame(root, bg="#FFFFFF")
        self.frame_buttons.pack(fill=tk.X, pady=15)

        # Кнопка старта
        self.btn_start = tk.Button(self.frame_buttons, text="Начать сверку", bg="#FF6F00", fg="#FFFFFF",
                                   font=("Segoe UI", 12, "bold"), command=self.start_processing, state=tk.DISABLED,
                                   activebackground="#E65C00", activeforeground="#FFFFFF", relief="flat", bd=0, pady=7)
        self.btn_start.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        # Кнопка повторного открытия отчета
        self.btn_show_last = tk.Button(self.frame_buttons, text="Показать прошлый отчет", bg="#333333", fg="#FFFFFF",
                                       font=("Segoe UI", 12, "bold"), command=self.reopen_summary, state=tk.DISABLED,
                                       activebackground="#444444", activeforeground="#FFFFFF", relief="flat", bd=0,
                                       pady=7)
        self.btn_show_last.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Прогресс-бар и статус
        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, mode='determinate',
                                        style="Orange.Horizontal.TProgressbar")
        self.progress.pack(fill=tk.X, pady=(0, 8))

        self.lbl_status = tk.Label(root, text="Ожидание запуска...", font=("Segoe UI", 10, "bold"), bg="#FFFFFF",
                                   fg="#FF6F00")
        self.lbl_status.pack(anchor=tk.W, pady=(0, 15))

        # Окно логов
        self.log_text = tk.Text(root, height=10, bg="#F9F9F9", fg="#333333", font=("Courier New", 10),
                                relief="flat", bd=1, highlightbackground="#E0E0E0", highlightcolor="#FF6F00",
                                highlightthickness=1, padx=10, pady=10, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Инфо блок
        info_text = (
            "• Скрипт проверит каждого студента и сохранит исправленный файл рядом с оригиналом.\n"
            "• Имена с 'е' и 'ё' проверяются автоматически. В итоговый файл запишется вариант с сайта.\n"
            "• Ненайденные студенты будут выделены красным цветом в документе."
        )
        self.lbl_info = tk.Label(root, text=info_text, justify=tk.LEFT, bg="#FFFFFF", fg="#757575",
                                 font=("Segoe UI", 9), anchor=tk.W)
        self.lbl_info.pack(fill=tk.X, pady=(15, 0))

        # Запуск проверки очереди сообщений
        self.check_queue()

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Выберите документ",
            filetypes=[("Word Documents", "*.docx")]
        )
        if file_path:
            self.file_path = file_path
            self.lbl_file_name.config(text=os.path.basename(file_path), fg="#333333", font=("Segoe UI", 10, "bold"))
            self.btn_start.config(state=tk.NORMAL)

    def log(self, message):
        """Отправка лога в queue GUI"""
        print(message)
        self.msg_queue.put({"type": "log", "text": message})

    def update_status(self, text, percent=0):
        """Отправка статуса в queue GUI"""
        self.msg_queue.put({"type": "progress", "text": text, "value": percent})

    def check_queue(self):
        """Регулярная проверка очереди на наличие сообщений из фонового потока"""
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                if msg["type"] == "log":
                    self.log_text.config(state=tk.NORMAL)
                    self.log_text.insert(tk.END, msg["text"] + "\n")
                    self.log_text.see(tk.END)
                    self.log_text.config(state=tk.DISABLED)

                elif msg["type"] == "progress":
                    self.progress["value"] = msg["value"]
                    self.lbl_status.config(text=msg["text"])

                elif msg["type"] == "done":
                    self.progress["value"] = 100
                    self.lbl_status.config(text="Проверка успешно завершена!")

                    # Сохраняем копию результатов в память приложения
                    self.last_changes = msg["changes"]
                    self.last_file_path = msg["path"]
                    # Делаем кнопку отчета активной
                    self.btn_show_last.config(state=tk.NORMAL)

                    self.show_summary_window(msg["changes"], msg["path"])

                    self.btn_select.config(state=tk.NORMAL)
                    self.btn_start.config(state=tk.NORMAL)

                elif msg["type"] == "error":
                    self.lbl_status.config(text="Ошибка!")
                    messagebox.showerror("Ошибка", msg["text"])
                    self.btn_select.config(state=tk.NORMAL)
                    self.btn_start.config(state=tk.NORMAL)
        except queue.Empty:
            pass

        self.root.after(100, self.check_queue)

    def show_summary_window(self, changes, file_path):
        """Создает всплывающее окно с отчетом об изменениях в фирменном стиле"""
        summary_win = tk.Toplevel(self.root)
        summary_win.title("Отчет об изменениях")
        summary_win.geometry("550x420")
        summary_win.configure(padx=20, pady=20, bg="#FFFFFF")

        lbl = tk.Label(summary_win, text="Внесенные изменения (группы и факультеты):", font=("Segoe UI", 12, "bold"),
                       bg="#FFFFFF", fg="#212121")
        lbl.pack(anchor=tk.W, pady=(0, 10))

        text_area = tk.Text(summary_win, wrap=tk.WORD, font=("Courier New", 10), bg="#F9F9F9", fg="#333333",
                            relief="flat", bd=1, highlightbackground="#E0E0E0", highlightthickness=1, padx=8, pady=8)
        text_area.pack(fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(text_area, command=text_area.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_area.config(yscrollcommand=scrollbar.set)

        if not changes:
            text_area.insert(tk.END, "Группы и факультеты остались без изменений.\n")
        else:
            for change in changes:
                text_area.insert(tk.END, f"• {change}\n")

        text_area.config(state=tk.DISABLED)

        lbl_file = tk.Label(summary_win, text=f"Файл успешно сохранен как:\n{os.path.basename(file_path)}",
                            fg="#FF6F00",
                            justify=tk.LEFT, bg="#FFFFFF", font=("Segoe UI", 10, "bold"))
        lbl_file.pack(anchor=tk.W, pady=(15, 0))

        btn_close = tk.Button(summary_win, text="Закрыть", command=summary_win.destroy, width=15,
                              font=("Segoe UI", 10, "bold"), bg="#333333", fg="#FFFFFF",
                              activebackground="#444444", activeforeground="#FFFFFF", relief="flat", bd=0, pady=5)
        btn_close.pack(anchor=tk.E, pady=(10, 0))

    def reopen_summary(self):
        """Позволяет повторно открыть окно последнего отчета"""
        if self.last_changes is not None and self.last_file_path is not None:
            self.show_summary_window(self.last_changes, self.last_file_path)

    def start_processing(self):
        if not self.file_path:
            return

        self.btn_select.config(state=tk.DISABLED)
        self.btn_start.config(state=tk.DISABLED)
        self.btn_show_last.config(state=tk.DISABLED)

        # Считываем значение умной проверки группы перед запуском потока
        use_smart_group = self.smart_group_var.get()

        # Сброс кэша прошлого отчета
        self.last_changes = None
        self.last_file_path = None

        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        self.progress["value"] = 0
        self.update_status("Инициализация браузера...")

        threading.Thread(target=self.process_file_worker, args=(self.file_path, use_smart_group), daemon=True).start()

    def process_file_worker(self, file_path, use_smart_group):
        try:
            self.log("Открытие документа...")
            doc = Document(file_path)
            if not doc.tables:
                self.msg_queue.put({"type": "error", "text": "В документе не найдены таблицы."})
                return

            target_table = None
            headers = []
            for tbl in doc.tables:
                if len(tbl.rows) == 0:
                    continue

                current_headers = [cell.text.strip().lower().replace(" ", "") for cell in tbl.rows[0].cells]
                if any('фио' in h or 'ф.и.о' in h for h in current_headers):
                    target_table = tbl
                    headers = current_headers
                    break

            if not target_table:
                self.msg_queue.put(
                    {"type": "error", "text": "В документе не найдена таблица со списком студентов (нет колонки ФИО)."})
                return

            table = target_table
            fio_idx, group_idx, fac_idx = None, None, None

            for i, h in enumerate(headers):
                if 'фио' in h or 'ф.и.о' in h:
                    fio_idx = i
                elif 'групп' in h:
                    group_idx = i
                elif 'факультет' in h:
                    fac_idx = i

            if fio_idx is None: fio_idx = 1
            if group_idx is None: group_idx = 2

            total_rows = len(table.rows) - 1
            changes_summary = []

            with sync_playwright() as p:
                self.log("Запуск браузера в фоновом (скрытом) режиме...")
                self.update_status("Запуск браузера...")

                browser = p.chromium.launch(headless=True, channel="msedge")
                context = browser.new_context()
                page = context.new_page()

                try:
                    page.goto(BASE_URL)
                    page.locator('input[name="login"]').fill(LOGIN)
                    page.locator('input[name="password"]').fill(PASSWORD)
                    page.locator('input[name="password"]').press("Enter")

                    search_input = page.locator("input[placeholder='Поиск']")
                    search_input.wait_for(state="visible", timeout=15000)
                    self.log("Авторизация успешна. Начинаем сверку.")

                    for idx, row in enumerate(table.rows[1:]):
                        if len(row.cells) <= max(fio_idx, group_idx):
                            continue

                        fio_cell = row.cells[fio_idx]
                        fio = fio_cell.text.strip()
                        doc_group = row.cells[group_idx].text.strip()
                        doc_faculty = row.cells[fac_idx].text.strip() if fac_idx is not None else None

                        if not fio or fio == doc_group:
                            continue

                        percent = int(((idx + 1) / total_rows) * 100)
                        self.update_status(f"Обработка: {idx + 1} из {total_rows} ({fio})", percent)

                        self.log("-" * 40)
                        self.log(f"[{idx + 1}/{total_rows}] Обработка: {fio}")

                        variations = get_name_variations(fio)
                        student_found = False
                        site_fio = None

                        for current_fio in variations:
                            search_input.fill("")
                            search_input.fill(current_fio)

                            listbox_options = page.locator("div[role='listbox'] div[role='option']")

                            try:
                                listbox_options.first.wait_for(state="visible", timeout=3000)
                                page.wait_for_timeout(1000)

                                options_count = listbox_options.count()
                                if options_count > 0:
                                    target_option = listbox_options.nth(options_count - 1)
                                    raw_text = target_option.inner_text().strip().split('\n')[0]

                                    site_fio = raw_text.split('(')[0].strip()

                                    target_option.click()
                                    student_found = True
                                    break
                            except Exception:
                                continue

                        if not student_found:
                            self.log(f"❌ Варианты для '{fio}' не найдены. Выделяем красным.")
                            for paragraph in fio_cell.paragraphs:
                                for run in paragraph.runs:
                                    run.font.color.rgb = RGBColor(255, 0, 0)

                            page.goto(BASE_URL)
                            search_input.wait_for(state="visible")
                            continue

                        try:
                            group_xpath = "//label[text()='Группа' or contains(text(), 'Группа')]/following-sibling::input"
                            page.locator(group_xpath).wait_for(state="attached", timeout=5000)
                            site_group = page.locator(group_xpath).input_value().strip()

                            site_faculty = None
                            if fac_idx is not None:
                                faculty_xpath = "//label[text()='Факультет' or contains(text(), 'Факультет')]/following-sibling::input"
                                site_faculty = page.locator(faculty_xpath).input_value().strip()

                            # --- Проверки и перезапись данных ---

                            # 1. Сверяем ФИО
                            if site_fio and site_fio != fio:
                                self.log(f"🔄 Обновляем ФИО: {fio} -> {site_fio}")
                                row.cells[fio_idx].text = site_fio

                            # 2. Сверяем Группу (С учетом «умной» проверки)
                            if doc_group != site_group:
                                should_change_group = True

                                # Если в документе пусто, всегда берем значение с сайта без проверок первой буквы
                                if not doc_group:
                                    should_change_group = True
                                elif use_smart_group:
                                    if site_group and doc_group[0].upper() == site_group[0].upper():
                                        should_change_group = True
                                    else:
                                        should_change_group = False
                                        self.log(
                                            f"⚠️ Пропущено изменение группы для {site_fio or fio}. Первая буква не совпала: файл [{doc_group}] vs сайт [{site_group}].")

                                if should_change_group:
                                    display_doc_group = doc_group if doc_group else "пусто"
                                    self.log(f"🔄 Обновляем группу: {display_doc_group} -> {site_group}")
                                    changes_summary.append(
                                        f"{site_fio or fio}: Группа [{display_doc_group}] ➔ [{site_group}]")
                                    row.cells[group_idx].text = site_group

                            # 3. Сверяем Факультет
                            if fac_idx is not None and doc_faculty != site_faculty:
                                display_doc_faculty = doc_faculty if doc_faculty else "пусто"
                                self.log(f"🔄 Обновляем факультет: {display_doc_faculty} -> {site_faculty}")
                                changes_summary.append(
                                    f"{site_fio or fio}: Факультет [{display_doc_faculty}] ➔ [{site_faculty}]")
                                row.cells[fac_idx].text = site_faculty

                        except Exception as e:
                            self.log(f"⚠️ Ошибка при сборе данных {fio}: {e}")

                        page.goto(BASE_URL)
                        search_input.wait_for(state="visible")

                finally:
                    browser.close()

            # Сохранение файла
            name, ext = os.path.splitext(file_path)
            updated_file_path = f"{name}_ОБНОВЛЕННЫЙ{ext}"
            doc.save(updated_file_path)

            self.log("✅ Проверка завершена! Файл сохранен.")

            self.msg_queue.put({
                "type": "done",
                "path": updated_file_path,
                "changes": changes_summary
            })

        except Exception as e:
            self.msg_queue.put({"type": "error", "text": str(e)})


if __name__ == '__main__':
    root = tk.Tk()
    app = DocCheckerApp(root)
    root.mainloop()