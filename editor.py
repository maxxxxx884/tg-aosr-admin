import json
import os
import sys
import subprocess
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

# Константы
CONFIG_FILE = Path('config.json')
DATA_FILE = Path('data.json')
REPORT_FILE = Path('report.txt')

STATUS_FOUND = 'found'
STATUS_NOT_FOUND = 'not_found'
STATUS_UNKNOWN = 'unknown'

ROW_TAG_NOT_FOUND = 'row_not_found'


class DataManager:
    """Отвечает за загрузку/сохранение данных, валидацию и отчёты."""

    def __init__(self, config_path: Path, data_path: Path):
        self.config_path = config_path
        self.data_path = data_path
        self.config = {}
        self.root_path = Path('.').resolve()
        self.data = []

    def load_config(self) -> None:
        if not self.config_path.exists():
            raise FileNotFoundError(f"Файл {self.config_path} не найден.")
        with self.config_path.open('r', encoding='utf-8') as f:
            self.config = json.load(f)
        # Нормализация root
        root = self.config.get('root', '.')
        self.root_path = Path(root).expanduser().resolve()
        if not self.root_path.exists():
            # Не падаем, но предупреждаем через исключение с особым типом
            raise FileNotFoundError(f"Путь root не существует: {self.root_path}")

    def load_data(self) -> None:
        if not self.data_path.exists():
            # Предложим создать пустой
            raise FileNotFoundError(f"Файл {self.data_path} не найден.")
        with self.data_path.open('r', encoding='utf-8') as f:
            self.data = json.load(f)
        if not isinstance(self.data, list):
            raise ValueError("Ожидался список объектов в data.json")

    def backup_data(self) -> Path:
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup = self.data_path.with_suffix(f'.backup_{ts}.json')
        with backup.open('w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=4)
        return backup

    @staticmethod
    def compute_status(value: str) -> str:
        val = (value or '').strip()
        if val and val.lower() != 'null':
            return STATUS_FOUND
        return STATUS_NOT_FOUND

    def save_data(self, data: list) -> None:
        # Резервное копирование на всякий случай
        try:
            self.backup_data()
        except Exception:
            # Бекап не обязателен, игнорируем ошибки бекапа
            pass
        with self.data_path.open('w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

    def generate_report_text(self, items: list) -> str:
        lines = []
        lines.append("ОТЧЕТ О ОТСУТСТВУЮЩИХ ЗНАЧЕНИЯХ")
        lines.append("=" * 50)
        lines.append(f"Сформирован: {datetime.now():%Y-%m-%d %H:%M:%S}")
        lines.append(f"Root: {self.root_path}")
        lines.append("")
        lines.append(f"Всего не найдено: {len(items)}")
        lines.append("")
        for i, item in enumerate(items, 1):
            lines.append(f"{i}. Data Name: {item.get('data_name', 'N/A')}")
            lines.append(f"   Файл: {item.get('file', 'N/A')}")
            lines.append(f"   Статус: {item.get('status', STATUS_NOT_FOUND)}")
            lines.append(f"   Extracted Value: {item.get('extracted_value', 'null')}")
            reason = item.get('reason', 'Причина не указана')
            lines.append(f"   Почему отсутствует: {reason}")
            kw = item.get('keywords', None)
            if isinstance(kw, (list, tuple)):
                lines.append(f"   Ключевые слова: {', '.join(map(str, kw))}")
            elif kw is not None:
                lines.append(f"   Ключевые слова: {kw}")
            lines.append("")
        return "\n".join(lines)

    def save_report(self, text: str, report_path: Path) -> None:
        with report_path.open('w', encoding='utf-8') as f:
            f.write(text)


class EditorApp:
    """GUI-редактор на Tkinter."""

    def __init__(self):
        self.dm = DataManager(CONFIG_FILE, DATA_FILE)

        self.root = tk.Tk()
        self.root.title("Document AI Editor")
        self.root.geometry("1100x760")
        self.root.minsize(900, 640)

        # Состояние
        self.data = []
        self.filtered_indices = []  # список индексов self.data, соответствующих текущей фильтрации/сорту
        self.selected_index = None  # индекс в self.data
        self.sort_state = {}  # column_id -> asc: bool
        self.unsaved_edit = False

        # UI переменные
        self.search_var = tk.StringVar()
        self.status_filter_var = tk.StringVar(value="all")
        self.edit_var = tk.StringVar()

        # Загрузка
        self.safe_load_all()

        # UI
        self.create_widgets()
        self.populate_tree()

        # Закрытие с подтверждением
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.log_info("Редактор запущен. Выберите элемент в таблице для редактирования.")

        self.root.mainloop()

    def safe_load_all(self):
        # Конфиг
        try:
            self.dm.load_config()
        except FileNotFoundError as e:
            messagebox.showerror("Ошибка", str(e))
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить config.json: {e}")

        # Данные
        try:
            self.dm.load_data()
            self.data = self.dm.data
        except FileNotFoundError:
            if messagebox.askyesno("Вопрос", f"{DATA_FILE} не найден. Создать пустой файл?"):
                self.data = []
                try:
                    self.dm.save_data(self.data)
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось создать {DATA_FILE}: {e}")
            else:
                self.data = []
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить data.json: {e}")
            self.data = []

    def create_widgets(self):
        # Топ панель: фильтр и поиск
        top = ttk.Frame(self.root)
        top.pack(fill=tk.X, padx=12, pady=(10, 6))

        ttk.Label(top, text="Статус:").pack(side=tk.LEFT)
        status_combo = ttk.Combobox(
            top,
            textvariable=self.status_filter_var,
            state="readonly",
            width=14,
            values=["all", STATUS_FOUND, STATUS_NOT_FOUND, STATUS_UNKNOWN]
        )
        status_combo.pack(side=tk.LEFT, padx=(6, 12))
        status_combo.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        ttk.Label(top, text="Поиск:").pack(side=tk.LEFT)
        search_entry = ttk.Entry(top, textvariable=self.search_var, width=36)
        search_entry.pack(side=tk.LEFT, padx=(6, 6))
        search_entry.bind("<Return>", lambda e: self.apply_filters())

        ttk.Button(top, text="Применить", command=self.apply_filters).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(top, text="Сбросить", command=self.reset_filters).pack(side=tk.LEFT)

        ttk.Button(top, text="Сформировать отчет", command=self.generate_report).pack(side=tk.RIGHT)

        # Таблица
        tree_frame = ttk.Frame(self.root)
        tree_frame.pack(padx=12, pady=6, fill=tk.BOTH, expand=True)

        y_scroll = ttk.Scrollbar(tree_frame, orient='vertical')
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        x_scroll = ttk.Scrollbar(tree_frame, orient='horizontal')
        x_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=('data_name', 'extracted_value', 'status', 'file'),
            show='headings',
            yscrollcommand=y_scroll.set,
            xscrollcommand=x_scroll.set,
            selectmode="browse"
        )
        y_scroll.config(command=self.tree.yview)
        x_scroll.config(command=self.tree.xview)

        # Заголовки
        self.tree.heading('data_name', text='Data Name', command=lambda: self.sort_by('data_name'))
        self.tree.heading('extracted_value', text='Extracted Value', command=lambda: self.sort_by('extracted_value'))
        self.tree.heading('status', text='Статус', command=lambda: self.sort_by('status'))
        self.tree.heading('file', text='Файл', command=lambda: self.sort_by('file'))

        # Ширины
        self.tree.column('data_name', width=220, anchor='w')
        self.tree.column('extracted_value', width=400, anchor='w')
        self.tree.column('status', width=110, anchor='center')
        self.tree.column('file', width=300, anchor='w')

        # Стили и подсветка
        style = ttk.Style(self.root)
        # Активируем тему по умолчанию и настраиваем цвета тега
        self.tree.tag_configure(ROW_TAG_NOT_FOUND, background="#fff3f3")  # светло-красный фон

        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)

        # Панель редактирования
        edit_frame = ttk.LabelFrame(self.root, text="Редактирование", padding=10)
        edit_frame.pack(padx=12, pady=8, fill=tk.X)

        ttk.Label(edit_frame, text="Выбранный элемент:").grid(row=0, column=0, sticky=tk.W)
        self.selected_label = ttk.Label(edit_frame, text="Не выбрано")
        self.selected_label.grid(row=0, column=1, sticky=tk.W, padx=(8, 0))

        ttk.Label(edit_frame, text="Extracted Value:").grid(row=1, column=0, sticky=tk.W, pady=(8, 0))
        self.edit_entry = ttk.Entry(edit_frame, textvariable=self.edit_var)
        self.edit_entry.grid(row=1, column=1, sticky=tk.EW, padx=(8, 0), pady=(8, 0))
        edit_frame.columnconfigure(1, weight=1)
        self.edit_var.trace_add('write', lambda *args: self.on_edit_changed())

        btns = ttk.Frame(edit_frame)
        btns.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))

        self.save_button = ttk.Button(btns, text="Сохранить в JSON", command=self.save_to_file, state='disabled')
        self.save_button.pack(side=tk.LEFT)

        self.open_file_button = ttk.Button(btns, text="Открыть файл", command=self.open_selected_file, state='disabled')
        self.open_file_button.pack(side=tk.LEFT, padx=(8, 0))

        ttk.Button(btns, text="Отменить изменения", command=self.revert_current).pack(side=tk.LEFT, padx=(8, 0))

        # Превью длинного значения
        preview_frame = ttk.LabelFrame(self.root, text="Превью значения", padding=6)
        preview_frame.pack(padx=12, pady=(0, 8), fill=tk.BOTH)
        self.preview_text = scrolledtext.ScrolledText(preview_frame, height=4, wrap='word')
        self.preview_text.pack(fill=tk.BOTH, expand=True)
        self.preview_text.configure(state='disabled')

        # Логи
        log_frame = ttk.LabelFrame(self.root, text="Логи", padding=6)
        log_frame.pack(padx=12, pady=(0, 12), fill=tk.BOTH)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log(self, level: str, text: str):
        ts = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert(tk.END, f"[{ts}] [{level}] {text}\n")
        self.log_text.see(tk.END)

    def log_info(self, text: str):
        self.log("INFO", text)

    def log_error(self, text: str):
        self.log("ERROR", text)

    def populate_tree(self):
        # Очистить
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Применить фильтры/поиск к self.data, сохранить соответствия индексов
        self.filtered_indices = self.filter_data_indices()

        for pos, idx in enumerate(self.filtered_indices):
            row = self.data[idx]
            values = (
                row.get('data_name', 'N/A'),
                self.shorten_text(row.get('extracted_value', '')),
                row.get('status', STATUS_UNKNOWN),
                row.get('file', 'N/A'),
            )
            tags = []
            if row.get('status') == STATUS_NOT_FOUND:
                tags.append(ROW_TAG_NOT_FOUND)
            iid = f"row-{idx}"  # стабильный iid, завязан на исходный индекс в self.data
            self.tree.insert('', tk.END, iid=iid, values=values, tags=tuple(tags))

        # Сброс выбранного
        self.selected_index = None
        self.selected_label.config(text="Не выбрано")
        self.edit_var.set('')
        self.update_preview('')
        self.save_button.config(state='disabled')
        self.open_file_button.config(state='disabled')

    def filter_data_indices(self):
        # Статус
        status_filter = self.status_filter_var.get()
        # Поиск по data_name/extracted_value/file (без регистра)
        query = (self.search_var.get() or '').strip().lower()

        indices = []
        for i, row in enumerate(self.data):
            if status_filter != "all":
                if row.get('status', STATUS_UNKNOWN) != status_filter:
                    continue
            if query:
                dn = str(row.get('data_name', '')).lower()
                ev = str(row.get('extracted_value', '')).lower()
                fp = str(row.get('file', '')).lower()
                if query not in dn and query not in ev and query not in fp:
                    continue
            indices.append(i)

        # Применить текущую сортировку, если задана
        if self.sort_state:
            col, asc = next(iter(self.sort_state.items()))
            indices.sort(key=lambda i: self.sort_key(self.data[i], col))
            if not asc:
                indices.reverse()
        return indices

    @staticmethod
    def shorten_text(text, limit=120):
        s = str(text or '')
        if len(s) <= limit:
            return s
        return s[:limit - 1] + "…"

    @staticmethod
    def sort_key(row: dict, col: str):
        val = row.get(col, '')
        # Универсальная строковая сортировка
        return str(val).lower()

    def sort_by(self, col: str):
        current = self.sort_state.get(col)
        # Сбросить все, переключить текущую колонку
        self.sort_state.clear()
        if current is None:
            self.sort_state[col] = True  # asc
        else:
            self.sort_state[col] = not current  # toggle
        self.populate_tree()

    def on_tree_select(self, event=None):
        selection = self.tree.selection()
        if not selection:
            return
        iid = selection[0]  # "row-<index>"
        try:
            idx = int(iid.split('-')[-1])
        except Exception:
            return
        if idx < 0 or idx >= len(self.data):
            return

        self.selected_index = idx
        row = self.data[idx]
        self.selected_label.config(text=row.get('data_name', 'N/A'))
        self.edit_var.set(row.get('extracted_value', ''))
        self.update_preview(row.get('extracted_value', ''))
        self.save_button.config(state='normal')

        # Проверим существование файла для кнопки "Открыть файл"
        relative_file = row.get('file', '')
        full_path = self.dm.root_path / relative_file if relative_file else None
        if full_path and full_path.exists():
            self.open_file_button.config(state='normal')
        else:
            self.open_file_button.config(state='disabled')

        self.unsaved_edit = False

    def on_edit_changed(self):
        # Помечаем наличие несохраненных правок
        self.unsaved_edit = True
        # Обновляем превью
        self.update_preview(self.edit_var.get())

    def update_preview(self, text: str):
        self.preview_text.configure(state='normal')
        self.preview_text.delete('1.0', tk.END)
        self.preview_text.insert(tk.END, str(text or ''))
        self.preview_text.configure(state='disabled')

    def open_selected_file(self):
        if self.selected_index is None:
            return
        relative_file = self.data[self.selected_index].get('file', '')
        if not relative_file:
            messagebox.showwarning("Предупреждение", "У выбранного элемента не указан файл.")
            return

        full_path = self.dm.root_path / relative_file
        if not full_path.exists():
            messagebox.showerror("Ошибка", f"Файл не найден: {full_path}")
            self.log_error(f"Файл не найден: {full_path}")
            return

        try:
            if sys.platform.startswith('win'):
                os.startfile(str(full_path))  # nosec
            elif sys.platform.startswith('darwin'):
                subprocess.Popen(['open', str(full_path)])
            else:
                # Linux / POSIX
                try:
                    subprocess.Popen(['xdg-open', str(full_path)])
                except FileNotFoundError:
                    messagebox.showerror("Ошибка", "Не найдена утилита xdg-open для открытия файлов.")
                    return
            self.log_info(f"Открыт файл: {full_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл: {e}")
            self.log_error(f"Ошибка открытия файла: {e}")

    def save_to_file(self):
        if self.selected_index is None:
            messagebox.showwarning("Предупреждение", "Сначала выберите элемент для редактирования.")
            return
        new_value = self.edit_var.get()
        old_value = self.data[self.selected_index].get('extracted_value', '')

        # Обновляем данные и статус
        self.data[self.selected_index]['extracted_value'] = new_value
        self.data[self.selected_index]['status'] = self.dm.compute_status(new_value)

        # Сохраняем
        try:
            self.dm.save_data(self.data)
        except Exception as e:
            self.log_error(f"Ошибка при сохранении: {e}")
            messagebox.showerror("Ошибка", f"Ошибка при сохранении: {e}")
            return

        # Обновляем таблицу без потери фильтрации/сортировки
        self.populate_tree()
        # Переустановим выбор на ту же строку, если она видна после фильтрации
        iid = f"row-{self.selected_index}"
        if iid in self.tree.get_children(''):
            self.tree.selection_set(iid)
            self.tree.see(iid)

        self.log_info(f"Изменено значение: '{old_value}' → '{new_value}'. Сохранено в {DATA_FILE}")
        messagebox.showinfo("Успех", f"Изменения сохранены в {DATA_FILE}")
        self.unsaved_edit = False

    def revert_current(self):
        if self.selected_index is None:
            return
        # Сбросить поле к текущему значению из self.data
        current_value = self.data[self.selected_index].get('extracted_value', '')
        self.edit_var.set(current_value)
        self.update_preview(current_value)
        self.unsaved_edit = False

    def apply_filters(self):
        self.populate_tree()
        self.log_info("Применены фильтры/поиск.")

    def reset_filters(self):
        self.search_var.set('')
        self.status_filter_var.set('all')
        self.sort_state.clear()
        self.populate_tree()
        self.log_info("Сброшены фильтры и сортировка.")

    def generate_report(self):
        not_found = [item for item in self.data if item.get('status') == STATUS_NOT_FOUND]
        if not not_found:
            messagebox.showinfo("Информация", "Не найдено элементов со статусом 'not_found'.")
            self.log_info("Отчет не сформирован: нет не найденных элементов.")
            return

        report_text = self.dm.generate_report_text(not_found)
        try:
            self.dm.save_report(report_text, REPORT_FILE)
            self.log_info(f"Отчет сохранен в {REPORT_FILE}")
            messagebox.showinfo("Успех", f"Отчет успешно сохранен в {REPORT_FILE}")
        except Exception as e:
            self.log_error(f"Ошибка при сохранении отчета: {e}")
            messagebox.showerror("Ошибка", f"Ошибка при сохранении отчета: {e}")

    def on_close(self):
        if self.unsaved_edit:
            if not messagebox.askyesno("Подтверждение", "Есть несохраненные изменения. Выйти без сохранения?"):
                return
        self.root.destroy()


def main():
    EditorApp()


if __name__ == "__main__":
    main()
