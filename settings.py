import json
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

CONFIG_PATH = Path('config.json')

TOOLTIP_BG = "lightyellow"
TOOLTIP_FONT = ("Consolas", 8)


class ToolTip:
    """Простой tooltip для виджетов."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.on_enter)
        self.widget.bind("<Leave>", self.on_leave)

    def on_enter(self, event=None):
        if not self.text:
            return
        try:
            x, y, _, _ = self.widget.bbox("insert")
        except Exception:
            x = y = 0
        x += self.widget.winfo_rootx() + 20
        y += self.widget.winfo_rooty() + 20
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self.tooltip,
            text=self.text,
            background=TOOLTIP_BG,
            relief="solid",
            borderwidth=1,
            font=TOOLTIP_FONT,
        )
        label.pack()

    def on_leave(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None


class SettingsGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Настройки парсинга")
        self.geometry("1700x550")

        self.root_dir = tk.StringVar()

        self._create_root_dir_section()
        self._create_rules_section()
        self._create_save_button()

        self.rows = []  # Список фреймов строк
        self._load_config()

    def _create_root_dir_section(self):
        """Создает секцию для выбора корневой папки."""
        tk.Label(self, text="Шаг 1. Выберите корневую папку").pack()
        entry = tk.Entry(self, textvariable=self.root_dir, font=("Consolas", 10))
        entry.pack(fill='x', padx=10)
        self._setup_entry_bindings(entry)
        tk.Button(self, text="Выбрать…", command=self._browse_dir).pack(pady=3)

    def _create_rules_section(self):
        """Создает секцию для списка правил."""
        frm = ttk.LabelFrame(self, text="Шаг 2. Файлы и ключевые слова")
        frm.pack(fill='both', expand=True, padx=5, pady=5)

        # Заголовки колонок
        header_frame = tk.Frame(frm)
        header_frame.pack(fill='x', padx=5, pady=2)
        headers = [
            ("Наименование данных", 70),
            ("Путь к файлу", 40),
            ("Имя файла", 25),
            (" ", 3),  # Пустая для кнопки
            ("Тип", 8),
            ("Ключевые слова", 25),
            ("Группа", 10),
            ("Действия", 15),
        ]
        for text, width in headers:
            tk.Label(header_frame, text=text, width=width, anchor='w').pack(side='left', padx=2)

        self.rows_container = tk.Canvas(frm)
        self.rows_container.pack(side='left', fill='both', expand=True)
        scrollbar = ttk.Scrollbar(frm, orient='vertical', command=self.rows_container.yview)
        scrollbar.pack(side='right', fill='y')
        self.rows_container.configure(yscrollcommand=scrollbar.set)

        self.rows_inner = tk.Frame(self.rows_container)
        self.rows_container.create_window((0, 0), window=self.rows_inner, anchor='nw')
        self.rows_inner.bind('<Configure>', lambda e: self.rows_container.configure(scrollregion=self.rows_container.bbox('all')))

        ttk.Button(self, text="+ добавить строку", command=self._add_row).pack(pady=5)

    def _create_save_button(self):
        """Создает кнопку сохранения."""
        tk.Button(self, text="Сохранить", command=self._save).pack(pady=8)

    def _setup_entry_bindings(self, entry_widget):
        """Настраивает стандартные горячие клавиши для Entry."""
        bindings = {
            '<Control-a>': self._select_all,
            '<Control-A>': self._select_all,
            '<Control-c>': self._copy_text,
            '<Control-C>': self._copy_text,
            '<Control-v>': self._paste_text,
            '<Control-V>': self._paste_text,
            '<Control-x>': self._cut_text,
            '<Control-X>': self._cut_text,
            '<Command-a>': self._select_all,
            '<Command-A>': self._select_all,
            '<Command-c>': self._copy_text,
            '<Command-C>': self._copy_text,
            '<Command-v>': self._paste_text,
            '<Command-V>': self._paste_text,
            '<Command-x>': self._cut_text,
            '<Command-X>': self._cut_text,
        }
        for key, func in bindings.items():
            entry_widget.bind(key, func)

    def _select_all(self, event):
        event.widget.select_range(0, tk.END)
        return 'break'

    def _copy_text(self, event):
        try:
            text = event.widget.selection_get()
            event.widget.clipboard_clear()
            event.widget.clipboard_append(text)
        except tk.TclError:
            pass
        return 'break'

    def _paste_text(self, event):
        try:
            text = event.widget.clipboard_get()
            cursor_pos = event.widget.index(tk.INSERT)
            try:
                start = event.widget.index(tk.SEL_FIRST)
                end = event.widget.index(tk.SEL_LAST)
                event.widget.delete(start, end)
                cursor_pos = start
            except tk.TclError:
                pass
            event.widget.insert(cursor_pos, text)
        except tk.TclError:
            pass
        return 'break'

    def _cut_text(self, event):
        try:
            text = event.widget.selection_get()
            event.widget.clipboard_clear()
            event.widget.clipboard_append(text)
            event.widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            pass
        return 'break'

    def _load_config(self):
        """Загружает сохраненную конфигурацию."""
        if not CONFIG_PATH.exists():
            self._add_row()
            return
        try:
            with CONFIG_PATH.open(encoding='utf-8') as f:
                cfg = json.load(f)
            self.root_dir.set(cfg.get('root', ''))
            items = cfg.get('items', [])
            if items:
                for item in items:
                    data_name = item.get('data_name', '')
                    file_path = item.get('file', '')
                    keywords = ', '.join(item.get('keywords', []))
                    group = item.get('group', '')
                    p = Path(file_path)
                    dir_part = p.parent.as_posix() if p.parent.as_posix() != '.' else ''
                    name_part = p.name
                    self._add_row(data_name=data_name, dir_path=dir_part, filename=name_part, full_path=file_path, kw=keywords, group=group)
            else:
                self._add_row()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить конфигурацию: {e}")
            self._add_row()

    def _get_file_type(self, file_path):
        """Определяет тип файла по расширению."""
        if not file_path:
            return "не выбран"
        ext = Path(file_path).suffix.lower()
        if ext in ['.doc', '.docx']:
            return 'Word'
        if ext in ['.xls', '.xlsx']:
            return 'Excel'
        if ext == '.pdf':
            return 'PDF'
        return f'неизвестный ({ext})'

    def _browse_dir(self):
        """Выбор корневой директории."""
        d = filedialog.askdirectory()
        if d:
            self.root_dir.set(d)

    def _browse_file(self, dir_entry, file_entry, type_label, row):
        """Выбор файла с учетом группы и обновлением путей."""
        start_dir = self._get_start_dir(dir_entry)
        filetypes = [
            ("Документы", "*.docx *.doc *.xlsx *.xls *.pdf"),
            ("Word документы", "*.docx *.doc"),
            ("Excel файлы", "*.xlsx *.xls"),
            ("PDF файлы", "*.pdf"),
            ("Все файлы", "*.*"),
        ]
        f = filedialog.askopenfilename(initialdir=start_dir, filetypes=filetypes)
        if not f:
            return

        root_dir = self.root_dir.get().strip()
        filename = Path(f).name
        full_path = self._get_relative_path(f, root_dir) or f

        group = self._get_row_group(row)
        rows_to_update = [r for r in self.rows if self._get_row_group(r) == group] if group else [row]

        for r in rows_to_update:
            r_widgets = r.winfo_children()
            r_dir_entry = r_widgets[1]
            r_file_entry = r_widgets[2]
            r_type_label = r_widgets[4]

            if not r_dir_entry.get().strip():
                dir_path = self._get_dir_path(f, root_dir)
                r_dir_entry.delete(0, tk.END)
                r_dir_entry.insert(0, dir_path)

            r_file_entry.delete(0, tk.END)
            r_file_entry.insert(0, filename)
            r_file_entry.full_path = full_path
            self._update_tooltip(r_file_entry, full_path)
            file_type = self._get_file_type(f)
            r_type_label.config(text=file_type)

    def _get_start_dir(self, dir_entry):
        """Определяет стартовую директорию для выбора файла."""
        dir_text = dir_entry.get().strip()
        root = self.root_dir.get().strip()
        if dir_text and root:
            return str((Path(root) / dir_text).resolve())
        return dir_text or root or None

    def _get_relative_path(self, file_path, root_dir):
        """Возвращает относительный путь, если возможно."""
        if not root_dir:
            return None
        try:
            return Path(file_path).resolve().relative_to(Path(root_dir).resolve()).as_posix()
        except ValueError:
            return None

    def _get_dir_path(self, file_path, root_dir):
        """Возвращает директорию файла (относительную или абсолютную)."""
        if root_dir:
            rel = self._get_relative_path(file_path, root_dir)
            if rel:
                return Path(rel).parent.as_posix() if Path(rel).parent.as_posix() != '.' else ''
        return str(Path(file_path).parent)

    def _update_tooltip(self, widget, text):
        """Обновляет или создает tooltip."""
        if hasattr(widget, 'tooltip'):
            widget.tooltip.text = text
        else:
            widget.tooltip = ToolTip(widget, text)

    def _update_file_type_manual(self, file_entry, type_label):
        """Обновляет тип файла при ручном вводе."""
        def on_change(*args):
            filename = file_entry.get()
            file_type = self._get_file_type(filename)
            type_label.config(text=file_type)
            file_entry.full_path = filename
            self._update_tooltip(file_entry, filename) if filename else setattr(file_entry, 'tooltip', None)
        file_entry.bind('<KeyRelease>', lambda e: on_change())

    def _get_row_group(self, row):
        """Получает группу строки."""
        widgets = row.winfo_children()
        return widgets[6].get().strip()

    def _make_row(self, data_name='', dir_path='', filename='', full_path='', kw='', group=''):
        """Создает фрейм строки."""
        row = tk.Frame(self.rows_inner, pady=2)

        # Наименование данных
        data_entry = tk.Entry(row, width=70)
        data_entry.insert(0, data_name)
        data_entry.pack(side='left', padx=2)
        self._setup_entry_bindings(data_entry)

        # Путь к файлу (директория)
        dir_entry = tk.Entry(row, width=40)
        dir_entry.insert(0, dir_path)
        dir_entry.pack(side='left', padx=2)
        self._setup_entry_bindings(dir_entry)

        # Имя файла
        file_entry = tk.Entry(row, width=25)
        file_entry.insert(0, filename)
        file_entry.full_path = full_path
        if full_path:
            self._update_tooltip(file_entry, full_path)
        file_entry.pack(side='left', padx=2)
        self._setup_entry_bindings(file_entry)

        # Кнопка выбора файла
        btn = tk.Button(row, text="…", command=lambda: self._browse_file(dir_entry, file_entry, type_label, row))
        btn.pack(side='left')

        # Тип файла
        type_source = full_path or filename
        type_label = tk.Label(row, text=self._get_file_type(type_source), width=8, anchor='w', relief='sunken')
        type_label.pack(side='left', padx=3)

        # Ключевые слова
        kw_entry = tk.Entry(row, width=25)
        kw_entry.insert(0, kw)
        kw_entry.pack(side='left', padx=4)
        self._setup_entry_bindings(kw_entry)

        # Группа
        group_entry = tk.Entry(row, width=10)
        group_entry.insert(0, group)
        group_entry.pack(side='left', padx=4)
        self._setup_entry_bindings(group_entry)

        # Действия
        actions_frame = tk.Frame(row)
        actions_frame.pack(side='left', padx=2)
        actions = [
            ("↑", lambda: self._move_row_up(row), "Переместить вверх"),
            ("↓", lambda: self._move_row_down(row), "Переместить вниз"),
            ("⧉", lambda: self._duplicate_row(row), "Дублировать строку"),
            ("✕", lambda: self._delete_row(row), "Удалить строку"),
        ]
        for text, cmd, tip in actions:
            btn = tk.Button(actions_frame, text=text, width=2, command=cmd)
            btn.pack(side='left', padx=1)
            ToolTip(btn, tip)

        self._update_file_type_manual(file_entry, type_label)
        return row

    def _add_row(self, **kwargs):
        """Добавляет строку в конец списка."""
        row = self._make_row(**kwargs)
        self.rows.append(row)
        row.pack(fill='x')
        self._refresh_rows()

    def _move_row_up(self, row):
        """Перемещает строку или группу вверх."""
        self._move_row(row, direction=-1)

    def _move_row_down(self, row):
        """Перемещает строку или группу вниз."""
        self._move_row(row, direction=1)

    def _move_row(self, row, direction):
        """Общая логика перемещения строки или группы."""
        group = self._get_row_group(row)
        if group:
            group_rows = [r for r in self.rows if self._get_row_group(r) == group]
            indices = sorted(self.rows.index(r) for r in group_rows)
            if (direction == -1 and indices[0] == 0) or (direction == 1 and indices[-1] == len(self.rows) - 1):
                return
            target_pos = indices[0] + direction if direction == -1 else indices[-1] + direction - len(indices) + 1
            removed = [self.rows.pop(i) for i in reversed(indices)]
            for i, r in enumerate(removed):
                self.rows.insert(target_pos + i, r)
        else:
            idx = self.rows.index(row)
            if (direction == -1 and idx == 0) or (direction == 1 and idx == len(self.rows) - 1):
                return
            self.rows.insert(idx + direction, self.rows.pop(idx))
        self._refresh_rows()

    def _duplicate_row(self, source_row):
        """Дублирует строку и вставляет под исходной."""
        idx = self.rows.index(source_row)
        widgets = source_row.winfo_children()
        params = {
            'data_name': widgets[0].get(),
            'dir_path': widgets[1].get(),
            'filename': widgets[2].get(),
            'full_path': getattr(widgets[2], 'full_path', ''),
            'kw': widgets[5].get(),
            'group': widgets[6].get(),
        }
        new_row = self._make_row(**params)
        self.rows.insert(idx + 1, new_row)
        self._refresh_rows()

    def _delete_row(self, row):
        """Удаляет строку."""
        if row in self.rows:
            self.rows.remove(row)
            row.destroy()
            self._refresh_rows()

    def _refresh_rows(self):
        """Перерисовывает все строки."""
        for r in self.rows_inner.winfo_children():
            r.pack_forget()
        for r in self.rows:
            r.pack(fill='x', pady=2)
        self.rows_container.configure(scrollregion=self.rows_container.bbox('all'))

    def _save(self):
        """Сохраняет конфигурацию без предупреждений."""
        root_dir = self.root_dir.get().strip()
        if not root_dir:
            messagebox.showerror("Ошибка", "Не выбрана корневая папка")
            return

        root_path = Path(root_dir).resolve()
        items = []

        for row in self.rows:
            widgets = row.winfo_children()
            data_name = widgets[0].get().strip()
            dir_text = widgets[1].get().strip()
            file_entry = widgets[2]
            kw = widgets[5].get().strip()
            group = widgets[6].get().strip()

            full_path = getattr(file_entry, 'full_path', '') or (Path(dir_text) / file_entry.get()).as_posix() if file_entry.get() else ''

            relative_path = full_path
            ftype = ''
            abs_path = None

            if full_path:
                try:
                    abs_path = (root_path / full_path).resolve() if not Path(full_path).is_absolute() else Path(full_path).resolve()
                    relative_path = abs_path.relative_to(root_path).as_posix()
                except ValueError:
                    relative_path = full_path  # Вне корня — сохраняем как есть
                except Exception:
                    relative_path = full_path  # Игнорируем ошибки

            if abs_path:
                ftype = self._get_file_type(abs_path).lower()
                if ftype not in ['word', 'excel', 'pdf']:
                    ftype = 'unknown' if ftype != 'не выбран' else ''

            items.append({
                "data_name": data_name,
                "file": relative_path,
                "type": ftype,
                "keywords": [k.strip() for k in kw.split(',') if k.strip()],
                "group": group
            })

        if not items:
            messagebox.showerror("Ошибка", "Нет строк для сохранения")
            return

        cfg = {"root": root_dir, "items": items}
        CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=4), encoding='utf-8')
        messagebox.showinfo("Готово", f"config.json сохранён. Обработано строк: {len(items)}")
        self.destroy()


if __name__ == "__main__":
    SettingsGUI().mainloop()
