# main.py (версия с перезапуском меню после закрытия модуля)

import sys
import tkinter as tk
from tkinter import messagebox

# Импортируем основные классы из других скриптов
import settings # Класс: SettingsGUI
import search # Класс: GUIApp
import editor # Класс: EditorApp
import upload # Класс: App
import OZR # Класс: ProductionJournalEditor
import ZVK # Класс: IncomingJournalEditor

def run_settings():
    try:
        settings.SettingsGUI().mainloop() # Запуск GUI настроек
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось запустить настройки: {e}")

def run_search():
    try:
        search.GUIApp() # Запуск GUI поиска
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось запустить поиск: {e}")

def run_editor():
    try:
        editor.EditorApp() # Запуск GUI редактора
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось запустить редактор: {e}")

def run_upload():
    try:
        upload.App().mainloop() # Запуск GUI выгрузки
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось запустить выгрузку: {e}")

def run_OZR():
    """Запуск редактора журналов производства работ"""
    try:
        root_ozr = tk.Tk()
        OZR.ProductionJournalEditor(root_ozr)
        root_ozr.mainloop()
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось запустить ОЖР: {e}")

def run_ZVK():
    """Запуск редактора журналов входного контроля"""
    try:
        root_zvk = tk.Tk()
        ZVK.IncomingJournalEditor(root_zvk)
        root_zvk.mainloop()
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось запустить ЖВК: {e}")

def main_menu():
    """Функция главного меню с рекурсивным вызовом после модуля"""
    root = tk.Tk()
    root.title("Главное меню")
    root.resizable(False, False)

    # Заголовок
    title_label = tk.Label(root, text="Главное меню", font=("Arial", 16, "bold"))
    title_label.pack(pady=10)

    buttons_config = [
        ("Настройки", run_settings),
        ("Запустить поиск", run_search),
        ("Редактор значений", run_editor),
        ("Выгрузка данных", run_upload),
        ("", None),  # Разделитель
        ("ОЖР (Журнал производства работ)", run_OZR),
        ("ЖВК (Журнал входного контроля)", run_ZVK),
    ]

    for text, func in buttons_config:
        if text == "":  # Разделитель
            separator = tk.Frame(root, height=2, bg="gray")
            separator.pack(fill=tk.X, padx=20, pady=5)
        else:
            tk.Button(root, text=text, width=30, height=2,
                     command=lambda f=func: [root.destroy(), f(), main_menu()]).pack(padx=20, pady=5)

    # Кнопка выхода
    tk.Button(root, text="Выход", width=30, height=2,
             bg="#ff6b6b", fg="white", font=("Arial", 10, "bold"),
             command=root.quit).pack(pady=20)

    root.mainloop()

# Запуск приложения
if __name__ == "__main__":
    main_menu()
