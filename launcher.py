import sys
import os
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import urllib.request
import platform
import webbrowser

# --- Константы и конфигурация ---
MIN_PYTHON_VERSION = (3, 7)
PYTHON_DOWNLOAD_URL = "https://www.python.org/downloads/"
APP_FILE = "cost_calculator.py"
REQUIRED_LIBRARIES = ["Pillow", "reportlab"]

FONT_BASE_URL = "https://github.com/dejavu-fonts/dejavu-fonts/raw/main/ttf/"
FONTS_TO_DOWNLOAD = {
    "DejaVuSans.ttf": FONT_BASE_URL + "DejaVuSans.ttf",
    "DejaVuSans-Bold.ttf": FONT_BASE_URL + "DejaVuSans-Bold.ttf"
}


class LauncherStatusWindow(tk.Tk):
    """Окно для отображения статуса проверок и установки."""
    def __init__(self):
        super().__init__()
        self.title("Запуск Калькулятора")
        self.geometry("500x350")
        self.resizable(False, False)

        self.status_text = tk.Text(self, wrap="word", height=15, state="disabled",
                                   bg="#f0f0f0", relief="solid", borderwidth=1, padx=10, pady=10)
        self.status_text.pack(pady=20, padx=20, fill="both", expand=True)

        self.progress = ttk.Progressbar(self, orient="horizontal", length=100, mode="determinate")
        self.progress.pack(pady=10, padx=20, fill="x")
        
        self.close_button = ttk.Button(self, text="Закрыть", state="disabled", command=self.destroy)
        self.close_button.pack(pady=10)

    def log(self, message):
        self.status_text.config(state="normal")
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state="disabled")
        self.update_idletasks()

    def enable_close(self):
        self.close_button.config(state="normal")

def check_and_install_libraries(status_window):
    """Проверяет и устанавливает необходимые библиотеки."""
    status_window.log("Проверка необходимых библиотек...")
    
    total_libs = len(REQUIRED_LIBRARIES)
    for i, lib in enumerate(REQUIRED_LIBRARIES):
        try:
            __import__(lib.split('==')[0])
            status_window.log(f"  - Библиотека '{lib}' уже установлена.")
        except ImportError:
            status_window.log(f"  - Библиотека '{lib}' не найдена. Начинаю установку...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", lib])
                status_window.log(f"  - Библиотека '{lib}' успешно установлена.")
            except subprocess.CalledProcessError:
                error_msg = f"Не удалось установить '{lib}'. Попробуйте установить ее вручную командой:\n\npip install {lib}"
                status_window.log(f"ОШИБКА: {error_msg}")
                messagebox.showerror("Ошибка установки", error_msg)
                return False
        status_window.progress['value'] = (i + 1) * (50 / total_libs)
        status_window.update_idletasks()
    
    status_window.log("Все библиотеки на месте.")
    return True

def check_and_download_fonts(status_window):
    """Проверяет и скачивает необходимые шрифты."""
    status_window.log("\nПроверка шрифтов для PDF...")
    all_fonts_present = True
    
    total_fonts = len(FONTS_TO_DOWNLOAD)
    
    for i, (filename, url) in enumerate(FONTS_TO_DOWNLOAD.items()):
        if os.path.exists(filename):
            status_window.log(f"  - Шрифт '{filename}' найден.")
            continue

        all_fonts_present = False
        status_window.log(f"  - Шрифт '{filename}' не найден. Начинаю загрузку...")
        try:
            # Placeholder for progress bar logic during download
            urllib.request.urlretrieve(url, filename)
            status_window.log(f"  - Шрифт '{filename}' успешно загружен.")

        except Exception as e:
            error_msg = f"Не удалось скачать шрифт '{filename}'.\n\nОшибка: {e}\n\nПожалуйста, скачайте его вручную и поместите в папку с программой."
            status_window.log(f"ОШИБКА: {error_msg}")
            webbrowser.open_new(url)
            messagebox.showerror("Ошибка загрузки шрифта", error_msg)
            return False
            
    if all_fonts_present:
        status_window.log("Все шрифты на месте.")

    status_window.progress['value'] = 100
    return True

def run_main_app():
    """Изолирует запуск основного приложения."""
    try:
        subprocess.Popen([sys.executable, APP_FILE])
    except FileNotFoundError:
        error_msg = f"Ошибка: Основной файл программы '{APP_FILE}' не найден."
        messagebox.showerror("Файл не найден", error_msg)
    except Exception as e:
        error_msg = f"Произошла непредвиденная ошибка при запуске приложения:\n\n{e}"
        messagebox.showerror("Критическая ошибка", error_msg)

def main():
    """Основная функция-обертка для проверок и запуска."""
    status_window = LauncherStatusWindow()
    
    if sys.version_info < MIN_PYTHON_VERSION:
        messagebox.showerror("Ошибка версии Python", f"Требуется Python {'.'.join(map(str, MIN_PYTHON_VERSION))} или новее.")
        status_window.destroy()
        return

    try:
        root = tk.Tk()
        root.withdraw()
        root.destroy()
    except tk.TclError:
        messagebox.showerror("Ошибка Tkinter", "Графическая библиотека Tkinter не найдена или не работает.")
        status_window.destroy()
        return

    if not check_and_install_libraries(status_window):
        status_window.enable_close()
        status_window.mainloop()
        return

    if not check_and_download_fonts(status_window):
        status_window.enable_close()
        status_window.mainloop()
        return

    status_window.log("\nВсе готово! Запускаем калькулятор...")
    status_window.after(1000, lambda: [run_main_app(), status_window.destroy()])
    status_window.mainloop()

if __name__ == "__main__":
    main()

