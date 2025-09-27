import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoAlertPresentException
import time
import shutil
import webbrowser
import zipfile  # NEU: Für ZIP-Handling
import subprocess
import sys

def install_packages():
    required_packages = ["requests", "selenium"]
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            print(f"Package '{package}' not found. Installing...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Beim Start ausführen
install_packages()

class ArchiwumDownloader:
    def __init__(self, master):
        self.master = master
        self.master.title("Szukajwarchiwach Downloader")
        self.master.geometry("735x710")
        self.stop_flag = False
        self.browser_choice = tk.StringVar()
        self.detected_browsers = ["Firefox", "Chrome", "Edge"]
        self.log_visible = True
        self.total_urls = 0
        self.completed_urls = 0
        self.current_language = "de"

        # Übersetzungen
        self.translations = {
            "de": {
                "homepage": "Szukaj w Archiwach",
                "url": "URL:",
                "add": "Hinzufügen",
                "dir": "Zielordner:",
                "choose": "Wählen",
                "browser": "Browser:",
                "download": "Herunterladen",
                "stop": "Stopp",
                "reset": "Zurücksetzen",
                "export": "Liste exportieren",
                "import": "Liste importieren",
                "log_area": "Log-Bereich",
                "progress": "Fortschritt",
                "queue": "Warteliste:",
                "status": "Status",
                "error_no_dir": "Bitte Zielordner angeben.",
                "error_no_url": "Bitte mindestens eine URL in die Warteliste einfügen.",
                "reset_log": "Warteliste zurückgesetzt.",
                "export_log": "Warteliste exportiert nach",
                "import_log": "Warteliste importiert von",
                "stopped": "Download gestoppt.",
                "started_browser": "Starte Browser:",
                "warn": "Warnfenster erkannt:",
                "warn_ok": "Warnfenster bestätigt.",
                "rb_search": "Suche nach Radio-Button...",
                "rb_ok": "Radio-Button aktiviert (per JavaScript).",
                "rb_error": "Fehler beim Aktivieren des Radio-Buttons:",
                "click_download": "Klicke auf 'Herunterladen'...",
                "dl_started": "Download gestartet für:",
                "moved": "Dateien verschoben nach:",
                "no_files": "Keine neuen Dateien gefunden für",
                "finished": "Alle Downloads abgeschlossen, Browser geschlossen.",
                "error": "Fehler:",
                "dl_error": "Fehler beim Download"
            },
            "en": {
                "homepage": "Szukaj w Archiwach",
                "url": "URL:",
                "add": "Add",
                "dir": "Target folder:",
                "choose": "Choose",
                "browser": "Browser:",
                "download": "Download",
                "stop": "Stop",
                "reset": "Reset",
                "export": "Export list",
                "import": "Import list",
                "log_area": "Log area",
                "progress": "Progress",
                "queue": "Queue:",
                "status": "Status",
                "error_no_dir": "Please specify a target folder.",
                "error_no_url": "Please add at least one URL to the queue.",
                "reset_log": "Queue reset.",
                "export_log": "Queue exported to",
                "import_log": "Queue imported from",
                "stopped": "Download stopped.",
                "started_browser": "Starting browser:",
                "warn": "Alert detected:",
                "warn_ok": "Alert confirmed.",
                "rb_search": "Looking for radio button...",
                "rb_ok": "Radio button activated (via JavaScript).",
                "rb_error": "Error activating radio button:",
                "click_download": "Clicking 'Download'...",
                "dl_started": "Download started for:",
                "moved": "Files moved to:",
                "no_files": "No new files found for",
                "finished": "All downloads completed, browser closed.",
                "error": "Error:",
                "dl_error": "Error downloading"
            },
            "pl": {
                "homepage": "Szukaj w Archiwach",
                "url": "URL:",
                "add": "Dodaj",
                "dir": "Folder docelowy:",
                "choose": "Wybierz",
                "browser": "Przeglądarka:",
                "download": "Pobierz",
                "stop": "Stop",
                "reset": "Resetuj",
                "export": "Eksportuj listę",
                "import": "Importuj listę",
                "log_area": "Obszar logów",
                "progress": "Postęp",
                "queue": "Lista oczekujących:",
                "status": "Status",
                "error_no_dir": "Podaj folder docelowy.",
                "error_no_url": "Dodaj co najmniej jeden URL do listy.",
                "reset_log": "Lista oczekujących zresetowana.",
                "export_log": "Lista wyeksportowana do",
                "import_log": "Lista zaimportowana z",
                "stopped": "Pobieranie zatrzymane.",
                "started_browser": "Uruchamiam przeglądarkę:",
                "warn": "Wykryto alert:",
                "warn_ok": "Alert potwierdzony.",
                "rb_search": "Szukam przycisku radiowego...",
                "rb_ok": "Przycisk radiowy aktywowany (JavaScript).",
                "rb_error": "Błąd przy aktywacji przycisku:",
                "click_download": "Klikam 'Pobierz'...",
                "dl_started": "Pobieranie rozpoczęte dla:",
                "moved": "Pliki przeniesione do:",
                "no_files": "Nie znaleziono nowych plików dla",
                "finished": "Wszystkie pobrania zakończone, przeglądarka zamknięta.",
                "error": "Błąd:",
                "dl_error": "Błąd pobierania"
            },
            "ru": {
                "homepage": "Szukaj w Archiwach",
                "url": "URL:",
                "add": "Добавить",
                "dir": "Папка назначения:",
                "choose": "Выбрать",
                "browser": "Браузер:",
                "download": "Скачать",
                "stop": "Стоп",
                "reset": "Сбросить",
                "export": "Экспорт списка",
                "import": "Импорт списка",
                "log_area": "Журнал",
                "progress": "Прогресс",
                "queue": "Очередь:",
                "status": "Статус",
                "error_no_dir": "Укажите папку назначения.",
                "error_no_url": "Добавьте хотя бы один URL в очередь.",
                "reset_log": "Очередь сброшена.",
                "export_log": "Очередь экспортирована в",
                "import_log": "Очередь импортирована из",
                "stopped": "Загрузка остановлена.",
                "started_browser": "Запуск браузера:",
                "warn": "Обнаружено предупреждение:",
                "warn_ok": "Предупреждение подтверждено.",
                "rb_search": "Поиск радио-кнопки...",
                "rb_ok": "Радио-кнопка активирована (JavaScript).",
                "rb_error": "Ошибка активации радио-кнопки:",
                "click_download": "Нажимаю 'Скачать'...",
                "dl_started": "Загрузка начата для:",
                "moved": "Файлы перемещены в:",
                "no_files": "Новых файлов не найдено для",
                "finished": "Все загрузки завершены, браузер закрыт.",
                "error": "Ошибка:",
                "dl_error": "Ошибка загрузки"
            }
        }

        # GUI
        self.create_widgets()
        self.update_language()

    # -------------------- GUI Methoden --------------------
    def create_widgets(self):
        pad_x = 10
        pad_y = 5
        btn_width = 18
        btn_height = 1
        font_bold = ("TkDefaultFont", 9, "bold")

        # Top-Frame für Home-Button + Sprachbuttons in Grid
        self.top_frame = tk.Frame(self.master)
        self.top_frame.grid(row=0, column=0, columnspan=3, padx=pad_x, pady=pad_y, sticky="we")
        self.top_frame.grid_columnconfigure(0, weight=1)
        self.top_frame.grid_columnconfigure(1, weight=0)

        # Homepage Button (links)
        self.homepage_button = tk.Button(
            self.top_frame, width=btn_width, height=btn_height,
            command=self.open_homepage, bg="#ff9999", fg="black",
            activebackground="#ff6666", font=font_bold
        )
        self.homepage_button.grid(row=0, column=0, sticky="w")

        # Sprach-Buttons (rechts) - TEXT ZENTRIERT
        self.lang_frame = tk.Frame(self.top_frame)
        self.lang_frame.grid(row=0, column=1, sticky="e")

        self.lang_buttons = {}
        for code, label in [("de", "Deutsch"), ("en", "English"), ("pl", "Polski"), ("ru", "Русский")]:
            btn = tk.Button(
                self.lang_frame, text=label, width=7, height=1,
                anchor="center", justify="center",
                bg="gray", fg="white", command=lambda c=code: self.set_language(c)
            )
            btn.pack(side="left", padx=2)
            self.lang_buttons[code] = btn

        # URL Eingabe
        self.url_label = tk.Label(self.master, width=15, anchor="w", font=font_bold)
        self.url_label.grid(row=1, column=0, padx=pad_x, pady=pad_y, sticky="w")
        self.url_entry = tk.Entry(self.master, width=60)
        self.url_entry.grid(row=1, column=1, padx=pad_x, pady=pad_y, sticky="w")
        self.url_entry.bind("<Return>", lambda event: self.add_url())
        self.add_button = tk.Button(
            self.master, width=btn_width, height=btn_height,
            command=self.add_url, bg="blue", fg="white", activebackground="#003399", font=font_bold
        )
        self.add_button.grid(row=1, column=2, padx=pad_x, pady=pad_y, sticky="w")

        # Zielordner
        self.dir_label = tk.Label(self.master, width=15, anchor="w", font=font_bold)
        self.dir_label.grid(row=2, column=0, padx=pad_x, pady=pad_y, sticky="w")
        self.dir_entry = tk.Entry(self.master, width=60)
        self.dir_entry.grid(row=2, column=1, padx=pad_x, pady=pad_y, sticky="w")
        self.dir_button = tk.Button(
            self.master, width=btn_width, height=btn_height, command=self.choose_directory,
            bg="blue", fg="white", activebackground="#003399", font=font_bold
        )
        self.dir_button.grid(row=2, column=2, padx=pad_x, pady=pad_y, sticky="w")

        # Browser-Auswahl
        self.browser_label = tk.Label(self.master, width=15, anchor="w", font=font_bold)
        self.browser_label.grid(row=3, column=0, padx=pad_x, pady=pad_y, sticky="w")
        self.browser_dropdown = ttk.Combobox(
            self.master, textvariable=self.browser_choice,
            values=self.detected_browsers, state="readonly", width=15
        )
        if self.detected_browsers:
            self.browser_choice.set(self.detected_browsers[0])
        self.browser_dropdown.grid(row=3, column=1, padx=pad_x, pady=pad_y, sticky="w")

        # Erste Button-Reihe
        self.start_button = tk.Button(
            self.master, width=btn_width, height=btn_height, command=self.start_download,
            bg="blue", fg="white", activebackground="#003399", font=font_bold
        )
        self.start_button.grid(row=4, column=0, padx=pad_x, pady=pad_y)
        self.stop_button = tk.Button(
            self.master, width=btn_width, height=btn_height, command=self.stop_download,
            bg="blue", fg="white", activebackground="#003399", font=font_bold
        )
        self.stop_button.grid(row=4, column=1, padx=pad_x, pady=pad_y)
        self.reset_button = tk.Button(
            self.master, width=btn_width, height=btn_height, command=self.reset_queue,
            bg="blue", fg="white", activebackground="#003399", font=font_bold
        )
        self.reset_button.grid(row=4, column=2, padx=pad_x, pady=pad_y)

        # Zweite Button-Reihe
        self.export_button = tk.Button(
            self.master, width=btn_width, height=btn_height, command=self.export_list,
            bg="blue", fg="white", activebackground="#003399", font=font_bold
        )
        self.export_button.grid(row=5, column=0, padx=pad_x, pady=pad_y)
        self.import_button = tk.Button(
            self.master, width=btn_width, height=btn_height, command=self.import_list,
            bg="blue", fg="white", activebackground="#003399", font=font_bold
        )
        self.import_button.grid(row=5, column=1, padx=pad_x, pady=pad_y)
        self.toggle_log_button = tk.Button(
            self.master, width=btn_width, height=btn_height, command=self.toggle_log,
            bg="blue", fg="white", activebackground="#003399", font=font_bold
        )
        self.toggle_log_button.grid(row=5, column=2, padx=pad_x, pady=pad_y)

        # Progressbar + Label
        self.progress = ttk.Progressbar(self.master, orient="horizontal", length=800, mode="determinate")
        self.progress.grid(row=6, column=0, columnspan=3, padx=pad_x, pady=(pad_y, 0), sticky="we")
        self.progress_label = tk.Label(self.master, anchor="center", font=font_bold)
        self.progress_label.grid(row=7, column=0, columnspan=3, padx=pad_x, pady=(0, pad_y), sticky="we")

        # Warteliste
        self.queue_label = tk.Label(self.master, anchor="w", font=font_bold)
        self.queue_label.grid(row=8, column=0, columnspan=3, padx=pad_x, pady=pad_y, sticky="w")
        self.queue_tree = ttk.Treeview(self.master, columns=("status", "url"), show="headings", height=10)
        self.queue_tree.column("status", width=80, anchor="center")
        self.queue_tree.column("url", width=700, anchor="w")
        self.queue_tree.grid(row=9, column=0, columnspan=3, padx=pad_x, pady=pad_y, sticky="nsew")
        self.queue_tree.tag_configure("oddrow", background="#f0f0f0", foreground="black")
        self.queue_tree.tag_configure("evenrow", background="#e0e0e0", foreground="black")
        self.queue_tree.bind("<Delete>", self.delete_selected)

        # Log-Bereich
        self.log_text = tk.Text(self.master, height=13, width=105, state="disabled")
        self.log_text.grid(row=10, column=0, columnspan=3, padx=pad_x, pady=pad_y, sticky="nsew")

        # Grid dynamisch
        self.master.grid_rowconfigure(9, weight=1)
        self.master.grid_rowconfigure(10, weight=1)
        self.master.grid_columnconfigure(1, weight=1)

    # -------------------- Sprach-Methoden --------------------
    def set_language(self, lang_code):
        self.current_language = lang_code
        self.update_language()

    def update_language(self):
        t = self.translations[self.current_language]
        self.homepage_button.config(text=t["homepage"])
        self.url_label.config(text=t["url"])
        self.add_button.config(text=t["add"])
        self.dir_label.config(text=t["dir"])
        self.dir_button.config(text=t["choose"])
        self.browser_label.config(text=t["browser"])
        self.start_button.config(text=t["download"])
        self.stop_button.config(text=t["stop"])
        self.reset_button.config(text=t["reset"])
        self.export_button.config(text=t["export"])
        self.import_button.config(text=t["import"])
        self.toggle_log_button.config(text=t["log_area"])
        self.progress_label.config(text=f"{t['progress']}: 0%")
        self.queue_label.config(text=t["queue"])
        self.queue_tree.heading("status", text=t["status"])
        self.queue_tree.heading("url", text="URL")

    # -------------------- GUI Hilfsfunktionen --------------------
    def open_homepage(self):
        webbrowser.open("https://www.szukajwarchiwach.gov.pl/de/strona_glowna")

    def log(self, msg):
        try:
            self.log_text.config(state="normal")
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state="disabled")
        except Exception:
            pass

    def choose_directory(self):
        folder = filedialog.askdirectory()
        if folder:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, folder)

    def add_url(self):
        url = self.url_entry.get().strip()
        if url:
            row_count = len(self.queue_tree.get_children())
            tag = "evenrow" if row_count % 2 == 0 else "oddrow"
            self.queue_tree.insert("", "end", values=("○", url), tags=(tag,))
            self.url_entry.delete(0, tk.END)

    def delete_selected(self, event=None):
        for item in self.queue_tree.selection():
            self.queue_tree.delete(item)
        self.recolor_rows()

    def reset_queue(self):
        t = self.translations[self.current_language]
        for item in self.queue_tree.get_children():
            self.queue_tree.delete(item)
        self.log(t["reset_log"])

    def recolor_rows(self):
        for index, item in enumerate(self.queue_tree.get_children()):
            tag = "evenrow" if index % 2 == 0 else "oddrow"
            self.queue_tree.item(item, tags=(tag,))

    def export_list(self):
        t = self.translations[self.current_language]
        file = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Textdateien", "*.txt")])
        if file:
            with open(file, "w", encoding="utf-8") as f:
                for item in self.queue_tree.get_children():
                    url = self.queue_tree.item(item, "values")[1]
                    f.write(url + "\n")
            self.log(f"{t['export_log']} {file}")

    def import_list(self):
        t = self.translations[self.current_language]
        file = filedialog.askopenfilename(filetypes=[("Textdateien", "*.txt")])
        if file:
            with open(file, "r", encoding="utf-8") as f:
                for line in f:
                    url = line.strip()
                    if url:
                        row_count = len(self.queue_tree.get_children())
                        tag = "evenrow" if row_count % 2 == 0 else "oddrow"
                        self.queue_tree.insert("", "end", values=("○", url), tags=(tag,))
            self.log(f"{t['import_log']} {file}")

    def toggle_log(self):
        if self.log_visible:
            self.log_text.grid_remove()
            self.master.geometry("735x510")
            self.log_visible = False
        else:
            self.log_text.grid()
            self.master.geometry("735x700")
            self.log_visible = True

    # -------------------- Download --------------------
    def start_download(self):
        t = self.translations[self.current_language]
        out_dir = self.dir_entry.get().strip()
        if not out_dir:
            messagebox.showerror(t["error"], t["error_no_dir"])
            return
        if len(self.queue_tree.get_children()) == 0:
            messagebox.showerror(t["error"], t["error_no_url"])
            return
        self.stop_flag = False
        urls = [(item, self.queue_tree.item(item, "values")[1]) for item in self.queue_tree.get_children()]
        self.total_urls = len(urls)
        self.completed_urls = 0
        threading.Thread(target=self.download_process, args=(urls, out_dir), daemon=True).start()

    def stop_download(self):
        t = self.translations[self.current_language]
        self.stop_flag = True
        self.log(t["stopped"])

    def update_progress(self):
        t = self.translations[self.current_language]
        if self.total_urls > 0:
            percent = int((self.completed_urls / self.total_urls) * 100)
            self.progress["value"] = percent
            self.progress_label.config(text=f"{t['progress']}: {percent}%")
        else:
            self.progress["value"] = 0
            self.progress_label.config(text=f"{t['progress']}: 0%")

    def get_subfolder_from_url(self, url, out_dir):
        match = re.findall(r"(\d+)$", url)
        folder_name = match[0] if match else "unbekannt"
        folder_path = os.path.join(out_dir, folder_name)
        os.makedirs(folder_path, exist_ok=True)
        return folder_path

    def handle_alert(self, driver):
        t = self.translations[self.current_language]
        try:
            alert = driver.switch_to.alert
            self.log(f"{t['warn']} {alert.text}")
            alert.accept()
            self.log(t["warn_ok"])
            return True
        except NoAlertPresentException:
            return False

    # -------------------- Kern: Download-Prozess --------------------
    def download_process(self, urls, out_dir):
        t = self.translations[self.current_language]
        try:
            browser = self.browser_choice.get()
            self.log(f"{t['started_browser']} {browser}")

            driver = None
            if browser == "Firefox":
                try:
                    options = FirefoxOptions()
                    driver = webdriver.Firefox(options=options)
                except Exception:
                    self.log("Firefox-Browser nicht gefunden.")
                    return
            elif browser == "Chrome":
                try:
                    options = ChromeOptions()
                    driver = webdriver.Chrome(options=options)
                except Exception:
                    self.log("Chrome-Browser nicht gefunden.")
                    return
            elif browser == "Edge":
                edge_path = self.find_edge_executable()
                if edge_path:
                    try:
                        options = EdgeOptions()
                        options.binary_location = edge_path
                        driver = webdriver.Edge(options=options)
                    except Exception:
                        self.log("Edge-Browser konnte nicht gestartet werden.")
                        return
                else:
                    self.log("Edge-Browser nicht gefunden.")
                    return

            if not driver:
                self.log("Kein gültiger Browser gefunden. Download abgebrochen.")
                return

            # Homepage öffnen + minimieren
            homepage_url = "https://www.szukajwarchiwach.gov.pl/de/strona_glowna"
            try:
                driver.minimize_window()
                try:
                     driver.get(homepage_url)
                except Exception:
                    pass
            except Exception as e:
                self.log(f"{t['error']} {e}")

            # Downloadverzeichnis
            download_dir = os.path.expanduser("~/Downloads")
            if not os.path.isdir(download_dir):
                download_dir = out_dir
                os.makedirs(download_dir, exist_ok=True)

            # initiale Dateien
            try:
                initial_global_files = set(os.listdir(download_dir))
            except Exception:
                initial_global_files = set()

            # Info pro URL vorbereiten
            items_info = []
            for item, url in urls:
                if self.stop_flag:
                    break
                try:
                    self.queue_tree.set(item, "status", "○")
                except Exception:
                    pass
                subfolder = self.get_subfolder_from_url(url, out_dir)
                items_info.append({
                    "item": item,
                    "url": url,
                    "subfolder": subfolder,
                    "click_time": None,
                    "moved_files": []
                })

            # Downloads auslösen
            for info in items_info:
                if self.stop_flag:
                    break
                url = info["url"]
                item = info["item"]
                try:
                    self.log(f"{t['dl_started']} {url}")
                    driver.get(url)
                    time.sleep(1.5)

                    self.handle_alert(driver)

                    try:
                        self.log(t["rb_search"])
                        radio_button = WebDriverWait(driver, 6).until(
                            EC.presence_of_element_located((By.ID, "_Jednostka_wyborSkanow_1"))
                        )
                        driver.execute_script("arguments[0].click();", radio_button)
                        self.log(t["rb_ok"])
                        time.sleep(1.5)
                    except Exception as e:
                        self.log(f"{t['rb_error']} {e}")

                    try:
                        self.log(t["click_download"])
                        xpath = ("//button[contains(normalize-space(.), 'Herunterladen') or "
                                 "contains(normalize-space(.), 'Pobierz') or "
                                 "contains(normalize-space(.), 'Download')]")
                        download_button = WebDriverWait(driver, 8).until(
                            EC.element_to_be_clickable((By.XPATH, xpath))
                        )
                        click_time = time.time()
                        info["click_time"] = click_time
                        driver.execute_script("arguments[0].click();", download_button)
                        self.log(f"{t['dl_started']} {url}")
                        time.sleep(1.5)
                    except Exception as e:
                        self.log(f"{t['error']} {e}")
                        info["click_time"] = time.time()

                    time.sleep(1.5)

                except Exception as e:
                    self.log(f"{t['error']} {e}")
                    try:
                        self.queue_tree.set(item, "status", "✖")
                    except Exception:
                        pass

                try:
                    self.queue_tree.set(item, "status", "…")
                except Exception:
                    pass

            # Warten auf alle Downloads
            self.log("Warte auf Abschluss aller Downloads...")
            final_files_set = self.wait_for_downloads_complete(download_dir, initial_global_files, timeout=900)
            new_files = [f for f in final_files_set if f not in initial_global_files]

            click_info_list = [(info["click_time"] if info["click_time"] else 0.0, info) for info in items_info]
            click_info_list.sort(key=lambda x: x[0])

            # Dateien verschieben + ZIP entpacken
            moved_summary = {}
            for filename in new_files:
                src = os.path.join(download_dir, filename)
                try:
                    file_mtime = os.path.getmtime(src)
                except Exception:
                    file_mtime = time.time()
                chosen_info = None
                chosen_time = -1
                for ct, info in click_info_list:
                    if ct <= file_mtime and ct >= chosen_time:
                        chosen_time = ct
                        chosen_info = info
                if chosen_info:
                    dest_folder = chosen_info["subfolder"]
                else:
                    dest_folder = os.path.join(out_dir, "unbekannt")
                    os.makedirs(dest_folder, exist_ok=True)
                dest = os.path.join(dest_folder, filename)
                try:
                    shutil.move(src, dest)
                    self.log(f"{t['moved']} {dest}")
                    chosen_info["moved_files"].append(dest)

                    # NEU: ZIP entpacken
                    if zipfile.is_zipfile(dest):
                        with zipfile.ZipFile(dest, 'r') as zip_ref:
                            zip_ref.extractall(dest_folder)
                        self.log(f"ZIP entpackt in {dest_folder}")
                except Exception as e:
                    self.log(f"{t['dl_error']} {filename}: {e}")

            # Status aktualisieren
            for info in items_info:
                try:
                    self.queue_tree.set(info["item"], "status", "✔")
                except Exception:
                    pass
                self.completed_urls += 1
                self.update_progress()

            self.log(t["finished"])

        finally:
            try:
                driver.quit()
            except Exception:
                pass

    def wait_for_downloads_complete(self, download_dir, initial_files, timeout=900):
        start_time = time.time()
        current_files = set(os.listdir(download_dir))
        while time.time() - start_time < timeout:
            new_files = set(os.listdir(download_dir)) - initial_files
            if all(not f.endswith(".crdownload") and not f.endswith(".part") for f in new_files):
                return set(os.listdir(download_dir))
            time.sleep(0.5)
        return set(os.listdir(download_dir))

    def find_edge_executable(self):
        import platform
        if platform.system() == "Windows":
            possible_paths = [
                r"C:/Program Files (x86)/Microsoft/EdgeCore/140.0.3485.81/msedge.exe",
                r"C:/Program Files/Microsoft/EdgeCore/140.0.3485.81/msedge.exe"
            ]
            for path in possible_paths:
                if os.path.isfile(path):
                    return path
        return None


if __name__ == "__main__":
    root = tk.Tk()
    app = ArchiwumDownloader(root)
    root.mainloop()
