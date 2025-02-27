import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import os
import datetime
import calendar
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkcalendar import Calendar, DateEntry
import locale
import bcrypt
from PIL import Image, ImageTk
import csv
import json
import webbrowser
import threading
import uuid
import logging
import shutil
from fpdf import FPDF
import re

# Setze deutsche Sprache
try:
    locale.setlocale(locale.LC_ALL, 'de_DE')
except:
    try:
        locale.setlocale(locale.LC_ALL, 'German_Germany')
    except:
        pass

# Konstanten
APP_NAME = "MitarbeiterPro"
VERSION = "1.0.0"
APPDATA_DIR = os.path.join(os.getenv('APPDATA'), APP_NAME)
DATABASE_PATH = os.path.join(APPDATA_DIR, 'employees.db')
LOG_PATH = os.path.join(APPDATA_DIR, 'logs')
EXPORT_PATH = os.path.join(APPDATA_DIR, 'exports')
BACKUP_PATH = os.path.join(APPDATA_DIR, 'backups')
CONFIG_PATH = os.path.join(APPDATA_DIR, 'config.json')
THEME_COLOR = "#3498db"
LIGHT_COLOR = "#ecf0f1"
DARK_COLOR = "#2c3e50"
WARNING_COLOR = "#e74c3c"
SUCCESS_COLOR = "#2ecc71"
NEUTRAL_COLOR = "#f39c12"

# Logger einrichten
def setup_logging():
    if not os.path.exists(LOG_PATH):
        os.makedirs(LOG_PATH)
    
    log_file = os.path.join(LOG_PATH, f'{datetime.datetime.now().strftime("%Y-%m-%d")}.log')
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(APP_NAME)

# Verzeichnisse erstellen
def setup_directories():
    for directory in [APPDATA_DIR, LOG_PATH, EXPORT_PATH, BACKUP_PATH]:
        if not os.path.exists(directory):
            os.makedirs(directory)

# Datenbank erstellen und initialisieren
def setup_database():
    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    
    # Mitarbeitertabelle
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS employees (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id TEXT UNIQUE,
        first_name TEXT NOT NULL,
        last_name TEXT NOT NULL,
        birth_date TEXT,
        address TEXT,
        phone TEXT,
        email TEXT,
        position TEXT,
        department TEXT,
        hire_date TEXT,
        salary REAL,
        status TEXT DEFAULT 'Aktiv',
        vacation_days_per_year INTEGER DEFAULT 30,
        sick_days_used INTEGER DEFAULT 0,
        profile_image TEXT,
        notes TEXT,
        created_at TEXT,
        updated_at TEXT
    )
    ''')
    
    # Urlaubstabelle
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS vacation (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER,
        start_date TEXT,
        end_date TEXT,
        days INTEGER,
        status TEXT DEFAULT 'Beantragt',
        approved_by TEXT,
        approved_date TEXT,
        notes TEXT,
        created_at TEXT,
        FOREIGN KEY (employee_id) REFERENCES employees (id)
    )
    ''')
    
    # Krankschreibungstabelle
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS sick_leave (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER,
        start_date TEXT,
        end_date TEXT,
        days INTEGER,
        medical_certificate BOOLEAN,
        notes TEXT,
        created_at TEXT,
        FOREIGN KEY (employee_id) REFERENCES employees (id)
    )
    ''')
    
    # Gehaltstabelle
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS salary_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER,
        amount REAL,
        effective_date TEXT,
        created_at TEXT,
        FOREIGN KEY (employee_id) REFERENCES employees (id)
    )
    ''')
    
    # Ausgabentabelle
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER,
        amount REAL,
        category TEXT,
        date TEXT,
        receipt_path TEXT,
        status TEXT DEFAULT 'Eingereicht',
        approved_by TEXT,
        approved_date TEXT,
        notes TEXT,
        created_at TEXT,
        FOREIGN KEY (employee_id) REFERENCES employees (id)
    )
    ''')
    
    # Arbeitszeitentabelle
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS working_time (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER,
        date TEXT,
        start_time TEXT,
        end_time TEXT,
        break_duration INTEGER,
        total_hours REAL,
        notes TEXT,
        created_at TEXT,
        FOREIGN KEY (employee_id) REFERENCES employees (id)
    )
    ''')
    
    # Benutzertabelle
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        password_hash TEXT,
        full_name TEXT,
        role TEXT,
        last_login TEXT,
        created_at TEXT,
        updated_at TEXT
    )
    ''')

    # Abteilungstabelle
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS departments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        description TEXT,
        manager_id INTEGER,
        created_at TEXT,
        updated_at TEXT,
        FOREIGN KEY (manager_id) REFERENCES employees (id)
    )
    ''')

    # Dokumententabelle
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS documents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER,
        document_type TEXT,
        file_path TEXT,
        upload_date TEXT,
        notes TEXT,
        created_at TEXT,
        FOREIGN KEY (employee_id) REFERENCES employees (id)
    )
    ''')
    
    # Standardabteilungen einfügen
    departments = [
        ('IT', 'Informationstechnologie', None),
        ('HR', 'Personalabteilung', None),
        ('Finanzen', 'Finanzabteilung', None),
        ('Vertrieb', 'Vertriebsabteilung', None),
        ('Marketing', 'Marketingabteilung', None)
    ]
    
    for dept in departments:
        try:
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            cursor.execute('''
            INSERT INTO departments (name, description, manager_id, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ''', (dept[0], dept[1], dept[2], current_time, current_time))
        except sqlite3.IntegrityError:
            # Abteilung existiert bereits
            pass
    
    # Standardadministrator erstellen
    try:
        admin_password = "admin123"
        hashed_password = bcrypt.hashpw(admin_password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
        current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        cursor.execute('''
        INSERT INTO users (username, password_hash, full_name, role, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?)
        ''', ('admin', hashed_password, 'Administrator', 'admin', current_time, current_time))
    except sqlite3.IntegrityError:
        # Admin existiert bereits
        pass
    
    conn.commit()
    conn.close()

# Konfiguration laden oder erstellen
def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        default_config = {
            "company_name": "Ihr Unternehmen",
            "company_address": "Musterstraße 1, 12345 Musterstadt",
            "company_phone": "+49 123 456789",
            "company_email": "info@ihrunternehmen.de",
            "company_logo": "",
            "vacation_days_default": 30,
            "working_hours_per_day": 8,
            "theme": "light",
            "language": "de",
            "backup_frequency": "daily",
            "last_backup": None
        }
        save_config(default_config)
        return default_config

# Konfiguration speichern
def save_config(config):
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4)

# Backup der Datenbank erstellen
def create_backup():
    if not os.path.exists(DATABASE_PATH):
        return False
    
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(BACKUP_PATH, f'employees_backup_{timestamp}.db')
    
    try:
        shutil.copy2(DATABASE_PATH, backup_file)
        
        # Aktualisiere letzte Backup-Zeit in Konfiguration
        config = load_config()
        config['last_backup'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        save_config(config)
        
        return True
    except Exception as e:
        logger.error(f"Backup fehlgeschlagen: {e}")
        return False

# Hilfsfunktion: Datum formatieren
def format_date(date_str, format_from="%Y-%m-%d", format_to="%d.%m.%Y"):
    if not date_str:
        return ""
    try:
        date_obj = datetime.datetime.strptime(date_str, format_from)
        return date_obj.strftime(format_to)
    except:
        return date_str

# Hilfsfunktion: Tage zwischen zwei Daten berechnen
def calculate_days(start_date, end_date, include_weekends=True):
    try:
        start = datetime.datetime.strptime(start_date, "%Y-%m-%d").date()
        end = datetime.datetime.strptime(end_date, "%Y-%m-%d").date()
        
        if include_weekends:
            return (end - start).days + 1
        else:
            days = 0
            current = start
            while current <= end:
                if current.weekday() < 5:  # 0-4 sind Montag bis Freitag
                    days += 1
                current += datetime.timedelta(days=1)
            return days
    except:
        return 0

# Login-Fenster
class LoginWindow:
    def __init__(self, root, on_successful_login):
        self.root = root
        self.on_successful_login = on_successful_login
        self.root.title(f"{APP_NAME} - Login")
        self.root.geometry("400x500")
        self.root.resizable(False, False)
        
        self.setup_ui()
        
    def setup_ui(self):
        # Container
        self.container = tk.Frame(self.root, bg=LIGHT_COLOR)
        self.container.place(relx=0.5, rely=0.5, anchor=tk.CENTER, relwidth=0.9, relheight=0.9)
        
        # Logo/Titel
        title_frame = tk.Frame(self.container, bg=LIGHT_COLOR)
        title_frame.pack(pady=(20, 30))
        
        title_label = tk.Label(title_frame, text=APP_NAME, font=("Arial", 24, "bold"), fg=THEME_COLOR, bg=LIGHT_COLOR)
        title_label.pack()
        
        version_label = tk.Label(title_frame, text=f"Version {VERSION}", font=("Arial", 10), fg=DARK_COLOR, bg=LIGHT_COLOR)
        version_label.pack()
        
        # Login Formular
        form_frame = tk.Frame(self.container, bg=LIGHT_COLOR)
        form_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=10)
        
        username_label = tk.Label(form_frame, text="Benutzername", font=("Arial", 12), anchor="w", bg=LIGHT_COLOR)
        username_label.pack(fill=tk.X, pady=(10, 5))
        
        self.username_entry = tk.Entry(form_frame, font=("Arial", 12), relief=tk.SOLID, bd=1)
        self.username_entry.pack(fill=tk.X, pady=(0, 15))
        
        password_label = tk.Label(form_frame, text="Passwort", font=("Arial", 12), anchor="w", bg=LIGHT_COLOR)
        password_label.pack(fill=tk.X, pady=(10, 5))
        
        self.password_entry = tk.Entry(form_frame, font=("Arial", 12), relief=tk.SOLID, bd=1, show="•")
        self.password_entry.pack(fill=tk.X, pady=(0, 15))
        
        # Login Button
        login_button = tk.Button(
            form_frame, 
            text="Anmelden", 
            font=("Arial", 12, "bold"), 
            bg=THEME_COLOR, 
            fg="white", 
            relief=tk.FLAT,
            padx=20,
            pady=10,
            command=self.login
        )
        login_button.pack(fill=tk.X, pady=(20, 0))
        
        # Status Label
        self.status_label = tk.Label(form_frame, text="", font=("Arial", 10), fg=WARNING_COLOR, bg=LIGHT_COLOR)
        self.status_label.pack(fill=tk.X, pady=(10, 0))
        
        # Copyright
        copyright_frame = tk.Frame(self.container, bg=LIGHT_COLOR)
        copyright_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
        copyright_label = tk.Label(
            copyright_frame, 
            text=f"© {datetime.datetime.now().year} {APP_NAME}", 
            font=("Arial", 8), 
            fg=DARK_COLOR, 
            bg=LIGHT_COLOR
        )
        copyright_label.pack()
        
    def login(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get()
        
        if not username or not password:
            self.status_label.config(text="Bitte Benutzername und Passwort eingeben")
            return
        
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        
        cursor.execute("SELECT id, password_hash, role FROM users WHERE username = ?", (username,))
        user = cursor.fetchone()
        
        if user and bcrypt.checkpw(password.encode('utf-8'), user[1].encode('utf-8')):
            # Update last login
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            cursor.execute("UPDATE users SET last_login = ?, updated_at = ? WHERE id = ?", 
                           (current_time, current_time, user[0]))
            conn.commit()
            
            # Log successful login
            logger.info(f"Benutzer {username} hat sich erfolgreich angemeldet.")
            
            # Call success callback with user info
            self.on_successful_login({"id": user[0], "username": username, "role": user[2]})
            
            self.root.destroy()
        else:
            self.status_label.config(text="Ungültiger Benutzername oder Passwort")
            logger.warning(f"Fehlgeschlagener Anmeldeversuch für Benutzer {username}")
        
        conn.close()

# Hauptanwendung
class EmployeeManagementSystem:
    def __init__(self, root, user):
        self.root = root
        self.user = user
        self.active_frame = None
        self.config = load_config()
        
        self.setup_ui()
        self.show_dashboard()
        
        # Überprüfe, ob ein Backup erstellt werden sollte
        self.check_backup_needs()
        
    def check_backup_needs(self):
        if not self.config.get('last_backup'):
            # Erstes Backup erstellen
            threading.Thread(target=create_backup).start()
            return
            
        last_backup = datetime.datetime.strptime(self.config['last_backup'], "%Y-%m-%d %H:%M:%S")
        now = datetime.datetime.now()
        
        if self.config['backup_frequency'] == 'daily' and (now - last_backup).days >= 1:
            threading.Thread(target=create_backup).start()
        elif self.config['backup_frequency'] == 'weekly' and (now - last_backup).days >= 7:
            threading.Thread(target=create_backup).start()
        elif self.config['backup_frequency'] == 'monthly' and (now - last_backup).days >= 30:
            threading.Thread(target=create_backup).start()
    
    def setup_ui(self):
        # Fenster konfigurieren
        self.root.title(f"{APP_NAME} - Mitarbeiterverwaltungssystem")
        self.root.state('zoomed')  # Maximiertes Fenster
        self.root.minsize(1200, 700)
        
        # Hauptcontainer
        self.main_container = tk.Frame(self.root)
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # Menü und Hauptbereich
        self.setup_menu()
        self.setup_content_area()
        
        # Statusleiste
        self.setup_status_bar()
    
    def setup_menu(self):
        # Seitenleiste
        self.sidebar = tk.Frame(self.main_container, bg=DARK_COLOR, width=220)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)  # Größe fixieren
        
        # Logo/Titel
        title_frame = tk.Frame(self.sidebar, bg=DARK_COLOR)
        title_frame.pack(fill=tk.X, padx=10, pady=(20, 30))
        
        title_label = tk.Label(
            title_frame, 
            text=APP_NAME, 
            font=("Arial", 16, "bold"), 
            fg="white", 
            bg=DARK_COLOR
        )
        title_label.pack(anchor=tk.W)
        
        # Menüpunkte
        menu_items = [
            ("Dashboard", self.show_dashboard, "🏠"),
            ("Mitarbeiter", self.show_employees, "👥"),
            ("Urlaub", self.show_vacation, "🏖️"),
            ("Krankschreibungen", self.show_sick_leave, "🏥"),
            ("Gehalt", self.show_salary, "💰"),
            ("Ausgaben", self.show_expenses, "💸"),
            ("Arbeitszeit", self.show_working_time, "⏱️"),
            ("Berichte", self.show_reports, "📊"),
            ("Einstellungen", self.show_settings, "⚙️")
        ]
        
        self.menu_buttons = {}
        
        for text, command, icon in menu_items:
            btn = tk.Button(
                self.sidebar,
                text=f"{icon} {text}",
                font=("Arial", 11),
                bg=DARK_COLOR,
                fg="white",
                bd=0,
                padx=10,
                pady=12,
                anchor=tk.W,
                justify=tk.LEFT,
                highlightthickness=0,
                activebackground=THEME_COLOR,
                activeforeground="white",
                command=command
            )
            btn.pack(fill=tk.X, padx=5, pady=2)
            self.menu_buttons[text] = btn
        
        # Benutzerinformationen
        user_frame = tk.Frame(self.sidebar, bg=DARK_COLOR)
        user_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=20)
        
        user_icon_label = tk.Label(
            user_frame, 
            text="👤", 
            font=("Arial", 16), 
            fg="white", 
            bg=DARK_COLOR
        )
        user_icon_label.pack(side=tk.LEFT, padx=(0, 5))
        
        user_label = tk.Label(
            user_frame, 
            text=self.user["username"], 
            font=("Arial", 10), 
            fg="white", 
            bg=DARK_COLOR, 
            anchor=tk.W
        )
        user_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        logout_button = tk.Button(
            user_frame,
            text="🚪",
            font=("Arial", 12),
            bg=DARK_COLOR,
            fg="white",
            bd=0,
            padx=5,
            activebackground=WARNING_COLOR,
            activeforeground="white",
            command=self.logout
        )
        logout_button.pack(side=tk.RIGHT)
    
    def setup_content_area(self):
        # Header und Content-Bereich
        self.content_container = tk.Frame(self.main_container, bg=LIGHT_COLOR)
        self.content_container.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Header
        self.header = tk.Frame(self.content_container, bg=LIGHT_COLOR, height=60)
        self.header.pack(fill=tk.X, padx=20, pady=10)
        
        self.header_title = tk.Label(
            self.header, 
            text="Dashboard", 
            font=("Arial", 18, "bold"), 
            fg=DARK_COLOR, 
            bg=LIGHT_COLOR
        )
        self.header_title.pack(side=tk.LEFT)
        
        # Aktuelle Zeit anzeigen
        self.time_label = tk.Label(
            self.header, 
            font=("Arial", 12), 
            fg=DARK_COLOR, 
            bg=LIGHT_COLOR
        )
        self.time_label.pack(side=tk.RIGHT)
        self.update_time()
        
        # Trennlinie
        separator = ttk.Separator(self.content_container, orient="horizontal")
        separator.pack(fill=tk.X, padx=20)
        
        # Hauptinhalt
        self.content_frame = tk.Frame(self.content_container, bg=LIGHT_COLOR)
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    
    def setup_status_bar(self):
        # Statusleiste
        self.status_bar = tk.Frame(self.root, bg=DARK_COLOR, height=25)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Version
        version_label = tk.Label(
            self.status_bar, 
            text=f"Version {VERSION}", 
            font=("Arial", 8), 
            fg="white", 
            bg=DARK_COLOR
        )
        version_label.pack(side=tk.RIGHT, padx=10)
        
        # Status
        self.status_label = tk.Label(
            self.status_bar, 
            text="Bereit", 
            font=("Arial", 8), 
            fg="white", 
            bg=DARK_COLOR
        )
        self.status_label.pack(side=tk.LEFT, padx=10)
    
    def update_time(self):
        current_time = datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
        self.time_label.config(text=current_time)
        self.root.after(1000, self.update_time)
    
    def update_status(self, message):
        self.status_label.config(text=message)
    
    def clear_content(self):
        # Bisherigen Inhalt entfernen
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        
        # Aktiven Button zurücksetzen
        for button in self.menu_buttons.values():
            button.config(bg=DARK_COLOR)
    
    def highlight_menu_button(self, button_name):
        if button_name in self.menu_buttons:
            self.menu_buttons[button_name].config(bg=THEME_COLOR)
    
    def logout(self):
        if messagebox.askyesno("Abmelden", "Möchten Sie sich wirklich abmelden?"):
            logger.info(f"Benutzer {self.user['username']} hat sich abgemeldet.")
            self.root.destroy()
            
            # Neue Anwendung starten
            new_root = tk.Tk()
            LoginWindow(new_root, self.on_successful_login)
            new_root.mainloop()
    
    def on_successful_login(self, user):
        new_root = tk.Tk()
        app = EmployeeManagementSystem(new_root, user)
        new_root.mainloop()
    
    # --- Hauptfunktionen für die verschiedenen Bereiche ---
    
    def show_dashboard(self):
        self.clear_content()
        self.header_title.config(text="Dashboard")
        self.highlight_menu_button("Dashboard")
        
        # Container für Statistiken
        stats_frame = tk.Frame(self.content_frame, bg=LIGHT_COLOR)
        stats_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Karten für verschiedene Statistiken
        self.create_stat_card(stats_frame, "Mitarbeiter", self.get_employee_count(), "👥", "#3498db")
        self.create_stat_card(stats_frame, "Aktuell im Urlaub", self.get_current_vacation_count(), "🏖️", "#2ecc71")
        self.create_stat_card(stats_frame, "Krank gemeldet", self.get_current_sick_count(), "🏥", "#e74c3c")
        self.create_stat_card(stats_frame, "Geburtstage diesen Monat", self.get_birthdays_this_month(), "🎂", "#f39c12")
        
        # Container für Diagramme
        charts_frame = tk.Frame(self.content_frame, bg=LIGHT_COLOR)
        charts_frame.pack(fill=tk.BOTH, expand=True)
        
        # Linkes Diagramm (Mitarbeiter nach Abteilung)
        left_chart_frame = tk.Frame(charts_frame, bg="white", bd=1, relief=tk.SOLID)
        left_chart_frame.grid(row=0, column=0, padx=(0, 10), pady=10, sticky="nsew")
        
        chart_title = tk.Label(left_chart_frame, text="Mitarbeiter nach Abteilung", font=("Arial", 12, "bold"), bg="white")
        chart_title.pack(pady=(10, 0))
        
        # Diagramm erstellen
        figure1 = plt.Figure(figsize=(5, 4), dpi=100)
        ax1 = figure1.add_subplot(111)
        departments, counts = self.get_employees_by_department()
        ax1.bar(departments, counts, color=THEME_COLOR)
        ax1.set_ylabel('Anzahl')
        ax1.set_title('')
        
        for i, v in enumerate(counts):
            ax1.text(i, v + 0.1, str(v), ha='center')
        
        canvas = FigureCanvasTkAgg(figure1, left_chart_frame)
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Rechtes Diagramm (Urlaubsstatistik)
        right_chart_frame = tk.Frame(charts_frame, bg="white", bd=1, relief=tk.SOLID)
        right_chart_frame.grid(row=0, column=1, padx=(10, 0), pady=10, sticky="nsew")
        
        chart_title = tk.Label(right_chart_frame, text="Urlaub & Krankheitstage", font=("Arial", 12, "bold"), bg="white")
        chart_title.pack(pady=(10, 0))
        
        # Diagramm erstellen
        figure2 = plt.Figure(figsize=(5, 4), dpi=100)
        ax2 = figure2.add_subplot(111)
        months = [calendar.month_name[i] for i in range(1, 13)]
        vacation_data = self.get_vacation_by_month()
        sick_data = self.get_sick_leave_by_month()
        
        ax2.plot(months, vacation_data, label='Urlaub', marker='o', color='#3498db')
        ax2.plot(months, sick_data, label='Krankheit', marker='s', color='#e74c3c')
        ax2.set_ylabel('Tage')
        ax2.set_title('')
        ax2.legend()
        
        canvas = FigureCanvasTkAgg(figure2, right_chart_frame)
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Grid-Konfiguration für gleichmäßige Größe
        charts_frame.grid_columnconfigure(0, weight=1)
        charts_frame.grid_columnconfigure(1, weight=1)
        charts_frame.grid_rowconfigure(0, weight=1)
        
        # Aktuelle Ereignisse (Geburtstage, Jubiläen, etc.)
        events_frame = tk.Frame(self.content_frame, bg=LIGHT_COLOR)
        events_frame.pack(fill=tk.X, pady=(20, 0))
        
        events_title = tk.Label(events_frame, text="Anstehende Ereignisse", font=("Arial", 14, "bold"), fg=DARK_COLOR, bg=LIGHT_COLOR)
        events_title.pack(anchor=tk.W, pady=(0, 10))
        
        events_container = tk.Frame(events_frame, bg="white", bd=1, relief=tk.SOLID)
        events_container.pack(fill=tk.X)
        
        # Ereignisse laden
        events = self.get_upcoming_events()
        
        if events:
            for event in events:
                event_frame = tk.Frame(events_container, bg="white", pady=5)
                event_frame.pack(fill=tk.X, padx=10)
                
                event_icon = tk.Label(event_frame, text=event["icon"], font=("Arial", 16), bg="white")
                event_icon.pack(side=tk.LEFT, padx=(5, 10))
                
                event_text = tk.Label(event_frame, text=event["text"], font=("Arial", 11), anchor=tk.W, bg="white")
                event_text.pack(side=tk.LEFT, fill=tk.X)
                
                event_date = tk.Label(event_frame, text=event["date"], font=("Arial", 10), fg="gray", bg="white")
                event_date.pack(side=tk.RIGHT, padx=10)
                
                # Trennlinie, außer für das letzte Element
                if event != events[-1]:
                    separator = ttk.Separator(events_container, orient="horizontal")
                    separator.pack(fill=tk.X, padx=20)
        else:
            no_events = tk.Label(events_container, text="Keine anstehenden Ereignisse", font=("Arial", 11), fg="gray", bg="white", pady=10)
            no_events.pack()
        
        self.update_status("Dashboard geladen")
    
    def create_stat_card(self, parent, title, value, icon, color):
        card = tk.Frame(parent, bg="white", bd=1, relief=tk.SOLID, padx=15, pady=15)
        card.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        icon_label = tk.Label(card, text=icon, font=("Arial", 24), fg=color, bg="white")
        icon_label.grid(row=0, column=0, rowspan=2, padx=(0, 15))
        
        title_label = tk.Label(card, text=title, font=("Arial", 11), fg=DARK_COLOR, bg="white")
        title_label.grid(row=0, column=1, sticky=tk.W)
        
        value_label = tk.Label(card, text=str(value), font=("Arial", 18, "bold"), fg=DARK_COLOR, bg="white")
        value_label.grid(row=1, column=1, sticky=tk.W)
    
    def get_employee_count(self):
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM employees WHERE status = 'Aktiv'")
        count = cursor.fetchone()[0]
        conn.close()
        return count
    
    def get_current_vacation_count(self):
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT COUNT(DISTINCT employee_id) FROM vacation 
            WHERE start_date <= ? AND end_date >= ? AND status = 'Genehmigt'
        """, (today, today))
        count = cursor.fetchone()[0]
        conn.close()
        return count
    
    def get_current_sick_count(self):
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT COUNT(DISTINCT employee_id) FROM sick_leave 
            WHERE start_date <= ? AND end_date >= ?
        """, (today, today))
        count = cursor.fetchone()[0]
        conn.close()
        return count
    
    def get_birthdays_this_month(self):
        current_month = datetime.datetime.now().month
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT COUNT(*) FROM employees 
            WHERE strftime('%m', birth_date) = ? AND status = 'Aktiv'
        """, (f"{current_month:02d}",))
        count = cursor.fetchone()[0]
        conn.close()
        return count
    
    def get_employees_by_department(self):
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT department, COUNT(*) 
            FROM employees 
            WHERE status = 'Aktiv' 
            GROUP BY department
        """)
        data = cursor.fetchall()
        conn.close()
        
        departments = [dept[0] if dept[0] else "Andere" for dept in data]
        counts = [dept[1] for dept in data]
        
        return departments, counts
    
    def get_vacation_by_month(self):
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT strftime('%m', start_date) as month, SUM(days) 
            FROM vacation 
            WHERE status = 'Genehmigt' AND strftime('%Y', start_date) = strftime('%Y', 'now')
            GROUP BY month
        """)
        data = {int(month): count for month, count in cursor.fetchall()}
        conn.close()
        
        # Alle Monate abdecken
        return [data.get(i, 0) for i in range(1, 13)]
    
    def get_sick_leave_by_month(self):
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT strftime('%m', start_date) as month, SUM(days) 
            FROM sick_leave 
            WHERE strftime('%Y', start_date) = strftime('%Y', 'now')
            GROUP BY month
        """)
        data = {int(month): count for month, count in cursor.fetchall()}
        conn.close()
        
        # Alle Monate abdecken
        return [data.get(i, 0) for i in range(1, 13)]
    
    def get_upcoming_events(self):
        events = []
        today = datetime.datetime.now().date()
        
        # Kommende Geburtstage
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        
        # Geburtstage in den nächsten 30 Tagen
        cursor.execute("""
            SELECT first_name, last_name, birth_date 
            FROM employees 
            WHERE status = 'Aktiv' 
            ORDER BY strftime('%m-%d', birth_date)
        """)
        
        employees = cursor.fetchall()
        for emp in employees:
            if emp[2]:  # Wenn Geburtsdatum vorhanden
                try:
                    birth_date = datetime.datetime.strptime(emp[2], "%Y-%m-%d").date()
                    
                    # Nächster Geburtstag in diesem Jahr
                    next_birthday = datetime.date(today.year, birth_date.month, birth_date.day)
                    
                    # Falls der Geburtstag dieses Jahr schon vorbei ist, zum nächsten Jahr wechseln
                    if next_birthday < today:
                        next_birthday = datetime.date(today.year + 1, birth_date.month, birth_date.day)
                    
                    # Nur Ereignisse in den nächsten 30 Tagen anzeigen
                    delta = (next_birthday - today).days
                    if 0 <= delta <= 30:
                        events.append({
                            "icon": "🎂",
                            "text": f"Geburtstag von {emp[0]} {emp[1]}",
                            "date": next_birthday.strftime("%d.%m.%Y")
                        })
                except:
                    pass
        
        # Jubiläen (Mitarbeiter, die X Jahre im Unternehmen sind)
        cursor.execute("""
            SELECT first_name, last_name, hire_date 
            FROM employees 
            WHERE status = 'Aktiv' 
            ORDER BY hire_date
        """)
        
        employees = cursor.fetchall()
        for emp in employees:
            if emp[2]:  # Wenn Einstellungsdatum vorhanden
                try:
                    hire_date = datetime.datetime.strptime(emp[2], "%Y-%m-%d").date()
                    
                    # Jubiläumsdatum in diesem Jahr
                    years_employed = today.year - hire_date.year
                    anniversary_date = datetime.date(today.year, hire_date.month, hire_date.day)
                    
                    # Falls das Jubiläum dieses Jahr schon vorbei ist, zum nächsten Jahr wechseln
                    if anniversary_date < today:
                        anniversary_date = datetime.date(today.year + 1, hire_date.month, hire_date.day)
                        years_employed += 1
                    
                    # Nur Ereignisse in den nächsten 30 Tagen und bei rundem Jubiläum (5, 10, 15, etc. Jahre)
                    delta = (anniversary_date - today).days
                    if 0 <= delta <= 30 and years_employed > 0 and years_employed % 5 == 0:
                        events.append({
                            "icon": "🏆",
                            "text": f"{years_employed}-jähriges Jubiläum von {emp[0]} {emp[1]}",
                            "date": anniversary_date.strftime("%d.%m.%Y")
                        })
                except:
                    pass
        
        # Kommender Urlaub
        cursor.execute("""
            SELECT e.first_name, e.last_name, v.start_date, v.end_date 
            FROM vacation v
            JOIN employees e ON v.employee_id = e.id
            WHERE v.status = 'Genehmigt' AND v.start_date >= ?
            ORDER BY v.start_date
            LIMIT 5
        """, (today.strftime("%Y-%m-%d"),))
        
        vacations = cursor.fetchall()
        for vac in vacations:
            try:
                start_date = datetime.datetime.strptime(vac[2], "%Y-%m-%d").date()
                
                # Nur Urlaub in den nächsten 14 Tagen anzeigen
                delta = (start_date - today).days
                if 0 <= delta <= 14:
                    events.append({
                        "icon": "🏖️",
                        "text": f"{vac[0]} {vac[1]} ist im Urlaub",
                        "date": f"{format_date(vac[2])} - {format_date(vac[3])}"
                    })
            except:
                pass
        
        conn.close()
        
        # Nach Datum sortieren
        events.sort(key=lambda x: datetime.datetime.strptime(x["date"].split(" - ")[0], "%d.%m.%Y"))
        
        return events[:10]  # Maximal 10 Ereignisse anzeigen
    
    def show_employees(self):
        self.clear_content()
        self.header_title.config(text="Mitarbeiterverwaltung")
        self.highlight_menu_button("Mitarbeiter")
        
        # Toolleiste erstellen
        toolbar = tk.Frame(self.content_frame, bg=LIGHT_COLOR)
        toolbar.pack(fill=tk.X, pady=(0, 10))
        
        # Suchfeld
        search_frame = tk.Frame(toolbar, bg=LIGHT_COLOR)
        search_frame.pack(side=tk.LEFT)
        
        search_label = tk.Label(search_frame, text="Suche:", bg=LIGHT_COLOR)
        search_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda name, index, mode: self.filter_employees())
        
        search_entry = tk.Entry(search_frame, textvariable=self.search_var, width=20)
        search_entry.pack(side=tk.LEFT)
        
        # Abteilungsfilter
        filter_frame = tk.Frame(toolbar, bg=LIGHT_COLOR)
        filter_frame.pack(side=tk.LEFT, padx=20)
        
        filter_label = tk.Label(filter_frame, text="Abteilung:", bg=LIGHT_COLOR)
        filter_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.department_var = tk.StringVar()
        self.department_var.set("Alle")
        self.department_var.trace_add("write", lambda name, index, mode: self.filter_employees())
        
        departments = ["Alle"] + self.get_departments()
        department_menu = ttk.Combobox(filter_frame, textvariable=self.department_var, values=departments, state="readonly", width=15)
        department_menu.pack(side=tk.LEFT)
        
        # Statusfilter
        status_frame = tk.Frame(toolbar, bg=LIGHT_COLOR)
        status_frame.pack(side=tk.LEFT)
        
        status_label = tk.Label(status_frame, text="Status:", bg=LIGHT_COLOR)
        status_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.status_var = tk.StringVar()
        self.status_var.set("Alle")
        self.status_var.trace_add("write", lambda name, index, mode: self.filter_employees())
        
        status_menu = ttk.Combobox(status_frame, textvariable=self.status_var, values=["Alle", "Aktiv", "Inaktiv"], state="readonly", width=10)
        status_menu.pack(side=tk.LEFT)
        
        # Buttons für Aktionen
        button_frame = tk.Frame(toolbar, bg=LIGHT_COLOR)
        button_frame.pack(side=tk.RIGHT)
        
        add_button = tk.Button(
            button_frame,
            text="Neuer Mitarbeiter",
            bg=THEME_COLOR,
            fg="white",
            padx=10,
            pady=2,
            relief=tk.FLAT,
            command=self.add_employee
        )
        add_button.pack(side=tk.RIGHT, padx=5)
        
        export_button = tk.Button(
            button_frame,
            text="Exportieren",
            bg=DARK_COLOR,
            fg="white",
            padx=10,
            pady=2,
            relief=tk.FLAT,
            command=lambda: self.export_data("employees")
        )
        export_button.pack(side=tk.RIGHT, padx=5)
        
        # Tabelle für Mitarbeiter
        table_frame = tk.Frame(self.content_frame, bg="white")
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar_y = tk.Scrollbar(table_frame)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        scrollbar_x = tk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Treeview für tabellarische Anzeige
        columns = ("id", "employee_id", "name", "department", "position", "hire_date", "status")
        self.employee_tree = ttk.Treeview(
            table_frame, 
            columns=columns,
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )
        
        # Spalten konfigurieren
        self.employee_tree.heading("id", text="ID")
        self.employee_tree.heading("employee_id", text="Personalnummer")
        self.employee_tree.heading("name", text="Name")
        self.employee_tree.heading("department", text="Abteilung")
        self.employee_tree.heading("position", text="Position")
        self.employee_tree.heading("hire_date", text="Einstellungsdatum")
        self.employee_tree.heading("status", text="Status")
        
        self.employee_tree.column("id", width=50, anchor=tk.CENTER)
        self.employee_tree.column("employee_id", width=120, anchor=tk.CENTER)
        self.employee_tree.column("name", width=200)
        self.employee_tree.column("department", width=150)
        self.employee_tree.column("position", width=150)
        self.employee_tree.column("hire_date", width=120, anchor=tk.CENTER)
        self.employee_tree.column("status", width=100, anchor=tk.CENTER)
        
        self.employee_tree.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbars mit Treeview verbinden
        scrollbar_y.config(command=self.employee_tree.yview)
        scrollbar_x.config(command=self.employee_tree.xview)
        
        # Kontextmenü für Zeilen
        self.employee_context_menu = tk.Menu(self.employee_tree, tearoff=0)
        self.employee_context_menu.add_command(label="Details anzeigen", command=self.view_employee)
        self.employee_context_menu.add_command(label="Bearbeiten", command=self.edit_employee)
        self.employee_context_menu.add_separator()
        self.employee_context_menu.add_command(label="Urlaub beantragen", command=self.request_vacation)
        self.employee_context_menu.add_command(label="Krankmeldung eintragen", command=self.report_sick_leave)
        self.employee_context_menu.add_separator()
        self.employee_context_menu.add_command(label="Dokument hochladen", command=self.upload_document)
        self.employee_context_menu.add_separator()
        self.employee_context_menu.add_command(label="Status ändern", command=self.change_employee_status)
        
        self.employee_tree.bind("<Button-3>", self.show_employee_context_menu)
        self.employee_tree.bind("<Double-1>", lambda event: self.view_employee())
        
        # Mitarbeiterdaten laden
        self.load_employees()
        self.update_status("Mitarbeiterverwaltung geladen")
    
    def show_employee_context_menu(self, event):
        try:
            iid = self.employee_tree.identify_row(event.y)
            if iid:
                self.employee_tree.selection_set(iid)
                self.employee_context_menu.post(event.x_root, event.y_root)
        except:
            pass
    
    def get_departments(self):
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM departments ORDER BY name")
        departments = [row[0] for row in cursor.fetchall()]
        conn.close()
        return departments
    
    def load_employees(self):
        # Alle bestehenden Einträge löschen
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)
        
        conn = sqlite3.connect(DATABASE_PATH)
        conn.row_factory = sqlite3.Row  # Ermöglicht Zugriff auf Spalten nach Namen
        cursor = conn.cursor()
        
        # Mitarbeiter laden
        cursor.execute("""
            SELECT id, employee_id, first_name, last_name, department, position, hire_date, status 
            FROM employees
            ORDER BY last_name, first_name
        """)
        
        for row in cursor.fetchall():
            formatted_date = format_date(row['hire_date']) if row['hire_date'] else ""
            
            self.employee_tree.insert(
                "", 
                tk.END, 
                values=(
                    row['id'],
                    row['employee_id'],
                    f"{row['last_name']}, {row['first_name']}",
                    row['department'],
                    row['position'],
                    formatted_date,
                    row['status']
                )
            )
        
        conn.close()
        
        # Filter anwenden, falls aktiv
        self.filter_employees()
    
    def filter_employees(self):
        search_term = self.search_var.get().lower()
        department_filter = self.department_var.get()
        status_filter = self.status_var.get()
        
        for item in self.employee_tree.get_children():
            values = self.employee_tree.item(item, "values")
            
            # Anwendung der Filter
            show_item = True
            
            # Suchfilter
            if search_term:
                if not (search_term in values[2].lower() or search_term in values[1].lower()):
                    show_item = False
            
            # Abteilungsfilter
            if department_filter != "Alle" and values[3] != department_filter:
                show_item = False
            
            # Statusfilter
            if status_filter != "Alle" and values[6] != status_filter:
                show_item = False
            
            # Sichtbarkeit anpassen
            if show_item:
                self.employee_tree.item(item, tags=())
            else:
                self.employee_tree.detach(item)
    
    def add_employee(self):
        EmployeeDialog(self.root, self.load_employees)
    
    def view_employee(self):
        selected_item = self.employee_tree.selection()
        if not selected_item:
            messagebox.showinfo("Information", "Bitte wählen Sie einen Mitarbeiter aus.")
            return
        
        employee_id = self.employee_tree.item(selected_item[0], "values")[0]
        EmployeeDetailDialog(self.root, employee_id, self.load_employees)
    
    def edit_employee(self):
        selected_item = self.employee_tree.selection()
        if not selected_item:
            messagebox.showinfo("Information", "Bitte wählen Sie einen Mitarbeiter aus.")
            return
        
        employee_id = self.employee_tree.item(selected_item[0], "values")[0]
        EmployeeDialog(self.root, self.load_employees, employee_id)
    
    def request_vacation(self):
        selected_item = self.employee_tree.selection()
        if not selected_item:
            messagebox.showinfo("Information", "Bitte wählen Sie einen Mitarbeiter aus.")
            return
        
        employee_id = self.employee_tree.item(selected_item[0], "values")[0]
        VacationDialog(self.root, employee_id)
    
    def report_sick_leave(self):
        selected_item = self.employee_tree.selection()
        if not selected_item:
            messagebox.showinfo("Information", "Bitte wählen Sie einen Mitarbeiter aus.")
            return
        
        employee_id = self.employee_tree.item(selected_item[0], "values")[0]
        SickLeaveDialog(self.root, employee_id)
    
    def upload_document(self):
        selected_item = self.employee_tree.selection()
        if not selected_item:
            messagebox.showinfo("Information", "Bitte wählen Sie einen Mitarbeiter aus.")
            return
        
        employee_id = self.employee_tree.item(selected_item[0], "values")[0]
        DocumentUploadDialog(self.root, employee_id)
    
    def change_employee_status(self):
        selected_item = self.employee_tree.selection()
        if not selected_item:
            messagebox.showinfo("Information", "Bitte wählen Sie einen Mitarbeiter aus.")
            return
        
        employee_data = self.employee_tree.item(selected_item[0], "values")
        employee_id = employee_data[0]
        current_status = employee_data[6]
        
        new_status = "Inaktiv" if current_status == "Aktiv" else "Aktiv"
        
        if messagebox.askyesno("Status ändern", f"Möchten Sie den Status des Mitarbeiters von '{current_status}' zu '{new_status}' ändern?"):
            conn = sqlite3.connect(DATABASE_PATH)
            cursor = conn.cursor()
            
            try:
                current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                cursor.execute("""
                    UPDATE employees 
                    SET status = ?, updated_at = ?
                    WHERE id = ?
                """, (new_status, current_time, employee_id))
                
                conn.commit()
                self.load_employees()
                self.update_status(f"Mitarbeiterstatus erfolgreich geändert")
                
                logger.info(f"Mitarbeiterstatus geändert: ID {employee_id}, neuer Status: {new_status}")
            except Exception as e:
                conn.rollback()
                messagebox.showerror("Fehler", f"Fehler beim Ändern des Mitarbeiterstatus: {str(e)}")
                logger.error(f"Fehler beim Ändern des Mitarbeiterstatus: {e}")
            finally:
                conn.close()
    
    def export_data(self, data_type):
        if data_type == "employees":
            # Exportdialog
            file_types = [
                ("CSV-Dateien", "*.csv"),
                ("Excel-Dateien", "*.xlsx"),
                ("PDF-Dateien", "*.pdf"),
                ("Alle Dateien", "*.*")
            ]
            
            export_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=file_types,
                initialdir=EXPORT_PATH,
                title="Mitarbeiterdaten exportieren"
            )
            
            if not export_path:
                return
            
            # Mitarbeiterdaten abrufen
            conn = sqlite3.connect(DATABASE_PATH)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT e.*, d.name as department_name
                FROM employees e
                LEFT JOIN departments d ON e.department = d.name
            """)
            
            employees = cursor.fetchall()
            conn.close()
            
            try:
                # In verschiedene Formate exportieren
                if export_path.endswith(".csv"):
                    self.export_to_csv(export_path, employees)
                elif export_path.endswith(".xlsx"):
                    messagebox.showinfo("Information", "Excel-Export ist in dieser Version nicht verfügbar.")
                elif export_path.endswith(".pdf"):
                    self.export_to_pdf(export_path, employees)
                else:
                    self.export_to_csv(export_path, employees)  # Standardmäßig als CSV
                
                self.update_status(f"Mitarbeiterdaten erfolgreich exportiert nach {export_path}")
                
                # Export-Ordner öffnen
                if os.path.exists(os.path.dirname(export_path)):
                    webbrowser.open(os.path.dirname(export_path))
            except Exception as e:
                messagebox.showerror("Exportfehler", f"Fehler beim Exportieren der Daten: {str(e)}")
                logger.error(f"Exportfehler: {e}")
    
    def export_to_csv(self, filepath, data):
        with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
            # Feldnamen aus den Daten extrahieren
            fieldnames = [key for key in data[0].keys()]
            
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for row in data:
                writer.writerow({key: row[key] for key in fieldnames})
    
    def export_to_pdf(self, filepath, data):
        pdf = FPDF()
        pdf.add_page()
        
        # Titel
        pdf.set_font("Arial", "B", 16)
        pdf.cell(190, 10, "Mitarbeiterliste", 0, 1, "C")
        pdf.ln(10)
        
        # Tabellenkopf
        pdf.set_font("Arial", "B", 10)
        pdf.cell(15, 10, "ID", 1, 0, "C")
        pdf.cell(40, 10, "Name", 1, 0, "C")
        pdf.cell(30, 10, "Personalnr.", 1, 0, "C")
        pdf.cell(40, 10, "Abteilung", 1, 0, "C")
        pdf.cell(40, 10, "Position", 1, 0, "C")
        pdf.cell(25, 10, "Status", 1, 1, "C")
        
        # Tabellendaten
        pdf.set_font("Arial", "", 10)
        for row in data:
            pdf.cell(15, 10, str(row['id']), 1, 0, "C")
            pdf.cell(40, 10, f"{row['first_name']} {row['last_name']}", 1, 0, "L")
            pdf.cell(30, 10, str(row['employee_id']), 1, 0, "L")
            pdf.cell(40, 10, str(row['department']), 1, 0, "L")
            pdf.cell(40, 10, str(row['position']), 1, 0, "L")
            pdf.cell(25, 10, str(row['status']), 1, 1, "C")
        
        # Fußzeile
        pdf.ln(10)
        pdf.set_font("Arial", "I", 8)
        pdf.cell(0, 10, f"Erstellt mit {APP_NAME} am {datetime.datetime.now().strftime('%d.%m.%Y %H:%M')}", 0, 0, "L")
        
        pdf.output(filepath)
    
    def show_vacation(self):
        self.clear_content()
        self.header_title.config(text="Urlaubsverwaltung")
        self.highlight_menu_button("Urlaub")
        
        