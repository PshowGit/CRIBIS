"""CRIBIS Export - applicazione desktop per esportare dati aziendali da CRIBIS X in un file Excel (.xls)."""
import json
import os
import queue
import re
import random
import sys
import threading
import time
import traceback
from datetime import datetime

import tkinter as tk

import customtkinter as ctk
import pandas as pd
import xlwt
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


# --- Percorsi: tutto vive accanto all'eseguibile (o allo script) ---
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = app_dir()
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
LOG_PATH = os.path.join(BASE_DIR, "cribis_export.log")

DEFAULT_CONFIG = {
    "LOGIN_URL": "https://www2.cribisx.com/Account/LogOn?ReturnUrl=%2f",
    "USERNAME": "",
    "PASSWORD": "",
    "INPUT_FILE": "",
    "OUTPUT_FILE": "cribis_dati_export.xls",
    "CHROME_DRIVER_PATH": "chromedriver.exe",
    "PROXY": "",
    "SHOW_BROWSER": True,
    "MIN_DELAY": 4,
    "MAX_DELAY": 9,
    "LONG_BREAK_EVERY": 100,
    "LONG_BREAK_SECONDS": 180,
    "DRIVER_RESTART_EVERY": 25,
    "USER_AGENTS_FILE": "user_agents.txt",
}

FALLBACK_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36"
)

SEARCH_TERM_HEADER = "Termine di Ricerca"
FIELDNAMES = [
    SEARCH_TERM_HEADER, "nome_azienda", "sede", "codice_fiscale", "iva",
    "cciaa_rea", "natura_giuridica", "attività", "attività_dichiarata",
]


def resolve_path(path):
    """I percorsi relativi sono sempre relativi alla cartella dell'applicazione."""
    return path if os.path.isabs(path) else os.path.join(BASE_DIR, path)


def output_path_for(config):
    """Percorso del file Excel di output (estensione sempre .xls)."""
    return os.path.splitext(resolve_path(config["OUTPUT_FILE"]))[0] + ".xls"


def load_config():
    config = dict(DEFAULT_CONFIG)
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            config.update(json.load(f))
    except (FileNotFoundError, json.JSONDecodeError):
        pass
    return config


def save_config(config):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def load_user_agents(path):
    try:
        with open(resolve_path(path), "r", encoding="utf-8") as f:
            agents = [line.strip() for line in f if line.strip()]
        if agents:
            return agents
    except OSError:
        pass
    return [FALLBACK_USER_AGENT]


# --- Lettura file di input ---
def _read_text(path):
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            with open(path, "r", encoding=encoding) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    raise ValueError("codifica del file non riconosciuta")


def read_search_terms(path):
    """Restituisce i termini di ricerca: una riga per i .txt, la prima colonna valida per csv/xls/xlsx."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".txt":
        values = _read_text(path).splitlines()
    elif ext in (".csv", ".xls", ".xlsx"):
        if ext == ".csv":
            df = pd.read_csv(path, sep=None, engine="python", header=None, dtype=str,
                             encoding="utf-8-sig", encoding_errors="replace", skip_blank_lines=True)
        else:
            df = pd.read_excel(path, header=None, dtype=str)
        values = []
        for column in df.columns:
            candidate = df[column].dropna().astype(str).str.strip()
            candidate = candidate[candidate != ""]
            if not candidate.empty:
                values = candidate.tolist()
                break
    else:
        raise ValueError(f"formato '{ext}' non supportato (usare .txt, .csv, .xls o .xlsx)")

    terms = [v.strip() for v in values if v and v.strip()]
    # Una prima riga senza cifre è un'intestazione (es. "Partita IVA"), non un termine di ricerca.
    if ext != ".txt" and terms and not any(ch.isdigit() for ch in terms[0]):
        terms = terms[1:]
    return terms


def describe_terms(terms):
    """Descrive il contenuto del file: codici IVA/fiscali oppure ragioni sociali."""
    is_vat = [bool(re.fullmatch(r"\d{9,11}|[A-Za-z0-9]{16}", t.replace(" ", ""))) for t in terms]
    vat = sum(is_vat)
    names = len(terms) - vat
    parts = []
    if vat:
        parts.append(f"{vat} IVA" if vat > 1 else "1 IVA")
    if names:
        parts.append(f"{names} Ragioni sociali" if names > 1 else "1 Ragione sociale")
    return " e ".join(parts) + (" trovate" if len(terms) > 1 else " trovata")


# --- Scraping ---
def safe_find(driver, xpath):
    try:
        return driver.find_element(By.XPATH, xpath).text.strip()
    except (NoSuchElementException, TimeoutException):
        return "N/D"


def extract_data(driver):
    field = "(//span[contains(@class,'upcas') and contains(@class,'link-silver-small')]/following-sibling::p)[{}]"
    return {
        "nome_azienda": safe_find(driver, "//*[@id='anagReport']/div[1]/div[1]/h1"),
        "sede": safe_find(driver, field.format(1)),
        "codice_fiscale": safe_find(driver, field.format(2)),
        "iva": safe_find(driver, field.format(3)),
        "cciaa_rea": safe_find(driver, field.format(4)),
        "natura_giuridica": safe_find(driver, field.format(5)),
        "attività": safe_find(driver, field.format(6)),
        "attività_dichiarata": safe_find(driver, field.format(7)),
    }


class XlsWriter:
    """Scrive il foglio .xls riga per riga. Il file viene risalvato a ogni riga,
    così i dati già raccolti non si perdono se l'esportazione viene interrotta."""

    def __init__(self, path, fieldnames):
        self.path = path
        self.fieldnames = fieldnames
        self.workbook = xlwt.Workbook(encoding="utf-8")
        self.sheet = self.workbook.add_sheet("Dati CRIBIS")
        self.next_row = 1
        bold = xlwt.easyxf("font: bold on")
        for col, name in enumerate(fieldnames):
            self.sheet.write(0, col, name, bold)
            self.sheet.col(col).width = 256 * (30 if col in (0, 1, 2) else 22)

    def writerow(self, row):
        if self.next_row > 65535:
            raise ValueError("raggiunto il limite di 65.535 righe del formato .xls")
        for col, name in enumerate(self.fieldnames):
            self.sheet.write(self.next_row, col, row.get(name, "N/D"))
        self.next_row += 1

    def save(self):
        try:
            self.workbook.save(self.path)
        except PermissionError as e:
            raise OSError(f"impossibile scrivere '{self.path}': il file è aperto in Excel? Chiuderlo e riprovare.") from e

    flush = save

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        try:
            self.save()
        except OSError:
            if exc[0] is None:
                raise


def quit_quietly(driver):
    try:
        driver.quit()
    except Exception:
        pass


class Scraper:
    """Esegue l'esportazione. Pensato per girare in un thread dedicato, mai nel thread della GUI."""

    def __init__(self, config, terms, log, progress):
        self.config = config
        self.terms = terms
        self.log = log
        self.progress = progress
        self.stop_event = threading.Event()
        self.driver = None
        self.output_path = output_path_for(config)
        self.summary = None
        self.user_agents = load_user_agents(config["USER_AGENTS_FILE"])

    def request_stop(self):
        """Chiamabile da qualunque thread: interrompe anche le attese Selenium chiudendo il browser."""
        self.stop_event.set()
        driver = self.driver
        if driver is not None:
            threading.Thread(target=quit_quietly, args=(driver,), daemon=True).start()

    def _sleep(self, seconds):
        """Attesa interrompibile. Restituisce True se è stato richiesto lo stop."""
        return self.stop_event.wait(seconds)

    def _setup_driver(self, user_agent):
        try:
            service = Service(ChromeDriverManager().install())
        except Exception as e:
            # Ripiego: Selenium Manager (incluso in Selenium) scarica il driver giusto per conto suo.
            self.log(f"⚠️ Download di ChromeDriver non riuscito ({e}). Provo con Selenium Manager.")
            service = Service()

        options = webdriver.ChromeOptions()
        if self.config["SHOW_BROWSER"]:
            options.add_argument("--start-maximized")
        else:
            options.add_argument("--headless=new")
            options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        options.add_argument(f"--user-agent={user_agent}")
        if self.config["PROXY"]:
            options.add_argument(f"--proxy-server={self.config['PROXY']}")

        driver = webdriver.Chrome(service=service, options=options)
        self.driver = driver
        try:
            driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": """
                    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                    window.navigator.chrome = { runtime: {} };
                    Object.defineProperty(navigator, 'languages', {get: () => ['it-IT','it']});
                    Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
                """
            })
        except Exception:
            pass
        return driver

    def _login(self, driver):
        self.log("Tentativo di login...")
        driver.get(self.config["LOGIN_URL"])
        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "UserName")))
            driver.find_element(By.NAME, "UserName").send_keys(self.config["USERNAME"])
            driver.find_element(By.NAME, "Password").send_keys(self.config["PASSWORD"])
            driver.find_element(By.NAME, "submit_pwd").click()
            WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "search-input")))
            self.log("✅ Login effettuato con successo.")
            return True
        except TimeoutException:
            if not self.stop_event.is_set():
                self.log("❌ Timeout durante il login. Controllare credenziali e URL.")
            return False
        except WebDriverException as e:
            if not self.stop_event.is_set():
                self.log(f"❌ Errore durante il login: {e.msg or e}")
            return False

    def _start_session(self):
        """Apre un nuovo browser con user-agent casuale e fa il login."""
        driver = self._setup_driver(random.choice(self.user_agents))
        if not self._login(driver):
            quit_quietly(driver)
            self.driver = None
            return None
        return driver

    def _search_term(self, driver, term):
        search_input = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.ID, "search-input")))
        search_input.clear()
        search_input.send_keys(term)

        WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "search-button"))).click()
        WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='root']//h2/a"))).click()

        try:
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='company-details-col-left']//p/i"))).click()
        except (TimeoutException, NoSuchElementException):
            pass
        return extract_data(driver)

    def run(self):
        cfg = self.config
        total = len(self.terms)
        results_found = 0
        done = 0
        error = None
        try:
            sheet = XlsWriter(self.output_path, FIELDNAMES)
            sheet.save()  # verifica subito che il file sia scrivibile, prima di aprire il browser
            self.log("Avvio del browser...")
            self.driver = self._start_session()
            if self.driver is None:
                if not self.stop_event.is_set():
                    error = "Login fallito: controllare utente, password e connessione."
                    self.log("❌ Login fallito. Operazione annullata.")
                return

            with sheet as writer:

                for idx, term in enumerate(self.terms, start=1):
                    if self.stop_event.is_set():
                        break
                    self.log(f"[{idx}/{total}] 🔍 Ricerca per: '{term}'")

                    restart_every = cfg["DRIVER_RESTART_EVERY"]
                    if restart_every and idx != 1 and (idx - 1) % restart_every == 0:
                        self.log("🔄 Riavvio del browser e cambio user-agent...")
                        quit_quietly(self.driver)
                        self.driver = self._start_session()
                        if self.driver is None:
                            if not self.stop_event.is_set():
                                error = "Login fallito dopo il riavvio del browser."
                                self.log("❌ Login fallito dopo il riavvio. Operazione interrotta.")
                            break

                    row = {name: "N/D" for name in FIELDNAMES}
                    row[SEARCH_TERM_HEADER] = term
                    max_retries = 3
                    for attempt in range(1, max_retries + 1):
                        try:
                            row.update(self._search_term(self.driver, term))
                            break
                        except (TimeoutException, NoSuchElementException, WebDriverException) as e:
                            if self.stop_event.is_set():
                                break
                            wait = 2 ** attempt + random.uniform(0.5, 1.5)
                            self.log(f"⚠️ Tentativo {attempt}/{max_retries} fallito per '{term}': "
                                     f"{getattr(e, 'msg', None) or type(e).__name__}. Attesa {wait:.1f}s.")
                            if self._sleep(wait):
                                break
                        except Exception as e:
                            self.log(f"❌ Errore imprevisto per '{term}': {e}")
                            self.log(traceback.format_exc())
                            break

                    if self.stop_event.is_set():
                        break

                    if row["nome_azienda"] == "N/D":
                        self.log(f"⚠️ Nessun risultato per '{term}'")
                    else:
                        results_found += 1
                        self.log(f"✅ Dati salvati per '{term}': {row['nome_azienda']}")
                    writer.writerow(row)
                    writer.flush()
                    done = idx
                    self.progress(done, total)

                    if self._sleep(random.uniform(cfg["MIN_DELAY"], cfg["MAX_DELAY"])):
                        break
                    long_every = cfg["LONG_BREAK_EVERY"]
                    if long_every and idx % long_every == 0 and idx != total:
                        self.log(f"⏸ Pausa lunga di {cfg['LONG_BREAK_SECONDS']}s ogni {long_every} richieste.")
                        if self._sleep(cfg["LONG_BREAK_SECONDS"]):
                            break
        except Exception as e:
            error = str(e)
            self.log(f"❌ Errore: {e}")
            self.log(traceback.format_exc())
        finally:
            if self.driver is not None:
                quit_quietly(self.driver)
                self.driver = None
            if self.stop_event.is_set():
                status, state = "stopped", "interrotta dall'utente"
            elif error:
                status, state = "error", "terminata con errore"
            else:
                status, state = "completed", "completata"
            self.summary = {"status": status, "error": error, "done": done, "total": total,
                            "found": results_found, "output_path": self.output_path}
            self.log(f"--- Esportazione {state} ---")
            self.log(f"Ricerche effettuate: {done}/{total} - Risultati trovati: {results_found}")
            self.log(f"📁 File Excel: {self.output_path}")


# --- GUI ---
WINDOW_BG = "#1d2a30"
STATUS_IDLE = "#c9d1d6"
STATUS_RUNNING = "#f2c14e"
STATUS_OK = "#5ff0a8"
STATUS_STOPPED = "#ffa94d"
STATUS_ERROR = "#ff7a7a"
PANEL_BG = "#30353a"
PANEL_BORDER = "#454c52"
FIELD_BG = "#1c2024"
LOG_BG = "#14181b"
LOG_TEXT = "#7fe3dc"
TEAL = "#1fa8a0"
TEAL_DARK = "#178680"

class App(ctk.CTk, TkinterDnD.DnDWrapper):
    POLL_MS = 100

    def __init__(self):
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("CRIBIS Export Tool")
        self.geometry("780x860")
        self.minsize(680, 800)

        self.config_data = load_config()
        self.messages = queue.Queue()
        self.scraper = None
        self.worker = None
        self.terms = []

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(self.POLL_MS, self._poll_messages)

    def _panel(self, row, title, **grid):
        """Pannello con titolo di sezione, come nel mockup. Restituisce il frame interno."""
        panel = ctk.CTkFrame(self, fg_color=PANEL_BG, border_width=1, border_color=PANEL_BORDER, corner_radius=10)
        panel.grid(row=row, column=0, padx=16, pady=(0, 12), sticky="nsew", **grid)
        panel.columnconfigure(0, weight=1)
        ctk.CTkLabel(panel, text=title, font=ctk.CTkFont(size=18, weight="bold"), anchor="w").grid(
            row=0, column=0, padx=14, pady=(10, 6), sticky="w")
        return panel

    def _build_ui(self):
        self.configure(fg_color=WINDOW_BG)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        label_font = ctk.CTkFont(size=15)
        entry_opts = dict(height=34, fg_color=FIELD_BG, border_color=PANEL_BORDER, text_color="white")

        # Impostazioni account
        account = self._panel(0, "⚙  IMPOSTAZIONI ACCOUNT")
        fields = ctk.CTkFrame(account, fg_color="transparent")
        fields.grid(row=1, column=0, padx=14, sticky="ew")
        fields.columnconfigure(1, weight=1)
        self.user_var = ctk.StringVar(value=self.config_data["USERNAME"])
        self.pass_var = ctk.StringVar(value=self.config_data["PASSWORD"])
        ctk.CTkLabel(fields, text="Utente:", font=label_font).grid(row=0, column=0, padx=(0, 12), pady=4, sticky="w")
        ctk.CTkEntry(fields, textvariable=self.user_var, **entry_opts).grid(row=0, column=1, pady=4, sticky="ew")
        ctk.CTkLabel(fields, text="Password:", font=label_font).grid(row=1, column=0, padx=(0, 12), pady=4, sticky="w")
        ctk.CTkEntry(fields, textvariable=self.pass_var, show="*", **entry_opts).grid(row=1, column=1, pady=4, sticky="ew")
        ctk.CTkButton(account, text="[ Salva Credenziali ]", height=36, fg_color=TEAL, hover_color=TEAL_DARK,
                      text_color="white", font=ctk.CTkFont(size=16, weight="bold"),
                      command=self._save_settings).grid(row=2, column=0, padx=14, pady=(8, 14), sticky="ew")

        # Input dati
        data = self._panel(1, "📂  INPUT DATI")
        self.drop_canvas = tk.Canvas(data, height=140, highlightthickness=0, bg=PANEL_BG, cursor="hand2")
        self.drop_canvas.grid(row=1, column=0, padx=14, sticky="ew")
        self.drop_canvas.bind("<Configure>", lambda _e: self._draw_drop_zone())
        self.drop_canvas.bind("<Button-1>", lambda _e: self._browse_input())
        self.drop_canvas.drop_target_register(DND_FILES)
        self.drop_canvas.dnd_bind("<<Drop>>", self._on_drop)
        self.file_label = ctk.CTkLabel(data, text="Nessun file selezionato", font=label_font, anchor="w")
        self.file_label.grid(row=2, column=0, padx=14, pady=(6, 12), sticky="w")

        # Processo e log
        process = self._panel(2, "🔄  PROCESSO E LOG")
        process.rowconfigure(2, weight=1)
        links = ctk.CTkFrame(process, fg_color="transparent")
        links.grid(row=0, column=0, padx=14, sticky="e")
        link_opts = dict(height=28, fg_color="transparent", border_width=1, border_color="#9aa1a7",
                         text_color="#d5dbe0", hover_color="#4a525a")
        ctk.CTkButton(links, text="📊 Apri file Excel", width=140, command=self._open_output_file,
                      **link_opts).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkButton(links, text="📄 Apri file di log", width=140, command=self._open_log_file,
                      **link_opts).grid(row=0, column=1)
        bar_row = ctk.CTkFrame(process, fg_color="transparent")
        bar_row.grid(row=1, column=0, padx=14, sticky="ew")
        bar_row.columnconfigure(0, weight=1)
        self.progress_bar = ctk.CTkProgressBar(bar_row, height=18, progress_color=TEAL, fg_color=FIELD_BG)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        self.progress_label = ctk.CTkLabel(bar_row, text="0%", width=50, font=ctk.CTkFont(size=16))
        self.progress_label.grid(row=0, column=1, padx=(10, 0))
        self.show_browser_var = ctk.BooleanVar(value=bool(self.config_data["SHOW_BROWSER"]))
        ctk.CTkSwitch(bar_row, text="Mostra il browser durante l'esecuzione", variable=self.show_browser_var,
                      progress_color=TEAL, font=ctk.CTkFont(size=14)).grid(
            row=1, column=0, columnspan=2, pady=(10, 0), sticky="w")
        self.status_label = ctk.CTkLabel(bar_row, text="In attesa", anchor="w", justify="left", wraplength=640,
                                         font=ctk.CTkFont(size=15, weight="bold"), text_color=STATUS_IDLE)
        self.status_label.grid(row=2, column=0, columnspan=2, pady=(8, 0), sticky="w")

        self.log_box = ctk.CTkTextbox(process, state="disabled", wrap="word", height=170, fg_color=LOG_BG,
                                      text_color=LOG_TEXT, border_width=1, border_color=PANEL_BORDER,
                                      font=ctk.CTkFont(family="Consolas", size=13))
        self.log_box.grid(row=2, column=0, padx=14, pady=10, sticky="nsew")

        buttons = ctk.CTkFrame(process, fg_color="transparent")
        buttons.grid(row=3, column=0, padx=14, pady=(0, 14), sticky="ew")
        buttons.columnconfigure((0, 1, 2), weight=1, uniform="btn")
        button_font = ctk.CTkFont(size=18, weight="bold")
        self.start_btn = ctk.CTkButton(buttons, text="[ ▶ AVVIA ]", height=46, font=button_font,
                                       fg_color="#2c5e50", hover_color="#367362", border_width=2,
                                       border_color="#3fcf8e", text_color="#5ff0a8", command=self._start)
        self.start_btn.grid(row=0, column=0, padx=(0, 6), sticky="ew")
        self.clear_btn = ctk.CTkButton(buttons, text="[ 🧹 PULISCI ]", height=46, font=button_font,
                                       fg_color="#3a4047", hover_color="#4a525a", border_width=2,
                                       border_color="#9aa1a7", text_color="#d5dbe0", command=self._clear)
        self.clear_btn.grid(row=0, column=1, padx=6, sticky="ew")
        self.stop_btn = ctk.CTkButton(buttons, text="[ ■ INTERROMPI ]", height=46, font=button_font,
                                      fg_color="#6a2f33", hover_color="#823a3f", border_width=2,
                                      border_color="#e05a5a", text_color="#ff7a7a", state="disabled",
                                      command=self._stop)
        self.stop_btn.grid(row=0, column=2, padx=(6, 0), sticky="ew")

    def _draw_drop_zone(self):
        canvas = self.drop_canvas
        canvas.delete("all")
        w, h = canvas.winfo_width(), canvas.winfo_height()
        canvas.create_rectangle(3, 3, w - 4, h - 4, outline="#9aa1a7", dash=(7, 5), width=2)
        canvas.create_text(w / 2, h / 2 - 14, text="📥  TRASCINA QUI IL TUO FILE  📥", fill="white",
                           font=("Segoe UI", 20, "bold"))
        canvas.create_text(w / 2, h / 2 + 20, text="(.txt, .csv, .xls, .xlsx)", fill="#c9d1d6",
                           font=("Segoe UI", 12))

    # -- input --
    def _on_drop(self, event):
        paths = self.tk.splitlist(event.data)
        if paths:
            self._load_input(paths[0])

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Seleziona il file di input",
            filetypes=[("File supportati", "*.txt *.csv *.xls *.xlsx"), ("Tutti i file", "*.*")])
        if path:
            self._load_input(path)

    def _load_input(self, path):
        try:
            terms = read_search_terms(path)
        except Exception as e:
            messagebox.showerror("Errore di lettura", f"Impossibile leggere il file:\n{e}")
            return
        if not terms:
            messagebox.showwarning("File vuoto", "Nel file non è stato trovato nessun termine di ricerca.")
            return
        self.terms = terms
        self.file_label.configure(text=f"File selezionato: {os.path.basename(path)} ({describe_terms(terms)})")
        self._log(f"File caricato: {path} ({len(terms)} termini)")

    # -- configurazione --
    def _collect_config(self):
        self.config_data.update({
            "USERNAME": self.user_var.get().strip(),
            "PASSWORD": self.pass_var.get(),
            "SHOW_BROWSER": bool(self.show_browser_var.get()),
        })
        return self.config_data

    def _save_settings(self):
        try:
            save_config(self._collect_config())
        except OSError as e:
            messagebox.showerror("Errore", f"Impossibile salvare le impostazioni:\n{e}")
            return
        self._log(f"Impostazioni salvate in {CONFIG_PATH}")

    # -- esecuzione --
    def _start(self):
        if not self.terms:
            messagebox.showwarning("Nessun file", "Trascina o seleziona prima un file di input.")
            return
        config = self._collect_config()
        if not config["USERNAME"] or not config["PASSWORD"]:
            messagebox.showwarning("Credenziali mancanti", "Compila utente e password.")
            return
        try:
            save_config(config)
        except OSError as e:
            self._log(f"⚠️ Impossibile salvare config.json: {e}")

        self.progress_bar.set(0)
        self.progress_label.configure(text="0%")
        self._set_status("⏳ Esportazione in corso...", STATUS_RUNNING)
        self.start_btn.configure(state="disabled")
        self.clear_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")

        self.scraper = Scraper(dict(config), list(self.terms),
                               log=lambda text: self.messages.put(("log", text)),
                               progress=lambda done, total: self.messages.put(("progress", done, total)))
        self.worker = threading.Thread(target=self._run_worker, daemon=True)
        self.worker.start()

    def _run_worker(self):
        try:
            self.scraper.run()
        finally:
            self.messages.put(("done", self.scraper.summary))

    def _open_file(self, path, title, missing):
        if not os.path.exists(path):
            messagebox.showinfo(title, missing)
            return
        try:
            os.startfile(path)
        except OSError as e:
            messagebox.showerror("Errore", f"Impossibile aprire il file:\n{e}")

    def _open_log_file(self):
        """Apre cribis_export.log (storico di tutte le sessioni) con il programma predefinito."""
        self._open_file(LOG_PATH, "Log", "Non c'è ancora nessun file di log.")

    def _open_output_file(self):
        """Apre il file Excel di output con il programma predefinito."""
        self._open_file(output_path_for(self.config_data), "File Excel",
                        "Il file Excel non esiste ancora: verrà creato alla prima esportazione.")

    def _clear(self):
        """Azzera log, avanzamento e file di input caricato (le credenziali restano)."""
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        self.progress_bar.set(0)
        self.progress_label.configure(text="0%")
        self.terms = []
        self.file_label.configure(text="Nessun file selezionato")
        self._set_status("In attesa", STATUS_IDLE)

    def _stop(self):
        if self.scraper:
            self._log("Interruzione in corso...")
            self._set_status("⏳ Interruzione in corso...", STATUS_STOPPED)
            self.stop_btn.configure(state="disabled")
            self.scraper.request_stop()

    def _on_close(self):
        if self.scraper and self.worker and self.worker.is_alive():
            if not messagebox.askyesno("Esportazione in corso", "L'esportazione è in corso. Interrompere e uscire?"):
                return
            self.scraper.request_stop()
            self.worker.join(timeout=8)
        self.destroy()

    # -- log e coda messaggi (thread GUI) --
    def _log(self, text):
        line = f"[{datetime.now():%H:%M:%S}] {text}"
        self.log_box.configure(state="normal")
        self.log_box.insert("end", line + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        try:
            with open(LOG_PATH, "a", encoding="utf-8") as f:
                f.write(f"{datetime.now():%Y-%m-%d} {line}\n")
        except OSError:
            pass

    def _set_status(self, text, color):
        self.status_label.configure(text=text, text_color=color)

    def _show_result(self, summary):
        """Esito dell'esportazione, scritto nel programma (nessuna finestra di dialogo)."""
        if not summary:
            return
        counts = f"{summary['done']}/{summary['total']} ricerche, {summary['found']} risultati trovati"
        if summary["status"] == "completed":
            self._set_status(f"✅ Esportazione completata — {counts}", STATUS_OK)
        elif summary["status"] == "stopped":
            self._set_status(f"⏹ Esportazione interrotta — {counts}", STATUS_STOPPED)
        else:
            self._set_status(f"❌ Esportazione non riuscita: {summary['error']}", STATUS_ERROR)

    def _poll_messages(self):
        try:
            while True:
                kind, *payload = self.messages.get_nowait()
                if kind == "log":
                    self._log(payload[0])
                elif kind == "progress":
                    done, total = payload
                    self.progress_bar.set(done / total if total else 0)
                    self.progress_label.configure(text=f"{int(100 * done / total) if total else 0}%")
                elif kind == "done":
                    self.start_btn.configure(state="normal")
                    self.clear_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
                    self._show_result(payload[0])
        except queue.Empty:
            pass
        self.after(self.POLL_MS, self._poll_messages)


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
