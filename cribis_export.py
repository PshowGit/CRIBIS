import csv
import os
import json
import time
import random
import traceback
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager


# --- Lettura configurazione ---
CONFIG_FILE = "config.json"
try:
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)
except FileNotFoundError:
    print(f"❌ File di configurazione '{CONFIG_FILE}' non trovato.")
    exit(1)

LOGIN_URL = config.get("LOGIN_URL", "")
USERNAME = config.get("USERNAME", "")
PASSWORD = config.get("PASSWORD", "")
INPUT_FILE = config.get("INPUT_FILE", "lettura.txt")
OUTPUT_FILE = config.get("OUTPUT_FILE", "dati_export.csv")
CHROME_DRIVER_PATH = config.get("CHROME_DRIVER_PATH", "chromedriver.exe")
PROXY = config.get("PROXY", "")
MIN_DELAY = config.get("MIN_DELAY", 4)
MAX_DELAY = config.get("MAX_DELAY", 9)
LONG_BREAK_EVERY = config.get("LONG_BREAK_EVERY", 100)
LONG_BREAK_SECONDS = config.get("LONG_BREAK_SECONDS", 180)
USER_AGENTS_FILE = config.get("USER_AGENTS_FILE", "user_agents.txt")
DRIVER_RESTART_EVERY = config.get("DRIVER_RESTART_EVERY", 25)

# --- Legge lista user-agent da file ---
try:
    with open(USER_AGENTS_FILE, "r", encoding="utf-8") as f:
        USER_AGENTS = [line.strip() for line in f if line.strip()]
    if not USER_AGENTS:
        raise ValueError("File user_agents.txt vuoto")
except Exception as e:
    print(f"❌ Errore nel leggere '{USER_AGENTS_FILE}': {e}")
    USER_AGENTS = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36"
    ]

# --- Funzioni di supporto ---
def random_delay():
    time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))

def long_break_if_needed(count):
    if LONG_BREAK_EVERY and count % LONG_BREAK_EVERY == 0 and count != 0:
        print(f"⏸ Pausa lunga di {LONG_BREAK_SECONDS}s ogni {LONG_BREAK_EVERY} richieste.")
        time.sleep(LONG_BREAK_SECONDS)

def setup_driver(user_agent=None):
    #service = Service(CHROME_DRIVER_PATH)
    # Creazione automatica del driver (senza dover scaricare manualmente il ChromeDriver)
    service = Service(ChromeDriverManager().install())
    
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

    if user_agent:
        options.add_argument(f"--user-agent={user_agent}")
    if PROXY:
        options.add_argument(f'--proxy-server={PROXY}')


    driver = webdriver.Chrome(service=service, options=options)
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

def login(driver):
    print("Tentativo di login...")
    driver.get(LOGIN_URL)
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "UserName")))
        driver.find_element(By.NAME, "UserName").send_keys(USERNAME)
        driver.find_element(By.NAME, "Password").send_keys(PASSWORD)
        driver.find_element(By.NAME, "submit_pwd").click()
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "search-input")))
        print("✅ Login effettuato con successo.")
        return True
    except TimeoutException:
        print("❌ Timeout durante il login. Controllare credenziali o URL.")
        return False
    except Exception as e:
        print(f"❌ Errore imprevisto durante il login: {e}")
        return False

def read_search_terms(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            terms = [line.strip() for line in f if line.strip()]
        print(f"Lettura di {len(terms)} termini di ricerca da '{file_path}'.")
        return terms
    except FileNotFoundError:
        print(f"❌ File di input '{file_path}' non trovato.")
        return []

def safe_find(driver, xpath):
    try:
        return driver.find_element(By.XPATH, xpath).text.strip()
    except (NoSuchElementException, TimeoutException):
        return "N/D"

def extract_data(driver):
    return {
        "nome_azienda": safe_find(driver, "//*[@id='anagReport']/div[1]/div[1]/h1"),
        "sede": safe_find(driver, "(//span[contains(@class,'upcas') and contains(@class,'link-silver-small')]/following-sibling::p)[1]"),
        "codice_fiscale": safe_find(driver, "(//span[contains(@class,'upcas') and contains(@class,'link-silver-small')]/following-sibling::p)[2]"),
        "iva": safe_find(driver, "(//span[contains(@class,'upcas') and contains(@class,'link-silver-small')]/following-sibling::p)[3]"),
        "cciaa_rea": safe_find(driver, "(//span[contains(@class,'upcas') and contains(@class,'link-silver-small')]/following-sibling::p)[4]"),
        "natura_giuridica": safe_find(driver, "(//span[contains(@class,'upcas') and contains(@class,'link-silver-small')]/following-sibling::p)[5]"),
        "attività": safe_find(driver, "(//span[contains(@class,'upcas') and contains(@class,'link-silver-small')]/following-sibling::p)[6]"),
        "attività_dichiarata": safe_find(driver, "(//span[contains(@class,'upcas') and contains(@class,'link-silver-small')]/following-sibling::p)[7]")
    }

# --- Main ---
def main():
    total_searches = 0
    results_found = 0
    csv_file_path = os.path.join(os.getcwd(), OUTPUT_FILE)

    search_terms = read_search_terms(INPUT_FILE)
    if not search_terms:
        print("Nessun termine di ricerca trovato. Uscita.")
        return

    # apre driver iniziale
    ua = random.choice(USER_AGENTS)
    driver = setup_driver(user_agent=ua)
    if not login(driver):
        print("❌ Login fallito. Uscita.")
        driver.quit()
        return

    with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = [
            "nome_azienda", "sede", "codice_fiscale", "iva",
            "cciaa_rea", "natura_giuridica", "attività", "attività_dichiarata"
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()

        for idx, term in enumerate(search_terms, start=1):
            total_searches += 1
            print(f"\n[{idx}/{len(search_terms)}] 🔍 Ricerca per: '{term}'")

            # Riavvio driver ogni DRIVER_RESTART_EVERY
            if DRIVER_RESTART_EVERY and idx % DRIVER_RESTART_EVERY == 1 and idx != 1:
                print(f"🔄 Riavvio driver e cambio user-agent...")
                try:
                    driver.quit()
                except Exception:
                    pass
                ua = random.choice(USER_AGENTS)
                driver = setup_driver(user_agent=ua)
                if not login(driver):
                    print("❌ Login fallito dopo riavvio driver. Esco.")
                    driver.quit()
                    return

            max_retries = 3
            for attempt in range(1, max_retries + 1):
                try:
                    search_input = WebDriverWait(driver, 15).until(
                        EC.visibility_of_element_located((By.ID, "search-input"))
                    )
                    search_input.clear()
                    search_input.send_keys(term)

                    search_button = WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.ID, "search-button"))
                    )
                    search_button.click()

                    link_result = WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[@id='root']//h2/a"))
                    )
                    link_result.click()

                    try:
                        collapse_icon = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//*[@id='company-details-col-left']//p/i"))
                        )
                        collapse_icon.click()
                    except (TimeoutException, NoSuchElementException):
                        pass

                    data = extract_data(driver)
                    if data["nome_azienda"] == "N/D":
                        print(f"⚠️ Nessun risultato trovato per '{term}'")
                    else:
                        results_found += 1
                        print(f"✅ Dati salvati per '{term}': {data['nome_azienda']}")
                    writer.writerow(data)

                    random_delay()
                    long_break_if_needed(idx)
                    break

                except (TimeoutException, NoSuchElementException, WebDriverException) as e:
                    wait = 2 ** attempt + random.uniform(0.5, 1.5)
                    print(f"⚠️ Tentativo {attempt}/{max_retries} fallito per '{term}': {e}. Attesa {wait:.1f}s.")
                    time.sleep(wait)
                except Exception as e:
                    print(f"❌ Errore imprevisto per '{term}': {e}")
                    traceback.print_exc()
                    writer.writerow({fn: "N/D" for fn in fieldnames})
                    break

    driver.quit()
    print("\n--- Riepilogo ---")
    print(f"🔹 Totale ricerche tentate: {total_searches}")
    print(f"🔹 Risultati trovati: {results_found}")
    print(f"📁 File CSV generato: {csv_file_path}")

if __name__ == "__main__":
    main()



