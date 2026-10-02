# CRIBIS Export

Applicazione desktop per Windows che, partendo da un elenco di P.IVA / codici fiscali, interroga il portale CRIBIS X ed esporta i dati anagrafici delle aziende in un file Excel (.xls).

## Download

Dalla pagina **Releases** del repository GitHub si scarica l'ultima versione: il file `CRIBIS_Export_Setup_<versione>.exe` è l'installer (consigliato), mentre `cribis_export.exe` è la versione portatile. Non serve installare nulla: salvarlo in una cartella a piacere (es. `C:\CRIBIS`) e fare doppio clic. I file di configurazione, il log e il risultato vengono creati nella stessa cartella.
Alla prima apertura Windows SmartScreen potrebbe avvisare che l'app non è riconosciuta: cliccare *Ulteriori informazioni* → *Esegui comunque*.

## Uso (utente finale)

1. Avviare `cribis_export.exe`. È necessario avere **Google Chrome** installato e una connessione a Internet (il driver di Chrome viene scaricato automaticamente).
2. Inserire **Utente** e **Password** e cliccare su **Salva Credenziali**. Le impostazioni vengono salvate in `config.json`, accanto all'eseguibile.
3. Trascinare nell'area tratteggiata il file con i termini di ricerca (oppure cliccarla per sceglierlo). Formati supportati:
   - `.txt`: una riga = un termine di ricerca;
   - `.csv`, `.xls`, `.xlsx`: viene letta la prima colonna che contiene dati. Un'eventuale intestazione (prima riga senza cifre, ad es. "Partita IVA") viene ignorata.
4. Cliccare **Avvia** (**Pulisci** azzera log, avanzamento e file caricato). Il pulsante **Apri file Excel** apre il risultato; **Apri file di log** mostra lo storico completo di tutte le sessioni (`cribis_export.log`), anche dopo aver usato Pulisci. Si apre Chrome ed effettua le ricerche una alla volta; la barra di avanzamento e il riquadro dei log mostrano cosa sta succedendo. Non chiudere la finestra di Chrome durante l'esecuzione.
   L'interruttore **Mostra il browser durante l'esecuzione** permette di lavorare con Chrome visibile (predefinito) o nascosto, in background.
   Alla fine, sotto la barra di avanzamento, compare l'esito (completata, interrotta o con errore) con il riepilogo.
5. **Interrompi** ferma l'operazione e chiude il browser: le righe già elaborate restano nel file Excel.

### File generati

Tutti i file vengono creati nella stessa cartella dell'eseguibile:

| File | Contenuto |
| --- | --- |
| `config.json` | Credenziali e impostazioni (non condividerlo) |
| `cribis_dati_export.xls` | Risultati (nome modificabile in `config.json`, `OUTPUT_FILE`) |
| `cribis_export.log` | Registro delle operazioni |
| `user_agents.txt` | Opzionale: elenco di user-agent, uno per riga |

Il file Excel contiene una riga per ogni termine, con la prima colonna **Termine di Ricerca** che riporta l'input a cui corrisponde; se nessun risultato viene trovato, gli altri campi valgono `N/D`. Il file viene aggiornato a ogni ricerca, ma va **chiuso in Excel prima di avviare** l'esportazione, altrimenti non può essere scritto. Il formato `.xls` ha un limite di 65.535 righe.

### Impostazioni avanzate

Modificabili direttamente in `config.json`: `LOGIN_URL` (indirizzo di login), `MIN_DELAY` / `MAX_DELAY` (pausa casuale tra le ricerche, in secondi), `LONG_BREAK_EVERY` / `LONG_BREAK_SECONDS` (pausa lunga periodica), `DRIVER_RESTART_EVERY` (riavvio del browser con nuovo user-agent), `PROXY`. Questi parametri servono a ridurre il rischio di blocchi da parte del portale.

## Sviluppo e build

Requisiti: Python 3.10+.

```powershell
pip install -r requirements.txt
python cribis_export.py          # avvio da sorgente
pyinstaller cribis_export.spec   # genera dist\cribis_export.exe (senza console)
```

### Installer (Inno Setup)

Dopo aver creato l'exe, l'installer Windows si genera con [Inno Setup 6](https://jrsoftware.org/isinfo.php):

```powershell
& "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe" installer\cribis_export.iss
# versione diversa: ... /DAppVersion=1.0.1 installer\cribis_export.iss
```

Il risultato è `installer\Output\CRIBIS_Export_Setup_<versione>.exe`. L'installazione è **per utente** (senza permessi di amministratore) in `%LOCALAPPDATA%\Programs\CRIBIS Export`, perché l'app scrive configurazione, log e risultati accanto all'eseguibile. Installa anche due file di esempio, `config.json` (senza credenziali) e `dati.txt` (un elenco con una P.IVA), senza mai sovrascrivere quelli già presenti: sono in `installer\esempio\` (il config si chiama `config.template.json` per non finire in `.gitignore`). Crea il collegamento nel menu Start (e sul desktop, a scelta) e la voce di disinstallazione, che rimuove anche `config.json` e il log ma lascia il file Excel dei risultati.

L'eseguibile creato è autonomo. `build/`, `dist/`, `config.json` e i file generati sono esclusi da git. Per pubblicare una nuova versione:

```powershell
gh release create v1.0.1 dist\cribis_export.exe --title "v1.0.1" --notes "Descrizione delle modifiche"
```

## Note sull'architettura

- Il lavoro con Selenium gira in un thread separato (`Scraper`); la GUI comunica con esso tramite una coda di messaggi, così non si blocca.
- "Interrompi" e la chiusura della finestra impostano un evento di stop e chiudono il browser, liberando anche le attese in corso.
- Gli XPath di `extract_data()` sono posizionali: se CRIBIS cambia il layout, i campi diventano `N/D`.
