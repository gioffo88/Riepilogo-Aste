# Aste Beni Mobili MO-BO-RE — automazione GitHub Actions

Replica automatica, su GitHub, della ricerca di aste giudiziarie di beni mobili
(veicoli, arte/oreficeria/orologeria/antiquariato) nelle province di Modena, Bologna
e Reggio Emilia andate deserte almeno 2 volte e ancora in vendita. Il risultato viene
salvato come file Excel su una cartella Google Drive ad ogni esecuzione.

## Cosa fa

1. Apre `ivgmodena.it` con un browser headless (Playwright) per catturare la chiave
   pubblica dell'API di ricerca (Typesense) e il token usato per l'API storico-inserzione
   — entrambi vengono ruotati periodicamente dal sito, per questo servono catturati "a caldo".
2. Interroga l'API di ricerca per tutti i lotti Veicoli / Arte-Oreficeria in vendita
   nelle 3 province.
3. Per ciascun lotto scarica lo storico delle aste precedenti e tiene solo quelli con
   almeno 2 esiti "Non aggiudicato".
4. Visita la scheda di ogni lotto idoneo per leggere cauzione e link ai documenti
   (perizia, avviso, ordinanza di vendita).
5. Confronta l'elenco con l'ultimo report trovato nella cartella Google Drive
   (stesso nome file, prefisso `Report_Aste_Beni_Mobili_MO_BO_RE`) per segnalare
   nuovi lotti e lotti scomparsi.
6. Genera il file Excel (`Riepilogo`, `Veicoli`, `Arte-Oreficeria-Antiquariato`) e lo
   carica sia su Google Drive sia come artifact della run di GitHub Actions (backup
   per 90 giorni anche se Drive non e' configurato).

## Cosa NON fa automaticamente

La colonna "Valore stimato da te" resta vuota: è una valutazione qualitativa (marca,
modello, anno, stato) che questa routine non prova a indovinare da sola, per non
scrivere numeri inventati in un foglio che sembra autorevole. Va compilata a mano o
collegata a un servizio di stima dedicato.

## Come attivarlo

### 1. Crea il repository

Crea un repository GitHub (anche privato) e carica dentro tutta questa cartella
mantenendo la struttura:

```
.github/workflows/aste_beni_mobili.yml
scripts/aste_scraper.py
requirements.txt
README.md
```

### 2. Crea un service account Google con accesso a Drive

1. Vai su [Google Cloud Console](https://console.cloud.google.com/) → crea (o riusa)
   un progetto.
2. Abilita la **Google Drive API** per quel progetto.
3. Crea un **Service Account** (IAM & Admin → Service Accounts → Create).
4. Genera una chiave JSON per il service account e scaricala.
5. Su Google Drive, crea (o scegli) la cartella dove vuoi salvare i report, e
   **condividila** con l'indirizzo email del service account (es.
   `nome-account@progetto.iam.gserviceaccount.com`), con permesso di modifica.
6. Copia l'ID della cartella dall'URL di Drive:
   `https://drive.google.com/drive/folders/QUESTO_E_L_ID`

### 3. Aggiungi i secret al repository GitHub

Repository → Settings → Secrets and variables → Actions → New repository secret:

- `GDRIVE_SA_KEY_B64`: il contenuto del file JSON del service account, **codificato in
  base64 su una riga sola**. Su Mac/Linux: `base64 -i chiave-service-account.json | tr -d '\n'`
  poi incolla il risultato.
- `GDRIVE_FOLDER_ID`: l'ID della cartella Drive del punto precedente.

### 4. (Opzionale) cambia la frequenza

Nel file `.github/workflows/aste_beni_mobili.yml`, la riga `cron: "0 7 * * 1"` esegue
il job ogni lunedì alle 7:00 UTC. Per cambiarla, usa [crontab.guru](https://crontab.guru).
GitHub Actions esegue i cron con qualche minuto di ritardo variabile: è normale.

### 5. Primo avvio

Vai su GitHub → tab **Actions** → workflow "Aste Beni Mobili MO-BO-RE" → **Run workflow**,
per testarlo subito senza aspettare il cron. Controlla i log dello step "Esegui lo
scraper e carica su Drive": se qualcosa nel sito è cambiato (struttura pagina, nomi dei
campi), il messaggio di errore indica dove intervenire nello script.

## Manutenzione prevista

Il sito Astagiudiziaria.com può cambiare struttura nel tempo (è già successo con la
rotazione della chiave Typesense durante la stesura di questo script). Punti più
fragili, in ordine di probabilità di rottura:

- `DOMAIN_BY_KEYWORD` in `aste_scraper.py`: mappa provincia → dominio del sito per le
  schede di dettaglio. Il dominio di Reggio Emilia non è stato verificato live (il
  sandbox di sviluppo non aveva accesso di rete al circuito Astagiudiziaria): se il job
  fallisce nello scraping dei lotti di Reggio Emilia, aggiorna questo dizionario con il
  dominio corretto.
- La regex della cauzione (`Cauzione[:\s]*€?...`) assume che la scheda annuncio mostri
  la label "Cauzione" seguita dall'importo: se il sito cambia il testo, aggiornare la
  regex in `scrape_lot_page`.
- `FILTER_BY`: se il portale rinomina le categorie (es. cambia il nome esatto di
  "ARTE -OREFICERIA - OROLOGERIA- ANTIQUARIATO"), la query Typesense smette di
  restituire risultati per quella categoria.

## Esecuzione locale (per test/debug)

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium

export GDRIVE_CREDENTIALS_FILE=/percorso/chiave-service-account.json   # opzionale
export GDRIVE_FOLDER_ID=xxxxxxxxxxxxxxxx                                # opzionale
python scripts/aste_scraper.py
```

Senza le due variabili d'ambiente, lo script gira comunque e salva solo il file in
`./output/`, senza toccare Google Drive.
