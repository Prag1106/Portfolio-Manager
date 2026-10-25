#!/usr/bin/env python3
"""Bot aggiornamento prezzi: apre 'Stock + ETF Report.xlsx' in Excel, lascia
aggiornare i dati Refinitiv/STOCKHISTORY, salva, sincronizza il DB locale e
(se configurato il token) carica il file su PythonAnywhere e sincronizza il
DB del sito live.

Uso:
    python3 refresh_prices.py               # tutto (Excel + sync locale + upload)
    python3 refresh_prices.py --no-upload   # salta l'upload su PythonAnywhere
    python3 refresh_prices.py --no-excel    # salta il refresh Excel (solo sync)
    python3 refresh_prices.py --scheduled   # launchd: esegue solo se l'ultimo
                                            # slot (giorno feriale 17:45) non è
                                            # ancora stato aggiornato

Token PythonAnywhere (Account -> API Token) va salvato in:
    ~/.pythonanywhere_token
"""
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime, timedelta
from urllib.parse import quote

BASE = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE, 'Stock + ETF Report.xlsx')
EXCEL_NAME = os.path.basename(EXCEL_PATH)
TOKEN_PATH = os.path.expanduser('~/.pythonanywhere_token')
PA_USER = 'Pragyesh'
PA_SITE = 'https://pragyesh.pythonanywhere.com'
PA_REMOTE_PATH = f'/home/{PA_USER}/Portfolio/{EXCEL_NAME}'
# Colonne Price @ Date + YTD % di entrambi i fogli: finché contengono errori
# (#BUSY!, #GETTING_DATA, ...) il refresh non è finito.
POLL_RANGES = [('Stock Report', 'F2:G200'), ('ETF Report', 'F2:G200')]
REFRESH_TIMEOUT_S = 300
# Slot giornaliero (lun-ven) e file con l'ora dell'ultimo aggiornamento riuscito.
# Il Mac a quell'ora è spesso chiuso/in stop: launchd riprova ogni 15 min con
# --scheduled e il bot recupera lo slot perso alla prima occasione utile.
SLOT_HOUR, SLOT_MINUTE = 17, 45
STATE_DIR = os.path.expanduser('~/Library/Application Support/portfolio-refresh')
STAMP_PATH = os.path.join(STATE_DIR, 'last_success')
FAIL_PATH = os.path.join(STATE_DIR, 'last_failure')
LOCK_PATH = os.path.join(STATE_DIR, 'lock')
# Richieste dalla pagina (app.py /api/prices/request): file locale + sito online.
REQ_PATH = os.path.join(BASE, '.price_refresh_request')
MIN_GAP = timedelta(minutes=15)       # niente refresh su richiesta se l'ultimo è più recente
RETRY_AFTER = timedelta(minutes=15)   # dopo un errore, attesa prima di riprovare
POLL_INTERVAL_S = 10


def log(msg):
    print(f'[{datetime.now().strftime("%H:%M:%S")}] {msg}', flush=True)


# Errori AppleScript transitori: capitano quando il Mac ha lo schermo spento e
# System Events/Excel rispondono lentamente (l'AppleEvent scade prima che il
# processo sia pronto). Non sono guasti veri: basta riprovare.
_TRANSIENT_OSA = ('-1712', 'timed out', '-600', '-609', '-1728', 'not running')


def osa(script, retries=3, backoff_s=5):
    last_err = ''
    for attempt in range(retries):
        r = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
        if r.returncode == 0:
            return r.stdout.strip()
        last_err = r.stderr.strip()
        if attempt < retries - 1 and any(t in last_err for t in _TRANSIENT_OSA):
            log(f'AppleScript transitorio ({last_err}); riprovo tra {backoff_s}s '
                f'(tentativo {attempt + 2}/{retries})...')
            time.sleep(backoff_s)
            continue
        break
    raise RuntimeError(f'AppleScript error: {last_err}')


def excel_was_running():
    return osa('tell application "System Events" to (name of processes) '
               'contains "Microsoft Excel"') == 'true'


def open_and_refresh():
    osa(f'''
    tell application "Microsoft Excel"
        set display alerts to false
        open (POSIX file "{EXCEL_PATH}")
        try
            refresh all workbook "{EXCEL_NAME}"
        end try
        try
            calculate
        end try
    end tell''')


def scan_cells():
    """(celle in caricamento, celle in errore) nelle colonne prezzo/YTD.

    In caricamento (#BUSY!, #GETTING_DATA) = bisogna aspettare.
    In errore (#N/A ecc.) = permanente, capita anche aprendo Excel a mano
    (es. YTD di Palantir/SIX): si segnala ma non si blocca il salvataggio.
    """
    loading, errors = 0, 0
    for sheet, rng in POLL_RANGES:
        out = osa(f'''
        tell application "Microsoft Excel"
            get string value of range "{rng}" of worksheet "{sheet}" ¬
                of workbook "{EXCEL_NAME}"
        end tell''')
        for cell in out.split(', '):
            c = cell.strip().upper()
            if 'BUSY' in c or 'GETTING' in c or 'RETRIEVING' in c:
                loading += 1
            elif '#' in c:
                errors += 1
    return loading, errors


def save_and_close(save=True, quit_excel=False):
    saving = 'yes' if save else 'no'
    osa(f'''
    tell application "Microsoft Excel"
        set display alerts to false
        try
            close workbook "{EXCEL_NAME}" saving {saving}
        end try
        set display alerts to true
    end tell''')
    if quit_excel:
        osa('tell application "Microsoft Excel" to quit saving no')


def refresh_excel():
    if not os.path.exists(EXCEL_PATH):
        log(f'ERRORE: {EXCEL_PATH} non trovato'); sys.exit(1)

    # Backup best-effort: su Mac addormentato copy2 può dare EDEADLK (errno 11,
    # "Resource deadlock avoided"). Il .bak è solo una rete di sicurezza — la
    # vera protezione è close-senza-salvare al timeout — quindi un backup
    # fallito non deve far abortire l'intero refresh.
    # Il .bak vecchio viene spesso scaricato da iCloud ("dataless"): sovrascriverlo
    # dà EDEADLK, quindi lo si elimina prima di copiare.
    for attempt in range(3):
        try:
            try:
                os.remove(EXCEL_PATH + '.bak')
            except FileNotFoundError:
                pass
            shutil.copy2(EXCEL_PATH, EXCEL_PATH + '.bak')
            break
        except OSError as e:
            log(f'Backup .bak fallito ({e}); '
                f'{"riprovo tra 5s" if attempt < 2 else "proseguo senza backup"}...')
            if attempt < 2:
                time.sleep(5)

    was_running = excel_was_running()
    log(f'Apro "{EXCEL_NAME}" in Excel (Excel già aperto: {"sì" if was_running else "no"})...')
    open_and_refresh()

    deadline = time.time() + REFRESH_TIMEOUT_S
    loading, errors = scan_cells()
    while loading and time.time() < deadline:
        log(f'In aggiornamento... {loading} celle in caricamento')
        time.sleep(POLL_INTERVAL_S)
        loading, errors = scan_cells()

    if loading:
        # Non salvare: meglio tenere gli ultimi valori buoni nel file
        log(f'ERRORE: dopo {REFRESH_TIMEOUT_S}s ci sono ancora {loading} celle in '
            'caricamento (Excel è connesso a internet e loggato in Microsoft 365?). '
            'Chiudo SENZA salvare.')
        save_and_close(save=False, quit_excel=not was_running)
        sys.exit(1)

    if errors:
        log(f'Nota: {errors} celle con errore permanente (succede anche aprendo '
            'Excel a mano, es. YTD Palantir su SIX). Salvo comunque.')
    log('Aggiornamento completato. Salvo e chiudo.')
    save_and_close(save=True, quit_excel=not was_running)


def last_due_slot(now=None):
    """Ultimo slot feriale 17:45 già passato (None se nessuno, impossibile in pratica)."""
    now = now or datetime.now()
    d = now.replace(hour=SLOT_HOUR, minute=SLOT_MINUTE, second=0, microsecond=0)
    if d > now:
        d -= timedelta(days=1)
    while d.weekday() >= 5:  # sab/dom
        d -= timedelta(days=1)
    return d


def last_success():
    try:
        return datetime.fromtimestamp(float(open(STAMP_PATH).read().strip()))
    except (OSError, ValueError):
        return None


def mark_success():
    os.makedirs(os.path.dirname(STAMP_PATH), exist_ok=True)
    with open(STAMP_PATH, 'w') as f:
        f.write(str(time.time()))


def lid_closed():
    """Coperchio chiuso = Mac in stop o DarkWake: Excel non è pilotabile."""
    r = subprocess.run(['ioreg', '-r', '-k', 'AppleClamshellState', '-d', '4'],
                       capture_output=True, text=True)
    return '"AppleClamshellState" = Yes' in r.stdout


def sync_local_db():
    sys.path.insert(0, BASE)
    from app import app  # noqa: E402
    with app.test_client() as client:
        for ep in ('/api/stock-universe/sync', '/api/etf-universe/sync'):
            r = client.post(ep)
            log(f'Sync locale {ep}: {r.status_code} {r.get_json()}')


def upload_and_sync_remote():
    if not os.path.exists(TOKEN_PATH):
        log(f'Upload saltato: nessun token in {TOKEN_PATH} '
            '(Account -> API Token su pythonanywhere.com)')
        return
    import requests
    token = open(TOKEN_PATH).read().strip()
    url = (f'https://www.pythonanywhere.com/api/v0/user/{PA_USER}/files/path'
           f'{quote(PA_REMOTE_PATH)}')
    with open(EXCEL_PATH, 'rb') as f:
        r = requests.post(url, headers={'Authorization': f'Token {token}'},
                          files={'content': f}, timeout=120)
    if r.status_code not in (200, 201):
        log(f'ERRORE upload su PythonAnywhere: {r.status_code} {r.text[:200]}')
        sys.exit(1)
    log(f'Excel caricato su PythonAnywhere ({r.status_code}).')

    for ep in ('/api/stock-universe/sync', '/api/etf-universe/sync'):
        r = requests.post(PA_SITE + ep, timeout=120)
        log(f'Sync remoto {ep}: {r.status_code} {r.text[:120]}')


def _read_ts(path):
    try:
        return datetime.fromtimestamp(float(open(path).read().strip()))
    except (OSError, ValueError):
        return None


def mark_failure():
    os.makedirs(STATE_DIR, exist_ok=True)
    with open(FAIL_PATH, 'w') as f:
        f.write(str(time.time()))


def page_request():
    """Richiesta più recente fatta aprendo la pagina (in locale o sul sito online)."""
    times = []
    t = _read_ts(REQ_PATH)
    if t:
        times.append(t)
    try:
        import requests
        t = requests.get(PA_SITE + '/api/prices/status', timeout=10).json().get('requested_at')
        if t:
            times.append(datetime.fromtimestamp(float(t)))
    except Exception:
        pass  # sito irraggiungibile: conta solo lo slot/la richiesta locale
    return max(times) if times else None


def due_reason():
    """Motivo per cui serve un refresh, o None."""
    now = datetime.now()
    done = last_success()
    slot = last_due_slot(now)
    if not done:
        return 'nessun aggiornamento registrato'
    if done < slot:
        return f'slot del {slot:%d/%m %H:%M} da aggiornare (ultimo riuscito: {done:%d/%m %H:%M})'
    req = page_request()
    if req and req > done and now - done >= MIN_GAP:
        return f'richiesta dalla pagina delle {req:%H:%M} (ultimo riuscito: {done:%d/%m %H:%M})'
    return None


def acquire_lock():
    """Un solo bot alla volta (launchd + pagina possono partire insieme)."""
    import fcntl
    os.makedirs(STATE_DIR, exist_ok=True)
    fh = open(LOCK_PATH, 'w')
    try:
        fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        return None
    return fh  # tenere il riferimento: il lock vale finché il processo è vivo


def run(args):
    if '--no-excel' not in args:
        refresh_excel()
    sync_local_db()
    if '--no-upload' not in args:
        upload_and_sync_remote()
    mark_success()
    log('Fatto.')


if __name__ == '__main__':
    args = sys.argv[1:]
    scheduled = '--scheduled' in args
    lock = acquire_lock()
    if not lock:
        if not scheduled:
            log('Un aggiornamento è già in corso: esco.')
        sys.exit(0)
    if scheduled:
        # Silenzioso finché non c'è niente da fare (gira ogni 5 min).
        if lid_closed():
            sys.exit(0)  # Mac chiuso/in stop: Excel non è pilotabile, riprova dopo
        failed = _read_ts(FAIL_PATH)
        if failed and (last_success() or datetime.min) < failed and datetime.now() - failed < RETRY_AFTER:
            sys.exit(0)  # errore recente: aspetta prima di riprovare
        reason = due_reason()
        if not reason:
            sys.exit(0)
        log(f'Avvio: {reason}.')
    try:
        run(args)
    except SystemExit as e:
        if e.code not in (0, None):
            mark_failure()
            log(f'Aggiornamento fallito: riprovo automaticamente tra {int(RETRY_AFTER.total_seconds() // 60)} min.')
        raise
    except Exception as e:
        mark_failure()
        import traceback
        traceback.print_exc()
        log(f'Aggiornamento fallito ({e}): riprovo automaticamente tra {int(RETRY_AFTER.total_seconds() // 60)} min.')
        sys.exit(1)
