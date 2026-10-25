from flask import Flask, render_template, request, jsonify, send_file
import sqlite3, json, io, os, math, gzip, functools, re
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.json.sort_keys = False  # preserve insertion order of dict keys in JSON responses
# Static assets carry a ?v=<mtime> cache-buster, so browsers can cache them hard.
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 30 * 24 * 3600
DB = os.path.join(os.path.dirname(__file__), 'portfolio.db')

# ── Excel-endpoint cache (keyed on file mtime → invalidates on upload/save) ────
DARTA_PATH = os.path.join(os.path.dirname(__file__), 'DARTA CHP+CHPV updated.xlsx')
_resp_cache = {}

def cache_by_mtime(path):
    """Cache a JSON GET endpoint's body until the underlying Excel file changes.
    `path` may be a string or a zero-arg callable returning the path."""
    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*a, **kw):
            p = path() if callable(path) else path
            try: mtime = os.path.getmtime(p)
            except OSError: mtime = -1
            hit = _resp_cache.get(fn.__name__)
            if hit and hit[0] == mtime:
                return app.response_class(hit[1], mimetype='application/json')
            resp = fn(*a, **kw)
            if resp.status_code == 200:
                _resp_cache[fn.__name__] = (mtime, resp.get_data())
            return resp
        return wrapper
    return deco

@app.after_request
def gzip_json(resp):
    """PythonAnywhere doesn't compress dynamic responses; large JSON payloads
    (dashboard, fund DB) shrink ~10x with gzip."""
    if (resp.status_code != 200 or resp.direct_passthrough
            or not resp.mimetype == 'application/json'
            or 'Content-Encoding' in resp.headers
            or 'gzip' not in (request.headers.get('Accept-Encoding') or '')):
        return resp
    body = resp.get_data()
    if len(body) < 1024:
        return resp
    resp.set_data(gzip.compress(body, 6))
    resp.headers['Content-Encoding'] = 'gzip'
    resp.headers['Content-Length'] = str(len(resp.get_data()))
    resp.headers.setdefault('Vary', 'Accept-Encoding')
    return resp

def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as db:
        db.execute('''CREATE TABLE IF NOT EXISTS clients (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )''')
        db.execute('''CREATE TABLE IF NOT EXISTS portfolios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id INTEGER NOT NULL,
            type TEXT NOT NULL,
            data TEXT NOT NULL,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (client_id) REFERENCES clients(id) ON DELETE CASCADE
        )''')
        db.execute('''CREATE TABLE IF NOT EXISTS consultants (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )''')
        # Migrate: clients belong to a financial consultant
        try:
            db.execute('ALTER TABLE clients ADD COLUMN consultant_id INTEGER REFERENCES consultants(id)')
        except Exception:
            pass
        if not db.execute('SELECT 1 FROM consultants LIMIT 1').fetchone():
            db.execute("INSERT INTO consultants (name) VALUES ('Nicola Francia')")
        # Any client without a consultant (pre-migration rows) goes to the first one
        first = db.execute('SELECT MIN(id) FROM consultants').fetchone()[0]
        db.execute('UPDATE clients SET consultant_id=? WHERE consultant_id IS NULL', (first,))
        db.commit()

init_db()

STOCK_EXCEL_PATH = os.path.join(os.path.dirname(__file__), 'Stock + ETF Report.xlsx')


def read_report_sheet(sheet_name, key_column):
    """Read a report sheet, auto-detecting the header row.

    Editing the workbook in Excel can leave a stray blank row above the header
    (seen 2026-09-15: a blank row 1 pushed the ETF header to row 2, so pandas
    read all columns as 'Unnamed' and the sync dropped every ETF). Scan the top
    rows for the one containing key_column and use it as the header.
    """
    import pandas as pd
    raw = pd.read_excel(STOCK_EXCEL_PATH, sheet_name=sheet_name, header=None)
    header_row = 0
    for i in range(min(10, len(raw))):
        if key_column in [str(x).strip() for x in raw.iloc[i].tolist()]:
            header_row = i
            break
    return pd.read_excel(STOCK_EXCEL_PATH, sheet_name=sheet_name, header=header_row)


def init_stock_db():
    with get_db() as db:
        db.execute('''CREATE TABLE IF NOT EXISTS stock_universe (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ticker TEXT,
            company TEXT NOT NULL UNIQUE,
            exchange TEXT,
            currency TEXT,
            price REAL,
            price_date TEXT,
            ytd_pct REAL,
            one_y_pct REAL,
            three_y_pct REAL,
            pe REAL,
            ranking INTEGER,
            industry TEXT,
            high_52w REAL,
            low_52w REAL,
            daily_chg_pct REAL,
            market_cap REAL,
            avg_volume REAL,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )''')
        db.commit()

init_stock_db()

@app.route('/api/stock-universe', methods=['GET'])
def get_stock_universe():
    db = get_db()
    rows = db.execute('SELECT * FROM stock_universe ORDER BY company').fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/stock-universe/sync', methods=['POST'])
def sync_stock_universe():
    import pandas as pd
    if not os.path.exists(STOCK_EXCEL_PATH):
        return jsonify({'error': 'Excel file not found'}), 404

    df = read_report_sheet('Stock Report', 'Company')
    now = datetime.now().isoformat()
    count = 0

    def safe(v):
        if v is None: return None
        try:
            f = float(v)
            return None if math.isnan(f) else f
        except: return None

    with get_db() as db:
        for _, row in df.iterrows():
            company = str(row.get('Company', '') or '').strip()
            if not company or company == 'nan': continue
            ticker = str(row.get('Ticker', '') or '').strip()
            if ticker in ('nan', 'None', ''): ticker = ''

            price_date = None
            if pd.notna(row.get('Date')):
                try: price_date = str(row['Date'])[:10]
                except: pass

            ranking = None
            if pd.notna(row.get('Ranking')):
                try: ranking = int(row['Ranking'])
                except: pass

            db.execute('''INSERT INTO stock_universe
                (ticker, company, exchange, currency, price, price_date, ytd_pct, one_y_pct,
                 three_y_pct, pe, ranking, industry, high_52w, low_52w, daily_chg_pct,
                 market_cap, avg_volume, updated_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(company) DO UPDATE SET
                ticker=excluded.ticker, exchange=excluded.exchange, currency=excluded.currency,
                price=excluded.price, price_date=excluded.price_date, ytd_pct=excluded.ytd_pct,
                one_y_pct=excluded.one_y_pct, three_y_pct=excluded.three_y_pct, pe=excluded.pe,
                ranking=excluded.ranking, industry=excluded.industry, high_52w=excluded.high_52w,
                low_52w=excluded.low_52w, daily_chg_pct=excluded.daily_chg_pct,
                market_cap=excluded.market_cap, avg_volume=excluded.avg_volume,
                updated_at=excluded.updated_at
            ''', (
                ticker, company,
                str(row.get('Exchange', '') or ''), str(row.get('Currency', '') or ''),
                safe(row.get('Price @ Date')), price_date,
                safe(row.get('YTD %')), safe(row.get('1Y %')), safe(row.get('3Y %')),
                safe(row.get('P/E')), ranking,
                str(row.get('Industry', '') or ''),
                safe(row.get('52 Day High')), safe(row.get('52 Day Low')),
                safe(row.get('Daily % Change')),
                safe(row.get('Market Cap')), safe(row.get('Average Volume')),
                now
            ))
            count += 1
        db.commit()

    return jsonify({'ok': True, 'synced': count})

@app.route('/api/stock-universe/search', methods=['GET'])
def search_stock_universe():
    q = request.args.get('q', '').strip()
    db = get_db()
    if not q:
        rows = db.execute('SELECT * FROM stock_universe ORDER BY company').fetchall()
    else:
        like = f'%{q}%'
        rows = db.execute(
            'SELECT * FROM stock_universe WHERE UPPER(ticker) LIKE UPPER(?) OR UPPER(company) LIKE UPPER(?)',
            (like, like)
        ).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/stock-universe/<int:sid>', methods=['PUT'])
def update_stock_ticker(sid):
    data = request.json
    db = get_db()
    db.execute('UPDATE stock_universe SET ticker=?, updated_at=? WHERE id=?',
               (data.get('ticker', ''), datetime.now().isoformat(), sid))
    db.commit()
    return jsonify({'ok': True})

@app.route('/api/stock-universe', methods=['POST'])
def add_stock():
    """Add a new ticker to Excel (col B) + DB placeholder. User's Refinitiv fills the rest."""
    import shutil
    from openpyxl import load_workbook
    from openpyxl.worksheet.formula import ArrayFormula

    data = request.json
    ticker = (data.get('ticker') or '').strip()
    if not ticker:
        return jsonify({'error': 'Ticker required'}), 400

    if not os.path.exists(STOCK_EXCEL_PATH):
        return jsonify({'error': 'Excel file not found'}), 404

    shutil.copy2(STOCK_EXCEL_PATH, STOCK_EXCEL_PATH + '.bak')
    wb = load_workbook(STOCK_EXCEL_PATH)
    ws = wb['Stock Report']
    last_row = ws.max_row
    new_row = last_row + 1
    tmpl = 2  # template row to copy formulas from

    # Col B = ticker (user input that Refinitiv reads)
    ws.cell(new_row, 2).value = ticker

    # Copy array formulas from template row for cols A, C-F, J-Q
    for col_idx in [1, 3, 4, 5, 6, 10, 11, 12, 13, 14, 15, 16, 17]:
        src = ws.cell(tmpl, col_idx).value
        dst = ws.cell(new_row, col_idx)
        if hasattr(src, 'text'):
            col_letter = dst.column_letter
            dst.value = ArrayFormula(f'{col_letter}{new_row}', src.text)
        elif src:
            dst.value = src

    # Cols G, H, I have hardcoded B2/E2 refs — replace with new row number
    for col_idx in [7, 8, 9]:
        formula = ws.cell(tmpl, col_idx).value
        if formula:
            new_formula = formula.replace(f'B{tmpl}', f'B{new_row}').replace(f'E{tmpl}', f'E{new_row}')
            ws.cell(new_row, col_idx).value = new_formula

    # Extend table reference
    for tbl in ws.tables.values():
        tbl.ref = f'A1:Q{new_row}'

    wb.save(STOCK_EXCEL_PATH)
    wb.close()

    # Add placeholder to DB
    now = datetime.now().isoformat()
    db = get_db()
    db.execute(
        'INSERT OR IGNORE INTO stock_universe (ticker, company, updated_at) VALUES (?,?,?)',
        (ticker, f'[{ticker}]', now)
    )
    db.commit()
    new_id = db.execute('SELECT id FROM stock_universe WHERE ticker=? AND company=?', (ticker, f'[{ticker}]')).fetchone()
    return jsonify({'ok': True, 'id': new_id['id'] if new_id else None,
                    'message': f'Ticker {ticker} aggiunto. Apri Excel, lascia aggiornare Refinitiv, poi premi Sincronizza.'})

@app.route('/api/stock-universe/<int:sid>', methods=['DELETE'])
def delete_stock(sid):
    """Remove stock from DB and delete its row from Excel."""
    import shutil, pandas as pd
    from openpyxl import load_workbook

    db = get_db()
    stock = db.execute('SELECT * FROM stock_universe WHERE id=?', (sid,)).fetchone()
    if not stock:
        return jsonify({'error': 'Not found'}), 404

    company = stock['company']

    if os.path.exists(STOCK_EXCEL_PATH):
        shutil.copy2(STOCK_EXCEL_PATH, STOCK_EXCEL_PATH + '.bak')
        # Find the Excel row by reading cached values (data_only)
        df = pd.read_excel(STOCK_EXCEL_PATH, sheet_name='Stock Report')
        excel_row = None
        for i, row in df.iterrows():
            if str(row.get('Company', '') or '').strip() == company:
                excel_row = i + 2  # +2 because header is row 1, df is 0-indexed
                break

        if excel_row:
            wb = load_workbook(STOCK_EXCEL_PATH)
            ws = wb['Stock Report']
            ws.delete_rows(excel_row)
            # Shrink table reference
            new_last = ws.max_row
            for tbl in ws.tables.values():
                tbl.ref = f'A1:Q{new_last}'
            wb.save(STOCK_EXCEL_PATH)
            wb.close()

    db.execute('DELETE FROM stock_universe WHERE id=?', (sid,))
    db.commit()
    return jsonify({'ok': True})

# ── ETF Universe endpoints ─────────────────────────────────────────────────────

def init_etf_db():
    with get_db() as db:
        db.execute('''CREATE TABLE IF NOT EXISTS etf_universe (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            etf_id TEXT NOT NULL UNIQUE,
            ticker TEXT,
            etf_name TEXT,
            exchange TEXT,
            currency TEXT,
            price REAL,
            price_date TEXT,
            ytd_pct REAL,
            one_y_pct REAL,
            three_y_pct REAL,
            mkt_cap REAL,
            type TEXT,
            sector TEXT,
            high_52w REAL,
            low_52w REAL,
            daily_chg_pct REAL,
            volume REAL,
            ter REAL,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )''')
        # Migrate: add ter column if it doesn't exist yet
        try:
            db.execute('ALTER TABLE etf_universe ADD COLUMN ter REAL')
        except Exception:
            pass
        db.commit()

init_etf_db()

@app.route('/api/etf-universe', methods=['GET'])
def get_etf_universe():
    db = get_db()
    rows = db.execute('SELECT * FROM etf_universe ORDER BY ticker').fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/etf-universe/sync', methods=['POST'])
def sync_etf_universe():
    import pandas as pd
    if not os.path.exists(STOCK_EXCEL_PATH):
        return jsonify({'error': 'Excel file not found'}), 404

    df = read_report_sheet('ETF Report', 'Ticker')
    now = datetime.now().isoformat()
    count = 0

    def safe(v):
        if v is None: return None
        try:
            f = float(v)
            return None if math.isnan(f) else f
        except: return None

    def safe_pct(v):
        """Parse a percentage value that may be a string like '0.55%' or a decimal like 0.0055."""
        if v is None: return None
        try:
            s = str(v).strip()
            if s in ('', 'nan', '#VALUE!'): return None
            if s.endswith('%'):
                return float(s[:-1]) / 100
            f = float(s)
            if math.isnan(f): return None
            # If value looks like a whole percentage (e.g. 0.55 meaning 0.55%), store as decimal
            return f if f < 1 else f / 100
        except: return None

    # Detect TER/Expense Ratio column — try exact match first, then case-insensitive
    ter_col = None
    for candidate in ['Expense ratio', 'Expense Ratio', 'expense ratio', 'TER', 'ter', 'Total Expense Ratio']:
        if candidate in df.columns:
            ter_col = candidate
            break
    if ter_col is None:
        # Case-insensitive fallback
        for col in df.columns:
            if col.strip().lower() in ('expense ratio', 'ter', 'total expense ratio'):
                ter_col = col
                break

    rows_to_insert = []
    for _, row in df.iterrows():
        ticker = str(row.get('Ticker', '') or '').strip()
        if not ticker or ticker in ('nan', '#VALUE!'): continue
        etf_name = str(row.get('ETF', '') or '').strip()
        if etf_name in ('nan', '#VALUE!', ''): etf_name = ''

        price_date = None
        if pd.notna(row.get('Date')):
            try: price_date = str(row['Date'])[:10]
            except: pass

        ter = safe_pct(row.get(ter_col)) if ter_col else None

        rows_to_insert.append((
            ticker,                          # etf_id (stable: use name from col A)
            ticker,                          # ticker
            etf_name if etf_name else ticker, # etf_name
            str(row.get('Exchange', '') or ''),
            str(row.get('Currency', '') or ''),
            safe(row.get('Price @ Date')), price_date,
            safe(row.get('YTD %')), safe(row.get('1Y %')), safe(row.get('3Y %')),
            safe(row.get('Mkt Cap')),
            str(row.get('Type', '') or ''), str(row.get('Sector', '') or ''),
            safe(row.get('52 week high')), safe(row.get('52 week low')),
            safe(row.get('Change (%)')), safe(row.get('Volume')),
            ter, now
        ))

    with get_db() as db:
        # Full replace — clears stale records when Refinitiv updates names
        db.execute('DELETE FROM etf_universe')
        for r in rows_to_insert:
            db.execute('''INSERT OR IGNORE INTO etf_universe
                (etf_id, ticker, etf_name, exchange, currency, price, price_date,
                 ytd_pct, one_y_pct, three_y_pct, mkt_cap, type, sector,
                 high_52w, low_52w, daily_chg_pct, volume, ter, updated_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ''', r)
            count += 1
        db.commit()

    return jsonify({'ok': True, 'synced': count})

@app.route('/api/etf-universe/search', methods=['GET'])
def search_etf_universe():
    q = request.args.get('q', '').strip()
    db = get_db()
    if not q:
        rows = db.execute('SELECT * FROM etf_universe ORDER BY ticker').fetchall()
    else:
        like = f'%{q}%'
        rows = db.execute(
            'SELECT * FROM etf_universe WHERE UPPER(ticker) LIKE UPPER(?) OR UPPER(etf_name) LIKE UPPER(?) OR UPPER(sector) LIKE UPPER(?)',
            (like, like, like)
        ).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/etf-universe', methods=['POST'])
def add_etf():
    """Add a new ETF identifier to Excel col B + DB placeholder."""
    import shutil
    from openpyxl import load_workbook
    from openpyxl.worksheet.formula import ArrayFormula

    data = request.json
    etf_id = (data.get('etf_id') or '').strip()
    if not etf_id:
        return jsonify({'error': 'ETF identifier required'}), 400
    if not os.path.exists(STOCK_EXCEL_PATH):
        return jsonify({'error': 'Excel file not found'}), 404

    shutil.copy2(STOCK_EXCEL_PATH, STOCK_EXCEL_PATH + '.bak')
    wb = load_workbook(STOCK_EXCEL_PATH)
    ws = wb['ETF Report']
    last_row = ws.max_row
    new_row = last_row + 1
    tmpl = 2

    # Col B = ETF identifier (user input Refinitiv reads)
    ws.cell(new_row, 2).value = etf_id

    # Copy array formulas: A, C-F, J, M-P
    for col_idx in [1, 3, 4, 5, 6, 10, 13, 14, 15, 16]:
        src = ws.cell(tmpl, col_idx).value
        dst = ws.cell(new_row, col_idx)
        if hasattr(src, 'text'):
            # Replace B2→B{new_row} in formula text for col A (_FV(B2,...))
            new_text = src.text.replace(f'B{tmpl}', f'B{new_row}')
            dst.value = ArrayFormula(f'{dst.column_letter}{new_row}', new_text)
        elif src:
            dst.value = src

    # G, H, I use F2, B2, E2 directly
    for col_idx in [7, 8, 9]:
        formula = ws.cell(tmpl, col_idx).value
        if formula:
            new_formula = (formula
                .replace(f'F{tmpl}', f'F{new_row}')
                .replace(f'B{tmpl}', f'B{new_row}')
                .replace(f'E{tmpl}', f'E{new_row}'))
            ws.cell(new_row, col_idx).value = new_formula

    # K (Type) and L (Sector) are plain text — leave blank for user to fill
    ws.cell(new_row, 11).value = ''
    ws.cell(new_row, 12).value = ''

    # Extend table
    for tbl in ws.tables.values():
        tbl.ref = f'A1:P{new_row}'

    wb.save(STOCK_EXCEL_PATH)
    wb.close()

    now = datetime.now().isoformat()
    db = get_db()
    db.execute(
        'INSERT OR IGNORE INTO etf_universe (etf_id, ticker, etf_name, updated_at) VALUES (?,?,?,?)',
        (etf_id, etf_id, f'[{etf_id}]', now)
    )
    db.commit()
    return jsonify({'ok': True,
                    'message': f'{etf_id} aggiunto. Apri Excel, lascia aggiornare Refinitiv, poi premi Sincronizza.'})

@app.route('/api/etf-universe/<int:sid>', methods=['DELETE'])
def delete_etf(sid):
    import shutil, pandas as pd
    from openpyxl import load_workbook

    db = get_db()
    etf = db.execute('SELECT * FROM etf_universe WHERE id=?', (sid,)).fetchone()
    if not etf:
        return jsonify({'error': 'Not found'}), 404

    ticker = etf['ticker']

    if os.path.exists(STOCK_EXCEL_PATH):
        shutil.copy2(STOCK_EXCEL_PATH, STOCK_EXCEL_PATH + '.bak')
        df = pd.read_excel(STOCK_EXCEL_PATH, sheet_name='ETF Report')
        excel_row = None
        for i, row in df.iterrows():
            if str(row.get('Ticker', '') or '').strip() == ticker:
                excel_row = i + 2
                break

        if excel_row:
            wb = load_workbook(STOCK_EXCEL_PATH)
            ws = wb['ETF Report']
            ws.delete_rows(excel_row)
            new_last = ws.max_row
            for tbl in ws.tables.values():
                tbl.ref = f'A1:P{new_last}'
            wb.save(STOCK_EXCEL_PATH)
            wb.close()

    db.execute('DELETE FROM etf_universe WHERE id=?', (sid,))
    db.commit()
    return jsonify({'ok': True})


# ── Excel category name → display name ────────────────────────────────────────
EXCEL_TO_DISPLAY = {
    # Azioni
    'Az. Area Euro':                        'Azioni Area Euro',
    'Az. Area Euro small cap':              'Azioni Area Euro Small Cap',
    'Az. Europa':                           'Azioni Europa',
    'Az. Europa small cap':                 'Azioni Europa Small Cap',
    'Az. Europa Growth':                    'Azioni Europa Growth',
    'Az. Europa Value':                     'Azioni Europa Value',
    'Az. Italia':                           'Azioni Italia',
    'Az. UK':                               'Azioni UK',
    'Az. Germania':                         'Azioni Germania',
    'Az. Spagna':                           'Azioni Spagna',
    'Az. Svizzera':                         'Azioni Svizzera',
    'Az.Svizzera small cap':                'Azioni Svizzera Small Cap',
    'Az. paesi nordici':                    'Azioni Paesi Nordici',
    'Az. USA':                              'Azioni USA',
    'Az. USA small cap':                    'Azioni USA Small Cap',
    'Az. USA Growth':                       'Azioni USA Growth',
    'Az. USA Value':                        'Azioni USA Value',
    'Az. Giappone':                         'Azioni Giappone',
    'Az. Giappone small cap':               'Azioni Giappone Small Cap',
    'Az. Asia Pacifico':                    'Azioni Asia Pacifico',
    'Az. Asia Pacifico ex Giapp.':          'Azioni Asia Pacifico ex Giappone',
    'Az. Pacifico':                         'Azioni Pacifico',
    'Az. Cina':                             'Azioni Cina',
    'Az. Grande Cina':                      'Azioni Grande Cina',
    'Az. India':                            'Azioni India',
    'Az. paesi emerg. Asia':                'Azioni Paesi Emergenti Asia',
    'Az. paesi emerg. Mondo':               'Azioni Paesi Emergenti Mondo',
    'Az. paesi emerg. America Latina':      'Azioni Paesi Emergenti America Latina',
    'Az. paesi emerg. altre zone':          'Azioni Paesi Emergenti Altre Zone',
    'Az. Altri paesi emerg.':               'Azioni Altri Paesi Emergenti',
    'Az. BRIC':                             'Azioni BRIC',
    'Az. Paesi Emerg. Europa e Russia':     'Azioni Paesi Emergenti Europa e Russia',
    'Az. globale':                          'Azioni Globale',
    'Az. globale small cap':                'Azioni Globale Small Cap',
    'Az. globale Growth':                   'Azioni Globale Growth',
    'Az. globale Value':                    'Azioni Globale Value',
    'Az. Oro':                              'Azioni Oro',
    'Az. real estate Europa':               'Azioni Real Estate Europa',
    'Az. real estate Mondo':                'Azioni Real Estate Mondo',
    'Az. Biotech':                          'Azioni Biotech',
    'Az. beni di consumo':                  'Azioni Beni di Consumo',
    'Az. ambiente':                         'Azioni Ambiente',
    'Az. energia materie prime oro':        'Azioni Energia Materie Prime Oro',
    'Az. industria':                        'Azioni Industria',
    'Az. salute - farmaceutico':            'Azioni Salute - Farmaceutico',
    'Az. Servizi di pubblica utilita':      'Azioni Servizi di Pubblica Utilità',
    'Az. servizi finanziari':               'Azioni Servizi Finanziari',
    'Az. tecnologia':                       'Azioni Tecnologia',
    'Az. telecomunicazioni':               'Azioni Telecomunicazioni',
    # Obbligazioni
    'Obblig. Euro':                         'Obbligazioni Euro',
    'Obblig. Euro breve term.':             'Obbligazioni Euro Breve Termine',
    'Obblig. Euro medio termine':           'Obbligazioni Euro Medio Termine',
    'Obblig. Euro lungo termine':           'Obbligazioni Euro Lungo Termine',
    'Obblig. Euro gov.':                    'Obbligazioni Euro Governativi',
    'Obblig. Euro gov. medio termine':      'Obbligazioni Euro Governativi Medio Termine',
    'Obblig. Euro gov. lungo termine':      'Obbligazioni Euro Governativi Lungo Termine',
    'Obblig. euro gov. breve termine':      'Obbligazioni Euro Governativi Breve Termine',
    'Obblig. Euro all maturities':          'Obbligazioni Euro Tutte le Scadenze',
    'Obblig. Euro corporate':               'Obbligazioni Euro Corporate',
    'Obblig. Euro high yield':              'Obbligazioni Euro High Yield',
    'Obblig. Euro a scadenza':              'Obbligazioni Euro a Scadenza',
    'Obblig. Europa':                       'Obbligazioni Europa',
    'Obblig. Europa High Yield':            'Obbligazioni Europa High Yield',
    'Obblig. Dollaro US':                   'Obbligazioni Dollaro US',
    'Obblig. Dollaro US breve term.':       'Obbligazioni Dollaro US Breve Termine',
    'Obblig. Dollaro US medio-lungo termine': 'Obbligazioni Dollaro US Medio-Lungo Termine',
    'Obblig. Dollaro US corporate':         'Obbligazioni Dollaro US Corporate',
    'Obblig. Dollaro US all mat':           'Obbligazioni Dollaro US Tutte le Scadenze',
    'Obblig. Dollaro US high yield':        'Obbligazioni Dollaro US High Yield',
    'Obblig. Asia':                         'Obbligazioni Asia',
    'Obblig. Paesi Emerg.':                 'Obbligazioni Paesi Emergenti',
    'Obblig. Paesi Emerg. Asia Local Ccy':  'Obbligazioni Paesi Emergenti Asia Valuta Locale',
    'Obblig. Paesi Emerg. Hard Currency':   'Obbligazioni Paesi Emergenti Hard Currency',
    'Obblig. Paesi Emerg. Local Currency':  'Obbligazioni Paesi Emergenti Valuta Locale',
    'Obblig. paesi emerg. Asia':            'Obbligazioni Paesi Emergenti Asia',
    'Obblig. globale':                      'Obbligazioni Globale',
    'Obblig. globale corporate':            'Obbligazioni Globale Corporate',
    'Obblig. globale high yield':           'Obbligazioni Globale High Yield',
    'Obblig. altre valute':                 'Obbligazioni Altre Valute',
    'Obblig. Franco svizzero':              'Obbligazioni Franco Svizzero',
    'Obblig. Global Inflation Linked':      'Obbligazioni Globale Inflation Linked',
    'Obblig. Indicizz. Inflation Linked':   'Obbligazioni Indicizzate Inflation Linked',
    'Obblig. Convertib. Europa':            'Obbligazioni Convertibili Europa',
    'Obblig. Convertib. Euro':              'Obbligazioni Convertibili Euro',
    'Obblig. Convertib. Glob.':             'Obbligazioni Convertibili Globale',
    'Obblig. Stutturati':                   'Obbligazioni Strutturati',
    # Monetari
    'Monetari altre valute europ':          'Monetari Altre Valute Europee',
    'Monetari ex Europa altre valute':      'Monetari ex Europa Altre Valute',
    # Bilanciati
    'Bilanc. Prud. Europa':                 'Bilanciati Prudenti Europa',
    'Bilanc. Prud. Global':                 'Bilanciati Prudenti Globale',
    'Bilanc. Prud. Dollaro US':             'Bilanciati Prudenti Dollaro US',
    'Bilanc. Equilib. Europa':              'Bilanciati Equilibrati Europa',
    'Bilanc. Equil. Global':                'Bilanciati Equilibrati Globale',
    'Bilanc. Equil. Dollaro US':            'Bilanciati Equilibrati Dollaro US',
    'Bilanc. Aggress. Global':              'Bilanciati Offensivi Globale',
    'Bilanc. Aggress. altre valute':        'Bilanciati Offensivi Altre Valute',
    'Bilanc. aggress. Dollaro US':          'Bilanciati Offensivi Dollaro US',
    'Bilanc. Aggress. Europa':              'Bilanciati Offensivi Europa',
    # Flessibili
    'Fless. Global':                        'Flessibili Globale',
    'Flessibili prudenti Europa':           'Flessibili Prudenti Europa',
    'Flessibili prudenti globale':          'Flessibili Prudenti Globale',
    'Fondi a scadenza pred. Euro':          'Fondi a Scadenza Euro',
    'Fondi  a garanzia o a formula Euro':   'Fondi a Garanzia Euro',
    # Performance Assoluta
    'Perf. ass. Long/Short eq.':            'Performance Assoluta Long/Short Equity',
    'Perf. Ass. Arbitr.Fus.-acquis. Euro':  'Performance Assoluta Arbitraggi Fusioni-Acquisizioni Euro',
    'Perf. assoluta strategia valute':      'Performance Assoluta Strategia Valute',
    'Perf. assoluta Market Neutral Euro':   'Performance Assoluta Market Neutral Euro',
    'Perf. assoluta tassi':                 'Performance Assoluta Tassi',
    'Perf. assoluta volatilita':            'Performance Assoluta Volatilità',
    'Perf. assoluta multi-strategia':       'Performance Assoluta Multi-Strategia',
    'Perf. ass. USD':                       'Performance Assoluta USD',
}

# ── Category hierarchy (all display names) ────────────────────────────────────
CATEGORY_TREE = {
    'equity': [
        {'name': 'Azioni Area Euro',         'children': ['Azioni Area Euro Small Cap']},
        {'name': 'Azioni Europa',            'children': ['Azioni Europa Small Cap', 'Azioni Europa Growth', 'Azioni Europa Value']},
        {'name': 'Azioni Italia'},
        {'name': 'Azioni UK'},
        {'name': 'Azioni Germania'},
        {'name': 'Azioni Spagna'},
        {'name': 'Azioni Svizzera',          'children': ['Azioni Svizzera Small Cap']},
        {'name': 'Azioni Paesi Nordici'},
        {'name': 'Azioni USA',               'children': ['Azioni USA Small Cap', 'Azioni USA Growth', 'Azioni USA Value']},
        {'name': 'Azioni Giappone',          'children': ['Azioni Giappone Small Cap']},
        {'name': 'Azioni Asia Pacifico',     'children': ['Azioni Asia Pacifico ex Giappone']},
        {'name': 'Azioni Pacifico'},
        {'name': 'Azioni Cina'},
        {'name': 'Azioni Grande Cina'},
        {'name': 'Azioni India'},
        {'name': 'Azioni Paesi Emergenti',   'children': [
            'Azioni Paesi Emergenti Asia', 'Azioni Paesi Emergenti Mondo',
            'Azioni Paesi Emergenti America Latina', 'Azioni Paesi Emergenti Altre Zone',
            'Azioni Altri Paesi Emergenti', 'Azioni BRIC', 'Azioni Paesi Emergenti Europa e Russia',
        ]},
        {'name': 'Azioni Globale',           'children': ['Azioni Globale Small Cap', 'Azioni Globale Growth', 'Azioni Globale Value']},
        {'name': 'Azioni Oro'},
        {'name': 'Azioni Real Estate Europa'},
        {'name': 'Azioni Real Estate Mondo'},
        {'name': 'Azioni Biotech'},
        {'name': 'Azioni Beni di Consumo'},
        {'name': 'Azioni Ambiente'},
        {'name': 'Azioni Energia Materie Prime Oro'},
        {'name': 'Azioni Industria'},
        {'name': 'Azioni Salute - Farmaceutico'},
        {'name': 'Azioni Servizi di Pubblica Utilità'},
        {'name': 'Azioni Servizi Finanziari'},
        {'name': 'Azioni Tecnologia'},
        {'name': 'Azioni Telecomunicazioni'},
    ],
    'bond': [
        {'name': 'Obbligazioni Euro',        'children': [
            'Obbligazioni Euro Breve Termine', 'Obbligazioni Euro Medio Termine',
            'Obbligazioni Euro Lungo Termine', 'Obbligazioni Euro Governativi',
            'Obbligazioni Euro Governativi Medio Termine', 'Obbligazioni Euro Governativi Lungo Termine',
            'Obbligazioni Euro Governativi Breve Termine', 'Obbligazioni Euro Tutte le Scadenze',
            'Obbligazioni Euro Corporate', 'Obbligazioni Euro High Yield', 'Obbligazioni Euro a Scadenza',
        ]},
        {'name': 'Obbligazioni Europa',      'children': ['Obbligazioni Europa High Yield']},
        {'name': 'Obbligazioni Dollaro US',  'children': [
            'Obbligazioni Dollaro US Breve Termine', 'Obbligazioni Dollaro US Medio-Lungo Termine',
            'Obbligazioni Dollaro US Corporate', 'Obbligazioni Dollaro US Tutte le Scadenze',
            'Obbligazioni Dollaro US High Yield',
        ]},
        {'name': 'Obbligazioni Asia'},
        {'name': 'Obbligazioni Paesi Emergenti', 'children': [
            'Obbligazioni Paesi Emergenti Asia Valuta Locale', 'Obbligazioni Paesi Emergenti Hard Currency',
            'Obbligazioni Paesi Emergenti Valuta Locale', 'Obbligazioni Paesi Emergenti Asia',
        ]},
        {'name': 'Obbligazioni Globale',     'children': ['Obbligazioni Globale Corporate', 'Obbligazioni Globale High Yield']},
        {'name': 'Obbligazioni Altre Valute','children': ['Obbligazioni Franco Svizzero']},
        {'name': 'Obbligazioni Inflation Linked', 'children': [
            'Obbligazioni Globale Inflation Linked', 'Obbligazioni Indicizzate Inflation Linked',
        ]},
        {'name': 'Obbligazioni Convertibili','children': [
            'Obbligazioni Convertibili Europa', 'Obbligazioni Convertibili Euro', 'Obbligazioni Convertibili Globale',
        ]},
        {'name': 'Obbligazioni Strutturati'},
        {'name': 'Monetari Europa',          'children': ['Monetari Euro', 'Monetari Altre Valute Europee']},
        {'name': 'Monetari ex Europa',       'children': ['Monetari Dollaro USA', 'Monetari ex Europa Altre Valute']},
    ],
    'bilanciato': [
        {'name': 'Bilanciati Prudenti',      'children': ['Bilanciati Prudenti Europa', 'Bilanciati Prudenti Globale', 'Bilanciati Prudenti Dollaro US']},
        {'name': 'Bilanciati Equilibrati',   'children': ['Bilanciati Equilibrati Europa', 'Bilanciati Equilibrati Globale', 'Bilanciati Equilibrati Dollaro US']},
        {'name': 'Bilanciati Offensivi',     'children': ['Bilanciati Offensivi Globale', 'Bilanciati Offensivi Altre Valute', 'Bilanciati Offensivi Dollaro US', 'Bilanciati Offensivi Europa']},
    ],
    'flessibili': [
        {'name': 'Flessibili Europa'},
        {'name': 'Flessibili Globale'},
        {'name': 'Flessibili Dollaro US'},
        {'name': 'Flessibili Prudenti Europa'},
        {'name': 'Flessibili Prudenti Globale'},
        {'name': 'Fondi a Scadenza Euro',    'children': ['Fondi a Garanzia Euro']},
    ],
    'commodity': [
        {'name': 'Commodities'},
    ],
    'altri': [
        {'name': 'Performance Assoluta Euro', 'children': [
            'Performance Assoluta Long/Short Equity', 'Performance Assoluta Arbitraggi Fusioni-Acquisizioni Euro',
            'Performance Assoluta Strategia Valute', 'Performance Assoluta Market Neutral Euro',
            'Performance Assoluta Tassi', 'Performance Assoluta Volatilità', 'Performance Assoluta Multi-Strategia',
        ]},
        {'name': 'Performance Assoluta Altre Valute', 'children': ['Performance Assoluta USD']},
        {'name': 'Altri'},
    ],
}

def _rename_fund_keys(funds_dict):
    """Rename Excel category keys to display names using EXCEL_TO_DISPLAY."""
    renamed = {}
    for k, v in funds_dict.items():
        renamed[EXCEL_TO_DISPLAY.get(k, k)] = v
    return renamed

def _aggregate_funds(funds_dict):
    """Rename Excel keys to display names, then pool parent-level entries."""
    funds_dict = _rename_fund_keys(funds_dict)
    for nodes in CATEGORY_TREE.values():
        for node in nodes:
            children = node.get('children', [])
            if not children:
                continue
            parent = node['name']
            combined = list(funds_dict.get(parent, []))
            for child in children:
                combined.extend(funds_dict.get(child, []))
            if combined:
                funds_dict[parent] = combined
    return funds_dict

def static_version():
    """Cache-buster: newest mtime among static assets, so redeploys invalidate automatically."""
    latest = 0
    static_dir = os.path.join(app.root_path, 'static')
    for root, _, files in os.walk(static_dir):
        for f in files:
            try: latest = max(latest, int(os.path.getmtime(os.path.join(root, f))))
            except OSError: pass
    return latest

@app.route('/')
def index():
    return render_template('index.html', v=static_version())

# ── Financial consultants ─────────────────────────────────────────────────────
@app.route('/api/consultants', methods=['GET'])
def get_consultants():
    db = get_db()
    rows = db.execute('SELECT id,name,created_at FROM consultants ORDER BY name').fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/consultants', methods=['POST'])
def create_consultant():
    name = (request.json or {}).get('name','').strip()
    if not name: return jsonify({'error':'Name required'}),400
    db = get_db()
    cur = db.execute('INSERT INTO consultants (name) VALUES (?)',(name,))
    db.commit()
    return jsonify({'id':cur.lastrowid,'name':name})

@app.route('/api/consultants/<int:fid>', methods=['PUT'])
def update_consultant(fid):
    name = (request.json or {}).get('name','').strip()
    if not name: return jsonify({'error':'Name required'}),400
    db = get_db()
    db.execute('UPDATE consultants SET name=? WHERE id=?',(name,fid))
    db.commit()
    return jsonify({'ok':True})

@app.route('/api/consultants/<int:fid>', methods=['DELETE'])
def delete_consultant(fid):
    # Refuse while clients are attached, so a click can't wipe a whole book.
    db = get_db()
    n = db.execute('SELECT COUNT(*) FROM clients WHERE consultant_id=?',(fid,)).fetchone()[0]
    if n: return jsonify({'error':f'Il consulente ha ancora {n} clienti'}),409
    db.execute('DELETE FROM consultants WHERE id=?',(fid,))
    db.commit()
    return jsonify({'ok':True})

# ── Clients ───────────────────────────────────────────────────────────────────
@app.route('/api/clients', methods=['GET'])
def get_clients():
    db = get_db()
    clients = db.execute('SELECT id,name,consultant_id,created_at,updated_at FROM clients ORDER BY name').fetchall()
    result = []
    for c in clients:
        portfolios = db.execute('SELECT type FROM portfolios WHERE client_id=?',(c['id'],)).fetchall()
        result.append({
            'id':c['id'],'name':c['name'],'consultant_id':c['consultant_id'],
            'created_at':c['created_at'],'updated_at':c['updated_at'],
            'portfolios':[p['type'] for p in portfolios]
        })
    return jsonify(result)

@app.route('/api/clients', methods=['POST'])
def create_client():
    data = request.json
    name = data.get('name','').strip()
    if not name: return jsonify({'error':'Name required'}),400
    db = get_db()
    fid = data.get('consultant_id') or db.execute('SELECT MIN(id) FROM consultants').fetchone()[0]
    cur = db.execute('INSERT INTO clients (name,consultant_id) VALUES (?,?)',(name,fid))
    db.commit()
    return jsonify({'id':cur.lastrowid,'name':name,'consultant_id':fid})

@app.route('/api/clients/<int:cid>', methods=['PUT'])
def update_client(cid):
    data = request.json
    db = get_db()
    # Move to another consultant (may be sent on its own, without a name)
    if 'consultant_id' in data:
        fid = data['consultant_id']
        if not db.execute('SELECT 1 FROM consultants WHERE id=?',(fid,)).fetchone():
            return jsonify({'error':'Consultant not found'}),404
        db.execute('UPDATE clients SET consultant_id=? WHERE id=?',(fid,cid))
        if 'name' not in data:
            db.commit()
            return jsonify({'ok':True})
    name = data.get('name','').strip()
    if not name: return jsonify({'error':'Name required'}),400
    db.execute('UPDATE clients SET name=?,updated_at=? WHERE id=?',(name,datetime.now().isoformat(),cid))
    db.commit()
    return jsonify({'ok':True})

@app.route('/api/clients/<int:cid>', methods=['DELETE'])
def delete_client(cid):
    db = get_db()
    db.execute('DELETE FROM clients WHERE id=?',(cid,))
    db.commit()
    return jsonify({'ok':True})

@app.route('/api/clients/<int:cid>/duplicate', methods=['POST'])
def duplicate_client(cid):
    db = get_db()
    orig = db.execute('SELECT name,consultant_id FROM clients WHERE id=?',(cid,)).fetchone()
    if not orig: return jsonify({'error':'Not found'}),404
    new_name = request.json.get('name', orig['name']+' (copia)').strip()
    cur = db.execute('INSERT INTO clients (name,consultant_id) VALUES (?,?)',(new_name,orig['consultant_id']))
    new_id = cur.lastrowid
    portfolios = db.execute('SELECT type,data FROM portfolios WHERE client_id=?',(cid,)).fetchall()
    now = datetime.now().isoformat()
    for p in portfolios:
        db.execute('INSERT INTO portfolios (client_id,type,data,updated_at) VALUES (?,?,?,?)',
                   (new_id,p['type'],p['data'],now))
    db.commit()
    return jsonify({'id':new_id,'name':new_name})

# ── PythonAnywhere self-service renewal ───────────────────────────────────────
PA_API_USER = 'Pragyesh'
PA_API_DOMAIN = 'pragyesh.pythonanywhere.com'

def _pa_api(method, path=''):
    """Call the PythonAnywhere webapps API. Returns (json, error)."""
    import urllib.request
    tok_path = os.path.expanduser('~/.pythonanywhere_token')
    if not os.path.exists(tok_path):
        return None, 'token non configurato sul server'
    token = open(tok_path).read().strip()
    url = f'https://www.pythonanywhere.com/api/v0/user/{PA_API_USER}/webapps/{PA_API_DOMAIN}/{path}'
    req = urllib.request.Request(url, method=method,
                                 headers={'Authorization': f'Token {token}'},
                                 data=b'' if method == 'POST' else None)
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            body = r.read()
            return (json.loads(body) if body else {}), None
    except Exception as e:
        return None, str(e)

_pa_status_cache = {'ts': 0, 'resp': None}

@app.route('/api/pa/status')
def pa_status():
    # The PA API can take seconds; cache 10 min so page loads don't block on it.
    import time
    now = time.time()
    if not request.args.get('fresh') and _pa_status_cache['resp'] and now - _pa_status_cache['ts'] < 600:
        return jsonify(_pa_status_cache['resp'])
    info, err = _pa_api('GET')
    if err: return jsonify({'error': err}), 503
    days = None
    try:
        from datetime import date
        y, m, d = map(int, info['expiry'].split('-'))
        days = (date(y, m, d) - date.today()).days
    except Exception: pass
    payload = {'expiry': info.get('expiry'), 'enabled': info.get('enabled'), 'days_left': days}
    _pa_status_cache['ts'] = now
    _pa_status_cache['resp'] = payload
    return jsonify(payload)

@app.route('/api/pa/renew', methods=['POST'])
def pa_renew():
    # Ottobre 2026: PythonAnywhere non rinnova più il piano gratuito con disable→enable
    # (verificato: la scadenza non cambia) e non c'è un endpoint API di rinnovo.
    # Il rinnovo si fa dal pulsante "Run until 1 month from today" nella pagina Web
    # dell'account: il frontend apre quella pagina (vedi paRenew in main.js).
    return jsonify({'error': 'Rinnovo da fare su pythonanywhere.com (Web → Run until 1 month from today)'}), 410

# ── Price refresh on page open ───────────────────────────────────────────────
# Opening the page records a refresh request. On the Mac (local run) the bot is
# started right away; online, the Mac's launchd job polls /api/prices/status
# every 5 min and runs the bot when a request is newer than the last refresh.
# refresh_prices.py itself throttles, locks and retries.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRICE_XLSX = os.path.join(BASE_DIR, 'Stock + ETF Report.xlsx')
PRICE_REQ = os.path.join(BASE_DIR, '.price_refresh_request')
PRICE_BOT = os.path.join(BASE_DIR, 'refresh_prices.py')
IS_MAC = __import__('sys').platform == 'darwin'
PRICE_FRESH_S = 15 * 60  # prices younger than this are not refreshed again

def _mtime(path):
    try: return os.path.getmtime(path)
    except OSError: return None

def _price_bot_running():
    if not IS_MAC: return False
    import subprocess
    r = subprocess.run(['pgrep', '-f', 'refresh_prices.py'], capture_output=True)
    return r.returncode == 0

@app.route('/api/prices/request', methods=['POST'])
def prices_request():
    import time, subprocess, sys
    now = time.time()
    upd = _mtime(PRICE_XLSX)
    if upd and now - upd < PRICE_FRESH_S:
        return jsonify({'fresh': True, 'prices_updated_at': upd, 'local': IS_MAC})
    with open(PRICE_REQ, 'w') as f: f.write(str(now))
    started = False
    if IS_MAC and os.path.exists(PRICE_BOT) and not _price_bot_running():
        subprocess.Popen([sys.executable, PRICE_BOT, '--scheduled'], cwd=BASE_DIR,
                         stdout=open(os.path.expanduser('~/Library/Logs/portfolio-refresh.log'), 'a'),
                         stderr=subprocess.STDOUT, start_new_session=True)
        started = True
    return jsonify({'requested_at': now, 'local': IS_MAC, 'started': started,
                    'prices_updated_at': _mtime(PRICE_XLSX)})

@app.route('/api/prices/status')
def prices_status():
    req = None
    try: req = float(open(PRICE_REQ).read().strip())
    except (OSError, ValueError): pass
    return jsonify({'requested_at': req, 'prices_updated_at': _mtime(PRICE_XLSX),
                    'running': _price_bot_running(), 'local': IS_MAC})

# ── Dashboard (all clients + portfolios in one call) ─────────────────────────
@app.route('/api/dashboard', methods=['GET'])
def get_dashboard():
    db = get_db()
    clients = db.execute('SELECT id,name,consultant_id,created_at,updated_at FROM clients ORDER BY name').fetchall()
    rows = db.execute('SELECT client_id,type,data,updated_at FROM portfolios').fetchall()
    by_client = {}
    for r in rows:
        by_client.setdefault(r['client_id'], {})[r['type']] = {
            'data': json.loads(r['data']), 'updated_at': r['updated_at']}
    return jsonify([{
        'id': c['id'], 'name': c['name'], 'consultant_id': c['consultant_id'],
        'created_at': c['created_at'], 'updated_at': c['updated_at'],
        'portfolios': list(by_client.get(c['id'], {}).keys()),
        'portfolios_data': by_client.get(c['id'], {}),
    } for c in clients])

# ── Portfolios ────────────────────────────────────────────────────────────────
@app.route('/api/clients/<int:cid>/portfolios', methods=['GET'])
def get_portfolios(cid):
    db = get_db()
    rows = db.execute('SELECT type,data,updated_at FROM portfolios WHERE client_id=?',(cid,)).fetchall()
    return jsonify({r['type']:{'data':json.loads(r['data']),'updated_at':r['updated_at']} for r in rows})

@app.route('/api/clients/<int:cid>/portfolios', methods=['POST'])
def save_portfolios(cid):
    data = request.json
    db = get_db()
    now = datetime.now().isoformat()
    existing = {r['type'] for r in db.execute('SELECT type FROM portfolios WHERE client_id=?',(cid,)).fetchall()}
    for ptype,pdata in data.items():
        if ptype in existing:
            db.execute('UPDATE portfolios SET data=?,updated_at=? WHERE client_id=? AND type=?',
                       (json.dumps(pdata),now,cid,ptype))
        else:
            db.execute('INSERT INTO portfolios (client_id,type,data,updated_at) VALUES (?,?,?,?)',
                       (cid,ptype,json.dumps(pdata),now))
    for t in existing:
        if t not in data:
            db.execute('DELETE FROM portfolios WHERE client_id=? AND type=?',(cid,t))
    db.execute('UPDATE clients SET updated_at=? WHERE id=?',(now,cid))
    db.commit()
    return jsonify({'ok':True})

# ── ISIN lookup proxy ─────────────────────────────────────────────────────────
@app.route('/api/isin/<isin>', methods=['GET'])
def lookup_isin(isin):
    import urllib.request, urllib.parse
    isin = isin.strip().upper()
    euro = ['MIL','GER','FRA','AMS','LSE','EBS','VIE','BRU','CPH','HEL','OSL','STK','MCE','SWX']
    hdrs = {'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/120','Accept':'application/json'}
    quotes = []
    for host in ['query1','query2']:
        try:
            url = f'https://{host}.finance.yahoo.com/v1/finance/search?q={urllib.parse.quote(isin)}&quotesCount=10&newsCount=0&listsCount=0&enableFuzzyQuery=false'
            req = urllib.request.Request(url, headers=hdrs)
            with urllib.request.urlopen(req, timeout=10) as r:
                quotes = json.loads(r.read()).get('quotes',[])
            if quotes: break
        except: continue
    if not quotes: return jsonify({'error':'not found'}),404
    best = (next((q for q in quotes if q.get('exchange') in euro and q.get('quoteType') in ('ETF','MUTUALFUND')),None)
         or next((q for q in quotes if q.get('quoteType') in ('ETF','MUTUALFUND')),None)
         or next((q for q in quotes if q.get('exchange') in euro),None)
         or quotes[0])
    name = best.get('longname') or best.get('shortname') or best.get('symbol',isin)
    qt = best.get('quoteType','')
    ftype = 'ETF' if 'ETF' in qt else 'Azione' if 'EQUITY' in qt else 'Obbligazione' if any(x in qt for x in ['BOND','INCOME']) else 'Fondo'
    symbol = best.get('symbol','')
    price = None
    if symbol:
        for host in ['query1','query2']:
            try:
                url = f'https://{host}.finance.yahoo.com/v8/finance/chart/{urllib.parse.quote(symbol)}?interval=1d&range=1d'
                req = urllib.request.Request(url, headers=hdrs)
                with urllib.request.urlopen(req, timeout=10) as r:
                    pd = json.loads(r.read())
                meta = pd.get('chart',{}).get('result',[{}])[0].get('meta',{})
                price = meta.get('regularMarketPrice') or meta.get('previousClose')
                if price: break
            except: continue
    return jsonify({'name':name,'price':price,'type':ftype,'ticker':symbol})

# ── FX rates (EUR base · no API key · ECB via Frankfurter, er-api fallback) ─────
_fx_cache = {'ts': 0, 'data': None}
_fx_hist = {}  # 'YYYY-MM-DD' -> data (historical rates never change → cache forever)

def get_fx_rates(date=None):
    """Return {'base':'EUR','date':..,'rates':{cur:rate}} of EUR→cur rates.
    With `date` (YYYY-MM-DD) returns the ECB rate for that day (Frankfurter
    historical, snaps to the prior business day on weekends/holidays); falls
    back to live rates if the historical fetch fails.
    Without a date: live rates cached 6h. Keyless: Frankfurter, fallback er-api.
    On total failure serves the last good value, else just {EUR:1}."""
    import time, urllib.request
    hdrs = {'User-Agent':'Mozilla/5.0','Accept':'application/json'}
    if date:
        if date in _fx_hist:
            return _fx_hist[date]
        try:
            req = urllib.request.Request(f'https://api.frankfurter.dev/v1/{date}?base=EUR', headers=hdrs)
            with urllib.request.urlopen(req, timeout=8) as r:
                j = json.loads(r.read())
            rates = j.get('rates') or {}
            if rates:
                rates['EUR'] = 1.0
                data = {'base':'EUR','date':j.get('date'),'rates':rates,'source':'frankfurter-hist'}
                _fx_hist[date] = data
                return data
        except Exception:
            pass
        return get_fx_rates()  # fall back to live
    now = time.time()
    if _fx_cache['data'] and (now - _fx_cache['ts'] < 6*3600):
        return _fx_cache['data']
    data = None
    try:  # Primary: Frankfurter (ECB reference rates)
        req = urllib.request.Request('https://api.frankfurter.dev/v1/latest?base=EUR', headers=hdrs)
        with urllib.request.urlopen(req, timeout=8) as r:
            j = json.loads(r.read())
        rates = j.get('rates') or {}
        if rates:
            rates['EUR'] = 1.0
            data = {'base':'EUR','date':j.get('date'),'rates':rates,'source':'frankfurter'}
    except Exception:
        data = None
    if not data:
        try:  # Fallback: open.er-api.com
            req = urllib.request.Request('https://open.er-api.com/v6/latest/EUR', headers=hdrs)
            with urllib.request.urlopen(req, timeout=8) as r:
                j = json.loads(r.read())
            rates = j.get('rates') or {}
            if rates:
                rates['EUR'] = 1.0
                data = {'base':'EUR','date':j.get('time_last_update_utc'),'rates':rates,'source':'er-api'}
        except Exception:
            data = None
    if data:
        _fx_cache['ts'] = now
        _fx_cache['data'] = data
        return data
    return _fx_cache['data'] or {'base':'EUR','date':None,'rates':{'EUR':1.0},'source':'none'}

@app.route('/api/fx', methods=['GET'])
def api_fx():
    date = (request.args.get('date') or '').strip()
    return jsonify(get_fx_rates(date if date else None))

# ── CG lookup ─────────────────────────────────────────────────────────────────
@app.route('/api/cg-lookup', methods=['GET'])
@cache_by_mtime(DARTA_PATH)
def cg_lookup():
    from openpyxl import load_workbook
    path = DARTA_PATH
    empty = {'isin': {}, 'name': {}}
    if not os.path.exists(path):
        return jsonify({'pro': empty, 'chpv': empty})
    def sc_fmt(v):
        if v is None: return ''
        v = _sc_decimal(v)
        try: return f'{float(v)*100:.4g}%'
        except: return ''
    wb = load_workbook(path, data_only=True, read_only=True)
    def fi(kw, hdrs):
        for i, h in enumerate(hdrs):
            if h and kw in str(h).lower(): return i
        return None
    pro_isin, pro_name = {}, {}
    ws_pro = wb['Challenge Pro']
    hdrs_pro = [c.value for c in next(ws_pro.iter_rows(min_row=1, max_row=1))]
    pi_name = fi('fondo interno', hdrs_pro)
    pi_isin = fi('isin under',    hdrs_pro)
    pi_cg   = fi('cg',           hdrs_pro)
    pi_sc   = fi('spese correnti',hdrs_pro)
    for row in ws_pro.iter_rows(min_row=2, values_only=True):
        if not row: continue
        cg = row[pi_cg] if pi_cg is not None and len(row) > pi_cg else None
        if cg is None: continue
        name = row[pi_name] if pi_name is not None and len(row) > pi_name else None
        isin = row[pi_isin] if pi_isin is not None and len(row) > pi_isin else None
        sc   = row[pi_sc]   if pi_sc   is not None and len(row) > pi_sc   else None
        nm = str(name).strip() if name else ''
        isin_k = str(isin).strip().upper() if isin else ''
        valid_isin = isin_k not in ('', 'NAN', '0', 'NONE')
        cg_s = str(cg).strip().replace(',', '.')
        sc_s = sc_fmt(sc)
        if valid_isin:
            pro_isin[isin_k] = {'cg': cg_s, 'sc': sc_s, 'name': nm}
        if nm:
            pro_name[nm.upper()] = {'cg': cg_s, 'sc': sc_s, 'isin': isin_k if valid_isin else ''}
    chpv_isin, chpv_name = {}, {}
    ws_chpv = wb['Challenge Private']
    hdrs_chpv = [c.value for c in next(ws_chpv.iter_rows(min_row=1, max_row=1))]
    ci_name = fi('fondo interno', hdrs_chpv)
    ci_isin = fi('isin under',    hdrs_chpv)
    ci_sc   = fi('spese correnti',hdrs_chpv)
    for row in ws_chpv.iter_rows(min_row=2, values_only=True):
        if not row: continue
        sc = row[ci_sc] if ci_sc is not None and len(row) > ci_sc else None
        sc_s = sc_fmt(sc)
        if not sc_s: continue
        name = row[ci_name] if ci_name is not None and len(row) > ci_name else None
        isin = row[ci_isin] if ci_isin is not None and len(row) > ci_isin else None
        nm = str(name).strip() if name else ''
        isin_k = str(isin).strip().upper() if isin else ''
        valid_isin = isin_k not in ('', 'NAN', '0', 'NONE')
        if valid_isin:
            chpv_isin[isin_k] = {'sc': sc_s, 'name': nm}
        if nm:
            chpv_name[nm.upper()] = {'sc': sc_s, 'isin': isin_k if valid_isin else ''}
    wb.close()
    return jsonify({'pro': {'isin': pro_isin, 'name': pro_name}, 'chpv': {'isin': chpv_isin, 'name': chpv_name}})

# ── Category tree endpoint ─────────────────────────────────────────────────────
@app.route('/api/category-tree', methods=['GET'])
def category_tree():
    return jsonify(CATEGORY_TREE)

# ── DARTA categories from Excel ────────────────────────────────────────────────
@app.route('/api/darta-categories', methods=['GET'])
def darta_categories():
    result = {group: [n['name'] for n in nodes] for group, nodes in CATEGORY_TREE.items()}
    return jsonify(result)

# ── DARTA funds per category ───────────────────────────────────────────────────
@app.route('/api/darta-funds-by-category', methods=['GET'])
@cache_by_mtime(DARTA_PATH)
def darta_funds_by_category():
    from openpyxl import load_workbook
    path = DARTA_PATH
    result = {'pro': {}, 'chpv': {}}
    if not os.path.exists(path):
        return jsonify(result)
    def sc_fmt(v):
        if v is None: return ''
        v = _sc_decimal(v)
        try: return f'{float(v)*100:.4g}%'
        except: return ''
    wb = load_workbook(path, data_only=True, read_only=True)
    sheets = {'Challenge Pro': 'pro', 'Challenge Private': 'chpv'}
    for sheet_name, key in sheets.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        hdrs = [str(h).strip().lower() if h else '' for h in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
        def fi(kw):
            for i, h in enumerate(hdrs):
                if kw in h: return i
            return None
        i_name = fi('fondo interno')
        i_isin = fi('isin under')
        i_cat  = fi('categoria quantalys')
        i_cg   = fi('cg')
        i_sc   = fi('spese correnti')
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or i_cat is None or len(row) <= i_cat: continue
            cat = row[i_cat]
            if not cat: continue
            cat = str(cat).strip()
            nm   = str(row[i_name]).strip() if i_name is not None and len(row) > i_name and row[i_name] else ''
            isin = str(row[i_isin]).strip().upper() if i_isin is not None and len(row) > i_isin and row[i_isin] else ''
            if isin in ('', 'NAN', 'NONE', '0'): isin = ''
            cg = ''
            if i_cg is not None and len(row) > i_cg and row[i_cg] is not None:
                cg = str(row[i_cg]).strip().replace(',', '.')
            sc = sc_fmt(row[i_sc]) if i_sc is not None and len(row) > i_sc else ''
            if not nm: continue
            if cat not in result[key]: result[key][cat] = []
            result[key][cat].append({'name': nm, 'isin': isin, 'cg': cg, 'sc': sc})
    wb.close()
    result['pro']  = _aggregate_funds(result['pro'])
    result['chpv'] = _aggregate_funds(result['chpv'])
    return jsonify(result)

# ── Fund database CRUD ─────────────────────────────────────────────────────────
@app.route('/api/fund-db', methods=['GET'])
@cache_by_mtime(DARTA_PATH)
def get_fund_db():
    from openpyxl import load_workbook
    path = DARTA_PATH
    result = {'pro': {'headers': [], 'rows': []}, 'chpv': {'headers': [], 'rows': []}}
    if not os.path.exists(path):
        return jsonify(result)
    wb = load_workbook(path, data_only=True, read_only=True)
    for sheet_name, key in [('Challenge Pro', 'pro'), ('Challenge Private', 'chpv')]:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        first = True
        for row in ws.iter_rows(min_row=1, values_only=True):
            if first:
                result[key]['headers'] = [str(h) if h is not None else '' for h in row]
                first = False
            else:
                vals = [str(v) if v is not None else None for v in row]
                # skip blank rows and legend/footnote rows
                non_empty = [v for v in vals if v not in (None, '', 'None')]
                if not non_empty: continue
                if any('LEGENDA' in str(v).upper() or 'COMMISSIONE' in str(v).upper() for v in vals if v): continue
                result[key]['rows'].append(vals)
    wb.close()
    return jsonify(result)

def _sc_decimal(v):
    """SPESE CORRENTI as a numeric decimal (0.0059); '0.59%' → 0.0059; 'N/A' kept as-is."""
    if v is None or isinstance(v, (int, float)):
        return float(v) if isinstance(v, int) else v
    s = str(v).strip().replace(',', '.')
    try:
        return float(s[:-1]) / 100 if s.endswith('%') else float(s)
    except ValueError:
        return v

@app.route('/api/fund-db', methods=['POST'])
def save_fund_db():
    import os, shutil
    from openpyxl import load_workbook
    path = os.path.join(os.path.dirname(__file__), 'DARTA CHP+CHPV updated.xlsx')
    if not os.path.exists(path):
        return jsonify({'error': 'File not found'}), 404
    data = request.json
    shutil.copy2(path, path + '.bak')
    wb = load_workbook(path)
    for sheet_name, key in [('Challenge Pro', 'pro'), ('Challenge Private', 'chpv')]:
        if key not in data or sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        hdr = [str(c.value).lower() if c.value is not None else '' for c in ws[1]]
        i_sc = next((i for i, h in enumerate(hdr) if 'spese correnti' in h), None)
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        for row in data[key].get('rows', []):
            vals = [v if v not in ('None', 'null', '') else None for v in row]
            if i_sc is not None and i_sc < len(vals):
                vals[i_sc] = _sc_decimal(vals[i_sc])
            ws.append(vals)
            if i_sc is not None and isinstance(vals[i_sc] if i_sc < len(vals) else None, float):
                ws.cell(row=ws.max_row, column=i_sc + 1).number_format = '0.00%'
    wb.save(path)
    wb.close()
    return jsonify({'ok': True})

# ── SICAV fund database CRUD ───────────────────────────────────────────────────
SICAV_PATH = os.path.join(os.path.dirname(__file__), 'Lista Fondi Quantalys.xlsx')

@app.route('/api/sicav-fund-db', methods=['GET'])
@cache_by_mtime(lambda: SICAV_PATH)
def get_sicav_fund_db():
    from openpyxl import load_workbook
    result = {'headers': [], 'rows': []}
    if not os.path.exists(SICAV_PATH):
        return jsonify(result)
    wb = load_workbook(SICAV_PATH, data_only=True, read_only=True)
    ws = wb.active
    first = True
    for row in ws.iter_rows(min_row=1, values_only=True):
        if first:
            result['headers'] = [str(h) if h is not None else '' for h in row]
            first = False
        else:
            result['rows'].append([str(v) if v is not None else None for v in row])
    wb.close()
    return jsonify(result)

@app.route('/api/sicav-fund-db', methods=['POST'])
def save_sicav_fund_db():
    import os, shutil
    from openpyxl import load_workbook
    if not os.path.exists(SICAV_PATH):
        return jsonify({'error': 'File not found'}), 404
    data = request.json
    shutil.copy2(SICAV_PATH, SICAV_PATH + '.bak')
    wb = load_workbook(SICAV_PATH)
    ws = wb.active
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    for row in data.get('rows', []):
        ws.append([v if v not in ('None', 'null', '') else None for v in row])
    wb.save(SICAV_PATH)
    wb.close()
    return jsonify({'ok': True})

@app.route('/api/sicav-funds-by-category', methods=['GET'])
@cache_by_mtime(lambda: SICAV_PATH)
def sicav_funds_by_category():
    from openpyxl import load_workbook
    result = {}
    if not os.path.exists(SICAV_PATH):
        return jsonify(result)
    wb = load_workbook(SICAV_PATH, data_only=True, read_only=True)
    ws = wb.active
    hdrs = [str(h).strip().lower() if h else '' for h in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    def fi(kw):
        for i, h in enumerate(hdrs):
            if kw in h: return i
        return None
    i_name = fi('name')
    i_isin = fi('isin')
    i_cat  = fi('categoria quantalys')
    i_fee  = fi('fee')
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or i_cat is None or len(row) <= i_cat: continue
        cat = row[i_cat]
        if not cat: continue
        cat = str(cat).strip()
        nm   = str(row[i_name]).strip() if i_name is not None and len(row) > i_name and row[i_name] else ''
        isin = str(row[i_isin]).strip().upper() if i_isin is not None and len(row) > i_isin and row[i_isin] else ''
        if isin in ('', 'NAN', 'NONE', '0'): isin = ''
        fee = ''
        if i_fee is not None and len(row) > i_fee and row[i_fee] is not None:
            try: fee = f'{float(row[i_fee])*100:.2f}%'
            except: fee = ''
        if not nm: continue
        if cat not in result: result[cat] = []
        result[cat].append({'name': nm, 'isin': isin, 'fee': fee})
    wb.close()
    return jsonify(_aggregate_funds(result))

# ── SICAV: import funds from Manuale Provvigionale PDF ──────────────────────────
# Parses the "Sezione Tabellare" commission manuals (Fondi/Sicav di Terzi e di
# Gruppo), matches by ISIN against the Quantalys Excel, auto-assigns the most
# suited Categoria Quantalys from the existing taxonomy, and (on commit) appends
# new funds + refreshes changed costs. Two-step: /analyze (preview) then /commit.

_CLASS_TOK = set('R RD RDM A AD AX AH AHD AT B BD C CD D E E2 F F2 F3 G GD H HR I ID J N NR P PD Q S T U V W X Y Z T2 M MD K KD ND W2 I2 IT WT PT RT'.split())
_CURR_TOK  = set('EUR USD GBP CHF JPY SEK NOK AUD CAD HKD SGD CNY CNH'.split())

def _sicav_pct_to_dec(s):
    """'1,50%' -> 0.015 ; '48,700%' -> 0.487 ; '' -> None (PDF uses comma decimals)."""
    if s is None: return None
    s = str(s).strip().replace('%', '').replace('.', '').replace(',', '.')
    if not s: return None
    try:
        return round(float(s) / 100.0, 6)
    except ValueError:
        return None

def _sicav_base_norm(name):
    """Normalise a fund name to a share-class/currency-agnostic key for matching."""
    import re
    s = re.sub(r'\(.*?\)', '', name or '').replace('"', ' ').upper()
    for d in ['HARD CLOSURE', 'SOFT CLOSURE', 'HEDGED', 'UNHEDGED', 'ATTIVITA RUSSE', 'ATTIVITÀ RUSSE']:
        s = s.replace(d, ' ')
    s = re.sub(r'[^A-Z0-9 ]', ' ', s)
    drop = _CLASS_TOK | _CURR_TOK | {'HEDGED', 'ACC', 'DIS', 'HR', 'PART', 'TRIMESTRALE', 'ANNUALE'}
    generic = {'FUND', 'FUNDS', 'FONDO', 'FONDI', 'THE', 'OF', 'ISF', 'FF', 'FAST', 'FTGF', 'SICAV', 'PLC'}
    core = [t for t in s.split() if t and t not in drop and t not in generic]
    return ' '.join(core)

def _sicav_top_category(q):
    if q.startswith('Az.'): return 'AZIONARIO'
    if q.startswith('Obblig'): return 'OBBLIGAZIONARIO'
    if q.startswith('Bilanc') or q.startswith('Fless'): return 'BILANCIATI'
    if q.startswith('Monetari'): return 'MONETARI'
    if q.startswith('Perf'): return 'PERFORMANCE ASSOLUTA'
    if q.startswith('Commodities'): return 'COMMODITIES'
    return 'PRODOTTI A GARANZIA O A FORMULA'

def _sicav_classify(name):
    """Best-effort Categoria Quantalys from a fund name, using only taxonomy values."""
    import re
    nm = name or ''
    for pre in ['Franklin Templeton Global Funds', 'FTGF', 'JPMorgan Investment Funds -',
                'JPMorgan Funds -', 'M&G (Lux)', 'MSIF', 'BGF', 'FF -', 'FAST']:
        if nm.startswith(pre): nm = nm[len(pre):]
    n = ' ' + re.sub(r'[^a-z0-9 ]', ' ', nm.lower()) + ' '
    def w(*ws):   return any(re.search(r'\b' + re.escape(x) + r'\b', n) for x in ws)
    def sub(*ws): return any(x in n for x in ws)
    n_nh = re.sub(r'(eur|usd|us dollar|dollar)\s+hedged', '', n)
    def euro(): return re.search(r'\b(euro|eur)\b', n_nh) is not None
    def usd():  return bool(re.search(r'\bus\b', n_nh) or re.search(r'\bu s\b', n_nh) or sub('us dollar', 'dollar', 'united states'))
    is_global = w('global', 'globale', 'world') or sub('globale')
    # --- targeted overrides for tricky brand names ---
    if sub('capital plus', 'kapital plus'): return 'Bilanc. Equil. Global'
    if sub('clean planet'): return 'Az. ambiente'
    if sub('euroland'): return 'Az. Europa Growth' if sub('growth') else 'Az. Area Euro'
    if sub('real return'): return 'Obblig. Global Inflation Linked'
    if sub('managed futures'): return 'Perf. assoluta multi-strategia'
    if re.search(r'long\s*/?\s*short', n):
        return 'Perf. assoluta multi-strategia' if sub('multi asset', 'multi-asset') else 'Perf. ass. Long/Short eq.'
    if sub('impact'): return 'Az. globale'
    if sub('breakout', 'brekout'): return 'Az. paesi emerg. Mondo'
    # --- real estate / commodities / convertibles / money market ---
    if sub('real estate', 'property', 'listed infrastructure', 'infrastructure', ' cities '):
        return 'Az. real estate Europa' if sub('europ') else 'Az. real estate Mondo'
    if sub('commodity', 'commodities'): return 'Commodities'
    if sub('convertible', 'convertib'):
        return 'Obblig. Convertib. Europa' if sub('europ') else 'Obblig. Convertib. Glob.'
    if sub('liquidit', 'money market'):
        return 'Monetari Dollaro USA' if usd() else 'Monetari Euro'
    # --- absolute return ---
    if sub('market neutral'): return 'Perf. assoluta Market Neutral Euro'
    if sub('absolute return', 'liquid alpha', 'alpha capture'):
        if sub('fixed income', 'bond', 'credit', 'tassi', 'rate'): return 'Perf. assoluta tassi'
        if sub('long', 'short', 'equity'): return 'Perf. ass. Long/Short eq.'
        return 'Perf. assoluta multi-strategia'
    # --- balanced / multi-asset / flexible ---
    if sub('multi-asset', 'multi asset', 'asset allocation', 'balanced', 'balanced-risk',
           'multi-strategy', 'multi strategy', 'capitale piu', 'strategy select', 'real assets',
           'target income', 'global themes', 'income builder', 'flexible income', 'step-in income',
           'climate transition multi'):
        if sub('conservative', 'prudent', 'prudenti', ' 15', '12%', ' 12'): return 'Bilanc. Prud. Global'
        if sub('growth', 'aggress', ' 70'): return 'Bilanc. Aggress. Global'
        if sub('flexible', 'flessibil'): return 'Fless. Global'
        return 'Bilanc. Equil. Global'
    # --- bonds ---
    hy = sub('high yield')
    is_bond = sub('bond', 'credit', 'fixed income', 'fixed maturity', 'duration', 'aggregate',
                  'core plus', 'short mid-term', 'government bond', 'green bond', 'green euro credit') \
              or sub('obbligazionar') or hy or ('reddito' in n and 'flessibile' not in n) or sub('riserva')
    if is_bond:
        infl = sub('inflation'); gov = sub('government', 'govern', ' gov '); corp = sub('corporate', 'credit')
        emerg = sub('emerging', 'emergente', ' emerg', 'china', 'asia', 'indonesia', 'thailand', 'latin')
        short = sub('short', 'breve', 'low duration', 'short term', 'short mid', 'mid-term', 'riserva', 'bt')
        if infl: return 'Obblig. Global Inflation Linked'
        if emerg:
            if sub('asia', 'china', 'indonesia', 'thailand'): return 'Obblig. paesi emerg. Asia'
            return 'Obblig. Paesi Emerg.'
        if hy:
            if euro(): return 'Obblig. Euro high yield'
            if usd():  return 'Obblig. Dollaro US high yield'
            return 'Obblig. globale high yield'
        if corp:
            if is_global: return 'Obblig. globale corporate'
            if euro():    return 'Obblig. Euro corporate'
            if usd():     return 'Obblig. Dollaro US all mat'
            return 'Obblig. globale corporate'
        if gov:
            if usd(): return 'Obblig. Dollaro US all mat'
            if euro(): return 'Obblig. Euro gov.'
            return 'Obblig. globale'
        if usd(): return 'Obblig. Dollaro US breve term.' if short else 'Obblig. Dollaro US all mat'
        if is_global: return 'Obblig. globale'
        if euro(): return 'Obblig. Euro breve term.' if short else 'Obblig. Euro all maturities'
        return 'Obblig. globale'
    if sub('flexible', 'flessibil'): return 'Fless. Global'
    # --- equity: sector ---
    if sub('technolog', 'fintech', 'innovation', 'disruptive', 'genomic', 'genetic', 'security', 'semiconduct') or w('ai', 'tech'):
        return 'Az. tecnologia'
    if sub('health', 'pharma'): return 'Az. salute - farmaceutico'
    if sub('biotech'): return 'Az. Biotech'
    if sub('water', 'waste', 'circular', 'timber', 'regeneration', 'clean energy', 'net zero',
           'green materials', 'brown to green', 'smart materials', 'environment', 'ecology', 'climate'):
        return 'Az. ambiente'
    if w('energy') or sub('oil', 'gold', ' oro'): return 'Az. energia materie prime oro'
    if sub('financ', 'banks'): return 'Az. servizi finanziari'
    if sub('consumer', 'consumi'): return 'Az. beni di consumo'
    if sub('utilit'): return 'Az. Servizi di pubblica utilita'
    # --- equity: geography ---
    if sub('china', 'cina', 'grande cina'): return 'Az. Cina'
    if sub('india'): return 'Az. India'
    if sub('japan', 'giappone'): return 'Az. Giappone'
    if sub('indonesia', 'thailand', 'asean'): return 'Az. paesi emerg. Asia'
    if sub('asia pacific', 'asia-pacific', 'pacific'): return 'Az. Asia Pacifico ex Giapp.'
    if sub('latin', 'brazil', 'brasil'): return 'Az. paesi emerg. America Latina'
    if sub('russia', 'russian', 'emerging europe', 'eastern europe'): return 'Az. Paesi Emerg. Europa e Russia'
    if sub('emerging', 'emergente', ' emerg'): return 'Az. paesi emerg. Mondo'
    if sub('italy', 'italia'): return 'Az. Italia'
    if w('uk') or sub('united kingdom'): return 'Az. UK'
    if sub('swiss', 'svizzera'): return 'Az. Svizzera'
    if sub('norden', 'nordic', 'nordici'): return 'Az. Europa'
    if sub('spain', 'spagna'): return 'Az. Spagna'
    if sub('germany', 'germania'): return 'Az. Germania'
    if sub('pan european', 'europe', 'european', 'europa'):
        if sub('growth'): return 'Az. Europa Growth'
        if sub('value', 'yield'): return 'Az. Europa Value'
        if sub('small'): return 'Az. Europa small cap'
        return 'Az. Europa'
    if usd() or sub('american'):
        if sub('growth'): return 'Az. USA Growth'
        if sub('value'): return 'Az. USA Value'
        if sub('small'): return 'Az. USA small cap'
        return 'Az. USA'
    if is_global or sub('alto potenziale', 'for income'):
        if sub('growth'): return 'Az. globale Growth'
        if sub('value'): return 'Az. globale Value'
        if sub('small'): return 'Az. globale small cap'
        return 'Az. globale'
    return 'Az. globale'  # safe default

def _sicav_pdf_extract(file_bytes):
    """Return list of dicts parsed from a Manuale Provvigionale (Sezione Tabellare) PDF."""
    import re, pdfplumber
    out = []
    isin_re = re.compile(r'^[A-Z]{2}[A-Z0-9]{9}[0-9]$')
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                for r in table:
                    if not r or not r[0] or len(r) < 10:
                        continue
                    isin = str(r[0]).replace('\n', ' ').strip()
                    if isin == 'ISIN' or not isin_re.match(isin):
                        continue
                    c = [(str(x).replace('\n', ' ').strip() if x is not None else '') for x in r]
                    valuta = c[6]
                    currency = re.split(r'\s+', valuta)[0].upper() if valuta else ''
                    out.append({
                        'isin': isin, 'societa': c[1], 'name': c[2], 'classe': c[3],
                        'proventi': c[4], 'periodicita': c[5], 'currency': currency,
                        'fee': _sicav_pct_to_dec(c[7]), 'bs_fee': _sicav_pct_to_dec(c[8]),
                        'bs': _sicav_pct_to_dec(c[9]),
                    })
    return out

def _sicav_load_excel_index():
    """Read the Quantalys Excel: header list + per-ISIN record for matching/updates."""
    from openpyxl import load_workbook
    from collections import Counter, defaultdict
    wb = load_workbook(SICAV_PATH, data_only=True, read_only=True)
    ws = wb.active
    hdr = [str(h) if h is not None else '' for h in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    idx = {h: i for i, h in enumerate(hdr)}
    def g(row, name):
        i = idx.get(name)
        return row[i] if i is not None and i < len(row) else None
    by_isin = {}
    base_cat = defaultdict(Counter)   # base-name -> (Category, Q)
    all_q = Counter()
    for row in ws.iter_rows(min_row=2, values_only=True):
        isin = g(row, 'ISIN')
        cat = g(row, 'Category'); q = g(row, 'Categoria Quantalys')
        if isin:
            by_isin[str(isin).strip()] = {
                'category': cat, 'cat_q': q, 'name': g(row, 'Name'),
                'fee': g(row, 'Fee'), 'bs_fee': g(row, 'BS%/Fee'), 'bs': g(row, 'BS%'),
            }
        if q:
            all_q[q] += 1
            if g(row, 'Name'):
                base_cat[_sicav_base_norm(g(row, 'Name'))][(cat, q)] += 1
    wb.close()
    return hdr, by_isin, base_cat, all_q

def _to_float(v):
    try:
        return round(float(v), 6)
    except (TypeError, ValueError):
        return None

@app.route('/api/sicav-import-pdf/analyze', methods=['POST'])
def sicav_import_pdf_analyze():
    """Parse an uploaded PDF; return new funds (auto-categorised) + cost changes. No write."""
    import re
    from collections import Counter, defaultdict
    try:
        import pdfplumber  # noqa: F401
    except ImportError:
        return jsonify({'error': 'Modulo pdfplumber non installato sul server. Esegui: pip install pdfplumber'}), 500
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'Nessun file ricevuto'}), 400
    data = f.read()
    try:
        parsed = _sicav_pdf_extract(data)
    except Exception as e:
        return jsonify({'error': f'Lettura PDF fallita: {e}'}), 400
    if not parsed:
        return jsonify({'error': 'Nessun fondo riconosciuto nel PDF (formato tabella non previsto).'}), 400

    hdr, by_isin, base_cat, all_q = _sicav_load_excel_index()
    existing = set(by_isin)
    # sibling maps built from PDF rows already present in the Excel (highest confidence)
    def norm(s): return re.sub(r'\s+', ' ', re.sub(r'[^A-Z0-9 ]', ' ', (s or '').upper())).strip()
    sib = defaultdict(Counter); pdf_base = defaultdict(Counter)
    for p in parsed:
        rec = by_isin.get(p['isin'])
        if rec and rec['cat_q']:
            key = (norm(p['societa']), norm(p['name']))
            sib[key][(rec['category'], rec['cat_q'])] += 1
            pdf_base[_sicav_base_norm(p['name'])][(rec['category'], rec['cat_q'])] += 1
    def top(cnt): return cnt.most_common(1)[0][0]

    new, updates = [], []
    for p in parsed:
        if p['isin'] in existing:
            rec = by_isin[p['isin']]
            old = {'fee': _to_float(rec['fee']), 'bs_fee': _to_float(rec['bs_fee']), 'bs': _to_float(rec['bs'])}
            neu = {'fee': p['fee'], 'bs_fee': p['bs_fee'], 'bs': p['bs']}
            if any(neu[k] is not None and neu[k] != old[k] for k in neu):
                updates.append({'isin': p['isin'], 'name': p['name'], 'old': old, 'new': neu})
            continue
        key = (norm(p['societa']), norm(p['name'])); bn = _sicav_base_norm(p['name'])
        if sib.get(key):
            cat, q = top(sib[key]); method = 'sibling'
        elif pdf_base.get(bn):
            cat, q = top(pdf_base[bn]); method = 'famiglia'
        elif base_cat.get(bn):
            cat, q = top(base_cat[bn]); method = 'famiglia'
        else:
            q = _sicav_classify(p['name']); cat = _sicav_top_category(q); method = 'auto'
        row = dict(p); row['category'] = cat; row['cat_q'] = q; row['method'] = method
        new.append(row)

    return jsonify({
        'ok': True,
        'categories': sorted(all_q),
        'stats': {'parsed': len(parsed), 'new': len(new), 'updates': len(updates),
                  'existing': len(parsed) - len(new) - len(updates)},
        'new': new, 'updates': updates,
    })

@app.route('/api/sicav-import-pdf/commit', methods=['POST'])
def sicav_import_pdf_commit():
    """Write reviewed rows: append new funds + apply selected cost updates. Backs up first."""
    import os, shutil
    from openpyxl import load_workbook
    if not os.path.exists(SICAV_PATH):
        return jsonify({'error': 'File Excel non trovato'}), 404
    payload = request.json or {}
    new_rows = payload.get('new', [])
    updates = payload.get('updates', [])
    if not new_rows and not updates:
        return jsonify({'error': 'Niente da salvare'}), 400
    shutil.copy2(SICAV_PATH, SICAV_PATH + '.bak')
    wb = load_workbook(SICAV_PATH)
    ws = wb.active
    hdr = [str(c.value) if c.value is not None else '' for c in ws[1]]
    idx = {h: i for i, h in enumerate(hdr)}

    # apply cost updates in-place (match by ISIN)
    i_isin = idx.get('ISIN')
    upd_map = {u['isin']: u for u in updates}
    n_upd = 0
    if i_isin is not None and upd_map:
        for row in ws.iter_rows(min_row=2):
            cell = row[i_isin]
            if cell.value and str(cell.value).strip() in upd_map:
                u = upd_map[str(cell.value).strip()]
                vals = u.get('new') or u   # analyze sends {old,new}; accept flat too
                wrote = False
                for col, key in (('Fee', 'fee'), ('BS%/Fee', 'bs_fee'), ('BS%', 'bs')):
                    ci = idx.get(col)
                    if ci is not None and vals.get(key) is not None:
                        row[ci].value = vals[key]
                        wrote = True
                if wrote:
                    n_upd += 1

    # append new rows
    field_by_header = {
        'Name': 'name', 'ISIN': 'isin', 'Società di Gestione': 'societa',
        'Categoria Quantalys': 'cat_q', 'Classe': 'classe', 'Proventi': 'proventi',
        'Currency': 'currency', 'Fee': 'fee', 'BS%/Fee': 'bs_fee', 'BS%': 'bs',
    }
    for r in new_rows:
        # keep top-level Category consistent with the (possibly edited) Categoria Quantalys
        cat = _sicav_top_category(r['cat_q']) if r.get('cat_q') else r.get('category')
        row = []
        for h in hdr:
            if h == 'Category': row.append(cat)
            elif h in field_by_header: row.append(r.get(field_by_header[h]))
            else: row.append(None)
        ws.append(row)
    wb.save(SICAV_PATH)
    wb.close()
    return jsonify({'ok': True, 'added': len(new_rows), 'updated': n_upd})

# ── DARTA: import costs/category from Challenge Pro/Private circolare PDF ────────
# The "Comunicazione Circolare" letters announce new funds; cost data sits in
# name-keyed tables (Commissioni di Gestione / Overperformance, and Base Spese BA).
# We match by FONDO INTERNO name and fill CG / SPESE CORRENTI / COMM. PERFORMANCE
# (Pro) or SPESE CORRENTI (Private), plus an auto-suggested Categoria Quantalys.

DARTA_PATH = os.path.join(os.path.dirname(__file__), 'DARTA CHP+CHPV updated.xlsx')
DARTA_SHEETS = {'pro': 'Challenge Pro', 'chpv': 'Challenge Private'}

def _darta_norm(s):
    return ' '.join(str(s or '').upper().split())

def _darta_parse_circolare(file_bytes):
    """From a circolare PDF, return {fund_name: {cg, overperf, base_spese}} (raw % strings)."""
    import re, pdfplumber
    cg, over, base = {}, {}, {}
    pct_re = re.compile(r'^-?\d+[.,]?\d*\s*%?$')
    def rows_of(table):
        for r in table:
            cells = [str(c).strip() for c in r if c not in (None, '')]
            if not cells:
                continue
            name, pcts = None, []
            for c in cells:
                if pct_re.match(c) or c.lower().replace('.', '') in ('na', 'n/a'):
                    pcts.append(c)
                elif name is None and len(c) > 4 and not c.isupper():
                    name = c
            if name:
                yield name, pcts
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for pg in pdf.pages:
            for t in pg.extract_tables():
                head = ' '.join(str(c) for c in (t[0] or []) if c).upper()
                if 'GESTIONE' in head or 'OVERPERFORMANCE' in head:
                    for nm, ps in rows_of(t[1:]):
                        if ps:
                            cg[_darta_norm(nm)] = (nm, ps[0])
                            if len(ps) > 1:
                                over[_darta_norm(nm)] = ps[1]
                elif 'BASE SPESE' in head:
                    for nm, ps in rows_of(t[1:]):
                        if ps:
                            base[_darta_norm(nm)] = ps[0]
    out = {}
    for key in set(cg) | set(base):
        disp = cg.get(key, (None, None))[0] or key
        out[key] = {'display': disp,
                    'cg': cg.get(key, (None, None))[1],
                    'overperf': over.get(key),
                    'base_spese': base.get(key)}
    return out

def _darta_norm_perf(v):
    if v is None:
        return None
    return '0%' if v.lower().replace('.', '') in ('na', 'n/a') else v

_DARTA_FAMILY_PREFIX = {'XCLASSIC', 'XTREND', 'XMULTI', 'MKT', 'T&GO', 'SOLUTION', 'MULTI'}

def _darta_derive_am(name):
    """Best-effort Asset Manager from a fund name (strip family prefix, take next token)."""
    toks = str(name or '').split()
    while toks and toks[0].upper().strip('&') in _DARTA_FAMILY_PREFIX:
        toks = toks[1:]
    return toks[0].upper() if toks else ''

def _darta_classify(name):
    """DARTA category suggestion: recognise target-maturity/buy&hold bond funds first,
    else defer to the shared classifier."""
    import re
    n = ' ' + re.sub(r'[^a-z0-9 ]', ' ', str(name or '').lower()) + ' '
    has_year = re.search(r'\b20\d{2}\b', n) is not None
    if ('bond target' in n) or ('target' in n and has_year) \
       or ('buy' in n and 'hold' in n) or ' b h ' in n or 'scadenza' in n:
        return 'Obblig. Euro a scadenza'
    return _sicav_classify(name)

@app.route('/api/darta-import-pdf/analyze', methods=['POST'])
def darta_import_pdf_analyze():
    from openpyxl import load_workbook
    from collections import Counter
    try:
        import pdfplumber  # noqa: F401
    except ImportError:
        return jsonify({'error': 'Modulo pdfplumber non installato sul server. Esegui: pip install pdfplumber'}), 500
    sheet_key = (request.form.get('sheet') or '').lower()
    if sheet_key not in DARTA_SHEETS:
        return jsonify({'error': 'Sheet non valido (usa pro o chpv)'}), 400
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'Nessun file ricevuto'}), 400
    try:
        funds = _darta_parse_circolare(f.read())
    except Exception as e:
        return jsonify({'error': f'Lettura PDF fallita: {e}'}), 400
    if not funds:
        return jsonify({'error': 'Nessuna tabella costi riconosciuta nel PDF.'}), 400

    sheet_name = DARTA_SHEETS[sheet_key]
    wb = load_workbook(DARTA_PATH, data_only=True, read_only=True)
    if sheet_name not in wb.sheetnames:
        return jsonify({'error': f'Foglio "{sheet_name}" non trovato'}), 404
    ws = wb[sheet_name]
    hdr = [str(h) if h is not None else '' for h in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    idx = {h: i for i, h in enumerate(hdr)}
    has_cg = 'CG (%)' in idx
    has_perf = 'COMM. PERFORMANCE' in idx
    by_name, all_q = {}, Counter()
    for row in ws.iter_rows(min_row=2, values_only=True):
        nm = row[idx['FONDO INTERNO']] if idx.get('FONDO INTERNO') is not None else None
        if not nm:
            continue
        q = row[idx['Categoria Quantalys']] if idx.get('Categoria Quantalys') is not None else None
        if q:
            all_q[q] += 1
        by_name[_darta_norm(nm)] = {
            'cg': row[idx['CG (%)']] if has_cg else None,
            'spese': row[idx['SPESE CORRENTI']] if idx.get('SPESE CORRENTI') is not None else None,
            'perf': row[idx['COMM. PERFORMANCE']] if has_perf else None,
            'cat_q': q,
        }
    wb.close()

    def norm_dec(v):
        try:
            return round(float(v), 6)
        except (TypeError, ValueError):
            return None

    matched, new_rows = [], []
    for key, p in funds.items():
        rec = by_name.get(key)
        if not rec:
            nr = {'name': p['display'], 'asset_manager': _darta_derive_am(p['display']),
                  'has_cg': has_cg, 'has_perf': has_perf,
                  'cg': p['cg'] if has_cg else None,
                  'spese': _sicav_pct_to_dec(p['base_spese']) if p['base_spese'] is not None else None,
                  'perf': _darta_norm_perf(p['overperf']) if has_perf else None,
                  'cat_q': _darta_classify(p['display'])}
            new_rows.append(nr)
            continue
        proposed = {'name': p['display'], 'has_cg': has_cg, 'has_perf': has_perf, 'changes': {}}
        # costs
        if has_cg and p['cg']:
            proposed['cg'] = p['cg']
            if str(rec['cg']).strip() != p['cg']:
                proposed['changes']['cg'] = {'old': rec['cg'], 'new': p['cg']}
        if p['base_spese'] is not None:
            new_sp = _sicav_pct_to_dec(p['base_spese'])
            proposed['spese'] = new_sp
            if norm_dec(rec['spese']) != new_sp:
                proposed['changes']['spese'] = {'old': rec['spese'], 'new': new_sp}
        if has_perf and p['overperf'] is not None:
            new_pf = _darta_norm_perf(p['overperf'])
            proposed['perf'] = new_pf
            if str(rec['perf']).strip() != str(new_pf):
                proposed['changes']['perf'] = {'old': rec['perf'], 'new': new_pf}
        # category: keep existing if present, else auto-suggest (never silently overwrite)
        proposed['cat_q'] = rec['cat_q'] or _darta_classify(p['display'])
        proposed['cat_was_empty'] = not rec['cat_q']
        matched.append(proposed)

    matched.sort(key=lambda m: m['name'])
    new_rows.sort(key=lambda m: m['name'])
    return jsonify({
        'ok': True, 'sheet': sheet_key, 'sheet_name': sheet_name,
        'categories': sorted(all_q),
        'stats': {'parsed': len(funds), 'matched': len(matched), 'new': len(new_rows)},
        'matched': matched, 'new': new_rows,
    })

@app.route('/api/darta-import-pdf/commit', methods=['POST'])
def darta_import_pdf_commit():
    import shutil
    from openpyxl import load_workbook
    payload = request.json or {}
    sheet_key = (payload.get('sheet') or '').lower()
    if sheet_key not in DARTA_SHEETS:
        return jsonify({'error': 'Sheet non valido'}), 400
    rows = payload.get('rows', [])          # matched updates (by name)
    new_rows = payload.get('new', [])       # unmatched funds to append
    if not rows and not new_rows:
        return jsonify({'error': 'Niente da salvare'}), 400
    if not os.path.exists(DARTA_PATH):
        return jsonify({'error': 'File DARTA non trovato'}), 404
    shutil.copy2(DARTA_PATH, DARTA_PATH + '.bak')
    wb = load_workbook(DARTA_PATH)
    ws = wb[DARTA_SHEETS[sheet_key]]
    hdr = [str(c.value) if c.value is not None else '' for c in ws[1]]
    idx = {h: i for i, h in enumerate(hdr)}
    i_name = idx.get('FONDO INTERNO')
    col_key = (('CG (%)', 'cg'), ('SPESE CORRENTI', 'spese'),
               ('COMM. PERFORMANCE', 'perf'), ('Categoria Quantalys', 'cat_q'),
               ('ASSET MANAGER', 'asset_manager'), ('FONDO INTERNO', 'name'))
    want = {_darta_norm(r['name']): r for r in rows}
    n_upd = 0
    for row in ws.iter_rows(min_row=2):
        if i_name is None or not row[i_name].value:
            continue
        r = want.get(_darta_norm(row[i_name].value))
        if not r:
            continue
        for col, key in col_key:
            ci = idx.get(col)
            if ci is not None and r.get(key) is not None and key not in ('name',):
                row[ci].value = r[key]
        n_upd += 1
    for r in new_rows:
        line = [None] * len(hdr)
        for col, key in col_key:
            ci = idx.get(col)
            if ci is not None and r.get(key) is not None:
                line[ci] = r[key]
        ws.append(line)
    wb.save(DARTA_PATH)
    wb.close()
    return jsonify({'ok': True, 'updated': n_upd, 'added': len(new_rows)})

# ── Excel export ──────────────────────────────────────────────────────────────
# ── Logo Allianz in alto a sinistra di ogni foglio Excel (riga 1 bianca) ──
XL_LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static', 'img', 'allianz_bank_fa_logo_blue.png')
def _xl_logo_row(ws):
    """Riga 1 bianca con il logo; ritorna la riga da cui parte il titolo."""
    ws.row_dimensions[1].height = 36
    try:
        from openpyxl.drawing.image import Image as XLImage
        from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
        from openpyxl.drawing.xdr import XDRPositiveSize2D
        from openpyxl.utils.units import pixels_to_EMU
        img = XLImage(XL_LOGO)
        h = 28; w = round(h * img.width / img.height)
        img.width, img.height = w, h
        img.anchor = OneCellAnchor(_from=AnchorMarker(col=1, colOff=0, row=0, rowOff=pixels_to_EMU(10)),
                                   ext=XDRPositiveSize2D(pixels_to_EMU(w), pixels_to_EMU(h)))
        ws.add_image(img)
    except Exception:
        pass  # senza Pillow/logo l'export funziona lo stesso, solo senza logo
    return 2

def _xlsx_response(wb, fname):
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    resp = send_file(buf,as_attachment=True,download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    # This export is generated live from the DB. The app sets a global 30-day
    # SEND_FILE_MAX_AGE_DEFAULT for static assets, but that would make the
    # browser cache this dynamic download at the same URL and serve a stale
    # file after the client is edited — so force a fresh fetch every time.
    resp.headers['Cache-Control'] = 'no-store, max-age=0'
    resp.headers.pop('Expires', None)
    return resp

@app.route('/api/clients/<int:cid>/export', methods=['GET'])
def export_excel(cid):
    db = get_db()
    client = db.execute('SELECT name FROM clients WHERE id=?',(cid,)).fetchone()
    if not client: return jsonify({'error':'Not found'}),404
    rows = db.execute('SELECT type,data FROM portfolios WHERE client_id=?',(cid,)).fetchall()
    wb = Workbook()
    wb.remove(wb.active)
    _write_client_sheets(wb, client, rows, request.args.get('overview_sheet','1')=='1', lambda label: label)
    if not wb.sheetnames:
        wb.create_sheet(title='Vuoto')
    return _xlsx_response(wb, f"{client['name'].replace(' ','_')}.xlsx")

# Export di tutti i clienti di un consulente: un foglio per cliente e portafoglio,
# solo le tabelle dei portafogli (niente panoramica).
_XL_SHORT = {'amm':'Amm. ETF','ammAz':'Amm. Azioni','ammEtf':'Amm. ETF','ammStock':'Amm. Azioni','darta':'DARTA','pac':'Easy Switch','sicav':'SICAV','sicavPac':'PAC'}
_XL_LABEL_SHORT = {'Amministrato ETF':'Amm. ETF','Amministrato Azioni':'Amm. Azioni','SICAV & Fondi':'SICAV'}
@app.route('/api/consultants/<int:fid>/export', methods=['GET'])
def export_consultant(fid):
    db = get_db()
    cons = db.execute('SELECT name FROM consultants WHERE id=?',(fid,)).fetchone()
    if not cons: return jsonify({'error':'Not found'}),404
    clients_ = db.execute('SELECT id,name FROM clients WHERE consultant_id=? ORDER BY name COLLATE NOCASE',(fid,)).fetchall()
    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    def title_for(cname):
        def _t(label):
            short = _XL_LABEL_SHORT.get(label, label)
            base = re.sub(r'[\[\]:*?/\\]', '', cname).strip() or 'Cliente'
            room = 31 - len(short) - 3
            t = f"{base[:max(1,room)].rstrip()} · {short}"
            n, cand = 2, t
            while cand.lower() in used:
                suf = f' ({n})'; cand = t[:31-len(suf)] + suf; n += 1
            used.add(cand.lower())
            return cand
        return _t
    for c in clients_:
        rows = db.execute('SELECT type,data FROM portfolios WHERE client_id=?',(c['id'],)).fetchall()
        _write_client_sheets(wb, c, rows, False, title_for(c['name']))
    if not wb.sheetnames:
        wb.create_sheet(title='Nessun cliente')
    return _xlsx_response(wb, f"Clienti_{cons['name'].replace(' ','_')}.xlsx")

def _write_client_sheets(wb, client, rows, overview_sheet, sheet_title):
    """Scrive in wb i fogli dei portafogli di un cliente (stile Allianz)."""
    # Per-portfolio-type option resolver. The frontend sends params namespaced
    # by portfolio type, e.g. "darta_isin=0", "pac_include=1". Anything not sent
    # defaults to enabled, so older/blank requests still export everything.

    def topt(pt, key):
        """Per-type column toggle (default True)."""
        v = request.args.get(f'{pt}_{key}')
        if v is not None:
            return v == '1'
        # back-compat: fall back to a global flag if present
        g = request.args.get(key)
        return (g != '0') if g is not None else True

    def type_included(pt):
        """Whether a whole portfolio type should be exported (default True)."""
        return request.args.get(f'{pt}_include','1') != '0'

    def build_col_spec(pt, is_pac, is_darta):
        """Return ordered list of {'key','hdr','width'} for DARTA/PAC sheets."""
        cols = [{'key':'blank_start','hdr':'','width':3}]
        cols.append({'key':'settore','hdr':'Settore','width':22})
        if topt(pt,'isin'):          cols.append({'key':'isin','hdr':'ISIN','width':16})
        cols.append({'key':'name','hdr':'Nome Fondo','width':34})
        if topt(pt,'ctrv'):          cols.append({'key':'ctrv','hdr':'CTRV','width':14})
        if topt(pt,'pct_ac'):        cols.append({'key':'pct_ac','hdr':'% Classe','width':10})
        if topt(pt,'cg'):            cols.append({'key':'cg','hdr':'CG %','width':9})
        if topt(pt,'sc'):            cols.append({'key':'sc','hdr':'SC %','width':9})
        if topt(pt,'cop'):           cols.append({'key':'cop','hdr':'Cop. %','width':9})
        if topt(pt,'comm_tot'):      cols.append({'key':'comm_tot','hdr':'Comm.Tot %','width':12})
        if topt(pt,'costo'):         cols.append({'key':'costo','hdr':'Costo/m' if is_pac else 'Costo','width':14})
        if is_pac and topt(pt,'costo_tot'): cols.append({'key':'costo_tot','hdr':'Costo Tot.','width':14})
        if topt(pt,'pct_port'):      cols.append({'key':'pct_port','hdr':'% Port.','width':12})
        if topt(pt,'qty'):           cols.append({'key':'qty','hdr':'Quantità','width':10})
        if topt(pt,'pct_settore'):   cols.append({'key':'pct_settore','hdr':'% Settore','width':10})
        cols.append({'key':'blank_end','hdr':'','width':3})
        return cols

    def build_col_spec_amm(pt):
        """Dynamic column spec for Amministrato ETF / Azioni sheets."""
        cols = [{'key':'blank_start','hdr':'','width':3}]
        cols.append({'key':'settore','hdr':'Settore','width':22})
        if topt(pt,'isin'):          cols.append({'key':'isin','hdr':'ISIN','width':16})
        cols.append({'key':'name','hdr':'Nome Fondo','width':34})
        cols.append({'key':'price','hdr':'Prezzo','width':13})
        if topt(pt,'ctrv'):          cols.append({'key':'ctrv','hdr':'CTRV','width':14})
        if topt(pt,'pct_ac'):        cols.append({'key':'pct_ac','hdr':'% Classe','width':10})
        if topt(pt,'ter'):            cols.append({'key':'ter','hdr':'TER %','width':9})
        if topt(pt,'costo'):         cols.append({'key':'costo','hdr':'Costo Ann.','width':14})
        if topt(pt,'pct_port'):      cols.append({'key':'pct_port','hdr':'% Port.','width':12})
        if topt(pt,'qty'):           cols.append({'key':'qty','hdr':'Quantità','width':12})
        if topt(pt,'pct_settore'):   cols.append({'key':'pct_settore','hdr':'% Settore','width':10})
        cols.append({'key':'blank_end','hdr':'','width':3})
        return cols

    def build_col_spec_sicav(pt, is_pac_like=False):
        """Dynamic column spec for SICAV & Fondi (and its PAC) sheets."""
        cols = [{'key':'blank_start','hdr':'','width':3}]
        cols.append({'key':'settore','hdr':'Settore','width':22})
        if topt(pt,'isin'):          cols.append({'key':'isin','hdr':'ISIN','width':16})
        cols.append({'key':'name','hdr':'Nome Fondo','width':34})
        cols.append({'key':'price','hdr':'Prezzo','width':13})
        if topt(pt,'pct_port'):      cols.append({'key':'pct_port','hdr':'% Port.','width':12})
        if topt(pt,'ctrv'):          cols.append({'key':'ctrv','hdr':'CTRV','width':14})
        if topt(pt,'pct_ac'):        cols.append({'key':'pct_ac','hdr':'% Classe','width':10})
        if topt(pt,'cg'):            cols.append({'key':'cg','hdr':'CG %','width':9})
        if topt(pt,'costo'):         cols.append({'key':'costo','hdr':'Costo/m' if is_pac_like else 'Costo','width':14})
        if is_pac_like and topt(pt,'costo_tot'): cols.append({'key':'costo_tot','hdr':'Costo Tot.','width':14})
        if topt(pt,'qty'):           cols.append({'key':'qty','hdr':'Quantità','width':10})
        if topt(pt,'pct_settore'):   cols.append({'key':'pct_settore','hdr':'% Settore','width':10})
        cols.append({'key':'blank_end','hdr':'','width':3})
        return cols

    # Pre-extract DARTA/SICAV data so Easy Switch (pac) and PAC (sicavPac)
    # can inherit plan, total, mesi, and cop_rate base from their parent
    _darta_data = {}
    _sicav_data = {}
    for _r in rows:
        if _r['type'] == 'darta':
            _darta_data = json.loads(_r['data'])
        elif _r['type'] == 'sicav':
            _sicav_data = json.loads(_r['data'])
    darta_total_for_pac = _darta_data.get('total', 0)
    PAC_PARENT_DATA = {'pac': _darta_data, 'sicavPac': _sicav_data}

    # Palette Allianz (come PDF e PowerPoint): blu istituzionale per fasce e totali,
    # bianco su blu per i valori, colori asset class del deck.
    NAVY='003781'; GOLD='0081BE'; GOLD_L='FFFFFF'; WHITE='FFFFFF'
    EQ_P='F0EBF4'; EQ_D='5E3F70'; EQ_M='8A679C'
    BD_P='E5F2F4'; BD_D='005A65'; BD_M='007D8C'
    CM_P='FEF5DA'; CM_D='7A5A00'; CM_M='FAB709'
    SG_P='E6F6FC'; SG_D='00709A'; SG_M='26B5E3'
    BGREY='E7E9EC'; LGREY='F7F7F7'; BORDER='DBDDE1'
    AC_COLORS = {
        'EQUITY':(EQ_P,EQ_D,EQ_M),'BOND':(BD_P,BD_D,BD_M),
        'COMMODITY':(CM_P,CM_D,CM_M),'START & GO':(SG_P,SG_D,SG_M)
    }
    DEF_COLOR = ('F0F0F0','3A3A3A','808080')

    def fl(c): return PatternFill('solid',fgColor=c)
    def fn(bold=False,color='1A1F2E',size=10,italic=False,mono=False):
        return Font(name='Arial',bold=bold,color=color,size=size,italic=italic)
    def al(h='left',v='center'): return Alignment(horizontal=h,vertical=v,wrap_text=False)
    thin = Side(style='thin',color=BORDER)
    def bdr(): return Border(left=thin,right=thin,top=thin,bottom=thin)

    TL = {'amm':'Amministrato ETF','ammAz':'Amministrato Azioni','ammEtf':'Amministrato ETF','ammStock':'Amministrato Azioni','darta':'DARTA','pac':'Easy Switch','sicav':'SICAV & Fondi','sicavPac':'PAC'}

    def parse_pct(s):
        if not s: return 0
        try: return float(str(s).replace(',','.').replace('%','').strip()) / 100
        except: return 0

    def get_copertura(total):
        if not total or total <= 0: return 0
        return 0.0025 if total <= 500000 else 0.0025 * 500000 / total

    # ── Panoramica generale (overview) sheet, placed first ──
    if overview_sheet:
        ov_items = []
        for _r in rows:
            _pt = _r['type']
            if not type_included(_pt):
                continue
            _d = json.loads(_r['data'])
            if _pt in ('ammEtf', 'ammStock'):
                _tot = sum((fr.get('ctrv') or 0) for fr in _d.get('rows', []))
            else:
                _tot = _d.get('total', 0) or 0
            if _tot and _tot > 0:
                ov_items.append((TL.get(_pt, _pt.upper()), _tot))
        if ov_items:
            ovs = wb.create_sheet(title='Panoramica')
            ovs.sheet_view.showGridLines = False
            for _cl, _w in (('A',3),('B',34),('C',20),('D',12),('E',3)):
                ovs.column_dimensions[_cl].width = _w
            rr = _xl_logo_row(ovs)
            ovs.row_dimensions[rr].height = 36
            for col in range(1,6): ovs.cell(rr,col).fill = fl(NAVY)
            ovs.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=5)
            tc = ovs.cell(rr,1,f'PANORAMICA GENERALE — {client["name"].upper()}')
            tc.font = fn(True,WHITE,13); tc.fill = fl(NAVY); tc.alignment = al('center')
            rr += 1
            ovs.row_dimensions[rr].height = 16
            for col in range(1,6): ovs.cell(rr,col).fill = fl(NAVY)
            ovs.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=5)
            sc2 = ovs.cell(rr,1,f'Data: {datetime.now().strftime("%d/%m/%Y")}')
            sc2.font = fn(False,GOLD_L,9,True); sc2.fill = fl(NAVY); sc2.alignment = al('center')
            rr += 2
            for ci,h in ((2,'Portafoglio'),(3,'Totale investito'),(4,'% sul totale')):
                hc = ovs.cell(rr,ci,h); hc.font = fn(True,WHITE,10); hc.fill = fl(NAVY)
                hc.alignment = al('center'); hc.border = bdr()
            rr += 1
            grand = sum(t for _,t in ov_items)
            hdr_row_ov = rr - 1
            data_start_ov = rr
            for i,(nm,tot) in enumerate(ov_items):
                bg = WHITE if (i % 2 == 0) else 'FAFBFC'
                c2 = ovs.cell(rr,2,nm); c2.font = fn(True,'1A1F2E',10); c2.fill = fl(bg); c2.border = bdr(); c2.alignment = al()
                c3 = ovs.cell(rr,3,tot); c3.font = fn(False,'1A1F2E',10); c3.fill = fl(bg); c3.border = bdr(); c3.alignment = al('right'); c3.number_format = '"€"#,##0'
                c4 = ovs.cell(rr,4,(tot/grand) if grand else 0); c4.font = fn(False,'1A1F2E',10); c4.fill = fl(bg); c4.border = bdr(); c4.alignment = al('right'); c4.number_format = '0.0%'
                rr += 1
            data_end_ov = rr - 1
            ovs.row_dimensions[rr].height = 22
            for col in range(1,6): ovs.cell(rr,col).fill = fl(NAVY)
            tcl = ovs.cell(rr,2,'TOTALE'); tcl.font = fn(True,WHITE,11); tcl.fill = fl(NAVY); tcl.alignment = al()
            tcv = ovs.cell(rr,3,grand); tcv.font = fn(True,GOLD_L,12); tcv.fill = fl(NAVY); tcv.alignment = al('right'); tcv.number_format = '"€"#,##0'
            tcp = ovs.cell(rr,4,1); tcp.font = fn(True,GOLD_L,11); tcp.fill = fl(NAVY); tcp.alignment = al('right'); tcp.number_format = '0.0%'

    for row in rows:
        pt = row['type']
        if not type_included(pt):
            continue
        pd_data = json.loads(row['data'])
        label = TL.get(pt,pt.upper())
        ws = wb.create_sheet(title=sheet_title(label))
        ws.sheet_view.showGridLines = False

        is_darta = (pt in ('darta', 'pac'))
        is_pac = (pt == 'pac')
        is_pac_like = (pt in ('pac', 'sicavPac'))
        is_amm_flat = (pt in ('ammEtf', 'ammStock'))
        is_amm = (pt in ('amm', 'ammAz', 'ammEtf', 'ammStock'))

        # ── Flat ammStock / ammEtf export (simple table with formulas) ──────────
        if is_amm_flat:
            flat_rows = pd_data.get('rows', [])
            is_etf_flat = (pt == 'ammEtf')

            # Lookup prices from DB
            _db = get_db()
            price_map = {}
            ytd_map = {}
            ter_map = {}
            name_map = {}
            sector_map = {}
            type_map = {}
            currency_map = {}
            fx_rates = get_fx_rates().get('rates', {})  # EUR→cur, for qty conversion
            if is_etf_flat:
                for _e in _db.execute('SELECT * FROM etf_universe').fetchall():
                    tk = (_e['ticker'] or '').strip()
                    if tk:
                        price_map[tk] = _e['price']
                        ytd_map[tk] = _e['ytd_pct']
                        ter_map[tk] = _e['ter']
                        name_map[tk] = _e['etf_name'] or tk
                        sector_map[tk] = _e['sector'] or ''
                        type_map[tk] = _e['type'] or ''
                        currency_map[tk] = (_e['currency'] or 'EUR').strip().upper()
            else:
                for _s in _db.execute('SELECT * FROM stock_universe').fetchall():
                    comp = (_s['company'] or '').strip()
                    if comp:
                        price_map[comp] = _s['price']
                        ytd_map[comp] = _s['ytd_pct']
                        sector_map[comp] = _s['industry'] or ''
                        name_map[comp] = _s['ticker'] or comp
                        currency_map[comp] = (_s['currency'] or 'EUR').strip().upper()

            # Column definitions
            if is_etf_flat:
                # For ETFs the Ticker column (Excel col A) already holds the
                # descriptive name, so the separate 'ETF Name' column is a
                # duplicate — drop it and widen Ticker to hold the full text.
                flat_cols = [
                    {'key':'ticker',  'hdr':'Ticker',   'width':34},
                    {'key':'sector',  'hdr':'Settore',  'width':20},
                    {'key':'etype',   'hdr':'Tipo',     'width':12},
                    {'key':'price',   'hdr':'Prezzo',   'width':13},
                    {'key':'ctrv',    'hdr':'CTRV (€)', 'width':14},
                    {'key':'ter',     'hdr':'TER %',    'width':9},
                    {'key':'costo',   'hdr':'Costo Ann.','width':14},
                    {'key':'pct_port','hdr':'% Port.',  'width':11},
                    {'key':'qty',     'hdr':'Quantità', 'width':12},
                    {'key':'ytd',     'hdr':'YTD %',    'width':10},
                ]
            else:
                flat_cols = [
                    {'key':'ticker',  'hdr':'Ticker',   'width':12},
                    {'key':'name',    'hdr':'Company',  'width':34},
                    {'key':'sector',  'hdr':'Settore',  'width':20},
                    {'key':'price',   'hdr':'Prezzo',   'width':13},
                    {'key':'ctrv',    'hdr':'CTRV (€)', 'width':14},
                    {'key':'pct_port','hdr':'% Port.',  'width':11},
                    {'key':'qty',     'hdr':'Quantità', 'width':12},
                    {'key':'ytd',     'hdr':'YTD %',    'width':10},
                ]

            # Drop toggleable columns the user disabled for this type
            _flat_optional = {'ctrv','ter','costo','pct_port','qty'}
            flat_cols = [c for c in flat_cols
                         if c['key'] not in _flat_optional or topt(pt, c['key'])]

            NCOLS_F = len(flat_cols) + 2  # +2 for blank margin cols
            BCOL = 1  # blank left
            for i, c in enumerate(flat_cols, 2):
                ws.column_dimensions[get_column_letter(i)].width = c['width']
            ws.column_dimensions['A'].width = 3
            ws.column_dimensions[get_column_letter(NCOLS_F)].width = 3

            flat_col_idx = {c['key']: i+2 for i, c in enumerate(flat_cols)}  # 1-indexed, offset by 1 for left blank

            def _fcl(key):
                ci = flat_col_idx.get(key)
                return get_column_letter(ci) if ci else None

            price_cl  = _fcl('price')
            ctrv_cl   = _fcl('ctrv')
            ter_cl    = _fcl('ter')
            costo_cl  = _fcl('costo')
            pct_cl    = _fcl('pct_port')
            qty_cl    = _fcl('qty')

            # Number formats
            flat_fmts = {
                flat_col_idx.get('price'):    '"€"#,##0.00',
                flat_col_idx.get('ctrv'):     '"€"#,##0',
                flat_col_idx.get('ter'):      '0.000%',
                flat_col_idx.get('costo'):    '"€"#,##0.00',
                flat_col_idx.get('pct_port'): '0.00%',
                flat_col_idx.get('qty'):      '#,##0.000',
                flat_col_idx.get('ytd'):      '0.00%',
            }
            flat_fmts = {k: v for k, v in flat_fmts.items() if k}

            r = _xl_logo_row(ws)
            ws.row_dimensions[r].height = 36
            for col in range(1, NCOLS_F+1): ws.cell(r, col).fill = fl(NAVY)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NCOLS_F)
            tc = ws.cell(r, 1, f'PORTFOLIO REPORT — {client["name"].upper()} — {label.upper()}')
            tc.font = fn(True, WHITE, 13); tc.fill = fl(NAVY); tc.alignment = al('center')
            r += 1

            total_ctrv = sum((fr.get('ctrv') or 0) for fr in flat_rows)

            ws.row_dimensions[r].height = 16
            for col in range(1, NCOLS_F+1): ws.cell(r, col).fill = fl(NAVY)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NCOLS_F)
            sub_text = f'Data: {datetime.now().strftime("%d/%m/%Y")}   |   Totale investito: €{total_ctrv:,.0f}'
            sc2 = ws.cell(r, 1, sub_text)
            sc2.font = fn(False, GOLD_L, 9, True); sc2.fill = fl(NAVY); sc2.alignment = al('center')
            r += 2

            # Column header row
            ws.row_dimensions[r].height = 15
            for col in range(1, NCOLS_F+1):
                ws.cell(r, col).fill = fl(NAVY)
            for i, c in enumerate(flat_cols, 2):
                hc = ws.cell(r, i, c['hdr'])
                hc.font = fn(True, WHITE, 9); hc.fill = fl(NAVY)
                hc.alignment = al('center'); hc.border = bdr()
            r += 1

            data_start = r
            pct_pending_flat = []
            name_ci = flat_col_idx.get('name') or flat_col_idx.get('ticker', 3)

            for fr in flat_rows:
                ticker = (fr.get('ticker') or '').strip()
                company = (fr.get('company') or '').strip()
                lookup_key = ticker if is_etf_flat else company
                ctrv_val = fr.get('ctrv') or 0
                price_val = price_map.get(lookup_key)
                ytd_val   = ytd_map.get(lookup_key)
                ter_val   = fr.get('ter') if fr.get('ter') is not None else ter_map.get(lookup_key)

                ws.row_dimensions[r].height = 15
                bg = WHITE if (r % 2 == 0) else 'FAFBFC'

                def _fv(key):
                    if key == 'ticker':   return ticker
                    if key == 'name':     return name_map.get(lookup_key, company or ticker)
                    if key == 'sector':   return sector_map.get(lookup_key, '')
                    if key == 'etype':    return type_map.get(lookup_key, '')
                    if key == 'price':    return price_val
                    if key == 'ctrv':     return ctrv_val if ctrv_val else None
                    if key == 'ter':      return ter_val
                    if key == 'costo':
                        if ctrv_cl and ter_cl:
                            return f'=IFERROR({ctrv_cl}{r}*{ter_cl}{r},0)'
                        return None
                    if key == 'pct_port': return None  # back-filled
                    if key == 'qty':
                        if ctrv_cl and price_cl:
                            cur = currency_map.get(lookup_key, 'EUR')
                            rate = fx_rates.get(cur, 1) or 1  # EUR→cur (price currency)
                            if abs(rate - 1) > 1e-9:
                                return f'=IFERROR({ctrv_cl}{r}*{rate}/{price_cl}{r},0)'
                            return f'=IFERROR({ctrv_cl}{r}/{price_cl}{r},0)'
                        return None
                    if key == 'ytd':      return ytd_val
                    return ''

                for i, c in enumerate(flat_cols, 2):
                    v = _fv(c['key'])
                    cell = ws.cell(r, i, v)
                    cell.fill = fl(bg); cell.border = bdr()
                    cell.alignment = al('right' if i > name_ci else 'left')
                    if i in flat_fmts:
                        cell.number_format = flat_fmts[i]
                    if c['key'] == 'ticker':
                        cell.font = fn(False, BD_D, 9, mono=True)
                    elif c['key'] == 'name':
                        cell.font = fn(True, '1A1F2E', 9)
                    else:
                        cell.font = fn(False, '1A1F2E' if v not in (None, '') else 'C8CDD8', 9)
                if pct_cl:
                    pct_pending_flat.append(r)
                r += 1

            data_end = r - 1

            # Grand total row
            ws.row_dimensions[r].height = 24
            grand_r_flat = r
            for col in range(1, NCOLS_F+1): ws.cell(r, col).fill = fl(NAVY)
            ws.cell(r, 2, 'TOTALE PORTAFOGLIO').font = fn(True, WHITE, 11)
            ws.cell(r, 2).fill = fl(NAVY); ws.cell(r, 2).alignment = al()
            if ctrv_cl:
                _gt_ctrv = f'=SUM({ctrv_cl}{data_start}:{ctrv_cl}{data_end})'
                gtc = ws.cell(r, flat_col_idx['ctrv'], _gt_ctrv)
                gtc.font = fn(True, GOLD_L, 12); gtc.fill = fl(NAVY)
                gtc.alignment = al('right'); gtc.number_format = '"€"#,##0'
            if is_etf_flat and costo_cl and flat_col_idx.get('costo'):
                _gt_costo = f'=SUM({costo_cl}{data_start}:{costo_cl}{data_end})'
                gcos = ws.cell(r, flat_col_idx['costo'], _gt_costo)
                gcos.font = fn(True, GOLD_L, 11); gcos.fill = fl(NAVY)
                gcos.alignment = al('right'); gcos.number_format = '"€"#,##0.00'
                if ter_cl and ctrv_cl:
                    _avg_ter = f'=IFERROR({costo_cl}{grand_r_flat}/{ctrv_cl}{grand_r_flat},0)'
                    gter = ws.cell(r, flat_col_idx['ter'], _avg_ter)
                    gter.font = fn(True, GOLD_L, 10); gter.fill = fl(NAVY)
                    gter.alignment = al('right'); gter.number_format = '0.000%'
            if pct_cl:
                pct_pending_flat.append(grand_r_flat)
                ws.cell(grand_r_flat, flat_col_idx['pct_port']).fill = fl(NAVY)

            # Back-fill % Port.
            if pct_cl and ctrv_cl:
                for _pr in pct_pending_flat:
                    cell = ws.cell(_pr, flat_col_idx['pct_port'])
                    cell.value = f'=IFERROR({ctrv_cl}{_pr}/{ctrv_cl}{grand_r_flat},0)'
                    cell.number_format = '0.00%'
                    cell.alignment = al('right')
                    cell.font = fn(True if _pr == grand_r_flat else False, GOLD_L if _pr == grand_r_flat else '1A1F2E', 11 if _pr == grand_r_flat else 9)
            r += 2
            continue  # skip rest of loop body for flat types

        if is_darta:
            col_spec = build_col_spec(pt, is_pac, is_darta)
            col_idx = {c['key']: i+1 for i,c in enumerate(col_spec)}
            NCOLS = len(col_spec)
            for i,c in enumerate(col_spec,1):
                ws.column_dimensions[get_column_letter(i)].width = c['width']
        elif is_amm:
            col_spec = build_col_spec_amm(pt)
            col_idx = {c['key']: i+1 for i,c in enumerate(col_spec)}
            NCOLS = len(col_spec)
            for i,c in enumerate(col_spec,1):
                ws.column_dimensions[get_column_letter(i)].width = c['width']
        else:
            col_spec = build_col_spec_sicav(pt, is_pac_like)
            col_idx = {c['key']: i+1 for i,c in enumerate(col_spec)}
            NCOLS = len(col_spec)
            for i,c in enumerate(col_spec,1):
                ws.column_dimensions[get_column_letter(i)].width = c['width']

        _isin_ci = col_idx.get('isin')
        _name_ci = col_idx.get('name', 4)

        def _cl(key, fb=None):
            ci = col_idx.get(key, fb)
            return get_column_letter(ci) if ci else None
        _ctrv_cl      = _cl('ctrv')
        _settore_cl   = get_column_letter(col_idx.get('settore', 2))
        _pct_port_ci  = col_idx.get('pct_port')
        _cg_cl        = _cl('cg');   _sc_cl  = _cl('sc')
        _cop_cl       = _cl('cop');  _comm_tot_cl = _cl('comm_tot')
        _costo_cl     = _cl('costo'); _costo_tot_cl = _cl('costo_tot')
        _price_cl     = _cl('price'); _ter_cl = _cl('ter')
        _num_fmts = {}
        for _k, _fmt in [
            ('ctrv',        '"€"#,##0'),
            ('cg',          '0.00%'),
            ('sc',          '0.00%'),
            ('cop',         '0.000%'),
            ('comm_tot',    '0.000%'),
            ('costo',       '"€"#,##0.00'),
            ('costo_tot',   '"€"#,##0.00'),
            ('pct_port',    '0.00%'),
            ('pct_ac',      '0.00%'),
            ('qty',         '#,##0.000'),
            ('pct_settore', '0.0%'),
            ('price',       '"€"#,##0.00'),
            ('ter',         '0.000%'),
        ]:
            _ci = col_idx.get(_k)
            if _ci: _num_fmts[_ci] = _fmt

        r = _xl_logo_row(ws)
        ws.row_dimensions[r].height = 36
        for col in range(1, NCOLS+1): ws.cell(r, col).fill = fl(NAVY)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NCOLS)
        title_text = f'PORTFOLIO REPORT — {client["name"].upper()} — {label.upper()}'
        tc = ws.cell(r, 1, title_text)
        tc.font = fn(True, WHITE, 13); tc.fill = fl(NAVY); tc.alignment = al('center')
        r += 1

        # Compute effective total. Easy Switch (pac) derives a monthly rata from
        # its DARTA source; PAC (sicavPac) is standalone: monthly rata = total ÷ months.
        mesi = 0
        if is_pac_like:
            if pt == 'pac':
                _parent = _darta_data
                mesi = _parent.get('pac',{}).get('mesi',15) or 15
                pac_cfg = _parent.get('pac',{})
                switch_total = pac_cfg.get('totalSwitch',0) or 0
                if not switch_total:  # backward-compat with legacy START & GO data
                    sg_pct = _parent.get('weights',{}).get('START & GO',{}).get('pct',0) / 100
                    switch_total = _parent.get('total',0) * sg_pct
                pac_actual = switch_total / mesi if mesi else 0
                total = pac_cfg.get('roundedTotal',0) if (pac_cfg.get('useRounded') and pac_cfg.get('roundedTotal',0)>0) else pac_actual
            else:  # sicavPac
                mesi = pd_data.get('mesi',0) or 0
                total = (pd_data.get('total',0) / mesi) if mesi else 0
        else:
            total = pd_data.get('total',0)

        # For darta/sicav: derive effective total from actual ctrvAbs sums
        effective_total = total
        if pt in ('darta', 'sicav') and total > 0:
            _eff_sum = 0
            _wts = pd_data.get('weights', {})
            _frows = pd_data.get('rows', {})
            for _ac, _acd in _wts.items():
                if _ac == 'START & GO': continue
                _ac_pct = _acd.get('pct', 0) / 100
                for _sub, _sp100 in _acd.get('subs', {}).items():
                    _sv = total * _ac_pct * (_sp100 / 100)
                    for _k, _rd in _frows.items():
                        if not _k.startswith(f'{pt}|{_ac}|{_sub}|'): continue
                        _ps = _rd.get('pctSettore') or 0
                        _ab = _rd.get('ctrvAbs')
                        _ctrv = round(_ab) if _ab is not None else (round(_sv * _ps) if _ps else 0)
                        _eff_sum += _ctrv
            if pt in ('darta', 'sicav'):
                _sg_pct = _wts.get('START & GO', {}).get('pct', 0) / 100
                _eff_sum += round(total * _sg_pct)
            if _eff_sum > 0:
                effective_total = _eff_sum

        ws.row_dimensions[r].height = 16
        for col in range(1, NCOLS+1): ws.cell(r, col).fill = fl(NAVY)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NCOLS)
        sub_text = f'Data: {datetime.now().strftime("%d/%m/%Y")}   |   Totale investito: €{effective_total:,.0f}'
        sc = ws.cell(r, 1, sub_text)
        sc.font = fn(False, GOLD_L, 9, True); sc.fill = fl(NAVY); sc.alignment = al('center')
        r += 2

        weights = pd_data.get('weights',{})
        fund_rows = pd_data.get('rows',{})
        chart_labels = []; chart_vals = []
        cop_rate_base = darta_total_for_pac if pt == 'pac' else total
        cop_rate = get_copertura(cop_rate_base) if is_darta else 0
        grand_costo = 0
        wter_num = 0; grand_amm_ctrv = 0; grand_ctrv_sum = 0
        _pct_port_pending = []; _section_total_rows = []

        weight_items = list(weights.items())

        for ac,acd in weight_items:
            if ac == 'START & GO': continue
            ac_pct = acd.get('pct',0)/100
            ac_val = total*ac_pct
            pale,dark,mid = AC_COLORS.get(ac,DEF_COLOR)
            chart_labels.append(ac); chart_vals.append(0)  # filled in after funds are summed

            ws.row_dimensions[r].height = 22
            ws.cell(r,2,f'  {ac}').font = fn(True,dark,12)
            ws.cell(r,2).fill = fl(pale); ws.cell(r,2).alignment = al()
            ws.cell(r,2).border = Border(top=Side(style='medium',color=mid))
            ctrv_col = col_idx.get('ctrv')
            pct_col  = col_idx.get('pct_port')
            if ctrv_col:
                ws.cell(r,ctrv_col,f'€{ac_val:,.0f}').font = fn(True,dark,11)
                ws.cell(r,ctrv_col).fill = fl(pale); ws.cell(r,ctrv_col).alignment = al('right')
            if pct_col:
                ws.cell(r,pct_col,f'{ac_pct*100:.2f}%').font = fn(True,dark,11)
                ws.cell(r,pct_col).fill = fl(pale); ws.cell(r,pct_col).alignment = al('right')
            for col in range(1,NCOLS+1):
                ws.cell(r,col).fill = fl(pale)
                ws.cell(r,col).border = Border(top=Side(style='medium',color=mid))
            ac_hdr_row = r; ac_ctrv_sum = 0; ac_costo_sum = 0
            r += 1

            ws.row_dimensions[r].height = 15
            hdrs = [c['hdr'] for c in col_spec]
            for ci,h in enumerate(hdrs,1):
                c = ws.cell(r,ci,h)
                c.font = fn(True,WHITE,9); c.fill = fl(NAVY)
                c.alignment = al('center'); c.border = bdr()
            r += 1
            _sec_fund_start = r
            _sec_fund_end   = r - 1  # safety: empty section → SUM of zero rows

            subs = list(acd.get('subs',{}).items())
            for sub,sp100 in subs:
                sub_pct = sp100/100; obj_pct = ac_pct*sub_pct; sv = total*obj_pct

                types = (['ETF'] if pt=='amm' else ['Azione']) if is_amm else (['Obbligazione','ETF'] if ac=='BOND' else ['Azione','ETF'])
                sub_ctrv = 0; sub_costo = 0

                # Pre-scan subsector for effective % Settore (uses ctrvAbs when available)
                sub_ctrv_pre = 0
                if is_amm:
                    for k_pre in fund_rows:
                        if not k_pre.startswith(f'{pt}|{ac}|{sub}|'): continue
                        rd_pre = fund_rows.get(k_pre,{})
                        ab_pre = rd_pre.get('ctrvAbs')
                        ps_pre = rd_pre.get('pctSettore',0)
                        sub_ctrv_pre += round(ab_pre) if ab_pre is not None else (round(sv*ps_pre) if ps_pre else 0)

                for ftype in types:
                    type_keys = sorted(
                        [k for k in fund_rows if k.startswith(f'{pt}|{ac}|{sub}|{ftype}|')],
                        key=lambda k: int(k.split('|')[4])
                    )
                    if not is_amm and not type_keys:
                        continue

                    for k in type_keys:
                        rd = fund_rows.get(k,{})
                        isin=rd.get('isin',''); nm=rd.get('name','')
                        price=rd.get('price'); ps=rd.get('pctSettore',0)
                        ab = rd.get('ctrvAbs')
                        ctrv = round(ab) if ab is not None else (round(sv*ps) if ps else 0)
                        if is_amm and not isin and not nm: continue
                        qty=ctrv/price if (price and price>0 and ctrv>0) else None
                        sub_ctrv += ctrv

                        ws.row_dimensions[r].height = 15
                        bg = WHITE if (r%2==0) else 'FAFBFC'

                        if is_darta:
                            cg_s = rd.get('cg',''); sc_s = rd.get('sc','')
                            cg_r = parse_pct(cg_s); sc_r = parse_pct(sc_s)
                            comm_tot = cg_r + sc_r + cop_rate
                            costo = ctrv * comm_tot if (ctrv > 0 and comm_tot > 0) else 0
                            sub_costo += costo; grand_costo += costo
                            # Build vals from col_spec key order
                            def _v(key):
                                if key in ('blank_start','blank_end'): return ''
                                if key == 'settore': return sub
                                if key == 'isin':    return isin if isin else '—'
                                if key == 'name':    return nm
                                if key == 'ctrv':    return ctrv if ctrv else None
                                if key == 'cg':      return cg_r  if cg_r  else None
                                if key == 'sc':      return sc_r  if sc_r  else None
                                if key == 'cop':     return cop_rate if cop_rate else None
                                if key == 'comm_tot':
                                    parts = []
                                    if _cg_cl:  parts.append(f'{_cg_cl}{r}')
                                    if _sc_cl:  parts.append(f'{_sc_cl}{r}')
                                    if _cop_cl: parts.append(f'{_cop_cl}{r}')
                                    return f'=IFERROR({"+".join(parts)},0)' if parts else None
                                if key == 'costo':
                                    if _ctrv_cl and _comm_tot_cl:
                                        return f'=IFERROR({_ctrv_cl}{r}*{_comm_tot_cl}{r},0)'
                                    return None
                                if key == 'costo_tot':
                                    if _costo_cl:
                                        return f'=IFERROR({_costo_cl}{r}*{mesi},0)'
                                    return None
                                if key == 'pct_port': return None  # back-filled after grand total
                                if key == 'qty': return qty if qty else None
                                if key == 'pct_ac':
                                    if _ctrv_cl:
                                        return f'=IFERROR({_ctrv_cl}{r}/{_ctrv_cl}{ac_hdr_row},0)'
                                    return None
                                if key == 'pct_settore':
                                    if _ctrv_cl:
                                        return (f'=IFERROR({_ctrv_cl}{r}/'
                                                f'SUMIF(${_settore_cl}:${_settore_cl},'
                                                f'${_settore_cl}{r},'
                                                f'{_ctrv_cl}:{_ctrv_cl}),0)')
                                    return None
                                return ''
                            vals = [_v(c['key']) for c in col_spec]
                        elif is_amm:
                            ter_s = rd.get('cg','')
                            ter_r = parse_pct(ter_s)
                            costo_amm = ctrv * ter_r if (ctrv > 0 and ter_r > 0) else 0
                            sub_costo += costo_amm; grand_costo += costo_amm
                            if ter_r>0 and ctrv>0: wter_num+=ctrv*ter_r
                            grand_amm_ctrv+=ctrv
                            def _v_amm(key):
                                if key in ('blank_start','blank_end'): return ''
                                if key == 'settore': return sub
                                if key == 'isin':    return isin if isin else '—'
                                if key == 'name':    return nm
                                if key == 'price':   return price if price else None
                                if key == 'ctrv':    return ctrv if ctrv else None
                                if key == 'ter':     return ter_r if ter_r else None
                                if key == 'costo':
                                    if _ctrv_cl and _ter_cl:
                                        return f'=IFERROR({_ctrv_cl}{r}*{_ter_cl}{r},0)'
                                    return None
                                if key == 'pct_port': return None  # back-filled
                                if key == 'qty':
                                    if _ctrv_cl and _price_cl:
                                        return f'=IFERROR({_ctrv_cl}{r}/{_price_cl}{r},0)'
                                    return qty if qty else None
                                if key == 'pct_ac':
                                    if _ctrv_cl:
                                        return f'=IFERROR({_ctrv_cl}{r}/{_ctrv_cl}{ac_hdr_row},0)'
                                    return None
                                if key == 'pct_settore':
                                    if _ctrv_cl:
                                        return (f'=IFERROR({_ctrv_cl}{r}/'
                                                f'SUMIF(${_settore_cl}:${_settore_cl},'
                                                f'${_settore_cl}{r},'
                                                f'{_ctrv_cl}:{_ctrv_cl}),0)')
                                    return None
                                return ''
                            vals = [_v_amm(c['key']) for c in col_spec]
                        else:
                            cg_s = rd.get('cg','')
                            cg_r = parse_pct(cg_s)
                            costo_sicav = ctrv * cg_r if (ctrv > 0 and cg_r > 0) else 0
                            sub_costo += costo_sicav; grand_costo += costo_sicav
                            def _v_sicav(key):
                                if key in ('blank_start','blank_end'): return ''
                                if key == 'settore': return sub
                                if key == 'isin':    return isin if isin else '—'
                                if key == 'name':    return nm
                                if key == 'price':   return price if price else None
                                if key == 'pct_port': return None  # back-filled
                                if key == 'ctrv':    return ctrv if ctrv else None
                                if key == 'cg':      return cg_r if cg_r else None
                                if key == 'costo':
                                    if _ctrv_cl and _cg_cl:
                                        return f'=IFERROR({_ctrv_cl}{r}*{_cg_cl}{r},0)'
                                    return None
                                if key == 'costo_tot':
                                    if _costo_cl:
                                        return f'=IFERROR({_costo_cl}{r}*{mesi},0)'
                                    return None
                                if key == 'qty':
                                    if _ctrv_cl and _price_cl:
                                        return f'=IFERROR({_ctrv_cl}{r}/{_price_cl}{r},0)'
                                    return qty if qty else None
                                if key == 'pct_ac':
                                    if _ctrv_cl:
                                        return f'=IFERROR({_ctrv_cl}{r}/{_ctrv_cl}{ac_hdr_row},0)'
                                    return None
                                if key == 'pct_settore':
                                    if _ctrv_cl:
                                        return (f'=IFERROR({_ctrv_cl}{r}/'
                                                f'SUMIF(${_settore_cl}:${_settore_cl},'
                                                f'${_settore_cl}{r},'
                                                f'{_ctrv_cl}:{_ctrv_cl}),0)')
                                    return None
                                return ''
                            vals = [_v_sicav(c['key']) for c in col_spec]

                        if _pct_port_ci:
                            _pct_port_pending.append(r)
                        for ci, v in enumerate(vals, 1):
                            c = ws.cell(r, ci, v)
                            c.fill = fl(bg); c.border = bdr()
                            c.alignment = al('right' if ci > _name_ci else 'left')
                            if ci in _num_fmts:
                                c.number_format = _num_fmts[ci]
                            if _isin_ci and ci == _isin_ci:
                                c.font = fn(False, BD_D, 9, mono=True)
                            elif ci == _name_ci:
                                c.font = fn(bool(nm), '1A1F2E', 9)
                            elif ci == _pct_port_ci:
                                # pct_port is None now but will be back-filled — use dark font
                                c.font = fn(False, '1A1F2E', 9)
                            else:
                                c.font = fn(False, '1A1F2E' if v not in (None, '—', '') else 'C8CDD8', 9)
                        _sec_fund_end = r
                        r += 1

                ac_ctrv_sum += sub_ctrv
                ac_costo_sum += sub_costo
            # Write actual CTRV sum back to the asset-class header row (formula)
            ac_pct_actual = ac_ctrv_sum / effective_total if effective_total else 0
            if chart_vals: chart_vals[-1] = round(ac_pct_actual * 100, 2)
            # Asset class senza fondi: SUM(E33:E32) includerebbe la cella stessa
            # (riferimento circolare → totale portafoglio e % a zero). Si scrive 0.
            _empty_sec = _sec_fund_end < _sec_fund_start
            def _secsum(cl):
                return 0 if _empty_sec else f'=SUM({cl}{_sec_fund_start}:{cl}{_sec_fund_end})'
            if _ctrv_cl:
                _hdr_ctrv = _secsum(_ctrv_cl)
                ws.cell(ac_hdr_row, ctrv_col, _hdr_ctrv).font = fn(True, dark, 11)
                ws.cell(ac_hdr_row, ctrv_col).fill = fl(pale)
                ws.cell(ac_hdr_row, ctrv_col).alignment = al('right')
                ws.cell(ac_hdr_row, ctrv_col).number_format = '"€"#,##0'
            if _pct_port_ci:
                _pct_port_pending.append(ac_hdr_row)
                ws.cell(ac_hdr_row, _pct_port_ci).fill = fl(pale)
                ws.cell(ac_hdr_row, _pct_port_ci).alignment = al('right')
                ws.cell(ac_hdr_row, _pct_port_ci).number_format = '0.0%'
            comm_tot_ci = col_idx.get('comm_tot')
            if is_darta and _costo_cl and _ctrv_cl and comm_tot_ci:
                _hdr_comm = 0 if _empty_sec else (f'=IFERROR(SUM({_costo_cl}{_sec_fund_start}:{_costo_cl}{_sec_fund_end})'
                             f'/SUM({_ctrv_cl}{_sec_fund_start}:{_ctrv_cl}{_sec_fund_end}),0)')
                ws.cell(ac_hdr_row, comm_tot_ci, _hdr_comm).font = fn(True, dark, 11)
                ws.cell(ac_hdr_row, comm_tot_ci).fill = fl(pale)
                ws.cell(ac_hdr_row, comm_tot_ci).alignment = al('right')
                ws.cell(ac_hdr_row, comm_tot_ci).number_format = '0.000%'
            # Asset-class summary row at bottom of section (SUM formulas)
            ws.row_dimensions[r].height = 18
            ws.cell(r, 2, f'  Totale {ac}').font = fn(True, dark, 11)
            ws.cell(r, 2).fill = fl(pale); ws.cell(r, 2).alignment = al()
            if _ctrv_cl:
                _tot_ctrv = _secsum(_ctrv_cl)
                ws.cell(r, ctrv_col, _tot_ctrv).font = fn(True, dark, 11)
                ws.cell(r, ctrv_col).fill = fl(pale); ws.cell(r, ctrv_col).alignment = al('right')
                ws.cell(r, ctrv_col).number_format = '"€"#,##0'
            _section_total_rows.append(r)
            if _pct_port_ci:
                _pct_port_pending.append(r)
                ws.cell(r, _pct_port_ci).fill = fl(pale)
                ws.cell(r, _pct_port_ci).alignment = al('right')
                ws.cell(r, _pct_port_ci).number_format = '0.0%'
            costo_ci = col_idx.get('costo'); costo_tot_ci = col_idx.get('costo_tot')
            if _costo_cl and costo_ci:
                _tot_costo = _secsum(_costo_cl)
                ws.cell(r, costo_ci, _tot_costo).font = fn(True, '003781', 11)
                ws.cell(r, costo_ci).fill = fl(pale); ws.cell(r, costo_ci).alignment = al('right')
                ws.cell(r, costo_ci).number_format = '"€"#,##0.00'
                if is_pac_like and _costo_tot_cl and costo_tot_ci:
                    _tot_costo_tot = _secsum(_costo_tot_cl)
                    ws.cell(r, costo_tot_ci, _tot_costo_tot).font = fn(True, '003781', 11)
                    ws.cell(r, costo_tot_ci).fill = fl(pale); ws.cell(r, costo_tot_ci).alignment = al('right')
                    ws.cell(r, costo_tot_ci).number_format = '"€"#,##0.00'
            for col in range(1, NCOLS+1):
                ws.cell(r, col).fill = fl(pale)
                ws.cell(r, col).border = Border(bottom=Side(style='medium', color=mid))
            grand_ctrv_sum += ac_ctrv_sum
            r += 2

        if True:  # TOTALE PORTAFOGLIO row is always rendered
            grand_r = r
            ws.row_dimensions[r].height = 24
            ws.cell(r, 2, 'TOTALE PORTAFOGLIO').font = fn(True, WHITE, 11)
            ws.cell(r, 2).fill = fl(NAVY); ws.cell(r, 2).alignment = al()
            ctrv_col = col_idx.get('ctrv')
            if _ctrv_cl and _section_total_rows:
                _grand_ctrv = f'=SUM({",".join([f"{_ctrv_cl}{tr}" for tr in _section_total_rows])})'
                ws.cell(r, ctrv_col, _grand_ctrv).font = fn(True, GOLD_L, 12)
                ws.cell(r, ctrv_col).fill = fl(NAVY); ws.cell(r, ctrv_col).alignment = al('right')
                ws.cell(r, ctrv_col).number_format = '"€"#,##0'
            else:
                ws.cell(r, ctrv_col, f'€{grand_ctrv_sum:,.0f}').font = fn(True, GOLD_L, 12)
                ws.cell(r, ctrv_col).fill = fl(NAVY); ws.cell(r, ctrv_col).alignment = al('right')
            costo_ci    = col_idx.get('costo')
            costo_tot_ci = col_idx.get('costo_tot')
            comm_tot_ci = col_idx.get('comm_tot')
            ter_ci      = col_idx.get('ter')
            if topt(pt,'comm_bar') and _section_total_rows:
                if _costo_cl and costo_ci:
                    _grand_costo = f'=SUM({",".join([f"{_costo_cl}{tr}" for tr in _section_total_rows])})'
                    ws.cell(r, costo_ci, _grand_costo).font = fn(True, GOLD_L, 12)
                    ws.cell(r, costo_ci).fill = fl(NAVY); ws.cell(r, costo_ci).alignment = al('right')
                    ws.cell(r, costo_ci).number_format = '"€"#,##0.00'
                if is_pac_like and _costo_tot_cl and costo_tot_ci:
                    _grand_costo_tot = f'=SUM({",".join([f"{_costo_tot_cl}{tr}" for tr in _section_total_rows])})'
                    ws.cell(r, costo_tot_ci, _grand_costo_tot).font = fn(True, GOLD_L, 12)
                    ws.cell(r, costo_tot_ci).fill = fl(NAVY); ws.cell(r, costo_tot_ci).alignment = al('right')
                    ws.cell(r, costo_tot_ci).number_format = '"€"#,##0.00'
                if comm_tot_ci:
                    # DARTA/PAC: weighted avg commission % label
                    label_ci = comm_tot_ci - 1
                    if label_ci >= 1:
                        ws.cell(r, label_ci, 'COSTO TOT.').font = fn(True, 'CDEAFE', 8)
                        ws.cell(r, label_ci).fill = fl(NAVY); ws.cell(r, label_ci).alignment = al('right')
                    if _costo_cl and _ctrv_cl:
                        _avg_pct = f'=IFERROR({_costo_cl}{grand_r}/{_ctrv_cl}{grand_r},0)'
                        ws.cell(r, comm_tot_ci, _avg_pct).font = fn(True, GOLD_L, 10)
                        ws.cell(r, comm_tot_ci).fill = fl(NAVY); ws.cell(r, comm_tot_ci).alignment = al('right')
                        ws.cell(r, comm_tot_ci).number_format = '0.00%'
                elif is_amm and _ter_cl and _costo_cl and _ctrv_cl:
                    # AMM: total cost sum in costo col, weighted avg TER in TER col
                    _ter_ci = col_idx.get('ter')
                    if _ter_ci:
                        _avg_ter = f'=IFERROR({_costo_cl}{grand_r}/{_ctrv_cl}{grand_r},0)'
                        ws.cell(r, _ter_ci, _avg_ter).font = fn(True, GOLD_L, 10)
                        ws.cell(r, _ter_ci).fill = fl(NAVY); ws.cell(r, _ter_ci).alignment = al('right')
                        ws.cell(r, _ter_ci).number_format = '0.000%'
                    # restore CTRV total (costo col already has cost total from earlier)
                    if ctrv_col and _ctrv_cl:
                        _rctrv = f'=SUM({",".join([f"{_ctrv_cl}{tr}" for tr in _section_total_rows])})'
                        ws.cell(r, ctrv_col, _rctrv).font = fn(True, GOLD_L, 12)
                        ws.cell(r, ctrv_col).fill = fl(NAVY); ws.cell(r, ctrv_col).alignment = al('right')
                        ws.cell(r, ctrv_col).number_format = '"€"#,##0'
                elif not is_darta and not is_amm and _cg_cl and _costo_cl and _ctrv_cl:
                    # SICAV: total cost sum in costo col, weighted avg CG% in CG col
                    _cg_ci = col_idx.get('cg')
                    if _cg_ci:
                        _avg_cg = f'=IFERROR({_costo_cl}{grand_r}/{_ctrv_cl}{grand_r},0)'
                        ws.cell(r, _cg_ci, _avg_cg).font = fn(True, GOLD_L, 10)
                        ws.cell(r, _cg_ci).fill = fl(NAVY); ws.cell(r, _cg_ci).alignment = al('right')
                        ws.cell(r, _cg_ci).number_format = '0.00%'
                    # restore CTRV total
                    if ctrv_col and _ctrv_cl:
                        _rctrv = f'=SUM({",".join([f"{_ctrv_cl}{tr}" for tr in _section_total_rows])})'
                        ws.cell(r, ctrv_col, _rctrv).font = fn(True, GOLD_L, 12)
                        ws.cell(r, ctrv_col).fill = fl(NAVY); ws.cell(r, ctrv_col).alignment = al('right')
                        ws.cell(r, ctrv_col).number_format = '"€"#,##0'
            if is_amm and wter_num > 0 and ter_ci:
                wter_pct = f'{wter_num/grand_amm_ctrv*100:.3f}%' if grand_amm_ctrv else '—'
                ws.cell(r, ter_ci, wter_pct).font = fn(True, GOLD_L, 10)
                ws.cell(r, ter_ci).fill = fl(NAVY); ws.cell(r, ter_ci).alignment = al('right')
            for col in range(1, NCOLS+1): ws.cell(r, col).fill = fl(NAVY)
            r += 1
            # ── Back-fill % Portafoglio for all fund / section / header rows ──
            if _pct_port_ci and _ctrv_cl and _section_total_rows:
                for _pr in _pct_port_pending:
                    ws.cell(_pr, _pct_port_ci, f'=IFERROR({_ctrv_cl}{_pr}/{_ctrv_cl}{grand_r},0)')
                    ws.cell(_pr, _pct_port_ci).number_format = '0.00%'
                    ws.cell(_pr, _pct_port_ci).alignment = al('right')
            if is_amm and wter_num > 0:
                ws.row_dimensions[r].height = 20
                ws.cell(r,2,'COSTO ANNUALE TER').font = fn(True,GOLD_L,10)
                ws.cell(r,2).fill = fl(NAVY); ws.cell(r,2).alignment = al()
                if ter_ci:
                    ws.cell(r,ter_ci,f'€{wter_num:,.2f}').font = fn(True,GOLD_L,11)
                    ws.cell(r,ter_ci).fill = fl(NAVY); ws.cell(r,ter_ci).alignment = al('right')
                for col in range(1,NCOLS+1): ws.cell(r,col).fill = fl(NAVY)
                r += 3
            else:
                r += 2
        else:
            r += 1


EXCEL_UPLOAD_TARGETS = {
    'stock': {'filename': 'Stock + ETF Report.xlsx', 'label': 'Report Azioni & ETF'},
    'darta': {'filename': 'DARTA CHP+CHPV updated.xlsx', 'label': 'Fondi Challenge Pro & Private'},
    'sicav': {'filename': 'Lista Fondi Quantalys.xlsx', 'label': 'Lista SICAV & Fondi (Quantalys)'},
}

@app.route('/api/excel-files', methods=['GET'])
def excel_files_info():
    out = []
    for key, t in EXCEL_UPLOAD_TARGETS.items():
        path = os.path.join(os.path.dirname(__file__), t['filename'])
        updated = None
        if os.path.exists(path):
            updated = datetime.fromtimestamp(os.path.getmtime(path)).strftime('%d/%m/%Y %H:%M')
        out.append({'key': key, 'filename': t['filename'], 'label': t['label'], 'updated': updated})
    return jsonify(out)

@app.route('/api/excel-files/<key>', methods=['POST'])
def upload_excel_file(key):
    if key not in EXCEL_UPLOAD_TARGETS:
        return jsonify({'error': 'Tipo file non valido'}), 400
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'Nessun file ricevuto'}), 400
    data = f.read()
    try:
        load_workbook(io.BytesIO(data), read_only=True).close()
    except Exception:
        return jsonify({'error': 'Il file non è un Excel (.xlsx) valido'}), 400
    path = os.path.join(os.path.dirname(__file__), EXCEL_UPLOAD_TARGETS[key]['filename'])
    tmp = path + '.tmp'
    with open(tmp, 'wb') as out:
        out.write(data)
    os.replace(tmp, path)
    return jsonify({'ok': True})

import os
if __name__=='__main__':
    port=int(os.environ.get('PORT',5001))
    # Debug è DISATTIVO di default: si attiva solo impostando esplicitamente
    # FLASK_DEBUG=1 in locale. In produzione (PythonAnywhere, ecc.) resta off,
    # così non si espongono console interattiva e stack trace.
    debug=os.environ.get('FLASK_DEBUG','') in ('1','true','True')
    app.run(debug=debug,host='0.0.0.0',port=port)
