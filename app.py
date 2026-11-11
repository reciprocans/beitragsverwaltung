import csv
import uuid
import re
import xlsxwriter
from openpyxl import load_workbook
import hmac
import io
import os
import secrets
import sqlite3
from contextlib import contextmanager
from datetime import datetime, date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from functools import wraps
from pathlib import Path

from flask import Flask, abort, flash, g, make_response, redirect, render_template, request, session, url_for, send_file, jsonify
from werkzeug.security import check_password_hash, generate_password_hash

app = Flask(__name__)
app.config.update(
    SECRET_KEY=os.environ.get('SECRET_KEY', ''),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    SESSION_COOKIE_SECURE=os.environ.get('COOKIE_SECURE', '0') == '1',
    PERMANENT_SESSION_LIFETIME=60 * 60 * 12,
    MAX_CONTENT_LENGTH=12 * 1024 * 1024,
)
DATA_DIR = Path(os.environ.get('DATA_DIR', '/data'))
DB_PATH = DATA_DIR / 'verein.db'
DOC_DIR = DATA_DIR / 'documents'
ALLOWED_DOC_EXT = {'.pdf', '.jpg', '.jpeg', '.png', '.docx'}
DOC_KINDS = {'mandate': 'SEPA-Mandat', 'application': 'Mitgliedsantrag', 'other': 'Sonstiges'}
CATEGORIES = {'single': 'Einzelmitgliedschaft', 'support': 'Fördermitglied', 'training': 'In Ausbildung', 'other': 'Sonstiges'}
DEFAULTS = {'single': 'DEFAULT_SINGLE_FEE', 'support': 'DEFAULT_SUPPORT_FEE', 'training': 'DEFAULT_TRAINING_FEE', 'other': 'DEFAULT_OTHER_FEE', 'family': 'DEFAULT_FAMILY_FEE', 'support_family': 'DEFAULT_SUPPORT_FAMILY_FEE'}
DEFAULT_VALUES = {'single': '60', 'support': '30', 'training': '20', 'other': '0', 'family': '100', 'support_family': '100'}


def money_cents(value):
    try:
        number = Decimal(str(value).strip().replace('€', '').replace(' ', '').replace(',', '.'))
        if not number.is_finite() or number < 0 or number > 10000000:
            raise ValueError
        return int((number * 100).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
    except (InvalidOperation, ValueError):
        raise ValueError('Bitte einen gültigen Betrag ab 0 € eingeben.')


def euro(value):
    val = Decimal(int(value or 0)) / 100
    return f'{val:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.') + ' €'


def amount_input(cents):
    return f'{Decimal(cents or 0) / 100:.2f}'.replace('.', ',')

app.jinja_env.filters['euro'] = euro
app.jinja_env.filters['amount_input'] = amount_input


def get_db():
    if 'db' not in g:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        g.db = sqlite3.connect(DB_PATH, timeout=20)
        g.db.row_factory = sqlite3.Row
        g.db.execute('PRAGMA foreign_keys = ON')
        g.db.execute('PRAGMA busy_timeout = 20000')
        g.db.execute('PRAGMA journal_mode = WAL')
    return g.db


@app.teardown_appcontext
def close_db(_error):
    db = g.pop('db', None)
    if db is not None:
        db.close()


def init_db():
    db = get_db()
    db.executescript('''
        CREATE TABLE IF NOT EXISTS families (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          name TEXT NOT NULL,
          address TEXT NOT NULL DEFAULT '',
          fee_cents INTEGER NOT NULL DEFAULT 0 CHECK(fee_cents >= 0),
          active INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS members (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          first_name TEXT NOT NULL,
          last_name TEXT NOT NULL,
          address TEXT NOT NULL DEFAULT '',
          family_id INTEGER REFERENCES families(id) ON DELETE SET NULL,
          category TEXT NOT NULL DEFAULT 'single' CHECK(category IN ('single','support','training','other')),
          fee_cents INTEGER NOT NULL DEFAULT 0 CHECK(fee_cents >= 0),
          active INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS dues (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          year INTEGER NOT NULL,
          family_id INTEGER REFERENCES families(id) ON DELETE SET NULL,
          member_id INTEGER REFERENCES members(id) ON DELETE SET NULL,
          payer_name TEXT NOT NULL,
          category_name TEXT NOT NULL,
          amount_cents INTEGER NOT NULL DEFAULT 0 CHECK(amount_cents >= 0),
          paid INTEGER NOT NULL DEFAULT 0,
          paid_at TEXT,
          UNIQUE(year, family_id),
          UNIQUE(year, member_id),
          CHECK(NOT (family_id IS NOT NULL AND member_id IS NOT NULL))
        );
        CREATE TABLE IF NOT EXISTS documents (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          member_id INTEGER REFERENCES members(id) ON DELETE CASCADE,
          family_id INTEGER REFERENCES families(id) ON DELETE CASCADE,
          kind TEXT NOT NULL CHECK(kind IN ('mandate','application','other')),
          original_name TEXT NOT NULL,
          stored_name TEXT NOT NULL UNIQUE,
          uploaded_at TEXT NOT NULL,
          CHECK ((member_id IS NOT NULL) != (family_id IS NOT NULL))
        );
        CREATE INDEX IF NOT EXISTS idx_documents_member ON documents(member_id);
        CREATE INDEX IF NOT EXISTS idx_documents_family ON documents(family_id);
        CREATE INDEX IF NOT EXISTS idx_dues_year ON dues(year);
        CREATE INDEX IF NOT EXISTS idx_members_family ON members(family_id);
    ''')
    # Additive migrations: existing member, payment and document records remain in place.
    for table, columns in {
        'members': [('street','TEXT NOT NULL DEFAULT \'\''),('postal_code','TEXT NOT NULL DEFAULT \'\''),('city','TEXT NOT NULL DEFAULT \'\''),('phone','TEXT NOT NULL DEFAULT \'\''),('gender','TEXT NOT NULL DEFAULT \'unknown\''),('birth_date','TEXT NOT NULL DEFAULT \'\''),('join_date','TEXT NOT NULL DEFAULT \'\''),('status','TEXT NOT NULL DEFAULT \'active\''),('fee_override','INTEGER NOT NULL DEFAULT 1'),('status_date',"TEXT NOT NULL DEFAULT ''"),('registered_cologne','INTEGER NOT NULL DEFAULT 0')],
        'families': [('street','TEXT NOT NULL DEFAULT \'\''),('postal_code','TEXT NOT NULL DEFAULT \'\''),('city','TEXT NOT NULL DEFAULT \'\''),('phone','TEXT NOT NULL DEFAULT \'\''),('fee_override','INTEGER NOT NULL DEFAULT 1'),('category',"TEXT NOT NULL DEFAULT 'family'")]
    }.items():
        present={r[1] for r in db.execute(f'PRAGMA table_info({table})')}
        for column, definition in columns:
            if column not in present:
                db.execute(f'ALTER TABLE {table} ADD COLUMN {column} {definition}')
        if table=='members':
            db.execute("UPDATE members SET status='inactive' WHERE active=0 AND status='active'")
    # V12: Familienart ohne Datenverlust zu bestehenden Familien hinzufügen.
    db.execute("UPDATE families SET category='family' WHERE category IS NULL OR category NOT IN ('family','support_family')")
    db.execute("""CREATE TABLE IF NOT EXISTS yearly_rates (
          year INTEGER NOT NULL, category TEXT NOT NULL,
          amount_cents INTEGER NOT NULL CHECK(amount_cents>=0),
          PRIMARY KEY(year,category))""")
    db.execute("""CREATE TABLE IF NOT EXISTS app_settings (
        key TEXT PRIMARY KEY, value TEXT NOT NULL
    )""")
    db.execute("INSERT OR IGNORE INTO app_settings(key,value) VALUES('jubilee_years','25,40,50,60,65,70')")
    # V8: actual incoming payment and optional discrepancy note.
    due_columns = {r[1] for r in db.execute('PRAGMA table_info(dues)')}
    if 'received_cents' not in due_columns:
        db.execute('ALTER TABLE dues ADD COLUMN received_cents INTEGER')
    if 'payment_note' not in due_columns:
        db.execute("ALTER TABLE dues ADD COLUMN payment_note TEXT NOT NULL DEFAULT ''")
    # The existing paid records have no separate transfer amount: assume the recorded due.
    db.execute('UPDATE dues SET received_cents=amount_cents WHERE paid=1 AND received_cents IS NULL')
    # V11: Short display name for family dues, including existing merged families.
    # Only the textual category is normalized; amounts, payments and dates stay intact.
    db.execute("UPDATE dues SET category_name='Familienbeitrag' WHERE category_name='Familie (Restbeitrag nach Zusammenführung)'")
    # Deleted payers must not appear as live dues. Preserve their original
    # invoices, including paid flags, in a separately viewable audit archive.
    db.execute("""CREATE TABLE IF NOT EXISTS deleted_dues (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        source_due_id INTEGER NOT NULL UNIQUE,
        year INTEGER NOT NULL,
        payer_name TEXT NOT NULL,
        category_name TEXT NOT NULL,
        amount_cents INTEGER NOT NULL,
        paid INTEGER NOT NULL,
        paid_at TEXT,
        archived_at TEXT NOT NULL,
        archive_reason TEXT NOT NULL
    )""")
    deleted_columns = {r[1] for r in db.execute('PRAGMA table_info(deleted_dues)')}
    if 'received_cents' not in deleted_columns:
        db.execute('ALTER TABLE deleted_dues ADD COLUMN received_cents INTEGER')
    if 'payment_note' not in deleted_columns:
        db.execute("ALTER TABLE deleted_dues ADD COLUMN payment_note TEXT NOT NULL DEFAULT ''")
    db.execute('UPDATE deleted_dues SET received_cents=amount_cents WHERE paid=1 AND received_cents IS NULL')
    # V4/V5 delete operations orphaned invoices by setting both references
    # to NULL. Migrate these old ghost records once, without double counting.
    archive_due_records(db, 'member_id IS NULL AND family_id IS NULL', (),
                        'Aus früherer Version übernommen (gelöschter Beitragszahler)')
    # V15: central role-based user accounts, persistent across restarts.
    db.execute("""CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY,
        role TEXT NOT NULL CHECK(role IN ('admin','editor','reader')),
        password_hash TEXT NOT NULL
    )""")
    for username, role, key in [('admin','admin','ADMIN_PASSWORD'),
                                ('redakteur','editor','EDITOR_PASSWORD'),
                                ('leser','reader','READER_PASSWORD')]:
        if not db.execute('SELECT 1 FROM users WHERE username=?',(username,)).fetchone():
            password = os.environ.get(key, '')
            if password:
                db.execute('INSERT INTO users(username,role,password_hash) VALUES(?,?,?)',
                           (username,role,generate_password_hash(password)))
    db.commit()


def archive_due_records(db, predicate, params, reason):
    """Move dues to an auditable archive; never show deleted payers as due."""
    db.execute(f"""INSERT OR IGNORE INTO deleted_dues
        (source_due_id,year,payer_name,category_name,amount_cents,paid,paid_at,archived_at,archive_reason,received_cents,payment_note)
        SELECT id,year,payer_name,category_name,amount_cents,paid,paid_at,?,?,received_cents,payment_note
        FROM dues WHERE {predicate}""", (datetime.now().isoformat(timespec='seconds'),reason,*params))
    db.execute(f'DELETE FROM dues WHERE {predicate}', params)


def jubilee_thresholds():
    value = get_db().execute("SELECT value FROM app_settings WHERE key='jubilee_years'").fetchone()
    return [int(v) for v in value['value'].split(',')] if value else [25,40,50,60,65,70]


def require_setup():
    key = app.config['SECRET_KEY']
    if not key or key == 'BITTE_ZUFAELLIGEN_LANGEN_SCHLUESSEL_SETZEN':
        raise RuntimeError('SECRET_KEY muss sicher gesetzt sein.')
    # All three accounts must be configured before exposing private data.
    for var in ('ADMIN_PASSWORD','EDITOR_PASSWORD','READER_PASSWORD'):
        if len(os.environ.get(var,'')) < 10:
            raise RuntimeError(f'{var} muss mindestens 10 Zeichen lang sein.')


ADMIN_ONLY = {'admin', 'import_excel', 'backup', 'deleted_dues', 'export', 'export_data'}
EDITOR_POST = {'member_new','member_edit','family_new','family_edit','upload_document','delete_document'}
# Members and families may be created/edited by editors; deletions and finance stay admin-only.


@app.before_request
def before():
    require_setup()
    init_db()
    if request.endpoint in ('login', 'static'):
        return
    if not session.get('authenticated') or session.get('role') not in ('admin','editor','reader'):
        return redirect(url_for('login'))
    role = session['role']
    if request.endpoint in ADMIN_ONLY and role != 'admin':
        abort(403)
    if request.method == 'POST':
        if role == 'reader' or (role == 'editor' and request.endpoint not in EDITOR_POST and request.endpoint != 'logout'):
            abort(403)
        supplied = request.form.get('_csrf', '')
        if not hmac.compare_digest(supplied, session.get('_csrf', '')) or not supplied:
            abort(400, 'Ungültiges Formular-Token. Bitte Seite neu laden.')


@app.context_processor
def inject_globals():
    if session.get('authenticated') and '_csrf' not in session:
        session['_csrf'] = secrets.token_urlsafe(32)
    return dict(csrf_token=session.get('_csrf', ''), now_year=datetime.now().year,
                categories=CATEGORIES, doc_kinds=DOC_KINDS, role=session.get('role'),
                username=session.get('username'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username','').strip().lower()
        password = request.form.get('password','')
        user = get_db().execute('SELECT * FROM users WHERE username=?',(username,)).fetchone()
        if user and check_password_hash(user['password_hash'], password):
            session.clear()
            session.permanent = True
            session['authenticated'] = True
            session['username'] = user['username']
            session['role'] = user['role']
            session['_csrf'] = secrets.token_urlsafe(32)
            return redirect(url_for('dashboard'))
        flash('Benutzername oder Passwort falsch.', 'danger')
    return render_template('login.html')


@app.post('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/users', methods=['GET','POST'])
def users_admin():
    if session.get('role') != 'admin':
        abort(403)
    db=get_db()
    if request.method=='POST':
        username=request.form.get('username','')
        password=request.form.get('new_password','')
        confirm=request.form.get('confirm_password','')
        if username not in ('admin','redakteur','leser'):
            abort(400)
        if len(password)<10 or password != confirm:
            flash('Passwort muss mindestens 10 Zeichen lang sein und zweimal übereinstimmen.','danger')
        else:
            db.execute('UPDATE users SET password_hash=? WHERE username=?',(generate_password_hash(password),username))
            db.commit()
            flash('Passwort für '+username+' geändert.','success')
        return redirect(url_for('users_admin'))
    return render_template('users.html',users=db.execute('SELECT username,role FROM users ORDER BY CASE username WHEN "admin" THEN 0 WHEN "redakteur" THEN 1 ELSE 2 END').fetchall())


def selected_year():
    try:
        year = int(request.values.get('year', datetime.now().year))
    except (ValueError, TypeError):
        abort(400, 'Ungültiges Jahr.')
    if year < 2000 or year > 2100:
        abort(400, 'Ungültiges Jahr.')
    return year


def rate_for_year(category, year):
    row = get_db().execute('SELECT amount_cents FROM yearly_rates WHERE category=? AND year<=? ORDER BY year DESC LIMIT 1', (category,year)).fetchone()
    return row['amount_cents'] if row else money_cents(os.getenv(DEFAULTS[category], DEFAULT_VALUES[category]))


def effective_fee(entity, category, year):
    return entity['fee_cents'] if entity['fee_override'] else rate_for_year(category, year)


def date_input(value):
    if not value:
        return ''
    if isinstance(value, (date, datetime)):
        return value.strftime('%Y-%m-%d')
    v=str(value).strip()
    for fmt in ('%Y-%m-%d','%d.%m.%Y','%d/%m/%Y','%d.%m.%y'):
        try:
            return datetime.strptime(v,fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    return ''


def refresh_due(year, member_id=None, family_id=None):
    """Explicit per-entity edit only; preserve past-year snapshots and payment flags."""
    if year != datetime.now().year:
        return
    db=get_db()
    if family_id:
        fam=db.execute('SELECT * FROM families WHERE id=?',(family_id,)).fetchone()
        if fam:
            db.execute('UPDATE dues SET amount_cents=?,payer_name=?,category_name=? WHERE year=? AND family_id=? AND paid=0', (effective_fee(fam,fam['category'],year),fam['name'],'Förderfamilienbeitrag' if fam['category']=='support_family' else 'Familienbeitrag',year,family_id))
    elif member_id:
        m=db.execute('SELECT * FROM members WHERE id=?',(member_id,)).fetchone()
        if m and (m['family_id'] is not None or m['status']!='active'):
            db.execute('DELETE FROM dues WHERE year=? AND member_id=? AND paid=0',(year,member_id))
        if m and m['family_id'] is None and m['status']=='active':
            db.execute('UPDATE dues SET amount_cents=?,payer_name=?,category_name=? WHERE year=? AND member_id=? AND paid=0', (effective_fee(m,m['category'],year),f"{m['last_name']}, {m['first_name']}",CATEGORIES[m['category']],year,member_id))
    db.commit()


def snapshot(year):
    """Create a dues snapshot for each active payer once per year.
    Existing amounts and payment flags are never silently rewritten.
    """
    db = get_db()
    families = db.execute('SELECT * FROM families WHERE active=1 AND EXISTS (SELECT 1 FROM members WHERE family_id=families.id AND active=1)').fetchall()
    for f in families:
        db.execute('''INSERT OR IGNORE INTO dues(year,family_id,payer_name,category_name,amount_cents)
                      VALUES(?,?,?,?,?)''', (year, f['id'], f['name'], 'Förderfamilienbeitrag' if f['category']=='support_family' else 'Familienbeitrag', effective_fee(f, f['category'], year)))
    singles = db.execute('SELECT * FROM members WHERE family_id IS NULL AND active=1').fetchall()
    for m in singles:
        db.execute('''INSERT OR IGNORE INTO dues(year,member_id,payer_name,category_name,amount_cents)
                      VALUES(?,?,?,?,?)''', (year, m['id'], f"{m['last_name']}, {m['first_name']}", CATEGORIES[m['category']], effective_fee(m, m['category'], year)))
    db.commit()


def cologne_due_flags(db, year):
    """Classify every due exactly once, including family dues (any active Cologne member)."""
    return {r['due_id']: bool(r['registered']) for r in db.execute("""
        SELECT d.id AS due_id,
          CASE WHEN d.family_id IS NOT NULL THEN
              EXISTS(SELECT 1 FROM members m WHERE m.family_id=d.family_id
                     AND m.status='active' AND m.registered_cologne=1)
          ELSE EXISTS(SELECT 1 FROM members m WHERE m.id=d.member_id
                     AND m.status='active' AND m.registered_cologne=1)
          END AS registered
        FROM dues d WHERE d.year=?""", (year,))}


def cologne_due_totals(rows, flags):
    result = {key: {'total':0, 'paid':0, 'missing':0, 'count':0, 'paid_count':0}
              for key in ('registered', 'not_registered')}
    for r in rows:
        item = result['registered' if flags.get(r['id'], False) else 'not_registered']
        item['total'] += r['amount_cents']
        item['count'] += 1
        if r['paid']:
            item['paid'] += r['received_cents'] if r['received_cents'] is not None else r['amount_cents']
            item['paid_count'] += 1
        else:
            item['missing'] += r['amount_cents']
    return result


@app.get('/')
def dashboard():
    year = selected_year()
    if year == datetime.now().year:
        snapshot(year)
    db = get_db()
    archive_years = [r[0] for r in db.execute('SELECT DISTINCT year FROM dues WHERE year < ? ORDER BY year DESC', (datetime.now().year,)).fetchall()]
    search = request.args.get('q', '').strip()
    show = request.args.get('show', 'all')
    if show not in ('all', 'paid', 'open'):
        show = 'all'
    base = db.execute('SELECT * FROM dues WHERE year=? ORDER BY payer_name COLLATE NOCASE', (year,)).fetchall()
    total = sum(x['amount_cents'] for x in base)
    paid = sum(x['received_cents'] if x['received_cents'] is not None else x['amount_cents'] for x in base if x['paid'])
    count_paid = sum(1 for x in base if x['paid'])
    outstanding = sum(x['amount_cents'] for x in base if not x['paid'])
    differences = sum(x['amount_cents'] - (x['received_cents'] if x['received_cents'] is not None else x['amount_cents']) for x in base if x['paid'])
    filtered = [r for r in base if (not search or search.casefold() in r['payer_name'].casefold() or search.casefold() in r['category_name'].casefold()) and (show == 'all' or (show == 'paid' and r['paid']) or (show == 'open' and not r['paid']))]
    # The Cologne flag is stored per member, not per family. A family line is
    # marked if at least one *active* associated member is registered in Cologne.
    # Single-person payments inherit that person's individual flag.
    cologne_for_due = cologne_due_flags(db, year)
    cologne_totals = cologne_due_totals(base, cologne_for_due)
    return render_template('dashboard.html', rows=filtered, cologne_for_due=cologne_for_due,
                           cologne_totals=cologne_totals, year=year, total=total, paid=paid,
                           missing=outstanding, differences=differences, total_count=len(base),
                           paid_count=count_paid, search=search, show=show,
                           archive_years=archive_years, current_year=datetime.now().year)


@app.post('/dues/<int:due_id>/paid')
def toggle_paid(due_id):
    year = selected_year()
    db = get_db()
    due = db.execute('SELECT * FROM dues WHERE id=? AND year=?', (due_id, year)).fetchone()
    if not due:
        abort(404)
    action = request.form.get('action', 'toggle')
    paid = request.form.get('paid') == '1'
    if action not in ('toggle', 'save'):
        abort(400)
    if action == 'save' and not due['paid']:
        abort(400, 'Nur bereits gebuchte Zahlungen können nachträglich bearbeitet werden.')
    if action == 'toggle' and due['paid'] and not paid:
        if request.form.get('confirm_remove') != '1':
            abort(400, 'Zum Entfernen einer Zahlung ist eine Bestätigung erforderlich.')
    if action == 'save':
        paid = True
    note = request.form.get('payment_note', '').strip()
    if len(note) > 2000:
        abort(400, 'Der Zahlungsvermerk ist zu lang (maximal 2000 Zeichen).')
    try:
        actual = money_cents(request.form.get('received_amount', '')) if paid else None
    except ValueError:
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
            return jsonify(error='Bitte einen gültigen Überweisungsbetrag eingeben.'), 400
        flash('Bitte einen gültigen Überweisungsbetrag eingeben.', 'danger')
        return redirect(url_for('dashboard', year=year, q=request.form.get('q',''), show=request.form.get('show','all')))
    if paid:
        paid_at = due['paid_at'] if due['paid'] and due['paid_at'] else date.today().isoformat()
    else:
        paid_at = None
    db.execute('UPDATE dues SET paid=?, paid_at=?, received_cents=?, payment_note=? WHERE id=?',
               (int(paid), paid_at, actual, note, due_id))
    db.commit()
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        snapshot = db.execute('SELECT * FROM dues WHERE year=?',(year,)).fetchall()
        cologne_totals = cologne_due_totals(snapshot, cologne_due_flags(db, year))
        total = sum(r['amount_cents'] for r in snapshot)
        received = sum((r['received_cents'] if r['received_cents'] is not None else r['amount_cents']) for r in snapshot if r['paid'])
        outstanding = sum(r['amount_cents'] for r in snapshot if not r['paid'])
        paid_count = sum(1 for r in snapshot if r['paid'])
        diff = due['amount_cents'] - actual if paid else 0
        return jsonify(ok=True, paid=paid, received_cents=actual, payment_note=note,
          difference_cents=diff, paid_at=paid_at or '', total_cents=total,
          total_count=len(snapshot), paid_cents=received, missing_cents=outstanding,
          paid_count=paid_count, formatted_received=euro(actual) if paid else '–',
          formatted_difference=euro(diff) if paid else '–',
          formatted_total=euro(total), formatted_paid=euro(received),
          formatted_missing=euro(outstanding), progress=round(received/total*100,1) if total else 0,
          cologne_totals=cologne_totals)
    flash('Zahlung und Vermerk gespeichert.' if paid else 'Zahlungsmarkierung nach Bestätigung entfernt.', 'success')
    return redirect(url_for('dashboard', year=year, q=request.form.get('q',''), show=request.form.get('show','all')))


MEMBER_SORTS = {
    'name': 'Nachname / Familie', 'first': 'Vorname', 'city': 'Ort',
    'birth': 'Geburtsdatum', 'join': 'Eintrittsdatum',
    'gender': 'Geschlecht', 'status': 'Mitgliedsstatus'
}
GENDER_VALUES = {'unknown', 'female', 'male', 'diverse'}


def member_list_context():
    # One deterministic grouping/order for listing and next/previous navigation.
    status = request.values.get('status', 'all')
    if status not in ('all', 'active', 'deceased', 'left', 'inactive'):
        status = 'all'
    sort = request.values.get('sort', 'name')
    if sort not in MEMBER_SORTS:
        sort = 'name'
    direction = request.values.get('dir', 'asc')
    if direction not in ('asc', 'desc'):
        direction = 'asc'
    db = get_db()
    records = list(db.execute('''SELECT m.*, f.name AS family_name FROM members m
        LEFT JOIN families f ON f.id=m.family_id
        WHERE (?='all' OR m.status=?)''', (status,status)))
    # Keep the family name row directly above its children, for all sort modes.
    groups = {}
    for r in records:
        group_key = ('f', r['family_id']) if r['family_id'] is not None else ('m', r['id'])
        groups.setdefault(group_key, []).append(r)

    def key_for(r):
        if sort == 'first':
            value = r['first_name']
        elif sort == 'city':
            value = r['city']
        elif sort == 'birth':
            value = r['birth_date']
        elif sort == 'join':
            value = r['join_date']
        elif sort == 'gender':
            value = {'female':'weiblich','male':'männlich','diverse':'divers','unknown':'zzzz'}.get(r['gender'],'zzzz')
        elif sort == 'status':
            value = r['status']
        else:
            value = r['last_name']
        return (str(value or '').casefold(), r['last_name'].casefold(), r['first_name'].casefold(), r['id'])

    reverse = direction == 'desc'
    # Group order for non-name sorts uses the first member in the group by that sort.
    for items in groups.values():
        items.sort(key=key_for, reverse=reverse)
    def group_key(item):
        group_id, group_members = item
        if sort == 'name' and group_id[0] == 'f':
            family = group_members[0]['family_name'] or ''
            return (family.casefold(), 0, group_members[0]['id'])
        return (key_for(group_members[0]), 1, group_members[0]['id']) if sort != 'name' else (key_for(group_members[0])[0], 1, group_members[0]['id'])
    ordered = sorted(groups.items(), key=group_key, reverse=reverse)
    rows = [m for _, persons in ordered for m in persons]
    return rows, status, sort, direction


@app.get('/members')
def members():
    rows, status, sort, direction = member_list_context()
    return render_template('members.html', rows=rows, status=status, sort=sort,
                           direction=direction, sorts=MEMBER_SORTS)


@app.post('/members/bulk-cologne')
def members_bulk_cologne():
    rows, status, sort, direction = member_list_context()
    selected = request.form.getlist('member_ids')
    value = request.form.get('registered_cologne', '')
    back = url_for('members', status=status, sort=sort, dir=direction)
    if value not in ('0', '1'):
        flash('Bitte Köln-Meldestatus auswählen.', 'danger')
        return redirect(back)
    if len(selected) > 5000 or not selected or any(not x.isdecimal() for x in selected):
        flash('Bitte mindestens eine gültige Person auswählen.', 'danger')
        return redirect(back)
    selected_ids = {int(x) for x in selected}
    visible_ids = {r['id'] for r in rows}
    if not selected_ids.issubset(visible_ids):
        abort(400, 'Ungültige Mitgliederauswahl.')
    db = get_db()
    db.executemany('UPDATE members SET registered_cologne=? WHERE id=?', [(int(value), i) for i in selected_ids])
    db.commit()
    flash(f'„In Köln gemeldet“ bei {len(selected_ids)} Person(en) aktualisiert.', 'success')
    return redirect(back)


def get_families():
    return get_db().execute('SELECT * FROM families ORDER BY name COLLATE NOCASE').fetchall()


def default_fee(category):
    return amount_input(money_cents(os.getenv(DEFAULTS[category], DEFAULT_VALUES[category])))


@app.route('/members/new', methods=['GET', 'POST'])
def member_new():
    return member_edit(None)


@app.route('/members/<int:member_id>/edit', methods=['GET','POST'])
def member_edit(member_id):
    db=get_db()
    member=db.execute('SELECT * FROM members WHERE id=?',(member_id,)).fetchone() if member_id else None
    if member_id and not member: abort(404)
    if request.method=='POST':
        first=request.form.get('first_name','').strip()[:150]
        last=request.form.get('last_name','').strip()[:150]
        cat=request.form.get('category','single')
        fid=request.form.get('family_id','').strip()
        family_id=int(fid) if fid.isdigit() else None
        if family_id and not db.execute('SELECT 1 FROM families WHERE id=?',(family_id,)).fetchone(): abort(400)
        status=request.form.get('status','active')
        gender=request.form.get('gender','unknown')
        if status not in ('active','inactive','deceased','left') or gender not in ('female','male','diverse','unknown') or cat not in CATEGORIES or not first or not last:
            flash('Bitte Namen, Status, Geschlecht und Beitragsart prüfen.','danger')
        else:
            try:
                fee=money_cents(request.form.get('fee','0'))
            except ValueError as e:
                flash(str(e),'danger')
            else:
                override=1 if request.form.get('fee_override')=='1' else 0
                cologne=int(request.form.get('registered_cologne')=='1')
                street=request.form.get('street','').strip()[:200]
                postal=request.form.get('postal_code','').strip()[:15]
                city=request.form.get('city','').strip()[:120]
                phone=request.form.get('phone','').strip()[:80]
                birth=date_input(request.form.get('birth_date',''))
                joined=date_input(request.form.get('join_date',''))
                status_date=date_input(request.form.get('status_date','')) if status in ('deceased','left') else ''
                if (request.form.get('birth_date') and not birth) or (request.form.get('join_date') and not joined) or (status in ('deceased','left') and request.form.get('status_date') and not status_date):
                    flash('Geburts- und Eintrittsdatum bitte im gültigen Format eingeben.','danger')
                else:
                    v=(first,last,street,postal,city,phone,gender,birth,joined,status,status_date,family_id,cat,fee,override,int(status=='active'),cologne)
                    if member_id:
                        db.execute("""UPDATE members SET first_name=?,last_name=?,street=?,postal_code=?,city=?,phone=?,gender=?,birth_date=?,join_date=?,status=?,status_date=?,family_id=?,category=?,fee_cents=?,fee_override=?,active=?,registered_cologne=? WHERE id=?""",v+(member_id,))
                    else:
                        cur=db.execute("""INSERT INTO members(first_name,last_name,street,postal_code,city,phone,gender,birth_date,join_date,status,status_date,family_id,category,fee_cents,fee_override,active,registered_cologne) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",v)
                        member_id=cur.lastrowid
                    db.commit()
                    refresh_due(datetime.now().year,member_id=member_id)
                    flash('Mitglied gespeichert. Historische Beiträge bleiben unverändert.','success')
                    return redirect(url_for('member_edit',member_id=member_id, status=request.form.get('nav_status','all'), sort=request.form.get('nav_sort','name'), dir=request.form.get('nav_dir','asc')))
    rows, nav_status, nav_sort, nav_dir = member_list_context()
    ids = [r['id'] for r in rows]
    if member_id and member_id not in ids:
        # A member may have been moved out of a filtered view while editing.
        nav_status = 'all'
        rows = list(db.execute('SELECT id FROM members ORDER BY last_name COLLATE NOCASE,first_name COLLATE NOCASE,id'))
        ids = [r['id'] for r in rows]
    ix = ids.index(member_id) if member_id and member_id in ids else -1
    prev_id = ids[ix-1] if ix>0 else None
    next_id = ids[ix+1] if ix>=0 and ix+1<len(ids) else None
    docs=db.execute('SELECT * FROM documents WHERE member_id=? ORDER BY uploaded_at DESC,id DESC',(member_id,)).fetchall() if member_id else []
    return render_template('member_form.html',member=member,families=get_families(),docs=docs,
                           year_fee=rate_for_year('single',datetime.now().year),rates={k:amount_input(rate_for_year(k,datetime.now().year)) for k in CATEGORIES}, prev_id=prev_id,next_id=next_id,nav_status=nav_status,nav_sort=nav_sort,nav_dir=nav_dir)


@app.post('/members/<int:member_id>/delete')
def delete_member(member_id):
    db = get_db()
    member = db.execute('SELECT * FROM members WHERE id=?', (member_id,)).fetchone()
    if not member:
        abort(404)
    attached = [r[0] for r in db.execute('SELECT stored_name FROM documents WHERE member_id=?', (member_id,))]
    try:
        # Remove *all* personal invoices from the live contribution overview,
        # including those of previous years; keep immutable audit copies.
        archive_due_records(db, 'member_id=?', (member_id,), 'Mitglied gelöscht')
        db.execute('DELETE FROM members WHERE id=?', (member_id,))
        if member['family_id'] is not None:
            # When the last active family member is removed, the family should
            # not keep an outstanding invoice in the live contribution list.
            fid = member['family_id']
            active_count = db.execute('SELECT COUNT(*) FROM members WHERE family_id=? AND active=1', (fid,)).fetchone()[0]
            if active_count == 0:
                archive_due_records(db, 'family_id=?', (fid,), 'Familie ohne aktive Mitglieder')
        db.commit()
    except Exception:
        db.rollback()
        raise
    for name in attached:
        (DOC_DIR / name).unlink(missing_ok=True)
    flash('Mitglied gelöscht. Zugehörige Beiträge wurden aus der Beitragsliste entfernt und im Löscharchiv gesichert.', 'success')
    return redirect(url_for('members'))


@app.get('/families')
def families():
    rows = get_db().execute('''SELECT f.*, COUNT(m.id) AS member_count, GROUP_CONCAT(m.first_name || ' ' || m.last_name, ', ') AS names
                              FROM families f LEFT JOIN members m ON m.family_id=f.id GROUP BY f.id ORDER BY f.name COLLATE NOCASE''').fetchall()
    return render_template('families.html', rows=rows)


@app.route('/families/new', methods=['GET', 'POST'])
def family_new():
    return family_edit(None)


@app.route('/families/<int:family_id>/edit', methods=['GET','POST'])
def family_edit(family_id):
    db=get_db()
    family=db.execute('SELECT * FROM families WHERE id=?',(family_id,)).fetchone() if family_id else None
    if family_id and not family: abort(404)
    if request.method=='POST':
        name=request.form.get('name','').strip()[:200]
        try: fee=money_cents(request.form.get('fee','0'))
        except ValueError as e:
            flash(str(e),'danger')
            return redirect(request.url)
        if not name:
            flash('Familienname fehlt.','danger')
            return redirect(request.url)
        status=int(request.form.get('active')=='1')
        category=request.form.get('category','family')
        if category not in ('family','support_family'): abort(400, 'Ungültige Familienart.')
        override=int(request.form.get('fee_override')=='1')
        v=(name,request.form.get('street','').strip()[:200],request.form.get('postal_code','').strip()[:15],request.form.get('city','').strip()[:120],request.form.get('phone','').strip()[:80],fee,override,status,category)
        # Save family and all newly entered people in one transaction.
        if family_id:
            db.execute('UPDATE families SET name=?,street=?,postal_code=?,city=?,phone=?,fee_cents=?,fee_override=?,active=?,category=? WHERE id=?',v+(family_id,))
        else:
            family_id=db.execute('INSERT INTO families(name,street,postal_code,city,phone,fee_cents,fee_override,active,category) VALUES (?,?,?,?,?,?,?,?,?)',v).lastrowid
        for first,last,gender,birth,joined,cologne in zip(request.form.getlist('person_first[]'),request.form.getlist('person_last[]'),request.form.getlist('person_gender[]'),request.form.getlist('person_birth[]'),request.form.getlist('person_join[]'),request.form.getlist('person_cologne[]')):
            first=first.strip()[:150];last=last.strip()[:150]
            if not first and not last: continue
            if not first or not last:
                db.rollback()
                flash('Bei jeder hinzugefügten Person sind Vor- und Nachname erforderlich.','danger')
                return redirect(url_for('family_edit',family_id=family_id)) if family else redirect(url_for('family_new'))
            gender=gender if gender in ('female','male','diverse','unknown') else 'unknown'
            db.execute("""INSERT INTO members(first_name,last_name,family_id,street,postal_code,city,phone,gender,birth_date,join_date,fee_override,fee_cents,active,status,registered_cologne) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(first,last,family_id,v[1],v[2],v[3],v[4],gender,date_input(birth),date_input(joined),0,rate_for_year('single',datetime.now().year),status,'active' if status else 'inactive',int(cologne=='1')))
        db.commit()
        refresh_due(datetime.now().year,family_id=family_id)
        flash('Familie samt Familienmitgliedern gespeichert. Nur ein Familienbeitrag wird berechnet.','success')
        return redirect(url_for('family_edit',family_id=family_id))
    docs=db.execute('SELECT * FROM documents WHERE family_id=? ORDER BY uploaded_at DESC,id DESC',(family_id,)).fetchall() if family_id else []
    persons=db.execute('SELECT * FROM members WHERE family_id=? ORDER BY last_name,first_name',(family_id,)).fetchall() if family_id else []
    return render_template('family_form.html',family=family,persons=persons,docs=docs,default_fee=amount_input(rate_for_year(family['category'] if family else 'family',datetime.now().year)), family_rates={k:amount_input(rate_for_year(k,datetime.now().year)) for k in ('family','support_family')})


@app.post('/families/<int:family_id>/delete')
def delete_family(family_id):
    db = get_db()
    family=db.execute('SELECT * FROM families WHERE id=?',(family_id,)).fetchone()
    if not family: abort(404)
    # Family members remain independent member files. The current-year family invoice
    # stays as a historical snapshot; newly-independent members get new dues next visit.
    ids=[r[0] for r in db.execute('SELECT id FROM members WHERE family_id=?',(family_id,))]
    year=datetime.now().year
    old=db.execute('SELECT * FROM dues WHERE family_id=? AND year=?',(family_id,year)).fetchone()
    db.execute('UPDATE members SET family_id=NULL WHERE family_id=?',(family_id,))
    # Paid family invoices are preserved; prevent new single dues for the rest
    # of the same year. Unpaid family invoices are replaced by single dues.
    if old and old['paid']:
        for mid in ids:
            db.execute("""INSERT OR IGNORE INTO dues(year,member_id,payer_name,category_name,amount_cents,paid)
                 SELECT ?,id,last_name || ', ' || first_name,'Nach Familienauflösung bereits bezahlt',0,1
                 FROM members WHERE id=?""",(year,mid))
    elif old:
        db.execute('DELETE FROM dues WHERE id=?',(old['id'],))
    db.execute('UPDATE dues SET family_id=NULL WHERE family_id=?',(family_id,))
    attached=[r[0] for r in db.execute('SELECT stored_name FROM documents WHERE family_id=?',(family_id,))]
    db.execute('DELETE FROM families WHERE id=?',(family_id,))
    db.commit()
    for name in attached: (DOC_DIR/name).unlink(missing_ok=True)
    flash(f'Familie gelöscht. {len(ids)} Person(en) bleiben erhalten. Historische Familienbeiträge bleiben dokumentiert.','success')
    return redirect(url_for('families'))


@app.route('/families/merge', methods=['GET','POST'])
def family_merge():
    db=get_db()
    candidates=db.execute("SELECT * FROM members WHERE family_id IS NULL ORDER BY last_name COLLATE NOCASE,first_name COLLATE NOCASE").fetchall()
    if request.method=='POST':
        raw=request.form.getlist('member_ids')
        try: ids=[int(x) for x in raw]
        except ValueError: abort(400)
        if len(set(ids))<2 or len(ids)!=len(set(ids)):
            flash('Bitte mindestens zwei unterschiedliche Personen auswählen.','danger')
            return redirect(url_for('family_merge'))
        persons=db.execute('SELECT * FROM members WHERE id IN ('+','.join('?'*len(ids))+') AND family_id IS NULL',ids).fetchall()
        if len(persons)!=len(ids): abort(400,'Mindestens eine Person ist bereits zugeordnet.')
        addr_id=int(request.form.get('address_source',ids[0]))
        source=next((x for x in persons if x['id']==addr_id and x['id'] in ids),None)
        if source is None: abort(400)
        name=request.form.get('name','').strip()[:200]
        if not name:
            flash('Familienname fehlt.','danger');return redirect(url_for('family_merge'))
        year=datetime.now().year
        category=request.form.get('category','family')
        if category not in ('family','support_family'): abort(400)
        fee=rate_for_year(category,year)
        # Prevent duplicate open invoices. Paid single-person dues remain audit records;
        # the remaining family invoice is reduced by payments made this year.
        try:
            fid=db.execute('''INSERT INTO families(name,street,postal_code,city,phone,fee_cents,fee_override,active,category)
                VALUES(?,?,?,?,?,?,0,1,?)''',(name,source['street'] or source['address'],source['postal_code'],source['city'],source['phone'],fee,category)).lastrowid
            paid_total=0
            for person in persons:
                old=db.execute('SELECT * FROM dues WHERE member_id=? AND year=?',(person['id'],year)).fetchone()
                if old:
                    if old['paid']: paid_total+=old['amount_cents']
                    else: db.execute('DELETE FROM dues WHERE id=?',(old['id'],))
                db.execute('UPDATE members SET family_id=?,street=?,postal_code=?,city=?,phone=? WHERE id=?',
                    (fid,source['street'] or source['address'],source['postal_code'],source['city'],source['phone'],person['id']))
            due=max(0,fee-paid_total)
            db.execute('''INSERT INTO dues(year,family_id,payer_name,category_name,amount_cents,paid)
                VALUES(?,?,?,?,?,?)''',(year,fid,name,'Förderfamilienbeitrag' if category=='support_family' else 'Familienbeitrag',due,int(due==0)))
            db.commit()
        except Exception:
            db.rollback();raise
        flash('Familie erstellt. Adressdaten übernommen, Personen zugeordnet. Offene Einzelbeiträge wurden ersetzt; bereits bezahlte Beiträge angerechnet.','success')
        return redirect(url_for('family_edit',family_id=fid))
    return render_template('family_merge.html',candidates=candidates)


@app.get('/export')
def export():
    year = selected_year()
    if year == datetime.now().year:
        snapshot(year)
    rows = get_db().execute('SELECT * FROM dues WHERE year=? ORDER BY payer_name COLLATE NOCASE', (year,)).fetchall()
    buff = io.StringIO()
    buff.write('\ufeff')
    writer = csv.writer(buff, delimiter=';')
    writer.writerow(['Jahr','Beitragspflichtiger','Beitragsart','Soll EUR','Bezahlt','Zahlungsdatum','Tatsächlicher Eingang EUR','Differenz EUR','Zahlungsvermerk'])
    for x in rows:
        writer.writerow([year, x['payer_name'],x['category_name'],amount_input(x['amount_cents']), 'Ja' if x['paid'] else 'Nein',x['paid_at'] or '', amount_input(x['received_cents']) if x['paid'] else '', amount_input(x['amount_cents'] - (x['received_cents'] if x['received_cents'] is not None else x['amount_cents'])) if x['paid'] else '', x['payment_note'] or ''])
    response = make_response(buff.getvalue())
    response.headers['Content-Type'] = 'text/csv; charset=utf-8'
    response.headers['Content-Disposition'] = f'attachment; filename="beitraege_{year}.csv"'
    return response


@app.get('/backup')
def backup():
    source = get_db()
    out = sqlite3.connect(':memory:')
    source.backup(out)
    temp = DATA_DIR / f'backup_{secrets.token_hex(8)}.sqlite3'
    try:
        disk = sqlite3.connect(temp)
        out.backup(disk)
        disk.close()
        data = temp.read_bytes()
    finally:
        out.close()
        temp.unlink(missing_ok=True)
    resp = make_response(data)
    resp.headers['Content-Type'] = 'application/octet-stream'
    resp.headers['Content-Disposition'] = f'attachment; filename="vereinsbeitraege_backup_{datetime.now():%Y-%m-%d}.sqlite3"'
    return resp



@app.post('/documents/<owner>/<int:owner_id>/upload')
def upload_document(owner, owner_id):
    if owner not in ('members', 'families'):
        abort(404)
    db = get_db()
    table = 'members' if owner == 'members' else 'families'
    column = 'member_id' if owner == 'members' else 'family_id'
    if not db.execute(f'SELECT 1 FROM {table} WHERE id=?', (owner_id,)).fetchone():
        abort(404)
    kind = request.form.get('kind', '')
    file = request.files.get('document')
    if kind not in DOC_KINDS or file is None or not file.filename:
        flash('Bitte Dokument und Dokumentart auswählen.', 'danger')
    else:
        original = file.filename.replace('\\', '/').split('/')[-1][:180]
        extension = Path(original).suffix.lower()
        if extension not in ALLOWED_DOC_EXT:
            flash('Nur PDF, JPG, PNG oder DOCX erlaubt.', 'danger')
        else:
            DOC_DIR.mkdir(parents=True, exist_ok=True)
            name = uuid.uuid4().hex + extension
            filepath = DOC_DIR / name
            try:
                file.save(filepath)
                if filepath.stat().st_size == 0:
                    filepath.unlink(missing_ok=True)
                    flash('Leere Dateien können nicht hochgeladen werden.', 'danger')
                else:
                    db.execute(f'INSERT INTO documents ({column},kind,original_name,stored_name,uploaded_at) VALUES (?,?,?,?,?)',
                               (owner_id, kind, original, name, datetime.now().strftime('%Y-%m-%d %H:%M')))
                    db.commit()
                    flash('Dokument sicher gespeichert.', 'success')
            except Exception:
                filepath.unlink(missing_ok=True)
                raise
    endpoint = 'member_edit' if owner == 'members' else 'family_edit'
    return redirect(url_for(endpoint, **{('member_id' if owner == 'members' else 'family_id'): owner_id}) + '#documents')


@app.get('/documents/<int:doc_id>/download')
def download_document(doc_id):
    doc = get_db().execute('SELECT * FROM documents WHERE id=?', (doc_id,)).fetchone()
    if not doc:
        abort(404)
    path = DOC_DIR / doc['stored_name']
    if not path.is_file():
        abort(404)
    response = send_file(path, as_attachment=True, download_name=doc['original_name'])
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['Cache-Control'] = 'no-store'
    return response


@app.post('/documents/<int:doc_id>/delete')
def delete_document(doc_id):
    db = get_db()
    doc = db.execute('SELECT * FROM documents WHERE id=?', (doc_id,)).fetchone()
    if not doc:
        abort(404)
    db.execute('DELETE FROM documents WHERE id=?', (doc_id,))
    db.commit()
    (DOC_DIR / doc['stored_name']).unlink(missing_ok=True)
    flash('Dokument gelöscht.', 'success')
    if doc['member_id']:
        return redirect(url_for('member_edit', member_id=doc['member_id']) + '#documents')
    return redirect(url_for('family_edit', family_id=doc['family_id']) + '#documents')


def export_dataset(kind, year):
    db = get_db()
    if kind == 'dues':
        if year == datetime.now().year:
            snapshot(year)
        cols = ['Jahr', 'Beitragspflichtiger', 'Beitragsart', 'Soll (EUR)', 'Bezahlt', 'Zahlungsdatum', 'Tatsächlicher Eingang (EUR)', 'Differenz (EUR)', 'Zahlungsvermerk']
        records = db.execute('SELECT * FROM dues WHERE year=? ORDER BY payer_name COLLATE NOCASE', (year,)).fetchall()
        values = [[r['year'], r['payer_name'], r['category_name'], r['amount_cents']/100, 'Ja' if r['paid'] else 'Nein', r['paid_at'] or '', r['received_cents']/100 if r['paid'] and r['received_cents'] is not None else (r['amount_cents']/100 if r['paid'] else None), (r['amount_cents'] - (r['received_cents'] if r['received_cents'] is not None else r['amount_cents']))/100 if r['paid'] else None, r['payment_note'] or ''] for r in records]
    elif kind == 'members':
        cols = ['ID','Vorname','Nachname','Straße','PLZ','Ort','Telefon','Geburtstag','Eintritt','Geschlecht','In Köln gemeldet','Status','Statusdatum','Familie','Beitragsart','Individueller Beitrag (EUR)']
        records = db.execute('SELECT m.*,f.name family_name FROM members m LEFT JOIN families f ON f.id=m.family_id ORDER BY m.last_name,m.first_name').fetchall()
        values = [[r['id'],r['first_name'],r['last_name'],r['street'],r['postal_code'],r['city'],r['phone'],r['birth_date'],r['join_date'],r['gender'],'Ja' if r['registered_cologne'] else 'Nein',r['status'],r['status_date'],r['family_name'] or '',CATEGORIES[r['category']],r['fee_cents']/100 if r['fee_override'] else None] for r in records]
    elif kind == 'families':
        cols = ['ID','Familie','Beitragsart','Straße','PLZ','Ort','Telefon','Individueller Beitrag (EUR)','Aktiv','Mitglieder']
        records = db.execute('SELECT f.*, GROUP_CONCAT(m.first_name || " " || m.last_name, ", ") names FROM families f LEFT JOIN members m ON m.family_id=f.id GROUP BY f.id ORDER BY f.name').fetchall()
        values = [[r['id'],r['name'],'Förderfamilie' if r['category']=='support_family' else 'Familie',r['street'],r['postal_code'],r['city'],r['phone'],r['fee_cents']/100 if r['fee_override'] else None,'Ja' if r['active'] else 'Nein',r['names'] or ''] for r in records]
    else:
        abort(404)
    return cols, values


@app.get('/export/data/<kind>.<fmt>')
def export_data(kind, fmt):
    if fmt not in ('xlsx','csv') or kind not in ('dues','members','families'):
        abort(404)
    year = selected_year()
    cols, values = export_dataset(kind, year)
    filename = f'{kind}_{year}' if kind == 'dues' else kind
    if fmt == 'csv':
        stream = io.StringIO()
        stream.write('\ufeff')
        w = csv.writer(stream, delimiter=';')
        w.writerow(cols)
        for record in values:
            w.writerow([str(v).replace('.', ',') if isinstance(v, float) else v for v in record])
        response = make_response(stream.getvalue())
        response.headers['Content-Type'] = 'text/csv; charset=utf-8'
        response.headers['Content-Disposition'] = f'attachment; filename="{filename}.csv"'
        return response
    stream = io.BytesIO()
    workbook = xlsxwriter.Workbook(stream, {'in_memory': True, 'strings_to_urls': False})
    sheet = workbook.add_worksheet('Beiträge' if kind == 'dues' else 'Mitglieder' if kind == 'members' else 'Familien')
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#243D6B', 'font_color': '#FFFFFF', 'border': 0, 'valign': 'vcenter'})
    currency_fmt = workbook.add_format({'num_format': '#,##0.00 "€"'})
    sheet.freeze_panes(1, 0)
    sheet.set_row(0, 26)
    sheet.set_column(0, len(cols)-1, 21)
    for j, name in enumerate(cols):
        sheet.write(0, j, name, header_fmt)
        if name in ('Adresse','Mitglieder','Beitragspflichtiger'):
            sheet.set_column(j,j,34)
        if 'EUR' in name:
            sheet.set_column(j,j,23)
    for i, record in enumerate(values,1):
        for j, value in enumerate(record):
            if isinstance(value, str) and value.startswith(('=', '+', '-', '@')):
                value = "'" + value
            sheet.write(i,j,value, currency_fmt if 'EUR' in cols[j] and isinstance(value,(int,float)) else None)
    if values:
        sheet.autofilter(0,0,len(values),len(cols)-1)
    workbook.close()
    stream.seek(0)
    return send_file(stream, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f'{filename}.xlsx')


@app.get('/archive')
def archive():
    rows = get_db().execute("""SELECT year, COUNT(*) total_count, COALESCE(SUM(amount_cents),0) total,
                             COALESCE(SUM(CASE WHEN paid=1 THEN amount_cents ELSE 0 END),0) paid
                             FROM dues WHERE year < ? GROUP BY year ORDER BY year DESC""", (datetime.now().year,)).fetchall()
    return render_template('archive.html', rows=rows)

@app.get('/deleted-dues')
def deleted_dues():
    db=get_db()
    rows=db.execute('SELECT * FROM deleted_dues ORDER BY archived_at DESC, year DESC, id DESC').fetchall()
    return render_template('deleted_dues.html', rows=rows)


@app.route('/admin',methods=['GET','POST'])
def admin():
    db=get_db()
    yr=selected_year()
    if request.method=='POST' and request.form.get('action')=='jubilees':
        raw=request.form.get('jubilee_years','')
        try:
            values=[int(x.strip()) for x in raw.split(',') if x.strip()]
            if not values or len(values)>30 or any(x<1 or x>120 for x in values): raise ValueError
            values=sorted(set(values))
        except ValueError:
            flash('Jubiläumsjahre bitte als Kommaliste mit Werten von 1 bis 120 eingeben.','danger')
        else:
            db.execute("UPDATE app_settings SET value=? WHERE key='jubilee_years'",(','.join(map(str,values)),))
            db.commit();flash('Jubiläumsstufen gespeichert.','success')
        return redirect(url_for('admin',year=yr))
    if request.method=='POST':
        try:
            for cat in DEFAULTS:
                amount=money_cents(request.form.get('rate_'+cat,''))
                db.execute('INSERT INTO yearly_rates(year,category,amount_cents) VALUES(?,?,?) ON CONFLICT(year,category) DO UPDATE SET amount_cents=excluded.amount_cents',(yr,cat,amount))
            # Übernahme in existierende, noch offene Jahresforderungen: Bei
            # individuellen Abweichungen sowie bezahlten Buchungen nichts ändern.
            db.execute("""UPDATE dues SET amount_cents=(
                  SELECT r.amount_cents FROM members m
                  JOIN yearly_rates r ON r.year=dues.year AND r.category=m.category
                  WHERE m.id=dues.member_id AND m.fee_override=0
                )
                WHERE year=? AND paid=0 AND member_id IS NOT NULL
                AND EXISTS (SELECT 1 FROM members m WHERE m.id=dues.member_id AND m.fee_override=0)
                AND EXISTS (SELECT 1 FROM members m JOIN yearly_rates r ON r.year=dues.year AND r.category=m.category WHERE m.id=dues.member_id)""",(yr,))
            # V12: offene Familienforderungen nach der jeweiligen Familienart
            # aktualisieren. Individuelle Beiträge und bezahlte Buchungen bleiben unverändert.
            db.execute("""UPDATE dues SET
                amount_cents=(SELECT r.amount_cents FROM families f JOIN yearly_rates r
                    ON r.year=dues.year AND r.category=f.category
                    WHERE f.id=dues.family_id),
                category_name=(SELECT CASE f.category WHEN 'support_family'
                    THEN 'Förderfamilienbeitrag' ELSE 'Familienbeitrag' END
                    FROM families f WHERE f.id=dues.family_id)
                WHERE year=? AND paid=0 AND family_id IS NOT NULL
                AND EXISTS (SELECT 1 FROM families f JOIN yearly_rates r
                    ON r.year=dues.year AND r.category=f.category
                    WHERE f.id=dues.family_id AND f.fee_override=0)""",(yr,))
            db.commit()
            flash(f'Beitragssätze für {yr} gespeichert und in offene Jahresbeiträge ohne individuelle Abweichung übernommen. Bereits bezahlte Buchungen bleiben unverändert.','success')
        except ValueError as e:
            db.rollback(); flash(str(e),'danger')
        return redirect(url_for('admin',year=yr))
    rates={cat:amount_input(rate_for_year(cat,yr)) for cat in DEFAULTS}
    years=[r[0] for r in db.execute('SELECT DISTINCT year FROM yearly_rates ORDER BY year DESC')]
    return render_template('admin.html',year=yr,rates=rates,years=years,jubilee_years=','.join(map(str,jubilee_thresholds())))


@app.route('/import',methods=['GET','POST'])
def import_excel():
    if request.method=='POST':
        uploaded=request.files.get('file')
        if not uploaded or not uploaded.filename.lower().endswith('.xlsx'):
            flash('Bitte eine .xlsx-Datei auswählen.','danger')
            return redirect(url_for('import_excel'))
        try:
            wb=load_workbook(uploaded,read_only=True,data_only=True)
            ws=wb.active
            iterator=ws.iter_rows(values_only=True)
            raw=next(iterator,None)
            if not raw: raise ValueError('Leere Datei.')
            headers={str(x or '').strip().lower():i for i,x in enumerate(raw)}
            for required in ['name','vorname','strasse','plz','ort']:
                if required not in headers: raise ValueError(f'Pflichtspalte fehlt: {required}')
            db=get_db(); added=0; skipped=0
            for row in iterator:
                def cell(k):
                    index=headers.get(k)
                    return row[index] if index is not None and index<len(row) else None
                last=str(cell('name') or '').strip()[:150]
                first=str(cell('vorname') or '').strip()[:150]
                if not first or not last: continue
                born=date_input(cell('geburtstag'))
                entered=date_input(cell('eintritt'))
                # Skip likely duplicates on repeated import; birth date distinguishes namesakes.
                if db.execute('SELECT 1 FROM members WHERE LOWER(first_name)=LOWER(?) AND LOWER(last_name)=LOWER(?) AND birth_date=?',(first,last,born)).fetchone():
                    skipped+=1;continue
                street=str(cell('strasse') or '').strip()[:200]
                postal=str(cell('plz') or '').strip()[:15]
                if postal.endswith('.0'): postal=postal[:-2]
                city=str(cell('ort') or '').strip()[:120]
                db.execute("""INSERT INTO members(first_name,last_name,street,postal_code,city,birth_date,join_date,category,fee_cents,fee_override,status,active) VALUES(?,?,?,?,?,?,?,?,?,?,?,1)""",(first,last,street,postal,city,born,entered,'single',rate_for_year('single',datetime.now().year),0,'active'))
                added+=1
            db.commit();wb.close()
            flash(f'Excel-Import abgeschlossen: {added} neu, {skipped} als Dubletten übersprungen. Familie und Geschlecht kannst du danach ergänzen.','success')
        except Exception as exc:
            flash(f'Import fehlgeschlagen: {exc}','danger')
        return redirect(url_for('import_excel'))
    return render_template('import.html')


@app.get('/jubilees')
def jubilees():
    year=selected_year()
    thresholds=jubilee_thresholds()
    members=get_db().execute("SELECT m.*,f.name family_name FROM members m LEFT JOIN families f ON f.id=m.family_id WHERE m.join_date<>'' AND m.status<>'deceased' AND m.status<>'left' ORDER BY m.last_name,m.first_name").fetchall()
    rows=[]
    for m in members:
        try: joined=date.fromisoformat(m['join_date'])
        except ValueError: continue
        duration=year-joined.year
        if duration in thresholds and duration>0:
            rows.append({'member':m,'years':duration,'anniversary':joined.replace(year=year).isoformat() if not (joined.month==2 and joined.day==29 and year%4!=0) else f'{year}-02-28'})
    rows.sort(key=lambda x:(-x['years'],x['member']['last_name'].casefold(),x['member']['first_name'].casefold()))
    return render_template('jubilees.html',year=year,rows=rows,thresholds=thresholds)


@app.get('/statistics')
def statistics():
    db=get_db()
    members=db.execute("SELECT * FROM members WHERE status='active'").fetchall()
    today=date.today()
    age_buckets=[('0–17',0,17),('18–29',18,29),('30–39',30,39),('40–49',40,49),('50–59',50,59),('60–69',60,69),('70–79',70,79),('80+',80,150)]
    ages=[];tenures=[]
    age_counts={a:0 for a,lo,hi in age_buckets}
    sex={'female':0,'male':0,'diverse':0,'unknown':0}
    tenure_counts={'< 5 Jahre':0,'5–9 Jahre':0,'10–19 Jahre':0,'20–29 Jahre':0,'30+ Jahre':0}
    for m in members:
        sex[m['gender'] if m['gender'] in sex else 'unknown']+=1
        try:
            birth=date.fromisoformat(m['birth_date']);age=today.year-birth.year-((today.month,today.day)<(birth.month,birth.day))
            if 0<=age<=150:
                ages.append(age)
                for k,lo,hi in age_buckets:
                    if lo<=age<=hi: age_counts[k]+=1;break
        except (ValueError,TypeError):pass
        try:
            joined=date.fromisoformat(m['join_date']);tenure=max(0,today.year-joined.year-((today.month,today.day)<(joined.month,joined.day)))
            tenures.append(tenure)
            k='< 5 Jahre' if tenure<5 else '5–9 Jahre' if tenure<10 else '10–19 Jahre' if tenure<20 else '20–29 Jahre' if tenure<30 else '30+ Jahre'
            tenure_counts[k]+=1
        except (ValueError,TypeError):pass
    status_counts={r['status']:r['n'] for r in db.execute('SELECT status,COUNT(*) n FROM members GROUP BY status')}
    years=list(range(today.year-9,today.year+1))
    entries={y:0 for y in years}
    exits={y:0 for y in years}
    deaths={y:0 for y in years}
    # Historical accessions from the join date; departures from the explicitly stored status date.
    # If no departure date is known, it cannot be assigned to a specific year.
    for m in db.execute('SELECT join_date,status,status_date FROM members'):
        try:
            y=date.fromisoformat(m['join_date']).year
            if y in entries: entries[y]+=1
        except ValueError: pass
        try:
            y=date.fromisoformat(m['status_date']).year
            if m['status']=='left' and y in exits: exits[y]+=1
            if m['status']=='deceased' and y in deaths: deaths[y]+=1
        except ValueError: pass
    flow=[{'year':y,'entries':entries[y],'exits':exits[y],'deaths':deaths[y]} for y in years]
    max_flow=max([max(entries.values(),default=0),max(exits.values(),default=0),max(deaths.values(),default=0),1])
    # Additional analytics; only existing, actively maintained records count.
    family_count = db.execute("SELECT COUNT(*) FROM families WHERE active=1 AND EXISTS (SELECT 1 FROM members WHERE members.family_id=families.id AND members.status='active')").fetchone()[0]
    assigned_family = sum(1 for m in members if m['family_id'] is not None)
    birthdays = [0]*12
    for m in members:
        try:
            b = date.fromisoformat(m['birth_date'])
            birthdays[b.month-1] += 1
        except (TypeError,ValueError): pass
    # Demografische Auswertungen, nur mit gültigen Personen-/Familiendaten.
    birth_decades = {}
    age_at_join = {'Unter 18': 0, '18–29': 0, '30–44': 0, '45–59': 0, '60+': 0}
    age_at_join_known = 0
    for m in members:
        try:
            b = date.fromisoformat(m['birth_date'])
            if 1900 <= b.year <= today.year:
                decade = (b.year // 10) * 10
                birth_decades[decade] = birth_decades.get(decade, 0) + 1
                try:
                    j = date.fromisoformat(m['join_date'])
                    age = j.year - b.year - ((j.month,j.day) < (b.month,b.day))
                    if 0 <= age <= 150 and j <= today:
                        bucket = 'Unter 18' if age < 18 else '18–29' if age < 30 else '30–44' if age < 45 else '45–59' if age < 60 else '60+'
                        age_at_join[bucket] += 1
                        age_at_join_known += 1
                except (ValueError,TypeError): pass
        except (ValueError,TypeError): pass
    birth_decades = dict(sorted(birth_decades.items()))
    max_decade = max(birth_decades.values(), default=0) or 1
    family_sizes = {'2 Personen':0,'3 Personen':0,'4 Personen':0,'5+ Personen':0}
    # Gruppierung nur aktiver Familienmitglieder; auch Ein-Personen-Familien separat zählen.
    people_per_family = {}
    for m in members:
        if m['family_id'] is not None:
            people_per_family[m['family_id']] = people_per_family.get(m['family_id'],0) + 1
    family_sizes = {'1 Person':0, **family_sizes}
    for size in people_per_family.values():
        label = '1 Person' if size == 1 else '2 Personen' if size == 2 else '3 Personen' if size == 3 else '4 Personen' if size == 4 else '5+ Personen'
        family_sizes[label] += 1
    max_family_size = max(family_sizes.values(), default=0) or 1
    max_age_at_join = max(age_at_join.values(), default=0) or 1
    largest_family = max(people_per_family.values(), default=0)
    net_flow=[{'year':x['year'], 'net':x['entries']-x['exits']-x['deaths']} for x in flow]
    next_jubilees=0
    for m in members:
        try:
            entered=date.fromisoformat(m['join_date'])
            if today.year+1-entered.year in jubilee_thresholds():next_jubilees+=1
        except (ValueError,TypeError):pass
    return render_template('statistics.html',family_count=family_count,
        family_people=assigned_family,birthdays=birthdays,
        known_birthdays=sum(birthdays), max_birthdays=max(max(birthdays),1), net_flow=net_flow, next_jubilees=next_jubilees,
        birth_decades=birth_decades, max_decade=max_decade, family_sizes=family_sizes, max_family_size=max_family_size, largest_family=largest_family, age_at_join=age_at_join, max_age_at_join=max_age_at_join, age_at_join_known=age_at_join_known,
        count=len(members),age_counts=age_counts,sex=sex,tenure_counts=tenure_counts,
        age_average=round(sum(ages)/len(ages),1) if ages else None,tenure_average=round(sum(tenures)/len(tenures),1) if tenures else None,
        known_ages=len(ages),known_tenures=len(tenures),status_counts=status_counts,
        max_age=max(age_counts.values(),default=0) or 1,max_tenure=max(tenure_counts.values(),default=0) or 1,flow=flow,max_flow=max_flow)


@app.get('/health')
def health():
    return {'status': 'ok'}


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000)
