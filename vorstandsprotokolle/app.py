import io
import os
import sqlite3
import secrets
import subprocess
import tempfile
import zipfile
from datetime import datetime, date
from pathlib import Path
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory, send_file, abort, flash
from werkzeug.utils import secure_filename
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get('DATA_DIR', '/data'))
UPLOAD_DIR = DATA_DIR / 'uploads'
PREVIEW_DIR = DATA_DIR / 'previews'
DB_PATH = DATA_DIR / 'app.db'

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
PREVIEW_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ['SECRET_KEY']
app.config['SESSION_COOKIE_SAMESITE']='Lax'
app.config['SESSION_COOKIE_HTTPONLY']=True
app.config['SESSION_COOKIE_SECURE']=os.environ.get('COOKIE_SECURE','0')=='1'
app.config['MAX_CONTENT_LENGTH'] = int(os.environ.get('MAX_UPLOAD_MB', '50')) * 1024 * 1024

ALLOWED_EXTENSIONS = {'pdf', 'doc', 'docx', 'xls', 'xlsx'}
CATEGORIES = ['Vorstandssitzungen', 'Veranstaltungen', 'Mitgliederlisten', 'Briefe', 'Sonstiges']
ROLE_LABELS = {'admin': 'Admin', 'editor': 'Redakteur', 'reader': 'Leser'}


def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = db()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            original_name TEXT NOT NULL,
            stored_name TEXT NOT NULL,
            preview_name TEXT NOT NULL,
            uploaded_at TEXT NOT NULL,
            category TEXT NOT NULL DEFAULT 'Sonstiges',
            comment TEXT NOT NULL DEFAULT '',
            document_date TEXT,
            archived INTEGER NOT NULL DEFAULT 0,
            archived_at TEXT
        )
    ''')
    cols = {row['name'] for row in conn.execute('PRAGMA table_info(documents)').fetchall()}
    migrations = {
        'category': "ALTER TABLE documents ADD COLUMN category TEXT NOT NULL DEFAULT 'Sonstiges'",
        'comment': "ALTER TABLE documents ADD COLUMN comment TEXT NOT NULL DEFAULT ''",
        'document_date': "ALTER TABLE documents ADD COLUMN document_date TEXT",
        'archived': "ALTER TABLE documents ADD COLUMN archived INTEGER NOT NULL DEFAULT 0",
        'archived_at': "ALTER TABLE documents ADD COLUMN archived_at TEXT",
    }
    for col, sql in migrations.items():
        if col not in cols:
            conn.execute(sql)

    # Bestehende Dokumente erhalten als Dokumentdatum zunächst ihr Upload-Datum.
    conn.execute("UPDATE documents SET document_date = substr(uploaded_at, 1, 10) WHERE document_date IS NULL OR document_date = ''")

    conn.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            role TEXT NOT NULL,
            password_hash TEXT NOT NULL
        )
    ''')

    conn.commit()
    conn.close()


def current_role():
    return session.get('role') if session.get('authenticated') else None


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get('authenticated'):
            return redirect('/login')
        return fn(*args, **kwargs)
    return wrapper


def roles_required(*roles):
    def decorator(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            if not session.get('authenticated'):
                return redirect('/login')
            if current_role() not in roles:
                abort(403)
            return fn(*args, **kwargs)
        return wrapper
    return decorator


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def valid_iso_date(value):
    try:
        return date.fromisoformat(value).isoformat()
    except (TypeError, ValueError):
        return None


def convert_to_pdf(input_path: Path, output_dir: Path) -> Path:
    ext = input_path.suffix.lower()
    if ext == '.pdf':
        target = output_dir / f"{input_path.stem}.pdf"
        if target.resolve() != input_path.resolve():
            target.write_bytes(input_path.read_bytes())
        return target

    result = subprocess.run(
        ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', str(output_dir), str(input_path)],
        capture_output=True, text=True, timeout=120
    )
    pdf_path = output_dir / f"{input_path.stem}.pdf"
    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(result.stderr or result.stdout or 'Konvertierung fehlgeschlagen')
    return pdf_path


@app.context_processor
def inject_globals():
    return {
        'current_role': current_role(),
        'current_username': session.get('username'),
        'csrf_token': session.get('_csrf',''),
        'role_labels': ROLE_LABELS,
    }


@app.route('/login', methods=['GET', 'POST'])
def login():
    return redirect('/login')

@app.route('/logout')
def logout():
    return redirect('/')

@app.before_request
def check_shared_auth_and_csrf():
    if request.endpoint in ('login','static'):
        return
    if not session.get('authenticated') or session.get('role') not in ('admin','editor','reader'):
        return redirect('/login')
    if request.method == 'POST':
        if request.form.get('_csrf','') != session.get('_csrf','') or not session.get('_csrf'):
            abort(400, 'Ungültiger Sicherheitscode. Bitte neu laden.')
        if session['role'] == 'reader':
            abort(403)

@app.route('/')
@login_required
def index():
    selected_category = request.args.get('category', '').strip()
    requested_year = request.args.get('year', '').strip()
    conn = db()
    years = [row['year'] for row in conn.execute(
        "SELECT DISTINCT substr(document_date,1,4) AS year FROM documents WHERE archived = 0 AND document_date IS NOT NULL AND document_date != '' ORDER BY year DESC"
    ).fetchall() if row['year']]

    selected_year = requested_year if requested_year in years else (years[0] if years else '')
    params = []
    where = ['archived = 0']
    if selected_year:
        where.append("substr(document_date,1,4) = ?")
        params.append(selected_year)
    if selected_category:
        where.append('category = ?')
        params.append(selected_category)

    docs = conn.execute(
        f"SELECT * FROM documents WHERE {' AND '.join(where)} ORDER BY document_date DESC, id DESC",
        params
    ).fetchall()
    conn.close()
    return render_template(
        'index.html', docs=docs, categories=CATEGORIES,
        selected_category=selected_category, years=years, selected_year=selected_year
    )


@app.route('/upload', methods=['POST'])
@roles_required('admin', 'editor')
def upload():
    title = request.form.get('title', '').strip()
    category = request.form.get('category', 'Sonstiges').strip() or 'Sonstiges'
    comment = request.form.get('comment', '').strip()
    document_date = valid_iso_date(request.form.get('document_date', '').strip())
    file = request.files.get('file')

    if not title:
        flash('Bitte einen Titel eingeben.', 'error')
        return redirect(url_for('index'))
    if not document_date:
        flash('Bitte ein gültiges Datum des Schriftstücks eingeben.', 'error')
        return redirect(url_for('index'))
    if not file or not file.filename:
        flash('Bitte eine Datei auswählen.', 'error')
        return redirect(url_for('index'))
    if not allowed_file(file.filename):
        flash('Erlaubt sind PDF, Word (DOC/DOCX) und Excel (XLS/XLSX).', 'error')
        return redirect(url_for('index'))
    if category not in CATEGORIES:
        category = 'Sonstiges'

    original = secure_filename(file.filename)
    suffix = Path(original).suffix.lower()
    token = secrets.token_hex(12)
    stored_name = f'{token}{suffix}'
    stored_path = UPLOAD_DIR / stored_name
    file.save(stored_path)

    preview_temp_dir = PREVIEW_DIR / token
    preview_temp_dir.mkdir(exist_ok=True)
    try:
        converted = convert_to_pdf(stored_path, preview_temp_dir)
        preview_name = f'{token}.pdf'
        (PREVIEW_DIR / preview_name).write_bytes(converted.read_bytes())
    except Exception as exc:
        stored_path.unlink(missing_ok=True)
        flash(f'Datei konnte nicht für die Vorschau verarbeitet werden: {exc}', 'error')
        return redirect(url_for('index'))
    finally:
        for p in preview_temp_dir.glob('*'):
            p.unlink(missing_ok=True)
        if preview_temp_dir.exists():
            preview_temp_dir.rmdir()

    conn = db()
    conn.execute(
        '''INSERT INTO documents
           (title, category, comment, document_date, original_name, stored_name, preview_name, uploaded_at, archived)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0)''',
        (title, category, comment, document_date, original, stored_name, preview_name, datetime.now().isoformat(timespec='seconds'))
    )
    conn.commit()
    conn.close()
    flash('Dokument hochgeladen.', 'success')
    return redirect(url_for('index', year=document_date[:4]))


@app.route('/edit/<int:doc_id>', methods=['POST'])
@roles_required('admin', 'editor')
def edit_document(doc_id):
    title = request.form.get('title', '').strip()
    category = request.form.get('category', 'Sonstiges').strip()
    comment = request.form.get('comment', '').strip()
    document_date = valid_iso_date(request.form.get('document_date', '').strip())
    if not title or not document_date:
        flash('Titel und Dokumentdatum dürfen nicht leer sein.', 'error')
        return redirect(url_for('view_document', doc_id=doc_id))
    if category not in CATEGORIES:
        category = 'Sonstiges'

    conn = db()
    doc = conn.execute('SELECT id FROM documents WHERE id = ?', (doc_id,)).fetchone()
    if not doc:
        conn.close()
        abort(404)
    conn.execute(
        'UPDATE documents SET title = ?, category = ?, comment = ?, document_date = ? WHERE id = ?',
        (title, category, comment, document_date, doc_id)
    )
    conn.commit()
    conn.close()
    flash('Angaben gespeichert.', 'success')
    return redirect(url_for('view_document', doc_id=doc_id))


@app.route('/view/<int:doc_id>')
@login_required
def view_document(doc_id):
    conn = db()
    doc = conn.execute('SELECT * FROM documents WHERE id = ?', (doc_id,)).fetchone()
    conn.close()
    if not doc:
        abort(404)
    return render_template('view.html', doc=doc, categories=CATEGORIES)


@app.route('/preview/<int:doc_id>')
@login_required
def preview(doc_id):
    conn = db()
    doc = conn.execute('SELECT * FROM documents WHERE id = ?', (doc_id,)).fetchone()
    conn.close()
    if not doc:
        abort(404)
    return send_from_directory(PREVIEW_DIR, doc['preview_name'], mimetype='application/pdf', as_attachment=False)


@app.route('/download/<int:doc_id>')
@login_required
def download(doc_id):
    conn = db()
    doc = conn.execute('SELECT * FROM documents WHERE id = ?', (doc_id,)).fetchone()
    conn.close()
    if not doc:
        abort(404)
    return send_from_directory(UPLOAD_DIR, doc['stored_name'], as_attachment=True, download_name=doc['original_name'])


@app.route('/delete/<int:doc_id>', methods=['POST'])
@roles_required('admin', 'editor')
def delete(doc_id):
    conn = db()
    doc = conn.execute('SELECT * FROM documents WHERE id = ?', (doc_id,)).fetchone()
    if not doc:
        conn.close()
        abort(404)
    conn.execute('DELETE FROM documents WHERE id = ?', (doc_id,))
    conn.commit()
    conn.close()
    (UPLOAD_DIR / doc['stored_name']).unlink(missing_ok=True)
    (PREVIEW_DIR / doc['preview_name']).unlink(missing_ok=True)
    flash('Dokument gelöscht.', 'success')
    return redirect(url_for('index'))


@app.route('/archive/<int:doc_id>', methods=['POST'])
@roles_required('admin')
def archive_document(doc_id):
    conn = db()
    doc = conn.execute('SELECT id FROM documents WHERE id = ?', (doc_id,)).fetchone()
    if not doc:
        conn.close()
        abort(404)
    conn.execute('UPDATE documents SET archived = 1, archived_at = ? WHERE id = ?',
                 (datetime.now().isoformat(timespec='seconds'), doc_id))
    conn.commit()
    conn.close()
    flash('Dokument archiviert.', 'success')
    return redirect(url_for('index'))


@app.route('/admin')
@roles_required('admin')
def admin():
    conn = db()
    archived_docs = conn.execute(
        'SELECT * FROM documents WHERE archived = 1 ORDER BY document_date DESC, id DESC'
    ).fetchall()
    users = conn.execute('SELECT username, role FROM users ORDER BY CASE role WHEN "admin" THEN 1 WHEN "editor" THEN 2 ELSE 3 END').fetchall()
    conn.close()
    return render_template('admin.html', archived_docs=archived_docs, users=users)


@app.route('/admin/restore/<int:doc_id>', methods=['POST'])
@roles_required('admin')
def restore_document(doc_id):
    conn = db()
    if not conn.execute('SELECT 1 FROM documents WHERE id = ?', (doc_id,)).fetchone():
        conn.close()
        abort(404)
    conn.execute('UPDATE documents SET archived = 0, archived_at = NULL WHERE id = ?', (doc_id,))
    conn.commit()
    conn.close()
    flash('Dokument aus dem Archiv wiederhergestellt.', 'success')
    return redirect(url_for('admin'))


@app.route('/admin/backup')
@roles_required('admin')
def backup():
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    with tempfile.TemporaryDirectory() as tmpdir:
        snapshot = Path(tmpdir) / 'app.db'
        source = sqlite3.connect(DB_PATH)
        target = sqlite3.connect(snapshot)
        source.backup(target)
        target.close()
        source.close()

        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.write(snapshot, 'app.db')
            for folder, prefix in [(UPLOAD_DIR, 'uploads'), (PREVIEW_DIR, 'previews')]:
                for path in folder.glob('*'):
                    if path.is_file():
                        zf.write(path, f'{prefix}/{path.name}')
        buffer.seek(0)
    return send_file(buffer, mimetype='application/zip', as_attachment=True,
                     download_name=f'vorstandsprotokolle-backup_{timestamp}.zip')


@app.errorhandler(403)
def forbidden(_):
    flash('Dafür hast du keine Berechtigung.', 'error')
    return redirect(url_for('index'))


@app.errorhandler(413)
def too_large(_):
    flash('Datei ist zu groß.', 'error')
    return redirect(url_for('index'))


init_db()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8787)
