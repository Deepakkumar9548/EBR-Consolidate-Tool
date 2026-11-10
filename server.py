import os
import sys
import time
import json
import sqlite3
import threading
import mimetypes
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
import re
import email
from email.policy import default
import openpyxl

# Setup workspace directories
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "ebr_enterprise.db")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
EXPORT_DIR = os.path.join(BASE_DIR, "exports")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(EXPORT_DIR, exist_ok=True)

def get_db_connection():
    conn = sqlite3.connect(DB_PATH, timeout=60.0)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    conn.execute("PRAGMA cache_size=-64000;") # 64MB Cache
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    c = conn.cursor()
    c.executescript('''
        CREATE TABLE IF NOT EXISTS jobs (
            job_id TEXT PRIMARY KEY,
            file_name TEXT,
            file_size INTEGER,
            total_rows INTEGER DEFAULT 0,
            processed_rows INTEGER DEFAULT 0,
            total_chunks INTEGER DEFAULT 0,
            completed_chunks INTEGER DEFAULT 0,
            chunk_size INTEGER DEFAULT 50000,
            status TEXT DEFAULT 'QUEUED',
            error_count INTEGER DEFAULT 0,
            sheet_info TEXT,
            start_time TEXT,
            end_time TEXT
        );

        CREATE TABLE IF NOT EXISTS job_checkpoints (
            checkpoint_id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id TEXT,
            chunk_index INTEGER,
            start_row INTEGER,
            end_row INTEGER,
            rows_committed INTEGER,
            status TEXT,
            created_at TEXT
        );

        CREATE TABLE IF NOT EXISTS master_data (
            record_id TEXT PRIMARY KEY,
            job_id TEXT,
            category TEXT,
            month_tag TEXT,
            enquiry_no TEXT,
            booking_no TEXT,
            invoice_no TEXT,
            contact_number TEXT,
            model_group TEXT,
            model_variant TEXT,
            dealer_name TEXT,
            location TEXT,
            region TEXT,
            pv_cv_ev TEXT,
            status TEXT,
            date_time TEXT,
            source_file TEXT,
            source_sheet TEXT,
            source_row INTEGER
        );

        CREATE TABLE IF NOT EXISTS error_data (
            error_id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id TEXT,
            source_file TEXT,
            source_sheet TEXT,
            source_row INTEGER,
            error_type TEXT,
            detail TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_master_contact ON master_data(contact_number);
        CREATE INDEX IF NOT EXISTS idx_master_enq ON master_data(enquiry_no);
        CREATE INDEX IF NOT EXISTS idx_master_bkg ON master_data(booking_no);
        CREATE INDEX IF NOT EXISTS idx_master_inv ON master_data(invoice_no);
        CREATE INDEX IF NOT EXISTS idx_master_job_cat ON master_data(job_id, category);
        CREATE INDEX IF NOT EXISTS idx_master_model ON master_data(model_group);
    ''')
    conn.commit()
    conn.close()

init_db()

# Normalization & Helpers
def normalize_header(raw):
    if raw is None: return ''
    s = str(raw).strip()
    s = re.sub(r'[\r\n\t]+', ' ', s)
    s = re.sub(r'[_\-\/]+', ' ', s)
    s = re.sub(r'[.,()\[\]#:;]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip().lower()
    return s

def clean_phone(v):
    if v is None: return ''
    s = str(v).strip().replace('.0', '').replace(' ', '')
    s = re.sub(r'^\+?91', '', s)
    if s.startswith('0') and len(s) == 11:
        s = s[1:]
    return s if s.isdigit() and len(s) == 10 else s

def parse_date_tag(v):
    if v is None: return '', ''
    if isinstance(v, datetime):
        return v.strftime('%b-%Y').upper(), v.strftime('%Y-%m-%d %H:%M:%S')
    s = str(v).strip()
    if not s: return '', ''
    for fmt in ('%Y-%m-%d %H:%M:%S', '%d-%b-%Y %I:%M %p', '%d-%b-%Y', '%Y-%m-%d', '%d/%m/%Y', '%d/%m/%Y %H:%M'):
        try:
            dt = datetime.strptime(s.split('.')[0], fmt)
            return dt.strftime('%b-%Y').upper(), dt.strftime('%Y-%m-%d %H:%M:%S')
        except:
            pass
    return '', s

SYNONYMS = {
    'ENQUIRY_NUMBER': ['enquiry number', 'dms enquiry number', 'enquiry no', 'lead id'],
    'ENQUIRY_DATE': ['enquiry date', 'enq assign date', 'stage 1 date', 'lead date'],
    'ENQUIRY_TYPE': ['enquiry type', 'enq type', 'lead type', 'type of enquiry', 'type desc', 'enquiry type name', 'type'],
    'ENQUIRY_SOURCE': ['enquiry source', 'source of contact', 'source', 'enq source'],
    'CONTACT_NUMBER': ['customer phone', 'customer mobile phone', 'phone no', 'phone number', 'contact number', 'mobile no', 'mobile number'],
    'MODEL_GROUP': ['product family', 'model group', 'model', 'vehicle model'],
    'MODEL_VARIANT': ['variant description', 'model variant', 'model description', 'variant'],
    'DEALER_NAME': ['dealer name', 'dealer parent name', 'dealership', 'delr name desc', 'dealer'],
    'LOCATION': ['dealer location', 'dealer location name', 'location', 'branch'],
    'REGION': ['zone', 'region', 'state', 'area office'],
    'BOOKING_NUMBER': ['booking number', 'dms booking number', 'booking no'],
    'BOOKING_DATE': ['created date', 'booking date and time', 'booking date'],
    'BOOKING_STATUS': ['booking stage', 'booking status', 'order status'],
    'INVOICE_NUMBER': ['invoice number', 'oem invoice number', 'inv no'],
    'INVOICE_DATE': ['invoice date and time', 'invoice date', 'billing date'],
    'INVOICE_STATUS': ['invoice status', 'retail status']
}

def map_headers(norm_headers):
    mapping = {}
    for field_key, phrases in SYNONYMS.items():
        matched_idx = None
        for p in phrases:
            for idx, nh in enumerate(norm_headers):
                if nh == p:
                    matched_idx = idx
                    break
            if matched_idx is not None: break
        if matched_idx is None:
            for p in phrases:
                for idx, nh in enumerate(norm_headers):
                    if p in nh or nh in p:
                        matched_idx = idx
                        break
                if matched_idx is not None: break
        mapping[field_key] = matched_idx
    return mapping

# Background Chunk Processing Engine (100% Non-Blocking Stream)
def run_chunk_processing(job_id, filepath, chunk_size=50000):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute('''
        INSERT INTO jobs (job_id, file_name, file_size, status, start_time)
        VALUES (?, ?, ?, 'PROCESSING', ?)
        ON CONFLICT(job_id) DO UPDATE SET status = 'PROCESSING', start_time = excluded.start_time
    ''', (job_id, os.path.basename(filepath), os.path.getsize(filepath) if os.path.exists(filepath) else 0, datetime.now().isoformat()))
    conn.commit()

    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        total_processed_job = 0

        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            row_iter = sheet.iter_rows(values_only=True)
            
            try:
                header_row = next(row_iter)
            except StopIteration:
                continue

            norm_headers = [normalize_header(h) for h in header_row]
            col_map = map_headers(norm_headers)

            # Determine category
            sheet_lower = sheet_name.lower() + ' ' + os.path.basename(filepath).lower()
            if 'book' in sheet_lower or col_map.get('BOOKING_NUMBER') is not None:
                category = 'Booking'
            elif 'retail' in sheet_lower or 'inv' in sheet_lower or col_map.get('INVOICE_NUMBER') is not None:
                category = 'Retail'
            else:
                category = 'Enquiry'

            chunk_batch = []
            chunk_index = 0
            row_idx = 1 # 1-based data row counter

            for r in row_iter:
                if not any(r): continue # skip blank
                row_idx += 1
                rec_id = f"{job_id}_{category[:3].upper()}_{row_idx:07d}"

                def get_v(field):
                    idx = col_map.get(field)
                    return str(r[idx]).strip() if (idx is not None and idx < len(r) and r[idx] is not None) else ''

                contact = clean_phone(get_v('CONTACT_NUMBER'))
                m_group = get_v('MODEL_GROUP')
                m_var = get_v('MODEL_VARIANT')
                dealer = get_v('DEALER_NAME')
                loc = get_v('LOCATION')
                reg = get_v('REGION')
                pv_cv_ev = 'EV' if ('EV' in m_group or 'EV' in m_var) else 'PV'

                if category == 'Enquiry':
                    month, dt_str = parse_date_tag(r[col_map['ENQUIRY_DATE']] if col_map.get('ENQUIRY_DATE') is not None else None)
                    enq_no = get_v('ENQUIRY_NUMBER')
                    # Working Point 1: Safely pick Enquiry Type first; fallback to Enquiry Source if Enquiry Type is empty or not present
                    enq_val = get_v('ENQUIRY_TYPE') or get_v('ENQUIRY_SOURCE')
                    bkg_no, inv_no, status = '', '', enq_val
                elif category == 'Booking':
                    month, dt_str = parse_date_tag(r[col_map['BOOKING_DATE']] if col_map.get('BOOKING_DATE') is not None else None)
                    bkg_no = get_v('BOOKING_NUMBER')
                    enq_no, inv_no, status = '', '', get_v('BOOKING_STATUS')
                else: # Retail
                    month, dt_str = parse_date_tag(r[col_map['INVOICE_DATE']] if col_map.get('INVOICE_DATE') is not None else None)
                    inv_no = get_v('INVOICE_NUMBER')
                    enq_no, bkg_no, status = '', '', get_v('INVOICE_STATUS')

                chunk_batch.append((
                    rec_id, job_id, category, month, enq_no, bkg_no, inv_no,
                    contact, m_group, m_var, dealer, loc, reg, pv_cv_ev, status,
                    dt_str, os.path.basename(filepath), sheet_name, row_idx
                ))

                if len(chunk_batch) >= chunk_size:
                    chunk_index += 1
                    c.executemany('''
                        INSERT OR REPLACE INTO master_data (
                            record_id, job_id, category, month_tag, enquiry_no, booking_no, invoice_no,
                            contact_number, model_group, model_variant, dealer_name, location, region, pv_cv_ev, status,
                            date_time, source_file, source_sheet, source_row
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    ''', chunk_batch)
                    
                    total_processed_job += len(chunk_batch)
                    c.execute('''
                        INSERT INTO job_checkpoints (job_id, chunk_index, start_row, end_row, rows_committed, status, created_at)
                        VALUES (?, ?, ?, ?, ?, 'COMMITTED', ?)
                    ''', (job_id, chunk_index, row_idx - len(chunk_batch) + 1, row_idx, len(chunk_batch), datetime.now().isoformat()))
                    
                    c.execute("UPDATE jobs SET processed_rows = ?, completed_chunks = ? WHERE job_id = ?",
                              (total_processed_job, chunk_index, job_id))
                    conn.commit()
                    chunk_batch.clear() # Memory freed

            # Flush residual batch
            if chunk_batch:
                chunk_index += 1
                c.executemany('''
                    INSERT OR REPLACE INTO master_data (
                        record_id, job_id, category, month_tag, enquiry_no, booking_no, invoice_no,
                        contact_number, model_group, model_variant, dealer_name, location, region, pv_cv_ev, status,
                        date_time, source_file, source_sheet, source_row
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ''', chunk_batch)
                total_processed_job += len(chunk_batch)
                c.execute('''
                    INSERT INTO job_checkpoints (job_id, chunk_index, start_row, end_row, rows_committed, status, created_at)
                    VALUES (?, ?, ?, ?, ?, 'COMMITTED', ?)
                ''', (job_id, chunk_index, row_idx - len(chunk_batch) + 1, row_idx, len(chunk_batch), datetime.now().isoformat()))
                c.execute("UPDATE jobs SET processed_rows = ?, completed_chunks = ? WHERE job_id = ?",
                          (total_processed_job, chunk_index, job_id))
                conn.commit()
                chunk_batch.clear()

        wb.close()
        c.execute("UPDATE jobs SET status = 'COMPLETED', end_time = ? WHERE job_id = ?", (datetime.now().isoformat(), job_id))
        conn.commit()
    except Exception as e:
        c.execute("UPDATE jobs SET status = 'FAILED', error_count = error_count + 1 WHERE job_id = ?", (job_id,))
        c.execute("INSERT INTO error_data (job_id, source_file, source_sheet, source_row, error_type, detail) VALUES (?, ?, '', 0, 'JOB_FAILED', ?)",
                  (job_id, os.path.basename(filepath), str(e)))
        conn.commit()
    finally:
        conn.close()

# REST HTTP Server Handler
class AppRequestHandler(BaseHTTPRequestHandler):
    def _send_json(self, data, status=200):
        body = json.dumps(data, default=str).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        query = parse_qs(parsed.query)

        if path == '/':
            html_file = os.path.join(BASE_DIR, "ebr_zip_converter_modern (2).html")
            if os.path.exists(html_file):
                with open(html_file, 'rb') as f:
                    content = f.read()
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.send_header('Content-Length', str(len(content)))
                self.end_headers()
                self.wfile.write(content)
            else:
                self._send_json({'error': 'UI HTML file not found'}, 404)
            return

        # GET /api/jobs/<job_id>/status
        if path.startswith('/api/jobs/') and path.endswith('/status'):
            job_id = path.split('/')[3]
            conn = get_db_connection()
            job = conn.execute("SELECT * FROM jobs WHERE job_id = ?", (job_id,)).fetchone()
            conn.close()
            if job:
                self._send_json(dict(job))
            else:
                self._send_json({'error': 'Job not found'}, 404)
            return

        # GET /api/records?job_id=...&category=...&page=1&limit=50&search=...
        if path == '/api/records':
            job_id = query.get('job_id', [''])[0]
            category = query.get('category', ['Enquiry'])[0]
            page = max(1, int(query.get('page', ['1'])[0]))
            limit = min(500, max(10, int(query.get('limit', ['50'])[0])))
            search = query.get('search', [''])[0].strip()
            offset = (page - 1) * limit

            conn = get_db_connection()
            where_clauses = ["category = ?"]
            params = [category]
            if job_id:
                where_clauses.append("job_id = ?")
                params.append(job_id)
            if search:
                where_clauses.append("(contact_number LIKE ? OR enquiry_no LIKE ? OR booking_no LIKE ? OR invoice_no LIKE ? OR model_group LIKE ? OR dealer_name LIKE ?)")
                sp = f"%{search}%"
                params.extend([sp, sp, sp, sp, sp, sp])

            where_sql = " WHERE " + " AND ".join(where_clauses)
            total_count = conn.execute(f"SELECT COUNT(*) FROM master_data {where_sql}", params).fetchone()[0]
            
            rows = conn.execute(f"SELECT * FROM master_data {where_sql} ORDER BY source_row ASC LIMIT ? OFFSET ?", params + [limit, offset]).fetchall()
            conn.close()

            self._send_json({
                'page': page,
                'limit': limit,
                'total_records': total_count,
                'total_pages': max(1, (total_count + limit - 1) // limit),
                'records': [dict(r) for r in rows]
            })
            return

        # GET /api/reports?job_id=...
        if path == '/api/reports':
            job_id = query.get('job_id', [''])[0]
            conn = get_db_connection()
            where_sql = "WHERE job_id = ?" if job_id else ""
            params = [job_id] if job_id else []

            counts = conn.execute(f"SELECT category, COUNT(*) as cnt FROM master_data {where_sql} GROUP BY category", params).fetchall()
            models = conn.execute(f"SELECT model_group, COUNT(*) as cnt FROM master_data {where_sql} GROUP BY model_group ORDER BY cnt DESC LIMIT 10", params).fetchall()
            dealers = conn.execute(f"SELECT dealer_name, COUNT(*) as cnt FROM master_data {where_sql} GROUP BY dealer_name ORDER BY cnt DESC LIMIT 10", params).fetchall()
            conn.close()

            self._send_json({
                'counts': {r['category']: r['cnt'] for r in counts},
                'top_models': [{ 'model': r['model_group'], 'count': r['cnt'] } for r in models],
                'top_dealers': [{ 'dealer': r['dealer_name'], 'count': r['cnt'] } for r in dealers]
            })
            return

        self._send_json({'error': 'Not Found'}, 404)

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path

        # POST /api/upload
        if path == '/api/upload':
            content_type = self.headers.get('Content-Type', '')
            if 'multipart/form-data' not in content_type:
                self._send_json({'error': 'Content-Type must be multipart/form-data'}, 400)
                return

            length = int(self.headers.get('Content-Length', 0))
            raw_data = self.rfile.read(length)

            # Parse multipart using email library
            msg_bytes = b'Content-Type: ' + content_type.encode('latin1') + b'\r\n\r\n' + raw_data
            msg = email.message_from_bytes(msg_bytes, policy=default)

            uploaded_file = None
            filename = "uploaded_file.xlsx"

            for part in msg.iter_parts():
                cd = part.get("Content-Disposition", "")
                if 'name="file"' in cd:
                    filename = part.get_filename() or "uploaded_data.xlsx"
                    uploaded_file = part.get_payload(decode=True)
                    break

            if not uploaded_file:
                self._send_json({'error': 'No file content received'}, 400)
                return

            job_id = f"JOB-{datetime.now().strftime('%Y%m%d')}-{int(time.time()*1000)%10000:04d}"
            save_path = os.path.join(UPLOAD_DIR, f"{job_id}_{filename}")
            with open(save_path, 'wb') as f:
                f.write(uploaded_file)

            file_size = os.path.getsize(save_path)
            
            # Metadata scan
            sheet_info = []
            try:
                wb = openpyxl.load_workbook(save_path, read_only=True)
                sheet_info = wb.sheetnames
                wb.close()
            except:
                sheet_info = ['Sheet1']

            conn = get_db_connection()
            conn.execute('''
                INSERT INTO jobs (job_id, file_name, file_size, sheet_info, status, start_time)
                VALUES (?, ?, ?, ?, 'QUEUED', ?)
            ''', (job_id, filename, file_size, json.dumps(sheet_info), datetime.now().isoformat()))
            conn.commit()
            conn.close()

            # Start background processing worker automatically
            t = threading.Thread(target=run_chunk_processing, args=(job_id, save_path, 50000), daemon=True)
            t.start()

            self._send_json({
                'message': 'File uploaded & job started successfully',
                'job_id': job_id,
                'file_name': filename,
                'file_size_mb': round(file_size / (1024 * 1024), 2),
                'sheets': sheet_info
            })
            return

        self._send_json({'error': 'Not Found'}, 404)

def run_server(port=8000):
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except:
            pass
    server = HTTPServer(('127.0.0.1', port), AppRequestHandler)
    print("================================================================")
    print(" [INFO] EBR High-Throughput Big-Data Server Engine running on:")
    print(f" >> http://127.0.0.1:{port}")
    print("================================================================")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server...")
        server.server_close()

if __name__ == '__main__':
    run_server()
