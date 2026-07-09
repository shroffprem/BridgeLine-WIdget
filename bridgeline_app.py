#!/usr/bin/env python3
"""
BridgeLine Accounts Widget — Google Sheets edition.
Hosted on Vercel via api/index.py (WSGI entrypoint into the `app` object below).
"""

import re
import csv
import json
import io
import threading
import time
import os
from datetime import datetime, date, timedelta, timezone
from flask import Flask, request, jsonify, render_template_string, Response

# ── Google Sheets setup ───────────────────────────────────────────────────────
import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

SPREADSHEET_ID = "1LKhDNyOd1u48UFgQafbz3oP4Ehgf1hJBt59F9A-8H7U"
SHEET_NAME     = "Accounts"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
IST            = timezone(timedelta(hours=5, minutes=30))

import requests

def trigger_ledger_rebuild():
    """Synchronous call to the Apps Script webhook so the Customer Ledger
    tab rebuilds immediately after this widget writes to Accounts/M Coll.
    Configure ledger_webhook_url / ledger_webhook_token in the Config sheet
    tab. If either is blank, this is a silent no-op. A failure here must
    never break a disbursement/collection save, so all errors are swallowed.

    Runs synchronously (not in a background thread) because Vercel serverless
    functions don't guarantee background threads survive after the HTTP
    response is sent — unlike the old Mac-only version, this call must
    complete (or time out) before the request returns.
    """
    global _accounts_cache
    _recent_activity_cache.clear()
    _accounts_cache = None
    cfg = load_config()
    url = (cfg.get("ledger_webhook_url") or "").strip()
    token = (cfg.get("ledger_webhook_token") or "").strip()
    if not url or not token:
        return
    try:
        # Apps Script /exec URLs always 302-redirect to the real content
        # URL, and that redirect drops a POST body (gets converted to GET
        # by both curl and requests, per standard 301/302 behavior). Using
        # GET with the token in the query string avoids that entirely.
        # rebuildLedger() reformats the whole sheet (borders, column
        # widths) every time and reliably takes 10-15s on the Apps Script
        # side, so the save request itself will visibly take that long.
        requests.get(url, params={"token": token}, timeout=15)
    except Exception as e:
        print(f"[ledger webhook] call failed: {e}")

def get_gspread_client():
    """Auth via a long-lived Google OAuth refresh token stored as Vercel env
    vars — no local token.pickle / browser consent flow, since this runs as
    a stateless serverless function with no persistent disk and no display.
    """
    creds = Credentials(
        None,
        refresh_token=os.environ["GOOGLE_REFRESH_TOKEN"],
        client_id=os.environ["GOOGLE_CLIENT_ID"],
        client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
        token_uri="https://oauth2.googleapis.com/token",
        scopes=SCOPES,
    )
    creds.refresh(Request())
    # BackOffHTTPClient retries 429 quota errors with exponential backoff —
    # without it a burst of reads (multi-step edits, recon rebuilds) can die
    # mid-operation and leave Accounts/M Coll half-updated.
    return gspread.authorize(creds, http_client=gspread.BackOffHTTPClient)

def get_sheet():
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    return sh.worksheet(SHEET_NAME)

REQUESTS_SHEET_NAME = "Requests"
REQUESTS_HEADERS = [
    "Request ID", "Submitted At", "Customer Name", "Cluster", "Branch",
    "Amount", "Account No", "IFSC", "Phone", "SO Name", "Gold Weight",
    "Status", "Disb ID", "Notes", "Bank"
]

def get_requests_sheet():
    gc = get_gspread_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(REQUESTS_SHEET_NAME)
    except Exception:
        ws = sh.add_worksheet(title=REQUESTS_SHEET_NAME, rows=1000, cols=len(REQUESTS_HEADERS))
        ws.append_row(REQUESTS_HEADERS)
        return sh, ws
    # Bank was appended to the schema after the tab was created with exactly
    # 14 columns — the grid must grow before col 15 is addressable.
    header = ws.row_values(1)
    if "Bank" not in header:
        if ws.col_count < len(REQUESTS_HEADERS):
            ws.resize(cols=len(REQUESTS_HEADERS))
        ws.update_cell(1, len(REQUESTS_HEADERS), "Bank")
    return sh, ws

# ── Bank bulk-upload templates ────────────────────────────────────────────────
# Declarative layouts for the bulk NEFT/RTGS files each portal accepts. Field
# names resolve against the per-request dict built in /requests/export-bulk;
# ('const', X) emits the literal X. Portals validate their own template files
# literally (IDFC especially), so adjust headers here after diffing against
# the portal's current downloaded template — never hardcode layouts elsewhere.
BANK_TEMPLATES = {
    'hdfc': {
        'label': 'HDFC NetBanking Plus',
        'filetype': 'xlsx',
        'filename': 'HDFC_Bulk_{date}.xlsx',
        'columns': [
            ('Beneficiary Name', 'customer'),
            ('Beneficiary Account Number', 'account_no'),
            ('IFSC Code', 'ifsc'),
            ('Amount', 'amount'),
            ('Payment Mode', 'mode'),
            ('Debit Account Number', 'debit_account'),
            ('Value Date', 'value_date'),
            ('Narration', 'narration'),
            ('Beneficiary Mobile', 'phone'),
        ],
    },
    'idfc': {
        'label': 'IDFC FIRST Bank',
        'filetype': 'xlsx',
        'filename': 'IDFC_Bulk_{date}.xlsx',
        'columns': [
            ('Beneficiary Name', 'customer'),
            ('Beneficiary Account Number', 'account_no'),
            ('IFSC', 'ifsc'),
            ('Amount', 'amount'),
            ('Payment Mode', 'mode'),
            ('Debit Account Number', 'debit_account'),
            ('Value Date', 'value_date'),
            ('Narration', 'narration'),
            ('Beneficiary Mobile', 'phone'),
            ('Beneficiary Email', ('const', '')),
        ],
    },
    'saraswat': {
        'label': 'Saraswat Bank',
        'filetype': 'csv',
        'filename': 'Saraswat_Bulk_{date}.csv',
        'columns': [
            ('Beneficiary Name', 'customer'),
            ('Beneficiary Account Number', 'account_no'),
            ('IFSC Code', 'ifsc'),
            ('Amount', 'amount'),
            ('Payment Mode', 'mode'),
            ('Narration', 'narration'),
            ('Debit Account Number', 'debit_account'),
            ('Value Date', 'value_date'),
        ],
    },
    'razorpayx': {
        'label': 'RazorpayX',
        'filetype': 'csv',
        'filename': 'RazorpayX_Bulk_{date}.csv',
        'columns': [
            ('RazorpayX Account Number', 'debit_account'),
            ('Payout Amount', 'amount'),
            ('Payout Currency', ('const', 'INR')),
            ('Payout Mode', 'mode'),
            ('Payout Purpose', ('const', 'payout')),
            ('Fund Account Type', ('const', 'bank_account')),
            ('Contact Name', 'customer'),
            ('Account IFSC', 'ifsc'),
            ('Account Number', 'account_no'),
            ('Contact Phone', 'phone'),
            ('Payout Narration', 'narration'),
            ('Payout Reference Id', 'request_id'),
        ],
    },
}

RTGS_MIN_AMOUNT = 200000

# ── Config ────────────────────────────────────────────────────────────────────
COMPANIES = ["HDB", "ICICI"]
CLUSTERS  = ["Bellary", "Hassan", "Hubli", "Mandya", "Mangalore", "Mysore", "Other"]
BRANCHES  = sorted([
    "Adyar", "Beejadi", "Bellary", "Chitradurga", "Chitrapady", "Davangere",
    "Hassan", "Hospet", "JP Nagar", "Kedinje", "Kollegala", "Kuvempu Nagar",
    "Mandya", "Mysore", "Other", "Puttur", "Saligrama", "Santhekatte",
    "Shankarpura", "Shikaripura", "Shimoga", "Thokoot", "Tumkur", "Udupi",
    "Vadarasse", "Valencia", "Vijay Nagar"
])

# ── Daily MIS Package (generate_mis.py, bundled in this repo) ─────────────────
# generate_mis.py and its image assets are bundled directly into this repo
# (no Drive-mount dynamic import — there's no Drive mount on a serverless host).

import generate_mis as mis

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")

# generate_mis.py expects these cached at its own LOGO_PATH/SIGNATURE_PATH/QR_PATH
# (tempfile.gettempdir(), which is writable per-invocation on Vercel too).
_MIS_ASSET_SOURCES = {
    "bl_logo.png":      "LOGO_PATH",
    "bl_signature.png": "SIGNATURE_PATH",
    "bl_qr.png":        "QR_PATH",
}

def ensure_mis_assets_cached():
    import shutil
    for src_name, attr in _MIS_ASSET_SOURCES.items():
        src_path = os.path.join(ASSETS_DIR, src_name)
        dest_path = getattr(mis, attr)
        if os.path.exists(src_path):
            shutil.copyfile(src_path, dest_path)

class _SheetShim:
    """Wraps gspread's get_all_values() rows so generate_mis.load_mcoll() /
    load_contacts() (written against openpyxl's ws.iter_rows API) work unchanged
    against live Sheets data instead of an Excel workbook."""
    def __init__(self, rows):
        self._rows = rows

    def iter_rows(self, min_row=1, max_row=None, values_only=True):
        start = min_row - 1
        end = max_row if max_row is not None else len(self._rows)
        for r in self._rows[start:end]:
            yield tuple(r)

class _WorkbookShim:
    def __init__(self, sheets):
        self._sheets = sheets

    @property
    def sheetnames(self):
        return list(self._sheets.keys())

    def __getitem__(self, name):
        return self._sheets[name]

# Numeric columns (0-indexed) in the Accounts sheet that parse_cases() reads
# with float(): amount, charges, gst, total, coll_amt, discount, balance.
# openpyxl(data_only=True) returns the computed number; gspread's
# get_all_values() returns the *displayed* string (e.g. "2,61,000.00"), so
# these need comma-stripping before generate_mis's float() calls see them.
_MIS_NUMERIC_COLS = {7, 8, 9, 10, 12, 13, 14}

def _clean_numeric_cell(v):
    if isinstance(v, str) and v.strip():
        cleaned = v.replace(',', '').strip()
        try:
            float(cleaned)
            return cleaned
        except ValueError:
            return v
    return v

def load_data_from_sheet(sh):
    """Live-Sheets equivalent of generate_mis.load_data(): same return shape
    (rows, db_raw, mcoll, (cluster_mgrs, branch_contacts)), sourced from the
    same spreadsheet the widget already reads/writes, instead of a manually
    downloaded Excel snapshot."""
    acc_vals = sh.worksheet(SHEET_NAME).get_all_values()
    # Find header row dynamically — it's the row containing "Disbursement ID"
    # (same approach as read_accounts_from_gsheet()). A row inserted/removed
    # above the header otherwise makes it get read as a data row and crashes
    # float() conversion on header text like 'Amount'.
    header_idx = next((i for i, r in enumerate(acc_vals) if r and 'Disbursement ID' in r), 1)
    raw_rows = list(acc_vals[header_idx + 1:])

    # Monthly archive tabs (cases relocated out of Accounts but still need
    # follow-up until Closed). Same column order as Accounts, but unlike the
    # comment here used to claim, not all of them are header-row-free (e.g.
    # 'Jun 26' has a title row + column-header row) — filter to actual case
    # rows by disb-id prefix, same as read_accounts_from_gsheet()'s archive
    # loop already does, instead of assuming a fixed/no-header shape.
    for archive_name in get_archive_tab_names():
        try:
            raw_rows += [r for r in sh.worksheet(archive_name).get_all_values()
                         if r and str(r[0]).strip().startswith('BLP-')]
        except gspread.exceptions.WorksheetNotFound:
            continue

    rows = []
    for raw in raw_rows:
        row = list(raw[:22])
        row += [None] * (22 - len(row))
        row = [None if c == '' else c for c in row]  # match openpyxl's blank-cell None
        for i in _MIS_NUMERIC_COLS:
            row[i] = _clean_numeric_cell(row[i])
        if not row[0] or str(row[0]).strip() == '':
            continue
        rows.append(row)

    sheets = {}
    try:
        mcoll_vals = sh.worksheet('M Coll').get_all_values()
        mcoll_vals = [
            [_clean_numeric_cell(c) if i == 2 else c for i, c in enumerate(r)]
            for r in mcoll_vals
        ]
        sheets['M Coll'] = _SheetShim(mcoll_vals)
    except gspread.exceptions.WorksheetNotFound:
        pass
    try:
        sheets['Contact'] = _SheetShim(sh.worksheet('Contact').get_all_values())
    except gspread.exceptions.WorksheetNotFound:
        pass
    wb_shim = _WorkbookShim(sheets)

    mcoll = mis.load_mcoll(wb_shim)
    cluster_mgrs, branch_contacts = mis.load_contacts(wb_shim)

    db_raw = {}
    try:
        db_vals = sh.worksheet('DashBoard').get_all_values()
        for raw in db_vals[:11]:
            label = raw[0] if len(raw) > 0 else None
            val = raw[1] if len(raw) > 1 else None
            if label:
                db_raw[str(label).strip()] = val
    except gspread.exceptions.WorksheetNotFound:
        pass

    return rows, db_raw, mcoll, (cluster_mgrs, branch_contacts)

# ── Extraction helpers ────────────────────────────────────────────────────────

def parse_inr_amount(text):
    for p in [
        r'INR\s+([\d,]+(?:\.\d+)?)',
        r'Rs\.?\s*([\d,]+(?:\.\d+)?)',
        r'debited for Rs\.([\d,]+(?:\.\d+)?)',
        r'(?:amount|transferred|transfer)[^\d]*([\d,]+)(?:\s*/-)?',
        r'([\d,]+)\s*/-',
    ]:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            try:
                val = float(m.group(1).replace(',', ''))
                if val >= 1000:
                    return val
            except ValueError:
                pass
    return None

def parse_date_from_message(text):
    patterns = [
        (r'on\s+(\d{2}[-/][A-Za-z]{3}[-/]\d{2,4})', ["%d-%b-%y", "%d-%b-%Y"]),
        (r'On\s+(\d{2}[-/]\d{2}[-/]\d{2,4})',        ["%d-%m-%y", "%d-%m-%Y"]),
        (r'on\s+(\d{2}[-/]\d{2}[-/]\d{2,4})',        ["%d-%m-%y", "%d-%m-%Y"]),
        (r'(\d{2}-[A-Z]{3}-\d{2})\b',                ["%d-%b-%y"]),
        (r'(\d{4}-\d{2}-\d{2})',                      ["%Y-%m-%d"]),
    ]
    for pattern, fmts in patterns:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            for fmt in fmts:
                try:
                    return datetime.strptime(m.group(1), fmt).strftime('%d-%m-%Y')
                except ValueError:
                    pass
    return date.today().strftime('%d-%m-%Y')

def extract_utr(text):
    t = text.strip()

    # Explicit UTR/Ref label in SMS-style messages
    m = re.search(r'(?:UTR|Ref(?:erence)?)\s*(?:No|Number|#|ID)?\.?\s*[:\-]?\s*([A-Z0-9]{8,22})', t, re.IGNORECASE)
    if m:
        val = m.group(1).upper()
        if not re.match(r'^\d{1,6}$', val):
            return val

    # Transaction ID label
    m = re.search(r'(?:Transaction|Txn|Trans)[\s_]?(?:ID|No|Ref)?\s*[:\-]?\s*([A-Z0-9]{8,22})', t, re.IGNORECASE)
    if m:
        val = m.group(1).upper()
        if not re.match(r'^\d{1,6}$', val):
            return val

    # Bank UTR codes: HDFCR (RTGS), HDFCH (NEFT), CNRBR, SBINR, KARBR, KARBN, SBIN4,
    # UTIBR, SIBLR, IOBAR, IOBAN, SUSBR, PKGBR, BDBLR, UBINR, BARBR, UJVNH, etc.
    # Pattern: 3-6 uppercase letters followed by any letter or digit, then 8+ digits
    m = re.search(r'\b([A-Z]{3,6}[A-Z0-9]\d{8,16})\b', t, re.IGNORECASE)
    if m:
        val = m.group(1).upper()
        # Reject pure IFSC codes (e.g. UBIN0905925 = 4 letters + 0 + 6 alphanum, no long digit block)
        if not re.match(r'^[A-Z]{4}0[A-Z0-9]{6}$', val):
            return val

    # IMPS: IMPS-REFNUM-... or IMPS Ref No: REFNUM
    m = re.search(r'IMPS[\s\-](\d{10,15})\b', t, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r'IMPS\s*Ref\s*(?:no|number|#)?\s*[:\-]?\s*(\d{10,15})', t, re.IGNORECASE)
    if m:
        return m.group(1)
    # IMPS with a name/company in between (e.g. "IMPS -TRIGOLDWYSE LLP- 617415701480
    # Avl bal INR") — find any standalone long digit run anywhere after IMPS
    m = re.search(r'IMPS\b.*?\b(\d{10,15})\b', t, re.IGNORECASE)
    if m:
        return m.group(1)

    # UPI: ref number is the 12-digit segment before the last -UPI or -PAYMENT suffix,
    # or before @VPA section. Pattern: UPI-...-DIGITS-WORD or UPI-...-DIGITS end-of-string
    m = re.search(r'UPI.*?-(\d{10,15})-?(?:[A-Z]*\s*$|UPI|PAYMENT)', t, re.IGNORECASE)
    if m:
        return m.group(1)
    # UPI with ref embedded anywhere as standalone 12-digit number
    if re.match(r'^UPI[\s\-]', t, re.IGNORECASE):
        m = re.search(r'\b(\d{12})\b', t)
        if m:
            return m.group(1)
    # "Credit Alert!" style: "... from VPA name@bank (UPI 615689900342)"
    m = re.search(r'\(UPI\s+(\d{10,15})\)', t, re.IGNORECASE)
    if m:
        return m.group(1)

    # Ref label
    m = re.search(r'\bref\b\s*[:\-#]?\s*([A-Z0-9]{8,22})', t, re.IGNORECASE)
    if m:
        val = m.group(1).upper()
        if not re.match(r'^\d{1,6}$', val):
            return val

    # 12-digit number at end of string (fallback)
    m = re.search(r'\b(\d{12})\s*\.?\s*$', t)
    if m:
        return m.group(1)

    return ''

def extract_sender_name(text):
    m = re.search(r'Cr-\w{8,}-(.+?)-(?:M[\s./]?S[\s./]?|Bridgeline|BRIDG)', text, re.IGNORECASE)
    if m:
        return m.group(1).strip().title()
    return ''

def extract_repayment(text):
    info = {
        'amount':  parse_inr_amount(text),
        'date':    parse_date_from_message(text),
        'utr':     extract_utr(text),
        'disb_id': '',
        'sender':  extract_sender_name(text),
    }
    m = re.search(r'BLP[-/]\d{6}[-/]\d{3}', text, re.IGNORECASE)
    if m:
        info['disb_id'] = m.group(0).upper()
    return info

def extract_disbursement(text):
    info = {
        'amount':   parse_inr_amount(text),
        'date':     parse_date_from_message(text),
        'customer': '',
        'company':  '',
        'cluster':  '',
        'branch':   '',
    }
    _exclude = {w.lower() for w in BRANCHES + CLUSTERS + COMPANIES} | {
        'name', 'branch', 'mobile', 'bank', 'ifsc', 'account', 'normal', 'bt',
        'pledged', 'value', 'transferred', 'charges', 'amount', 'net', 'wt',
    }

    def _valid_name(n):
        n = n.strip()
        if not (2 < len(n) < 60):
            return False
        words = re.split(r'\s+', n.lower())
        if any(w in _exclude for w in words):
            return False
        return True

    SEP = r'[\s:*\-]+'
    for p in [
        r'(?:customer\s+name|customer)' + SEP + r'([A-Za-z][A-Za-z\s\.]+?)(?:\n|,|/|$)',
        r'(?:^|\b)name' + SEP + r'([A-Za-z][A-Za-z\s\.]+?)(?:\n|,|/|$)',
        r'(?:borrower|client)' + SEP + r'([A-Za-z][A-Za-z\s\.]+?)(?:\n|,|/|$)',
    ]:
        m = re.search(p, text, re.IGNORECASE | re.MULTILINE)
        if m:
            name = m.group(1).strip()
            if _valid_name(name):
                info['customer'] = name.title()
                break

    if re.search(r'\bHDB\b', text, re.IGNORECASE):
        info['company'] = 'HDB'
    elif re.search(r'\bICICI\b', text, re.IGNORECASE):
        info['company'] = 'ICICI'

    for c in CLUSTERS:
        if re.search(r'\b' + re.escape(c) + r'\b', text, re.IGNORECASE):
            info['cluster'] = c
            break

    for b in BRANCHES:
        if re.search(r'\b' + re.escape(b) + r'\b', text, re.IGNORECASE):
            info['branch'] = b
            break

    return info

# ── Data layer ────────────────────────────────────────────────────────────────

import pandas as pd

COL = {
    'disb_id': 1, 'date': 2, 'customer': 3, 'chq': 4, 'company': 5,
    'cluster': 6, 'branch': 7, 'amount': 8, 'charges': 9, 'gst': 10,
    'total': 11, 'coll_date': 12, 'coll_amount': 13, 'discount': 14,
    'balance': 15, 'tat': 16, 'current_date': 17, 'discrepancy': 18,
    'status': 19, 'srv_branch': 20, 'srv_cluster': 21, 'debit_note': 22,
    'credit_note': 23,
    'remarks': 24,
}

def _to_num(v):
    try:
        if pd.isna(v): return 0.0
        return float(str(v).replace(',','').replace('₹','').strip())
    except Exception:
        return 0.0

def calc_total(amount, charges=None, gst=None):
    amt = float(amount or 0)
    ch  = round(float(charges or amt * 0.005), 2)
    g   = round(float(gst or ch * 0.18), 2)
    return round(amt + ch + g, 2)

FOLLOWUP_SHEET_NAME = 'Apr/May26'  # legacy follow-up tab, grandfathered into active_archive_tabs

def get_archive_tab_names():
    """Names of monthly archive tabs (e.g. 'June 26') still being merged into
    every Accounts read, sourced from the Config sheet's 'active_archive_tabs'
    list (JSON array) instead of a hardcoded name — so a brand-new tab the
    monthly rollover creates is picked up automatically, with no code change.
    Falls back to just the legacy FOLLOWUP_SHEET_NAME if the Config key is
    missing (e.g. first run before it's been seeded).
    """
    cfg = load_config()
    tabs = cfg.get('active_archive_tabs')
    if not tabs:
        return [FOLLOWUP_SHEET_NAME]
    return tabs

_accounts_cache = None  # (timestamp, rows)
_ACCOUNTS_CACHE_TTL = 15  # seconds

def read_accounts_from_gsheet():
    """Cached: this is the shared hot path for lookup_case(), get_open_cases(),
    _case_detail()'s fallback, get_recent_activity(), and more — each call
    re-scans Accounts + every active archive tab (4-5 Sheets reads). Left
    uncached, ordinary usage (dashboard load + opening Edit Case + a save)
    burns through the per-minute Sheets quota fast, and gspread's backoff
    then silently retries instead of failing, making saves take 20-40s.
    A warm serverless instance reuses this across requests within the TTL;
    trigger_ledger_rebuild() clears it after every write so a save's own
    follow-up read never sees stale data."""
    global _accounts_cache
    if _accounts_cache and (time.time() - _accounts_cache[0]) < _ACCOUNTS_CACHE_TTL:
        return _accounts_cache[1]
    rows = _read_accounts_from_gsheet_uncached()
    _accounts_cache = (time.time(), rows)
    return rows

def _read_accounts_from_gsheet_uncached():
    sh = get_gspread_client().open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet(SHEET_NAME)
    all_vals = ws.get_all_values()
    # Find header row dynamically — it's the row containing "Disbursement ID"
    header_idx = next((i for i, r in enumerate(all_vals) if r and 'Disbursement ID' in r), 1)
    headers = all_vals[header_idx]
    rows = []
    for i, row in enumerate(all_vals[header_idx + 1:], start=header_idx + 2):
        if row and row[0].startswith('BLP-'):
            record = {'_row': i, '_sheet': SHEET_NAME}
            for j, h in enumerate(headers):
                record[h] = row[j] if j < len(row) else ''
            rows.append(record)

    # Monthly archive tabs (e.g. 'Apr/May26', 'June 26', ...). No header row —
    # columns are in the same order as Accounts, so reuse its header list.
    # Keep reading a tab until every case in it shows Overdue Status = Closed.
    for archive_name in get_archive_tab_names():
        try:
            ws2 = sh.worksheet(archive_name)
            vals2 = ws2.get_all_values()
            for i, row in enumerate(vals2, start=1):
                if row and row[0].startswith('BLP-'):
                    record = {'_row': i, '_sheet': archive_name}
                    for j, h in enumerate(headers):
                        record[h] = row[j] if j < len(row) else ''
                    rows.append(record)
        except gspread.exceptions.WorksheetNotFound:
            continue

    return rows

def read_mcoll_from_gsheet():
    """Each row here is one individual instalment/payment against a case,
    with its own date and amount - unlike Accounts.'Collected Amount', which
    is a running total across every payment ever made for that case."""
    sh = get_gspread_client().open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet('M Coll')
    except gspread.exceptions.WorksheetNotFound:
        return []
    all_vals = ws.get_all_values()
    rows = []
    for row in all_vals[1:]:
        if row and row[0]:
            rows.append({
                'disb_id': row[0],
                'coll_date': row[1] if len(row) > 1 else '',
                'amount': row[2] if len(row) > 2 else '',
            })
    return rows

def get_payment_events(records, mcoll_rows):
    """Flattens collections into individual dated payment events, instead of
    the cumulative 'Collected Amount' column on the Accounts row. Filtering
    the cumulative column by date (e.g. 'collected today') is wrong whenever
    a case has more than one instalment, since that column already includes
    every prior instalment too. A case with M Coll history uses those rows
    (one per instalment); a case with no M Coll history was paid in a single
    one-shot payment, so its Accounts row IS that one event.
    """
    mcoll_by_disb = {}
    for m in mcoll_rows:
        mcoll_by_disb.setdefault(m['disb_id'], []).append(m)

    events = []
    for r in records:
        did = r.get('Disbursement ID', '')
        name = r.get('Customer Name', '')
        if did in mcoll_by_disb:
            for m in mcoll_by_disb[did]:
                d = parse_disb_date(m['coll_date'])
                amt = _to_num(m['amount'])
                if d and amt:
                    events.append({'disb_id': did, 'customer': name, 'date': d, 'amount': amt})
        else:
            cd = r.get('Collected   Date', '') or r.get('Collected Date', '')
            d = parse_disb_date(cd)
            amt = _to_num(r.get('Collected Amount', 0))
            if d and amt:
                events.append({'disb_id': did, 'customer': name, 'date': d, 'amount': amt})
    return events

def read_contacts():
    """Read staff directory from Contact sheet."""
    try:
        sh = get_gspread_client().open_by_key(SPREADSHEET_ID)
        ws = sh.worksheet('Contact')
        rows = ws.get_all_values()
        if not rows:
            return []
        result = []
        current_cluster = ''
        for row in rows[1:]:
            if not any(r.strip() for r in row):
                continue
            cluster = row[0].strip() if len(row) > 0 and row[0].strip() else current_cluster
            if row[0].strip():
                current_cluster = cluster
            record = {
                'cluster':     cluster,
                'name':        row[1].strip() if len(row) > 1 else '',
                'designation': row[2].strip() if len(row) > 2 else '',
                'branch':      row[3].strip() if len(row) > 3 else '',
                'phone':       row[4].strip().rstrip('.') if len(row) > 4 else '',
                'email':       row[5].strip() if len(row) > 5 else '',
            }
            if record['name']:
                result.append(record)
        return result
    except Exception as e:
        raise RuntimeError(f'Contact sheet error: {e}')

def save_contacts(contacts):
    """Write full contacts list back to Contact sheet.

    Sheets has no transactions, so clear()-then-update() has a window where a
    failed update leaves the tab wiped with no rollback. Writing the new data
    first (a plain update() is retriable/idempotent on its own) and only
    clearing now-unused trailing rows afterward means a failure before the
    write completes leaves the old data intact instead of blank."""
    sh = get_gspread_client().open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet('Contact')
    prev_row_count = len(ws.get_all_values())
    rows = [['Cluster', 'Name', 'Designation', 'Branch', 'Phone', 'Email']]
    prev_cluster = ''
    for c in contacts:
        cluster = c.get('cluster', '').strip()
        rows.append([
            cluster if cluster != prev_cluster else '',
            c.get('name', '').strip(),
            c.get('designation', '').strip(),
            c.get('branch', '').strip(),
            c.get('phone', '').strip(),
            c.get('email', '').strip(),
        ])
        if cluster:
            prev_cluster = cluster
    ws.update('A1', rows)
    if prev_row_count > len(rows):
        ws.batch_clear([f'A{len(rows) + 1}:Z{prev_row_count}'])
    ws.format('A1:F1', {
        'textFormat': {'bold': True, 'foregroundColor': {'red':1,'green':1,'blue':1}},
        'backgroundColor': {'red': 0.1, 'green': 0.23, 'blue': 0.36}
    })

def _parse_case(r):
    amount    = _to_num(r.get('Amount', 0))
    total     = _to_num(r.get('Total', 0)) or calc_total(amount)
    collected = _to_num(r.get('Collected Amount', 0))
    balance   = max(0, total - collected)
    return amount, total, collected, balance

def get_open_cases():
    records = read_accounts_from_gsheet()
    cases = []
    for r in records:
        status = r.get('Overdue Status', '').strip()
        if status and status != 'Closed':
            amount, total, collected, balance = _parse_case(r)
            cases.append({
                'disb_id':   r.get('Disbursement ID', ''),
                'customer':  r.get('Customer Name', ''),
                'amount':    amount,
                'total':     total,
                'collected': collected,
                'balance':   balance,
                'status':    status,
            })
    return cases

def lookup_case(disb_id):
    records = read_accounts_from_gsheet()
    for r in records:
        if r.get('Disbursement ID', '').upper() == disb_id.upper():
            amount, total, collected, balance = _parse_case(r)
            return {
                'found':     True,
                'row':       r['_row'],
                'sheet':     r.get('_sheet', SHEET_NAME),
                'customer':  r.get('Customer Name', ''),
                'amount':    amount,
                'total':     total,
                'collected': collected,
                'balance':   balance,
                'status':    r.get('Overdue Status', '').strip(),
            }
    return {'found': False}

def get_next_seq():
    records = read_accounts_from_gsheet()
    max_seq = 0
    for r in records:
        m = re.search(r'-(\d{3})$', r.get('Disbursement ID', ''))
        if m:
            max_seq = max(max_seq, int(m.group(1)))
    return max_seq + 1

def save_disbursement(data):
    ws  = get_sheet()
    seq = get_next_seq()
    try:
        d = datetime.strptime(data['date'], '%d-%m-%Y')
    except Exception:
        d = datetime.today()
    ddmmyy  = d.strftime('%d%m%y')
    disb_id = f"BLP-{ddmmyy}-{seq:03d}"

    amount  = float(data['amount'])
    charges = round(amount * 0.005, 2)
    gst     = round(charges * 0.18, 2)
    total   = round(amount + charges + gst, 2)

    all_vals = ws.get_all_values()
    next_row = len(all_vals) + 1

    row_data = [''] * 24
    row_data[COL['disb_id']-1]      = disb_id
    row_data[COL['date']-1]         = d.strftime('%d-%m-%Y')
    row_data[COL['customer']-1]     = data.get('customer', '')
    row_data[COL['chq']-1]          = data.get('chq', '')
    row_data[COL['company']-1]      = data.get('company', '')
    row_data[COL['cluster']-1]      = data.get('cluster', '')
    row_data[COL['branch']-1]       = data.get('branch', '')
    row_data[COL['amount']-1]       = amount
    row_data[COL['charges']-1]      = charges
    row_data[COL['gst']-1]          = gst
    row_data[COL['total']-1]        = total
    row_data[COL['balance']-1]      = total
    row_data[COL['current_date']-1] = datetime.today().strftime('%d-%m-%Y')
    row_data[COL['status']-1]       = 'Follow Up!'
    row_data[COL['srv_branch']-1]   = data.get('serviced_branch', '')
    row_data[COL['srv_cluster']-1]  = data.get('serviced_cluster', '')
    row_data[COL['debit_note']-1]   = data.get('utr', '')
    row_data[COL['remarks']-1]      = data.get('remarks', '')

    ws.insert_row(row_data, next_row)
    trigger_ledger_rebuild()

    request_id = data.get('request_id', '').strip()
    if request_id:
        try:
            _mark_request_disbursed(request_id, disb_id)
        except Exception:
            pass  # don't fail the disbursement if Requests update errors

    return disb_id

def _cell_num(ws, row, col):
    val = ws.cell(row, col).value or '0'
    try:
        return float(str(val).replace(',', '').replace('₹', '').strip())
    except Exception:
        return 0.0

def save_repayment(data):
    disb_id = data['disb_id'].strip().upper()
    info    = lookup_case(disb_id)
    if not info['found']:
        raise ValueError(f"Disbursement ID '{disb_id}' not found.")

    try:
        coll_date = datetime.strptime(data['date'], '%d-%m-%Y').strftime('%d-%m-%Y')
    except Exception:
        coll_date = datetime.today().strftime('%d-%m-%Y')

    amount   = float(data['amount'])
    discount = float(data.get('discount', 0) or 0)
    utr      = data.get('utr', '').strip()
    raw_msg  = data.get('raw_msg', '')

    sh  = get_gspread_client().open_by_key(SPREADSHEET_ID)
    ws  = sh.worksheet(info.get('sheet', SHEET_NAME))
    row = info['row']

    existing = _cell_num(ws, row, COL['coll_amount'])
    total    = _cell_num(ws, row, COL['total']) or calc_total(_cell_num(ws, row, COL['amount']))
    new_coll = existing + amount
    new_bal  = max(0, total - new_coll - discount)
    # Sub-rupee residue (rounding leftovers from discount/instalment math)
    # counts as fully settled, matching the >= 1.0 "open" threshold
    # generate_mis.py already uses for its own reports.
    new_status = 'Closed' if new_bal < 1 else info['status']

    updates = [
        (row, COL['coll_date'],   coll_date),
        (row, COL['coll_amount'], new_coll),
        (row, COL['status'],      new_status),
        # Balance was previously left stale in the sheet (never written
        # after creation) — only the in-memory API response had the right
        # number. Write it back so the raw sheet matches reality too.
        (row, COL['balance'],     new_bal),
    ]
    if discount:
        updates.append((row, COL['discount'], discount))

    if utr:
        existing_utr = (ws.cell(row, COL['credit_note']).value or '').strip()
        combined_utr = f"{existing_utr}, {utr}" if existing_utr else utr
        updates.append((row, COL['credit_note'], combined_utr))

    remarks = data.get('remarks', '').strip()
    if remarks:
        existing_rem = (ws.cell(row, COL['remarks']).value or '').strip()
        combined_rem = f"{existing_rem} | {remarks}" if existing_rem else remarks
        updates.append((row, COL['remarks'], combined_rem))

    ws.batch_update([{
        'range': gspread.utils.rowcol_to_a1(r, c),
        'values': [[v]]
    } for r, c, v in updates])

    # Record every repayment in M Coll, including single one-shot full
    # closures -- M Coll needs to be the one definitive payment history
    # everywhere it's read (the Customer Ledger, MIS reports, etc.), not
    # just for multi-instalment cases.
    mcoll_error = None
    try:
        mc = sh.worksheet('M Coll')
        mc_vals = mc.get_all_values()
        next_mc_row = len(mc_vals) + 1
        mc.insert_row([
            disb_id, coll_date, amount,
            utr or raw_msg, info.get('customer', ''),
        ], next_mc_row)
    except Exception as e:
        # Previously a bare `except: pass` — Accounts had already been
        # updated above, so a failure here silently detached this payment
        # from M Coll (the definitive payment-history log used by the
        # Customer Ledger and MIS reports), undetected until reconciliation
        # didn't match. Surface it two ways: durably, in the Accounts row's
        # own Remarks cell (so anyone auditing this case later sees it even
        # if nobody was watching the API response at the time), and in the
        # API response so the officer sees a warning immediately.
        mcoll_error = str(e)
        try:
            existing_rem = (ws.cell(row, COL['remarks']).value or '').strip()
            warn = f"M Coll log FAILED for {coll_date} payment of Rs.{amount}: {mcoll_error}"
            combined = f"{existing_rem} | {warn}" if existing_rem else warn
            ws.update_cell(row, COL['remarks'], combined)
        except Exception:
            pass  # best-effort — the warning-write itself must not blow up the request

    trigger_ledger_rebuild()
    result = {'new_collected': new_coll, 'new_balance': new_bal, 'status': new_status}
    if mcoll_error:
        result['mcoll_warning'] = (
            f"Payment recorded on the account but NOT logged to M Coll (payment "
            f"history): {mcoll_error}. It's noted in Remarks — check manually."
        )
    return result

# ── Case editing (fix mistakes in saved disbursements / repayments) ───────────

def _parse_any_date(s):
    for f in ('%d-%m-%Y', '%d/%m/%Y', '%d-%b-%Y', '%Y-%m-%d', '%d-%b-%y'):
        try:
            return datetime.strptime(str(s).strip(), f)
        except Exception:
            pass
    return None

_recent_activity_cache = {}  # {days: (timestamp, events)}
_RECENT_ACTIVITY_TTL = 25  # seconds

def get_recent_activity(days=7):
    """Cached wrapper: this view fires automatically every time the Edit
    Case tab opens and costs 5-6 Sheets reads (Accounts + every active
    archive tab + M Coll). Left uncached, repeated tab opens compete for
    the same per-minute Sheets quota as the actual Load Case / Save
    requests, making THOSE more likely to hit 429 and hang. A warm
    serverless instance reuses this module-level cache across requests
    within the TTL window; a cold start just recomputes once."""
    cached = _recent_activity_cache.get(days)
    if cached and (time.time() - cached[0]) < _RECENT_ACTIVITY_TTL:
        return cached[1]
    events = _get_recent_activity_uncached(days)
    _recent_activity_cache[days] = (time.time(), events)
    return events

def _get_recent_activity_uncached(days=7):
    """Every individual disbursement AND every individual repayment recorded
    in the last N days — a flat, unfiltered event log rather than one
    deduped row per case, so a case with several separate repayments today
    (e.g. Gunashekar with 2 repayments) shows each one instead of collapsing
    them into a single 'last activity' summary.

    Disbursements come from read_accounts_from_gsheet() (merges Accounts +
    every active monthly archive tab). Repayments come straight from M Coll,
    the definitive payment-event log — cluster/branch are looked up from the
    matching Accounts/archive record since M Coll doesn't carry them."""
    cutoff = datetime.today() - timedelta(days=days)
    records = read_accounts_from_gsheet()
    by_disb_id = {r.get('Disbursement ID', ''): r for r in records}

    events = []
    for r in records:
        d = _parse_any_date(r.get('Disbursement Date', ''))
        if not d or d < cutoff:
            continue
        events.append({
            'type': 'Disbursement',
            'disb_id':  r.get('Disbursement ID', ''),
            'customer': r.get('Customer Name', ''),
            'cluster':  r.get('Cluster', ''),
            'branch':   r.get('Branch', ''),
            'amount':   _to_num(r.get('Amount', 0)),
            'date':     r.get('Disbursement Date', ''),
            'utr':      r.get('Debit Note', ''),
            'status':   r.get('Overdue Status', '').strip(),
            'sheet':    r.get('_sheet', SHEET_NAME),
            '_sortkey': d,
        })

    try:
        mc_rows = get_gspread_client().open_by_key(SPREADSHEET_ID).worksheet('M Coll').get_all_values()[1:]
    except Exception:
        mc_rows = []
    for row in mc_rows:
        if not row or not row[0]:
            continue
        d = _parse_any_date(row[1] if len(row) > 1 else '')
        if not d or d < cutoff:
            continue
        disb_id = row[0].strip()
        case = by_disb_id.get(disb_id, {})
        events.append({
            'type': 'Repayment',
            'disb_id':  disb_id,
            'customer': row[4] if len(row) > 4 else case.get('Customer Name', ''),
            'cluster':  case.get('Cluster', ''),
            'branch':   case.get('Branch', ''),
            'amount':   _to_num(row[2]) if len(row) > 2 else 0,
            'date':     row[1] if len(row) > 1 else '',
            'utr':      row[3] if len(row) > 3 else '',
            'status':   case.get('Overdue Status', '').strip(),
            'sheet':    case.get('_sheet', SHEET_NAME),
            '_sortkey': d,
        })

    events.sort(key=lambda e: e['_sortkey'], reverse=True)
    for e in events:
        del e['_sortkey']
    return events

def _get_case_payments(sh, disb_id):
    """M Coll rows for one case, each carrying its live row number so edits/
    deletes can target it. Pulled into its own helper so callers that
    already read M Coll for another reason (recomputing collected/balance)
    can build this list from data they already have in hand, instead of
    _case_detail() re-reading the whole M Coll sheet from scratch."""
    payments = []
    try:
        mc = sh.worksheet('M Coll')
        for i, r in enumerate(mc.get_all_values(), start=1):
            if r and str(r[0]).strip().upper() == disb_id:
                payments.append({
                    'mc_row': i,
                    'date':   r[1] if len(r) > 1 else '',
                    'amount': _to_num(r[2]) if len(r) > 2 else 0,
                    'utr':    r[3] if len(r) > 3 else '',
                })
    except Exception:
        pass
    return payments

def _case_detail(disb_id, known=None):
    """Everything the Edit Case screen needs: the Accounts row's editable +
    computed fields, and the case's individual payments from M Coll.

    `known`, if given, is {'sheet', 'row', 'payments'} already resolved by
    the caller's own write (update_disbursement / _recompute_case_from_mcoll)
    — this skips a second full lookup_case() (which re-scans Accounts plus
    every archive tab) and, when payments are already known, a second full
    M Coll read. Without it, both are looked up fresh here as before."""
    disb_id = disb_id.strip().upper()
    sh = get_gspread_client().open_by_key(SPREADSHEET_ID)

    if known:
        sheet, row, payments = known['sheet'], known['row'], known.get('payments')
    else:
        info = lookup_case(disb_id)
        if not info['found']:
            return None
        sheet, row, payments = info['sheet'], info['row'], None

    ws = sh.worksheet(sheet)
    vals = ws.row_values(row)
    def cell(key):
        i = COL[key] - 1
        return vals[i] if i < len(vals) else ''

    if payments is None:
        payments = _get_case_payments(sh, disb_id)

    total     = _to_num(cell('total')) or calc_total(_to_num(cell('amount')))
    collected = _to_num(cell('coll_amount'))
    discount  = _to_num(cell('discount'))
    # Balance is computed live from Total/Collected/Discount, same as
    # _parse_case() everywhere else in the app — the sheet's stored Balance
    # column (col O) is a stale snapshot written once at creation and never
    # updated by save_repayment()/save_disbursement(), so it must not be
    # read directly here.
    balance = max(0, total - collected - discount)

    return {
        'disb_id': disb_id, 'sheet': sheet,
        'date': cell('date'), 'customer': cell('customer'), 'chq': cell('chq'),
        'company': cell('company'), 'cluster': cell('cluster'), 'branch': cell('branch'),
        'amount': _to_num(cell('amount')), 'charges': _to_num(cell('charges')),
        'gst': _to_num(cell('gst')), 'total': total,
        'discount': discount, 'collected': collected,
        'balance': balance, 'status': cell('status'),
        'srv_branch': cell('srv_branch'), 'srv_cluster': cell('srv_cluster'),
        'debit_note': cell('debit_note'), 'remarks': cell('remarks'),
        'payments': payments,
    }

EDITABLE_TEXT_FIELDS = ('date', 'customer', 'chq', 'company', 'cluster', 'branch',
                        'srv_branch', 'srv_cluster', 'debit_note', 'remarks')

def update_disbursement(data):
    """Writes corrected disbursement fields back to the case's Accounts row.
    An amount change recomputes the derived money chain (charges 0.5%, GST
    18% on charges, total) and re-derives status/balance against existing
    collections."""
    disb_id = data['disb_id'].strip().upper()
    info = lookup_case(disb_id)
    if not info['found']:
        raise ValueError(f"Disbursement ID '{disb_id}' not found.")
    sh  = get_gspread_client().open_by_key(SPREADSHEET_ID)
    ws  = sh.worksheet(info.get('sheet', SHEET_NAME))
    row = info['row']

    updates = [(row, COL[k], str(data[k]).strip())
               for k in EDITABLE_TEXT_FIELDS if k in data]

    if str(data.get('amount', '')).strip() != '':
        amount  = float(data['amount'])
        charges = round(amount * 0.005, 2)
        gst     = round(charges * 0.18, 2)
        total   = round(amount + charges + gst, 2)
        collected = _cell_num(ws, row, COL['coll_amount'])
        discount  = _cell_num(ws, row, COL['discount'])
        bal       = max(0, total - collected - discount)
        # Same sub-₹1 closed threshold as save_repayment; a previously Closed
        # case whose corrected amount leaves money owed reopens.
        status = 'Closed' if bal < 1 else (
            'Follow Up!' if info['status'] in ('', 'Closed') else info['status'])
        updates += [(row, COL['amount'],  amount),
                    (row, COL['charges'], charges),
                    (row, COL['gst'],     gst),
                    (row, COL['total'],   total),
                    (row, COL['status'],  status),
                    (row, COL['balance'], bal)]

    if updates:
        ws.batch_update([{
            'range': gspread.utils.rowcol_to_a1(r, c),
            'values': [[v]]
        } for r, c, v in updates])
        trigger_ledger_rebuild()

    # Handed back to the route so it can build the response via _case_detail()
    # without a second full lookup_case() (Accounts + every archive tab scan).
    return {'sheet': info['sheet'], 'row': row}

def _recompute_case_from_mcoll(disb_id):
    """After an M Coll payment row is added, edited, or deleted, re-derive
    the Accounts row's aggregates from the full M Coll history — M Coll is
    the one definitive payment record. Returns the resolved sheet/row plus
    the case's payments (built from the M Coll data already read here) so
    the caller's response can skip a second lookup_case() + M Coll read."""
    disb_id = disb_id.strip().upper()
    info = lookup_case(disb_id)
    if not info['found']:
        raise ValueError(f"Disbursement ID '{disb_id}' not found.")
    sh  = get_gspread_client().open_by_key(SPREADSHEET_ID)
    ws  = sh.worksheet(info.get('sheet', SHEET_NAME))
    row = info['row']

    mc_all = sh.worksheet('M Coll').get_all_values()
    payments = [{
        'mc_row': i, 'date': r[1] if len(r) > 1 else '',
        'amount': _to_num(r[2]) if len(r) > 2 else 0, 'utr': r[3] if len(r) > 3 else '',
    } for i, r in enumerate(mc_all, start=1) if r and str(r[0]).strip().upper() == disb_id]
    pays = [r for r in mc_all if r and str(r[0]).strip().upper() == disb_id]
    collected = sum(_to_num(r[2]) if len(r) > 2 else 0 for r in pays)

    def _pd(dstr):
        for f in ('%d-%m-%Y', '%d/%m/%Y', '%d-%b-%Y', '%Y-%m-%d'):
            try:
                return datetime.strptime(str(dstr).strip(), f)
            except Exception:
                pass
        return None
    dated = [(d, r[1]) for r in pays if len(r) > 1 for d in [_pd(r[1])] if d]
    coll_date = max(dated)[1] if dated else (pays[-1][1] if pays and len(pays[-1]) > 1 else '')

    total    = _cell_num(ws, row, COL['total']) or calc_total(_cell_num(ws, row, COL['amount']))
    discount = _cell_num(ws, row, COL['discount'])
    bal      = max(0, total - collected - discount)
    status   = 'Closed' if bal < 1 else 'Follow Up!'

    ws.batch_update([{
        'range': gspread.utils.rowcol_to_a1(row, c),
        'values': [[v]]
    } for c, v in [(COL['coll_date'], coll_date), (COL['coll_amount'], collected),
                   (COL['status'], status), (COL['balance'], bal)]])
    trigger_ledger_rebuild()

    return {'sheet': info['sheet'], 'row': row, 'payments': payments}

def _mcoll_row_checked(sh, disb_id, mc_row):
    """Guard against stale row numbers: the M Coll row being targeted must
    still belong to this case."""
    mc = sh.worksheet('M Coll')
    current = mc.row_values(mc_row)
    if not current or str(current[0]).strip().upper() != disb_id.strip().upper():
        raise ValueError('That payment row has moved — reload the case and try again.')
    return mc

def update_repayment(data):
    disb_id = data['disb_id'].strip().upper()
    mc_row  = int(data['mc_row'])
    sh = get_gspread_client().open_by_key(SPREADSHEET_ID)
    mc = _mcoll_row_checked(sh, disb_id, mc_row)
    mc.batch_update([
        {'range': gspread.utils.rowcol_to_a1(mc_row, 2), 'values': [[str(data.get('date', '')).strip()]]},
        {'range': gspread.utils.rowcol_to_a1(mc_row, 3), 'values': [[float(data['amount'])]]},
        {'range': gspread.utils.rowcol_to_a1(mc_row, 4), 'values': [[str(data.get('utr', '')).strip()]]},
    ])
    return _recompute_case_from_mcoll(disb_id)

def delete_repayment(data):
    disb_id = data['disb_id'].strip().upper()
    mc_row  = int(data['mc_row'])
    sh = get_gspread_client().open_by_key(SPREADSHEET_ID)
    mc = _mcoll_row_checked(sh, disb_id, mc_row)
    mc.delete_rows(mc_row)
    return _recompute_case_from_mcoll(disb_id)

def add_repayment(data):
    """Append a new M Coll payment row from the Edit Case tab — the same
    write save_repayment() does, minus the WhatsApp-extraction fields that
    only apply to the standalone Repayment tab's paste-a-message flow."""
    disb_id = data['disb_id'].strip().upper()
    info = lookup_case(disb_id)
    if not info['found']:
        raise ValueError(f"Disbursement ID '{disb_id}' not found.")
    sh = get_gspread_client().open_by_key(SPREADSHEET_ID)
    mc = sh.worksheet('M Coll')
    next_row = len(mc.get_all_values()) + 1
    mc.insert_row([
        disb_id,
        str(data.get('date', '')).strip(),
        float(data['amount']),
        str(data.get('utr', '')).strip(),
        info.get('customer', ''),
    ], next_row)
    return _recompute_case_from_mcoll(disb_id)

# ── Config (stored in a 'Config' tab of the same spreadsheet) ─────────────────
# No persistent local disk exists on a serverless host, so config lives in the
# Sheet (key/value rows, one per top-level key; dict/list values JSON-encoded
# into the cell) instead of a local bridgeline_config.json file.

CONFIG_SHEET_NAME = "Config"

DEFAULT_CONFIG = {
    "whatsapp_groups": {c: "" for c in CLUSTERS},
    "report_time": "09:00",
    "overdue_threshold_days": 7,
    "custom_types": [],
    "ledger_webhook_url": "",
    "ledger_webhook_token": "",
    # Registered bank accounts used for reconciliation, e.g.
    # [{"name": "HDFC xx0923", "account_number": "5010..."}, ...]. A
    # transfer narration containing another registered account's number is
    # an internal movement between our own accounts (Capital In/Out), not
    # a real disbursement/expense/collection.
    "bank_accounts": [],
    # Defaults for the Bank File tab's bulk-upload exports (overridable per
    # export in the UI).
    "bulk_debit_account": "",
    "bulk_narration": "BridgeLine Disbursement",
}

def load_config():
    try:
        ws = get_gspread_client().open_by_key(SPREADSHEET_ID).worksheet(CONFIG_SHEET_NAME)
        rows = ws.get_all_values()
        cfg = {}
        for row in rows[1:] if rows and rows[0] and rows[0][0].strip().lower() == 'key' else rows:
            if len(row) < 2 or not row[0].strip():
                continue
            key, raw = row[0].strip(), row[1]
            try:
                cfg[key] = json.loads(raw)
            except (json.JSONDecodeError, TypeError):
                cfg[key] = raw  # plain strings (report_time, webhook url/token) stay as-is
        for k, v in DEFAULT_CONFIG.items():
            cfg.setdefault(k, v)
        return cfg
    except Exception:
        return DEFAULT_CONFIG.copy()

def save_config(cfg):
    # Write-then-clear-trailing (not clear-then-write) — see save_contacts()
    # for why: Sheets has no transactions, so this shrinks the failure
    # window to "the write itself fails" (old config intact) instead of
    # "clear succeeded, write failed" (config wiped, unrecoverable).
    ws = get_gspread_client().open_by_key(SPREADSHEET_ID).worksheet(CONFIG_SHEET_NAME)
    prev_row_count = len(ws.get_all_values())
    rows = [['key', 'value']]
    for k, v in cfg.items():
        val = json.dumps(v) if isinstance(v, (dict, list)) else str(v)
        rows.append([k, val])
    ws.update(range_name='A1', values=rows)
    if prev_row_count > len(rows):
        ws.batch_clear([f'A{len(rows) + 1}:Z{prev_row_count}'])

# ── Today summary ─────────────────────────────────────────────────────────────

def parse_disb_date(s):
    for fmt in ['%d-%b-%Y','%d-%m-%Y','%Y-%m-%d','%d/%m/%Y','%d-%b-%y','%-d-%b-%Y']:
        try:
            return datetime.strptime(str(s).strip(), fmt)
        except:
            pass
    return None

RECON_LOG_SHEET_NAME = "Recon Log"

def _parse_recon_date(s):
    if not s:
        return None
    d = parse_disb_date(s)
    if d:
        return d
    for fmt in ('%d/%m/%y', '%d/%m/%Y'):
        try:
            return datetime.strptime(str(s).strip(), fmt)
        except Exception:
            pass
    return None

def get_latest_bank_balance():
    """Read the most recent closing balance + date from the Recon Log sheet
    tab, per registered bank account, and combine them. Each /reconcile/save
    call appends one row here (tagged with an Account column) — this
    replaces reading back a persistent Daily Reconciliation.xlsx from a
    local/Drive path, since the hosted version no longer keeps that file
    anywhere central (each reconciliation save is now a one-off in-browser
    download).

    With multiple pooled bank accounts, "the bank balance" is the SUM of
    each account's own latest closing balance (last appended row for that
    account). The "since this date" cutoff used elsewhere to compute
    disbursed/collected-since-reconciliation uses the EARLIEST of the
    accounts' latest-reconciled dates — conservative, so we never miss
    counting a transaction that happened after one account's reconciliation
    but before another's.
    """
    try:
        ws = get_gspread_client().open_by_key(SPREADSHEET_ID).worksheet(RECON_LOG_SHEET_NAME)
        rows = ws.get_all_values()[1:]  # skip header
        if not rows:
            return None, None
        # Only count accounts currently registered in Config — a blank/
        # unrecognised Account tag is orphaned data (e.g. rows saved before
        # multi-account support existed, or a since-removed account) and
        # must not be summed in as if it were still a live account.
        registered = {a.get('name', '').strip() for a in (load_config().get('bank_accounts') or [])}
        latest_by_account = {}
        for row in rows:
            acct = row[4].strip() if len(row) > 4 else ''
            if registered and acct not in registered:
                continue
            latest_by_account[acct] = row  # last write wins -> most recent append for that account
        closings, dates = [], []
        for row in latest_by_account.values():
            if len(row) > 2 and row[2]:
                try:
                    closings.append(float(str(row[2]).replace(',', '').replace('₹', '').strip()))
                except ValueError:
                    pass
            d = _parse_recon_date(row[0] if row else None)
            if d:
                dates.append(d)
        if not closings:
            return None, None
        total_closing = sum(closings)
        earliest_date = min(dates) if dates else None
        return total_closing, earliest_date
    except Exception:
        return None, None

def get_today_summary():
    records = read_accounts_from_gsheet()
    mcoll_rows = read_mcoll_from_gsheet()
    payment_events = get_payment_events(records, mcoll_rows)

    bank_balance, bank_date = get_latest_bank_balance()
    today = datetime.today().date()

    disbursed_today = 0
    collected_today = 0
    total_outstanding = 0
    disbursed_since_bank = 0
    collected_since_bank = 0
    disb_rows, coll_rows, out_rows = [], [], []

    for r in records:
        did  = r.get('Disbursement ID', '')
        name = r.get('Customer Name', '')
        dd = r.get('Disbursement Date','')
        d = parse_disb_date(dd)
        if d and d.date() == today:
            amt = _to_num(r.get('Amount', 0))
            disbursed_today += amt
            disb_rows.append({'disb_id': did, 'customer': name, 'amount': amt})
        if d and bank_date and d.date() > bank_date.date():
            disbursed_since_bank += _to_num(r.get('Amount', 0))
        if r.get('Overdue Status','').strip() not in ('Closed',''):
            _, total, collected, balance = _parse_case(r)
            total_outstanding += balance
            if balance > 0:
                out_rows.append({'disb_id': did, 'customer': name, 'amount': balance})

    # 'Collected Today' / 'collected since bank reconciliation' must sum just
    # the individual payment events on/after the relevant date - not the
    # cumulative Accounts column, which mixes in every prior instalment too.
    for ev in payment_events:
        ev_date = ev['date'].date()
        if ev_date == today:
            collected_today += ev['amount']
            coll_rows.append({'disb_id': ev['disb_id'], 'customer': ev['customer'], 'amount': ev['amount']})
        if bank_date and ev_date > bank_date.date():
            collected_since_bank += ev['amount']

    available = None
    if bank_balance is not None:
        available = bank_balance - disbursed_since_bank + collected_since_bank
    return {
        'disbursed_today':   disbursed_today,
        'collected_today':   collected_today,
        'total_outstanding': total_outstanding,
        'available_for_disbursement': available,
        'bank_balance': bank_balance,
        'bank_balance_date': bank_date.strftime('%d %b %Y') if bank_date else None,
        'disbursed_since_bank': disbursed_since_bank,
        'collected_since_bank': collected_since_bank,
        'date': datetime.today().strftime('%d %b %Y'),
        'disb_rows': disb_rows,
        'coll_rows': coll_rows,
        'out_rows': out_rows,
    }

# ── Bank Reconciliation ───────────────────────────────────────────────────────

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def parse_bank_statement(filepath, filename):
    """Parse bank statement CSV/Excel. Returns transactions + opening/closing balance."""
    ext = filename.lower().rsplit('.', 1)[-1]
    try:
        if ext == 'csv':
            df = pd.read_csv(filepath, header=None, dtype=str, encoding='utf-8', on_bad_lines='skip')
        elif ext == 'xls':
            df = pd.read_excel(filepath, header=None, dtype=str, engine='xlrd')
        else:
            df = pd.read_excel(filepath, header=None, dtype=str, engine='openpyxl')
    except Exception as e:
        raise ValueError(f"Could not read file: {e}")

    # Scan for opening / closing balance in metadata rows
    opening_balance = closing_balance = None
    for _, row in df.iterrows():
        cells = [str(c).strip() for c in row if pd.notna(c) and str(c).strip()]
        combined = ' '.join(cells).lower()
        if opening_balance is None and 'opening' in combined and 'balance' in combined:
            for c in cells:
                v = _clean_amount(c)
                if v: opening_balance = v; break
        if closing_balance is None and ('closing' in combined or 'available balance' in combined) and 'balance' in combined:
            for c in cells:
                v = _clean_amount(c)
                if v: closing_balance = v; break

    # Find transaction header row
    header_row = None
    for i, row in df.iterrows():
        cells = [str(c).lower() for c in row if pd.notna(c)]
        combined = ' '.join(cells)
        if ('debit' in combined or 'withdrawal' in combined or 'credit' in combined or 'deposit' in combined) \
                and ('date' in combined or 'narration' in combined or 'description' in combined):
            header_row = i; break

    if header_row is None:
        raise ValueError("Could not detect header row in the statement.")

    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = [str(c).strip().lower() if pd.notna(c) else f'col_{i}' for i, c in enumerate(df.columns)]

    col_map = {}
    for col in df.columns:
        cl = str(col).lower()
        if any(k in cl for k in ['value date', 'txn date', 'transaction date', 'posting date', 'date']):
            col_map.setdefault('date', col)
        if any(k in cl for k in ['narration', 'description', 'particulars', 'remarks', 'details']):
            col_map.setdefault('description', col)
        if any(k in cl for k in ['utr', 'ref no', 'chq/ref', 'cheque', 'reference']):
            col_map.setdefault('utr', col)
        if any(k in cl for k in ['debit', 'withdrawal', 'dr amount']):
            col_map.setdefault('debit', col)
        if any(k in cl for k in ['credit', 'deposit', 'cr amount']):
            col_map.setdefault('credit', col)
        if 'balance' in cl:
            col_map.setdefault('balance', col)

    # Keywords that identify summary/footer rows — skip these
    SKIP_KEYWORDS = {'opening balance', 'closing balance', 'total', 'available balance',
                     'ledger balance', 'brought forward', 'carried forward', 'statement summary',
                     'end of statement', 'generated on', 'dr count', 'cr count',
                     'debit count', 'credit count', 'page no', '****'}

    rows = []
    last_balance = None
    for _, row in df.iterrows():
        date_val = str(row.get(col_map.get('date', ''), '') or '').strip()
        desc     = str(row.get(col_map.get('description', ''), '') or '').strip()
        utr_raw  = str(row.get(col_map.get('utr', ''), '') or '').strip()
        debit    = _clean_amount(row.get(col_map.get('debit', ''), ''))
        credit   = _clean_amount(row.get(col_map.get('credit', ''), ''))
        balance  = _clean_amount(row.get(col_map.get('balance', ''), ''))

        if not date_val or date_val.lower() in ('nan', 'none', '', 'date'):
            continue
        # Skip separator/asterisk rows
        if set(date_val.replace('*','').replace('-','').strip()) == set():
            continue
        # Skip rows where "date" cell is a pure number (statement summary totals row e.g. 5704953.02)
        try:
            float(date_val.replace(',', ''))
            continue
        except ValueError:
            pass
        # Skip summary/footer rows by keyword anywhere in the row
        row_text = ' '.join(str(v).lower() for v in row.values if pd.notna(v))
        desc_l = desc.lower()
        if any(kw in desc_l or kw in date_val.lower() or kw in row_text for kw in SKIP_KEYWORDS):
            continue
        # Skip rows that look like count/summary lines (no real date format)
        if not re.search(r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}', date_val):
            continue
        if debit == 0 and credit == 0:
            continue

        # UTR strategy:
        #   Credits  → Chq/Ref column IS the RTGS/NEFT UTR (e.g. KARBR52026...)
        #   Debits   → Chq/Ref column has HDFC sequential number (000...159) — useless
        #              Real UTR is at the END of the narration (e.g. -HDFCR52026060165094792)
        chq_clean = utr_raw if utr_raw.lower() not in ('nan','none','') else ''
        if credit > 0:
            # For credits, prefer the Chq/Ref col; fall back to narration extraction
            utr = chq_clean if (chq_clean and not re.match(r'^0+\d{0,6}$', chq_clean)) \
                  else extract_utr(desc)
        else:
            # For debits, always extract from narration (Chq col has sequential nums)
            utr = extract_utr(desc)
            # If narration extraction failed, try the chq col as last resort
            if not utr and chq_clean and not re.match(r'^0+\d{0,6}$', chq_clean):
                utr = chq_clean

        if balance:
            last_balance = balance

        rows.append({'date': date_val, 'description': desc, 'utr': utr,
                     'debit': debit, 'credit': credit, 'balance': balance})

    if rows:
        if opening_balance is None and rows[0]['balance']:
            first = rows[0]
            opening_balance = round(first['balance'] + first['debit'] - first['credit'], 2)
        if closing_balance is None and last_balance:
            closing_balance = last_balance

    return {'transactions': rows,
            'opening_balance': opening_balance or 0,
            'closing_balance': closing_balance or 0}

def _clean_amount(v):
    try:
        s = str(v or '').replace(',', '').replace('₹', '').replace(' ', '').strip()
        if not s or s.lower() in ('nan', 'none', '-', ''):
            return 0.0
        return float(s)
    except Exception:
        return 0.0

def _xl_styles():
    thin  = Side(style='thin',   color='D0DCE8')
    return {
        'thin_border': Border(left=thin, right=thin, top=thin, bottom=thin),
        'hdr_fill':    PatternFill("solid", fgColor="1A3A5C"),
        'hdr_font':    Font(bold=True, color="FFFFFF", size=9, name='Arial'),
        'grn_fill':    PatternFill("solid", fgColor="E2EFDA"),   # matched / closed
        'amb_fill':    PatternFill("solid", fgColor="FFF2CC"),   # near match / partial
        'red_fill':    PatternFill("solid", fgColor="FFE0E0"),   # unmatched / review
        'blu_fill':    PatternFill("solid", fgColor="DDEEFF"),   # capital / other
        'tot_fill':    PatternFill("solid", fgColor="F2F2F2"),
        'tot_font':    Font(bold=True, size=9, name='Arial'),
        'norm_font':   Font(size=9, name='Arial'),
        'bold_font':   Font(bold=True, size=9, name='Arial'),
        'title_font':  Font(bold=True, size=11, name='Arial', color='1A3A5C'),
        'sub_font':    Font(bold=True, size=9, name='Arial', color='555555'),
    }

def _hdr(ws, row, cols, s):
    for i, (label, width) in enumerate(cols, 1):
        c = ws.cell(row, i, label)
        c.fill = s['hdr_fill']; c.font = s['hdr_font']; c.border = s['thin_border']
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[row].height = 26

def _row(ws, r, vals, s, right_cols=(), fill=None, font=None):
    f = fill or (s['grn_fill'] if r % 2 == 0 else PatternFill("solid", fgColor="FFFFFF"))
    for i, v in enumerate(vals, 1):
        c = ws.cell(r, i, v)
        c.fill = f; c.font = font or s['norm_font']; c.border = s['thin_border']
        if i in right_cols: c.alignment = Alignment(horizontal='right')

def _n(v):
    """Number → Indian-comma string, blank if zero/None."""
    try:
        f = float(v)
        return f"{f:,.2f}" if f else ''
    except Exception:
        return ''

def _title(ws, text, ncols, s):
    ws.merge_cells(f'A1:{get_column_letter(ncols)}1')
    c = ws.cell(1, 1, text)
    c.font = s['title_font']; c.alignment = Alignment(horizontal='left')
    ws.row_dimensions[1].height = 22

def _legend(ws, text, ncols, s):
    ws.merge_cells(f'A2:{get_column_letter(ncols)}2')
    c = ws.cell(2, 1, text)
    c.font = Font(italic=True, size=8, name='Arial', color='555555')
    c.alignment = Alignment(horizontal='left')

def _expense_category(desc):
    d = desc.lower()
    if 'pos' in d or 'swipe' in d:                                    return 'Card Purchase (POS)'
    if any(k in d for k in ['claude','openai','chatgpt','base44','lovable','subscription','me dc si']):
                                                                       return 'Software Subscription'
    if any(k in d for k in ['airtel','jio','telecom','broadband']):   return 'Telecom'
    if any(k in d for k in ['gst/bank','gst/gst','markup','dcc','dc intl','bank charg','service charge']):
                                                                       return 'Bank Charges / GST'
    if 'salary' in d or 'sal ' in d:                                  return 'Salary'
    if 'partner' in d or 'withdrawal' in d or 'proprietor' in d:     return 'Partner Share'
    if any(k in d for k in ['pradaan','pradan']):                     return 'Pradaan Routing'
    return 'Miscellaneous'

def _is_fd_booking(desc):
    return 'fd booked' in desc.lower() or 'fd - booked' in desc.lower()

def _is_expense_debit(desc, amt):
    """Return True if this debit is an operating expense (not a loan disbursement)."""
    d = desc.lower()
    # Fixed deposits are capital movements, not expenses
    if _is_fd_booking(desc):
        return False
    # RTGS/NEFT DR with large amounts are almost always disbursements
    if re.search(r'rtgs\s+dr|neft\s+dr', d) and amt >= 50000:
        return False
    # Partner share / owner withdrawals
    if any(k in d for k in ['partshare', 'part share', 'prem narayan', 's a prem']):
        return True
    # POS, card, subscriptions, bank charges, UPI small payments → expense
    if any(k in d for k in ['pos ', ' pos', 'me dc', 'dc intl', 'markup', 'gst/bank',
                              'subscription', 'claude', 'openai', 'chatgpt', 'base44',
                              'airtel', 'jio', 'salary', 'partner']):
        return True
    # Small UPI/IMPS debits (< 25000) not RTGS/NEFT DR → expense
    if amt < 25000 and re.search(r'^(upi|imps|ft\s)', d):
        return True
    return False

def _date_in_period(date_str, period_month):
    """Return True if date_str (any format DD/MM/YY, DD-MM-YYYY, etc.) falls in period_month (e.g. 'Jun 2026')."""
    if not date_str or not period_month:
        return False
    try:
        # period_month e.g. "Jun 2026" → month=6, year=2026
        from datetime import datetime as _dt
        pm = _dt.strptime(period_month, '%b %Y')
        for fmt in ('%d/%m/%Y', '%d-%m-%Y', '%d/%m/%y', '%d-%m-%y',
                    '%Y-%m-%d', '%m/%d/%Y', '%d %b %Y', '%d %B %Y'):
            try:
                d = _dt.strptime(str(date_str).strip(), fmt)
                return d.month == pm.month and d.year == pm.year
            except ValueError:
                continue
    except Exception:
        pass
    return False

# ── Transaction matcher ───────────────────────────────────────────────────────

def _match_transactions(txns, records, mc_rows):
    """Auto-classify + match each bank transaction against Accounts + M Coll."""

    # Registered bank accounts (Config 'bank_accounts') — a transfer
    # narration containing one of THESE account numbers is money moving
    # between our own accounts (Capital Out on the sending statement,
    # Capital In on the receiving one), not a real disbursement/expense/
    # collection. Distinct from acct_to_disb below, which indexes CUSTOMER
    # beneficiary account numbers for FT collection matching.
    own_account_numbers = set()
    try:
        for acc in (load_config().get('bank_accounts') or []):
            num = str(acc.get('account_number', '')).strip()
            if num:
                own_account_numbers.add(num)
    except Exception:
        pass

    # UTR → disb_id from debit note column (disbursements going out)
    utr_to_disb     = {}
    utr_to_coll_acct = {}  # UTR/ref in Credit Note of Accounts → disb_id (collection receipts recorded in sheet)
    amt_to_disb_id  = {}   # round(amt) → disb_id  (for amount-only fallback)
    for r in records:
        did      = r.get('Disbursement ID','').strip()
        deb_note = str(r.get('Debit Note','') or '').strip()
        crd_note = str(r.get('Credit Note','') or '').strip()
        amt      = _to_num(r.get('Amount', 0))
        coll_amt = _to_num(r.get('Collected Amount', 0) or r.get('Collected   Amount', 0))
        # Debit Note → disb UTR map
        for u in re.split(r'[,;\s]+', deb_note):
            u = u.strip().upper()
            if len(u) >= 8: utr_to_disb[u] = did
        # Also extract numeric/alphanumeric refs from any SMS-style debit note text
        for num in re.findall(r'\b([A-Z0-9]{10,22})\b', deb_note.upper()):
            utr_to_disb[num] = did
        # Credit Note → collection UTR map (IMPS/NEFT/UPI refs embedded in SMS)
        for u in re.split(r'[,;\s]+', crd_note):
            u = u.strip().upper()
            if len(u) >= 8: utr_to_coll_acct[u] = did
        for num in re.findall(r'\b([A-Z0-9]{10,22})\b', crd_note.upper()):
            utr_to_coll_acct[num] = did
        if amt: amt_to_disb_id.setdefault(round(amt), []).append(did)
        if coll_amt: amt_to_disb_id.setdefault(round(coll_amt), []).append(did)

    # UTR → disb_id from M Coll (collections coming in)
    utr_to_coll   = {}
    amt_to_coll   = {}   # round(amt) → disb_id
    for mc in mc_rows:
        did = (mc[0] if len(mc)>0 else '').strip()
        raw = (mc[3] if len(mc)>3 else '').strip()
        amt = _clean_amount(mc[2]) if len(mc)>2 else 0
        if did:
            # Map the full Credit Note value
            if raw: utr_to_coll[raw.upper()] = did
            # Also extract any IMPS/NEFT/RTGS ref numbers embedded in SMS text (10-15 digits)
            for num in re.findall(r'\b\d{10,15}\b', raw):
                utr_to_coll[num] = did
            # Extract bank UTR codes (e.g. KARBR5..., HDFCR5...)
            for code in re.findall(r'[A-Z]{4,}[A-Z0-9]{6,}', raw.upper()):
                utr_to_coll[code] = did
        if amt and did: amt_to_coll.setdefault(round(amt), []).append(did)

    # Customer name → disb_id (for Pradaan / name-based match)
    name_to_disb = {}   # full name → disb_id
    words_to_disb = []  # [(significant_words_set, disb_id, full_name)] for fuzzy matching
    STOP_WORDS = {'mr', 'mrs', 'ms', 'dr', 'the', 'and', 'of', 'to', 'in', 'for', 'a', 'an'}
    for r in records:
        did  = r.get('Disbursement ID','').strip()
        name = r.get('Customer Name','').strip().lower()
        if name:
            name_to_disb[name] = did
            words = {w for w in re.split(r'[\s\.\-]+', name) if len(w) >= 3 and w not in STOP_WORDS}
            if words:
                words_to_disb.append((words, did, name))

    def _fuzzy_name_match(desc):
        """Return (disb_id, score) for best name match in desc. Score = matched word count."""
        dl = desc.lower()
        best_did, best_score = '', 0
        for words, did, full_name in words_to_disb:
            score = sum(1 for w in words if w in dl)
            ratio = score / len(words) if words else 0
            # Need ≥2 words matched OR full name is single meaningful word with exact match
            if score >= 2 or (len(words) == 1 and ratio == 1.0):
                if score > best_score:
                    best_score, best_did = score, did
        return best_did, best_score

    # HDFC account number → disb_id (for FT/internal transfers)
    # Pattern: "PRANAV T P DR - 50100551677276 - PRANAV T P" — the 14-digit number is beneficiary acct
    # We index by account number so "FT -BRIDGELINE PARTNERS CR - 50100551677276 - PRANAV T P" can match
    acct_to_disb = {}
    for tx in txns:
        desc = tx.get('description', '')
        if tx['debit'] > 0:
            m = re.search(r'\b(5010\d{10})\b', desc)  # HDFC account numbers start with 5010
            if m:
                acct_to_disb[m.group(1)] = None  # placeholder; will resolve via amount below

    result = []
    for tx in txns:
        desc  = tx.get('description', '')
        utr   = tx.get('utr', '').strip().upper()
        dr, cr = tx['debit'], tx['credit']
        tx_type = tx_ref = tx_basis = tx_notes = ''

        own_acct_m = None
        if own_account_numbers:
            for num in own_account_numbers:
                if num in desc:
                    own_acct_m = num
                    break

        if dr > 0:
            if own_acct_m:
                tx_type = 'Capital Out'; tx_notes = f'Internal transfer to own account {own_acct_m}'
            elif _is_fd_booking(desc):
                tx_type = 'FD Booking'; tx_notes = 'Fixed Deposit — internal capital movement'
            elif _is_expense_debit(desc, dr):
                tx_type  = 'Expense'
                tx_notes = _expense_category(desc)
            else:
                # 1. UTR match
                matched = utr_to_disb.get(utr)
                if not matched:
                    desc_utr = extract_utr(desc).upper()
                    matched  = utr_to_disb.get(desc_utr)
                # IFSC-in-narration: Debit Note may contain an IFSC (e.g. UBIN0905925)
                # that appears inside the bank narration (e.g. "NEFT DR-UBIN0905925-BASAVARAJ...")
                if not matched:
                    for r in records:
                        did2     = r.get('Disbursement ID','').strip()
                        dn2      = str(r.get('Debit Note','') or '').strip().upper()
                        ifsc_m   = re.match(r'^([A-Z]{4}0[A-Z0-9]{6})$', dn2)
                        if ifsc_m and ifsc_m.group(1) in desc.upper():
                            matched = did2
                            break

                if matched:
                    tx_type = 'Disbursement'; tx_ref = matched
                    tx_basis = 'UTR'; tx_notes = 'Accounts (Debit Note)'
                else:
                    amt_cands = amt_to_disb_id.get(round(dr), [])
                    name_did, name_score = _fuzzy_name_match(desc)
                    # 2. Amount + name (strongest non-UTR match)
                    if amt_cands and name_did and name_did in amt_cands:
                        tx_type = 'Disbursement'; tx_ref = name_did
                        tx_basis = 'Amount+Name'; tx_notes = 'Accounts (Disb)'
                    # 3. Amount only
                    elif amt_cands:
                        tx_type = 'Disbursement'; tx_ref = amt_cands[0]
                        tx_basis = 'Amount'; tx_notes = 'Accounts (Disb)'
                    # 4. Name only (with decent score)
                    elif name_did and name_score >= 2:
                        tx_type = 'Disbursement'; tx_ref = name_did
                        tx_basis = 'Name'; tx_notes = 'Fuzzy name — confirm'
                    else:
                        tx_type = 'Expense'; tx_notes = _expense_category(desc)

        elif cr > 0:
            if own_acct_m:
                tx_type = 'Capital In'; tx_notes = f'Internal transfer from own account {own_acct_m}'

            # ₹1 test credits — banks/borrowers send ₹1 to verify account before full payment
            elif cr == 1.0:
                tx_type = 'Test Credit'; tx_notes = 'Penny verification — ignore'

            # Pradaan routing = collection routed via Pradaan account
            elif re.search(r'pradaan|pradan', desc, re.IGNORECASE):
                cands = amt_to_coll.get(round(cr), [])
                tx_type = 'Collection (via Pradaan)'
                tx_ref  = cands[0] if cands else ''
                tx_basis= 'Amount' if cands else '—'
                tx_notes= 'Pradaan Routing'

            else:
                # Match to collection by UTR — check M Coll first, then Accounts Credit Note
                matched = utr_to_coll.get(utr)
                if not matched:
                    desc_utr = extract_utr(desc).upper()
                    matched  = utr_to_coll.get(desc_utr)
                if not matched:
                    matched = utr_to_coll_acct.get(utr)

                # FT (internal fund transfer) — match via HDFC account number in narration
                # e.g. "FT -BRIDGELINE PARTNERS CR - 50100551677276 - PRANAV T P"
                # The same account number appears in the debit narration for that disbursement
                if not matched and re.match(r'FT\s*[\-–]', desc, re.IGNORECASE):
                    acct_m = re.search(r'\b(5010\d{10})\b', desc)
                    if acct_m:
                        acct_num = acct_m.group(1)
                        # Find the debit transaction with this account number and match its Disb ID
                        for other in txns:
                            if other['debit'] > 0 and acct_num in other.get('description', ''):
                                cands = amt_to_disb_id.get(round(other['debit']), [])
                                if cands:
                                    matched = cands[0]
                                    break

                if matched:
                    tx_type = 'Collection'; tx_ref = matched
                    tx_basis = 'UTR'; tx_notes = 'Matched Collections'
                else:
                    coll_cands = amt_to_coll.get(round(cr), [])
                    name_did, name_score = _fuzzy_name_match(desc)
                    # Amount + name
                    if coll_cands and name_did and name_did in coll_cands:
                        tx_type = 'Collection'; tx_ref = name_did
                        tx_basis = 'Amount+Name'; tx_notes = 'M Coll + name'
                    # Amount only (M Coll)
                    elif coll_cands:
                        tx_type = 'Collection'; tx_ref = coll_cands[0]
                        tx_basis = 'Amount'; tx_notes = 'Fuzzy — confirm'
                    # Amount match against Accounts disbursement amounts (repayment of full loan)
                    elif name_did and name_score >= 2:
                        # Check if amount also matches for higher confidence
                        disb_cands = amt_to_disb_id.get(round(cr), [])
                        if disb_cands and name_did in disb_cands:
                            tx_type = 'Collection'; tx_ref = name_did
                            tx_basis = 'Amount+Name'; tx_notes = 'Name+amount match'
                        else:
                            tx_type = 'Collection'; tx_ref = name_did
                            tx_basis = 'Name'; tx_notes = 'Fuzzy name — confirm'
                    else:
                        if cr < 5000 or re.search(r'ft\s+-\s+cr|capital|transfer from', desc, re.IGNORECASE):
                            tx_type = 'Capital In'; tx_notes = '—'
                        else:
                            tx_type = 'Collection'; tx_ref = ''
                            tx_basis = '—'; tx_notes = 'Review — no match found'

        # Honour manual override from widget
        override = tx.get('type_override', '').strip()
        if override and override != 'Skip':
            tx_type  = override
            tx_basis = 'Manual'
        row_rem = tx.get('row_remarks', '').strip()
        if row_rem:
            tx_notes = f"{tx_notes} | {row_rem}" if tx_notes else row_rem

        result.append({**tx, 'type': tx_type, 'matched_ref': tx_ref,
                        'match_basis': tx_basis, 'match_notes': tx_notes})
    return result

# ── Sheet 1: Statement ────────────────────────────────────────────────────────

def _sheet_statement(wb, period_label, remarks, opening, closing, classified_txns, s):
    sname = 'Statement'
    if sname not in wb.sheetnames:
        ws = wb.create_sheet(sname)
        _title(ws, f'BridgeLine Partners — {remarks or "HDFC Bank A/c"} | {period_label} Reconciliation', 10, s)
        _legend(ws, 'Capital In | Disbursement | Collection | Expense  —  Type auto-classified. Update Matched Book Ref / Match Basis manually where needed.', 10, s)
        cols = [('Date',12),('Narration',50),('UTR / Ref No.',24),
                ('Withdrawal (₹)',16),('Deposit (₹)',16),('Closing Bal (₹)',16),
                ('Type',24),('Matched Book Ref',18),('Match Basis',14),('Match Source / Notes',28)]
        _hdr(ws, 3, cols, s)
        next_row = 4
    else:
        ws = wb[sname]
        next_row = ws.max_row + 2

    ws.merge_cells(f'A{next_row}:J{next_row}')
    c = ws.cell(next_row, 1,
        f"▶  {period_label}  —  Opening: ₹{opening:,.2f}  |  Closing: ₹{closing:,.2f}  |  {remarks or ''}")
    c.fill = s['hdr_fill']; c.font = Font(bold=True, size=9, name='Arial', color='FFFFFF')
    c.alignment = Alignment(horizontal='left')
    next_row += 1

    type_fill = {
        'Disbursement':           s['amb_fill'],
        'Collection':             s['grn_fill'],
        'Collection (via Pradaan)':s['grn_fill'],
        'Expense':                s['red_fill'],
        'Capital In':             s['blu_fill'],
        'Capital Out':            s['blu_fill'],
        'FD Booking':             s['blu_fill'],
    }
    total_dr = total_cr = 0.0
    for tx in classified_txns:
        total_dr += tx['debit']; total_cr += tx['credit']
        fill = type_fill.get(tx['type'], PatternFill("solid", fgColor="FFFFFF"))
        for ci, val in enumerate([
            tx['date'], tx['description'], tx.get('utr',''),
            _n(tx['debit']), _n(tx['credit']), _n(tx['balance']),
            tx['type'], tx['matched_ref'], tx['match_basis'], tx['match_notes']
        ], 1):
            c = ws.cell(next_row, ci, val)
            c.fill = fill; c.font = s['norm_font']; c.border = s['thin_border']
            if ci in (4,5,6): c.alignment = Alignment(horizontal='right')
        next_row += 1

    for ci, val in enumerate(['','TOTALS','',_n(total_dr),_n(total_cr),'','','','',''], 1):
        c = ws.cell(next_row, ci, val)
        c.fill = s['tot_fill']; c.font = s['tot_font']; c.border = s['thin_border']
        if ci in (4,5): c.alignment = Alignment(horizontal='right')

    ws.freeze_panes = 'A4'
    return total_dr, total_cr

# ── Sheet 2: Disbursement Recon ───────────────────────────────────────────────

def _sheet_disb_recon(wb, period_label, records, classified_txns, s, period_month, remarks_map=None):
    """Only disbursements from the current period month."""
    sname = 'Disbursement Recon'
    cols = [('Disb ID',18),('Disb Date',14),('Customer',26),('Book Amount (₹)',16),
            ('Charges (₹)',12),('GST (₹)',10),('Bank Date',12),('Bank Withdrawal (₹)',18),
            ('Disb UTR',26),('Match Basis',14),('Status',18),('Remarks',34)]
    ncols = len(cols)
    if sname not in wb.sheetnames:
        ws = wb.create_sheet(sname)
        _title(ws, 'Disbursement Reconciliation — Books vs Bank Debits', ncols, s)
        _legend(ws, 'Green = UTR matched  |  Amber = amount/date matched  |  Red = not found in bank', ncols, s)
        _hdr(ws, 3, cols, s)
        row = 4
    else:
        ws = wb[sname]
        row = ws.max_row + 2

    # Period separator
    ws.merge_cells(f'A{row}:{get_column_letter(ncols)}{row}')
    c = ws.cell(row, 1, f'▶  {period_label}')
    c.fill = s['hdr_fill']; c.font = Font(bold=True, size=9, name='Arial', color='FFFFFF')
    c.alignment = Alignment(horizontal='left')
    row += 1

    bank_by_ref = {}
    for tx in classified_txns:
        if tx['debit'] > 0 and tx['matched_ref']:
            bank_by_ref.setdefault(tx['matched_ref'], tx)

    for r in records:
        ddate = str(r.get('Disbursement Date','') or '')
        did   = r.get('Disbursement ID','')
        # Only show June 2026 disbursements (or any that matched in bank this period)
        if not _date_in_period(ddate, period_month) and did not in bank_by_ref:
            continue

        amt     = _to_num(r.get('Amount', 0))
        charges = _to_num(r.get('Charges', 0) or r.get('Processing Charges', 0))
        gst     = _to_num(r.get('GST(18%)', 0) or r.get('GST', 0))
        utr     = str(r.get('Debit Note','') or '').strip()
        bank    = bank_by_ref.get(did)
        if bank:
            basis  = bank['match_basis']
            bdelta = abs(amt - bank['debit'])
            status = 'Matched' if bdelta <= 1 else f'Matched (Δ₹{bdelta:,.0f})'
            fill   = s['grn_fill'] if basis == 'UTR' else s['amb_fill']
            bdate  = bank['date']; bamt = _n(bank['debit'])
        else:
            basis = 'Not found'; status = 'Not found in bank'; fill = s['red_fill']
            bdate = bamt = ''

        remark = (remarks_map or {}).get(did, '')
        _row(ws, row, [did, ddate, r.get('Customer Name',''),
                        _n(amt), _n(charges), _n(gst),
                        bdate, bamt, utr, basis, status, remark],
             s, right_cols=(4,5,6,8), fill=fill)
        row += 1

    ws.freeze_panes = 'A4'


# ── Sheet 3: Collection Recon ─────────────────────────────────────────────────

def _sheet_coll_recon(wb, period_label, records, mc_rows, classified_txns, s, period_month, remarks_map=None):
    sname = 'Collection Recon'
    cols = [('Disb ID',18),('Customer',24),('Branch',14),('Disbursed (₹)',16),
            ('Status',10),('Pay Date',12),('Payment (₹)',16),('Collection UTR',24),
            ('Narration (short)',36),('Collected (₹)',16),('Outstanding (₹)',16),('Remarks',34)]
    if sname not in wb.sheetnames:
        ws = wb.create_sheet(sname)
        _title(ws, f'Collection Reconciliation — Payment level', 12, s)
        _legend(ws, 'Green = Closed  |  Amber = Partial.  One row per payment.', 12, s)
        _hdr(ws, 3, cols, s)
        row = 4
    else:
        ws = wb[sname]
        row = ws.max_row + 2

    # Period separator
    ws.merge_cells(f'A{row}:L{row}')
    c = ws.cell(row, 1, f'\u25b6  {period_label}')
    c.fill = s['hdr_fill']; c.font = Font(bold=True, size=9, name='Arial', color='FFFFFF')
    c.alignment = Alignment(horizontal='left')
    row += 1

    # Bank credits indexed by UTR for narration lookup; also by amount for
    # backfilling a UTR when the M Coll note has none embedded.
    bank_cr_by_utr = {}
    bank_cr_by_amt = {}
    for tx in classified_txns:
        if tx['credit'] > 0:
            if tx.get('utr'):
                bank_cr_by_utr[tx['utr'].upper()] = tx
            bank_cr_by_amt.setdefault(round(tx['credit']), []).append(tx)

    # M Coll grouped by disb_id — filter to payments made this period
    mc_map = {}
    for mc in mc_rows:
        did  = (mc[0] if len(mc)>0 else '').strip()
        pdate= (mc[1] if len(mc)>1 else '')
        if did and (not period_month or _date_in_period(str(pdate), period_month)):
            mc_map.setdefault(did, []).append(mc)

    # Also include any disb that has a bank match this period even without M Coll entry
    coll_disb_ids = set(mc_map.keys())
    for tx in classified_txns:
        if tx['credit'] > 0 and tx['matched_ref']:
            coll_disb_ids.add(tx['matched_ref'])

    for r in records:
        did = r.get('Disbursement ID','')
        if did not in coll_disb_ids:
            continue

        amt, total, collected, balance = _parse_case(r)
        status   = r.get('Overdue Status','').strip()
        sfill    = s['grn_fill'] if status == 'Closed' else s['amb_fill']
        payments = mc_map.get(did, [])

        if not payments:
            # Matched via bank credit only — add a single row
            bank_tx = next((tx for tx in classified_txns
                            if tx['credit'] > 0 and tx['matched_ref'] == did), None)
            if bank_tx:
                _row(ws, row, [
                    did, r.get('Customer Name',''), r.get('Branch',''),
                    _n(total), status,
                    bank_tx['date'], _n(bank_tx['credit']), bank_tx.get('utr',''),
                    bank_tx['description'][:50], _n(collected), _n(balance),
                    (remarks_map or {}).get(did, '')
                ], s, right_cols=(4,7,10,11), fill=sfill)
                row += 1
            continue

        first = True
        for mc in payments:
            # M Coll col D holds `utr or raw_msg` (see save_repayment) — pull
            # the UTR out of SMS text when that's what landed there, and if
            # there's still none, backfill from a unique same-amount bank
            # credit in this period.
            mc_raw  = str(mc[3] if len(mc)>3 else '').strip()
            mc_amt  = _clean_amount(mc[2]) if len(mc)>2 else 0
            mc_utr  = extract_utr(mc_raw)
            if not mc_utr and re.fullmatch(r'[A-Za-z0-9\-/]{6,25}', mc_raw or ''):
                mc_utr = mc_raw  # already a bare ref, just not SMS-shaped
            bank_tx = bank_cr_by_utr.get(mc_utr.upper()) if mc_utr else None
            if not mc_utr:
                amt_matches = bank_cr_by_amt.get(round(mc_amt), [])
                if len(amt_matches) == 1:
                    bank_tx = amt_matches[0]
                    mc_utr  = bank_tx.get('utr', '')
            narr    = bank_tx['description'][:50] + '…' if bank_tx else '—'

            _row(ws, row, [
                did if first else '',
                r.get('Customer Name','') if first else '',
                r.get('Branch','') if first else '',
                _n(total) if first else '',
                status if first else '',
                mc[1] if len(mc)>1 else '', _n(mc_amt), mc_utr, narr,
                _n(collected) if first else '',
                _n(balance) if first else '',
                (remarks_map or {}).get(did, '') if first else '',
            ], s, right_cols=(4,7,10,11), fill=sfill)
            first = False; row += 1

    ws.freeze_panes = 'A4'

# ── Sheet 4: Expenses ─────────────────────────────────────────────────────────

def _sheet_expenses(wb, period_label, classified_txns, s):
    sname = 'Expenses'
    cols = [('Date',12),('Narration',54),('UTR / Ref No.',24),('Amount (₹)',16),('Category',28)]
    if sname not in wb.sheetnames:
        ws = wb.create_sheet(sname)
        _title(ws, f'Operational Expenses', 5, s)
        _legend(ws, 'Auto-extracted — POS, subscriptions, bank charges, salaries, partner share.', 5, s)
        _hdr(ws, 3, cols, s)
        row = 4
    else:
        ws = wb[sname]
        row = ws.max_row + 2

    # Period separator
    ws.merge_cells(f'A{row}:E{row}')
    c = ws.cell(row, 1, f'\u25b6  {period_label}')
    c.fill = s['hdr_fill']; c.font = Font(bold=True, size=9, name='Arial', color='FFFFFF')
    c.alignment = Alignment(horizontal='left')
    row += 1

    expenses = [tx for tx in classified_txns if tx['type'] == 'Expense' and tx['debit'] > 0]
    total = 0.0
    for tx in expenses:
        total += tx['debit']
        _row(ws, row, [
            tx['date'], tx['description'],
            tx.get('utr','') or extract_utr(tx['description']),
            _n(tx['debit']), tx['match_notes']
        ], s, right_cols=(4,))
        row += 1

    for ci, val in enumerate(['','TOTAL','',_n(total),''], 1):
        c = ws.cell(row, ci, val)
        c.fill = s['tot_fill']; c.font = s['tot_font']; c.border = s['thin_border']
        if ci == 4: c.alignment = Alignment(horizontal='right')

    ws.freeze_panes = 'A4'
    return expenses, total

# ── Sheet 5: Mapped (Disb ID ↔ Disb UTR ↔ Collection UTR) ───────────────

def _sheet_mapped(wb, period_label, records, mc_rows, classified_txns, s, period_month=None):
    """Only cases with activity in the period month: disbursed in-month, an
    M Coll payment in-month, or a bank debit/credit match in this period's
    statement. Without this filter every case ever (incl. archive tabs)
    repeats in each period section."""
    sname = 'Mapped'
    cols = [('Disb ID',18),('Customer',26),('Disb Date',12),('Disb Amount (₹)',16),
            ('Disb UTR (Bank)',26),('Coll Date',12),('Coll Amount (₹)',16),
            ('Collection UTR',26),('Match Basis',14),('Status',18)]
    if sname not in wb.sheetnames:
        ws = wb.create_sheet(sname)
        _title(ws, 'UTR Mapping — Disbursements & Collections', 10, s)
        _legend(ws, 'One row per payment event. Disb UTR from bank debits; Coll UTR from bank credits.', 10, s)
        _hdr(ws, 3, cols, s)
        row = 4
    else:
        ws = wb[sname]
        row = ws.max_row + 2

    ws.merge_cells(f'A{row}:J{row}')
    c = ws.cell(row, 1, f'▶  {period_label}')
    c.fill = s['hdr_fill']; c.font = Font(bold=True, size=9, name='Arial', color='FFFFFF')
    c.alignment = Alignment(horizontal='left')
    row += 1

    bank_debit_by_ref  = {}
    bank_credit_by_utr = {}
    bank_credit_refs   = set()
    bank_cr_by_amt     = {}
    for tx in classified_txns:
        if tx['debit'] > 0 and tx['matched_ref']:
            bank_debit_by_ref[tx['matched_ref']] = tx
        if tx['credit'] > 0:
            if tx.get('utr'):
                bank_credit_by_utr[tx['utr'].upper()] = tx
            if tx['matched_ref']:
                bank_credit_refs.add(tx['matched_ref'])
            bank_cr_by_amt.setdefault(round(tx['credit']), []).append(tx)

    # Payments filtered to the period month — pre-period instalments don't
    # belong in this period's mapping.
    mc_map = {}
    for mc in mc_rows:
        did   = (mc[0] if len(mc) > 0 else '').strip()
        pdate = str(mc[1] if len(mc) > 1 else '')
        if did and (not period_month or _date_in_period(pdate, period_month)):
            mc_map.setdefault(did, []).append(mc)

    for r in records:
        did   = r.get('Disbursement ID', '')
        ddate = str(r.get('Disbursement Date', '') or '')
        # Strictly this period: disbursed in-month, paid in-month, or matched
        # against this period's bank statement.
        if period_month and not (_date_in_period(ddate, period_month)
                                 or did in mc_map
                                 or did in bank_debit_by_ref
                                 or did in bank_credit_refs):
            continue
        amt   = _to_num(r.get('Amount', 0))
        books_utr = str(r.get('Debit Note', '') or '').strip()

        bank_dr   = bank_debit_by_ref.get(did)
        bank_utr  = bank_dr.get('utr', '') if bank_dr else books_utr
        bank_ddate = bank_dr['date'] if bank_dr else ddate
        dr_basis  = bank_dr['match_basis'] if bank_dr else ('Books' if books_utr else 'Not found')
        dr_fill   = s['grn_fill'] if bank_dr and bank_dr['match_basis'] == 'UTR' else (
                    s['amb_fill'] if bank_dr else s['red_fill'])

        payments = mc_map.get(did, [])
        if not payments:
            _row(ws, row, [did, r.get('Customer Name',''), bank_ddate, _n(amt),
                           bank_utr or books_utr, '', '', '', dr_basis,
                           r.get('Overdue Status','')], s, right_cols=(4,7,8), fill=dr_fill)
            row += 1
        else:
            first = True
            for mc in payments:
                # Same UTR handling as Collection Recon: col D may hold raw
                # SMS text — extract, else backfill from a unique same-amount
                # bank credit.
                mc_raw  = str(mc[3] if len(mc) > 3 else '').strip()
                mc_amt  = _clean_amount(mc[2]) if len(mc) > 2 else 0
                mc_date = mc[1] if len(mc) > 1 else ''
                mc_utr  = extract_utr(mc_raw)
                if not mc_utr and re.fullmatch(r'[A-Za-z0-9\-/]{6,25}', mc_raw or ''):
                    mc_utr = mc_raw  # already a bare ref, just not SMS-shaped
                bank_cr = bank_credit_by_utr.get(mc_utr.upper()) if mc_utr else None
                if not mc_utr:
                    amt_matches = bank_cr_by_amt.get(round(mc_amt), [])
                    if len(amt_matches) == 1:
                        bank_cr = amt_matches[0]
                        mc_utr  = bank_cr.get('utr', '')
                row_fill = (dr_fill if first else (s['grn_fill'] if bank_cr else s['amb_fill']))
                _row(ws, row, [
                    did if first else '',
                    r.get('Customer Name','') if first else '',
                    bank_ddate if first else '',
                    _n(amt) if first else '',
                    (bank_utr or books_utr) if first else '',
                    mc_date, _n(mc_amt), mc_utr,
                    dr_basis if first else '',
                    r.get('Overdue Status','') if first else '',
                ], s, right_cols=(4,7,8), fill=row_fill)
                first = False; row += 1

    ws.freeze_panes = 'A4'

# ── Main save function ────────────────────────────────────────────────────────

RECON_TXNS_SHEET_NAME = "Recon Txns"
RECON_TXNS_HEADERS = [
    "Batch ID", "Recon Date", "Account", "Period Label", "Opening", "Closing",
    "Remarks", "Txn Date", "Description", "UTR", "Debit", "Credit", "Balance",
    "Type", "Matched Ref", "Match Basis", "Match Notes"
]

def get_recon_txns_sheet(sh):
    try:
        return sh.worksheet(RECON_TXNS_SHEET_NAME)
    except Exception:
        ws = sh.add_worksheet(title=RECON_TXNS_SHEET_NAME, rows=5000, cols=len(RECON_TXNS_HEADERS))
        ws.append_row(RECON_TXNS_HEADERS)
        return ws

def _recon_num(v):
    try:
        return float(str(v).replace(',', '').strip() or 0)
    except Exception:
        return 0.0

def _load_recon_periods(ws):
    """Read the Recon Txns tab and return the stored periods sorted
    chronologically. Where the same (recon_date, account) was saved more than
    once, only the latest Batch ID wins — re-running a day's recon replaces
    that period instead of duplicating it."""
    all_vals = ws.get_all_values()
    # batches[(recon_date, account)] = {batch_id, meta, txns}
    batches = {}
    for row in all_vals[1:]:
        if len(row) < 14 or not row[0]:
            continue
        key = (row[1], row[2])
        batch_id = row[0]
        if key not in batches or batch_id > batches[key]['batch_id']:
            batches[key] = {'batch_id': batch_id,
                            'recon_date': row[1], 'account': row[2],
                            'period_label': row[3],
                            'opening': _recon_num(row[4]), 'closing': _recon_num(row[5]),
                            'remarks': row[6], 'txns': []}
        # Placeholder rows (period saved with zero transactions) carry the
        # metadata only — don't surface them as ledger lines.
        if batches[key]['batch_id'] == batch_id and (row[8] or _recon_num(row[10]) or _recon_num(row[11])):
            batches[key]['txns'].append({
                'date':        row[7],
                'description': row[8],
                'utr':         row[9],
                'debit':       _recon_num(row[10]),
                'credit':      _recon_num(row[11]),
                'balance':     _recon_num(row[12]),
                'type':        row[13],
                'matched_ref': row[14] if len(row) > 14 else '',
                'match_basis': row[15] if len(row) > 15 else '',
                'match_notes': row[16] if len(row) > 16 else '',
            })

    def _period_sort_key(p):
        try:
            return datetime.strptime(p['recon_date'], '%d-%m-%Y')
        except Exception:
            return datetime.min
    return sorted(batches.values(), key=_period_sort_key)

def save_reconciliation(recon_date, opening_balance, closing_balance, transactions,
                         remarks='', remarks_map=None, account=''):
    """Persists this period's classified transactions to the 'Recon Txns'
    sheet tab, then rebuilds the FULL cumulative reconciliation workbook from
    every stored period and returns its bytes for an in-browser download.

    The Google Sheet is the system of record — the user never re-uploads a
    previous file to keep history; every download contains all periods saved
    so far. Saving the same (recon_date, account) again replaces that period
    (latest Batch ID wins in _load_recon_periods), so corrections are safe.

    Each call also appends one summary row to the 'Recon Log' tab so
    get_latest_bank_balance() / the dashboard's bank-balance figure keeps
    working without reading back any saved file.
    """
    try:
        records = read_accounts_from_gsheet()
    except Exception:
        records = []
    try:
        sh      = get_gspread_client().open_by_key(SPREADSHEET_ID)
        mc_ws   = sh.worksheet('M Coll')
        mc_rows = mc_ws.get_all_values()[1:]
    except Exception:
        sh      = get_gspread_client().open_by_key(SPREADSHEET_ID)
        mc_rows = []

    # Derive period label from date  e.g. "Jun 2026"
    try:
        period_label = datetime.strptime(recon_date, '%d-%m-%Y').strftime('%b %Y')
    except Exception:
        period_label = recon_date

    # Classify and match bank transactions; drop any manually marked Skip
    classified = [tx for tx in _match_transactions(transactions, records, mc_rows)
                  if tx.get('type_override', '') != 'Skip']
    rm = remarks_map or {}

    # Persist this period's classified rows to the Recon Txns tab (one batch
    # write), then read everything back so current + historical periods flow
    # through one rebuild path.
    txns_ws = get_recon_txns_sheet(sh)
    batch_id = datetime.now().strftime('%Y%m%d%H%M%S')
    txns_ws.append_rows([[
        batch_id, recon_date, account, period_label,
        opening_balance, closing_balance, remarks,
        tx['date'], tx['description'], tx.get('utr', ''),
        tx['debit'], tx['credit'], tx['balance'],
        tx['type'], tx['matched_ref'], tx['match_basis'], tx['match_notes'],
    ] for tx in classified] or [[batch_id, recon_date, account, period_label,
                                 opening_balance, closing_balance, remarks,
                                 '', '', '', 0, 0, 0, '', '', '', '']])
    periods = _load_recon_periods(txns_ws)

    s = _xl_styles()
    wb = openpyxl.Workbook()
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    # Rebuild every stored period in chronological order. The _sheet_*
    # builders append below existing content (ws.max_row + 2) when the sheet
    # already exists, so periods stack naturally. Historical periods use their
    # STORED classified values — never re-matched, so past results don't shift
    # as live sheet data changes. Mapped is built once (it's a full-book UTR
    # map with no period filter — per-period calls would duplicate the book).
    total_dr = total_cr = 0.0
    for p in periods:
        p_label = p['period_label']  # month scope, e.g. "Jul 2026" — drives filtering
        # Daily sections need the day in the header, otherwise consecutive
        # saves within a month are indistinguishable.
        day_label = f"{p['recon_date']} ({p_label})"
        sheet_label = f"{day_label} — {p['account']}" if p['account'] else day_label
        dr, cr = _sheet_statement(wb, sheet_label, p['remarks'], p['opening'],
                                  p['closing'], p['txns'], s)
        is_current = (p['recon_date'] == recon_date and p['account'] == account)
        if is_current:
            total_dr, total_cr = dr, cr
        _sheet_disb_recon(wb, day_label, records, p['txns'], s, p_label,
                          rm if is_current else {})
        _sheet_coll_recon(wb, day_label, records, mc_rows, p['txns'], s, p_label,
                          rm if is_current else {})
        _sheet_expenses(wb, day_label, p['txns'], s)
        _sheet_mapped(wb, day_label, records, mc_rows, p['txns'], s, p_label)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    try:
        log_ws = sh.worksheet(RECON_LOG_SHEET_NAME)
        log_ws.append_row([recon_date, opening_balance, closing_balance,
                            datetime.today().strftime('%d-%m-%Y %H:%M'), account])
    except Exception as e:
        print(f"[Recon Log] append failed: {e}")

    filename = f"Daily Reconciliation - {account or 'All Accounts'}.xlsx"

    return {
        'total_debit':  total_dr,
        'total_credit': total_cr,
        'closing':      closing_balance,
        'rows_saved':   len(transactions),
        'filename':     filename,
        'file_bytes':   buf.getvalue(),
    }

# ── Flask app ─────────────────────────────────────────────────────────────────

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB upload limit

SETUP_HTML = """<!DOCTYPE html><html><head><meta charset='UTF-8'>
<title>BridgeLine — Setup</title>
<style>body{font-family:Arial;max-width:680px;margin:40px auto;padding:20px;background:#eef2f7}
h1{color:#1a3a5c}
.step{background:white;border-radius:8px;padding:18px;margin:12px 0;box-shadow:0 1px 4px rgba(0,0,0,.1)}
a{color:#2d5986}
</style></head><body>
<h1>BridgeLine Widget</h1>
<div class='step'>This widget is hosted and always-on — there's nothing to set up.
Google Sheets access is configured once via environment variables on the server,
not per-device. If you're seeing an authorization error, contact the admin to
check the hosting environment's Google credentials rather than re-running any
local setup.</div>
<p><a href='/'>← Back to widget</a></p>
</body></html>"""

HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>BridgeLine Accounts</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:Arial,sans-serif;background:#eef2f7;color:#222}
  header{background:#1a3a5c;color:white;padding:18px 24px;display:flex;justify-content:space-between;align-items:center}
  header h1{font-size:1.4rem}
  header p{font-size:.85rem;color:#a8c4e0;margin-top:2px}
  header a{color:#a8c4e0;font-size:.8rem;text-decoration:none}
  .container{max-width:1360px;margin:24px auto;padding:0 16px}
  .tabs{display:flex;gap:6px;margin-bottom:0;flex-wrap:wrap}
  .tab-btn{padding:10px 18px;border:none;border-radius:8px 8px 0 0;background:#c8d8ec;color:#1a3a5c;font-size:.88rem;cursor:pointer;font-weight:600}
  .tab-btn.active{background:white}
  .tab-content{display:none;background:white;border-radius:0 8px 8px 8px;padding:24px;box-shadow:0 2px 8px rgba(0,0,0,.1)}
  .tab-content.active{display:block}
  .section{border:1px solid #d0dce8;border-radius:6px;padding:14px 16px;margin-bottom:16px;background:#f8fbff}
  .section h3{font-size:.82rem;color:#2d5986;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px}
  textarea{width:100%;border:1px solid #b0c4d8;border-radius:4px;padding:8px;font-size:.88rem;resize:vertical;min-height:80px;font-family:Arial}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  .field{display:flex;flex-direction:column;gap:4px}
  .field label{font-size:.8rem;font-weight:600;color:#444}
  input,select{width:100%;padding:7px 10px;border:1px solid #b0c4d8;border-radius:4px;font-size:.9rem;background:white}
  input:focus,select:focus{outline:none;border-color:#2d5986;box-shadow:0 0 0 2px #d0dce8}
  input[readonly]{background:#f0f4f8;color:#555}
  .btn{padding:9px 16px;border:none;border-radius:5px;font-size:.88rem;cursor:pointer;font-weight:600}
  .btn-extract{background:#e8f0fe;color:#1a3a5c;margin-top:8px}
  .btn-extract:hover{background:#c8d8f8}
  .btn-save{background:#1a5c3a;color:white;width:100%;padding:13px;font-size:1rem;margin-top:8px;border-radius:6px}
  .btn-save:hover{background:#14472d}
  .btn-lookup{background:#2d5986;color:white;padding:7px 14px}
  .lookup-row{display:flex;gap:8px;align-items:center}
  .lookup-row input{flex:1}
  .info-box{background:#e8f4ec;border:1px solid #a8d4b8;border-radius:4px;padding:8px 12px;font-size:.85rem;color:#1a5c3a;margin-top:8px;display:none}
  .info-box.error{background:#fde8e8;border-color:#f5b0b0;color:#a00}
  .status{margin-top:10px;padding:10px 14px;border-radius:5px;font-size:.88rem;display:none}
  .status.success{background:#e8f4ec;color:#1a5c3a;border:1px solid #a8d4b8}
  .status.error{background:#fde8e8;color:#a00;border:1px solid #f5b0b0}
  .hint{font-size:.75rem;color:#888;margin-top:2px}
  .spinner{display:none;color:#888;font-size:.85rem;margin-top:6px}
  /* Summary banner */
  .summary-bar{display:grid;grid-template-columns:repeat(5,1fr);gap:8px;margin-bottom:14px}
  .sum-card{background:white;border-radius:8px;padding:12px 14px;box-shadow:0 1px 4px rgba(0,0,0,.1);text-align:center}
  .sum-card .val{font-size:1.1rem;font-weight:700;color:#1a3a5c}
  .sum-card .lbl{font-size:.72rem;color:#888;margin-top:2px}
  .sum-card.clickable{cursor:pointer;transition:transform .08s,box-shadow .08s}
  .sum-card.clickable:hover{transform:translateY(-2px);box-shadow:0 4px 10px rgba(0,0,0,.15)}
  .sum-modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:999;align-items:center;justify-content:center}
  .sum-modal-bg.show{display:flex}
  .sum-modal{background:#fff;border-radius:10px;max-width:480px;width:92%;max-height:75vh;overflow-y:auto;box-shadow:0 8px 30px rgba(0,0,0,.3)}
  .sum-modal-head{background:#1a3a5c;color:#fff;padding:14px 18px;border-radius:10px 10px 0 0;display:flex;justify-content:space-between;align-items:center}
  .sum-modal-head h3{margin:0;font-size:1rem}
  .sum-modal-head button{background:none;border:none;color:#fff;font-size:1.3rem;cursor:pointer;line-height:1}
  .sum-modal table{width:100%;font-size:.85rem;border-collapse:collapse}
  .sum-modal th{text-align:left;padding:8px 14px;background:#f0f4f8;border-bottom:1px solid #dde}
  .sum-modal td{padding:7px 14px;border-bottom:1px solid #eef}
  .sum-modal tr:hover td{background:#f8fbff}
  .sum-modal .empty{padding:20px;text-align:center;color:#888;font-size:.85rem}
  /* Calculator */
  .calc-display{background:#1a3a5c;color:#fff;font-size:2rem;text-align:right;padding:16px 20px;border-radius:8px;margin-bottom:12px;min-height:70px;word-break:break-all;line-height:1.2}
  .calc-display .calc-expr{font-size:.9rem;color:#a8c4e0;min-height:22px;margin-bottom:4px}
  .calc-keys{display:grid;grid-template-columns:repeat(4,1fr);gap:8px}
  .calc-key{padding:16px 8px;border:none;border-radius:6px;font-size:1.1rem;font-weight:700;cursor:pointer;transition:opacity .1s}
  .calc-key:active{opacity:.7}
  .k-num{background:#e8f0fe;color:#1a3a5c}
  .k-op{background:#2d5986;color:white}
  .k-eq{background:#1a5c3a;color:white}
  .k-clear{background:#c00;color:white}
  .k-back{background:#b8860b;color:white}
  .k-zero{grid-column:span 2}
  /* Reconciliation */
  .recon-summary{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:12px 0}
  .recon-card{background:#1a3a5c;color:white;border-radius:8px;padding:12px;text-align:center}
  .recon-card .rv{font-size:1.1rem;font-weight:700;margin-bottom:2px}
  .recon-card .rl{font-size:.72rem;color:#a8c4e0}
  .recon-table{width:100%;font-size:.82rem;border-collapse:collapse;margin-top:10px}
  .recon-table th{background:#1a3a5c;color:white;padding:7px 10px;text-align:left}
  .recon-table td{padding:6px 10px;border-bottom:1px solid #eef2f7}
  .recon-table tr:hover td{background:#f8fbff}
  .upload-area{border:2px dashed #b0c4d8;border-radius:8px;padding:30px;text-align:center;color:#888;cursor:pointer;margin-bottom:12px}
  .upload-area:hover{border-color:#2d5986;color:#2d5986;background:#f0f4fc}
  .upload-area input[type=file]{display:none}
  /* Settings */
  .settings-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}
  /* Date picker */
  .date-wrap{position:relative}
  .date-wrap input[type=text]{padding-right:32px;box-sizing:border-box;width:100%}
  .date-cal-btn{position:absolute;right:4px;top:50%;transform:translateY(-50%);background:none;border:none;cursor:pointer;font-size:1.05rem;color:#2d5986;padding:4px;line-height:1}
  .date-cal-btn:hover{color:#1a3a5c}
  .date-native{position:absolute;right:4px;top:50%;transform:translateY(-50%);width:24px;height:24px;opacity:0;border:none;padding:0}
</style>
</head>
<body>
<header>
  <div><h1>BridgeLine Partners</h1><p>Accounts Entry Widget</p></div>
  <a href="/setup" target="_blank">⚙ Setup / Help</a>
</header>
<div class="container">
  <!-- Summary Banner -->
  <div class="summary-bar" id="summary-bar">
    <div class="sum-card clickable" onclick="openSumModal('disb')"><div class="val" id="s-disb">—</div><div class="lbl">Disbursed Today</div></div>
    <div class="sum-card clickable" onclick="openSumModal('coll')"><div class="val" id="s-coll">—</div><div class="lbl">Collected Today</div></div>
    <div class="sum-card clickable" onclick="openSumModal('out')"><div class="val" id="s-out">—</div><div class="lbl">Total Outstanding</div></div>
    <div class="sum-card clickable" onclick="openSumModal('avail')"><div class="val" id="s-avail">—</div><div class="lbl">Available to Disburse</div></div>
    <div class="sum-card"><div class="val" id="s-date">—</div><div class="lbl">As of</div></div>
  </div>

  <div class="sum-modal-bg" id="sum-modal-bg" onclick="if(event.target===this) closeSumModal()">
    <div class="sum-modal">
      <div class="sum-modal-head"><h3 id="sum-modal-title">Details</h3><button onclick="closeSumModal()">&times;</button></div>
      <div id="sum-modal-body"></div>
    </div>
  </div>

  <div class="tabs">
    <button class="tab-btn active" onclick="showTab('disb',this); loadPendingRequests()">➕ New Disbursement</button>
    <button class="tab-btn" onclick="showTab('repa',this); loadOpenCases()">💰 Repayment</button>
    <button class="tab-btn" onclick="showTab('bulk',this); loadBulkRequests()">💸 Bank File</button>
    <button class="tab-btn" onclick="showTab('calc',this)">🧮 Calculator</button>
    <button class="tab-btn" onclick="showTab('recon',this)">🏦 Bank Reconciliation</button>
    <button class="tab-btn" onclick="showTab('invoice',this); loadInvoiceCases()">📄 Invoice</button>
    <button class="tab-btn" onclick="showTab('edit',this); loadRecentCases()">✏️ Edit Case</button>
    <button class="tab-btn" onclick="showTab('contacts',this); loadContacts()">👥 Contacts</button>
    <button class="tab-btn" onclick="showTab('settings',this); loadSettings()">⚙ Settings</button>
    <button class="tab-btn" id="mis-btn" onclick="generateMis()" style="background:#1a3a5c;color:#fff">📦 Generate MIS Package</button>
  </div>
  <div id="mis-status" style="padding:0 18px;font-size:.85rem"></div>

  <!-- DISBURSEMENT TAB -->
  <div id="disb" class="tab-content active">
    <div class="section" id="pending-requests-section">
      <h3 style="cursor:pointer;user-select:none;" onclick="togglePendingReqs()">
        📥 Pending Field Requests <span id="pending-count" style="font-size:13px;color:#888;font-weight:400;"></span>
        <span id="pending-chevron" style="float:right;font-size:13px;">▼</span>
      </h3>
      <div id="pending-requests-body">
        <div id="pending-requests-list" style="margin-top:8px;"></div>
        <div id="pending-selected-info" style="display:none;margin-top:10px;padding:10px 12px;background:#f0f7ff;border:1px solid #b3d4f5;border-radius:6px;font-size:13px;line-height:1.6;"></div>
        <input type="hidden" id="d-request-id">
      </div>
    </div>
    <div class="section">
      <h3>1 — Paste Customer / Disbursement Note</h3>
      <textarea id="d-msg" placeholder="Paste customer note, WhatsApp message, or disbursement details here..."></textarea>
      <button class="btn btn-extract" onclick="extractDisb()">⚡ Extract Details</button>
    </div>
    <div class="section">
      <h3>2 — Paste Bank Confirmation (for UTR)</h3>
      <textarea id="d-utr-msg" placeholder="Paste your bank's outgoing transfer SMS here to extract UTR..."></textarea>
      <button class="btn btn-extract" onclick="extractDisbUTR()">⚡ Extract UTR</button>
    </div>
    <div class="section">
      <h3>Disbursement Details</h3>
      <div class="grid">
        <div class="field"><label>Date (DD-MM-YYYY) *</label>
          <div class="date-wrap">
            <input type="text" id="d-date">
            <input type="date" class="date-native" id="d-date-native" onchange="_pickDate('d-date', this.value)">
            <button type="button" class="date-cal-btn" onclick="_openDatePicker('d-date-native')">📅</button>
          </div>
        </div>
        <div class="field"><label>Customer Name *</label><input type="text" id="d-customer" placeholder="Full name"></div>
        <div class="field"><label>Company *</label>
          <input type="text" id="d-company" list="company-list" placeholder="Select or type new...">
          <datalist id="company-list">""" + \
          "".join(f'<option value="{c}">' for c in COMPANIES) + """</datalist></div>
        <div class="field"><label>Cluster *</label>
          <input type="text" id="d-cluster" list="cluster-list" placeholder="Select or type new...">
          <datalist id="cluster-list">""" + \
          "".join(f'<option value="{c}">' for c in CLUSTERS) + """</datalist></div>
        <div class="field"><label>Branch *</label>
          <input type="text" id="d-branch" list="branch-list" placeholder="Select or type new...">
          <datalist id="branch-list">""" + \
          "".join(f'<option value="{b}">' for b in BRANCHES) + """</datalist></div>
        <div class="field"><label>Cheque No. (optional)</label><input type="text" id="d-chq"></div>
        <div class="field"><label>Amount (₹) *</label>
          <input type="number" id="d-amount" placeholder="e.g. 500000" oninput="calcCharges()"></div>
        <div class="field"><label>Charges (0.5%)</label>
          <input type="text" id="d-charges" readonly><span class="hint">Auto-calculated</span></div>
        <div class="field"><label>GST 18% on Charges</label>
          <input type="text" id="d-gst" readonly><span class="hint">Auto-calculated</span></div>
        <div class="field"><label>Total Receivable (₹)</label>
          <input type="text" id="d-total" readonly><span class="hint">Auto-calculated</span></div>
        <div class="field"><label>Serviced Branch</label>
          <input type="text" id="d-srv-branch" list="srv-branch-list" placeholder="Select or type new...">
          <datalist id="srv-branch-list">""" + \
          "".join(f'<option value="{b}">' for b in BRANCHES) + """</datalist></div>
        <div class="field"><label>Serviced Cluster</label>
          <input type="text" id="d-srv-cluster" list="srv-cluster-list" placeholder="Select or type new...">
          <datalist id="srv-cluster-list">""" + \
          "".join(f'<option value="{c}">' for c in CLUSTERS) + """</datalist></div>
        <div class="field" style="grid-column:1/-1"><label>Disbursement UTR / Debit Note</label>
          <input type="text" id="d-utr" placeholder="Auto-extracted or enter manually"></div>
        <div class="field" style="grid-column:1/-1"><label>Remarks</label>
          <input type="text" id="d-remarks" placeholder="e.g. Urgent case, referred by X, special rate approved…"></div>
      </div>
    </div>
    <button class="btn btn-save" onclick="saveDisb()">✅ Save Disbursement to Google Sheet</button>
    <div id="d-status" class="status"></div>
  </div>

  <!-- REPAYMENT TAB -->
  <div id="repa" class="tab-content">
    <div class="section">
      <h3>Paste Bank SMS / Payment Message</h3>
      <textarea id="r-msg" placeholder="Paste HDFC/SBI/IMPS bank message here..."></textarea>
      <button class="btn btn-extract" onclick="extractRepa()">⚡ Extract from Message</button>
    </div>
    <div class="section">
      <h3>Select Open Case</h3>
      <select id="r-open-cases" onchange="onCaseSelect()" style="width:100%;font-size:.85rem">
        <option value="">Loading open cases...</option>
      </select>
      <div class="spinner" id="cases-spinner">Loading...</div>
      <div id="r-info" class="info-box"></div>
    </div>
    <div class="section">
      <h3>Repayment Details</h3>
      <div class="grid">
        <div class="field"><label>Collection Date (DD-MM-YYYY) *</label>
          <div class="date-wrap">
            <input type="text" id="r-date">
            <input type="date" class="date-native" id="r-date-native" onchange="_pickDate('r-date', this.value)">
            <button type="button" class="date-cal-btn" onclick="_openDatePicker('r-date-native')">📅</button>
          </div>
        </div>
        <div class="field"><label>Amount Received (₹) *</label>
          <input type="number" id="r-amount" placeholder="e.g. 603540"></div>
        <div class="field"><label>UTR / Reference No.</label>
          <input type="text" id="r-utr" placeholder="e.g. BKIDR52026..."></div>
        <div class="field"><label>Discount (₹)</label>
          <input type="number" id="r-discount" placeholder="0" value="0"></div>
        <div class="field" style="grid-column:1/-1"><label>Remarks</label>
          <input type="text" id="r-remarks" placeholder="e.g. Part payment, cheque cleared, customer requested receipt…"></div>
      </div>
    </div>
    <button class="btn btn-save" onclick="saveRepa()">✅ Save Repayment to Google Sheet</button>
    <div id="r-status" class="status"></div>
  </div>

  <!-- CALCULATOR TAB -->
  <div id="calc" class="tab-content">
    <div style="max-width:360px;margin:0 auto">
      <div class="calc-display">
        <div class="calc-expr" id="calc-expr"></div>
        <div id="calc-disp">0</div>
      </div>
      <div class="calc-keys">
        <button class="calc-key k-clear" onclick="calcClear()">C</button>
        <button class="calc-key k-back"  onclick="calcBack()">⌫</button>
        <button class="calc-key k-op"    onclick="calcOp('%')">%</button>
        <button class="calc-key k-op"    onclick="calcOp('/')">÷</button>

        <button class="calc-key k-num"   onclick="calcNum('7')">7</button>
        <button class="calc-key k-num"   onclick="calcNum('8')">8</button>
        <button class="calc-key k-num"   onclick="calcNum('9')">9</button>
        <button class="calc-key k-op"    onclick="calcOp('*')">×</button>

        <button class="calc-key k-num"   onclick="calcNum('4')">4</button>
        <button class="calc-key k-num"   onclick="calcNum('5')">5</button>
        <button class="calc-key k-num"   onclick="calcNum('6')">6</button>
        <button class="calc-key k-op"    onclick="calcOp('-')">−</button>

        <button class="calc-key k-num"   onclick="calcNum('1')">1</button>
        <button class="calc-key k-num"   onclick="calcNum('2')">2</button>
        <button class="calc-key k-num"   onclick="calcNum('3')">3</button>
        <button class="calc-key k-op"    onclick="calcOp('+')">+</button>

        <button class="calc-key k-num k-zero" onclick="calcNum('0')">0</button>
        <button class="calc-key k-num"   onclick="calcNum('.')">.</button>
        <button class="calc-key k-eq"    onclick="calcEq()">=</button>
      </div>
      <div style="margin-top:14px;text-align:center;font-size:.8rem;color:#888">
        Tip: results are auto-formatted in Indian numbering (₹ lakhs / crores)
      </div>
    </div>
  </div>

  <!-- BANK RECONCILIATION TAB -->
  <div id="recon" class="tab-content">
    <div class="section">
      <h3>Daily Bank Statement Upload</h3>
      <div class="grid" style="margin-bottom:12px">
        <div class="field">
          <label>Reconciliation Date *</label>
          <div class="date-wrap">
            <input type="text" id="rec-date" placeholder="DD-MM-YYYY">
            <input type="date" class="date-native" id="rec-date-native" onchange="_pickDate('rec-date', this.value)">
            <button type="button" class="date-cal-btn" onclick="_openDatePicker('rec-date-native')">📅</button>
          </div>
        </div>
        <div class="field">
          <label>Opening Balance (₹) <span style="color:#888;font-weight:400">(auto-read from statement)</span></label>
          <input type="number" id="rec-opening" placeholder="Auto-extracted from statement" step="0.01">
        </div>
        <div class="field">
          <label>Closing Balance (₹) <span style="color:#888;font-weight:400">(auto-read from statement)</span></label>
          <input type="number" id="rec-closing" placeholder="Auto-extracted from statement" step="0.01" readonly style="background:#f0f4f8">
        </div>
        <div class="field">
          <label>Bank Account *</label>
          <input type="text" id="rec-account" list="rec-account-list" placeholder="Select or type new...">
          <datalist id="rec-account-list"></datalist>
        </div>
        <div class="field" style="grid-column:1/-1">
          <label>Remarks / Account Name</label>
          <input type="text" id="rec-remarks" placeholder="e.g. HDFC Current A/c XX0923">
        </div>
      </div>
      <label style="font-size:.8rem;font-weight:600;color:#444;display:block;margin-bottom:6px">Upload Bank Statement (CSV or Excel)</label>
      <div class="upload-area" id="upload-area" onclick="document.getElementById('rec-file').click()">
        <input type="file" id="rec-file" accept=".csv,.xlsx,.xls" onchange="onFileSelect(this)">
        <div id="upload-label">📂 Click to upload bank statement<br><span style="font-size:.78rem">Supports CSV, Excel (.xlsx/.xls)</span></div>
      </div>
      <button class="btn btn-extract" style="margin:0" onclick="parseStatement()">⚡ Parse Statement</button>
    </div>

    <!-- Reconciliation output -->
    <div id="recon-preview-section" style="display:none">

      <!-- Balances summary strip -->
      <div class="recon-summary" id="recon-summary" style="margin-bottom:14px"></div>

      <!-- Auto-reconciled confidence strip -->
      <div class="section" style="background:#f0faf4;border-color:#a8d5b5">
        <h3 style="color:#1a5c3a">✅ Auto-Reconciled <span id="confident-badge" style="font-weight:400;color:#555"></span></h3>
        <div id="confident-summary" style="font-size:.83rem;color:#333;line-height:1.8"></div>
      </div>

      <!-- Review queue — only uncertain entries -->
      <div class="section" id="review-section" style="display:none">
        <h3>⚠️ Needs Your Input <span id="review-badge" style="font-weight:400;color:#888"></span></h3>
        <p style="font-size:.8rem;color:#666;margin:0 0 10px">These couldn't be confidently matched. Set the correct <b>Type</b> and optionally add <b>Remarks</b> — then save.</p>
        <div style="overflow-x:auto">
          <table class="recon-table">
            <thead><tr>
              <th>Date</th><th>Description</th><th>UTR</th>
              <th style="text-align:right">Amount</th>
              <th style="width:60px">Dr/Cr</th>
              <th style="width:80px">Auto Type</th>
              <th style="width:190px">Correct Type</th>
              <th style="min-width:260px">Remarks</th>
            </tr></thead>
            <tbody id="review-tbody"></tbody>
          </table>
        </div>
      </div>

      <p style="font-size:.8rem;color:#666;margin:8px 0 4px">History is stored automatically — every download contains <b>all periods saved so far</b>. No need to re-upload previous files.</p>
      <button class="btn btn-save" onclick="saveRecon()" style="margin-top:4px">💾 Complete Reconciliation &amp; Save Excel</button>
      <div id="recon-status" class="status"></div>
    </div>
  </div>

  <!-- SETTINGS TAB -->
  <div id="contacts" class="tab-content">
    <div class="section">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;flex-wrap:wrap;gap:8px">
        <h3 style="margin:0">Staff Directory</h3>
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
          <input type="text" id="contact-search" placeholder="Search…"
            oninput="filterContacts(this.value)"
            style="padding:6px 10px;border:1px solid #ccd;border-radius:6px;font-size:.85rem;width:180px">
          <button class="btn" onclick="addContactRow()" style="padding:6px 14px;font-size:.82rem;background:#1a3a5c;color:#fff;border:none;border-radius:6px;cursor:pointer">+ Add Person</button>
          <button class="btn" onclick="addCluster()" style="padding:6px 14px;font-size:.82rem;background:#2e7d32;color:#fff;border:none;border-radius:6px;cursor:pointer">+ Add Cluster</button>
          <button class="btn btn-save" id="contacts-save-btn" onclick="saveContacts()" style="padding:6px 14px;font-size:.82rem;display:none">💾 Save</button>
        </div>
      </div>
      <div id="contacts-status" class="status"></div>
      <div id="contacts-body"></div>
    </div>
  </div>

  <!-- BANK FILE TAB -->
  <div id="bulk" class="tab-content">
    <div class="section">
      <h3>📥 Pending Requests <span id="bulk-count" style="font-size:13px;color:#888;font-weight:400;"></span>
        <button onclick="loadBulkRequests()" style="float:right;font-size:12px;padding:3px 10px;cursor:pointer;border:1px solid #ccc;border-radius:4px;background:#fff;">↻ Refresh</button></h3>
      <div style="overflow-x:auto;">
        <table id="bulk-table" style="width:100%;border-collapse:collapse;font-size:13px;display:none;">
          <thead><tr style="text-align:left;border-bottom:2px solid #dce8f5;">
            <th style="padding:6px 8px;"><input type="checkbox" id="bulk-select-all" onchange="toggleBulkSelectAll(this.checked)"></th>
            <th style="padding:6px 8px;">Request ID</th><th style="padding:6px 8px;">Submitted</th>
            <th style="padding:6px 8px;">Customer</th><th style="padding:6px 8px;">Cluster / Branch</th>
            <th style="padding:6px 8px;">Amount</th><th style="padding:6px 8px;">Account</th>
            <th style="padding:6px 8px;">IFSC</th><th style="padding:6px 8px;">Bank</th>
            <th style="padding:6px 8px;">Phone</th><th style="padding:6px 8px;">SO</th>
            <th style="padding:6px 8px;">Gold g</th>
          </tr></thead>
          <tbody id="bulk-tbody"></tbody>
        </table>
      </div>
      <div id="bulk-empty" style="color:#aaa;font-size:13px;">Loading...</div>
      <div id="bulk-footer" style="display:none;margin-top:10px;padding:8px 12px;background:#f0f7ff;border:1px solid #b3d4f5;border-radius:6px;font-size:13px;font-weight:600;"></div>
      <button class="btn" style="margin-top:10px;width:auto;padding:8px 18px;background:#fff;color:#c0392b;border:1px solid #e0b4b4;" onclick="deleteBulkRequests()">🗑 Delete Selected</button>
      <div id="bulk-delete-status" class="status"></div>
    </div>
    <div class="section">
      <h3>⚙ Export Bank File</h3>
      <div class="grid">
        <div class="field"><label>Bank Format *</label>
          <select id="bulk-bank" style="width:100%;padding:9px;border:1px solid #ccc;border-radius:6px;font-size:.9rem;">""" + \
          "".join(f'<option value="{k}">{v["label"]} (.{v["filetype"]})</option>' for k, v in BANK_TEMPLATES.items()) + """</select></div>
        <div class="field"><label>Debit Account</label>
          <input type="text" id="bulk-debit" placeholder="Your BridgeLine account no."></div>
        <div class="field"><label>Narration / Reference</label>
          <input type="text" id="bulk-narration" maxlength="20" placeholder="Shows on beneficiary statement"></div>
        <div class="field"><label>Value Date (DD/MM/YYYY)</label>
          <input type="text" id="bulk-value-date"></div>
      </div>
      <button class="btn btn-save" style="margin-top:12px" onclick="exportBulk()">💸 Generate &amp; Download</button>
      <div id="bulk-status" class="status"></div>
      <p style="font-size:12px;color:#888;margin-top:8px;">Selected requests are marked <b>Exported</b> in the sheet after download — they leave this list but stay visible in New Disbursement for recording. RTGS is auto-assigned for amounts ≥ ₹2,00,000, NEFT below.</p>
    </div>
    <div class="section">
      <h3>📋 Paste WhatsApp Request</h3>
      <p style="font-size:12px;color:#888;">For requests that arrive as a DISBURSEMENT REQUEST message instead of through the field app. Parsed requests join the same pending queue.</p>
      <textarea id="bulk-paste" placeholder="Paste the DISBURSEMENT REQUEST message here..."></textarea>
      <button class="btn btn-extract" onclick="parseBulkMsg()">⚡ Parse Message</button>
      <div id="bulk-parse-preview" style="display:none;margin-top:12px;">
        <div class="grid">
          <div class="field"><label>Customer Name *</label><input type="text" id="bp-customer"></div>
          <div class="field"><label>Amount (₹) *</label><input type="number" id="bp-amount"></div>
          <div class="field"><label>Account No *</label><input type="text" id="bp-account"></div>
          <div class="field"><label>IFSC *</label><input type="text" id="bp-ifsc" style="text-transform:uppercase"></div>
          <div class="field"><label>Bank Name</label><input type="text" id="bp-bank"></div>
          <div class="field"><label>Phone</label><input type="text" id="bp-phone"></div>
          <div class="field"><label>Cluster</label><input type="text" id="bp-cluster" list="cluster-list"></div>
          <div class="field"><label>Branch</label><input type="text" id="bp-branch" list="branch-list"></div>
          <div class="field"><label>SO Name</label><input type="text" id="bp-so"></div>
          <div class="field"><label>Gold Wt (gms)</label><input type="text" id="bp-gold"></div>
        </div>
        <div id="bulk-parse-warnings" style="margin-top:8px;font-size:12px;color:#b8860b;"></div>
        <button class="btn btn-save" style="margin-top:10px" onclick="addParsedRequest()">➕ Add to Queue</button>
      </div>
      <div id="bulk-paste-status" class="status"></div>
    </div>
  </div>

  <div id="settings" class="tab-content">
    <div class="section">
      <h3>General Settings</h3>
      <div style="display:flex;gap:12px;align-items:center;margin-bottom:12px">
        <label style="font-size:.85rem;font-weight:600">Daily Report Time:</label>
        <input type="time" id="report-time" style="width:120px">
      </div>
      <div style="display:flex;gap:12px;align-items:center;margin-bottom:12px">
        <label style="font-size:.85rem;font-weight:600">Bulk Export Debit Account:</label>
        <input type="text" id="set-bulk-debit" style="width:220px" placeholder="Default debit account no.">
      </div>
      <div style="display:flex;gap:12px;align-items:center;margin-bottom:12px">
        <label style="font-size:.85rem;font-weight:600">Bulk Export Narration:</label>
        <input type="text" id="set-bulk-narration" style="width:220px" maxlength="20">
      </div>
      <button class="btn btn-save" onclick="saveSettings()" style="width:auto;padding:10px 24px">💾 Save Settings</button>
      <div id="settings-status" class="status"></div>
    </div>
    <div class="section">
      <h3>Reconciliation &amp; MIS Reports</h3>
      <p style="font-size:.85rem;color:#555">The Daily Reconciliation Excel and the MIS PDF package are generated fresh on each run and download directly to your browser — keep your own copy if you want a running archive.</p>
    </div>
  </div>

  <!-- EDIT CASE TAB -->
  <div id="edit" class="tab-content">
    <div class="section">
      <h3>Load a Case</h3>
      <div style="display:flex;gap:10px;align-items:flex-end">
        <div class="field" style="flex:1"><label>Disbursement ID</label>
          <input type="text" id="e-disb-id" placeholder="e.g. BLP-010726-001" style="text-transform:uppercase" oninput="this.value=this.value.toUpperCase()"></div>
        <button class="btn btn-extract" style="width:auto;padding:10px 24px" onclick="loadEditCase()">🔍 Load Case</button>
      </div>
      <div id="e-load-status" class="status"></div>
    </div>
    <div class="section" id="e-recent-section">
      <h3 style="cursor:pointer;user-select:none;" onclick="toggleRecentCases()">
        🕐 Recently Recorded (last 7 days) <span id="e-recent-count" style="font-size:13px;color:#888;font-weight:400;"></span>
        <span id="e-recent-chevron" style="float:right;font-size:13px;">▼</span>
      </h3>
      <div id="e-recent-body"><div id="e-recent-list" style="margin-top:8px;"></div></div>
    </div>
    <div id="e-body" style="display:none">
      <div class="section">
        <h3>Disbursement Details <span id="e-case-id" style="font-weight:400;color:#888;font-size:13px"></span></h3>
        <div class="grid">
          <div class="field"><label>Date (DD-MM-YYYY)</label><input type="text" id="e-date"></div>
          <div class="field"><label>Customer Name</label><input type="text" id="e-customer"></div>
          <div class="field"><label>Company</label>
            <input type="text" id="e-company" list="company-list"></div>
          <div class="field"><label>Cluster</label>
            <input type="text" id="e-cluster" list="cluster-list"></div>
          <div class="field"><label>Branch</label>
            <input type="text" id="e-branch" list="branch-list"></div>
          <div class="field"><label>Cheque No.</label><input type="text" id="e-chq"></div>
          <div class="field"><label>Amount (₹)</label>
            <input type="number" id="e-amount"><span class="hint">Changing this recomputes charges, GST, total &amp; status</span></div>
          <div class="field"><label>Serviced Branch</label>
            <input type="text" id="e-srv-branch" list="srv-branch-list"></div>
          <div class="field"><label>Serviced Cluster</label>
            <input type="text" id="e-srv-cluster" list="srv-cluster-list"></div>
          <div class="field" style="grid-column:1/-1"><label>Disbursement UTR / Debit Note</label>
            <input type="text" id="e-debit-note"></div>
          <div class="field" style="grid-column:1/-1"><label>Remarks</label>
            <input type="text" id="e-remarks"></div>
        </div>
        <div id="e-computed" style="margin-top:10px;padding:10px 12px;background:#f0f7ff;border:1px solid #b3d4f5;border-radius:6px;font-size:13px;line-height:1.7"></div>
        <button class="btn btn-save" style="margin-top:12px" onclick="saveEditDisb()">💾 Save Disbursement Changes</button>
        <div id="e-disb-status" class="status"></div>
      </div>
      <div class="section">
        <h3>Repayments</h3>
        <div style="overflow-x:auto">
          <table class="recon-table" style="width:100%">
            <thead><tr><th>Date</th><th style="text-align:right">Amount (₹)</th><th>UTR / Note</th><th style="width:130px">Actions</th></tr></thead>
            <tbody id="e-pay-tbody"></tbody>
          </table>
        </div>
        <div style="display:flex;gap:10px;align-items:flex-end;margin-top:14px;padding-top:14px;border-top:1px solid #e0e6ed">
          <div class="field" style="flex:0 0 130px;margin:0"><label>Date</label>
            <input type="text" id="e-new-pay-date" placeholder="DD-MM-YYYY"></div>
          <div class="field" style="flex:0 0 140px;margin:0"><label>Amount (₹)</label>
            <input type="number" id="e-new-pay-amount"></div>
          <div class="field" style="flex:1;margin:0"><label>UTR / Note</label>
            <input type="text" id="e-new-pay-utr"></div>
          <button class="btn btn-save" style="width:auto;padding:10px 20px" onclick="addPayRow()">➕ Add Payment</button>
        </div>
        <div id="e-pay-status" class="status"></div>
      </div>
    </div>
  </div>

  <!-- INVOICE TAB -->
  <div id="invoice" class="tab-content">
    <div class="section">
      <h3>Generate Invoice &amp; Ledger</h3>
      <div class="field" style="margin-bottom:14px">
        <label>Select Open Case</label>
        <select id="inv-open-cases" onchange="onInvCaseSelect()" style="width:100%;font-size:.85rem">
          <option value="">Loading open cases...</option>
        </select>
      </div>
      <div class="field" style="margin-bottom:14px">
        <label>Or enter Disbursement ID manually <span style="font-weight:400;color:#888">(for any case incl. closed)</span></label>
        <input type="text" id="inv-manual-id" placeholder="e.g. BLP-010726-001" style="text-transform:uppercase" oninput="this.value=this.value.toUpperCase()">
      </div>
      <div id="inv-case-info" style="display:none;padding:10px 12px;background:#f0f7ff;border:1px solid #b3d4f5;border-radius:6px;font-size:13px;margin-bottom:14px;line-height:1.6;"></div>
      <button class="btn btn-save" onclick="generateInvoice()" id="inv-btn" style="width:auto;padding:10px 28px">📄 Generate &amp; Download PDF</button>
      <div id="inv-status" class="status" style="margin-top:10px"></div>
    </div>
  </div>

</div>

<script>
let openCasesData = [];
let reconTxns     = [];
let _pendingReqsOpen = true;

async function loadPendingRequests() {
  const list = document.getElementById('pending-requests-list');
  const countEl = document.getElementById('pending-count');
  list.innerHTML = '<span style="color:#888;font-size:13px;">Loading...</span>';
  try {
    const items = await (await fetch('/requests/pending?include=exported')).json();
    if (!Array.isArray(items) || items.length === 0) {
      list.innerHTML = '<span style="color:#aaa;font-size:13px;">No pending requests</span>';
      countEl.textContent = '';
      return;
    }
    countEl.textContent = `(${items.length})`;
    list.innerHTML = '';
    items.forEach(r => {
      const card = document.createElement('div');
      card.style.cssText = 'border:1px solid #dce8f5;border-radius:6px;padding:10px 12px;margin-bottom:8px;cursor:pointer;background:#fff;transition:background 0.15s;';
      card.onmouseenter = () => card.style.background = '#f0f7ff';
      card.onmouseleave = () => card.style.background = '#fff';
      const amt = Number(r.amount) > 0 ? '₹' + Number(r.amount).toLocaleString('en-IN', {maximumFractionDigits:0}) : r.amount;
      const badge = r.status === 'Exported'
        ? ' <span style="font-size:11px;background:#fff3cd;color:#856404;border:1px solid #ffeeba;border-radius:4px;padding:1px 6px;font-weight:400;">Exported</span>'
        : '';
      card.innerHTML = `<button class="req-del-btn" title="Delete request" style="float:right;font-size:12px;padding:2px 8px;cursor:pointer;border:1px solid #e0b4b4;border-radius:4px;background:#fff;color:#c0392b;">🗑</button>
        <div style="font-weight:600;font-size:14px;">${r.customer}${badge}</div>
        <div style="font-size:12px;color:#555;margin-top:3px;">${r.cluster} · ${r.branch} · ${amt}</div>
        <div style="font-size:11px;color:#999;margin-top:2px;">${r.request_id} · ${r.submitted_at}</div>`;
      card.addEventListener('click', () => selectPendingRequest(r));
      card.querySelector('.req-del-btn').addEventListener('click', (ev) => {
        ev.stopPropagation();
        deletePendingRequest(r);
      });
      list.appendChild(card);
    });
  } catch(e) {
    list.innerHTML = `<span style="color:#c00;font-size:13px;">Error loading requests: ${e.message}</span>`;
  }
}

async function deletePendingRequest(r) {
  const reason = prompt(`Delete ${r.request_id} (${r.customer}, ${r.cluster}/${r.branch}) from the queue?\n\nReason (optional — saved to the sheet):`);
  if (reason === null) return;
  if (!confirm(`Confirm: remove ${r.request_id} from the queue? The row stays in the sheet marked Deleted.`)) return;
  try {
    const res = await (await fetch('/requests/delete', {method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({request_ids: [r.request_id], reason: reason.trim()})})).json();
    if (!res.ok) throw new Error(res.error || 'Failed');
    // If the deleted request was selected into the form, clear the link so
    // saving the disbursement doesn't try to mark a Deleted row Disbursed.
    if (document.getElementById('d-request-id').value === r.request_id) {
      document.getElementById('d-request-id').value = '';
      document.getElementById('pending-selected-info').style.display = 'none';
    }
    loadPendingRequests();
  } catch(e) {
    alert('Delete failed: ' + e.message);
  }
}

function selectPendingRequest(r) {
  document.getElementById('d-customer').value = r.customer || '';
  document.getElementById('d-cluster').value  = r.cluster  || '';
  document.getElementById('d-branch').value   = r.branch   || '';
  if (r.amount) { document.getElementById('d-amount').value = r.amount; calcCharges(); }
  document.getElementById('d-request-id').value = r.request_id;

  const info = document.getElementById('pending-selected-info');
  const amt = Number(r.amount) > 0 ? '₹' + Number(r.amount).toLocaleString('en-IN', {maximumFractionDigits:0}) : r.amount;
  info.innerHTML = `<b>Selected:</b> ${r.customer} &nbsp;·&nbsp; ${r.cluster} / ${r.branch} &nbsp;·&nbsp; ${amt}<br>
    <b>Account:</b> ${r.account_no || '—'} &nbsp; <b>IFSC:</b> ${r.ifsc || '—'}<br>
    <b>Phone:</b> ${r.phone || '—'} &nbsp; <b>SO:</b> ${r.so_name || '—'} &nbsp; <b>Gold:</b> ${r.gold_weight || '—'} gms<br>
    <span style="color:#2a7ae2;font-size:11px;">${r.request_id}</span>
    <button onclick="clearDisb()" style="float:right;font-size:11px;padding:2px 8px;cursor:pointer;border:1px solid #ccc;border-radius:4px;background:#fff;">✕ Clear</button>`;
  info.style.display = 'block';
  document.getElementById('d-customer').scrollIntoView({ behavior: 'smooth', block: 'center' });
}

// ── Edit Case tab ────────────────────────────────────────────────────────────
let _editCase = null;
let _recentCasesOpen = true;

async function loadRecentCases() {
  const list = document.getElementById('e-recent-list');
  const countEl = document.getElementById('e-recent-count');
  list.innerHTML = '<span style="color:#888;font-size:13px;">Loading...</span>';
  try {
    const r = await _fetchJsonWithTimeout('/case/recent?days=7', {});
    if (!r.ok) throw new Error(r.error || 'Failed to load');
    const events = r.events;
    if (!events.length) {
      list.innerHTML = '<span style="color:#aaa;font-size:13px;">No disbursements or repayments recorded in the last 7 days</span>';
      countEl.textContent = '';
      return;
    }
    countEl.textContent = `(${events.length})`;
    list.innerHTML = '';
    events.forEach(e => {
      const card = document.createElement('div');
      card.style.cssText = 'border:1px solid #dce8f5;border-radius:6px;padding:10px 12px;margin-bottom:8px;cursor:pointer;background:#fff;transition:background 0.15s;';
      card.onmouseenter = () => card.style.background = '#f0f7ff';
      card.onmouseleave = () => card.style.background = '#fff';
      const isDisb = e.type === 'Disbursement';
      const typeTag = isDisb
        ? '<span style="background:#e8f0fe;color:#1a3a5c;font-size:10px;padding:1px 6px;border-radius:3px;">➕ DISBURSEMENT</span>'
        : '<span style="background:#e8f4ec;color:#155724;font-size:10px;padding:1px 6px;border-radius:3px;">💰 REPAYMENT</span>';
      const archTag = e.sheet && e.sheet !== 'Accounts'
        ? ` <span style="background:#fff3cd;color:#8a6500;font-size:10px;padding:1px 6px;border-radius:3px;">📁 ${e.sheet}</span>` : '';
      card.innerHTML = `<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px">
          <div style="font-weight:600;font-size:14px;">${e.customer} <span style="font-weight:400;color:#888;font-size:12px">${e.disb_id}</span></div>
          <div style="font-weight:700;font-size:14px;color:${isDisb ? '#1a3a5c' : '#155724'}">₹${fmt(e.amount)}</div>
        </div>
        <div style="margin-top:3px">${typeTag}${archTag}</div>
        <div style="font-size:12px;color:#555;margin-top:5px;">${e.cluster} · ${e.branch} &nbsp;·&nbsp; ${e.status || '—'}${e.utr ? ' &nbsp;·&nbsp; ' + e.utr : ''}</div>
        <div style="font-size:11px;color:#999;margin-top:2px;">${e.date}</div>`;
      card.addEventListener('click', () => {
        document.getElementById('e-disb-id').value = e.disb_id;
        loadEditCase();
      });
      list.appendChild(card);
    });
  } catch(e) {
    list.innerHTML = `<span style="color:#c00;font-size:13px;">Error: ${e.message}</span>`;
  }
}

function toggleRecentCases() {
  _recentCasesOpen = !_recentCasesOpen;
  document.getElementById('e-recent-body').style.display = _recentCasesOpen ? '' : 'none';
  document.getElementById('e-recent-chevron').textContent = _recentCasesOpen ? '▼' : '▶';
}

async function loadEditCase() {
  const id = document.getElementById('e-disb-id').value.trim().toUpperCase();
  if (!id) return showStatus('e-load-status','error','Enter a Disbursement ID.');
  showStatus('e-load-status','success','Loading...');
  try {
    const r = await _fetchJsonWithTimeout(`/case/${encodeURIComponent(id)}/detail`, {});
    if (!r.ok) throw new Error(r.error || 'Not found');
    renderEditCase(r.case);
    showStatus('e-load-status','success',`✅ Loaded ${id}`);
  } catch(e) {
    document.getElementById('e-body').style.display = 'none';
    showStatus('e-load-status','error','❌ ' + e.message);
  }
}

function renderEditCase(c) {
  _editCase = c;
  document.getElementById('e-body').style.display = '';
  document.getElementById('e-case-id').textContent = `— ${c.disb_id} (${c.sheet})`;
  document.getElementById('e-body').scrollIntoView({ behavior: 'smooth', block: 'start' });
  const set = (id,v) => document.getElementById(id).value = v ?? '';
  set('e-date', c.date); set('e-customer', c.customer); set('e-chq', c.chq);
  set('e-company', c.company); set('e-cluster', c.cluster); set('e-branch', c.branch);
  set('e-amount', c.amount); set('e-srv-branch', c.srv_branch); set('e-srv-cluster', c.srv_cluster);
  set('e-debit-note', c.debit_note); set('e-remarks', c.remarks);
  document.getElementById('e-new-pay-date').value = new Date().toLocaleDateString('en-GB').replace(/\//g,'-');
  document.getElementById('e-computed').innerHTML =
    `<b>Computed:</b> Charges ₹${fmtDec(c.charges)} &nbsp;|&nbsp; GST ₹${fmtDec(c.gst)} &nbsp;|&nbsp; Total ₹${fmtDec(c.total)}` +
    ` &nbsp;|&nbsp; Collected ₹${fmtDec(c.collected)} &nbsp;|&nbsp; Discount ₹${fmtDec(c.discount)}` +
    ` &nbsp;|&nbsp; <b>Balance ₹${fmtDec(c.balance)}</b> &nbsp;|&nbsp; Status: <b>${c.status || '—'}</b>`;
  renderEditPayments();
}

function renderEditPayments() {
  const tb = document.getElementById('e-pay-tbody');
  tb.innerHTML = '';
  if (!_editCase.payments.length) {
    tb.innerHTML = '<tr><td colspan="4" style="color:#aaa">No payments recorded in M Coll for this case.</td></tr>';
    return;
  }
  _editCase.payments.forEach((p, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${p.date}</td><td style="text-align:right">${fmtDec(p.amount)}</td>` +
      `<td style="max-width:340px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${(p.utr||'').replace(/"/g,'&quot;')}">${p.utr || '—'}</td>` +
      `<td><button onclick="editPayRow(${i})" style="padding:3px 10px;font-size:12px;cursor:pointer">✏️ Edit</button> ` +
      `<button onclick="delPayRow(${i})" style="padding:3px 10px;font-size:12px;cursor:pointer;color:#a00">🗑</button></td>`;
    tb.appendChild(tr);
  });
}

function editPayRow(i) {
  const p = _editCase.payments[i];
  const tr = document.getElementById('e-pay-tbody').children[i];
  tr.innerHTML = `<td><input type="text" id="ep-date-${i}" value="${p.date}" style="width:110px"></td>` +
    `<td style="text-align:right"><input type="number" id="ep-amt-${i}" value="${p.amount}" style="width:120px;text-align:right"></td>` +
    `<td><input type="text" id="ep-utr-${i}" value="${(p.utr||'').replace(/"/g,'&quot;')}" style="width:100%"></td>` +
    `<td><button onclick="savePayRow(${i})" style="padding:3px 10px;font-size:12px;cursor:pointer;color:#155724">✅ Save</button> ` +
    `<button onclick="renderEditPayments()" style="padding:3px 10px;font-size:12px;cursor:pointer">✕</button></td>`;
}

// 55s client-side ceiling so a slow/hung request (e.g. Sheets API retrying
// a rate-limit under the hood) always resolves to a visible error instead
// of leaving the button/status stuck on "Saving..." forever. Set above the
// server's own 60s function timeout isn't possible client-side, but 55s
// gives real Sheets calls (which can legitimately take 20-40s under
// quota pressure) room to finish rather than aborting a write that would
// have succeeded a few seconds later.
async function _fetchJsonWithTimeout(url, opts, ms) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), ms || 55000);
  try {
    const res = await fetch(url, {...opts, signal: ctrl.signal});
    return await res.json();
  } catch (e) {
    if (e.name === 'AbortError') throw new Error('Request timed out — the sheet may be busy, please try again.');
    throw e;
  } finally {
    clearTimeout(timer);
  }
}

async function savePayRow(i) {
  const p = _editCase.payments[i];
  const body = {
    disb_id: _editCase.disb_id, mc_row: p.mc_row,
    date:   document.getElementById(`ep-date-${i}`).value.trim(),
    amount: document.getElementById(`ep-amt-${i}`).value,
    utr:    document.getElementById(`ep-utr-${i}`).value.trim(),
  };
  if (!body.amount) return showStatus('e-pay-status','error','Amount is required.');
  showStatus('e-pay-status','success','Saving...');
  try {
    const r = await _fetchJsonWithTimeout('/case/update-repayment', {method:'POST',
      headers:{'Content-Type':'application/json'}, body: JSON.stringify(body)});
    if (r.ok) { renderEditCase(r.case); loadRecentCases(); showStatus('e-pay-status','success','✅ Payment updated — balance & status recomputed.'); }
    else showStatus('e-pay-status','error','❌ ' + r.error);
  } catch (e) {
    showStatus('e-pay-status','error','❌ ' + e.message);
  }
}

async function delPayRow(i) {
  const p = _editCase.payments[i];
  if (!confirm(`Delete this payment?\n\n${p.date}  ₹${fmtDec(p.amount)}\n${p.utr || ''}\n\nThe case balance and status will be recomputed.`)) return;
  showStatus('e-pay-status','success','Deleting...');
  try {
    const r = await _fetchJsonWithTimeout('/case/delete-repayment', {method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({disb_id: _editCase.disb_id, mc_row: p.mc_row})});
    if (r.ok) { renderEditCase(r.case); loadRecentCases(); showStatus('e-pay-status','success','✅ Payment deleted — balance & status recomputed.'); }
    else showStatus('e-pay-status','error','❌ ' + r.error);
  } catch (e) {
    showStatus('e-pay-status','error','❌ ' + e.message);
  }
}

async function addPayRow() {
  if (!_editCase) return;
  const body = {
    disb_id: _editCase.disb_id,
    date:   document.getElementById('e-new-pay-date').value.trim(),
    amount: document.getElementById('e-new-pay-amount').value,
    utr:    document.getElementById('e-new-pay-utr').value.trim(),
  };
  if (!body.date || !body.amount) return showStatus('e-pay-status','error','Date and Amount are required.');
  showStatus('e-pay-status','success','Adding...');
  try {
    const r = await _fetchJsonWithTimeout('/case/add-repayment', {method:'POST',
      headers:{'Content-Type':'application/json'}, body: JSON.stringify(body)});
    if (r.ok) {
      renderEditCase(r.case); loadRecentCases();
      document.getElementById('e-new-pay-amount').value = '';
      document.getElementById('e-new-pay-utr').value = '';
      showStatus('e-pay-status','success','✅ Payment added — balance & status recomputed.');
    } else showStatus('e-pay-status','error','❌ ' + r.error);
  } catch (e) {
    showStatus('e-pay-status','error','❌ ' + e.message);
  }
}

async function saveEditDisb() {
  if (!_editCase) return;
  const g = id => document.getElementById(id).value.trim();
  const body = {
    disb_id: _editCase.disb_id,
    date: g('e-date'), customer: g('e-customer'), chq: g('e-chq'),
    company: g('e-company'), cluster: g('e-cluster'), branch: g('e-branch'),
    amount: g('e-amount'), srv_branch: g('e-srv-branch'), srv_cluster: g('e-srv-cluster'),
    debit_note: g('e-debit-note'), remarks: g('e-remarks'),
  };
  if (!body.customer || !body.amount) return showStatus('e-disb-status','error','Customer and Amount are required.');
  const btn = event.target; btn.disabled = true; btn.textContent = 'Saving...';
  try {
    const r = await _fetchJsonWithTimeout('/case/update-disbursement', {method:'POST',
      headers:{'Content-Type':'application/json'}, body: JSON.stringify(body)});
    if (r.ok) { renderEditCase(r.case); loadRecentCases(); showStatus('e-disb-status','success','✅ Saved — derived values recomputed.'); }
    else showStatus('e-disb-status','error','❌ ' + r.error);
  } catch (e) {
    showStatus('e-disb-status','error','❌ ' + e.message);
  } finally {
    btn.disabled = false; btn.textContent = '💾 Save Disbursement Changes';
  }
}

// ── Invoice tab ──────────────────────────────────────────────────────────────
async function loadInvoiceCases() {
  const sel = document.getElementById('inv-open-cases');
  sel.innerHTML = '<option value="">Loading...</option>';
  try {
    const cases = await (await fetch('/open-cases')).json();
    sel.innerHTML = '<option value="">— Select an open case —</option>';
    cases.forEach(c => {
      const opt = document.createElement('option');
      opt.value = c.disb_id;
      opt.textContent = `${c.disb_id}  |  ${c.customer}  |  ₹${fmt(c.balance)} due`;
      sel.appendChild(opt);
    });
  } catch(e) {
    sel.innerHTML = '<option value="">Error loading cases</option>';
  }
}

function onInvCaseSelect() {
  const id = document.getElementById('inv-open-cases').value;
  document.getElementById('inv-manual-id').value = '';
  const info = document.getElementById('inv-case-info');
  if (!id) { info.style.display = 'none'; return; }
  const c = openCasesData.find(x => x.disb_id === id);
  if (c) {
    info.innerHTML = `<b>${c.customer}</b> &nbsp;|&nbsp; Disbursed: ₹${fmt(c.amount)} &nbsp;|&nbsp; Total Due: ₹${fmt(c.total)} &nbsp;|&nbsp; <b>Balance: ₹${fmt(c.balance)}</b> &nbsp;|&nbsp; ${c.status}`;
    info.style.display = 'block';
  }
}

async function generateInvoice() {
  const fromDrop = document.getElementById('inv-open-cases').value;
  const fromText = document.getElementById('inv-manual-id').value.trim().toUpperCase();
  const disb_id = fromText || fromDrop;
  if (!disb_id) return showStatus('inv-status', 'error', 'Select a case or enter a Disbursement ID.');
  const btn = document.getElementById('inv-btn');
  btn.disabled = true; btn.textContent = '⏳ Generating...';
  showStatus('inv-status', 'success', 'Fetching data and building PDF...');
  try {
    const r = await fetch('/generate-invoice', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({disb_id}),
    });
    if (!r.ok) {
      const err = await r.json().catch(() => ({error: `HTTP ${r.status}`}));
      throw new Error(err.error || `HTTP ${r.status}`);
    }
    const blob = await r.blob();
    const disposition = r.headers.get('Content-Disposition') || '';
    const match = disposition.match(/filename="(.+)"/);
    const filename = match ? match[1] : `${disb_id} Invoice and Ledger.pdf`;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    showStatus('inv-status', 'success', `✅ Downloaded: ${filename}`);
  } catch(e) {
    showStatus('inv-status', 'error', '❌ ' + e.message);
  } finally {
    btn.disabled = false; btn.textContent = '📄 Generate & Download PDF';
  }
}

function togglePendingReqs() {
  _pendingReqsOpen = !_pendingReqsOpen;
  document.getElementById('pending-requests-body').style.display = _pendingReqsOpen ? '' : 'none';
  document.getElementById('pending-chevron').textContent = _pendingReqsOpen ? '▼' : '▶';
}

// ── Date picker helper ────────────────────────────────────────────────────
function _openDatePicker(nativeId) {
  const inp = document.getElementById(nativeId);
  const textId = nativeId.replace('-native', '');
  const txtVal = (document.getElementById(textId).value || '').trim();
  const m = txtVal.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (m) inp.value = `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  if (inp.showPicker) { try { inp.showPicker(); return; } catch(e) {} }
  inp.click();
}
function _pickDate(textId, isoVal) {
  if (!isoVal) return;
  const [y,m,d] = isoVal.split('-');
  document.getElementById(textId).value = `${d}-${m}-${y}`;
}

function showTab(id, btn) {
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  btn.classList.add('active');
}

function calcCharges() {
  const amt = parseFloat(document.getElementById('d-amount').value) || 0;
  const ch  = Math.round(amt * 0.005 * 100) / 100;
  const gst = Math.round(ch * 0.18 * 100) / 100;
  const tot = amt + ch + gst;
  document.getElementById('d-charges').value = ch.toLocaleString('en-IN', {minimumFractionDigits:2});
  document.getElementById('d-gst').value     = gst.toLocaleString('en-IN', {minimumFractionDigits:2});
  document.getElementById('d-total').value   = tot.toLocaleString('en-IN', {minimumFractionDigits:2});
}

async function extractDisb() {
  const msg = document.getElementById('d-msg').value.trim();
  if (!msg) return;
  const d = await (await fetch('/extract/disbursement', {method:'POST',
    headers:{'Content-Type':'application/json'}, body: JSON.stringify({message:msg})})).json();
  if (d.date)     document.getElementById('d-date').value = d.date;
  if (d.customer) document.getElementById('d-customer').value = d.customer;
  if (d.company)  document.getElementById('d-company').value = d.company;
  if (d.cluster)  document.getElementById('d-cluster').value = d.cluster;
  if (d.branch)   document.getElementById('d-branch').value = d.branch;
  if (d.amount)  { document.getElementById('d-amount').value = d.amount; calcCharges(); }
}

async function extractDisbUTR() {
  const msg = document.getElementById('d-utr-msg').value.trim();
  if (!msg) return;
  const d = await (await fetch('/extract/repayment', {method:'POST',
    headers:{'Content-Type':'application/json'}, body: JSON.stringify({message:msg})})).json();
  if (d.utr) {
    document.getElementById('d-utr').value = d.utr;
    if (d.date && !document.getElementById('d-date').value)
      document.getElementById('d-date').value = d.date;
  } else {
    document.getElementById('d-utr').value = '';
    alert('No UTR found in that message.');
  }
}

async function extractRepa() {
  const msg = document.getElementById('r-msg').value.trim();
  if (!msg) return;
  const d = await (await fetch('/extract/repayment', {method:'POST',
    headers:{'Content-Type':'application/json'}, body: JSON.stringify({message:msg})})).json();
  if (d.date)   document.getElementById('r-date').value = d.date;
  if (d.amount) document.getElementById('r-amount').value = d.amount;
  if (d.utr)    document.getElementById('r-utr').value = d.utr;
  if (d.disb_id) {
    const sel = document.getElementById('r-open-cases');
    for (let i = 0; i < sel.options.length; i++) {
      if (sel.options[i].value === d.disb_id) { sel.value = d.disb_id; break; }
    }
    onCaseSelect();
  }
  if (d.sender) showInfo(`Sender: <b>${d.sender}</b> — select the matching case above.`, false);
}

async function loadOpenCases() {
  const sel = document.getElementById('r-open-cases');
  const spinner = document.getElementById('cases-spinner');
  sel.innerHTML = '<option value="">Loading...</option>';
  spinner.style.display = 'block';
  try {
    const cases = await (await fetch('/open-cases')).json();
    openCasesData = cases;
    sel.innerHTML = '<option value="">— Select an open case —</option>';
    cases.forEach(c => {
      const opt = document.createElement('option');
      opt.value = c.disb_id;
      opt.textContent = `${c.disb_id}  |  ${c.customer}  |  ₹${fmt(c.balance)} due`;
      sel.appendChild(opt);
    });
  } catch(e) {
    sel.innerHTML = '<option value="">Error loading — check credentials</option>';
  }
  spinner.style.display = 'none';
}

function onCaseSelect() {
  const id = document.getElementById('r-open-cases').value;
  const c  = openCasesData.find(x => x.disb_id === id);
  if (!c) { document.getElementById('r-info').style.display='none'; return; }
  showInfo(`<b>${c.customer}</b> &nbsp;|&nbsp; Disbursed: ₹${fmt(c.amount)} &nbsp;|&nbsp; Total Due: ₹${fmt(c.total)} &nbsp;|&nbsp; Collected: ₹${fmt(c.collected)} &nbsp;|&nbsp; <b>Balance: ₹${fmt(c.balance)}</b> &nbsp;|&nbsp; ${c.status}`, false);
}

function showInfo(msg, isError) {
  const box = document.getElementById('r-info');
  box.style.display = 'block';
  box.className = 'info-box' + (isError ? ' error' : '');
  box.innerHTML = msg;
}

function fmt(n) { return Number(n).toLocaleString('en-IN', {maximumFractionDigits:0}); }
function fmtDec(n) { return Number(n).toLocaleString('en-IN', {minimumFractionDigits:2, maximumFractionDigits:2}); }

async function generateMis() {
  const btn = document.getElementById('mis-btn');
  btn.disabled = true; btn.textContent = '⏳ Generating...';
  try {
    const r = await fetch('/generate-mis', {method: 'POST'});
    if (!r.ok) {
      const err = await r.json().catch(() => ({error: 'Unknown error'}));
      throw new Error(err.error || `HTTP ${r.status}`);
    }
    const blob = await r.blob();
    const disposition = r.headers.get('Content-Disposition') || '';
    const match = disposition.match(/filename="(.+)"/);
    const filename = match ? match[1] : 'BridgeLine MIS Package.zip';
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    showStatus('mis-status', 'success', `✅ Downloaded ${filename}`);
  } catch (e) {
    showStatus('mis-status', 'error', '❌ ' + e.message);
  } finally {
    btn.disabled = false; btn.textContent = '📦 Generate MIS Package';
  }
}

async function saveDisb() {
  const data = {
    date: document.getElementById('d-date').value.trim(),
    customer: document.getElementById('d-customer').value.trim(),
    chq: document.getElementById('d-chq').value.trim(),
    company: document.getElementById('d-company').value.trim(),
    cluster: document.getElementById('d-cluster').value.trim(),
    branch: document.getElementById('d-branch').value.trim(),
    amount: document.getElementById('d-amount').value,
    serviced_branch: document.getElementById('d-srv-branch').value.trim(),
    serviced_cluster: document.getElementById('d-srv-cluster').value.trim(),
    utr:        document.getElementById('d-utr').value.trim(),
    remarks:    document.getElementById('d-remarks').value.trim(),
    request_id: document.getElementById('d-request-id').value.trim(),
  };
  if (!data.customer || !data.amount || !data.company || !data.cluster || !data.branch)
    return showStatus('d-status','error','Please fill all required (*) fields.');
  const btn = event.target; btn.disabled = true; btn.textContent = 'Saving...';
  const r = await (await fetch('/save/disbursement', {method:'POST',
    headers:{'Content-Type':'application/json'}, body: JSON.stringify(data)})).json();
  btn.disabled = false; btn.textContent = '✅ Save Disbursement to Google Sheet';
  if (r.ok) { showStatus('d-status','success',`✅ Saved! ID: ${r.disb_id}`); clearDisb(); loadPendingRequests(); }
  else       showStatus('d-status','error','❌ ' + r.error);
}

async function saveRepa() {
  const disb_id = document.getElementById('r-open-cases').value;
  if (!disb_id) return showStatus('r-status','error','Please select an open case.');
  const amount = document.getElementById('r-amount').value;
  if (!amount)  return showStatus('r-status','error','Amount is required.');
  const data = {
    disb_id,
    date:     document.getElementById('r-date').value.trim(),
    amount,
    utr:      document.getElementById('r-utr').value.trim(),
    discount: document.getElementById('r-discount').value || 0,
    raw_msg:  document.getElementById('r-msg').value.trim(),
    remarks:  document.getElementById('r-remarks').value.trim(),
  };
  const btn = event.target; btn.disabled = true; btn.textContent = 'Saving...';
  const r = await (await fetch('/save/repayment', {method:'POST',
    headers:{'Content-Type':'application/json'}, body: JSON.stringify(data)})).json();
  btn.disabled = false; btn.textContent = '✅ Save Repayment to Google Sheet';
  if (r.ok) {
    if (r.mcoll_warning) {
      showStatus('r-status','error',`⚠️ ₹${fmt(amount)} recorded (balance ₹${fmt(r.new_balance)}) but NOT logged to payment history — ${r.mcoll_warning}`);
    } else {
      showStatus('r-status','success',`✅ ₹${fmt(amount)} recorded. New balance: ₹${fmt(r.new_balance)}. Status: ${r.status}`);
    }
    document.getElementById('r-amount').value = '';
    document.getElementById('r-utr').value = '';
    document.getElementById('r-discount').value = '0';
    document.getElementById('r-msg').value = '';
    document.getElementById('r-remarks').value = '';
    loadOpenCases();
  } else showStatus('r-status','error','❌ ' + r.error);
}

function showStatus(id, type, msg) {
  const el = document.getElementById(id);
  el.className = 'status ' + type;
  el.textContent = msg;
  el.style.display = 'block';
  setTimeout(() => el.style.display = 'none', 9000);
}

function clearDisb() {
  ['d-date','d-customer','d-chq','d-amount','d-charges','d-gst','d-total','d-utr','d-remarks'].forEach(id =>
    document.getElementById(id).value = '');
  ['d-company','d-cluster','d-branch'].forEach(id =>
    document.getElementById(id).value = '');
  ['d-srv-branch','d-srv-cluster'].forEach(id =>
    document.getElementById(id).value = '');
  document.getElementById('d-msg').value = '';
  document.getElementById('d-utr-msg').value = '';
  document.getElementById('d-date').value = new Date().toLocaleDateString('en-GB').replace(/\\//g,'-');
  document.getElementById('d-request-id').value = '';
  document.getElementById('pending-selected-info').style.display = 'none';
}

// ── Summary banner ──────────────────────────────────────────────────────────
let _summaryData = {};
async function loadSummary() {
  try {
    const d = await (await fetch('/summary')).json();
    _summaryData = d;
    document.getElementById('s-disb').textContent = (d.disbursed_today||0) > 0 ? '₹'+fmt(d.disbursed_today) : '—';
    document.getElementById('s-coll').textContent = (d.collected_today||0) > 0 ? '₹'+fmt(d.collected_today) : '—';
    document.getElementById('s-out').textContent  = d.total_outstanding != null ? '₹'+fmt(d.total_outstanding) : '—';
    document.getElementById('s-avail').textContent = d.available_for_disbursement != null ? '₹'+fmt(d.available_for_disbursement) : '—';
    document.getElementById('s-date').textContent = d.date;
  } catch(e) {}
}

function openSumModal(kind) {
  const d = _summaryData;
  const cfgs = {
    disb: {title: 'Disbursed Today', rows: d.disb_rows || [], emptyMsg: 'No disbursements recorded today.'},
    coll: {title: 'Collected Today', rows: d.coll_rows || [], emptyMsg: 'No collections recorded today.'},
    out:  {title: 'Total Outstanding', rows: d.out_rows || [], emptyMsg: 'No open balances.'},
    avail:{title: 'Available to Disburse', rows: null, emptyMsg: ''},
  };
  const cfg = cfgs[kind];
  document.getElementById('sum-modal-title').textContent = cfg.title;
  const body = document.getElementById('sum-modal-body');
  if (kind === 'avail') {
    if (d.available_for_disbursement == null) {
      body.innerHTML = `<div class="empty">No bank reconciliation found yet — upload a bank statement under "Bank Reconciliation" to enable this.</div>`;
    } else {
      body.innerHTML = `<div style="padding:18px">
        <table style="width:100%;font-size:.88rem">
          <tr><td style="padding:6px 0;color:#555">Bank Closing Balance (${d.bank_balance_date||''})</td><td style="padding:6px 0;text-align:right;font-weight:600">₹${fmt(d.bank_balance||0)}</td></tr>
          <tr><td style="padding:6px 0;color:#555">Plus: Collected since then</td><td style="padding:6px 0;text-align:right;font-weight:600;color:#1a5c3a">+₹${fmt(d.collected_since_bank||0)}</td></tr>
          <tr><td style="padding:6px 0;color:#555">Less: Disbursed since then</td><td style="padding:6px 0;text-align:right;font-weight:600;color:#c00">−₹${fmt(d.disbursed_since_bank||0)}</td></tr>
          <tr style="border-top:2px solid #1a3a5c"><td style="padding:8px 0;font-weight:700">Available to Disburse</td><td style="padding:8px 0;text-align:right;font-weight:700;color:#1a5c3a">₹${fmt(d.available_for_disbursement||0)}</td></tr>
        </table></div>`;
    }
  } else if (!cfg.rows.length) {
    body.innerHTML = `<div class="empty">${cfg.emptyMsg}</div>`;
  } else {
    let total = cfg.rows.reduce((s,r) => s + (r.amount||0), 0);
    body.innerHTML = `<table>
      <thead><tr><th>Disb ID</th><th>Customer</th><th style="text-align:right">Amount</th></tr></thead>
      <tbody>${cfg.rows.map(r => `<tr><td>${r.disb_id||''}</td><td>${r.customer||''}</td><td style="text-align:right">₹${fmt(r.amount||0)}</td></tr>`).join('')}</tbody>
      <tfoot><tr style="font-weight:700;background:#f0f4f8"><td colspan="2" style="padding:8px 14px">Total</td><td style="padding:8px 14px;text-align:right">₹${fmt(total)}</td></tr></tfoot>
    </table>`;
  }
  document.getElementById('sum-modal-bg').classList.add('show');
}
function closeSumModal() {
  document.getElementById('sum-modal-bg').classList.remove('show');
}

// ── Calculator ───────────────────────────────────────────────────────────────
let calcExpr = '';
let calcCurrent = '0';
let calcJustEq  = false;

function calcUpdate() {
  document.getElementById('calc-disp').textContent =
    parseFloat(calcCurrent).toLocaleString('en-IN', {maximumFractionDigits:8});
  document.getElementById('calc-expr').textContent = calcExpr;
}

function calcNum(ch) {
  if (calcJustEq) { calcCurrent = ''; calcExpr = ''; calcJustEq = false; }
  if (ch === '.' && calcCurrent.includes('.')) return;
  if (calcCurrent === '0' && ch !== '.') calcCurrent = ch;
  else calcCurrent += ch;
  calcUpdate();
}

function calcOp(op) {
  calcJustEq = false;
  calcExpr = calcCurrent + ' ' + op + ' ';
  calcCurrent = '0';
  calcUpdate();
}

function calcEq() {
  if (!calcExpr) return;
  try {
    const expr = calcExpr + calcCurrent;
    const result = Function('"use strict"; return (' + expr + ')')();
    calcExpr = expr + ' =';
    calcCurrent = String(parseFloat(result.toFixed(8)));
    calcJustEq = true;
    calcUpdate();
  } catch(e) {}
}

function calcClear() {
  calcExpr = ''; calcCurrent = '0'; calcJustEq = false; calcUpdate();
}

function calcBack() {
  if (calcJustEq) { calcClear(); return; }
  calcCurrent = calcCurrent.length > 1 ? calcCurrent.slice(0,-1) : '0';
  calcUpdate();
}

// Keyboard support for calculator
document.addEventListener('keydown', e => {
  const calcTab = document.getElementById('calc');
  if (!calcTab.classList.contains('active')) return;
  if (e.key >= '0' && e.key <= '9') calcNum(e.key);
  else if (e.key === '.') calcNum('.');
  else if (e.key === '+') calcOp('+');
  else if (e.key === '-') calcOp('-');
  else if (e.key === '*') calcOp('*');
  else if (e.key === '/') { e.preventDefault(); calcOp('/'); }
  else if (e.key === '%') calcOp('%');
  else if (e.key === 'Enter' || e.key === '=') calcEq();
  else if (e.key === 'Backspace') calcBack();
  else if (e.key === 'Escape') calcClear();
});

// ── Reconciliation ───────────────────────────────────────────────────────────
function onFileSelect(input) {
  const name = input.files[0] ? input.files[0].name : '';
  document.getElementById('upload-label').innerHTML =
    name ? `📄 <b>${name}</b> selected` : '📂 Click to upload bank statement';
}

async function parseStatement() {
  const fileInput = document.getElementById('rec-file');
  if (!fileInput.files.length) return alert('Please upload a bank statement file first.');
  const recDate = document.getElementById('rec-date').value.trim();
  if (!recDate) return alert('Please enter the reconciliation date.');

  const formData = new FormData();
  formData.append('file', fileInput.files[0]);

  const btn = event.target; btn.disabled = true; btn.textContent = 'Parsing...';
  try {
    const r = await (await fetch('/reconcile/parse', {method:'POST', body: formData})).json();
    btn.disabled = false; btn.textContent = '⚡ Parse Statement';
    if (!r.ok) return alert('❌ ' + r.error);

    reconTxns = r.transactions;
    if (r.opening_balance) document.getElementById('rec-opening').value = r.opening_balance;
    if (r.closing_balance) document.getElementById('rec-closing').value = r.closing_balance;
    renderReconResult(r);
  } catch(e) {
    btn.disabled = false; btn.textContent = '⚡ Parse Statement';
    alert('Error parsing file: ' + e);
  }
}

const BASE_TYPES = ['Disbursement','Collection','Collection (via Pradaan)','Expense','Capital In','Capital Out','Skip'];
const TYPE_COLOR = {
  'Disbursement':'#fff8e1','Collection':'#e8f5e9','Collection (via Pradaan)':'#e8f5e9',
  'Expense':'#fce4ec','Capital In':'#e3f2fd','Capital Out':'#e3f2fd','Skip':'#eeeeee'
};

let _customTypes = [];

async function loadCustomTypes() {
  try {
    const cfg = await (await fetch('/config')).json();
    _customTypes = cfg.custom_types || [];
  } catch { _customTypes = []; }
}

async function saveCustomType(t) {
  if (!t || BASE_TYPES.includes(t) || _customTypes.includes(t)) return;
  _customTypes.push(t);
  await fetch('/config', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({custom_types: _customTypes})});
}

let _bankAccounts = [];

async function loadBankAccounts() {
  try {
    const cfg = await (await fetch('/config')).json();
    _bankAccounts = cfg.bank_accounts || [];
  } catch { _bankAccounts = []; }
  const dl = document.getElementById('rec-account-list');
  if (dl) dl.innerHTML = _bankAccounts.map(a => `<option value="${a.name}">`).join('');
}

// If the typed account name isn't already registered, prompt for its
// account number (used for Capital In/Out transfer matching between our
// own accounts) and save it to Config so it shows up in the dropdown and
// transfer-matching going forward.
async function registerAccountIfNew(name) {
  if (_bankAccounts.some(a => a.name === name)) return;
  const acctNum = (prompt(`New bank account "${name}" — enter its account number (used to auto-detect transfers between your own accounts). Leave blank to skip.`) || '').trim();
  _bankAccounts.push({name, account_number: acctNum});
  await fetch('/config', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({bank_accounts: _bankAccounts})});
  await loadBankAccounts();
}

function allTypeOpts() {
  return ['', ...BASE_TYPES, ..._customTypes];
}
function buildTypeSelect(i, val='') {
  const opts = allTypeOpts().map(t =>
    `<option value="${t}" ${t===val?'selected':''}>${t||'— auto —'}</option>`).join('');
  return `<select data-row="${i}" onchange="onTypeChange(this)"
    style="width:100%;font-size:.82rem;padding:4px 6px;border:1px solid #b0c4d8;border-radius:5px;background:white">${opts}</select>`;
}
function rowBg(type) { return TYPE_COLOR[type] || 'white'; }

function onTypeChange(sel) {
  const i   = sel.dataset.row;
  const val = sel.value;
  document.getElementById('txrow-'+i).style.background = rowBg(val);
}

function renderReconResult(r) {
  const opening = r.opening_balance || 0;
  const closing = r.closing_balance || 0;
  const totalDr = r.transactions.reduce((s,t) => s+t.debit, 0);
  const totalCr = r.transactions.reduce((s,t) => s+t.credit, 0);

  document.getElementById("recon-summary").innerHTML = `
    <div class="recon-card"><div class="rv">\u20b9${fmtDec(opening)}</div><div class="rl">Opening</div></div>
    <div class="recon-card" style="background:#c00"><div class="rv">\u20b9${fmtDec(totalDr)}</div><div class="rl">Total Debits</div></div>
    <div class="recon-card" style="background:#1a5c3a"><div class="rv">\u20b9${fmtDec(totalCr)}</div><div class="rl">Total Credits</div></div>
    <div class="recon-card" style="background:#b8860b"><div class="rv">\u20b9${fmtDec(closing)}</div><div class="rl">Closing</div></div>`;

  const conf = r.confident || [];
  document.getElementById("confident-badge").textContent =
    `\u2014 ${conf.length} of ${r.transactions.length} entries matched automatically`;
  const byType = {};
  conf.forEach(tx => {
    byType[tx.type] = byType[tx.type] || {count:0, dr:0, cr:0};
    byType[tx.type].count++;
    byType[tx.type].dr += tx.debit;
    byType[tx.type].cr += tx.credit;
  });
  document.getElementById("confident-summary").innerHTML = Object.entries(byType).map(([t,v]) =>
    `<span style="display:inline-block;margin-right:24px"><b>${t}</b>: ${v.count} txns`
    + (v.dr ? ` &nbsp;Dr \u20b9${fmtDec(v.dr)}` : "")
    + (v.cr ? ` &nbsp;Cr \u20b9${fmtDec(v.cr)}` : "") + `</span>`
  ).join("");

  const rev = r.review || [];
  window._reviewTxns = rev;
  if (rev.length) {
    document.getElementById("review-section").style.display = "block";
    document.getElementById("review-badge").textContent = `\u2014 ${rev.length} entries need review`;
    document.getElementById("review-tbody").innerHTML = rev.map((tx, i) => {
      const desc = (tx.description||"").replace(/</g,"&lt;");
      const amt  = tx.debit ? "\u20b9"+fmtDec(tx.debit) : "\u20b9"+fmtDec(tx.credit);
      const drCr = tx.debit ? `<span style="color:#c00;font-weight:600">Dr</span>`
                            : `<span style="color:#1a5c3a;font-weight:600">Cr</span>`;
      const autoColor = ({Disbursement:"#b8860b",Collection:"#1a5c3a",Expense:"#c00",
        "FD Booking":"#1565c0","Capital In":"#1565c0","Capital Out":"#1565c0"})[tx.type] || "#555";
      return `<tr style="background:${i%2?"#f8fbff":"white"}">
        <td style="white-space:nowrap;font-size:.82rem">${tx.date}</td>
        <td style="font-size:.8rem;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${desc}">${desc}</td>
        <td style="font-size:.78rem;color:#666;max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${tx.utr||"\u2014"}</td>
        <td style="text-align:right;font-size:.82rem;font-weight:600">${amt}</td>
        <td style="text-align:center">${drCr}</td>
        <td style="font-size:.75rem;color:${autoColor};font-weight:600;white-space:nowrap">${tx.type||"\u2014"}</td>
        <td>${buildTypeSelect(i, tx.type)}</td>
        <td><input type="text" data-review-row="${i}" placeholder="remarks\u2026"
            style="width:100%;font-size:.82rem;padding:4px 6px;border:1px solid #b0c4d8;border-radius:5px;box-sizing:border-box"></td>
      </tr>`;
    }).join("");
  } else {
    document.getElementById("review-section").style.display = "none";
  }
  document.getElementById("recon-preview-section").style.display = "block";
  document.getElementById("recon-status").style.display = "none";
}


async function saveRecon() {
  if (!reconTxns.length) return alert('No transactions to save.');
  const account = document.getElementById('rec-account').value.trim();
  if (!account) return alert('Please select or enter the bank account this statement belongs to.');
  await registerAccountIfNew(account);

  // Collect corrections from the review queue and apply back to full transaction list
  const reviewCorrections = {};
  const newTypes = [];
  (window._reviewTxns || []).forEach((tx, i) => {
    const sel = document.querySelector(`#review-tbody select[data-row="${i}"]`);
    const inp = document.querySelector(`input[data-review-row="${i}"]`);
    const type_override = sel ? sel.value.trim() : '';
    const row_remarks   = inp ? inp.value.trim() : '';
    if (row_remarks && !BASE_TYPES.includes(row_remarks) && !_customTypes.includes(row_remarks) && !newTypes.includes(row_remarks))
      newTypes.push(row_remarks);
    // Match back to full list by date+description+amount
    reviewCorrections[`${tx.date}|${tx.description}|${tx.debit}|${tx.credit}`] = {type_override, row_remarks};
  });
  for (const t of newTypes) await saveCustomType(t);

  const txns = reconTxns.map(tx => {
    const key = `${tx.date}|${tx.description}|${tx.debit}|${tx.credit}`;
    const fix = reviewCorrections[key] || {};
    return {...tx, type_override: fix.type_override||'', row_remarks: fix.row_remarks||''};
  });

  const remarksMap = {};

  const data = {
    date:        document.getElementById('rec-date').value.trim(),
    opening:     parseFloat(document.getElementById('rec-opening').value) || 0,
    closing:     parseFloat(document.getElementById('rec-closing').value) || 0,
    remarks:     document.getElementById('rec-remarks').value.trim(),
    account:     account,
    transactions: txns,
    remarks_map: remarksMap,
  };
  const btn = event.target; btn.disabled = true; btn.textContent = 'Saving...';
  try {
    const resp = await fetch('/reconcile/save', {method:'POST',
      headers:{'Content-Type':'application/json'}, body: JSON.stringify(data)});
    if (!resp.ok) {
      const err = await resp.json().catch(() => ({error: 'Unknown error'}));
      throw new Error(err.error || `HTTP ${resp.status}`);
    }
    const rowsSaved = resp.headers.get('X-Rows-Saved') || txns.length;
    const closingBal = resp.headers.get('X-Closing-Balance') || data.closing;
    const blob = await resp.blob();
    const disposition = resp.headers.get('Content-Disposition') || '';
    const match = disposition.match(/filename="(.+)"/);
    const filename = match ? match[1] : 'Daily Reconciliation.xlsx';
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    showStatus('recon-status','success',
      `✅ Downloaded ${filename} (full history) — ${rowsSaved} transactions this period, Closing Balance: ₹${fmtDec(closingBal)}`);
  } catch (e) {
    showStatus('recon-status','error','❌ ' + e.message);
  } finally {
    btn.disabled = false; btn.textContent = '💾 Save to Daily Reconciliation Excel';
  }
}

// ── Settings tab ─────────────────────────────────────────────────────────────
let _allContacts = [];
let _contactsDirty = false;

async function loadContacts() {
  document.getElementById('contacts-body').innerHTML = '<p style="color:#888;font-size:.85rem">Loading…</p>';
  const d = await (await fetch('/contacts')).json();
  if (!d.ok && d.error) {
    document.getElementById('contacts-body').innerHTML =
      `<p style="color:#c00;font-size:.85rem">⚠️ Could not load contacts: ${d.error}</p>`;
    return;
  }
  _allContacts = (d.contacts || []).map((c,i) => ({...c, _id: i}));
  _contactsDirty = false;
  renderContactsTable(_allContacts);
}

function filterContacts(q) {
  const ql = q.toLowerCase();
  const filtered = ql ? _allContacts.filter(c =>
    [c.cluster, c.name, c.branch, c.designation, c.phone, c.email].some(v => (v||'').toLowerCase().includes(ql))
  ) : _allContacts;
  renderContactsTable(filtered);
}

function _markDirty() {
  _contactsDirty = true;
  document.getElementById('contacts-save-btn').style.display = '';
}

function _cellInput(val, field, id, placeholder) {
  return `<input data-id="${id}" data-field="${field}" value="${(val||'').replace(/"/g,'&quot;')}"
    placeholder="${placeholder}" oninput="_editContact(this)"
    style="width:100%;border:none;background:transparent;font-size:.83rem;font-family:inherit;padding:0;outline:none;min-width:60px">`;
}

function renderContactsTable(list) {
  const clusters = {};
  list.forEach(c => { (clusters[c.cluster] = clusters[c.cluster]||[]).push(c); });
  const DESIG_ORDER = {'CLUSTER MANAGER':0,'BRANCH HEAD':1,'BRANCH MANAGER':1};
  const cols = ['Name','Designation','Branch','Phone','Email',''];
  let html = '';
  for (const [cluster, members] of Object.entries(clusters)) {
    members.sort((a,b) => (DESIG_ORDER[(a.designation||'').toUpperCase()]??9) - (DESIG_ORDER[(b.designation||'').toUpperCase()]??9));
    html += `<div style="margin-bottom:20px">
      <div style="display:flex;align-items:center;gap:8px;background:#1a3a5c;color:#fff;padding:6px 12px;border-radius:6px 6px 0 0">
        <input value="${cluster}" data-cluster-old="${cluster}" onchange="_renameCluster(this)"
          style="background:transparent;border:none;color:#fff;font-weight:700;font-size:.85rem;letter-spacing:.5px;flex:1;outline:none;cursor:pointer"
          title="Click to rename cluster">
        <button onclick="addContactRow('${cluster}')" title="Add person to this cluster"
          style="background:rgba(255,255,255,.2);border:none;color:#fff;border-radius:4px;padding:2px 8px;cursor:pointer;font-size:.8rem">+ Person</button>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:.83rem">
        <thead><tr style="background:#f0f4f8">
          ${cols.map(c => `<th style="padding:6px 10px;text-align:left;border-bottom:1px solid #dde;font-size:.8rem">${c}</th>`).join('')}
        </tr></thead><tbody>`;
    members.forEach((c,i) => {
      const isManager = (c.designation||'').toUpperCase().includes('CLUSTER');
      const bg = isManager ? '#fffbe6' : (i%2===0?'#fff':'#f9fbfd');
      html += `<tr style="background:${bg}" data-id="${c._id}">
        <td style="padding:4px 8px;border-bottom:1px solid #eef;font-weight:${isManager?700:400}">${_cellInput(c.name,'name',c._id,'Name')}</td>
        <td style="padding:4px 8px;border-bottom:1px solid #eef">${_cellInput(c.designation,'designation',c._id,'e.g. Branch Manager')}</td>
        <td style="padding:4px 8px;border-bottom:1px solid #eef">${_cellInput(c.branch,'branch',c._id,'Branch')}</td>
        <td style="padding:4px 8px;border-bottom:1px solid #eef">${_cellInput(c.phone,'phone',c._id,'Phone')}</td>
        <td style="padding:4px 8px;border-bottom:1px solid #eef">${_cellInput(c.email,'email',c._id,'Email')}</td>
        <td style="padding:4px 8px;border-bottom:1px solid #eef;text-align:center">
          <button onclick="_deleteContact(${c._id})" title="Remove"
            style="background:none;border:none;color:#c00;cursor:pointer;font-size:1rem;line-height:1">🗑</button>
        </td></tr>`;
    });
    html += '</tbody></table></div>';
  }
  document.getElementById('contacts-body').innerHTML = html || '<p style="color:#888">No contacts found.</p>';
}

function _editContact(inp) {
  const id = parseInt(inp.dataset.id), field = inp.dataset.field;
  const c = _allContacts.find(x => x._id === id);
  if (c) { c[field] = inp.value; _markDirty(); }
}

function _renameCluster(inp) {
  const oldName = inp.dataset.clusterOld, newName = inp.value.trim();
  if (!newName || newName === oldName) return;
  _allContacts.forEach(c => { if (c.cluster === oldName) c.cluster = newName; });
  inp.dataset.clusterOld = newName;
  _markDirty();
}

function _deleteContact(id) {
  _allContacts = _allContacts.filter(c => c._id !== id);
  _markDirty();
  filterContacts(document.getElementById('contact-search').value);
}

function addContactRow(cluster) {
  const clusters = [...new Set(_allContacts.map(c => c.cluster))];
  const targetCluster = cluster || clusters[0] || 'New Cluster';
  const newId = Math.max(0, ..._allContacts.map(c => c._id)) + 1;
  // Insert after last member of that cluster
  const idx = _allContacts.map(c=>c.cluster).lastIndexOf(targetCluster);
  const newRow = {_id: newId, cluster: targetCluster, name:'', designation:'Branch Manager', branch:'', phone:'', email:''};
  if (idx >= 0) _allContacts.splice(idx + 1, 0, newRow);
  else _allContacts.push(newRow);
  _markDirty();
  filterContacts(document.getElementById('contact-search').value);
  // Focus the new row's name input
  setTimeout(() => {
    const inp = document.querySelector(`input[data-id="${newId}"][data-field="name"]`);
    if (inp) inp.focus();
  }, 50);
}

function addCluster() {
  const name = prompt('New cluster name:');
  if (!name || !name.trim()) return;
  const newId = Math.max(0, ..._allContacts.map(c => c._id)) + 1;
  _allContacts.push({_id: newId, cluster: name.trim(), name:'', designation:'CLUSTER MANAGER', branch:'', phone:'', email:''});
  _markDirty();
  filterContacts(document.getElementById('contact-search').value);
}

async function saveContacts() {
  const btn = document.getElementById('contacts-save-btn');
  btn.textContent = '⏳ Saving…'; btn.disabled = true;
  const payload = _allContacts.filter(c => c.name.trim()).map(({_id, ...rest}) => rest);
  const r = await (await fetch('/contacts', {method:'POST',
    headers:{'Content-Type':'application/json'}, body: JSON.stringify({contacts: payload})})).json();
  btn.textContent = '💾 Save'; btn.disabled = false;
  if (r.ok) {
    _contactsDirty = false; btn.style.display = 'none';
    showStatus('contacts-status', 'success', '✅ Contacts saved to Google Sheet');
  } else {
    showStatus('contacts-status', 'error', '❌ ' + r.error);
  }
}

async function loadSettings() {
  const cfg = await (await fetch('/config')).json();
  document.getElementById('report-time').value = cfg.report_time || '09:00';
  document.getElementById('set-bulk-debit').value = cfg.bulk_debit_account || '';
  document.getElementById('set-bulk-narration').value = cfg.bulk_narration || '';
}

async function saveSettings() {
  const cfg = {
    report_time: document.getElementById('report-time').value,
    bulk_debit_account: document.getElementById('set-bulk-debit').value.trim(),
    bulk_narration: document.getElementById('set-bulk-narration').value.trim(),
  };
  const r = await (await fetch('/config', {method:'POST',
    headers:{'Content-Type':'application/json'}, body: JSON.stringify(cfg)})).json();
  showStatus('settings-status', r.ok ? 'success' : 'error', r.ok ? '✅ Settings saved!' : '❌ '+r.error);
}

// ── Bank File tab ─────────────────────────────────────────────────────────────
let bulkRequests = [];

function _fmtInr(v) {
  const n = Number(v);
  return n > 0 ? '₹' + n.toLocaleString('en-IN', {maximumFractionDigits:0}) : (v || '—');
}

async function loadBulkRequests() {
  const tbody = document.getElementById('bulk-tbody');
  const empty = document.getElementById('bulk-empty');
  const table = document.getElementById('bulk-table');
  empty.style.display = ''; empty.textContent = 'Loading...';
  try {
    const items = await (await fetch('/requests/pending')).json();
    if (items.error) throw new Error(items.error);
    bulkRequests = items;
    document.getElementById('bulk-count').textContent = items.length ? `(${items.length})` : '';
    if (!items.length) {
      table.style.display = 'none';
      empty.textContent = 'No pending requests';
      updateBulkFooter();
      return;
    }
    tbody.innerHTML = '';
    items.forEach(r => {
      const tr = document.createElement('tr');
      tr.style.borderBottom = '1px solid #eef3f9';
      tr.innerHTML = `
        <td style="padding:6px 8px;"><input type="checkbox" class="bulk-check" value="${r.request_id}" onchange="updateBulkFooter()"></td>
        <td style="padding:6px 8px;white-space:nowrap;color:#2a7ae2;">${r.request_id}</td>
        <td style="padding:6px 8px;white-space:nowrap;font-size:12px;color:#888;">${r.submitted_at}</td>
        <td style="padding:6px 8px;font-weight:600;">${r.customer}</td>
        <td style="padding:6px 8px;">${r.cluster} / ${r.branch}</td>
        <td style="padding:6px 8px;white-space:nowrap;">${_fmtInr(r.amount)}</td>
        <td style="padding:6px 8px;font-family:monospace;">${r.account_no || '—'}</td>
        <td style="padding:6px 8px;font-family:monospace;">${r.ifsc || '—'}</td>
        <td style="padding:6px 8px;">${r.bank || '—'}</td>
        <td style="padding:6px 8px;">${r.phone || '—'}</td>
        <td style="padding:6px 8px;">${r.so_name || '—'}</td>
        <td style="padding:6px 8px;">${r.gold_weight || '—'}</td>`;
      tbody.appendChild(tr);
    });
    table.style.display = '';
    empty.style.display = 'none';
    document.getElementById('bulk-select-all').checked = false;
    updateBulkFooter();
  } catch(e) {
    table.style.display = 'none';
    empty.textContent = 'Error: ' + e.message;
  }
}

function toggleBulkSelectAll(checked) {
  document.querySelectorAll('.bulk-check').forEach(c => c.checked = checked);
  updateBulkFooter();
}

function _selectedBulkIds() {
  return Array.from(document.querySelectorAll('.bulk-check:checked')).map(c => c.value);
}

function updateBulkFooter() {
  const ids = _selectedBulkIds();
  const footer = document.getElementById('bulk-footer');
  if (!ids.length) { footer.style.display = 'none'; return; }
  const total = bulkRequests.filter(r => ids.includes(r.request_id))
    .reduce((s, r) => s + (Number(String(r.amount).replace(/[^\\d.]/g,'')) || 0), 0);
  footer.style.display = '';
  footer.textContent = `${ids.length} selected · total ${_fmtInr(total)}`;
}

async function exportBulk() {
  const ids = _selectedBulkIds();
  if (!ids.length) { showStatus('bulk-status', 'error', '❌ Select at least one request'); return; }
  const body = {
    request_ids: ids,
    bank: document.getElementById('bulk-bank').value,
    debit_account: document.getElementById('bulk-debit').value.trim(),
    narration: document.getElementById('bulk-narration').value.trim(),
    value_date: document.getElementById('bulk-value-date').value.trim(),
  };
  showStatus('bulk-status', 'success', '⏳ Generating...');
  try {
    const r = await fetch('/requests/export-bulk', {method:'POST',
      headers:{'Content-Type':'application/json'}, body: JSON.stringify(body)});
    if (!r.ok) {
      const err = await r.json().catch(() => ({error: `HTTP ${r.status}`}));
      if (r.status === 409) {
        showStatus('bulk-status', 'error',
          `❌ Already exported/disbursed: ${(err.request_ids||[]).join(', ')} — list refreshed`);
        loadBulkRequests();
        return;
      }
      throw new Error(err.error || `HTTP ${r.status}`);
    }
    const blob = await r.blob();
    const disposition = r.headers.get('Content-Disposition') || '';
    const match = disposition.match(/filename="(.+)"/);
    const filename = match ? match[1] : 'bulk_payment.csv';
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    showStatus('bulk-status', 'success', `✅ Downloaded ${filename} — ${ids.length} request(s) marked Exported`);
    loadBulkRequests();
  } catch(e) {
    showStatus('bulk-status', 'error', '❌ ' + e.message);
  }
}

async function deleteBulkRequests() {
  const ids = _selectedBulkIds();
  if (!ids.length) { showStatus('bulk-delete-status', 'error', '❌ Select at least one request'); return; }
  const reason = prompt(`Delete ${ids.length} request(s) from the queue?\n\nReason (optional — saved to the sheet):`);
  if (reason === null) return;
  if (!confirm(`Confirm: remove ${ids.join(', ')} from the queue? The row stays in the sheet marked Deleted.`)) return;
  try {
    const r = await (await fetch('/requests/delete', {method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({request_ids: ids, reason: reason.trim()})})).json();
    if (!r.ok) throw new Error(r.error || 'Failed');
    showStatus('bulk-delete-status', 'success', `✅ Deleted ${r.deleted} request(s)`);
    loadBulkRequests();
  } catch(e) {
    showStatus('bulk-delete-status', 'error', '❌ ' + e.message);
  }
}

async function parseBulkMsg() {
  const msg = document.getElementById('bulk-paste').value;
  if (!msg.trim()) return;
  const r = await (await fetch('/requests/parse', {method:'POST',
    headers:{'Content-Type':'application/json'}, body: JSON.stringify({message: msg})})).json();
  const f = r.fields || {};
  document.getElementById('bp-customer').value = f.customer || '';
  document.getElementById('bp-amount').value = f.amount || '';
  document.getElementById('bp-account').value = f.account_no || '';
  document.getElementById('bp-ifsc').value = f.ifsc || '';
  document.getElementById('bp-bank').value = f.bank || '';
  document.getElementById('bp-phone').value = f.phone || '';
  document.getElementById('bp-cluster').value = f.cluster || '';
  document.getElementById('bp-branch').value = f.branch || '';
  document.getElementById('bp-so').value = f.so_name || '';
  document.getElementById('bp-gold').value = f.gold_weight || '';
  document.getElementById('bulk-parse-warnings').innerHTML =
    (r.warnings || []).map(w => `⚠ ${w}`).join('<br>');
  document.getElementById('bulk-parse-preview').style.display = '';
}

async function addParsedRequest() {
  const body = {
    customer: document.getElementById('bp-customer').value.trim(),
    amount: document.getElementById('bp-amount').value.trim(),
    account_no: document.getElementById('bp-account').value.trim(),
    ifsc: document.getElementById('bp-ifsc').value.trim().toUpperCase(),
    bank: document.getElementById('bp-bank').value.trim(),
    phone: document.getElementById('bp-phone').value.trim(),
    cluster: document.getElementById('bp-cluster').value.trim(),
    branch: document.getElementById('bp-branch').value.trim(),
    so_name: document.getElementById('bp-so').value.trim(),
    gold_weight: document.getElementById('bp-gold').value.trim(),
  };
  const r = await (await fetch('/requests/add', {method:'POST',
    headers:{'Content-Type':'application/json'}, body: JSON.stringify(body)})).json();
  if (r.ok) {
    showStatus('bulk-paste-status', 'success', `✅ Added to queue as ${r.request_id}`);
    document.getElementById('bulk-paste').value = '';
    document.getElementById('bulk-parse-preview').style.display = 'none';
    loadBulkRequests();
  } else {
    showStatus('bulk-paste-status', 'error', '❌ ' + (r.error || 'Failed'));
  }
}

window.onload = async () => {
  const today = new Date().toLocaleDateString('en-GB').replace(/\\//g,'-');
  document.getElementById('d-date').value = today;
  document.getElementById('r-date').value = today;
  document.getElementById('rec-date').value = today;
  document.getElementById('bulk-value-date').value = new Date().toLocaleDateString('en-GB');
  await loadCustomTypes();
  await loadBankAccounts();
  loadSummary();
  loadPendingRequests();
  try {
    const cfg = await (await fetch('/config')).json();
    document.getElementById('bulk-debit').value = cfg.bulk_debit_account || '';
    document.getElementById('bulk-narration').value = cfg.bulk_narration || '';
  } catch(e) {}
};
</script>
</body></html>"""

@app.route('/')
def index():
    return render_template_string(HTML)

@app.route('/setup')
def setup():
    return render_template_string(SETUP_HTML)

@app.route('/extract/disbursement', methods=['POST'])
def api_extract_disbursement():
    return jsonify(extract_disbursement(request.json.get('message', '')))

@app.route('/extract/repayment', methods=['POST'])
def api_extract_repayment():
    return jsonify(extract_repayment(request.json.get('message', '')))

@app.route('/open-cases')
def api_open_cases():
    try:
        return jsonify(get_open_cases())
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def _mark_request_disbursed(request_id, disb_id):
    _, ws = get_requests_sheet()
    all_vals = ws.get_all_values()
    for i, row in enumerate(all_vals[1:], start=2):
        if row and row[0] == request_id:
            ws.update_cell(i, 12, 'Disbursed')
            ws.update_cell(i, 13, disb_id)
            return

def _request_row_to_item(row, idx):
    def col(name, default_i):
        i = idx.get(name, default_i)
        return row[i] if len(row) > i else ''
    return {
        'request_id':  col('Request ID', 0),
        'submitted_at': col('Submitted At', 1),
        'customer':    col('Customer Name', 2),
        'cluster':     col('Cluster', 3),
        'branch':      col('Branch', 4),
        'amount':      col('Amount', 5),
        'account_no':  col('Account No', 6),
        'ifsc':        col('IFSC', 7),
        'phone':       col('Phone', 8),
        'so_name':     col('SO Name', 9),
        'gold_weight': col('Gold Weight', 10),
        'status':      col('Status', 11),
        'bank':        col('Bank', 14),
    }

@app.route('/requests/pending')
def api_requests_pending():
    try:
        # ?include=exported keeps just-exported requests visible (used by the
        # New Disbursement picker so click-to-fill still works after a bank
        # run); the Bank File tab calls this plain and sees Pending only.
        wanted = {'Pending'}
        if request.args.get('include') == 'exported':
            wanted.add('Exported')
        _, ws = get_requests_sheet()
        all_vals = ws.get_all_values()
        if len(all_vals) < 2:
            return jsonify([])
        idx = {h: i for i, h in enumerate(all_vals[0])}
        items = []
        for row in reversed(all_vals[1:]):
            if not row or not row[0]:
                continue
            status = row[idx.get('Status', 11)] if len(row) > 11 else ''
            if status not in wanted:
                continue
            items.append(_request_row_to_item(row, idx))
        return jsonify(items)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def _build_bulk_csv(template, rows):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow([h for h, _ in template['columns']])
    for r in rows:
        w.writerow([_resolve_bulk_field(f, r) for _, f in template['columns']])
    return buf.getvalue().encode('utf-8')

def _build_bulk_xlsx(template, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([h for h, _ in template['columns']])
    for r in rows:
        ws.append([_resolve_bulk_field(f, r) for _, f in template['columns']])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def _resolve_bulk_field(field, row):
    if isinstance(field, tuple) and field[0] == 'const':
        return field[1]
    return row.get(field, '')

def _clean_request_amount(raw):
    """'₹11,75,000' / 'Rs. 1175000.00' → numeric; portals reject ₹ and commas.
    Extracts the first number pattern rather than stripping globally, so the
    dot in a 'Rs.' prefix can't corrupt the value. Distinct from the recon
    _clean_amount(), which must preserve negative signs."""
    m = re.search(r'\d[\d,]*(?:\.\d{1,2})?', str(raw or ''))
    if not m:
        return 0
    val = float(m.group(0).replace(',', ''))
    return int(val) if val == int(val) else val

@app.route('/requests/export-bulk', methods=['POST'])
def api_requests_export_bulk():
    try:
        body = request.json or {}
        ids = body.get('request_ids') or []
        bank_key = body.get('bank', '')
        template = BANK_TEMPLATES.get(bank_key)
        if not template:
            return jsonify({'ok': False, 'error': f'Unknown bank format: {bank_key}'}), 400
        if not ids:
            return jsonify({'ok': False, 'error': 'No requests selected'}), 400

        cfg = load_config()
        debit_account = body.get('debit_account') or cfg.get('bulk_debit_account', '')
        narration = body.get('narration') or cfg.get('bulk_narration', '')
        value_date = body.get('value_date') or datetime.now(IST).strftime('%d/%m/%Y')

        _, ws = get_requests_sheet()
        all_vals = ws.get_all_values()
        idx = {h: i for i, h in enumerate(all_vals[0])} if all_vals else {}
        by_id = {}
        for rownum, row in enumerate(all_vals[1:], start=2):
            if row and row[0] in ids:
                by_id[row[0]] = (rownum, row)

        missing = [i for i in ids if i not in by_id]
        if missing:
            return jsonify({'ok': False, 'error': 'Requests not found',
                            'request_ids': missing}), 404
        # Double-export guard: anything not Pending has already been exported
        # or disbursed — refuse the whole batch so nothing is paid twice.
        conflicts = [i for i in ids
                     if by_id[i][1][idx.get('Status', 11)] != 'Pending']
        if conflicts:
            return jsonify({'ok': False, 'error': 'Not pending (already exported or disbursed?)',
                            'request_ids': conflicts}), 409

        export_rows = []
        for req_id in ids:
            item = _request_row_to_item(by_id[req_id][1], idx)
            amount = _clean_request_amount(item['amount'])
            export_rows.append({
                **item,
                'amount': amount,
                'mode': 'RTGS' if amount >= RTGS_MIN_AMOUNT else 'NEFT',
                'debit_account': debit_account,
                'narration': narration,
                'value_date': value_date,
                'bank_name': item['bank'],
            })

        if template['filetype'] == 'csv':
            file_bytes = _build_bulk_csv(template, export_rows)
            mimetype = 'text/csv'
        else:
            file_bytes = _build_bulk_xlsx(template, export_rows)
            mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        filename = template['filename'].format(date=datetime.now(IST).strftime('%d-%m-%Y'))

        # Mark AFTER the file is built (a marked-but-unbuilt file is the bad
        # direction), in ONE batch_update so a partial write can't split the
        # batch between Exported and Pending.
        stamp = f"Exported {template['label']} {datetime.now(IST).strftime('%d-%m-%Y %H:%M')} IST"
        notes_col = idx.get('Notes', 13)
        status_letter = chr(ord('A') + idx.get('Status', 11))
        notes_letter = chr(ord('A') + notes_col)
        updates = []
        for req_id in ids:
            rownum, row = by_id[req_id]
            old_note = row[notes_col] if len(row) > notes_col else ''
            note = f"{old_note} | {stamp}" if old_note.strip() else stamp
            updates.append({'range': f'{status_letter}{rownum}', 'values': [['Exported']]})
            updates.append({'range': f'{notes_letter}{rownum}', 'values': [[note]]})
        ws.batch_update(updates)

        return Response(file_bytes, mimetype=mimetype,
                        headers={'Content-Disposition': f'attachment; filename="{filename}"'})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

# ── WhatsApp DISBURSEMENT REQUEST message parser ─────────────────────────────
# Label-keyed, for the exact message format the field app generates. The
# generic extract_disbursement() stays untouched — it serves the disb tab's
# free-form notes; this one only trusts explicit "Label: value" lines.
REQUEST_LABELS = [
    ('customer',    ['customer name', 'name']),
    ('phone',       ['phone no', 'phone number', 'phone', 'mobile']),
    ('amount',      ['amount required', 'amount']),
    ('account_no',  ['account no', 'account number', 'a/c no', 'ac no']),
    ('ifsc',        ['ifsc code', 'ifsc']),
    ('bank',        ['bank name', 'bank']),
    ('cluster',     ['cluster']),
    ('branch',      ['branch']),
    ('so_name',     ['so name']),
    ('gold_weight', ['gold wt (gms)', 'gold wt', 'gold weight']),
    ('date',        ['date']),
]

def parse_request_message(text):
    fields, warnings = {}, []
    # Longest label first so "Customer Name" wins over "Name" and
    # "IFSC Code" over "IFSC".
    label_map = sorted(
        [(lbl, key) for key, labels in REQUEST_LABELS for lbl in labels],
        key=lambda x: -len(x[0]))
    for line in (text or '').splitlines():
        m = re.match(r'^\s*([^:\-]+?)\s*[:\-]\s*(.+?)\s*$', line)
        if not m:
            continue
        label = re.sub(r'\s+', ' ', m.group(1).strip().lower())
        value = m.group(2).strip()
        for lbl, key in label_map:
            if label == lbl and key not in fields:
                fields[key] = value
                break
    if 'amount' in fields:
        fields['amount'] = _clean_request_amount(fields['amount'])
        if not fields['amount']:
            warnings.append('Amount could not be read as a number')
    if 'ifsc' in fields:
        fields['ifsc'] = fields['ifsc'].upper().replace(' ', '')
        if not re.match(r'^[A-Z]{4}0[A-Z0-9]{6}$', fields['ifsc']):
            warnings.append(f"IFSC '{fields['ifsc']}' doesn't look valid (AAAA0XXXXXX)")
    if 'phone' in fields:
        fields['phone'] = re.sub(r'\D', '', fields['phone'])
        if len(fields['phone']) != 10:
            warnings.append(f"Phone '{fields['phone']}' is not 10 digits")
    if 'account_no' in fields:
        fields['account_no'] = re.sub(r'\s', '', fields['account_no'])
    for required in ('customer', 'amount', 'account_no', 'ifsc'):
        if not fields.get(required):
            warnings.append(f'Missing: {required}')
    return fields, warnings

def generate_request_id(ws):
    # Ported verbatim from disbursement-form-deploy/api/submit-request.py so
    # widget-added and field-app requests share one REQ-DDMMYY-NNN sequence.
    # IST, not server time (Vercel's Python runtime is UTC) — otherwise a
    # request added after 5:30pm IST gets tomorrow's date-prefix.
    today = datetime.now(IST)
    prefix = f"REQ-{today.strftime('%d%m%y')}-"
    seq = 1
    for row in ws.get_all_values()[1:]:
        if row and row[0].startswith(prefix):
            try:
                n = int(row[0].split("-")[-1])
                if n >= seq:
                    seq = n + 1
            except ValueError:
                pass
    return f"{prefix}{seq:03d}"

@app.route('/requests/parse', methods=['POST'])
def api_requests_parse():
    fields, warnings = parse_request_message((request.json or {}).get('message', ''))
    return jsonify({'fields': fields, 'warnings': warnings})

@app.route('/requests/delete', methods=['POST'])
def api_requests_delete():
    try:
        body = request.json or {}
        ids = body.get('request_ids') or []
        reason = (body.get('reason') or '').strip()
        if not ids:
            return jsonify({'ok': False, 'error': 'No requests selected'}), 400
        _, ws = get_requests_sheet()
        all_vals = ws.get_all_values()
        idx = {h: i for i, h in enumerate(all_vals[0])} if all_vals else {}
        by_id = {}
        for rownum, row in enumerate(all_vals[1:], start=2):
            if row and row[0] in ids:
                by_id[row[0]] = (rownum, row)
        missing = [i for i in ids if i not in by_id]
        if missing:
            return jsonify({'ok': False, 'error': 'Requests not found',
                            'request_ids': missing}), 404
        # Soft delete: Status → Deleted, reason + timestamp appended to Notes.
        # The row stays in the sheet as the audit trail of what was removed
        # and why; every queue view filters on Status so it disappears
        # from pending/export/disb-picker immediately.
        stamp = f"Deleted{': ' + reason if reason else ''} {datetime.now(IST).strftime('%d-%m-%Y %H:%M')} IST"
        notes_col = idx.get('Notes', 13)
        status_letter = chr(ord('A') + idx.get('Status', 11))
        notes_letter = chr(ord('A') + notes_col)
        updates = []
        for req_id in ids:
            rownum, row = by_id[req_id]
            old_note = row[notes_col] if len(row) > notes_col else ''
            note = f"{old_note} | {stamp}" if old_note.strip() else stamp
            updates.append({'range': f'{status_letter}{rownum}', 'values': [['Deleted']]})
            updates.append({'range': f'{notes_letter}{rownum}', 'values': [[note]]})
        ws.batch_update(updates)
        return jsonify({'ok': True, 'deleted': len(ids)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

def _flag_id_collision(ws, req_id):
    """Best-effort detection (not prevention) of the ID-generation race: two
    concurrent submissions can compute the same REQ-DDMMYY-NNN before either
    has appended its row. Re-scanning right after our own append can catch
    it — not airtight (both racers may read a self-consistent but stale
    view), but turns a silent duplicate ID into a flagged one a human will
    notice instead of nothing. Mirrors disbursement-form-deploy/api/
    submit-request.py so both writers behave the same way. Never blocks the
    submission itself."""
    try:
        all_vals = ws.get_all_values()
        header = all_vals[0]
        notes_col = header.index('Notes') if 'Notes' in header else 13
        matches = [i for i, row in enumerate(all_vals[1:], start=2) if row and row[0] == req_id]
        if len(matches) <= 1:
            return
        marker = f"ID COLLISION: {len(matches)} requests share {req_id} - verify manually"
        for rownum in matches:
            row = all_vals[rownum - 1]
            old_note = row[notes_col] if len(row) > notes_col else ''
            if 'ID COLLISION' in old_note:
                continue
            note = f"{old_note} | {marker}" if old_note.strip() else marker
            ws.update_cell(rownum, notes_col + 1, note)
    except Exception:
        pass  # auxiliary safety net — a failure here must never surface as a submit failure

def _find_duplicate_request(all_vals, account_no, amount):
    """Same beneficiary account + same amount submitted today = duplicate,
    whatever its status — a request already exported or disbursed today is
    the worst kind of duplicate. Digits-only compare so '2,00,000' matches."""
    digits = lambda s: re.sub(r'\D', '', str(s or ''))
    today = datetime.now(IST).strftime('%d-%m-%Y')
    acct, amt = digits(account_no), digits(amount)
    if not acct or not amt:
        return None
    for row in all_vals[1:]:
        if len(row) < 12 or not row[0] or not row[1].startswith(today):
            continue
        if row[11] == 'Deleted':
            continue  # deleted requests don't block resubmission
        if digits(row[6]) == acct and digits(row[5]) == amt:
            return {'request_id': row[0], 'submitted_at': row[1], 'status': row[11]}
    return None

@app.route('/requests/add', methods=['POST'])
def api_requests_add():
    try:
        body = request.json or {}
        if not body.get('customer') or not body.get('account_no'):
            return jsonify({'ok': False, 'error': 'Customer and account number are required'}), 400
        _, ws = get_requests_sheet()
        all_vals = ws.get_all_values()
        dup = _find_duplicate_request(all_vals, body.get('account_no'), body.get('amount'))
        if dup:
            return jsonify({'ok': False, 'duplicate': True,
                            'error': f"Duplicate of {dup['request_id']} ({dup['status']}, {dup['submitted_at']})"}), 409
        req_id = generate_request_id(ws)
        ws.append_row([
            req_id,
            datetime.now(IST).strftime("%d-%m-%Y %H:%M IST"),
            body.get('customer', ''),
            body.get('cluster', ''),
            body.get('branch', ''),
            str(body.get('amount', '')),
            body.get('account_no', ''),
            body.get('ifsc', ''),
            body.get('phone', ''),
            body.get('so_name', ''),
            str(body.get('gold_weight', '')),
            'Pending',
            '',
            'Added via widget paste',
            body.get('bank', ''),
        ])
        _flag_id_collision(ws, req_id)
        return jsonify({'ok': True, 'request_id': req_id})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/save/disbursement', methods=['POST'])
def api_save_disbursement():
    try:
        return jsonify({'ok': True, 'disb_id': save_disbursement(request.json)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})

@app.route('/save/repayment', methods=['POST'])
def api_save_repayment():
    try:
        return jsonify({'ok': True, **save_repayment(request.json)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})

@app.route('/case/recent')
def api_case_recent():
    try:
        days = int(request.args.get('days', 7))
        return jsonify({'ok': True, 'events': get_recent_activity(days)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/case/<disb_id>/detail')
def api_case_detail(disb_id):
    try:
        detail = _case_detail(disb_id)
        if not detail:
            return jsonify({'ok': False, 'error': f'{disb_id.upper()} not found'}), 404
        return jsonify({'ok': True, 'case': detail})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/case/update-disbursement', methods=['POST'])
def api_case_update_disbursement():
    try:
        known = update_disbursement(request.json)
        return jsonify({'ok': True, 'case': _case_detail(request.json['disb_id'], known=known)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})

@app.route('/case/update-repayment', methods=['POST'])
def api_case_update_repayment():
    try:
        known = update_repayment(request.json)
        return jsonify({'ok': True, 'case': _case_detail(request.json['disb_id'], known=known)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})

@app.route('/case/delete-repayment', methods=['POST'])
def api_case_delete_repayment():
    try:
        known = delete_repayment(request.json)
        return jsonify({'ok': True, 'case': _case_detail(request.json['disb_id'], known=known)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})

@app.route('/case/add-repayment', methods=['POST'])
def api_case_add_repayment():
    try:
        known = add_repayment(request.json)
        return jsonify({'ok': True, 'case': _case_detail(request.json['disb_id'], known=known)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})

def _load_for_invoice(sh, disb_id):
    """Read only Accounts + M Coll (2 API calls) instead of the full
    load_data_from_sheet() which also reads archives, Contact, DashBoard."""
    acc_vals = sh.worksheet(SHEET_NAME).get_all_values()
    header_idx = next((i for i, r in enumerate(acc_vals) if r and 'Disbursement ID' in r), 1)
    raw_rows = [r for r in acc_vals[header_idx + 1:] if r and str(r[0]).strip().startswith('BLP-')]

    # Only scan archive tabs if the case isn't in the main Accounts sheet
    if not any(str(r[0]).strip().upper() == disb_id for r in raw_rows):
        for archive_name in get_archive_tab_names():
            try:
                archive_rows = [r for r in sh.worksheet(archive_name).get_all_values()
                                if r and str(r[0]).strip().startswith('BLP-')]
                raw_rows += archive_rows
                if any(str(r[0]).strip().upper() == disb_id for r in archive_rows):
                    break
            except gspread.exceptions.WorksheetNotFound:
                continue

    rows = []
    for raw in raw_rows:
        row = list(raw[:22]) + [None] * max(0, 22 - len(raw))
        row = [None if c == '' else c for c in row]
        for i in _MIS_NUMERIC_COLS:
            row[i] = _clean_numeric_cell(row[i])
        if row[0]:
            rows.append(row)

    sheets = {}
    try:
        mcoll_vals = sh.worksheet('M Coll').get_all_values()
        mcoll_vals = [[_clean_numeric_cell(c) if i == 2 else c for i, c in enumerate(r)]
                      for r in mcoll_vals]
        sheets['M Coll'] = _SheetShim(mcoll_vals)
    except gspread.exceptions.WorksheetNotFound:
        pass

    mcoll = mis.load_mcoll(_WorkbookShim(sheets))
    return rows, mcoll

@app.route('/generate-invoice', methods=['POST'])
def api_generate_invoice():
    try:
        disb_id = (request.json or {}).get('disb_id', '').strip().upper()
        if not disb_id:
            return jsonify({'error': 'disb_id required'}), 400
        ensure_mis_assets_cached()
        sh = get_gspread_client().open_by_key(SPREADSHEET_ID)
        rows, mcoll = _load_for_invoice(sh, disb_id)
        _, _, all_cases_full = mis.parse_cases(rows, mcoll)
        case = next((c for c in all_cases_full if c['id'].upper() == disb_id), None)
        if not case:
            return jsonify({'error': f'{disb_id} not found'}), 404
        mcoll_entry = mcoll.get(disb_id)
        paid = case['status'] == 'Closed'
        pdf_bytes = mis.generate_invoice_ledger(case, mcoll_entry=mcoll_entry, paid_in_full=paid)
        safe = case['customer'].replace('/', '-')
        filename = f"{disb_id} {safe} Invoice and Ledger.pdf"
        return Response(pdf_bytes, mimetype='application/pdf',
            headers={'Content-Disposition': f'attachment; filename="{filename}"'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generate-mis', methods=['POST'])
def api_generate_mis():
    """Runs the daily MIS pipeline live from this Google Sheet (no manual
    Excel export / CLI step) and returns the same ZIP generate_mis.py's CLI
    path produces."""
    try:
        ensure_mis_assets_cached()

        sh = get_gspread_client().open_by_key(SPREADSHEET_ID)
        rows, db_raw, mcoll, (cluster_mgrs, branch_contacts) = load_data_from_sheet(sh)
        open_cases, all_cases, all_cases_full = mis.parse_cases(rows, mcoll)
        metrics = mis.compute_dashboard_metrics(all_cases_full, db_raw)

        try:
            mis.write_claude_dashboard_to_sheet(sh, metrics)
        except Exception as e:
            print(f'WARNING: Could not write Claude_Dashboard: {e}')

        active_clusters = sorted(set(c['cluster'] for c in open_cases))
        zip_bytes = mis.build_zip(open_cases, all_cases, all_cases_full, metrics,
                                   cluster_mgrs, branch_contacts, active_clusters, mcoll_raw=mcoll)

        date_human = mis.TODAY.strftime('%d-%b-%Y')
        filename = f'{date_human} BridgeLine MIS Package.zip'
        return Response(
            zip_bytes,
            mimetype='application/zip',
            headers={'Content-Disposition': f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@app.route('/summary')
def api_summary():
    try:
        return jsonify(get_today_summary())
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/config', methods=['GET'])
def api_get_config():
    return jsonify(load_config())

@app.route('/config', methods=['POST'])
def api_save_config():
    try:
        cfg = load_config()
        cfg.update(request.json)
        save_config(cfg)
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})

@app.route('/reconcile/parse', methods=['POST'])
def api_reconcile_parse():
    if 'file' not in request.files:
        return jsonify({'ok': False, 'error': 'No file uploaded'})
    f = request.files['file']
    if not f.filename:
        return jsonify({'ok': False, 'error': 'Empty filename'})
    import tempfile
    suffix = '.' + f.filename.rsplit('.', 1)[-1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        f.save(tmp.name)
        tmp_path = tmp.name
    try:
        result = parse_bank_statement(tmp_path, f.filename)
        # Run classification immediately so UI can split confident vs review
        try:
            records = read_accounts_from_gsheet()
            sh      = get_gspread_client().open_by_key(SPREADSHEET_ID)
            mc_rows = sh.worksheet('M Coll').get_all_values()[1:]
        except Exception:
            records, mc_rows = [], []
        classified = _match_transactions(result['transactions'], records, mc_rows)

        def _needs_review(tx):
            if tx['match_basis'] in ('UTR', 'Manual'):         return False
            if tx['type'] in ('FD Booking', 'Test Credit'):    return False
            if tx['credit'] > 0 and tx['credit'] < 100:        return False
            return True

        confident = [tx for tx in classified if not _needs_review(tx)]
        review    = [tx for tx in classified if _needs_review(tx)]

        return jsonify({
            'ok': True,
            'opening_balance': result['opening_balance'],
            'closing_balance': result['closing_balance'],
            'transactions':    classified,   # full list for save
            'confident':       confident,
            'review':          review,
            'confident_count': len(confident),
            'review_count':    len(review),
        })
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})
    finally:
        os.unlink(tmp_path)

@app.route('/contacts', methods=['GET'])
def api_contacts_get():
    try:
        return jsonify({'ok': True, 'contacts': read_contacts()})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e), 'contacts': []})

@app.route('/contacts', methods=['POST'])
def api_contacts_post():
    try:
        contacts = request.json.get('contacts', [])
        save_contacts(contacts)
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})

@app.route('/reconcile/records', methods=['GET'])
def api_reconcile_records():
    try:
        records = read_accounts_from_gsheet()
        out = [{'id': r.get('Disbursement ID',''), 'customer': r.get('Customer Name',''),
                'branch': r.get('Branch','')} for r in records if r.get('Disbursement ID','').strip()]
        return jsonify({'ok': True, 'records': out})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e), 'records': []})

@app.route('/reconcile/save', methods=['POST'])
def api_reconcile_save():
    """Persists the period to the 'Recon Txns' sheet tab and returns the FULL
    cumulative reconciliation workbook (all saved periods) as a direct
    download — see save_reconciliation()'s docstring. No file re-upload is
    needed to keep history; the Google Sheet is the system of record.
    Summary numbers (rows saved, closing balance) are echoed back in response
    headers so the frontend can show a status message without parsing a JSON
    body on a binary response.
    """
    try:
        data = request.json
        result = save_reconciliation(
            recon_date=data['date'],
            opening_balance=float(data.get('opening') or 0),
            closing_balance=float(data.get('closing') or 0),
            transactions=data['transactions'],
            remarks=data.get('remarks', ''),
            remarks_map=data.get('remarks_map', {}),
            account=data.get('account', ''),
        )
        return Response(
            result['file_bytes'],
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={
                'Content-Disposition': f'attachment; filename="{result["filename"]}"',
                'X-Rows-Saved': str(result['rows_saved']),
                'X-Closing-Balance': str(result['closing']),
            },
        )
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

# Hosted on Vercel via api/index.py (WSGI). No local dev-server / ngrok /
# browser-open block needed here — that only applied to the Mac-only version.
