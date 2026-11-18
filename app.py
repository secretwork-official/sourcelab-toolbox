import csv, difflib, hashlib, html, io, ipaddress, json, os, re, secrets, shutil, subprocess, tempfile, threading, time, zipfile, zlib
from urllib.parse import urlparse
import barcode, markdown, pytesseract, qrcode, requests, xmltodict, yaml
from barcode.writer import ImageWriter
from flask import Flask, Response, abort, jsonify, request, send_file
from openpyxl import Workbook, load_workbook
from pdf2image import convert_from_bytes
from PIL import Image, ImageOps
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from werkzeug.middleware.proxy_fix import ProxyFix
try:
    import pillow_heif; pillow_heif.register_heif_opener()
except Exception:
    pass

Image.MAX_IMAGE_PIXELS = 40_000_000
MAX_MB = int(os.environ.get('MAX_MB', 20))
app = Flask(__name__, static_folder='static', static_url_path='/static')
app.config['MAX_CONTENT_LENGTH'] = MAX_MB * 1024 * 1024
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
FILES = os.path.join(tempfile.gettempdir(), 'hosted'); os.makedirs(FILES, exist_ok=True)

class Bad(Exception):
    def __init__(s, m, c=400): super().__init__(m); s.code = c

HITS = {}
def limit(kind, n, per=60):
    k = (kind, request.remote_addr); now = time.time()
    h = [t for t in HITS.get(k, []) if now - t < per]; h.append(now); HITS[k] = h
    if len(HITS) > 5000: HITS.clear()
    if len(h) > n: raise Bad('Too many requests. Please wait a while and try again.', 429)

def ip_hash(): return hashlib.sha256(('tb:' + (request.remote_addr or 'x')).encode()).hexdigest()[:32]
def stem(n): return re.sub(r'[^\w.-]', '_', os.path.splitext(os.path.basename(n or 'file'))[0])[:60] or 'file'
def num(v, d, lo, hi):
    try: x = float(v)
    except (TypeError, ValueError): x = d
    return min(max(x, lo), hi)
def F_(name, data, mime='application/octet-stream'): return ('file', name, data, mime)
def T_(s): return ('text', str(s))
def zip_(items):
    b = io.BytesIO()
    with zipfile.ZipFile(b, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, d in items: z.writestr(n, d)
    return F_('result.zip', b.getvalue(), 'application/zip')
def pack(items): return F_(*items[0]) if len(items) == 1 else zip_(items)
def need(fs, n=1, what='file'):
    if len(fs) < n: raise Bad(f'Please add at least {n} {what}(s).')

TOOLS = []
def tool(id, name, cat, desc, files=None, fields=(), heavy=False, rate=20):
    def d(fn): TOOLS.append(dict(id=id, name=name, cat=cat, desc=desc, files=files, fields=list(fields), heavy=heavy, rate=rate, h=fn)); return fn
    return d
def F(n, l, t='text', d='', o=None): return dict(n=n, l=l, t=t, d=d, o=o)
PDF1 = {'accept': '.pdf', 'multiple': False}; PDFN = {'accept': '.pdf', 'multiple': True}
IMG = {'accept': 'image/*,.heic,.heif', 'multiple': True}; AV = {'accept': 'audio/*,video/*', 'multiple': False}
DATAF = {'accept': '.csv,.xlsx,.json,.yaml,.yml,.xml', 'multiple': False}

# ---------------- PDF ----------------
def rd(f):
    try:
        r = PdfReader(io.BytesIO(f.read()))
        if r.is_encrypted and not r.decrypt(''): raise Bad('This PDF is password protected. Unlock it first.')
        return r
    except Bad: raise
    except Exception: raise Bad('Could not read this PDF.')
def wr(w): b = io.BytesIO(); w.write(b); return b.getvalue()
def pages(spec, n):
    if not (spec or '').strip(): return list(range(n))
    out = []
    for part in spec.replace(' ', '').split(','):
        m = re.fullmatch(r'(\d*)-(\d*)|(\d+)', part)
        if not m: raise Bad('Write pages like 1-3,5,8-')
        a, b = (int(m.group(3)),) * 2 if m.group(3) else (int(m.group(1) or 1), int(m.group(2) or n))
        if a < 1 or b > n or a > b: raise Bad(f'Pages must be between 1 and {n}.')
        out += range(a - 1, b)
    return out
PG = F('pages', 'Pages (blank = all). Example: 1-3,5,8-')

@tool('pdf-merge', 'Merge PDF', 'PDF', 'Join several PDFs into one.', PDFN)
def pdf_merge(fs, fm):
    need(fs, 2, 'PDF'); w = PdfWriter()
    for f in fs:
        for p in rd(f).pages: w.add_page(p)
    return F_('merged.pdf', wr(w), 'application/pdf')
