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

@tool('pdf-extract', 'Split / extract pages', 'PDF', 'Keep only the pages you choose, or split into one PDF per page.', PDF1, [PG, F('mode', 'Output', 'select', 'One PDF with these pages', ['One PDF with these pages', 'Separate PDF per page (ZIP)'])])
def pdf_extract(fs, fm):
    r = rd(fs[0]); idx = pages(fm.get('pages'), len(r.pages))
    if fm.get('mode', '').startswith('Separate'):
        items = []
        for i in idx:
            w = PdfWriter(); w.add_page(r.pages[i]); items.append((f'page-{i + 1}.pdf', wr(w)))
        return zip_(items)
    w = PdfWriter()
    for i in idx: w.add_page(r.pages[i])
    return F_('extracted.pdf', wr(w), 'application/pdf')

@tool('pdf-remove', 'Remove pages', 'PDF', 'Delete pages from a PDF.', PDF1, [F('pages', 'Pages to remove. Example: 2,4-6')])
def pdf_remove(fs, fm):
    r = rd(fs[0]); gone = set(pages(fm.get('pages'), len(r.pages))); w = PdfWriter()
    for i, p in enumerate(r.pages):
        if i not in gone: w.add_page(p)
    if not len(w.pages): raise Bad('That would remove every page.')
    return F_('trimmed.pdf', wr(w), 'application/pdf')

@tool('pdf-rotate', 'Rotate PDF pages', 'PDF', 'Rotate all or some pages.', PDF1, [F('angle', 'Rotate clockwise', 'select', '90', ['90', '180', '270']), PG])
def pdf_rotate(fs, fm):
    r = rd(fs[0]); sel = set(pages(fm.get('pages'), len(r.pages))); a = int(fm.get('angle', 90)); w = PdfWriter()
    if a not in (90, 180, 270): raise Bad('Angle must be 90, 180 or 270.')
    for i, p in enumerate(r.pages):
        w.add_page(p)
        if i in sel: w.pages[-1].rotate(a)
    return F_('rotated.pdf', wr(w), 'application/pdf')

@tool('pdf-compress', 'Compress PDF', 'PDF', 'Make a PDF smaller (images are recompressed).', PDF1, [F('quality', 'Image quality 20-90', 'number', '60')], heavy=True)
def pdf_compress(fs, fm):
    r = rd(fs[0]); w = PdfWriter(clone_from=r); q = int(num(fm.get('quality'), 60, 20, 90))
    for p in w.pages:
        try:
            for im in p.images: im.replace(im.image, quality=q)
        except Exception: pass
        p.compress_content_streams()
    w.compress_identical_objects(remove_identicals=True, remove_orphans=True)
    return F_('compressed.pdf', wr(w), 'application/pdf')

@tool('pdf-lock', 'Lock PDF with password', 'PDF', 'Encrypt a PDF (AES-256).', PDF1, [F('password', 'Password', 'password')])
def pdf_lock(fs, fm):
    pw = fm.get('password', '')
    if not pw: raise Bad('Enter a password.')
    w = PdfWriter(clone_from=rd(fs[0])); w.encrypt(pw, algorithm='AES-256')
    return F_('locked.pdf', wr(w), 'application/pdf')

@tool('pdf-unlock', 'Unlock PDF', 'PDF', 'Remove the password from a PDF you own (you must know the password).', PDF1, [F('password', 'Current password', 'password')])
def pdf_unlock(fs, fm):
    try:
        r = PdfReader(io.BytesIO(fs[0].read()))
        if r.is_encrypted and not r.decrypt(fm.get('password', '')): raise Bad('Wrong password.')
        return F_('unlocked.pdf', wr(PdfWriter(clone_from=r)), 'application/pdf')
    except Bad: raise
    except Exception: raise Bad('Could not read this PDF.')

def overlay(page, draw):
    w, h = float(page.mediabox.width), float(page.mediabox.height); b = io.BytesIO()
    c = canvas.Canvas(b, pagesize=(w, h)); draw(c, w, h); c.save(); b.seek(0); page.merge_page(PdfReader(b).pages[0])

@tool('pdf-watermark', 'Add text watermark', 'PDF', 'Stamp diagonal text on every page (Latin letters only).', PDF1, [F('text', 'Watermark text', 'text', 'CONFIDENTIAL')])
def pdf_watermark(fs, fm):
    t = (fm.get('text') or '').strip()[:60].encode('latin-1', 'replace').decode('latin-1')
    if not t: raise Bad('Enter the watermark text.')
    w = PdfWriter(clone_from=rd(fs[0]))
    def draw(c, W, H):
        c.saveState(); c.setFillColorRGB(.5, .5, .5, alpha=.3); c.translate(W / 2, H / 2); c.rotate(45)
        c.setFont('Helvetica-Bold', min(W, H) / 9); c.drawCentredString(0, 0, t); c.restoreState()
    for p in w.pages: overlay(p, draw)
    return F_('watermarked.pdf', wr(w), 'application/pdf')

@tool('pdf-pagenumbers', 'Add page numbers', 'PDF', 'Number every page at the bottom.', PDF1)
def pdf_pagenumbers(fs, fm):
    w = PdfWriter(clone_from=rd(fs[0])); n = len(w.pages)
    for i, p in enumerate(w.pages): overlay(p, lambda c, W, H, i=i: (c.setFont('Helvetica', 10), c.drawCentredString(W / 2, 20, f'{i + 1} / {n}')))
    return F_('numbered.pdf', wr(w), 'application/pdf')

@tool('pdf-text', 'Extract text from PDF', 'PDF', 'Copy the text out of a PDF (scanned PDFs need OCR on images).', PDF1)
def pdf_text(fs, fm):
    r = rd(fs[0]); t = '\n\n'.join((p.extract_text() or '') for p in r.pages).strip()
    return T_(t or 'No selectable text found. If this is a scan, convert pages to images and use Image to text (OCR).')

@tool('pdf-to-images', 'PDF to images', 'PDF', 'Turn pages into PNG images (max 30 pages).', PDF1, [PG, F('dpi', 'Quality (DPI) 72-200', 'number', '110')], heavy=True)
def pdf_to_images(fs, fm):
    raw = fs[0].read(); r = rd(type('R', (), {'read': lambda self: raw})())
    idx = pages(fm.get('pages'), len(r.pages))[:30]; dpi = int(num(fm.get('dpi'), 110, 72, 200)); items = []
    for i in idx:
        im = convert_from_bytes(raw, dpi=dpi, first_page=i + 1, last_page=i + 1)[0]; b = io.BytesIO(); im.save(b, 'PNG'); items.append((f'page-{i + 1}.png', b.getvalue()))
    return pack(items)

# ---------------- Images ----------------
EXT = {'PNG': 'png', 'JPEG': 'jpg', 'WEBP': 'webp', 'GIF': 'gif', 'BMP': 'bmp', 'TIFF': 'tiff'}
def im(f):
    try:
        i = Image.open(f); fmt = i.format; i.load(); i = ImageOps.exif_transpose(i); i.fmt = fmt if fmt in EXT else 'PNG'; return i
    except Exception: raise Bad('Could not read this image.')
def enc(i, fmt, q=85):
    if fmt in ('JPEG', 'BMP') and i.mode not in ('RGB', 'L'):
        bg = Image.new('RGB', i.size, 'white'); i = i.convert('RGBA'); bg.paste(i, mask=i.split()[3]); i = bg
    if fmt == 'GIF' and i.mode not in ('P', 'L', 'RGB'): i = i.convert('RGBA')
    b = io.BytesIO(); kw = {'quality': q, 'optimize': True} if fmt in ('JPEG', 'WEBP') else {'optimize': True} if fmt == 'PNG' else {}
    i.save(b, fmt, **kw); return b.getvalue()
def per_image(fs, make):
    need(fs, 1, 'image'); return pack([make(im(f), stem(f.filename)) for f in fs[:20]])
MIME = {'png': 'image/png', 'jpg': 'image/jpeg', 'webp': 'image/webp', 'gif': 'image/gif'}
FMTS = ['PNG', 'JPEG', 'WEBP', 'GIF', 'BMP', 'TIFF']

@tool('img-convert', 'Convert image format', 'Images', 'JPG, PNG, WEBP, GIF, BMP, TIFF and iPhone HEIC.', IMG, [F('format', 'Convert to', 'select', 'PNG', FMTS), F('quality', 'Quality 30-100', 'number', '90')])
def img_convert(fs, fm):
    f = fm.get('format', 'PNG'); f = f if f in FMTS else 'PNG'; q = int(num(fm.get('quality'), 90, 30, 100))
    return per_image(fs, lambda i, s: (f'{s}.{EXT[f]}', enc(i, f, q)))

@tool('img-compress', 'Compress image', 'Images', 'Shrink file size of photos.', IMG, [F('quality', 'Quality 20-90', 'number', '65'), F('maxw', 'Max width in px (0 = keep)', 'number', '1920')])
def img_compress(fs, fm):
    q = int(num(fm.get('quality'), 65, 20, 90)); mw = int(num(fm.get('maxw'), 1920, 0, 8000))
    def mk(i, s):
        f = 'WEBP' if i.fmt == 'WEBP' else 'PNG' if i.fmt == 'PNG' else 'JPEG'
        if mw and i.width > mw: i = i.resize((mw, round(i.height * mw / i.width)), Image.LANCZOS)
        return (f'{s}-small.{EXT[f]}', enc(i, f, q))
    return per_image(fs, mk)

@tool('img-resize', 'Resize image', 'Images', 'Change width and height or scale by percent.', IMG, [F('width', 'Width px (0 = auto)', 'number', '0'), F('height', 'Height px (0 = auto)', 'number', '0'), F('percent', 'Or scale % (0 = off)', 'number', '50')])
def img_resize(fs, fm):
    W, H, P = int(num(fm.get('width'), 0, 0, 10000)), int(num(fm.get('height'), 0, 0, 10000)), num(fm.get('percent'), 0, 0, 500)
    def mk(i, s):
        if P: w, h = round(i.width * P / 100), round(i.height * P / 100)
        elif W and H: w, h = W, H
        elif W: w, h = W, round(i.height * W / i.width)
        elif H: w, h = round(i.width * H / i.height), H
        else: raise Bad('Enter a width, a height or a percent.')
        return (f'{s}-resized.{EXT[i.fmt]}', enc(i.resize((max(w, 1), max(h, 1)), Image.LANCZOS), i.fmt))
    return per_image(fs, mk)

@tool('img-crop', 'Crop image', 'Images', 'Cut pixels off each edge.', IMG, [F('l', 'Left px', 'number', '0'), F('t', 'Top px', 'number', '0'), F('r', 'Right px', 'number', '0'), F('b', 'Bottom px', 'number', '0')])
def img_crop(fs, fm):
    l, t, r, b = (int(num(fm.get(k), 0, 0, 20000)) for k in 'ltrb')
    def mk(i, s):
        if l + r >= i.width or t + b >= i.height: raise Bad('Crop is larger than the image.')
        return (f'{s}-cropped.{EXT[i.fmt]}', enc(i.crop((l, t, i.width - r, i.height - b)), i.fmt))
    return per_image(fs, mk)

@tool('img-rotate', 'Rotate / flip image', 'Images', 'Rotate or mirror images.', IMG, [F('mode', 'Action', 'select', 'Rotate 90 right', ['Rotate 90 right', 'Rotate 90 left', 'Rotate 180', 'Flip horizontal', 'Flip vertical'])])
def img_rotate(fs, fm):
    m = fm.get('mode', ''); ops = {'Rotate 90 right': lambda i: i.rotate(-90, expand=True), 'Rotate 90 left': lambda i: i.rotate(90, expand=True), 'Rotate 180': lambda i: i.rotate(180), 'Flip horizontal': ImageOps.mirror, 'Flip vertical': ImageOps.flip}
    if m not in ops: raise Bad('Choose an action.')
    return per_image(fs, lambda i, s: (f'{s}-edited.{EXT[i.fmt]}', enc(ops[m](i), i.fmt)))

@tool('img-strip-exif', 'Remove photo metadata (EXIF)', 'Images', 'Strip location and camera data before sharing photos.', IMG)
def img_strip(fs, fm):
    def mk(i, s): i.info = {}; return (f'{s}-clean.{EXT[i.fmt]}', enc(i, i.fmt, 95))
    return per_image(fs, mk)

@tool('img-ocr', 'Image to text (OCR)', 'Images', 'Read text from a photo or screenshot.', {'accept': 'image/*,.heic', 'multiple': False}, [F('lang', 'Language', 'select', 'English', ['English', 'Hindi', 'English + Hindi'])], heavy=True, rate=8)
def img_ocr(fs, fm):
    lang = {'English': 'eng', 'Hindi': 'hin', 'English + Hindi': 'eng+hin'}.get(fm.get('lang'), 'eng')
    try: t = pytesseract.image_to_string(im(fs[0]).convert('RGB'), lang=lang, timeout=90)
    except (pytesseract.TesseractNotFoundError, RuntimeError, pytesseract.TesseractError): raise Bad('OCR is not available right now.', 503)
    return T_(t.strip() or 'No text found.')

@tool('img-favicon', 'Make favicon (ICO)', 'Images', 'Turn a square logo into favicon.ico with many sizes.', {'accept': 'image/*', 'multiple': False})
def img_favicon(fs, fm):
    i = im(fs[0]).convert('RGBA'); b = io.BytesIO(); i.save(b, 'ICO', sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]); return F_('favicon.ico', b.getvalue(), 'image/x-icon')

@tool('img-gif', 'Images to animated GIF', 'Images', 'Combine pictures into a GIF (max 60 frames).', IMG, [F('ms', 'Milliseconds per frame', 'number', '500')])
def img_gif(fs, fm):
    need(fs, 2, 'image'); fr = [im(f).convert('RGB') for f in fs[:60]]; fr = [fr[0]] + [x.resize(fr[0].size) for x in fr[1:]]
    b = io.BytesIO(); fr[0].save(b, 'GIF', save_all=True, append_images=fr[1:], duration=int(num(fm.get('ms'), 500, 50, 5000)), loop=0); return F_('animation.gif', b.getvalue(), 'image/gif')

@tool('img-palette', 'Color palette from image', 'Images', 'Find the main colors of a picture.', {'accept': 'image/*', 'multiple': False})
def img_palette(fs, fm):
    q = im(fs[0]).convert('RGB').resize((120, 120)).quantize(6); pal = q.getpalette()[:18]; cnt = sorted(q.getcolors(), reverse=True); tot = sum(c for c, _ in cnt)
    return T_('\n'.join(f'#{pal[k*3]:02X}{pal[k*3+1]:02X}{pal[k*3+2]:02X}  {c * 100 // tot}%' for c, k in cnt))

@tool('img-to-pdf', 'Images to PDF', 'PDF', 'Put photos into one PDF.', IMG)
def img_to_pdf(fs, fm):
    need(fs, 1, 'image'); L = [im(f).convert('RGB') for f in fs[:40]]; b = io.BytesIO(); L[0].save(b, 'PDF', save_all=True, append_images=L[1:], resolution=100); return F_('images.pdf', b.getvalue(), 'application/pdf')

# ---------------- Video & audio ----------------
AVX = {'mp4', 'mkv', 'mov', 'webm', 'avi', 'flv', 'wmv', '3gp', 'mp3', 'wav', 'aac', 'm4a', 'ogg', 'flac', 'opus', 'wma', 'gif'}
def ff(f, args, ext, pre=(), timeout=240):
    src_ext = (f.filename or '').rsplit('.', 1)[-1].lower()
    if src_ext not in AVX: raise Bad('This audio/video type is not supported.')
    d = tempfile.mkdtemp()
    try:
        src, dst = os.path.join(d, 'in.' + src_ext), os.path.join(d, 'out.' + ext); f.save(src)
        try: p = subprocess.run(['ffmpeg', '-y', '-hide_banner', '-loglevel', 'error', '-nostdin', *pre, '-i', src, *args, dst], capture_output=True, timeout=timeout)
        except subprocess.TimeoutExpired: raise Bad('That took too long. Try a shorter or smaller file.')
        except FileNotFoundError: raise Bad('Media tools are not available right now.', 503)
        if p.returncode or not os.path.exists(dst): raise Bad('Could not process this file. Is it a valid audio/video file?')
        return open(dst, 'rb').read()
    finally: shutil.rmtree(d, ignore_errors=True)
VID = ['-c:v', 'libx264', '-preset', 'veryfast', '-crf', '27', '-pix_fmt', 'yuv420p', '-c:a', 'aac', '-b:a', '128k']
CODECS = {'mp4': [*VID, '-movflags', '+faststart'], 'mkv': VID, 'webm': ['-c:v', 'libvpx', '-b:v', '1M', '-c:a', 'libvorbis'],
          'mp3': ['-vn', '-c:a', 'libmp3lame', '-q:a', '2'], 'wav': ['-vn'], 'm4a': ['-vn', '-c:a', 'aac', '-b:a', '192k'], 'ogg': ['-vn', '-c:a', 'libvorbis', '-q:a', '5'], 'flac': ['-vn']}
MIMES = {'mp4': 'video/mp4', 'mkv': 'video/x-matroska', 'webm': 'video/webm', 'mp3': 'audio/mpeg', 'wav': 'audio/wav', 'm4a': 'audio/mp4', 'ogg': 'audio/ogg', 'flac': 'audio/flac'}

@tool('av-convert', 'Convert video / audio', 'Video & Audio', 'MP4, MKV, WEBM, MP3, WAV, M4A, OGG, FLAC.', AV, [F('format', 'Convert to', 'select', 'mp4', list(CODECS))], heavy=True, rate=6)
def av_convert(fs, fm):
    e = fm.get('format', 'mp4'); e = e if e in CODECS else 'mp4'
    return F_(f'{stem(fs[0].filename)}.{e}', ff(fs[0], CODECS[e], e), MIMES[e])

@tool('av-extract-audio', 'Extract audio from video', 'Video & Audio', 'Get an MP3 from any video.', AV, heavy=True, rate=6)
def av_extract(fs, fm): return F_(f'{stem(fs[0].filename)}.mp3', ff(fs[0], CODECS['mp3'], 'mp3'), 'audio/mpeg')

@tool('video-compress', 'Compress video', 'Video & Audio', 'Smaller MP4, up to 720p.', AV, [F('crf', 'Compression 24 (better) to 35 (smaller)', 'number', '30')], heavy=True, rate=6)
def video_compress(fs, fm):
    c = str(int(num(fm.get('crf'), 30, 24, 35)))
    return F_(f'{stem(fs[0].filename)}-small.mp4', ff(fs[0], ['-vf', "scale=-2:'min(720,ih)'", '-c:v', 'libx264', '-preset', 'veryfast', '-crf', c, '-pix_fmt', 'yuv420p', '-c:a', 'aac', '-b:a', '96k', '-movflags', '+faststart'], 'mp4'), 'video/mp4')

@tool('video-gif', 'Video to GIF', 'Video & Audio', 'Make a GIF from the first seconds of a video.', AV, [F('start', 'Start second', 'number', '0'), F('dur', 'Length in seconds (max 15)', 'number', '6'), F('fps', 'Frames per second', 'number', '10')], heavy=True, rate=6)
def video_gif(fs, fm):
    fps = int(num(fm.get('fps'), 10, 3, 20))
    return F_(f'{stem(fs[0].filename)}.gif', ff(fs[0], ['-t', str(num(fm.get('dur'), 6, 1, 15)), '-vf', f'fps={fps},scale=min(480\\,iw):-1:flags=lanczos', '-loop', '0'], 'gif', pre=['-ss', str(num(fm.get('start'), 0, 0, 36000))]), 'image/gif')

@tool('av-trim', 'Trim / cut video or audio', 'Video & Audio', 'Keep only a part of a file.', AV, [F('start', 'Start second', 'number', '0'), F('dur', 'Length in seconds', 'number', '10')], heavy=True, rate=6)
def av_trim(fs, fm):
    e = (fs[0].filename or '').rsplit('.', 1)[-1].lower()
    if e not in CODECS: e = 'mp4'
    return F_(f'{stem(fs[0].filename)}-trim.{e}', ff(fs[0], ['-t', str(num(fm.get('dur'), 10, 0.5, 3600)), *CODECS[e]], e, pre=['-ss', str(num(fm.get('start'), 0, 0, 36000))]), MIMES[e])

# ---------------- Data & docs ----------------
def load_data(f):
    ext = (f.filename or '').rsplit('.', 1)[-1].lower(); raw = f.read()
    try:
        if ext == 'csv': return list(csv.DictReader(io.StringIO(raw.decode('utf-8-sig'))))
        if ext in ('xlsx', 'xlsm'):
            rows = list(load_workbook(io.BytesIO(raw), read_only=True, data_only=True).worksheets[0].iter_rows(values_only=True))
            head = [str(c) if c is not None else f'col{i + 1}' for i, c in enumerate(rows[0])]; return [dict(zip(head, r)) for r in rows[1:]]
        if ext == 'json': return json.loads(raw)
        if ext in ('yaml', 'yml'): return yaml.safe_load(raw)
        if ext == 'xml': return xmltodict.parse(raw)
    except Exception: raise Bad('Could not read this file. Check that it is valid.')
    raise Bad('Use a .csv, .xlsx, .json, .yaml or .xml file.')
def rows_of(d):
    if isinstance(d, dict):
        for v in d.values():
            if isinstance(v, list) and v and all(isinstance(x, dict) for x in v): return v
        return [d]
    return [x if isinstance(x, dict) else {'value': x} for x in d] if isinstance(d, list) else [{'value': d}]

@tool('data-convert', 'CSV / Excel / JSON / YAML / XML', 'Data & Docs', 'Convert between data formats.', DATAF, [F('to', 'Convert to', 'select', 'json', ['json', 'csv', 'xlsx', 'yaml', 'xml'])])
def data_convert(fs, fm):
    d = load_data(fs[0]); to = fm.get('to', 'json'); s = stem(fs[0].filename); clean = json.loads(json.dumps(d, default=str))
    if to == 'json': return F_(f'{s}.json', json.dumps(clean, indent=2, ensure_ascii=False).encode(), 'application/json')
    if to == 'yaml': return F_(f'{s}.yaml', yaml.safe_dump(clean, allow_unicode=True, sort_keys=False).encode(), 'text/yaml')
    if to == 'xml':
        try: return F_(f'{s}.xml', xmltodict.unparse({'root': {'item': clean}} if isinstance(clean, list) else clean if len(clean) == 1 else {'root': clean}, pretty=True).encode(), 'application/xml')
        except Exception: raise Bad('This data cannot be written as XML.')
    rows = rows_of(clean); cols = list(dict.fromkeys(k for r in rows for k in r))
    flat = lambda v: json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v
    if to == 'csv':
        b = io.StringIO(); w = csv.writer(b); w.writerow(cols); [w.writerow([flat(r.get(c)) for c in cols]) for r in rows]; return F_(f'{s}.csv', b.getvalue().encode('utf-8-sig'), 'text/csv')
    wb = Workbook(); ws = wb.active; ws.append(cols); [ws.append([flat(r.get(c)) for c in cols]) for r in rows]; b = io.BytesIO(); wb.save(b)
    return F_(f'{s}.xlsx', b.getvalue(), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@tool('md-to-html', 'Markdown to HTML', 'Data & Docs', 'Get the HTML code for Markdown text.', None, [F('text', 'Markdown', 'textarea')])
def md_html(fs, fm): return T_(markdown.markdown(fm.get('text', '')[:200000], extensions=['tables', 'fenced_code']))

@tool('text-diff', 'Compare two texts', 'Data & Docs', 'See what changed between two versions.', None, [F('a', 'Original', 'textarea'), F('b', 'Changed', 'textarea')])
def text_diff(fs, fm):
    d = '\n'.join(difflib.unified_diff(fm.get('a', '').splitlines(), fm.get('b', '').splitlines(), 'original', 'changed', lineterm=''))
    return T_(d or 'The two texts are identical.')

@tool('zip-create', 'Make a ZIP file', 'Data & Docs', 'Put several files in one ZIP.', {'accept': '*/*', 'multiple': True})
def zip_create(fs, fm):
    need(fs, 1); items = []; seen = set()
    for f in fs[:50]:
        n = re.sub(r'[^\w. -]', '_', os.path.basename(f.filename or 'file'))[:100] or 'file'
        while n in seen: n = '_' + n
        seen.add(n); items.append((n, f.read()))
    return zip_(items)

@tool('hash', 'Hash generator (MD5, SHA, CRC32)', 'Developer', 'Checksums for any text.', None, [F('text', 'Text', 'textarea')])
def hash_(fs, fm):
    b = fm.get('text', '').encode()
    return T_('\n'.join([f'{a.upper()}: {hashlib.new(a, b).hexdigest()}' for a in ('md5', 'sha1', 'sha256', 'sha512')] + [f'CRC32: {zlib.crc32(b) & 0xffffffff:08x}']))

@tool('qr', 'QR code generator', 'Developer', 'Make a QR code image for text or a link.', None, [F('text', 'Text or link', 'textarea'), F('size', 'Size 4-20', 'number', '10')])
def qr_(fs, fm):
    t = fm.get('text', '').strip()
    if not t or len(t) > 2000: raise Bad('Enter 1 to 2000 characters.')
    q = qrcode.QRCode(box_size=int(num(fm.get('size'), 10, 4, 20)), border=3); q.add_data(t); q.make(fit=True); b = io.BytesIO(); q.make_image().save(b, 'PNG'); return F_('qr.png', b.getvalue(), 'image/png')

@tool('barcode', 'Barcode generator (Code 128)', 'Developer', 'Letters and numbers to a barcode image.', None, [F('text', 'Text (ASCII only)')])
def barcode_(fs, fm):
    t = fm.get('text', '').strip()
    if not t or len(t) > 60 or not t.isascii(): raise Bad('Use 1 to 60 ASCII characters.')
    b = io.BytesIO(); barcode.get('code128', t, writer=ImageWriter()).write(b); return F_('barcode.png', b.getvalue(), 'image/png')

# ---------------- API routes ----------------
@app.errorhandler(Bad)
def bad(e): return jsonify(error=str(e)), e.code
@app.errorhandler(413)
def big(e): return jsonify(error=f'File too large. The limit is {MAX_MB} MB.'), 413
@app.errorhandler(404)
def nf(e): return (jsonify(error='Not found.'), 404) if request.path.startswith('/api') else (page('Not found', '<p>That page does not exist.</p>', 404))
@app.errorhandler(Exception)
def boom(e):
    if hasattr(e, 'code') and isinstance(e.code, int): return jsonify(error=getattr(e, 'description', 'Error')), e.code
    app.logger.exception('error'); return jsonify(error='Something went wrong. Please try again.'), 500

BY = {t['id']: t for t in TOOLS}
HEAVY = threading.BoundedSemaphore(2)
@app.get('/api/tools')
def tools(): return jsonify([{k: t[k] for k in ('id', 'name', 'cat', 'desc', 'files', 'fields')} for t in TOOLS])

@app.post('/api/run/<tid>')
def run(tid):
    t = BY.get(tid) or abort(404); limit('run', 90); limit('run:' + tid, t['rate'])
    fs = [f for f in request.files.getlist('files') if f and f.filename]
    if t['files'] and not fs: raise Bad('Please choose a file.')
    if len(fs) > 50: raise Bad('Too many files.')
    if t['heavy'] and not HEAVY.acquire(timeout=3): raise Bad('The server is busy. Try again in a moment.', 503)
    try: r = t['h'](fs, request.form)
    finally:
        if t['heavy']: HEAVY.release()
    if r[0] == 'text': return jsonify(text=r[1])
    return send_file(io.BytesIO(r[2]), mimetype=r[3], as_attachment=True, download_name=r[1])

# ---------------- Pages ----------------
def page(title, body, code=200):
    return Response(f'<!doctype html><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1"><meta name=robots content=noindex><title>{html.escape(title)}</title><link rel=stylesheet href=/static/style.css><main class=narrow><h1>{html.escape(title)}</h1>{body}<p><a href=/>Back to SourceLab Toolbox</a></p></main>', code, mimetype='text/html')

@app.get('/')
def index(): return app.send_static_file('index.html')
@app.get('/healthz')
def healthz(): return 'ok'
@app.get('/robots.txt')
def robots(): return Response('User-agent: *\nDisallow: /f/\nDisallow: /s/\nDisallow: /api/\n', mimetype='text/plain')

@app.after_request
def headers(r):
    r.headers.setdefault('X-Content-Type-Options', 'nosniff'); r.headers.setdefault('X-Frame-Options', 'DENY'); r.headers.setdefault('Referrer-Policy', 'no-referrer')
    r.headers.setdefault('Strict-Transport-Security', 'max-age=63072000; includeSubDomains')
    r.headers.setdefault('Content-Security-Policy', "default-src 'self'; img-src 'self' data: blob:; style-src 'self' 'unsafe-inline'; connect-src 'self' https:; frame-ancestors 'none'; base-uri 'none'; form-action 'self'")
    return r

# ---------------- Link shortener ----------------
SB, SK, SG = (os.environ.get(k) for k in ('LINKS_SB_URL', 'LINKS_SB_KEY', 'LINKS_GATE')); MEM = {}
def rpc(fn, **a):
    r = requests.post(f'{SB}/rest/v1/rpc/{fn}', json={'p_gate': SG, **a}, headers={'apikey': SK}, timeout=8); r.raise_for_status(); return r.json()
def link_new(code, url):
    if SB: return rpc('link_create', p_code=code, p_url=url, p_ip=ip_hash())
    if code in MEM: return 'exists'
    MEM[code] = [url, 0]; return 'ok'
def link_get(code):
    if SB: return rpc('link_get', p_code=code)
    v = MEM.get(code); return v[0] if v and v[1] < 3 else None
def link_flag(code):
    if SB: return rpc('link_report', p_code=code)
    if code in MEM: MEM[code][1] += 1
NOSHORT = {'bit.ly', 't.co', 'tinyurl.com', 'goo.gl', 'is.gd', 'cutt.ly', 'rebrand.ly', 'shorturl.at', 'ow.ly', 'buff.ly', 't.me', 'rb.gy', 'tiny.cc'}
def clean_url(u):
    u = (u or '').strip(); p = urlparse(u)
    if len(u) > 2000 or p.scheme not in ('http', 'https') or not p.hostname: raise Bad('Enter a full link starting with http:// or https://')
    h = p.hostname.lower().rstrip('.')
    if p.username or p.password: raise Bad('Links with a username or password are not allowed.')
    if '.' not in h or h in NOSHORT or h.endswith(('.local', '.internal', '.onion', '.localhost')) or h == request.host.split(':')[0]: raise Bad('This link cannot be shortened.')
    try:
        if not ipaddress.ip_address(h).is_global: raise Bad('This link cannot be shortened.')
    except ValueError: pass
    return u
@app.post('/api/shorten')
def shorten():
    limit('short', 10, 3600); u = clean_url((request.get_json(silent=True) or {}).get('url'))
    for _ in range(6):
        code = ''.join(secrets.choice('0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ') for _ in range(6)); r = link_new(code, u)
        if r == 'ok': return jsonify(url=f'{request.host_url}s/{code}', code=code)
        if r == 'rate': raise Bad('Daily link limit reached. Try again tomorrow.', 429)
    raise Bad('Could not create a link. Try again.', 503)
@app.get('/s/<code>')
def go(code):
    if not re.fullmatch(r'[0-9A-Za-z]{6}', code): abort(404)
    u = link_get(code)
    if not u: return page('Link not found', '<p>This link does not exist, expired, or was removed after reports.</p>', 404)
    e = html.escape(u)
    return page('You are leaving SourceLab Toolbox', f'<p>This short link points to:</p><p class=key>{e}</p><p>Continue only if you trust this address. Never enter passwords or payment details on a site you do not recognise.</p><p><a class=btn rel="nofollow noopener noreferrer" href="{e}">Continue</a></p><form method=post action="/s/{code}/report"><button class=alt>Report this link</button></form>')
@app.post('/s/<code>/report')
def report_link(code):
    limit('report', 10, 3600)
    if re.fullmatch(r'[0-9A-Za-z]{6}', code): link_flag(code)
    return page('Thank you', '<p>Your report was recorded. Links with several reports are disabled automatically.</p>')

# ---------------- 24-hour file hosting ----------------
BLOCK = {'exe', 'msi', 'bat', 'cmd', 'com', 'scr', 'js', 'mjs', 'vbs', 'vbe', 'ps1', 'apk', 'jar', 'html', 'htm', 'xhtml', 'svg', 'php', 'sh', 'dll', 'lnk', 'reg', 'hta', 'dmg', 'pkg', 'deb', 'iso'}
TTL = 24 * 3600
def sweep():
    now = time.time()
    for fn in os.listdir(FILES):
        if fn.endswith('.json'):
            p = os.path.join(FILES, fn)
            try: exp = json.load(open(p))['exp']
            except Exception: exp = 0
            if exp < now:
                for e in ('.json', '.bin'):
                    try: os.remove(p[:-5] + e)
                    except OSError: pass
def sweeper():
    while True:
        time.sleep(600)
        try: sweep()
        except Exception: pass
threading.Thread(target=sweeper, daemon=True).start()
def meta(i):
    if not re.fullmatch(r'[A-Za-z0-9_-]{10,40}', i): abort(404)
    try: m = json.load(open(os.path.join(FILES, i + '.json')))
    except Exception: abort(404)
    if m['exp'] < time.time(): abort(404)
    return m
@app.post('/api/host')
def host():
    limit('host', 10, 3600); f = request.files.get('files')
    if not f or not f.filename: raise Bad('Please choose a file.')
    name = re.sub(r'[^\w. -]', '_', os.path.basename(f.filename))[:100] or 'file'
    if any(p.lower() in BLOCK for p in name.split('.')[1:]): raise Bad('This file type is not allowed (programs, scripts, HTML and SVG are blocked).')
    sweep()
    if sum(os.path.getsize(os.path.join(FILES, x)) for x in os.listdir(FILES) if x.endswith('.bin')) > 400 * 1024 * 1024: raise Bad('Storage is full right now. Try again later.', 503)
    i, tok = secrets.token_urlsafe(12), secrets.token_urlsafe(12); f.save(os.path.join(FILES, i + '.bin'))
    json.dump({'name': name, 'exp': time.time() + TTL, 'tok': tok, 'size': os.path.getsize(os.path.join(FILES, i + '.bin'))}, open(os.path.join(FILES, i + '.json'), 'w'))
    return jsonify(url=f'{request.host_url}f/{i}', manage=f'{request.host_url}f/{i}/manage/{tok}', expires_in='24 hours')
@app.get('/f/<i>')
def fpage(i):
    m = meta(i)
    return page('Download file', f'<p class=key>{html.escape(m["name"])}</p><p>{m["size"] // 1024} KB. Expires in about {max(int((m["exp"] - time.time()) // 3600), 0)} hours. Files are uploaded by strangers and are not scanned. Open only files you trust.</p><p><a class=btn href="/f/{i}/get">Download</a></p><form method=post action="/f/{i}/report"><button class=alt>Report this file</button></form>')
@app.get('/f/<i>/get')
def fget(i):
    m = meta(i)
    r = send_file(os.path.join(FILES, i + '.bin'), mimetype='application/octet-stream', as_attachment=True, download_name=m['name'])
    r.headers['Content-Security-Policy'] = 'sandbox'; r.headers['X-Robots-Tag'] = 'noindex'; return r
@app.post('/f/<i>/report')
def freport(i):
    limit('report', 10, 3600); m = meta(i); m['reports'] = m.get('reports', 0) + 1
    if m['reports'] >= 3: m['exp'] = 0
    json.dump(m, open(os.path.join(FILES, i + '.json'), 'w')); app.logger.warning('file reported %s', i)
    return page('Thank you', '<p>Your report was recorded. Files with several reports are removed automatically.</p>')
@app.route('/f/<i>/manage/<tok>', methods=['GET', 'POST'])
def fmanage(i, tok):
    m = meta(i)
    if not secrets.compare_digest(m['tok'], tok): abort(404)
    if request.method == 'POST':
        for e in ('.json', '.bin'): os.remove(os.path.join(FILES, i + e))
        return page('Deleted', '<p>Your file was deleted.</p>')
    return page('Manage your file', f'<p class=key>{html.escape(m["name"])}</p><form method=post><button class=del>Delete now</button></form>')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 10000)))
