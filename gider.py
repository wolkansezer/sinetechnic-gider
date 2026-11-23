# -*- coding: utf-8 -*-
"""
SİNETECHNİC GİDER — Gider fişi toplama ve muhasebeciye gönderme programı
Sahibi: Volkan Sezer
"""
import base64
import ctypes
import json
import os
import re
import secrets
import shutil
import smtplib
import socket
import sqlite3
import sys
import threading
import zipfile
from datetime import date, datetime
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from urllib.parse import parse_qs, quote, urlparse
import urllib.error
import urllib.request

from PIL import Image, ImageOps, ImageEnhance
from PySide6.QtCore import (QDate, QLocale, QObject, QRunnable, Qt, QThreadPool, QTimer,
                            QUrl, Signal)
from PySide6.QtGui import (QAction, QColor, QDesktopServices, QIcon, QImage,
                           QPixmap, QTransform)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox,
                               QComboBox, QDateEdit, QDialog,
                               QDialogButtonBox, QDoubleSpinBox, QFileDialog,
                               QFormLayout, QFrame, QGridLayout, QHBoxLayout,
                               QHeaderView, QLabel, QLineEdit, QMainWindow,
                               QMessageBox, QPushButton, QScrollArea,
                               QSpinBox, QSplitter, QTableWidget,
                               QTableWidgetItem, QTextEdit, QVBoxLayout,
                               QWidget)

APP_NAME = "SİNETECHNİC GİDER"
APP_VERSION = "1.2.0"
OWNER = "Volkan Sezer"
WEBSITE = "https://www.sinetechnic.com"
EMAIL = "volkan.sezer@sinetechnic.com"
DISCORD = "https://discord.gg/nvqc84AHZf"

AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz",
         "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
KATEGORILER = ["Akaryakıt", "Yedek Parça / Malzeme", "Yemek", "Konaklama",
               "Ulaşım / Otoyol / Otopark", "Telefon / İnternet", "Kırtasiye",
               "Kargo", "Araç Bakım", "Diğer"]
ODEMELER = ["Kredi Kartı", "Banka Kartı", "Nakit", "Havale / EFT"]
BELGE_TURLERI = ["ÖKC Fişi", "e-Arşiv Fatura", "e-Fatura", "Kağıt Fatura",
                 "Gider Pusulası", "Diğer"]
IMG_EXT = (".jpg", ".jpeg", ".png", ".heic", ".webp", ".bmp")
ALL_EXT = IMG_EXT + (".pdf",)


# ----------------------------------------------------------------- yollar
def resource_path(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def data_root():
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    if not os.path.isdir(docs):
        docs = os.path.expanduser("~")
    root = os.path.join(docs, "SINETECHNIC GIDER")
    os.makedirs(os.path.join(root, "fisler"), exist_ok=True)
    os.makedirs(os.path.join(root, "paketler"), exist_ok=True)
    return root


ROOT = data_root()
DB_PATH = os.path.join(ROOT, "gider.db")
SETTINGS_PATH = os.path.join(ROOT, "ayarlar.json")


# ----------------------------------------------------------------- şifre koruma (Windows DPAPI)
class _BLOB(ctypes.Structure):
    _fields_ = [("cbData", ctypes.c_uint32), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _dpapi(data: bytes, protect: bool) -> bytes:
    if os.name != "nt":
        return data
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    buf = ctypes.create_string_buffer(data, len(data))
    blob_in = _BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    blob_out = _BLOB()
    fn = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    if protect:
        ok = fn(ctypes.byref(blob_in), "SINETECHNIC", None, None, None, 0, ctypes.byref(blob_out))
    else:
        ok = fn(ctypes.byref(blob_in), None, None, None, None, 0, ctypes.byref(blob_out))
    if not ok:
        raise OSError("DPAPI hatası")
    out = ctypes.string_at(blob_out.pbData, blob_out.cbData)
    kernel32.LocalFree(blob_out.pbData)
    return out


def sifrele(text: str) -> str:
    if not text:
        return ""
    try:
        return "dp:" + base64.b64encode(_dpapi(text.encode("utf-8"), True)).decode()
    except Exception:
        return "b64:" + base64.b64encode(text.encode("utf-8")).decode()


def coz(text: str) -> str:
    if not text:
        return ""
    try:
        if text.startswith("dp:"):
            return _dpapi(base64.b64decode(text[3:]), False).decode("utf-8")
        if text.startswith("b64:"):
            return base64.b64decode(text[4:]).decode("utf-8")
    except Exception:
        return ""
    return text


# ----------------------------------------------------------------- ayarlar
DEFAULT_SETTINGS = {
    "firma": "SİNETECHNİC",
    "muhasebeci_ad": "",
    "muhasebeci_eposta": "",
    "bana_kopya": True,
    "smtp_host": "smtp-mail.outlook.com",
    "smtp_port": 587,
    "smtp_guvenlik": "STARTTLS",
    "smtp_kullanici": "",
    "smtp_sifre": "",
    "telefon_port": 8765,
    "token": "",
    "otomatik_ocr": True,
    "tarama_modu": True,
    "github_repo": "wolkansezer/sinetechnic-gider-veri",
    "github_anahtar": "",
    "telefon_sayfa": "https://wolkansezer.github.io/sinetechnic-gider/",
    "bulut_aralik_sn": 60,
    "teslim_gunu": 14,
    "otomatik_gonder": True,
    "otomatik_gun": 10,
    "son_oto_deneme": "",
}


def load_settings():
    s = dict(DEFAULT_SETTINGS)
    if os.path.exists(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                s.update(json.load(f))
        except Exception:
            pass
    if not s.get("token"):
        s["token"] = secrets.token_urlsafe(8)
        save_settings(s)
    return s


def save_settings(s):
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)


# ----------------------------------------------------------------- veritabanı
class DB:
    def __init__(self, path):
        self.lock = threading.Lock()
        self.con = sqlite3.connect(path, check_same_thread=False)
        self.con.row_factory = sqlite3.Row
        self.con.execute("""CREATE TABLE IF NOT EXISTS fis(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tarih TEXT NOT NULL,
            firma TEXT DEFAULT '',
            vkn TEXT DEFAULT '',
            belge_turu TEXT DEFAULT 'ÖKC Fişi',
            belge_no TEXT DEFAULT '',
            kategori TEXT DEFAULT 'Diğer',
            tutar REAL DEFAULT 0,
            kdv_orani REAL DEFAULT 20,
            kdv REAL DEFAULT 0,
            odeme TEXT DEFAULT 'Kredi Kartı',
            plaka TEXT DEFAULT '',
            aciklama TEXT DEFAULT '',
            dosya TEXT DEFAULT '',
            kaynak TEXT DEFAULT 'PC',
            ocr_metin TEXT DEFAULT '',
            eklenme TEXT,
            gonderildi TEXT
        )""")
        cols = {r[1] for r in self.con.execute("PRAGMA table_info(fis)")}
        if "tarih_kesin" not in cols:
            self.con.execute("ALTER TABLE fis ADD COLUMN tarih_kesin INTEGER DEFAULT 0")
        if "ek_no" not in cols:
            self.con.execute("ALTER TABLE fis ADD COLUMN ek_no INTEGER")
        self.con.commit()

    def add(self, **kw):
        kw.setdefault("eklenme", datetime.now().isoformat(timespec="seconds"))
        cols = ",".join(kw.keys())
        q = ",".join("?" * len(kw))
        with self.lock:
            cur = self.con.execute(f"INSERT INTO fis({cols}) VALUES({q})", list(kw.values()))
            self.con.commit()
            return cur.lastrowid

    def update(self, fid, **kw):
        if not kw:
            return
        sets = ",".join(f"{k}=?" for k in kw)
        with self.lock:
            self.con.execute(f"UPDATE fis SET {sets} WHERE id=?", list(kw.values()) + [fid])
            self.con.commit()

    def delete(self, fid):
        with self.lock:
            self.con.execute("DELETE FROM fis WHERE id=?", (fid,))
            self.con.commit()

    def get(self, fid):
        with self.lock:
            return self.con.execute("SELECT * FROM fis WHERE id=?", (fid,)).fetchone()

    def month(self, ym):
        with self.lock:
            return self.con.execute(
                "SELECT * FROM fis WHERE substr(tarih,1,7)=? ORDER BY tarih, id", (ym,)).fetchall()

    def months(self):
        with self.lock:
            rows = self.con.execute(
                "SELECT DISTINCT substr(tarih,1,7) m FROM fis ORDER BY m DESC").fetchall()
        return [r["m"] for r in rows]

    def unsent_months(self):
        with self.lock:
            rows = self.con.execute(
                "SELECT substr(tarih,1,7) m, COUNT(*) n FROM fis WHERE gonderildi IS NULL GROUP BY m ORDER BY m").fetchall()
        return [(r["m"], r["n"]) for r in rows]

    def unsent_count(self, ym):
        with self.lock:
            return self.con.execute(
                "SELECT COUNT(*) FROM fis WHERE substr(tarih,1,7)=? AND gonderildi IS NULL",
                (ym,)).fetchone()[0]


# ----------------------------------------------------------------- yardımcılar
def tl(v):
    s = f"{v:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".") + " TL"


def teslim_tarihi(ym, gun=14):
    y, m = map(int, ym.split("-"))
    y2, m2 = (y + 1, 1) if m == 12 else (y, m + 1)
    return date(y2, m2, gun)


def onceki_ay(d=None):
    d = d or date.today()
    return (date(d.year - 1, 12, 1) if d.month == 1 else date(d.year, d.month - 1, 1)).strftime("%Y-%m")


def ay_adi(ym):
    y, m = ym.split("-")
    return f"{AYLAR[int(m) - 1]} {y}"


def guvenli_ad(s, n=30):
    tr = str.maketrans("çğıöşüÇĞİÖŞÜ ", "cgiosuCGIOSU_")
    s = (s or "").translate(tr)
    s = re.sub(r"[^A-Za-z0-9_\-]", "", s)
    return s[:n] or "FIS"


def yerel_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def _heif():
    try:
        from pillow_heif import register_heif_opener  # type: ignore
        register_heif_opener()
    except Exception:
        pass


def log_yaz(dosya, mesaj):
    try:
        with open(os.path.join(ROOT, dosya), "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {mesaj}\n")
    except Exception:
        pass


# ---------------- belge tarama (fişi fotoğraftan bul, düzleştir, taranmış görünüm ver)
def _sirala(pts):
    import numpy as np
    pts = pts.reshape(4, 2).astype("float32")
    s = pts.sum(1)
    d = np.diff(pts, axis=1).ravel()
    return np.array([pts[s.argmin()], pts[d.argmin()], pts[s.argmax()], pts[d.argmax()]], dtype="float32")


def belge_bul(img):
    """Fotoğraftaki en büyük açık renkli belgeyi (fişi) bulur. (köşeler, kesin_mi) veya None."""
    import cv2
    import numpy as np
    h, w = img.shape[:2]
    sc = 1000.0 / max(h, w)
    small = cv2.resize(img, (max(1, int(w * sc)), max(1, int(h * sc))), interpolation=cv2.INTER_AREA)
    gray = cv2.cvtColor(small, cv2.COLOR_BGR2GRAY)
    hsv = cv2.cvtColor(small, cv2.COLOR_BGR2HSV)
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    t1, _ = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    ust = blur[blur > t1]
    t2 = cv2.threshold(ust.reshape(-1, 1), 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[0] if ust.size > 100 else t1
    beyaz = cv2.inRange(hsv, (0, 0, 0), (180, 70, 255))
    adaylar = [cv2.bitwise_and(cv2.threshold(blur, t1, 255, cv2.THRESH_BINARY)[1], beyaz)]
    for t in (t2, max(t1, t2 - 15)):
        m = cv2.bitwise_and(cv2.threshold(blur, t, 255, cv2.THRESH_BINARY)[1], beyaz)
        m = cv2.morphologyEx(m, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_RECT, (25, 25)))
        m = cv2.morphologyEx(m, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (21, 21)))
        adaylar.append(m)
    edges = cv2.dilate(cv2.Canny(blur, 50, 150), cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5)), iterations=2)
    adaylar.append(edges)
    alan_top = small.shape[0] * small.shape[1]
    en_iyi = None
    for m in adaylar:
        cnts, _ = cv2.findContours(m, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        for c in sorted(cnts, key=cv2.contourArea, reverse=True)[:5]:
            a = cv2.contourArea(c)
            if a < 0.05 * alan_top:
                break
            if a > 0.97 * alan_top:
                continue
            ap = cv2.approxPolyDP(c, 0.02 * cv2.arcLength(c, True), True)
            dort = len(ap) == 4 and cv2.isContourConvex(ap)
            q = ap if dort else cv2.boxPoints(cv2.minAreaRect(c)).astype("int32")
            mm = np.zeros(gray.shape, np.uint8)
            cv2.drawContours(mm, [q.reshape(-1, 1, 2)], -1, 255, -1)
            parlak = cv2.mean(gray, mask=mm)[0]
            dolu = min(1.0, a / (cv2.contourArea(q.reshape(-1, 1, 2).astype("float32")) or 1))
            skor = a * (parlak / 255.0) ** 4 * dolu ** 2 * (1.2 if dort else 1.0)
            if en_iyi is None or skor > en_iyi[0]:
                en_iyi = (skor, q, dort)
    if en_iyi is None:
        return None
    return _sirala(en_iyi[1].astype("float32") / sc), en_iyi[2]


def _duzlestir(img, q):
    import cv2
    import numpy as np
    tl, tr, br, bl = q
    W = int(max(np.linalg.norm(br - bl), np.linalg.norm(tr - tl)))
    H = int(max(np.linalg.norm(tr - br), np.linalg.norm(tl - bl)))
    if W < 50 or H < 50:
        return img
    M = cv2.getPerspectiveTransform(q, np.array([[0, 0], [W - 1, 0], [W - 1, H - 1], [0, H - 1]], dtype="float32"))
    return cv2.warpPerspective(img, M, (W, H), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)


def _tarama_gorunumu(img):
    import cv2
    import numpy as np
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) if img.ndim == 3 else img
    k = max(15, (min(gray.shape) // 20) | 1)
    bg = cv2.medianBlur(cv2.dilate(gray, np.ones((7, 7), np.uint8)), k)
    norm = cv2.divide(gray, bg, scale=255)
    norm = cv2.normalize(norm, None, 0, 255, cv2.NORM_MINMAX)
    lut = np.array([255 * (i / 255.0) ** 2.2 for i in range(256)], np.uint8)
    norm = cv2.normalize(cv2.LUT(norm, lut), None, 0, 255, cv2.NORM_MINMAX)
    blur = cv2.GaussianBlur(norm, (0, 0), 1.0)
    return cv2.addWeighted(norm, 1.5, blur, -0.5, 0)


def belge_tara(pil_img):
    """PIL görüntüsünden taranmış belge (PIL, gri) üretir. (görüntü, fiş_bulundu_mu)"""
    import cv2
    import numpy as np
    img = cv2.cvtColor(np.array(pil_img.convert("RGB")), cv2.COLOR_RGB2BGR)
    bulundu = False
    try:
        r = belge_bul(img)
        if r:
            q, kesin = r
            img = _duzlestir(img, q)
            bulundu = True
            if not kesin:
                r2 = belge_bul(img)
                if r2:
                    a2 = cv2.contourArea(r2[0].reshape(-1, 1, 2))
                    if 0.45 < a2 / float(img.shape[0] * img.shape[1]) < 0.95:
                        img = _duzlestir(img, r2[0])
    except Exception as e:
        log_yaz("tarama_log.txt", f"belge bulunamadı: {e!r}")
    h, w = img.shape[:2]
    if w < 1200:
        f = 1200.0 / w
        img = cv2.resize(img, (int(w * f), int(h * f)), interpolation=cv2.INTER_CUBIC)
    if h > 4000 or w > 2400:
        f = min(4000.0 / h, 2400.0 / w)
        img = cv2.resize(img, (int(w * f), int(h * f)), interpolation=cv2.INTER_AREA)
    return Image.fromarray(_tarama_gorunumu(img)), bulundu


def gorsel_hazirla(src_bytes, tarama=True):
    """(kaydedilecek JPEG, orijinal JPEG) döndürür. Tarama açıksa fiş bulunup düzleştirilir."""
    _heif()
    im = ImageOps.exif_transpose(Image.open(BytesIO(src_bytes))).convert("RGB")
    im.thumbnail((3000, 3000))
    orj = BytesIO()
    im.save(orj, "JPEG", quality=85, optimize=True)
    if not tarama:
        return orj.getvalue(), orj.getvalue()
    try:
        sc, _ = belge_tara(im)
    except Exception as e:
        log_yaz("tarama_log.txt", f"tarama hatası: {e!r}")
        sc = ImageEnhance.Sharpness(ImageOps.autocontrast(ImageOps.grayscale(im), cutoff=1)).enhance(1.6)
    out = BytesIO()
    sc.convert("L").save(out, "JPEG", quality=88, optimize=True)
    return out.getvalue(), orj.getvalue()


def orj_yolu(rel):
    p = abs_path(rel)
    return os.path.join(os.path.dirname(p), "orj", os.path.basename(p))


def dosya_kaydet(data: bytes, ext: str, tarih: str, tarama=True):
    """Belgeyi fisler/YYYY-MM klasörüne kaydeder (orijinali fisler/YYYY-MM/orj altında), ROOT'a göreli yolu döndürür."""
    folder = os.path.join(ROOT, "fisler", tarih[:7])
    os.makedirs(folder, exist_ok=True)
    ext = ext.lower()
    orj = None
    if ext in IMG_EXT:
        try:
            data, orj = gorsel_hazirla(data, tarama)
            ext = ".jpg"
        except Exception as e:
            log_yaz("tarama_log.txt", f"görsel hazırlanamadı: {e!r}")
    name = datetime.now().strftime("%Y%m%d_%H%M%S_") + secrets.token_hex(3) + ext
    path = os.path.join(folder, name)
    with open(path, "wb") as f:
        f.write(data)
    if orj:
        os.makedirs(os.path.join(folder, "orj"), exist_ok=True)
        with open(os.path.join(folder, "orj", name), "wb") as f:
            f.write(orj)
    return os.path.relpath(path, ROOT)


def abs_path(rel):
    return os.path.join(ROOT, rel) if rel else ""


def dosya_tasi(rel, yeni_ym):
    """Belgeyi (ve orijinalini) başka ay klasörüne taşır, yeni göreli yolu döndürür."""
    src = abs_path(rel)
    dd = os.path.join(ROOT, "fisler", yeni_ym)
    os.makedirs(dd, exist_ok=True)
    dst = os.path.join(dd, os.path.basename(src))
    shutil.move(src, dst)
    o = orj_yolu(rel)
    if os.path.exists(o):
        os.makedirs(os.path.join(dd, "orj"), exist_ok=True)
        shutil.move(o, os.path.join(dd, "orj", os.path.basename(src)))
    return os.path.relpath(dst, ROOT)


# ---------------- OCR: gömülü RapidOCR (birincil) + Windows OCR (yedek)
_ocr_motor = None
_ocr_kilit = threading.Lock()


def _rapid_ocr(img):
    global _ocr_motor
    from rapidocr_onnxruntime import RapidOCR  # type: ignore
    import numpy as np
    with _ocr_kilit:
        if _ocr_motor is None:
            _ocr_motor = RapidOCR()
        res, _ = _ocr_motor(np.array(img.convert("RGB")))
    if not res:
        return ""
    kutular = []
    for box, txt, conf in res:
        ys = [p[1] for p in box]
        xs = [p[0] for p in box]
        kutular.append(((min(ys) + max(ys)) / 2, max(ys) - min(ys), min(xs), txt))
    kutular.sort()
    satirlar, cur = [], []
    for k in kutular:
        if cur and abs(k[0] - cur[-1][0]) > max(8, 0.55 * max(k[1], cur[-1][1])):
            satirlar.append(cur)
            cur = []
        cur.append(k)
    if cur:
        satirlar.append(cur)
    return "\n".join("  ".join(t[3] for t in sorted(s, key=lambda x: x[2])) for s in satirlar)


def _win_ocr(img):
    import asyncio
    import winocr  # type: ignore
    for lang in ("tr", "tr-TR", "en-US", "en"):
        try:
            res = asyncio.run(winocr.recognize_pil(img, lang))
            lines = [getattr(ln, "text", "") for ln in (getattr(res, "lines", None) or [])]
            text = "\n".join(lines) if lines else (getattr(res, "text", None) or "")
            if text:
                return text
        except Exception as e:
            log_yaz("ocr_log.txt", f"winocr {lang}: {e!r}")
    return ""


def ocr_metin(path):
    if not path or not path.lower().endswith((".jpg", ".jpeg", ".png", ".bmp", ".webp")):
        return ""
    img = Image.open(path).convert("RGB")
    for ad, fn in (("rapidocr", _rapid_ocr), ("winocr", _win_ocr)):
        try:
            text = fn(img)
            if text and len(text.strip()) > 10:
                return text
        except Exception as e:
            log_yaz("ocr_log.txt", f"{ad} hatası: {e!r}")
    return ""


TUTAR_RE = re.compile(r"\*?\s?(\d{1,3}(?:[.\s]\d{3})*,\d{2}|\d+,\d{2})")


def _tutar(s):
    try:
        return float(s.replace(" ", "").replace(".", "").replace(",", "."))
    except Exception:
        return None


def sayi_oku(s):
    s = (s or "").strip().replace(" ", "").replace("₺", "").replace("TL", "")
    if not s:
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


KATEGORI_KELIME = [
    ("Akaryakıt", r"MOTORIN|BENZIN|\bLPG\b|AKARYAKIT|OPET|SHELL|PETROL|TOTAL|AYTEMIZ|\bBP\b|DIZEL|EURODIESEL|KURSUNSUZ|POMPA"),
    ("Yemek", r"DONER|KEBAP|LOKANTA|RESTORAN|RESTAURANT|CAFE|KAFE|PIDE|BURGER|PIZZA|CORBA|AYRAN|SUAVYRAN|COLA|BAKLAVA|PASTANE|SIMIT|TOST|YEMEK|MENU"),
    ("Ulaşım / Otoyol / Otopark", r"OTOPARK|OTOYOL|KOPRU|HGS|OGS|TAKSI|ISPARK"),
    ("Konaklama", r"OTEL|HOTEL|PANSIYON|KONAKLAMA"),
    ("Kargo", r"KARGO|YURTICI|ARAS|MNG|SURAT|PTT"),
    ("Kırtasiye", r"KIRTASIYE|KAGIT|TONER|KARTUS"),
    ("Telefon / İnternet", r"TURKCELL|VODAFONE|TURK TELEKOM|SUPERONLINE|INTERNET"),
    ("Araç Bakım", r"LASTIK|OTO SERVIS|YAG DEGISIM|OTO YIKAMA|AKU"),
    ("Yedek Parça / Malzeme", r"ELEKTRIK|HIRDAVAT|YAPI MARKET|KOCTAS|BAUHAUS|LAMBA|KABLO|ELEKTRONIK"),
]


def fis_ayristir(text):
    """OCR metninden alanları tahmin eder."""
    out = {}
    if not text:
        return out
    tr_map = str.maketrans("İIĞÜŞÖÇığüşöç", "IIGUSOCiguscc")
    up = text.translate(tr_map).upper()
    lines = [l.strip() for l in up.splitlines() if l.strip()]

    for m in re.finditer(r"\b(\d{1,2})\s?[./-]\s?(\d{1,2})\s?[./-]\s?(\d{4}|\d{2})\b", up):
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        y = y + 2000 if y < 100 else y
        try:
            dt = date(y, mo, d)
        except ValueError:
            continue
        if date(2020, 1, 1) <= dt <= date.today():
            out["tarih"] = dt.isoformat()
            break

    toplam = kdv = None
    for i, ln in enumerate(lines):
        amounts = [a for a in (_tutar(x) for x in TUTAR_RE.findall(ln)) if a is not None]
        if not amounts and i + 1 < len(lines) and re.search(r"TOP", ln):
            amounts = [a for a in (_tutar(x) for x in TUTAR_RE.findall(lines[i + 1])) if a is not None]
        if not amounts:
            continue
        sade = ln.replace(" ", "")
        if ("TOPKDV" in sade or "TOPLAMKDV" in sade) and kdv is None:
            kdv = amounts[-1]
        elif ("TOPLAM" in sade or re.search(r"\bTOP\b", ln) or "GENELTOPLAM" in sade or "ODENECEK" in sade) and "KDV" not in sade:
            toplam = max(toplam or 0, amounts[-1])
    if toplam is None:
        alla = [a for a in (_tutar(x) for x in TUTAR_RE.findall(up)) if a]
        if alla:
            toplam = max(alla)
    if toplam:
        out["tutar"] = toplam
    if kdv and toplam and kdv < toplam:
        out["kdv"] = kdv
        oran = kdv / max(0.01, toplam - kdv) * 100
        out["kdv_orani"] = min((1.0, 10.0, 20.0), key=lambda o: abs(o - oran))

    m = re.search(r"F[I1]S\s*N[O0]\s*[:.]?\s*(\d+)", up)
    if m:
        out["belge_no"] = m.group(1)
    m = re.search(r"(?:VKN|V\.?\s?NO|VERGI\s*NO|\bV\.?\s?D\.?\b|\bVD\b)[^\d]{0,25}(\d{10,11})", up)
    if m:
        out["vkn"] = m.group(1)
    m = re.search(r"PLAKA[^A-Z0-9]{0,5}(\d{2}\s?[A-Z]{1,3}\s?\d{2,4})", up)
    if m:
        out["plaka"] = m.group(1).replace(" ", "")

    for kat, pat in KATEGORI_KELIME:
        if re.search(pat, up):
            out["kategori"] = kat
            if kat == "Akaryakıt":
                out.setdefault("kdv_orani", 20.0)
            break
    for ln in lines[:6]:
        if len(re.sub(r"[^A-Z]", "", ln)) >= 4 and not re.search(r"\d{1,2}[./-]\d{1,2}[./-]\d{2}", ln):
            out["firma"] = ln.title()[:60]
            break
    return out


# ----------------------------------------------------------------- telefon sunucusu
TELEFON_HTML = """<!doctype html><html lang="tr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>SİNETECHNİC GİDER</title><style>
*{box-sizing:border-box}body{margin:0;font-family:-apple-system,Segoe UI,sans-serif;background:#0e1626;color:#e8edf5}
header{padding:18px;background:#121c30;border-bottom:3px solid #ff7a1a}
h1{margin:0;font-size:19px;letter-spacing:.5px}h1 span{color:#ff7a1a}
main{padding:16px}label{display:block;font-size:13px;color:#9fb0c8;margin:14px 0 6px}
.big{display:flex;align-items:center;justify-content:center;width:100%;height:130px;border-radius:16px;
background:#ff7a1a;color:#fff;font-size:22px;font-weight:700;border:none}
input,select{width:100%;padding:13px;border-radius:10px;border:1px solid #2a3a58;background:#16223a;color:#fff;font-size:16px}
.row{display:flex;gap:10px}.row>div{flex:1}
#prev{display:flex;flex-wrap:wrap;gap:8px;margin-top:12px}#prev img{width:80px;height:100px;object-fit:cover;border-radius:8px}
button.send{margin-top:20px;width:100%;padding:16px;border-radius:12px;border:none;background:#2e8b57;color:#fff;font-size:18px;font-weight:700}
#msg{margin-top:16px;padding:12px;border-radius:10px;display:none;text-align:center;font-size:16px}
.ok{background:#1f5132}.err{background:#6b1f1f}
</style></head><body><header><h1>SİNETECHNİC <span>GİDER</span></h1></header><main>
<label class="big" for="f">📷 Fiş Fotoğrafı Çek</label>
<input id="f" type="file" accept="image/*,application/pdf" capture="environment" multiple style="display:none">
<div id="prev"></div>
<label>Fiş tarihi</label><input id="tarih" type="date">
<label>Tutar (boş bırakılırsa program okumaya çalışır)</label><input id="tutar" inputmode="decimal" placeholder="örn. 2500,00">
<div class="row"><div><label>Kategori</label><select id="kategori">%KATEGORI%</select></div>
<div><label>Ödeme</label><select id="odeme">%ODEME%</select></div></div>
<div class="row"><div><label>Plaka</label><input id="plaka" placeholder="34 ABC 123"></div>
<div><label>Not</label><input id="not" placeholder="isteğe bağlı"></div></div>
<button class="send" id="s">Programa Gönder</button><div id="msg"></div></main>
<script>
const f=document.getElementById('f'),prev=document.getElementById('prev'),msg=document.getElementById('msg');
let files=[];(function(){const d=new Date();document.getElementById('tarih').value=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');})();
f.onchange=()=>{for(const x of f.files){files.push(x);const i=document.createElement('img');
if(x.type.startsWith('image/')){i.src=URL.createObjectURL(x);}else{i.alt='PDF';}prev.appendChild(i);}f.value='';};
function show(t,c){msg.textContent=t;msg.className=c;msg.style.display='block';}
document.getElementById('s').onclick=async()=>{
 if(!files.length){show('Önce fiş fotoğrafı çekin','err');return;}
 const fd=new FormData();files.forEach((x,i)=>fd.append('dosya',x,x.name||('fis'+i+'.jpg')));
 ['tutar','kategori','odeme','plaka','not','tarih'].forEach(k=>fd.append(k,document.getElementById(k).value));
 show('Gönderiliyor…','ok');
 try{const r=await fetch('upload'+location.search,{method:'POST',body:fd});const j=await r.json();
  if(j.ok){show('✅ '+j.adet+' belge programa eklendi','ok');files=[];prev.innerHTML='';
   ['tutar','plaka','not'].forEach(k=>document.getElementById(k).value='');}
  else show('Hata: '+j.hata,'err');}catch(e){show('Bağlantı hatası: bilgisayar açık ve aynı Wi-Fi ağında mı?','err');}
};
</script></body></html>"""


class Bridge(QObject):
    yeni_fis = Signal(list)
    ocr_bitti = Signal(int)
    bulut_durum = Signal(str)


class TelefonSunucu:
    def __init__(self, settings, db, bridge):
        self.s, self.db, self.bridge = settings, db, bridge
        self.httpd = None
        self.hata = ""

    @property
    def url(self):
        return f"http://{yerel_ip()}:{self.s['telefon_port']}/?t={self.s['token']}"

    def start(self):
        outer = self
        kat = '<option value="">Otomatik (fişten)</option>' + "".join(f"<option>{k}</option>" for k in KATEGORILER)
        odm = "".join(f"<option>{k}</option>" for k in ODEMELER)
        page = TELEFON_HTML.replace("%KATEGORI%", kat).replace("%ODEME%", odm).encode("utf-8")

        class H(BaseHTTPRequestHandler):
            def log_message(self, *a):
                pass

            def _auth(self):
                q = parse_qs(urlparse(self.path).query)
                return q.get("t", [""])[0] == outer.s["token"]

            def _send(self, code, body, ctype):
                self.send_response(code)
                self.send_header("Content-Type", ctype)
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def _json(self, obj, code=200):
                self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                           "application/json; charset=utf-8")

            def do_GET(self):
                if not self._auth():
                    return self._send(403, "Yetkisiz. QR kodu programdan tekrar okutun.".encode("utf-8"),
                                      "text/plain; charset=utf-8")
                self._send(200, page, "text/html; charset=utf-8")

            def do_POST(self):
                if not self._auth() or not urlparse(self.path).path.endswith("/upload"):
                    return self._json({"ok": False, "hata": "yetkisiz"}, 403)
                try:
                    n = int(self.headers.get("Content-Length", "0"))
                    if n > 200 * 1024 * 1024:
                        return self._json({"ok": False, "hata": "dosya çok büyük"}, 413)
                    raw = self.rfile.read(n)
                    msg = BytesParser(policy=policy.default).parsebytes(
                        b"Content-Type: " + self.headers["Content-Type"].encode() + b"\r\n\r\n" + raw)
                    fields, files = {}, []
                    for part in msg.iter_parts():
                        name = part.get_param("name", header="content-disposition")
                        fname = part.get_filename()
                        payload = part.get_payload(decode=True) or b""
                        if fname:
                            files.append((fname, payload))
                        elif name:
                            fields[name] = payload.decode("utf-8", "ignore").strip()
                    ids = outer.kaydet(files, fields)
                    self._json({"ok": True, "adet": len(ids)})
                except Exception as e:
                    self._json({"ok": False, "hata": str(e)}, 500)

        try:
            self.httpd = ThreadingHTTPServer(("0.0.0.0", int(self.s["telefon_port"])), H)
            threading.Thread(target=self.httpd.serve_forever, daemon=True).start()
            self.hata = ""
        except OSError as e:
            self.hata = f"Port {self.s['telefon_port']} açılamadı: {e}"

    def stop(self):
        if self.httpd:
            self.httpd.shutdown()
            self.httpd.server_close()
            self.httpd = None

    def kaydet(self, files, fields):
        ids = []
        tutar = sayi_oku(fields.get("tutar", ""))
        for fname, data in files:
            if not data:
                continue
            ext = os.path.splitext(fname)[1].lower() or ".jpg"
            today = date.today().isoformat()
            secilen = fields.get("tarih") or today
            if not re.match(r"^\d{4}-\d{2}-\d{2}$", secilen):
                secilen = today
            rel = dosya_kaydet(data, ext, secilen, self.s.get("tarama_modu", True))
            kat = fields.get("kategori") or "Diğer"
            fid = outer_add(self.db, rel, secilen, kaynak="Telefon", kategori=kat,
                            tarih_kesin=1 if secilen != today else 0,
                            odeme=fields.get("odeme") or "Kredi Kartı",
                            plaka=fields.get("plaka", "").upper(),
                            aciklama=fields.get("not", ""), tutar=tutar or 0)
            ids.append(fid)
        if ids:
            self.bridge.yeni_fis.emit(ids)
        return ids


def outer_add(db, rel, tarih, **kw):
    tutar = kw.get("tutar") or 0
    oran = 20.0
    kdv = round(tutar - tutar / (1 + oran / 100), 2) if tutar else 0
    belge = "e-Arşiv Fatura" if rel.lower().endswith(".pdf") else "ÖKC Fişi"
    return db.add(tarih=tarih, dosya=rel, kdv_orani=oran, kdv=kdv, belge_turu=belge, **kw)


# ----------------------------------------------------------------- bulut (GitHub özel depo) senkron
GH_API = "https://api.github.com"


def gh_anahtar(s):
    return coz(s.get("github_anahtar", ""))


def bulut_link(s):
    """Telefonu tek seferde bağlayan kurulum linki (anahtar # sonrasında; sunucuya gitmez)."""
    tok, repo = gh_anahtar(s), (s.get("github_repo") or "").strip()
    sayfa = (s.get("telefon_sayfa") or "").strip()
    if not (tok and repo and sayfa):
        return ""
    return f"{sayfa}#r={quote(repo, safe='/')}&k={quote(tok)}"


def gh_istek(method, url, tok, body=None, raw=False, timeout=60):
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        "Authorization": f"Bearer {tok}",
        "Accept": "application/vnd.github.raw" if raw else "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "SinetechnicGider/1.1"})
    if data is not None:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, r.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()


def meta_coz(name):
    """2026-11-23__17000000abcd__<base64url-json>.jpg"""
    try:
        part = name.split("__")[2].rsplit(".", 1)[0]
        part += "=" * (-len(part) % 4)
        return json.loads(base64.urlsafe_b64decode(part).decode("utf-8"))
    except Exception:
        m = re.match(r"^(\d{4}-\d{2}-\d{2})", name)
        return {"tarih": m.group(1)} if m else {}


_bulut_kilit = threading.Lock()


class BulutIs(QRunnable):
    def __init__(self, s, db, bridge):
        super().__init__()
        self.s, self.db, self.bridge = s, db, bridge

    def run(self):
        if not _bulut_kilit.acquire(blocking=False):
            return
        try:
            tok, repo = gh_anahtar(self.s), (self.s.get("github_repo") or "").strip()
            if not tok or not repo:
                self.bridge.bulut_durum.emit("kapalı (Ayarlar → GitHub)")
                return
            st, raw = gh_istek("GET", f"{GH_API}/repos/{repo}/contents/gelen", tok)
            if st == 404:
                self.bridge.bulut_durum.emit(f"bağlı · bekleyen yok · {datetime.now():%H:%M}")
                return
            if st in (401, 403):
                self.bridge.bulut_durum.emit("⚠ GitHub anahtarı geçersiz/yetkisiz")
                return
            if st != 200:
                self.bridge.bulut_durum.emit(f"⚠ GitHub hata {st}")
                return
            liste = json.loads(raw.decode("utf-8"))
            ids = []
            for f in sorted(liste, key=lambda x: x.get("name", "")):
                if f.get("type") != "file":
                    continue
                name, path, sha = f["name"], f["path"], f["sha"]
                ext = os.path.splitext(name)[1].lower()
                if ext not in ALL_EXT:
                    continue
                st, data = gh_istek("GET", f"{GH_API}/repos/{repo}/contents/{quote(path)}", tok, raw=True, timeout=120)
                if st != 200 or not data:
                    continue
                meta = meta_coz(name)
                tarih = meta.get("tarih") or date.today().isoformat()
                if not re.match(r"^\d{4}-\d{2}-\d{2}$", tarih):
                    tarih = date.today().isoformat()
                rel = dosya_kaydet(data, ext, tarih, self.s.get("tarama_modu", True))
                yeni = outer_add(self.db, rel, tarih, kaynak="Telefon (GitHub)",
                                 kategori=meta.get("kategori") or "Diğer",
                                 odeme=meta.get("odeme") or "Kredi Kartı",
                                 plaka=(meta.get("plaka") or "").upper(),
                                 aciklama=meta.get("not") or "",
                                 tutar=sayi_oku(str(meta.get("tutar") or "")) or 0,
                                 tarih_kesin=1 if meta.get("tarih_kesin") else 0)
                ids.append(yeni)
                gh_istek("DELETE", f"{GH_API}/repos/{repo}/contents/{quote(path)}", tok,
                         {"message": f"alindi {name[:10]}", "sha": sha})
            self.bridge.bulut_durum.emit(f"bağlı · son kontrol {datetime.now():%H:%M}")
            if ids:
                self.bridge.yeni_fis.emit(ids)
        except Exception as e:
            self.bridge.bulut_durum.emit(f"⚠ ulaşılamadı ({type(e).__name__})")
        finally:
            _bulut_kilit.release()


# ----------------------------------------------------------------- arka plan OCR
class OcrIs(QRunnable):
    def __init__(self, db, fid, bridge, zorla=False, tara=False, tarama=True):
        super().__init__()
        self.db, self.fid, self.bridge = db, fid, bridge
        self.zorla, self.tara, self.tarama = zorla, tara, tarama

    def run(self):
        row = self.db.get(self.fid)
        if not row:
            return
        if self.tara and row["dosya"].lower().endswith(IMG_EXT):
            try:
                kaynak = orj_yolu(row["dosya"])
                kaynak = kaynak if os.path.exists(kaynak) else abs_path(row["dosya"])
                with open(kaynak, "rb") as f:
                    yeni, orj = gorsel_hazirla(f.read(), True)
                if not os.path.exists(orj_yolu(row["dosya"])):
                    os.makedirs(os.path.dirname(orj_yolu(row["dosya"])), exist_ok=True)
                    with open(orj_yolu(row["dosya"]), "wb") as f:
                        f.write(orj)
                with open(abs_path(row["dosya"]), "wb") as f:
                    f.write(yeni)
            except Exception as e:
                log_yaz("tarama_log.txt", f"yeniden tarama: {e!r}")
        text = ocr_metin(abs_path(row["dosya"]))
        if not text:
            self.bridge.ocr_bitti.emit(self.fid)
            return
        p = fis_ayristir(text)
        upd = {"ocr_metin": text}
        z = self.zorla
        for k in ("firma", "vkn", "belge_no", "plaka"):
            if p.get(k) and (z or not row[k]):
                upd[k] = p[k]
        if p.get("tarih") and (z or not row["tarih_kesin"]):
            try:
                od = date.fromisoformat(p["tarih"])
                if date.today().replace(year=date.today().year - 1) <= od <= date.today():
                    upd["tarih"] = p["tarih"]
            except ValueError:
                pass
        if p.get("kategori") and (z or row["kategori"] in ("", "Diğer")):
            upd["kategori"] = p["kategori"]
        if p.get("tutar") and (z or not row["tutar"]):
            upd["tutar"] = p["tutar"]
        tutar = upd.get("tutar", row["tutar"]) or 0
        if p.get("kdv_orani"):
            upd["kdv_orani"] = p["kdv_orani"]
        if p.get("kdv"):
            upd["kdv"] = p["kdv"]
        elif tutar:
            oran = upd.get("kdv_orani", row["kdv_orani"]) or 20
            upd["kdv"] = round(tutar - tutar / (1 + oran / 100), 2)
        # tarih değiştiyse dosyayı doğru ay klasörüne taşı
        if "tarih" in upd and upd["tarih"][:7] != row["tarih"][:7] and row["dosya"]:
            try:
                upd["dosya"] = dosya_tasi(row["dosya"], upd["tarih"][:7])
            except Exception as e:
                log_yaz("tarama_log.txt", f"taşıma: {e!r}")
        self.db.update(self.fid, **upd)
        self.bridge.ocr_bitti.emit(self.fid)


# ----------------------------------------------------------------- muhasebeci paketi
def paket_olustur(db, ym, settings, rows=None, ek_no=0):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfgen import canvas

    tum = db.month(ym)
    sira = {r["id"]: i for i, r in enumerate(tum, 1)}
    rows = list(rows) if rows is not None else tum
    if not rows:
        raise ValueError("Bu ayda gönderilecek fiş yok.")
    firma = guvenli_ad(settings.get("firma", "FIRMA"), 20)
    ek_etiket = f" - EK {ek_no}" if ek_no else ""
    out_dir = os.path.join(ROOT, "paketler", ym + (f"_EK{ek_no}" if ek_no else ""))
    if os.path.isdir(out_dir):
        shutil.rmtree(out_dir, ignore_errors=True)
    belge_dir = os.path.join(out_dir, "belgeler")
    os.makedirs(belge_dir, exist_ok=True)

    # belgeleri sıralı adlandırarak kopyala
    kopyalar = []
    for r in rows:
        i = sira.get(r["id"], 0)
        src = abs_path(r["dosya"])
        if src and os.path.exists(src):
            ext = os.path.splitext(src)[1]
            tut = f"{(r['tutar'] or 0):.2f}".replace(".", ",")
            ad = f"{i:03d}_{r['tarih']}_{guvenli_ad(r['firma'], 25)}_{tut}TL{ext}"
            dst = os.path.join(belge_dir, ad)
            shutil.copy2(src, dst)
            kopyalar.append((i, r, dst))
        else:
            kopyalar.append((i, r, None))

    # Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Gider Listesi"
    baslik = f"{settings.get('firma', '')} — {ay_adi(ym)} Gider Listesi{ek_etiket}"
    ws.append([baslik])
    ws.merge_cells("A1:N1")
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="121C30")
    ws["A1"].alignment = Alignment(horizontal="center")
    hdr = ["Sıra", "Tarih", "Belge Türü", "Belge No", "Firma / Satıcı", "VKN/TCKN",
           "Kategori", "Matrah", "KDV %", "KDV Tutarı", "Toplam", "Ödeme", "Plaka", "Açıklama"]
    ws.append(hdr)
    thin = Side(style="thin", color="BFC7D5")
    for c in ws[2]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="FF7A1A")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
    for r in rows:
        i = sira.get(r["id"], 0)
        matrah = round((r["tutar"] or 0) - (r["kdv"] or 0), 2)
        ws.append([i, datetime.strptime(r["tarih"], "%Y-%m-%d").date(), r["belge_turu"], r["belge_no"],
                   r["firma"], r["vkn"], r["kategori"], matrah, (r["kdv_orani"] or 0) / 100,
                   r["kdv"] or 0, r["tutar"] or 0, r["odeme"], r["plaka"], r["aciklama"]])
    last = ws.max_row
    ws.append(["", "", "", "", "TOPLAM", "", "", f"=SUM(H3:H{last})", "",
               f"=SUM(J3:J{last})", f"=SUM(K3:K{last})"])
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="E8EDF5")
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        for c in row:
            c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
        row[1].number_format = "DD.MM.YYYY"
        row[8].number_format = "0%"
        for idx in (7, 9, 10):
            row[idx].number_format = '#,##0.00 "₺"'
    widths = [6, 12, 15, 14, 32, 13, 22, 13, 8, 13, 14, 13, 11, 30]
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = "A3"

    ws2 = wb.create_sheet("Özet")
    ws2.append(["Kategori", "Adet", "KDV", "Toplam"])
    kat = {}
    for r in rows:
        k = kat.setdefault(r["kategori"], [0, 0.0, 0.0])
        k[0] += 1
        k[1] += r["kdv"] or 0
        k[2] += r["tutar"] or 0
    for k, v in sorted(kat.items()):
        ws2.append([k, v[0], round(v[1], 2), round(v[2], 2)])
    ws2.append([])
    ws2.append(["KDV Oranı", "Adet", "KDV", "Toplam"])
    oran = {}
    for r in rows:
        o = oran.setdefault(f"%{r['kdv_orani']:g}", [0, 0.0, 0.0])
        o[0] += 1
        o[1] += r["kdv"] or 0
        o[2] += r["tutar"] or 0
    for k, v in sorted(oran.items()):
        ws2.append([k, v[0], round(v[1], 2), round(v[2], 2)])
    for c in ("A", "B", "C", "D"):
        ws2.column_dimensions[c].width = 26
    for row in ws2.iter_rows():
        if row[0].value in ("Kategori", "KDV Oranı"):
            for c in row:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="121C30")
        else:
            for c in row[2:]:
                c.number_format = '#,##0.00 "₺"'
    ek_dosya = f"_EK{ek_no}" if ek_no else ""
    xlsx = os.path.join(out_dir, f"{ym}_{firma}_Gider_Listesi{ek_dosya}.xlsx")
    wb.save(xlsx)

    # Birleşik PDF (fotoğraflı fişler, sayfa başına bir belge)
    font = "Helvetica"
    for fp in (r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\segoeui.ttf",
               "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"):
        if os.path.exists(fp):
            try:
                pdfmetrics.registerFont(TTFont("TRFont", fp))
                font = "TRFont"
                break
            except Exception:
                pass
    pdf = os.path.join(out_dir, f"{ym}_{firma}_Fisler{ek_dosya}.pdf")
    c = canvas.Canvas(pdf, pagesize=A4)
    W, H = A4
    c.setFont(font, 18)
    c.drawString(40, H - 60, f"{settings.get('firma', '')}")
    c.setFont(font, 14)
    c.drawString(40, H - 85, f"{ay_adi(ym)} — Gider Belgeleri{ek_etiket}")
    c.setFont(font, 10)
    toplam = sum((r["tutar"] or 0) for r in rows)
    kdv_t = sum((r["kdv"] or 0) for r in rows)
    c.drawString(40, H - 110, f"Belge sayısı: {len(rows)}    Toplam: {tl(toplam)}    KDV: {tl(kdv_t)}")
    y = H - 140
    for r in rows:
        i = sira.get(r["id"], 0)
        if y < 50:
            c.showPage()
            c.setFont(font, 10)
            y = H - 50
        c.drawString(40, y, f"{i:03d}  {r['tarih'][8:10]}.{r['tarih'][5:7]}.{r['tarih'][:4]}  "
                            f"{(r['firma'] or '-')[:38]:<38}  {r['kategori'][:18]:<18}  {tl(r['tutar'] or 0)}")
        y -= 15
    c.showPage()
    pdf_ekleri = []
    for i, r, dst in kopyalar:
        if not dst:
            continue
        if dst.lower().endswith(".pdf"):
            pdf_ekleri.append(dst)
            continue
        c.setFont(font, 11)
        c.drawString(30, H - 30, f"No {i:03d} | {r['tarih']} | {(r['firma'] or '-')[:40]} | "
                                 f"{tl(r['tutar'] or 0)} | {r['kategori']}")
        try:
            img = ImageReader(dst)
            iw, ih = img.getSize()
            maxw, maxh = W - 60, H - 80
            sc = min(maxw / iw, maxh / ih)
            c.drawImage(img, (W - iw * sc) / 2, (H - 45 - ih * sc), iw * sc, ih * sc)
        except Exception as e:
            c.drawString(30, H - 60, f"Görsel eklenemedi: {e}")
        c.showPage()
    c.save()

    # PDF belgeleri (e-arşiv) birleşik PDF'e ekle
    if pdf_ekleri:
        try:
            from pypdf import PdfReader, PdfWriter  # type: ignore
            w = PdfWriter()
            for p in [pdf] + pdf_ekleri:
                for pg in PdfReader(p).pages:
                    w.add_page(pg)
            tmp = pdf + ".tmp"
            with open(tmp, "wb") as f:
                w.write(f)
            os.replace(tmp, pdf)
        except Exception:
            pass

    # ZIP
    zp = os.path.join(out_dir, f"{ym}_{firma}_Gider_Paketi{ek_dosya}.zip")
    with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(xlsx, os.path.basename(xlsx))
        z.write(pdf, os.path.basename(pdf))
        for _, _, dst in kopyalar:
            if dst:
                z.write(dst, "belgeler/" + os.path.basename(dst))
    return {"klasor": out_dir, "xlsx": xlsx, "pdf": pdf, "zip": zp, "ek_no": ek_no,
            "adet": len(rows), "toplam": toplam, "kdv": kdv_t}


def eposta_gonder(settings, paket, ym, ekler):
    to = settings.get("muhasebeci_eposta", "").strip()
    user = settings.get("smtp_kullanici", "").strip()
    pw = coz(settings.get("smtp_sifre", ""))
    if not to or not user or not pw:
        raise ValueError("Ayarlar'da muhasebeci e-postası ve e-posta (SMTP) bilgileri eksik.")
    msg = EmailMessage()
    ek = paket.get("ek_no") or 0
    msg["Subject"] = f"{settings.get('firma', '')} - {ay_adi(ym)} Gider Fişleri" + (f" - EK {ek}" if ek else "")
    msg["From"] = user
    msg["To"] = to
    if settings.get("bana_kopya"):
        msg["Cc"] = user
    hitap = settings.get("muhasebeci_ad") or ""
    msg.set_content(
        f"Merhaba {hitap},\n\n{ay_adi(ym)} ayına ait gider belgeleri ektedir."
        + (f" Bu gönderim daha önce iletilen paketin EK {ek} belgeleridir (sonradan ulaşan fişler)." if ek else "")
        + f"\nSon teslim tarihi: {teslim_tarihi(ym, int(settings.get('teslim_gunu', 14))):%d.%m.%Y}\n\n"
        f"Belge sayısı : {paket['adet']}\nToplam tutar : {tl(paket['toplam'])}\n"
        f"Toplam KDV   : {tl(paket['kdv'])}\n\n"
        f"Ekler: Excel gider listesi ve tarihe göre sıralı birleşik PDF.\n\n"
        f"İyi çalışmalar,\n{settings.get('firma', '')}\n\n-- {APP_NAME} ile gönderilmiştir.")
    toplam_boyut = 0
    for p in ekler:
        with open(p, "rb") as f:
            data = f.read()
        toplam_boyut += len(data)
        if p.endswith(".xlsx"):
            mt, st = "application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif p.endswith(".pdf"):
            mt, st = "application", "pdf"
        else:
            mt, st = "application", "zip"
        msg.add_attachment(data, maintype=mt, subtype=st, filename=os.path.basename(p))
    if toplam_boyut > 24 * 1024 * 1024:
        raise ValueError(f"Ekler {toplam_boyut / 1048576:.1f} MB — e-posta sınırını (~25 MB) aşıyor. "
                         "Paket klasörünü açıp Drive/WeTransfer ile gönderin.")
    host, port = settings["smtp_host"], int(settings["smtp_port"])
    if settings.get("smtp_guvenlik") == "SSL":
        srv = smtplib.SMTP_SSL(host, port, timeout=60)
    else:
        srv = smtplib.SMTP(host, port, timeout=60)
        srv.ehlo()
        srv.starttls()
        srv.ehlo()
    srv.login(user, pw)
    srv.send_message(msg)
    srv.quit()


# ----------------------------------------------------------------- arayüz
STYLE = """
* { font-family: 'Segoe UI', Arial; font-size: 10pt; }
QMainWindow, QDialog { background: #0e1626; }
QWidget { color: #e8edf5; }
QLabel#title { font-size: 18pt; font-weight: 700; color: #ffffff; }
QLabel#accent { color: #ff7a1a; font-size: 18pt; font-weight: 700; }
QLabel#sub { color: #8ea0bd; }
QFrame#card { background: #16223a; border: 1px solid #22324f; border-radius: 12px; }
QLabel#cardval { font-size: 16pt; font-weight: 700; color: #ffffff; }
QLabel#cardlbl { color: #8ea0bd; font-size: 9pt; }
QPushButton { background: #1d2b47; border: 1px solid #2c3f63; border-radius: 8px; padding: 8px 14px; }
QPushButton:hover { background: #27385b; }
QPushButton#primary { background: #ff7a1a; border: none; color: #fff; font-weight: 700; }
QPushButton#primary:hover { background: #ff8f3d; }
QPushButton#green { background: #2e8b57; border: none; color: #fff; font-weight: 700; }
QPushButton#danger { background: #7a2430; border: none; }
QTableWidget { background: #121c30; alternate-background-color: #16223a; gridline-color: #22324f;
               border: 1px solid #22324f; border-radius: 8px; selection-background-color: #ff7a1a; }
QHeaderView::section { background: #1d2b47; color: #cfd8e6; padding: 6px; border: none; font-weight: 600; }
QLineEdit, QComboBox, QDateEdit, QDoubleSpinBox, QSpinBox, QTextEdit {
    background: #121c30; border: 1px solid #2c3f63; border-radius: 6px; padding: 6px; }
QComboBox QAbstractItemView { background: #121c30; selection-background-color: #ff7a1a; }
QScrollArea { border: 1px solid #22324f; border-radius: 8px; background: #0b1220; }
QLabel#banner { background: #5a3a12; border: 1px solid #ff7a1a; border-radius: 8px; padding: 8px; }
QCheckBox::indicator { width: 16px; height: 16px; }
"""


def card(label):
    f = QFrame()
    f.setObjectName("card")
    lay = QVBoxLayout(f)
    lay.setContentsMargins(14, 10, 14, 10)
    v = QLabel("-")
    v.setObjectName("cardval")
    l = QLabel(label)
    l.setObjectName("cardlbl")
    lay.addWidget(v)
    lay.addWidget(l)
    return f, v


def qr_label(text, box=8):
    qr = QLabel()
    qr.setAlignment(Qt.AlignCenter)
    try:
        import qrcode
        img = qrcode.make(text, box_size=box, border=2).convert("RGB")
        b = BytesIO()
        img.save(b, "PNG")
        qr.setPixmap(QPixmap.fromImage(QImage.fromData(b.getvalue())))
    except Exception:
        qr.setText("(QR oluşturulamadı)")
    return qr


class TelefonDialog(QDialog):
    def __init__(self, parent, sunucu, settings):
        super().__init__(parent)
        self.setWindowTitle("Telefondan Fiş Ekle")
        lay = QVBoxLayout(self)
        link = bulut_link(settings)
        if link:
            lay.addWidget(QLabel("<h3>☁ Her yerden (mobil veri dahil)</h3>"
                                 "1) iPhone kamerasıyla QR kodu okut — telefon bir kez bağlanır.<br>"
                                 "2) Safari'de <b>Paylaş → Ana Ekrana Ekle</b> de — uygulama gibi açılır.<br>"
                                 "3) Fiş çek → <b>Gönder</b>. Program açık olduğunda fişler otomatik gelir.<br>"
                                 "<span style='color:#ff9a4d'>Bu QR'da GitHub anahtarınız var, kimseyle paylaşmayın.</span>"))
            lay.addWidget(qr_label(link, 7))
            le = QLineEdit(link)
            le.setReadOnly(True)
            lay.addWidget(le)
        else:
            lay.addWidget(QLabel("☁ Her yerden kullanmak için Ayarlar → <b>GitHub</b> bölümünü doldurun."))
        lay.addWidget(QLabel("<hr><b>📶 Aynı Wi-Fi'dayken (yerel bağlantı)</b>"))
        if sunucu.hata:
            lay.addWidget(QLabel(sunucu.hata + "\nAyarlar'dan başka bir port seçin."))
        else:
            info = QLabel("Telefon ve bilgisayar aynı ağdaysa bu QR ile fişler anında gelir.")
            info.setWordWrap(True)
            lay.addWidget(info)
            lay.addWidget(qr_label(sunucu.url, 4 if link else 7))
            url = QLineEdit(sunucu.url)
            url.setReadOnly(True)
            lay.addWidget(url)
            lay.addWidget(QLabel("<span style='color:#8ea0bd'>İlk açılışta Windows Güvenlik Duvarı "
                                 "izin sorarsa <b>Özel ağlar</b> için izin verin.</span>"))
        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)


class AyarlarDialog(QDialog):
    def __init__(self, parent, s):
        super().__init__(parent)
        self.s = s
        self.setWindowTitle("Ayarlar")
        self.setMinimumWidth(480)
        f = QFormLayout(self)
        self.firma = QLineEdit(s["firma"])
        self.m_ad = QLineEdit(s["muhasebeci_ad"])
        self.m_ep = QLineEdit(s["muhasebeci_eposta"])
        self.kopya = QCheckBox("Gönderilen e-postanın kopyası bana da gelsin")
        self.kopya.setChecked(bool(s.get("bana_kopya")))
        self.host = QLineEdit(s["smtp_host"])
        self.port = QSpinBox()
        self.port.setRange(1, 65535)
        self.port.setValue(int(s["smtp_port"]))
        self.guv = QComboBox()
        self.guv.addItems(["STARTTLS", "SSL"])
        self.guv.setCurrentText(s.get("smtp_guvenlik", "STARTTLS"))
        self.user = QLineEdit(s["smtp_kullanici"])
        self.pw = QLineEdit(coz(s.get("smtp_sifre", "")))
        self.pw.setEchoMode(QLineEdit.Password)
        self.tport = QSpinBox()
        self.tport.setRange(1024, 65535)
        self.tport.setValue(int(s["telefon_port"]))
        self.ocr = QCheckBox("Fişleri otomatik oku (Windows OCR)")
        self.ocr.setChecked(bool(s.get("otomatik_ocr")))
        self.tarama = QCheckBox("Fotoğrafları tarama görünümüne çevir (siyah-beyaz, net)")
        self.tarama.setChecked(bool(s.get("tarama_modu")))
        hazir = QComboBox()
        hazir.addItems(["Hazır ayar seç…", "Outlook / Hotmail", "Gmail", "Yandex", "Yaani Mail"])
        hazir.currentTextChanged.connect(self.hazir_sec)

        f.addRow(QLabel("<b>Firma</b>"))
        f.addRow("Firma adı:", self.firma)
        f.addRow(QLabel("<b>Muhasebeci</b>"))
        f.addRow("Adı:", self.m_ad)
        f.addRow("E-posta:", self.m_ep)
        f.addRow("", self.kopya)
        f.addRow(QLabel("<b>Gönderen e-posta (SMTP)</b>"))
        f.addRow("Hazır ayar:", hazir)
        f.addRow("Sunucu:", self.host)
        f.addRow("Port:", self.port)
        f.addRow("Güvenlik:", self.guv)
        f.addRow("E-posta adresi:", self.user)
        f.addRow("Şifre / Uygulama şifresi:", self.pw)
        not_ = QLabel("<span style='color:#8ea0bd'>Gmail ve Outlook'ta normal şifre yerine "
                      "<b>uygulama şifresi</b> gerekir (hesapta iki adımlı doğrulama açık olmalı). "
                      "Şifre Windows'un kendi şifrelemesiyle (DPAPI) saklanır.</span>")
        not_.setWordWrap(True)
        f.addRow(not_)
        f.addRow(QLabel("<b>📅 Teslim protokolü</b>"))
        self.teslim = QSpinBox()
        self.teslim.setRange(1, 28)
        self.teslim.setValue(int(s.get("teslim_gunu", 14)))
        self.teslim.setSuffix(". gün (sonraki ay)")
        self.oto = QCheckBox("Geçen ayın fişlerini otomatik e-postala")
        self.oto.setChecked(bool(s.get("otomatik_gonder")))
        self.otogun = QSpinBox()
        self.otogun.setRange(1, 28)
        self.otogun.setValue(int(s.get("otomatik_gun", 10)))
        self.otogun.setSuffix(". gün")
        f.addRow("Muhasebeci son teslim:", self.teslim)
        f.addRow("", self.oto)
        f.addRow("Otomatik gönderim günü:", self.otogun)
        f.addRow(QLabel("<b>☁ GitHub (her yerden telefonla ekleme)</b>"))
        self.gh_repo = QLineEdit(s.get("github_repo", ""))
        self.gh_tok = QLineEdit(gh_anahtar(s))
        self.gh_tok.setEchoMode(QLineEdit.Password)
        self.gh_tok.setPlaceholderText("github_pat_… (yalnızca bu depoya Contents: Read and write)")
        self.gh_sayfa = QLineEdit(s.get("telefon_sayfa", ""))
        f.addRow("GitHub özel depo:", self.gh_repo)
        f.addRow("GitHub anahtarı:", self.gh_tok)
        f.addRow("Telefon sayfası:", self.gh_sayfa)
        f.addRow(QLabel("<b>Diğer</b>"))
        f.addRow("Telefon bağlantı portu:", self.tport)
        f.addRow("", self.ocr)
        f.addRow("", self.tarama)
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.kaydet)
        bb.rejected.connect(self.reject)
        f.addRow(bb)

    def hazir_sec(self, t):
        presets = {"Outlook / Hotmail": ("smtp-mail.outlook.com", 587, "STARTTLS"),
                   "Gmail": ("smtp.gmail.com", 587, "STARTTLS"),
                   "Yandex": ("smtp.yandex.com", 465, "SSL"),
                   "Yaani Mail": ("smtp.yaani.com", 587, "STARTTLS")}
        if t in presets:
            h, p, g = presets[t]
            self.host.setText(h)
            self.port.setValue(p)
            self.guv.setCurrentText(g)

    def kaydet(self):
        s = self.s
        s["firma"] = self.firma.text().strip()
        s["muhasebeci_ad"] = self.m_ad.text().strip()
        s["muhasebeci_eposta"] = self.m_ep.text().strip()
        s["bana_kopya"] = self.kopya.isChecked()
        s["smtp_host"] = self.host.text().strip()
        s["smtp_port"] = self.port.value()
        s["smtp_guvenlik"] = self.guv.currentText()
        s["smtp_kullanici"] = self.user.text().strip()
        s["smtp_sifre"] = sifrele(self.pw.text())
        s["telefon_port"] = self.tport.value()
        s["otomatik_ocr"] = self.ocr.isChecked()
        s["tarama_modu"] = self.tarama.isChecked()
        s["teslim_gunu"] = self.teslim.value()
        s["otomatik_gonder"] = self.oto.isChecked()
        s["otomatik_gun"] = min(self.otogun.value(), self.teslim.value())
        s["github_repo"] = self.gh_repo.text().strip()
        s["github_anahtar"] = sifrele(self.gh_tok.text().strip())
        s["telefon_sayfa"] = self.gh_sayfa.text().strip()
        save_settings(s)
        self.accept()


class GonderDialog(QDialog):
    def __init__(self, parent, ym, adet, toplam, alici):
        super().__init__(parent)
        self.setWindowTitle("Muhasebeciye Gönder")
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(f"<h3>{ay_adi(ym)}</h3>{adet} belge · Toplam <b>{tl(toplam)}</b><br>"
                             f"Alıcı: <b>{alici or '(Ayarlar’da tanımlı değil)'}</b>"))
        self.xlsx = QCheckBox("Excel gider listesi")
        self.pdf = QCheckBox("Birleşik PDF (tüm fişler sıralı)")
        self.zip = QCheckBox("ZIP (orijinal dosyalar dahil)")
        self.xlsx.setChecked(True)
        self.pdf.setChecked(True)
        for w in (self.xlsx, self.pdf, self.zip):
            lay.addWidget(w)
        h = QHBoxLayout()
        self.b_mail = QPushButton("📧 E-posta ile Gönder")
        self.b_mail.setObjectName("primary")
        self.b_klasor = QPushButton("📁 Sadece Paketi Hazırla")
        iptal = QPushButton("İptal")
        h.addWidget(self.b_mail)
        h.addWidget(self.b_klasor)
        h.addWidget(iptal)
        lay.addLayout(h)
        self.secim = None
        self.b_mail.clicked.connect(lambda: self._sec("mail"))
        self.b_klasor.clicked.connect(lambda: self._sec("klasor"))
        iptal.clicked.connect(self.reject)

    def _sec(self, s):
        self.secim = s
        self.accept()


class ImageView(QScrollArea):
    def __init__(self):
        super().__init__()
        self.lbl = QLabel("Fiş seçin")
        self.lbl.setAlignment(Qt.AlignCenter)
        self.lbl.setStyleSheet("color:#5f7090;")
        self.setWidget(self.lbl)
        self.setWidgetResizable(True)
        self.pix = None

    def set_path(self, path):
        if path and os.path.exists(path) and path.lower().endswith(IMG_EXT):
            self.pix = QPixmap(path)
            self._fit()
        elif path and path.lower().endswith(".pdf"):
            self.pix = None
            self.lbl.setPixmap(QPixmap())
            self.lbl.setText("📄 PDF belge\n(Aç butonuyla görüntüleyin)")
        else:
            self.pix = None
            self.lbl.setPixmap(QPixmap())
            self.lbl.setText("Fiş seçin")

    def _fit(self):
        if self.pix and not self.pix.isNull():
            sz = self.viewport().size()
            self.lbl.setPixmap(self.pix.scaled(sz.width() - 8, sz.height() - 8,
                                               Qt.KeepAspectRatio, Qt.SmoothTransformation))

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._fit()


class Main(QMainWindow):
    COLS = ["No", "Tarih", "Firma", "Kategori", "Tutar", "KDV", "Ödeme", "Plaka", "Durum"]

    def __init__(self):
        super().__init__()
        self.s = load_settings()
        self.db = DB(DB_PATH)
        self.bridge = Bridge()
        self.bridge.yeni_fis.connect(self.yeni_fisler)
        self.bridge.ocr_bitti.connect(self.ocr_bitti)
        self.bridge.bulut_durum.connect(self._bulut_durum)
        self.bulut_txt = "kontrol ediliyor…"
        self.pool = QThreadPool.globalInstance()
        self.sunucu = TelefonSunucu(self.s, self.db, self.bridge)
        self.sunucu.start()
        self.cur_id = None
        self.loading = False

        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        ico = resource_path(os.path.join("assets", "icon.ico"))
        if os.path.exists(ico):
            self.setWindowIcon(QIcon(ico))
        self.resize(1360, 820)
        self.setAcceptDrops(True)
        self._ui()
        self.ay_listesi()
        self.yenile()
        self.oto_hata = ""
        QTimer.singleShot(600, self.hatirlatma)
        self.saat_timer = QTimer(self)
        self.saat_timer.timeout.connect(self.hatirlatma)
        self.saat_timer.timeout.connect(self.otomatik_kontrol)
        self.saat_timer.start(60 * 60 * 1000)
        QTimer.singleShot(3 * 60 * 1000, self.otomatik_kontrol)
        self.bulut_timer = QTimer(self)
        self.bulut_timer.timeout.connect(self.bulut_kontrol)
        self.bulut_timer.start(max(20, int(self.s.get("bulut_aralik_sn", 60))) * 1000)
        QTimer.singleShot(1500, self.bulut_kontrol)

    # ------------------------------------------------ arayüz kurulumu
    def _ui(self):
        w = QWidget()
        root = QVBoxLayout(w)
        root.setContentsMargins(16, 12, 16, 12)
        root.setSpacing(10)

        top = QHBoxLayout()
        t1 = QLabel("SİNETECHNİC")
        t1.setObjectName("title")
        t2 = QLabel("GİDER")
        t2.setObjectName("accent")
        top.addWidget(t1)
        top.addWidget(t2)
        sub = QLabel("  Fiş topla · sırala · muhasebeciye gönder")
        sub.setObjectName("sub")
        top.addWidget(sub)
        top.addStretch()
        top.addWidget(QLabel("Ay:"))
        self.ay = QComboBox()
        self.ay.setMinimumWidth(150)
        self.ay.currentIndexChanged.connect(self.yenile)
        top.addWidget(self.ay)
        self.teslim_lbl = QLabel()
        self.teslim_lbl.setObjectName("sub")
        top.addWidget(self.teslim_lbl)
        root.addLayout(top)

        bar = QHBoxLayout()
        b_tel = QPushButton("📱 Telefondan Ekle")
        b_tel.setObjectName("primary")
        b_tel.clicked.connect(self.telefon)
        b_bul = QPushButton("☁ Buluttan Al")
        b_bul.clicked.connect(self.bulut_kontrol)
        b_dos = QPushButton("➕ Dosyadan Ekle")
        b_dos.clicked.connect(self.dosyadan)
        b_gon = QPushButton("📤 Muhasebeciye Gönder")
        b_gon.setObjectName("green")
        b_gon.clicked.connect(self.gonder)
        b_kl = QPushButton("📁 Klasörü Aç")
        b_kl.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(ROOT)))
        b_ay = QPushButton("⚙ Ayarlar")
        b_ay.clicked.connect(self.ayarlar)
        b_hk = QPushButton("ℹ Hakkında")
        b_hk.clicked.connect(self.hakkinda)
        self.ara = QLineEdit()
        self.ara.setPlaceholderText("🔎 Ara (firma, plaka, açıklama)…")
        self.ara.textChanged.connect(self.tablo_doldur)
        for b in (b_tel, b_bul, b_dos, b_gon, b_kl, b_ay, b_hk):
            bar.addWidget(b)
        bar.addStretch()
        bar.addWidget(self.ara, 1)
        root.addLayout(bar)

        self.banner = QLabel()
        self.banner.setObjectName("banner")
        self.banner.setVisible(False)
        root.addWidget(self.banner)

        cards = QHBoxLayout()
        c1, self.v_adet = card("Belge sayısı")
        c2, self.v_top = card("Toplam tutar")
        c3, self.v_kdv = card("Toplam KDV")
        c4, self.v_yakit = card("Akaryakıt")
        c5, self.v_bek = card("Gönderilmedi")
        for c in (c1, c2, c3, c4, c5):
            cards.addWidget(c)
        root.addLayout(cards)

        split = QSplitter(Qt.Horizontal)
        self.tbl = QTableWidget(0, len(self.COLS))
        self.tbl.setHorizontalHeaderLabels(self.COLS)
        self.tbl.setAlternatingRowColors(True)
        self.tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl.verticalHeader().setVisible(False)
        hh = self.tbl.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.Stretch)
        self.tbl.itemSelectionChanged.connect(self.secildi)
        split.addWidget(self.tbl)

        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(8, 0, 0, 0)
        self.view = ImageView()
        self.view.setMinimumHeight(300)
        rl.addWidget(self.view, 3)
        ib = QHBoxLayout()
        b_rot = QPushButton("⟳ Döndür")
        b_rot.clicked.connect(self.dondur)
        b_open = QPushButton("🔍 Aç")
        b_open.clicked.connect(self.ac)
        b_ocr = QPushButton("🧾 Fişi Oku")
        b_ocr.setToolTip("Fişi tekrar okuyup boş/yanlış alanları doldurur")
        b_ocr.clicked.connect(lambda: self.ocr_calistir(zorla=True))
        b_tara = QPushButton("📄 Yeniden Tara")
        b_tara.setToolTip("Orijinal fotoğraftan fişi tekrar bulup düzleştirir ve okur")
        b_tara.clicked.connect(lambda: self.ocr_calistir(zorla=True, tara=True))
        self.b_orj = QPushButton("🖼 Orijinal")
        self.b_orj.setCheckable(True)
        self.b_orj.setToolTip("Taranmış hal ile orijinal fotoğraf arasında geçiş")
        self.b_orj.toggled.connect(self.orj_goster)
        ib.addWidget(b_rot)
        ib.addWidget(b_open)
        ib.addWidget(b_tara)
        ib.addWidget(b_ocr)
        ib.addWidget(self.b_orj)
        rl.addLayout(ib)

        form = QGridLayout()
        self.f_tarih = QDateEdit()
        self.f_tarih.setCalendarPopup(True)
        self.f_tarih.setDisplayFormat("dd.MM.yyyy")
        self.f_firma = QLineEdit()
        self.f_vkn = QLineEdit()
        self.f_tur = QComboBox()
        self.f_tur.addItems(BELGE_TURLERI)
        self.f_no = QLineEdit()
        self.f_kat = QComboBox()
        self.f_kat.setEditable(True)
        self.f_kat.addItems(KATEGORILER)
        self.f_tutar = QDoubleSpinBox()
        self.f_tutar.setRange(0, 99999999)
        self.f_tutar.setDecimals(2)
        self.f_tutar.setSuffix(" ₺")
        self.f_tutar.setGroupSeparatorShown(True)
        self.f_oran = QComboBox()
        self.f_oran.addItems(["20", "10", "1", "0"])
        self.f_kdv = QDoubleSpinBox()
        self.f_kdv.setRange(0, 99999999)
        self.f_kdv.setDecimals(2)
        self.f_kdv.setSuffix(" ₺")
        self.f_odeme = QComboBox()
        self.f_odeme.addItems(ODEMELER)
        self.f_plaka = QLineEdit()
        self.f_acik = QLineEdit()
        self.f_tutar.valueChanged.connect(self.kdv_hesapla)
        self.f_oran.currentTextChanged.connect(self.kdv_hesapla)
        r = 0
        for lbl, wdg, lbl2, wdg2 in [
            ("Tarih", self.f_tarih, "Belge türü", self.f_tur),
            ("Firma", self.f_firma, "VKN/TCKN", self.f_vkn),
            ("Belge no", self.f_no, "Kategori", self.f_kat),
            ("Tutar (KDV dahil)", self.f_tutar, "KDV %", self.f_oran),
            ("KDV tutarı", self.f_kdv, "Ödeme", self.f_odeme),
            ("Plaka", self.f_plaka, "Açıklama", self.f_acik)]:
            form.addWidget(QLabel(lbl), r, 0)
            form.addWidget(wdg, r, 1)
            form.addWidget(QLabel(lbl2), r, 2)
            form.addWidget(wdg2, r, 3)
            r += 1
        rl.addLayout(form)
        fb = QHBoxLayout()
        b_kay = QPushButton("💾 Kaydet")
        b_kay.setObjectName("primary")
        b_kay.clicked.connect(self.kaydet)
        b_sil = QPushButton("🗑 Sil")
        b_sil.setObjectName("danger")
        b_sil.clicked.connect(self.sil)
        fb.addWidget(b_kay)
        fb.addWidget(b_sil)
        rl.addLayout(fb)
        split.addWidget(right)
        split.setSizes([760, 580])
        root.addWidget(split, 1)

        self.status = QLabel()
        self.status.setObjectName("sub")
        root.addWidget(self.status)
        self.setCentralWidget(w)

        sc = QAction(self)
        sc.setShortcut("Ctrl+S")
        sc.triggered.connect(self.kaydet)
        self.addAction(sc)
        dl = QAction(self)
        dl.setShortcut("Delete")
        dl.triggered.connect(self.sil)
        self.addAction(dl)

    # ------------------------------------------------ veri
    def ym(self):
        return self.ay.currentData() or date.today().strftime("%Y-%m")

    def ay_listesi(self, sec=None):
        cur = sec or (self.ay.currentData() if self.ay.count() else None)
        ms = set(self.db.months())
        today = date.today()
        ms.add(today.strftime("%Y-%m"))
        prev = date(today.year - (today.month == 1), (today.month - 2) % 12 + 1, 1)
        ms.add(prev.strftime("%Y-%m"))
        self.ay.blockSignals(True)
        self.ay.clear()
        for m in sorted(ms, reverse=True):
            self.ay.addItem(ay_adi(m), m)
        idx = self.ay.findData(cur or today.strftime("%Y-%m"))
        self.ay.setCurrentIndex(max(idx, 0))
        self.ay.blockSignals(False)

    def yenile(self):
        self.rows = self.db.month(self.ym())
        self.tablo_doldur()
        top = sum(r["tutar"] or 0 for r in self.rows)
        kdv = sum(r["kdv"] or 0 for r in self.rows)
        yak = sum(r["tutar"] or 0 for r in self.rows if r["kategori"] == "Akaryakıt")
        bek = sum(1 for r in self.rows if not r["gonderildi"])
        self.v_adet.setText(str(len(self.rows)))
        self.v_top.setText(tl(top))
        self.v_kdv.setText(tl(kdv))
        self.v_yakit.setText(tl(yak))
        self.v_bek.setText(str(bek))
        dl = teslim_tarihi(self.ym(), int(self.s.get("teslim_gunu", 14)))
        if not self.rows:
            durum = ""
        elif bek == 0:
            durum = "  ✅ gönderildi"
        elif date.today() > dl:
            durum = "  🔴 süre geçti"
        else:
            durum = f"  ({(dl - date.today()).days} gün)"
        self.teslim_lbl.setText(f"  Muhasebeci teslimi: {dl:%d.%m.%Y}{durum}")
        port_ok = "" if not self.sunucu.hata else f"  ⚠ {self.sunucu.hata}"
        self.status.setText(f"Veri klasörü: {ROOT}   ·   Wi-Fi bağlantısı: "
                            f"{'açık' if self.sunucu.httpd else 'kapalı'}{port_ok}   ·   ☁ Bulut: {self.bulut_txt}")

    def tablo_doldur(self):
        q = self.ara.text().strip().lower()
        rows = [r for r in self.rows if not q or q in " ".join(
            str(r[k] or "") for k in ("firma", "plaka", "aciklama", "kategori", "belge_no")).lower()]
        self.tbl.setRowCount(len(rows))
        sel_row = -1
        for i, r in enumerate(rows):
            eksik = not r["tutar"] or not r["firma"]
            durum = "✅ Gönderildi" if r["gonderildi"] else ("⚠ Eksik bilgi" if eksik else "Bekliyor")
            vals = [str(self.rows.index(r) + 1), datetime.strptime(r["tarih"], "%Y-%m-%d").strftime("%d.%m.%Y"),
                    r["firma"] or "—", r["kategori"], tl(r["tutar"] or 0), tl(r["kdv"] or 0),
                    r["odeme"], r["plaka"], durum]
            for j, v in enumerate(vals):
                it = QTableWidgetItem(v)
                it.setData(Qt.UserRole, r["id"])
                if j in (4, 5):
                    it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if j == 8 and eksik and not r["gonderildi"]:
                    it.setForeground(QColor("#ffb347"))
                if j == 8 and r["gonderildi"]:
                    it.setForeground(QColor("#5fd38d"))
                self.tbl.setItem(i, j, it)
            if r["id"] == self.cur_id:
                sel_row = i
        if sel_row >= 0:
            self.tbl.selectRow(sel_row)
        elif not rows:
            self.form_temizle()

    def secildi(self):
        items = self.tbl.selectedItems()
        if not items:
            return
        fid = items[0].data(Qt.UserRole)
        r = self.db.get(fid)
        if not r:
            return
        self.cur_id = fid
        self.loading = True
        y, m, d = map(int, r["tarih"].split("-"))
        self.f_tarih.setDate(QDate(y, m, d))
        self.f_firma.setText(r["firma"] or "")
        self.f_vkn.setText(r["vkn"] or "")
        self.f_tur.setCurrentText(r["belge_turu"] or "ÖKC Fişi")
        self.f_no.setText(r["belge_no"] or "")
        self.f_kat.setCurrentText(r["kategori"] or "Diğer")
        self.f_tutar.setValue(r["tutar"] or 0)
        self.f_oran.setCurrentText(f"{(r['kdv_orani'] or 0):g}")
        self.f_kdv.setValue(r["kdv"] or 0)
        self.f_odeme.setCurrentText(r["odeme"] or "Kredi Kartı")
        self.f_plaka.setText(r["plaka"] or "")
        self.f_acik.setText(r["aciklama"] or "")
        self.loading = False
        self.b_orj.blockSignals(True)
        self.b_orj.setChecked(False)
        self.b_orj.setEnabled(bool(r["dosya"]) and os.path.exists(orj_yolu(r["dosya"])))
        self.b_orj.blockSignals(False)
        self.view.set_path(abs_path(r["dosya"]))

    def form_temizle(self):
        self.cur_id = None
        self.view.set_path("")

    def kdv_hesapla(self):
        if self.loading:
            return
        t = self.f_tutar.value()
        o = float(self.f_oran.currentText() or 0)
        self.f_kdv.setValue(round(t - t / (1 + o / 100), 2) if t else 0)

    def kaydet(self):
        if not self.cur_id:
            return
        r = self.db.get(self.cur_id)
        yeni_tarih = self.f_tarih.date().toString("yyyy-MM-dd")
        upd = dict(tarih=yeni_tarih, firma=self.f_firma.text().strip(), vkn=self.f_vkn.text().strip(),
                   belge_turu=self.f_tur.currentText(), belge_no=self.f_no.text().strip(),
                   kategori=self.f_kat.currentText().strip(), tutar=self.f_tutar.value(),
                   kdv_orani=float(self.f_oran.currentText() or 0), kdv=self.f_kdv.value(),
                   odeme=self.f_odeme.currentText(), plaka=self.f_plaka.text().strip().upper(),
                   aciklama=self.f_acik.text().strip())
        upd["tarih_kesin"] = 1
        if r["dosya"] and yeni_tarih[:7] != r["tarih"][:7]:
            try:
                upd["dosya"] = dosya_tasi(r["dosya"], yeni_tarih[:7])
            except Exception:
                pass
        self.db.update(self.cur_id, **upd)
        if yeni_tarih[:7] != self.ym():
            self.ay_listesi(yeni_tarih[:7])
        self.yenile()
        self.status.setText(f"✔ Kaydedildi — {self.f_firma.text()} {tl(self.f_tutar.value())}")

    def sil(self):
        if not self.cur_id:
            return
        if QMessageBox.question(self, "Sil", "Bu fiş ve dosyası silinsin mi?") != QMessageBox.Yes:
            return
        r = self.db.get(self.cur_id)
        p = abs_path(r["dosya"])
        if p and os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass
        self.db.delete(self.cur_id)
        self.cur_id = None
        self.yenile()

    def dondur(self):
        if not self.cur_id:
            return
        p = abs_path(self.db.get(self.cur_id)["dosya"])
        if p.lower().endswith(IMG_EXT):
            im = Image.open(p)
            im.rotate(-90, expand=True).save(p, quality=90)
            self.view.set_path(p)

    def ac(self):
        if self.cur_id:
            QDesktopServices.openUrl(QUrl.fromLocalFile(abs_path(self.db.get(self.cur_id)["dosya"])))

    def ocr_calistir(self, fid=None, zorla=False, tara=False):
        fid = fid or self.cur_id
        if not fid:
            return
        self.status.setText("📄 Fiş taranıyor ve okunuyor…" if tara else "🧾 Fiş okunuyor… (ilk seferde 10-20 sn sürebilir)")
        self.pool.start(OcrIs(self.db, fid, self.bridge, zorla=zorla, tara=tara))

    def orj_goster(self, acik):
        if not self.cur_id:
            return
        r = self.db.get(self.cur_id)
        o = orj_yolu(r["dosya"]) if r and r["dosya"] else ""
        self.view.set_path(o if acik and o and os.path.exists(o) else abs_path(r["dosya"]))

    def ocr_bitti(self, fid):
        r = self.db.get(fid)
        if r and r["ocr_metin"]:
            self.status.setText(f"🧾 Fiş okundu: {r['firma'] or '-'} {tl(r['tutar'] or 0)} — kontrol edip kaydedin")
        elif r:
            self.status.setText("ℹ Otomatik okuma yapılamadı; bilgileri elle girin. "
                                "(Windows Ayarlar → Dil'den Türkçe dil paketi ekli olmalı)")
        self.ay_listesi(r["tarih"][:7] if r else None)
        self.yenile()

    # ------------------------------------------------ ekleme
    def yeni_fisler(self, ids):
        self.ay_listesi(date.today().strftime("%Y-%m"))
        if ids:
            self.cur_id = ids[-1]
        self.yenile()
        self.status.setText(f"📱 Telefondan {len(ids)} belge geldi")
        if not self.isActiveWindow():
            self.setWindowTitle(f"({len(ids)} yeni fiş) {APP_NAME} v{APP_VERSION}")
            QTimer.singleShot(15000, lambda: self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}"))
        QApplication.alert(self)
        if self.s.get("otomatik_ocr"):
            for i in ids:
                self.ocr_calistir(i)

    def ekle_dosyalar(self, paths):
        ids = []
        for p in paths:
            if not p.lower().endswith(ALL_EXT):
                continue
            with open(p, "rb") as f:
                data = f.read()
            t = date.fromtimestamp(os.path.getmtime(p)).isoformat()
            rel = dosya_kaydet(data, os.path.splitext(p)[1], t, self.s.get("tarama_modu", True))
            ids.append(outer_add(self.db, rel, t, kaynak="PC"))
        if ids:
            self.yeni_fisler(ids)
            self.status.setText(f"➕ {len(ids)} belge eklendi")

    def dosyadan(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Fiş / fatura seç", os.path.expanduser("~"),
                                                "Belgeler (*.jpg *.jpeg *.png *.heic *.webp *.pdf)")
        self.ekle_dosyalar(paths)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        self.ekle_dosyalar([u.toLocalFile() for u in e.mimeData().urls()])

    # ------------------------------------------------ gönderim
    def gonder(self):
        ym = self.ym()
        rows = self.db.month(ym)
        if not rows:
            QMessageBox.information(self, "Gönder", "Bu ayda fiş yok.")
            return
        yeni = [r for r in rows if not r["gonderildi"]]
        eski = len(rows) - len(yeni)
        hedef, ek_no = rows, 0
        if eski and yeni:
            mb = QMessageBox(self)
            mb.setWindowTitle("EK gönderim")
            mb.setText(f"{ay_adi(ym)} için {eski} fiş daha önce gönderildi, {len(yeni)} yeni fiş var.")
            b_ek = mb.addButton(f"Sadece yeni {len(yeni)} fişi EK olarak gönder", QMessageBox.AcceptRole)
            b_tum = mb.addButton("Tümünü yeniden gönder", QMessageBox.DestructiveRole)
            mb.addButton("İptal", QMessageBox.RejectRole)
            mb.exec()
            if mb.clickedButton() == b_ek:
                hedef, ek_no = yeni, self.sonraki_ek(rows)
            elif mb.clickedButton() != b_tum:
                return
        elif eski and not yeni:
            if QMessageBox.question(self, "Zaten gönderildi",
                                    f"{ay_adi(ym)} fişlerinin hepsi gönderilmiş. Yeniden gönderilsin mi?") != QMessageBox.Yes:
                return
        eksik = [r for r in hedef if not r["tutar"]]
        if eksik and QMessageBox.question(
                self, "Eksik bilgi", f"{len(eksik)} fişte tutar girilmemiş. Yine de devam edilsin mi?") != QMessageBox.Yes:
            return
        d = GonderDialog(self, ym, len(hedef), sum(r["tutar"] or 0 for r in hedef), self.s.get("muhasebeci_eposta"))
        if ek_no:
            d.setWindowTitle(f"Muhasebeciye Gönder — EK {ek_no}")
        if d.exec() != QDialog.Accepted:
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            ekler = [k for k, cb in (("xlsx", d.xlsx), ("pdf", d.pdf), ("zip", d.zip)) if cb.isChecked()]
            if d.secim == "mail":
                self.paket_gonder(ym, hedef, ek_no, ekler)
                QApplication.restoreOverrideCursor()
                QMessageBox.information(self, "Gönderildi",
                                        f"✅ {ay_adi(ym)}{f' EK {ek_no}' if ek_no else ''} fişleri "
                                        f"{self.s['muhasebeci_eposta']} adresine gönderildi.")
            else:
                paket = paket_olustur(self.db, ym, self.s, hedef, ek_no)
                QApplication.restoreOverrideCursor()
                QDesktopServices.openUrl(QUrl.fromLocalFile(paket["klasor"]))
                if QMessageBox.question(self, "Paket hazır",
                                        "Paket klasörü açıldı. Bu fişleri 'gönderildi' olarak işaretleyeyim mi?") == QMessageBox.Yes:
                    self.isaretle(hedef, ek_no)
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "Hata", f"İşlem tamamlanamadı:\n{e}")
        self.yenile()
        self.hatirlatma()

    def sonraki_ek(self, rows):
        return max([r["ek_no"] or 0 for r in rows] + [0]) + 1

    def isaretle(self, rows, ek_no):
        now = datetime.now().isoformat(timespec="seconds")
        for r in rows:
            self.db.update(r["id"], gonderildi=now, ek_no=ek_no)

    def paket_gonder(self, ym, rows, ek_no, ekler=("xlsx", "pdf")):
        paket = paket_olustur(self.db, ym, self.s, rows, ek_no)
        eposta_gonder(self.s, paket, ym, [paket[k] for k in ekler])
        self.isaretle(rows, ek_no)
        return paket

    def smtp_hazir(self):
        return bool(self.s.get("muhasebeci_eposta") and self.s.get("smtp_kullanici") and self.s.get("smtp_sifre"))

    def otomatik_kontrol(self):
        """Teslim protokolü: geçmiş ayların gönderilmemiş fişlerini otomatik gün geldiğinde e-postalar."""
        if not self.s.get("otomatik_gonder") or not self.smtp_hazir():
            return
        today = date.today()
        if self.s.get("son_oto_deneme") == today.isoformat():
            return
        gun = int(self.s.get("teslim_gunu", 14))
        oto = min(int(self.s.get("otomatik_gun", 10)), gun)
        cur = today.strftime("%Y-%m")
        sonuc, hata = [], []
        for ym, n in self.db.unsent_months():
            if ym >= cur:
                continue
            dl = teslim_tarihi(ym, gun)
            if today < date(dl.year, dl.month, oto):
                continue
            rows = self.db.month(ym)
            yeni = [r for r in rows if not r["gonderildi"]]
            ek_no = self.sonraki_ek(rows) if len(yeni) < len(rows) else 0
            try:
                self.paket_gonder(ym, yeni, ek_no)
                sonuc.append(f"{ay_adi(ym)}{f' EK {ek_no}' if ek_no else ''}: {len(yeni)} fiş")
            except Exception as e:
                hata.append(f"{ay_adi(ym)}: {e}")
        if sonuc and not hata:
            self.s["son_oto_deneme"] = today.isoformat()
            save_settings(self.s)
        if sonuc:
            self.status.setText("📤 Otomatik gönderildi → " + ", ".join(sonuc))
            QApplication.alert(self)
        if hata:
            self.oto_hata = "; ".join(hata)
        else:
            self.oto_hata = ""
        self.yenile()
        self.hatirlatma()

    def bulut_kontrol(self):
        if self.s.get("github_anahtar"):
            self.pool.start(BulutIs(self.s, self.db, self.bridge))
        else:
            self._bulut_durum("kapalı (Ayarlar → GitHub)")

    def _bulut_durum(self, t):
        self.bulut_txt = t
        if not self.status.text().startswith(("📱", "🧾", "➕", "✔")):
            self.yenile()

    def hatirlatma(self):
        today = date.today()
        gun = int(self.s.get("teslim_gunu", 14))
        oto = min(int(self.s.get("otomatik_gun", 10)), gun)
        cur = today.strftime("%Y-%m")
        satirlar, acil = [], False
        for ym, n in self.db.unsent_months():
            if ym >= cur:
                continue
            dl = teslim_tarihi(ym, gun)
            kal = (dl - today).days
            if kal < 0:
                acil = True
                satirlar.append(f"🔴 <b>{ay_adi(ym)}</b>: {n} fiş gönderilmedi — teslim günü "
                                f"<b>{dl:%d.%m.%Y}</b> GEÇTİ. Hemen gönderin.")
            else:
                acil = acil or kal <= 3
                oto_txt = ""
                if self.s.get("otomatik_gonder"):
                    oto_txt = (f" · otomatik gönderim {date(dl.year, dl.month, oto):%d.%m.%Y}"
                               if self.smtp_hazir() else " · ⚠ otomatik gönderim için Ayarlar'da e-posta bilgilerini girin")
                satirlar.append(f"⏰ <b>{ay_adi(ym)}</b>: {n} fiş — son teslim <b>{dl:%d.%m.%Y}</b> "
                                f"({'BUGÜN' if kal == 0 else f'{kal} gün kaldı'}){oto_txt}")
        if getattr(self, "oto_hata", ""):
            acil = True
            satirlar.append(f"❌ Otomatik gönderim başarısız: {self.oto_hata}")
        if satirlar:
            self.banner.setText("<br>".join(satirlar))
            self.banner.setStyleSheet("background:#5c1a22;border:1px solid #e0485a;" if acil else "")
            self.banner.setVisible(True)
        else:
            self.banner.setVisible(False)

    # ------------------------------------------------ diğer
    def telefon(self):
        TelefonDialog(self, self.sunucu, self.s).exec()

    def ayarlar(self):
        eski_port = self.s["telefon_port"]
        if AyarlarDialog(self, self.s).exec() == QDialog.Accepted:
            if self.s["telefon_port"] != eski_port:
                self.sunucu.stop()
                self.sunucu.start()
            self.bulut_kontrol()
            self.hatirlatma()
        self.yenile()

    def hakkinda(self):
        QMessageBox.about(self, "Hakkında",
                          f"<h2>{APP_NAME}</h2>Sürüm {APP_VERSION}<br><br>"
                          f"Gider fişlerini telefondan toplar, sıralar ve her ay muhasebeciye gönderir.<br><br>"
                          f"<b>Sahibi:</b> {OWNER}<br>"
                          f"<b>Web:</b> <a href='{WEBSITE}'>www.sinetechnic.com</a><br>"
                          f"<b>E-posta:</b> <a href='mailto:{EMAIL}'>{EMAIL}</a><br>"
                          f"<b>Discord:</b> <a href='{DISCORD}'>{DISCORD}</a>")

    def closeEvent(self, e):
        self.sunucu.stop()
        super().closeEvent(e)


def ocr_oz_test(img_path, out_path):
    """Derlenmiş exe'nin tarama+OCR'ı içerdiğini doğrulamak için (CI): gider.exe --ocr-test fis.jpg sonuc.txt"""
    try:
        sc, bulundu = belge_tara(Image.open(img_path))
        p = os.path.join(ROOT, "_ocr_test.jpg")
        sc.save(p)
        metin = ocr_metin(p)
        alan = fis_ayristir(metin)
        sonuc = f"bulundu={bulundu}\n{json.dumps(alan, ensure_ascii=False)}\n---\n{metin}"
    except Exception as e:
        sonuc = f"HATA {e!r}"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(sonuc)


def main():
    if len(sys.argv) >= 4 and sys.argv[1] == "--ocr-test":
        ocr_oz_test(sys.argv[2], sys.argv[3])
        return
    if os.name == "nt":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Sinetechnic.Gider")
        except Exception:
            pass
    QLocale.setDefault(QLocale(QLocale.Turkish, QLocale.Turkey))
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLE)
    w = Main()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
