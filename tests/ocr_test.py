"""CI kontrolü: tarama + OCR + ayrıştırma Windows'ta çalışıyor mu?"""
import os, sys, tempfile
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ["USERPROFILE"] = os.environ["HOME"] = tempfile.mkdtemp()
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import gider
from PIL import Image

src = os.path.join(os.path.dirname(__file__), "ornek_fis.jpg")
sc, bulundu = gider.belge_tara(Image.open(src))
print("fiş bulundu:", bulundu, sc.size)
p = os.path.join(gider.ROOT, "t.jpg"); sc.save(p)
for ad, fn in (("rapidocr", gider._rapid_ocr), ("winocr", gider._win_ocr)):
    try:
        t = fn(sc.convert("RGB")); print(f"--- {ad}:\n{t}")
    except Exception as e:
        print(f"--- {ad} HATA: {e!r}")
metin = gider.ocr_metin(p)
alan = gider.fis_ayristir(metin)
print("AYRIŞTIRMA:", alan)
assert bulundu, "fiş bulunamadı"
assert alan.get("tutar") == 320.0, "tutar okunamadı"
assert alan.get("tarih") == "2026-10-30", "tarih okunamadı"
print("OK")
