"""GitHub Actions'ta çalışır: telefondan gelen 'muhasebeciye gönder' komutunu işler.

Kullanım: python bulut_gonder.py <veri_deposu_klasoru> <komut.json>
Ortam: SMTP_SIFRE (GitHub secret)
"""
import json
import os
import shutil
import sys
import tempfile
import traceback
from datetime import datetime

VERI = os.path.abspath(sys.argv[1])
KOMUT = os.path.abspath(sys.argv[2])
os.environ["HOME"] = os.environ["USERPROFILE"] = tempfile.mkdtemp()
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import gider  # noqa: E402

GELEN_ID_BAS = 1_000_000


class _Sinyal:
    def emit(self, *a):
        pass


class _Kopru:
    ocr_bitti = _Sinyal()


def oku_json(p, varsayilan):
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return varsayilan


def yaz_json(p, veri):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(veri, f, ensure_ascii=False, indent=1)


def sonuc_yaz(ok, mesaj, **ek):
    ad = os.path.basename(KOMUT)
    yaz_json(os.path.join(VERI, "komut_sonuc", ad),
             dict(ok=ok, mesaj=mesaj, zaman=datetime.now().isoformat(timespec="seconds"), **ek))
    print(("OK: " if ok else "HATA: ") + mesaj)


def main():
    komut = oku_json(KOMUT, {})
    ym = komut.get("ay", "")
    if len(ym) != 7:
        return sonuc_yaz(False, "Komutta ay bilgisi yok")
    s = dict(gider.DEFAULT_SETTINGS)
    s.update(oku_json(os.path.join(VERI, "ayar", "gonderim.json"), {}))
    s["smtp_sifre"] = os.environ.get("SMTP_SIFRE", "")
    if komut.get("alici"):
        s["muhasebeci_eposta"] = komut["alici"]
    gider.save_settings(s)
    if not s["smtp_sifre"]:
        return sonuc_yaz(False, "GitHub'da SMTP_SIFRE secret'ı eklenmemiş (depo → Settings → Secrets → Actions).")
    if not s.get("muhasebeci_eposta") or not s.get("smtp_kullanici"):
        return sonuc_yaz(False, "E-posta ayarları buluta henüz yüklenmemiş. PC programında Ayarlar'ı doldurup bir kez açık bırakın.")

    arsiv = os.path.join(VERI, "arsiv", ym)
    gonderim_p = os.path.join(arsiv, "gonderim.json")
    gonderim = oku_json(gonderim_p, {"gonderimler": []})
    giden_id = {i for g in gonderim["gonderimler"] for i in g.get("kayit_idler", [])}
    giden_ad = {a for g in gonderim["gonderimler"] for a in g.get("gelen_adlar", [])}
    db = gider.DB(gider.DB_PATH)

    # 1) PC'nin işleyip arşive koyduğu fişler
    for k in oku_json(os.path.join(arsiv, "kayitlar.json"), []):
        src = os.path.join(VERI, k.get("dosya") or "")
        if not k.get("dosya") or not os.path.exists(src):
            continue
        hedef_klasor = os.path.join(gider.ROOT, "fisler", ym)
        os.makedirs(hedef_klasor, exist_ok=True)
        dst = os.path.join(hedef_klasor, os.path.basename(src))
        shutil.copy2(src, dst)
        kw = {a: k.get(a) for a in gider.KAYIT_ALANLAR if a in k}
        if kw.get("id") in giden_id:
            kw["gonderildi"] = kw.get("gonderildi") or "bulut"
        kw["dosya"] = os.path.relpath(dst, gider.ROOT)
        kw["tarih_kesin"] = 1
        db.add(**kw)

    # 2) Telefondan gelip PC'ye henüz inmemiş fişler (bu ayınkiler)
    gelen_klasor = os.path.join(VERI, "gelen")
    gelen_map = {}
    if os.path.isdir(gelen_klasor):
        nid = GELEN_ID_BAS
        for ad in sorted(os.listdir(gelen_klasor)):
            ext = os.path.splitext(ad)[1].lower()
            if ext not in gider.ALL_EXT or ad in giden_ad:
                continue
            meta = gider.meta_coz(ad)
            tarih = meta.get("tarih") or ""
            if tarih[:7] != ym:
                continue
            with open(os.path.join(gelen_klasor, ad), "rb") as f:
                rel = gider.dosya_kaydet(f.read(), ext, tarih, True)
            nid += 1
            gider.outer_add(db, rel, tarih, id=nid, kaynak="Telefon (bulut)", kaynak_adi=ad,
                            kategori=meta.get("kategori") or "Diğer", odeme=meta.get("odeme") or "Kredi Kartı",
                            plaka=(meta.get("plaka") or "").upper(), aciklama=meta.get("not") or "",
                            tutar=gider.sayi_oku(str(meta.get("tutar") or "")) or 0,
                            tarih_kesin=1 if meta.get("tarih_kesin") else 0)
            gider.OcrIs(db, nid, _Kopru()).run()
            gelen_map[nid] = ad

    tum = db.month(ym)
    rows = [r for r in tum if not r["gonderildi"]]
    if not rows:
        return sonuc_yaz(False, f"{gider.ay_adi(ym)} için gönderilmemiş fiş yok.")
    onceden = any(r["gonderildi"] for r in tum) or bool(gonderim["gonderimler"])
    ek_no = (max([g.get("ek_no") or 0 for g in gonderim["gonderimler"]] + [0]) + 1) if onceden else 0
    paket = gider.paket_olustur(db, ym, s, rows, ek_no)
    gider.eposta_gonder(s, paket, ym, [paket["xlsx"], paket["pdf"]])

    gonderim["gonderimler"].append({
        "zaman": datetime.now().isoformat(timespec="seconds"), "ek_no": ek_no, "kaynak": "telefon",
        "kayit_idler": [r["id"] for r in rows if r["id"] < GELEN_ID_BAS],
        "gelen_adlar": [gelen_map[r["id"]] for r in rows if r["id"] in gelen_map],
        "adet": paket["adet"], "toplam": round(paket["toplam"], 2)})
    yaz_json(gonderim_p, gonderim)
    etiket = f" EK {ek_no}" if ek_no else ""
    sonuc_yaz(True, f"{gider.ay_adi(ym)}{etiket}: {paket['adet']} belge, toplam {gider.tl(paket['toplam'])} "
                    f"→ {s['muhasebeci_eposta']} adresine gönderildi.",
              adet=paket["adet"], toplam=round(paket["toplam"], 2), ek_no=ek_no)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        traceback.print_exc()
        sonuc_yaz(False, f"Gönderim hatası: {e}")
    finally:
        try:
            os.remove(KOMUT)
        except OSError:
            pass
