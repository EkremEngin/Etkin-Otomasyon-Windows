#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ETKİN OTOMASYON — PyInstaller TEK GİRİŞ NOKTASI (dispatcher).

Bu dosya .exe/.app'in ana giriş noktasıdır. İlk argümana (token) göre yönlendirir:
  • token YOK           → GUI aç (izin_gui — İzin | DGS iki modül)
  • "orchestrator" ...  → İZİN giriş+onay motoru (izin_otomasyon.main)
  • "uploader" ...      → İZİN belge yükle+onay worker'ı (izin_onaya_dosyali_yildiz.main)
  • "logincheck" [cdp]  → login/park probe (izin_login_check.main) — İZİN ve DGS ORTAK kullanır
  • "dgs" --park <KOD>  → DGS giriş/onay/kontrol (dgs_park.main → park modülü → ortak motor)
  • "paths"             → teşhis: resume/done dosyalarının YAZILDIĞI klasörü basar (JSON)

GUI ve orkestratör alt-süreçlerini `izin_frozen.worker_cmd()` ile açtığından, donmuş exe kendini
bu token'larla çağırır. Script modunda (geliştirme) doğrudan `python3 izin_gui.py` da çalışır.
"""
from __future__ import annotations

import os
import sys

_TOKENS = ("orchestrator", "uploader", "logincheck", "dgs", "paths")


def _repair_stdio():
    """Windows'ta windowed (console=False) PyInstaller exe'de worker stdio'su ÜÇ türlü bozulur:
      1) Çift-tıkla açılan GUI'de konsol yok → sys.stdout/stderr = None → print() AttributeError.
      2) GUI worker'ı PIPE ile açınca fd 1/2 GEÇERLİDİR, ama Python stream'i sistem ANSI kod sayfasıyla
         (Türkçe/Batı Windows'ta cp1254/cp1252) kurar → "ş/İ/ö" kodlanamaz → UnicodeEncodeError, worker çöker.
      3) Donmuş exe PYTHONUNBUFFERED'ı ONURLANDIRMAZ → stdout BLOK-tamponlu → satırlar iş bitene kadar
         PIPE'a düşmez → GUI'de "canlı işlem akışı" görünmez (hepsi sonda toplu gelir).
    Çözüm: stream'i UTF-8 + SATIR-tamponuna sabitle (her \n'de flush). GUI okuma tarafını zaten utf-8 açıyor
    (izin_gui Popen encoding='utf-8'). (Kanıt: donmuş 'streamtest'te encoding-only ilk→son aralık 0.00s =
    tamponlu; line_buffering ile ~2s = canlı. macOS/Unix'te varsayılan utf-8+satır; bu onarım orada zararsız.)"""
    for name, fd in (("stdout", 1), ("stderr", 2)):
        stream = getattr(sys, name, None)
        if stream is None:                                  # (1) konsol yok → fd 1/2'den utf-8 + satır-tamponu ile kur
            try:
                setattr(sys, name, os.fdopen(fd, "w", buffering=1, encoding="utf-8", errors="replace"))
            except Exception:
                try:
                    setattr(sys, name, open(os.devnull, "w"))
                except Exception:
                    pass
        else:                                               # (2)+(3) cp125x→utf-8 + satır-tamponu (canlı akış)
            try:
                stream.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
            except Exception:
                pass


def main():
    _repair_stdio()
    if len(sys.argv) >= 2 and sys.argv[1] in _TOKENS:
        mode = sys.argv.pop(1)                              # token'ı düş → alt argparse normal görür
        if mode == "orchestrator":
            import izin_otomasyon
            izin_otomasyon.main()
        elif mode == "uploader":
            import izin_onaya_dosyali_yildiz
            izin_onaya_dosyali_yildiz.main()
        elif mode == "dgs":
            import dgs_park                                  # --park <KOD>'u kendi ayrıştırır
            dgs_park.main()
        elif mode == "paths":
            _print_paths()
        else:
            import izin_login_check
            izin_login_check.main()
    else:
        import izin_gui
        izin_gui.main()


def _print_paths():
    """Teşhis: resume/done dosyaları NEREYE yazılıyor?

    Motorlar done dosyasını göreli yolla açıyor → sürecin CWD'sine düşer; GUI o CWD'yi
    izin_frozen.data_dir()'den veriyor. Donmuş exe'de __file__ tabanlı SCRIPT_DIR, PyInstaller'ın
    _MEIPASS GEÇİCİ klasörüdür (çıkışta silinir) → resume oraya yazılırsa uçar, mükerrer kayıt olur.
    Bu token ikisini yan yana basar; data_dir, meipass'e EŞİT ÇIKMAMALI.
    """
    import json
    import izin_frozen
    meipass = getattr(sys, "_MEIPASS", None)
    data = izin_frozen.data_dir()
    print(json.dumps({
        "frozen": izin_frozen.is_frozen(),
        "data_dir": data,                       # ← resume/done/istisna dosyaları BURAYA yazılır
        "script_dir": izin_frozen.SCRIPT_DIR,   # ← __file__ tabanlı (donmuşta _MEIPASS = geçici!)
        "meipass": meipass,
        "executable": sys.executable,
        "data_dir_gecici_mi": bool(meipass and os.path.abspath(data) == os.path.abspath(meipass)),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
