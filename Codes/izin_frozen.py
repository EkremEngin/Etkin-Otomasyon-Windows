#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Donmuş (.exe/.app) uyumlu alt-süreç başlatma yardımcısı.

Sorun: PyInstaller ile paketlenince `sys.executable` artık python değil, UYGULAMANIN KENDİSİ; ve
diskte `izin_otomasyon.py` gibi bir dosya YOK (hepsi exe içine gömülü). Bu yüzden GUI'nin/orkestratörün
`[sys.executable, "izin_otomasyon.py", ...]` şeklinde alt-süreç açması donmuş modda ÇALIŞMAZ
(exe'yi tekrar GUI olarak açar). Çözüm: donmuş modda exe kendini bir TOKEN ile çağırır
(`izin_app.py` bu token'a göre doğru worker'a yönlendirir); script modunda normal .py çağrılır.

worker_cmd(mode) → alt-süreç komutunun BAŞI döner; çağıran sonuna argümanları ekler.
"""
from __future__ import annotations

import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# mode token → script modunda çağrılacak .py
_SCRIPTS = {
    "orchestrator": "izin_otomasyon.py",
    "uploader": "izin_onaya_dosyali_yildiz.py",
    "logincheck": "izin_login_check.py",
    "dgs": "dgs_park.py",          # tek DGS giriş noktası; park `--park <KOD>` ile seçilir
}


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def worker_cmd(mode: str) -> list:
    if mode not in _SCRIPTS:
        raise ValueError(f"bilinmeyen worker modu: {mode!r}")
    if is_frozen():
        return [sys.executable, mode]                       # exe kendini token'la çağırır
    return [sys.executable, os.path.join(SCRIPT_DIR, _SCRIPTS[mode])]


def data_dir() -> str:
    """Resume/done dosyalarının (dgs_done_*.txt · izin_done_*.txt) yazıldığı KALICI klasör.

    🔴 NEDEN VAR: motorlar done dosyasını ÇIPLAK GÖRELİ yolla açıyor (dgs_poc.py:
    `done_file = f"dgs_done_{lokasyon}_{sheet}.txt"`) → dosya sürecin CWD'sine düşer. Donmuş exe'de
    SCRIPT_DIR (=__file__'ın klasörü) PyInstaller'ın _MEIPASS GEÇİCİ klasörüdür ve süreç biter bitmez
    SİLİNİR. Oraya cwd verilirse resume dosyaları her koşuda buharlaşır → ikinci koşu kimseyi atlamaz
    → MÜKERRER KAYIT. Bu yüzden donmuş modda EXE'NİN DURDUĞU klasör kullanılır.

    Kural (operatöre anlatılacak hali): "exe nerede duruyorsa resume dosyaları da orada."
    ETKN_DATA_DIR env değişkeni her şeyi ezer (kaçış kapısı).
    """
    env = os.environ.get("ETKN_DATA_DIR")
    if env:
        return os.path.abspath(os.path.expanduser(env))
    if not is_frozen():
        return SCRIPT_DIR                                   # geliştirme: proje klasörü (mevcut davranış)
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    # macOS .app: .../IzinOtomasyonu.app/Contents/MacOS/IzinOtomasyonu → paketin DIŞINA çık
    suffix = os.sep + os.path.join("Contents", "MacOS")
    if exe_dir.endswith(suffix):
        exe_dir = os.path.dirname(exe_dir[: -len(suffix)])  # .app'i içeren klasör
    return exe_dir
