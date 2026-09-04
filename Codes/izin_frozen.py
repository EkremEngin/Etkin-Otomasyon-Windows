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


# ==========================================================================
# Kullanıcı klasörleri — Masaüstü / İndirilenler
# ==========================================================================
# 🔴 WINDOWS TUZAĞI: os.path.expanduser("~/Desktop") mac'te hep doğrudur, WINDOWS'TA ÇOĞU ZAMAN YOKTUR.
# OneDrive kuruluysa Masaüstü C:\Users\<ad>\OneDrive\Desktop'a taşınır. Var olmayan bir klasörü dosya
# diyaloğuna `initialdir` olarak verirsen Windows'un yerel diyaloğu BOŞ açılır — kullanıcı "hiçbir
# dosya görünmüyor, klasörler boş" der. (2026-09-04'te izin belgesi seçiminde tam olarak bu yaşandı.)
#
# Doğrusu Windows'a SORMAKTIR: kabuk klasörleri kayıt defterinde tutulur ve yönlendirme yapıldığında
# oradaki değer de güncellenir. macOS/Linux'ta bu fonksiyonlar sessizce ~/Desktop'a düşer.

_WIN_MASAUSTU = "Desktop"
_WIN_INDIRILENLER = "{374DE290-123F-4565-9164-39C4925E467B}"   # İndirilenler'in bilinen-klasör GUID'i
_WIN_BELGELER = "Personal"                                     # Belgeler kayıt defterinde böyle geçer


def _win_kabuk_klasoru(anahtar: str) -> "str | None":
    """Windows kayıt defterinden gerçek kabuk klasörü yolu. Windows dışında / okunamazsa None."""
    if not sys.platform.startswith("win"):
        return None
    try:
        import winreg
    except Exception:  # noqa
        return None
    # "User Shell Folders" ham değeri tutar (%USERPROFILE%\... gibi), "Shell Folders" çözülmüşünü.
    for alt in (r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, alt) as k:
                yol = os.path.expandvars(winreg.QueryValueEx(k, anahtar)[0])
                if os.path.isdir(yol):
                    return yol
        except OSError:
            continue
    return None


def _ilk_var_olan(*adaylar) -> "str | None":
    """Adaylardan diskte GERÇEKTEN var olan ilk klasör. Hiçbiri yoksa None."""
    for a in adaylar:
        if a and os.path.isdir(a):
            return a
    return None


def masaustu() -> str:
    """Kullanıcının Masaüstü klasörü. Bulunamazsa ev klasörü — asla var olmayan bir yol dönmez."""
    od = (os.environ.get("OneDrive") or os.environ.get("OneDriveConsumer")
          or os.environ.get("OneDriveCommercial"))
    return _ilk_var_olan(
        _win_kabuk_klasoru(_WIN_MASAUSTU),
        # ⚠ Türkçe Windows'ta OneDrive altındaki klasörün FİZİKSEL adı "Masaüstü"dür ("Desktop" değil);
        #   yalnız İngilizcesini arayınca bu yedek hiç tutmuyordu.
        os.path.join(od, "Masaüstü") if od else None,
        os.path.join(od, "Desktop") if od else None,
        os.path.expanduser("~/Masaüstü"),
        os.path.expanduser("~/Desktop"),
    ) or os.path.expanduser("~")


def belgeler() -> str:
    """Kullanıcının Belgeler klasörü. E-postayla gelen Excel'in Türkçe Windows'ta en olası yeri."""
    od = (os.environ.get("OneDrive") or os.environ.get("OneDriveConsumer")
          or os.environ.get("OneDriveCommercial"))
    return _ilk_var_olan(
        _win_kabuk_klasoru(_WIN_BELGELER),
        os.path.join(od, "Belgeler") if od else None,
        os.path.join(od, "Documents") if od else None,
        os.path.expanduser("~/Belgeler"),
        os.path.expanduser("~/Documents"),
    ) or os.path.expanduser("~")


def indirilenler() -> str:
    """Kullanıcının İndirilenler klasörü. Bulunamazsa ev klasörü."""
    return _ilk_var_olan(
        _win_kabuk_klasoru(_WIN_INDIRILENLER),
        os.path.expanduser("~/Downloads"),
        os.path.expanduser("~/İndirilenler"),
    ) or os.path.expanduser("~")


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
    # macOS .app: .../Etkin Otomasyon.app/Contents/MacOS/Etkin Otomasyon → paketin DIŞINA çık
    suffix = os.sep + os.path.join("Contents", "MacOS")
    if exe_dir.endswith(suffix):
        exe_dir = os.path.dirname(exe_dir[: -len(suffix)])  # .app'i içeren klasör
    return exe_dir
