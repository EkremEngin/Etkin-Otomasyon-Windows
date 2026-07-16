# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec — İzin Otomasyonu tek-dosya (.exe / mac binary).
Build:  pyinstaller izin_app.spec        (Windows'ta .exe, macOS'ta unix binary üretir)
Çıktı:  dist/IzinOtomasyonu(.exe)

Not: Playwright KENDİ TARAYICISINI bundle ETMEZ — kullanıcının Chrome'una CDP ile bağlanır.
Bu yüzden `playwright install` (tarayıcı indirme) GEREKMEZ; sadece playwright PAKETİ + node driver bundle edilir.
"""
from PyInstaller.utils.hooks import collect_all
import sys

# playwright'ın node driver'ı + veri dosyaları (connect_over_cdp bunu ister)
datas, binaries, hiddenimports = collect_all('playwright')
ctk_datas, ctk_binaries, ctk_hiddenimports = collect_all('customtkinter')
datas += ctk_datas
binaries += ctk_binaries
hiddenimports += ctk_hiddenimports
datas += [
    ('assets/etkn-logo.png', 'assets'),
    ('assets/etkn-logo-hd.png', 'assets'),
    ('assets/etkn-app-icon.png', 'assets'),
    ('assets/etkn-app-icon.ico', 'assets'),   # Windows taskbar: çok-boyutlu ikon (iconbitmap runtime'da okur)
]

# exe dosya ikonu: Windows .ico ister (Explorer + taskbar net); macOS .png/.icns kullanır.
app_icon = 'assets/etkn-app-icon.ico' if sys.platform.startswith('win') else 'assets/etkn-app-icon.png'

# izin_app dinamik (fonksiyon içi) import ettiği için modülleri açıkça bildir
hiddenimports += [
    'izin_gui', 'izin_otomasyon', 'izin_onaya_dosyali_yildiz', 'izin_login_check',
    'izin_poc', 'izin_onaya', 'izin_belge', 'izin_data_v2', 'izin_frozen',
    'openpyxl', 'PIL', 'customtkinter', 'darkdetect', 'tkinter', 'tkinter.filedialog',
    'tkinter.messagebox',
]

# DGS modülü. dgs_park.main() park scriptini importlib ile ÇALIŞMA ANINDA yüklüyor (statik import yok)
# → PyInstaller bunları göremez, tek tek bildirilmeli. Biri eksikse o park exe'de ModuleNotFoundError verir.
hiddenimports += [
    'dgs_park',                                                       # tek giriş noktası (--park <KOD>)
    'dgs_tpi', 'dgs_bv', 'dgs_tpiz', 'dgs_yildiz', 'dgs_ulutek',      # park scriptleri (overrides taşır)
    'dgs_poc', 'dgs_onaya', 'dgs_rapor_kontrol',                      # ortak motor
]

a = Analysis(
    ['izin_app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Etkin Otomasyon',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,          # GUI penceresi (siyah konsol yok). Worker'lar stdio-repair ile PIPE'a yazar.
    disable_windowed_traceback=False,
    argv_emulation=False,   # mac: dosya-sürükle argv emülasyonu gerekmiyor
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=app_icon,
)

# macOS'ta ham Mach-O dosyası Finder'da uygulama gibi davranmaz. Aynı build'den
# çift tıklanabilir .app paketi de çıkar; Windows build'lerinde yalnız .exe üretilir.
if sys.platform == "darwin":
    app = BUNDLE(
        exe,
        name='Etkin Otomasyon.app',
        bundle_identifier='com.fev.izin',
        icon='assets/etkn-app-icon.png',
        info_plist={
            'CFBundleDisplayName': 'Etkin Otomasyon',
            'CFBundleShortVersionString': '1.0.0',
            'NSHighResolutionCapable': True,
        },
    )
