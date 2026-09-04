#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETKİN OTOMASYON — Masaüstü Arayüzü (öğretici, İKİ MODÜLLÜ)
==========================================================
Python bilmeyen bir kullanıcının (Begüm/Tolga) işi çift-tıkla yapabilmesi için tkinter tabanlı arayüz.
PyInstaller ile Windows .exe ve macOS .app olarak paketlenir.

MODÜLLER (üstteki İZİN | DGS seçicisi):
  • İZİN → yıllık izin girişi + belge yükleme + onay        (motor: izin_otomasyon.py)
  • DGS  → gelir vergisi istisnası puantaj girişi + onay     (motor: dgs_park.py → dgs_tpi/bv/... → dgs_poc)

Akış (her iki modülde de aynı): park → Chrome/login → Excel → kapsam → çalıştır (canlı log).

Motor DOSYALARINA DOKUNMAZ — hepsini ALT-SÜREÇ olarak çağırır (playwright'ın sync API'si GUI thread'ini
bozar; ayrıca dgs_poc/dgs_onaya PRIME scriptlerdir, bkz. PRIME_SCRIPTS.md).
"""
from __future__ import annotations   # Python 3.9 uyumu: 'str | None' annotation'ları ertele

import os
import re
import sys
import json
import queue
import shutil
import datetime
import threading
import subprocess
import urllib.parse
import urllib.request

import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image

# --- Portalsız motor modülleri (playwright YÜKLEMEZ; arayüz ana thread'inde güvenli) ---
import izin_data_v2 as data
import izin_belge
import izin_frozen   # donmuş (.exe) uyumlu alt-süreç komutları + kalıcı resume klasörü
import dgs_park      # DGS park kayıt defteri (dgs_poc'u SADECE run() içinde import eder → playwright gelmez)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Resume/done dosyalarının yazıldığı KALICI klasör. Alt-süreçlere cwd olarak BU verilir — SCRIPT_DIR
# DEĞİL: donmuş exe'de SCRIPT_DIR geçici _MEIPASS'tir, oraya yazılan resume dosyaları uçar → mükerrer kayıt.
DATA_DIR = izin_frozen.data_dir()


def asset_path(relative_path: str) -> str:
    """Kaynak kodda ve PyInstaller paketinde aynı marka varlığına ulaş."""
    return os.path.join(getattr(sys, "_MEIPASS", SCRIPT_DIR), relative_path)


BRAND_LOGO_PATH = asset_path(os.path.join("assets", "etkn-logo-hd.png"))
APP_ICON_PATH = asset_path(os.path.join("assets", "etkn-app-icon.png"))
APP_ICO_PATH = asset_path(os.path.join("assets", "etkn-app-icon.ico"))   # Windows: çok-boyutlu .ico → taskbar NET
PROFILE_DIR = os.path.expanduser("~/dgs-chrome-profile")
DEBUG_PORT = 9222
CDP_URL = f"http://localhost:{DEBUG_PORT}"

# Tk her iki platformda da yazı tipi adını kendisi çözer. Başlıkları sistemin
# kendi ailesiyle çizmek, uygulamanın Windows'ta ve macOS'ta doğal hissetmesini sağlar.
FONT_UI = "SF Pro Text" if sys.platform == "darwin" else "Segoe UI"
FONT_DISPLAY = "SF Pro Display" if sys.platform == "darwin" else "Segoe UI"
FONT_MONO = "SF Mono" if sys.platform == "darwin" else "Consolas"

# Canlı log yazı boyu.
# 🔴 WINDOWS'TA NEDEN DAHA BÜYÜK: uygulama DPI-farkında (main() içinde SetProcessDpiAwareness) →
# Windows onu KENDİ ölçeğiyle büyütmez, ölçeklemeyi biz yaparız. Ölçeğimiz ise tasarımı (1420x920)
# ekrana sığdırmaktan geliyor; %150 ölçekli 1080p bir dizüstünde bu ~1.0 çıkıyor, yani sistemin
# büyüttüğü diğer uygulamaların yanında yazı KÜÇÜK kalıyor. Mono 10 punto orada okunmuyordu.
# Kullanıcı log penceresindeki A− / A+ ile değiştirebilir; seçim gui_ayarlar.json'a yazılır.
LOG_YAZI_VARSAYILAN = 12 if sys.platform.startswith("win") else 10
LOG_YAZI_ALT, LOG_YAZI_UST = 8, 26
AYAR_DOSYASI = os.path.join(DATA_DIR, "gui_ayarlar.json")

# Dosya diyaloğu uzantı filtreleri.
# ⚠ Desen listesi TUPLE olarak verilir — tek string içinde boşlukla ayırmak ("*.xlsx *.xlsm")
#   platformlar arasında güvenilir değil. Büyük harfli karşılıkları da yazılı: macOS/Linux'ta Tk
#   eşleştirmeyi KENDİ yapar ve büyük/küçük harf duyarlıdır; "RAPOR.XLSX" aksi halde listelenmez.
# ⚠ Her filtreye "Tüm dosyalar" eklenir: filtre beklenmedik bir şey yaparsa kullanıcı yine de
#   dosyasını görebilsin — kilitlenip "hiçbir şey görünmüyor" durumuna düşmesin.
EXCEL_TIPLERI = [("Excel dosyası", ("*.xlsx", "*.xlsm", "*.XLSX", "*.XLSM")),
                 ("Tüm dosyalar", "*.*")]
PDF_TIPLERI = [("PDF belgesi", ("*.pdf", "*.PDF")), ("Tüm dosyalar", "*.*")]

# 🔴 SÜRÜM DAMGASI — exe'nin İÇİNE gömülür. Depodaki 'Etkin Otomasyon.exe' bir BUILD ÇIKTISIDIR:
# kodu güncelleyip depoyu yenilemek onu DEĞİŞTİRMEZ, yeniden build edilene kadar eski kodu çalıştırır.
# Bir kez "yeni sürümü indirdim ama hiçbir şey değişmemiş" diye vakit kaybedildi (2026-09-04).
# Bu damga arayüzün üst şeridinde ve log'un ilk satırında görünür → hangi build olduğu belli olur.
SURUM = "2026-09-04"

def _domain(url: str) -> str:
    return urllib.parse.urlsplit(url or "").netloc.lower()


# ==========================================================================
# GÖRÜNTÜ KATMANI — İZİN ve DGS sekmelerinde kullanıcıya AYNI teknokent adı görünsün.
# ⚠ İŞLEYİŞE ETKİSİ YOK: bunlar YALNIZ combobox/chip metnidir. İçeride park hâlâ
#   kendi koduyla çözülür (--park argümanı, resume dosya adları, Excel okuma, portal
#   karşılaştırmaları hepsi eskisiyle BİREBİR aynı). Anahtar = portal domaini (iki
#   modülde de ortak olan tek şey; kodlar farklı olduğu için kod anahtar OLARAK KULLANILMAZ).
# ==========================================================================
PARK_DISPLAY = {
    "argeportal.teknoparkistanbul.com.tr": "Teknopark İstanbul",
    "argeportal.bilisimvadisi.com.tr":     "Bilişim Vadisi",
    "argeportal.teknoparkizmir.com.tr":    "Teknopark İzmir",
    "argeportal.yildizteknopark.com.tr":   "Yıldız Teknopark",
    "argeportal.ulutek.com.tr":            "Ulutek Teknopark",
}


def park_display(p) -> str:
    """Park (İZİN ya da DGS) → kullanıcıya görünen ORTAK ad. Salt görüntü; kod içeride sabit kalır."""
    return PARK_DISPLAY.get(_domain(p.portal_url), p.ad)


# Otomasyon kapsamındaki parklar (Teknoera olan ARI/ODTÜ hariç)
# --- YEDEK (eski hali — combobox İZİN park KODUNU gösteriyordu): -------------
# AUTO_PARKS = [c for c, p in data.PARKS.items() if p.otomasyon]
# ----------------------------------------------------------------------------
AUTO_PARKS = [park_display(p) for p in data.PARKS.values() if p.otomasyon]
# Görünen ad → İZİN park kodu (combobox seçimini içerideki koda çevirir; işleyiş değişmez)
IZIN_DISPLAY_TO_CODE = {park_display(p): c for c, p in data.PARKS.items() if p.otomasyon}

MODE_TR = {
    "": ("PDF'siz", "Bu park onay için BELGE İSTEMEZ — giriş sonrası doğrudan onaya gider."),
    "per_person": ("Kişi-başı belge", "Her personelin KENDİ izin formu (PDF) gerekir."),
    "ortak": ("Ortak tek belge", "Firma TÜM çalışanları tek antetli/kaşeli yazıyla (tek PDF) yollar."),
}


# ==========================================================================
# DGS modülü — park eşlemesi ve dönem
# ==========================================================================
# ⚠ AYNI PORTAL, FARKLI KOD: İZİN kodları ≠ DGS kodları.
#       İYTE→TPIz · YTP→Yıldız · ULUTEK→Ulutek   (TPI ve BV aynı)
# DGS kodu = Excel'in "Lokasyonu SGK" (F) kolonundaki değer; resume dosya adlarında da o geçer.
# izin_login_check İZİN kodunu döndürdüğü için park karşılaştırması PORTAL URL'si üzerinden yapılır.
# --- YEDEK (eski hali — combobox DGS KODUNU gösteriyordu): -------------------
# DGS_PARKS = list(dgs_park.PARKS)
# ----------------------------------------------------------------------------
DGS_PARKS = [park_display(p) for p in dgs_park.PARKS.values()]   # kullanıcıya ortak ad; işleyiş aynı
# Görünen ad → DGS park kodu (combobox seçimini içerideki koda çevirir; içeride hâlâ TPIz/Yıldız/Ulutek)
DGS_DISPLAY_TO_CODE = {park_display(p): code for code, p in dgs_park.PARKS.items()}

DGS_MODLAR = {
    "Giriş + Onay": ("giris", "Excel'den puantajı girer ve onaya gönderir. NORMALDE HEP BUNU KULLAN."),
    "Sadece Onay": ("onay", "Yeni giriş yapmaz; önceki koşudan taslakta kalmış kayıtları onaya gönderir."),
    "SGK Kontrol": ("kontrol", "İş bittikten sonra son kontrol: SGK raporunda herkes 'Gün = 30' mu?"),
}

_TR_AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
             "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]


# _domain: yukarıda (GÖRÜNTÜ KATMANI'ndan önce) tanımlandı.
DGS_BY_PORTAL = {_domain(p.portal_url): code for code, p in dgs_park.PARKS.items()}


def dgs_donem_ay() -> str:
    """DGS dönemi = bir ÖNCEKİ ay (motorun otomatik türettiği kuralın aynısı).
    Yalnız GÖSTERİM için — gerçek dönemi dgs_poc türetir (burada dgs_poc import EDİLMEZ: playwright'ı
    arayüz sürecine sokmayalım)."""
    t = datetime.date.today()
    return _TR_AYLAR[(t.month - 2) % 12]


def dgs_done_sayisi(park_code: str, ay: str) -> int:
    """Resume dosyasındaki kişi sayısı — bunlar TEKRAR GİRİLMEZ (mükerrer kayıt kalkanı)."""
    try:
        with open(os.path.join(DATA_DIR, f"dgs_done_{park_code}_{ay}.txt"), encoding="utf-8") as fh:
            return sum(1 for satir in fh if satir.strip())
    except OSError:
        return 0


# --- İstisna kişiler: arama/eşleştirme yardımcıları ---------------------------
_TR_HARF = {'ç': 'c', 'Ç': 'c', 'ğ': 'g', 'Ğ': 'g', 'ı': 'i', 'İ': 'i', 'I': 'i',
            'ö': 'o', 'Ö': 'o', 'ş': 's', 'Ş': 's', 'ü': 'u', 'Ü': 'u'}


def _fold_tr(s: str) -> str:
    """Türkçe-duyarsız normalize — dgs_onaya.fold ile AYNI kural (motor da bununla eşliyor)."""
    s = (s or "").replace("̇", "")
    return " ".join("".join(_TR_HARF.get(c, c) for c in s).lower().split())


def _tc_rakam(s) -> str:
    return "".join(ch for ch in str(s or "") if ch.isdigit())


def _ist_key(k: dict) -> str:
    """İstisna kimlik anahtarı: T.C. varsa T.C. (evlilik/kızlık soyadına dayanıklı), yoksa ad."""
    return _tc_rakam(k.get("tc")) or _fold_tr(k.get("ad", ""))


# Motorun (dgs_poc) run sonu bastığı başarısızlık satırı: "  !! BAŞARISIZ: <ad> — <sebep>"
# (kaynak: dgs_poc.py:1465). GUI bunu canlı akıştan yakalayıp run bitiminde "tekrar dene" teklif eder.
_BASARISIZ_RE = re.compile(r"!!\s*BAŞARISIZ:\s*(.+?)\s*[—–]\s*(.+)")
# Canlı ilerleme: motor her kişinin başında "=== [13/38] AHMET YILMAZ (proje: …) ===" basar →
# GUI durum çubuğunda "İşleniyor: 13/38 · AHMET YILMAZ" göstermek için yakala.
_ILERLEME_RE = re.compile(r"===\s*\[(\d+)/(\d+)\]\s*(.+?)\s*===")
# Kapanış sonu: otomasyonun tamamlayamadığı (MANUEL giriş gereken) kişiler → run bitince popup'ta göster.
_MANUEL_RE = re.compile(r"<<<MANUEL>>>(.*)")


# Nötr kömür zemin + markanın turuncusu. Sıcak renk yalnız etkileşim ve vurgu için
# kullanılır; böylece uzun operasyonlarda ekran daha sakin, hiyerarşi daha nettir.
UI = {
    "bg": "#0A0C0F",
    "surface": "#12151A",
    "surface_raised": "#181C22",
    "card": "#161A20",
    "card_hover": "#222832",
    "input": "#0E1116",
    "border": "#29313C",
    "border_strong": "#3B4655",
    "text": "#F5F7FA",
    "muted": "#AAB3C1",
    "subtle": "#768193",
    "primary": "#FF6B1A",
    "primary_hover": "#FF843D",
    "accent_soft": "#2B1D16",
    "accent_text": "#FFB58A",
    "cyan": "#54C8FF",
    "success": "#39D69C",
    "warning": "#F5C761",
    "danger": "#FF6F78",
    "danger_soft": "#402126",
    "danger_hover": "#5A2A31",
}

# Canlı log renkleri — log penceresi ve ana penceredeki tek satırlık önizleme aynı sözlüğü kullanır.
_LOG_RENK = {"success": UI["success"], "warning": UI["warning"],
             "danger": UI["danger"], "command": UI["cyan"]}
# Bellekte tutulan en fazla log parçası (pencere kapalıyken de birikir). TAM kayıt gui_canli_log.txt'de.
LOG_TAVAN = 4000


# ==========================================================================
# Chrome başlatma (çapraz-platform)
# ==========================================================================
def find_chrome() -> str | None:
    if sys.platform == "darwin":
        cands = ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"]
    elif sys.platform.startswith("win"):
        cands = [
            os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
        ]
    else:
        cands = ["google-chrome", "google-chrome-stable", "chromium-browser", "chromium"]
    for c in cands:
        if os.path.sep in c or c.endswith(".exe"):
            if os.path.exists(c):
                return c
        else:
            w = shutil.which(c)
            if w:
                return w
    return None


def debug_port_up(timeout=1.5) -> bool:
    try:
        with urllib.request.urlopen(f"{CDP_URL}/json/version", timeout=timeout):
            return True
    except Exception:
        return False


def launch_chrome(portal_url: str) -> tuple[bool, str]:
    """Debug-portlu Chrome'u aç (zaten açıksa portal URL'sini yeni sekmede açar)."""
    chrome = find_chrome()
    if not chrome:
        return False, ("Google Chrome bulunamadı. Kurulu mu? (Windows'ta Program Files, "
                       "macOS'ta /Applications)")
    args = [chrome, f"--remote-debugging-port={DEBUG_PORT}",
            f"--user-data-dir={PROFILE_DIR}", portal_url]
    try:
        subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:  # noqa
        return False, f"Chrome başlatılamadı: {e}"
    return True, "Chrome açılıyor — pencerede login ol + Cloudflare'i geç."


# ==========================================================================
# Pencere ölçüsü — DPI hesabı (Windows'ta canlı log panelini kaçıran hata buradaydı)
# ==========================================================================
# Arayüz bu MANTIKSAL boyuta göre çizildi (kart genişlikleri, wraplength'ler, sağdaki kartlar).
# NOT: canlı log artık ana pencerede DEĞİL, ayrı pencerede (bkz. IzinGUI._build_log) — ana pencerenin
# yükseklik ihtiyacı eskisinden düşük. Alt pencereler ayrıca `_ust_pencere_plani` ile ekrana sığdırılır.
# TASARIM_H 920'den 820'ye ÇEKİLDİ (2026-09-04): canlı log ana pencereden çıkıp ayrı pencereye
# taşınınca burada o kadar yüksekliğe ihtiyaç kalmadı. Ölçek "tasarımı çalışma alanına sığdır"
# hesabından geliyor ve yaygın Windows ekranlarında darboğaz HEP yükseklikti — tasarım kısalınca
# ölçek yükseliyor, yani YAZILAR VE WIDGET'LAR BÜYÜYOR: 1080p@%150'de 1.02 → 1.14, 1366x768'de
# 0.73 → 0.82. ("Windows'ta yazılar aşırı küçük" şikâyeti.) 820, ölçülen asgari boyun (760) üstünde.
TASARIM_W, TASARIM_H = 1420, 820


def pencere_plani(sw: int, sh: int, dpi: float, win: bool) -> dict:
    """Ekrana SIĞAN pencere planı: ölçek çarpanı + mantıksal boyut + konum + minsize.

    🔴 WINDOWS TUZAĞI (canlı log görünmüyordu): CustomTkinter `geometry()` ve `minsize()`
    değerlerini ekranın DPI çarpanıyla ÇARPAR (Windows %150 ölçek → 1.5×). Eski kod ölçüyü
    FİZİKSEL pikselle veriyordu (1420x920) → pencere 2130x1380 açılıyor, sağ sütunun DİBİNDEKİ
    canlı log paneli ekranın/görev çubuğunun altında kalıyordu. minsize de (1100x760 → 1650x1140)
    şiştiği için kullanıcı pencereyi küçültüp log'u geri getiremiyordu. macOS'ta dpi=1 olduğundan
    sorun hiç görünmedi — bu yüzden yalnız Windows'ta çıktı.

    Çözüm: (1) çalışma alanı fiziksel piksel olarak hesaplanır (görev çubuğu + başlık payı düşülür),
    (2) tasarım oraya sığmıyorsa ÖLÇEK kısılır — yazı ve widget BİRLİKTE küçülür, oran bozulmaz,
    (3) geometry'ye ölçeğe BÖLÜNMÜŞ mantıksal değer verilir → fiziksel sonuç tam çalışma alanı kadar.
    """
    # Görev çubuğu / menü çubuğu + pencere başlığı payı — Tk çalışma alanını doğrudan vermiyor.
    if win:
        yatay_pay, dikey_pay = int(48 * dpi), int(96 * dpi)   # kenarlar + görev çubuğu + başlık
    else:
        yatay_pay, dikey_pay = 64, 120                        # menü çubuğu + Dock
    alan_w, alan_h = max(640, sw - yatay_pay), max(480, sh - dikey_pay)

    # Tasarım çalışma alanına sığmıyorsa ölçeği kıs (0.7 = okunabilirlik tabanı).
    olcek = max(0.7, min(dpi, alan_w / TASARIM_W, alan_h / TASARIM_H))

    genislik = int(min(TASARIM_W, alan_w / olcek))
    yukseklik = int(min(TASARIM_H, alan_h / olcek))
    return {
        "olcek": olcek,
        "genislik": genislik,
        "yukseklik": yukseklik,
        "x": max(0, (sw - int(genislik * olcek)) // 2),
        "y": max(0, (sh - int(yukseklik * olcek)) // 2),
        # Minsize de MANTIKSAL birimde (CTk ölçekle çarpacak). 760 tabanı ölçüldü: bunun altında
        # sağ sütun kısalıyor ve canlı log kartı kırpılmaya başlıyor.
        "min_w": int(min(1120, alan_w / olcek)),
        "min_h": int(min(680, alan_h / olcek)),   # tasarım kısaldı → taban da indi (bkz. TASARIM_H)
        # Ekran 0.7 tabanında bile dar: Windows'ta tam ekran aç → görev çubuğunu WM'in kendisi
        # hesaplar, alt paneller kesin kesilmez.
        "dar": win and (alan_w < TASARIM_W * 0.7 or alan_h < TASARIM_H * 0.7),
    }


# ==========================================================================
# Ana uygulama
# ==========================================================================
class IzinGUI:
    def __init__(self, root: ctk.CTk):
        self.root = root
        root.title("Etkin Otomasyon")
        self._fit_window()
        root.configure(fg_color=UI["bg"])
        # Pencere/taskbar ikonu. Windows'ta çok-boyutlu .ico + iconbitmap NET görünür; iconphoto tek dev
        # PNG'yi kaba ölçekleyip taskbar'da tırtıklı/soluk gösteriyordu. macOS/Linux'ta iconphoto kullanılır.
        self._app_icon = None
        _icon_ok = False
        if sys.platform.startswith("win"):
            try:
                root.iconbitmap(default=APP_ICO_PATH)   # default= → tüm pencereler/dialoglar da alır
                _icon_ok = True
            except Exception:
                _icon_ok = False
        if not _icon_ok:
            try:
                self._app_icon = tk.PhotoImage(file=APP_ICON_PATH)
                root.iconphoto(True, self._app_icon)
            except Exception:
                pass

        # --- İZİN modülü ---
        self.excel_path = tk.StringVar()
        self.belge_path = tk.StringVar()   # per_person=klasör, ortak=dosya
        self.commit = tk.BooleanVar(value=True)
        self.onayla = tk.BooleanVar(value=True)

        # --- DGS modülü (TAMAMEN AYRI state: Excel'i de park kodları da farklı) ---
        self.dgs_excel = tk.StringVar()
        self.dgs_person = tk.StringVar()
        self.dgs_limit = tk.StringVar()
        # commit'i CHECKBOX DEĞİL, butonlar belirler: "BAŞLAT"=gerçek, "GÜVENLİ DENEME"=yazmaz.
        # Operatör "commit" kavramını bilmek zorunda değil (tak-çalıştır kararı, 2026-07-14).
        self.dgs_commit = tk.BooleanVar(value=False)
        self.dgs_onayla = tk.BooleanVar(value=True)
        # Destek VARSAYILAN DAHİL — Mayıs'ta Destek'çiler de girildi (TPI 10, BV 3, Yıldız 1);
        # kapalı unutulursa sessizce eksik kalıyorlar (Mayıs'taki 10-kişi vakası).
        self.dgs_destek = tk.BooleanVar(value=True)
        # İstisna kişiler — otomasyonun DOKUNMAYACAĞI kişiler. [{"ad","tc","girildi"}]
        # Park başına diskte tutulur (dgs_istisna_<park>_<ay>.txt) → uygulama kapansa da kaybolmaz.
        self.istisna: "list[dict]" = []

        # Çalışan iş TEK: iki modül aynı Chrome/CDP oturumunu paylaşıyor → aynı anda ikisi koşamaz.
        self.proc = None
        self._run_btns: "list[ctk.CTkButton]" = []
        self._stop_btns: "list[ctk.CTkButton]" = []
        self.out_q: "queue.Queue[str]" = queue.Queue()

        # Başarısız kişi takibi: motor "!! BAŞARISIZ: <ad> — <sebep>" bastıkça toplanır; run bitince
        # (yalnız GERÇEK DGS-giriş koşularında) "tekrar dene" penceresi açılır. _retry_baglami None ise teklif yok.
        self._son_hatalar: "list[dict]" = []      # [{"ad","mesaj"}]
        self._retry_baglami = None                # {"modul","park",…} | None  → "Tekrar Dene" tuşunun bağlamı
        self._kapanis_asamasi = False             # True iken biten koşu KAPANIŞ'tır → tekrar kapanış tetikleme
        # İZİN tekrar-denemesi kişi kişi koşar: motorun --person'u TEK ad alıyor (izin_otomasyon.py:110),
        # liste dosyası yok. Bu yüzden seçilen adlar kuyruğa alınır, her koşu bitince sıradaki başlar.
        self._izin_retry_kuyruk: "list[str]" = []
        self._izin_retry_aktif = None             # şu an denenen ad (koşu bitince sonucu buna yazılır)
        self._izin_retry_sonuc: "list[dict]" = []  # [{"ad","ok","mesaj"}] → zincir sonunda özet

        # Kalıcı kullanıcı tercihleri (şimdilik yalnız log yazı boyu). Bozuk/eksik dosya sorun değil:
        # okunamazsa varsayılana düşeriz, uygulama yine açılır.
        self._ayarlar = self._ayar_oku()
        self._log_yazi_boy = self._ayarlar.get("log_yazi_boy", LOG_YAZI_VARSAYILAN)
        if not isinstance(self._log_yazi_boy, int) or not (LOG_YAZI_ALT <= self._log_yazi_boy <= LOG_YAZI_UST):
            self._log_yazi_boy = LOG_YAZI_VARSAYILAN

        # Dosya diyaloglarının en son açıldığı klasör (anahtar → yol). Operatör her seferinde
        # aynı yere gitmek zorunda kalmasın; ayrıca varsayılan klasör yoksa buradan kurtarırız.
        self._son_klasor: "dict[str, str]" = {}

        # Log GEÇMİŞİ bellekte tutulur: akış artık ana pencerede DEĞİL, ayrı pencerede gösteriliyor
        # (neden: _build_log). Pencere kapalıyken de birikir, açılınca oraya dökülür. Tavanlı liste.
        self._log_gecmis: "list[tuple[str, str | None]]" = []
        self._log_win = None            # açık log penceresi (CTkToplevel) | None
        self._log_win_txt = None        # o pencerenin metin kutusu | None
        self._log_retry_btn = None      # log penceresindeki "tekrar dene" tuşu | None
        self._log_ustte = None          # "en üstte tut" anahtarı (pencere açılınca kurulur) | None
        self._log_win_kapatildi = False  # operatör pencereyi bilerek kapattı mı? (otomatik açmayı susturur)
        self._retry_durum = None        # {"toplam","denenebilir"} | None → iki retry tuşunun ortak durumu

        # Canlı log'un DOSYA kopyası — gözetimli testte Claude `tail -f` ile birebir izler
        # (ekran görüntüsü izni gerekmez). Her açılışta ayraç atılır; dosya DATA_DIR'de kalır.
        self._log_dosyasi = None
        try:
            self._log_dosyasi = open(os.path.join(DATA_DIR, "gui_canli_log.txt"), "a", encoding="utf-8")
            self._log_dosyasi.write(f"\n===== GUI AÇILDI {datetime.datetime.now():%Y-%m-%d %H:%M:%S} "
                                    f"(frozen={izin_frozen.is_frozen()}) =====\n")
            self._log_dosyasi.flush()
        except OSError:
            pass                                   # log dosyası açılamazsa GUI yine de çalışır

        self._configure_styles()
        self._build()
        self._switch_module("İZİN")
        self.root.after(120, self._drain_log)
        # 🪟 WINDOWS: log penceresi AÇILIŞTA gelsin — operatör tuşu aramasın (kullanıcı isteği).
        # macOS'ta kapalı başlar (ilk koşuda kendiliğinden açılıyor); tek satır değiştirip
        # ikisinde de açtırabilirsin. Ana pencere çizilmeden açılamaz: konum hesabı
        # (_ust_pencere_plani) ana pencerenin köşesini okuyor → _acilista_log_penceresi bekler.
        if sys.platform.startswith("win"):
            self.root.after(400, self._acilista_log_penceresi)

    # ---------- arayüz kurulum ----------
    def _fit_window(self):
        """Ekrana taşmayan, ortalanmış başlangıç boyutu (13" Mac ve Windows laptop dahil).

        Hesabın tamamı ve Windows'taki DPI tuzağı için: `pencere_plani` docstring'i.
        """
        root = self.root
        try:
            # Ekranın DPI çarpanı: macOS → 1, Windows %125 → 1.25, %150 → 1.5.
            dpi = float(ctk.ScalingTracker.window_dpi_scaling_dict.get(root, 1.0)) or 1.0
        except Exception:
            dpi = 1.0
        p = pencere_plani(root.winfo_screenwidth(), root.winfo_screenheight(), dpi,
                          sys.platform.startswith("win"))
        self._olcek = p["olcek"]                      # kart/log ölçüleri de bunu kullanır
        # CTk'nin uyguladığı çarpan = DPI × buradaki global çarpan → hedef ölçeğe BÖLEREK ayarla.
        ctk.set_widget_scaling(p["olcek"] / dpi)
        ctk.set_window_scaling(p["olcek"] / dpi)
        root.geometry(f"{p['genislik']}x{p['yukseklik']}+{p['x']}+{p['y']}")
        root.minsize(p["min_w"], p["min_h"])
        # Log'a basılır: "log görünmüyor" şikâyetinde hangi ekranda ne hesaplandığı tek satırda belli olsun.
        self._ekran_bilgi = (f"{root.winfo_screenwidth()}x{root.winfo_screenheight()} · dpi={dpi:g} · "
                             f"ölçek={p['olcek']:.2f} · pencere={int(p['genislik'] * p['olcek'])}"
                             f"x{int(p['yukseklik'] * p['olcek'])} · {sys.platform}")
        if p["dar"]:
            try:
                root.state("zoomed")                 # yalnız Windows; başka yerde TclError → yut
            except Exception:
                pass

    def _configure_styles(self):
        """Ortak CustomTkinter görünümü.

        ⚠️ Ölçek (widget/window scaling) BURADA AYARLANMAZ — `_fit_window` ekrana göre hesaplar.
        Buraya sabit 1.0 yazmak Windows'ta o hesabı eziyordu (bkz. `pencere_plani`).
        """
        ctk.set_appearance_mode("dark")

    def _card(self, parent, number: str, title: str, subtitle: str):
        """Başlığı, numara rozeti ve ince çerçevesi olan yeniden kullanılabilir kart."""
        card = ctk.CTkFrame(parent, fg_color=UI["card"], corner_radius=18,
                            border_width=1, border_color=UI["border"])
        heading = ctk.CTkFrame(card, fg_color="transparent")
        heading.pack(fill="x", padx=18, pady=(16, 10))
        ctk.CTkLabel(heading, text=number, fg_color=UI["accent_soft"], text_color=UI["accent_text"],
                     width=38, height=28, corner_radius=10,
                     font=(FONT_UI, 9, "bold")).pack(side="left")
        title_box = ctk.CTkFrame(heading, fg_color="transparent")
        title_box.pack(side="left", fill="x", expand=True, padx=12)
        ctk.CTkLabel(title_box, text=title.upper(), text_color=UI["text"],
                     font=(FONT_UI, 11, "bold"), anchor="w").pack(fill="x")
        ctk.CTkLabel(title_box, text=subtitle, text_color=UI["muted"],
                     font=(FONT_UI, 9), anchor="w").pack(fill="x", pady=(3, 0))
        ctk.CTkFrame(card, fg_color=UI["border"], height=1, corner_radius=0).pack(fill="x", padx=18)
        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=18, pady=16)
        return card, body

    def _gelismis_alan(self, parent):
        """Katlanır 'Gelişmiş ayarlar' bölümü — operatörün (Begüm/Tolga) bilmesi GEREKMEYEN ayarlar
        (tek kişi, limit, commit ayrıntıları) burada gizli durur; varsayılanlar tak-çalıştır içindir."""
        holder = ctk.CTkFrame(parent, fg_color="transparent")
        holder.pack(fill="x", pady=(10, 0))
        icerik = ctk.CTkFrame(holder, fg_color=UI["input"], corner_radius=14,
                              border_width=1, border_color=UI["border"])

        def ac_kapa():
            if icerik.winfo_manager():
                icerik.pack_forget()
                btn.configure(text="Gelişmiş ayarlar  ▸")
            else:
                icerik.pack(fill="x", pady=(6, 0))
                btn.configure(text="Gelişmiş ayarlar  ▾")

        btn = ctk.CTkButton(holder, text="Gelişmiş ayarlar  ▸", command=ac_kapa, height=26,
                            corner_radius=12, fg_color="transparent", hover_color=UI["card_hover"],
                            text_color=UI["subtle"], anchor="w", font=("Helvetica", 9, "bold"))
        btn.pack(fill="x")
        return icerik

    def _make_step(self, parent, number: str, label: str, active=False):
        cell = ctk.CTkFrame(parent, fg_color="transparent")
        cell.pack(side="left", padx=(0, 8))
        color = UI["primary"] if active else UI["border"]
        fg = UI["text"] if active else UI["muted"]
        ctk.CTkLabel(cell, text=number, fg_color=color, text_color="white", width=24, height=24,
                     corner_radius=12, font=(FONT_UI, 9, "bold")).pack(side="left")
        ctk.CTkLabel(cell, text=label, text_color=fg, font=(FONT_UI, 9, "bold"),
                     padx=6).pack(side="left")
        return cell

    def _set_status(self, text: str, tone="info"):
        if not hasattr(self, "status_dot"):
            return
        color = {
            "info": UI["cyan"], "success": UI["success"], "warning": UI["warning"],
            "danger": UI["danger"], "running": UI["primary"],
        }.get(tone, UI["cyan"])
        self.status_dot.configure(text_color=color)
        self.status_text.configure(text=text)

    def _build(self):
        shell = ctk.CTkFrame(self.root, fg_color="transparent")
        shell.pack(fill="both", expand=True, padx=24, pady=22)

        # Başlık alanı
        header = ctk.CTkFrame(shell, fg_color=UI["surface"], corner_radius=20,
                              border_width=1, border_color=UI["border"])
        header.pack(fill="x")
        ctk.CTkFrame(header, height=3, fg_color=UI["primary"], corner_radius=3).pack(fill="x", padx=20)
        header_inner = ctk.CTkFrame(header, fg_color="transparent")
        header_inner.pack(fill="x")
        brand = ctk.CTkFrame(header_inner, fg_color="transparent")
        brand.pack(side="left", fill="x", expand=True, padx=20, pady=14)
        logo_tile = ctk.CTkFrame(brand, width=72, height=72, fg_color="#F6F0E9", corner_radius=18)
        logo_tile.pack(side="left", padx=(0, 16))
        logo_tile.pack_propagate(False)
        self.brand_logo = None
        try:
            logo_image = Image.open(BRAND_LOGO_PATH)
            self.brand_logo = ctk.CTkImage(light_image=logo_image, dark_image=logo_image, size=(64, 64))
            ctk.CTkLabel(logo_tile, text="", image=self.brand_logo).pack(expand=True)
        except Exception:
            ctk.CTkLabel(logo_tile, text="ETKN", text_color=UI["primary"],
                         font=(FONT_UI, 13, "bold")).pack(expand=True)
        brand_text = ctk.CTkFrame(brand, fg_color="transparent")
        brand_text.pack(side="left", fill="y")
        ctk.CTkLabel(brand_text, text="ETKİN PROJE  /  OPERASYON PLATFORMU", text_color=UI["primary"],
                     font=(FONT_UI, 9, "bold"), anchor="w").pack(fill="x", pady=(2, 2))
        self.head_title = ctk.CTkLabel(brand_text, text="İzin Operasyon Merkezi", text_color=UI["text"],
                                       font=(FONT_DISPLAY, 24, "bold"), anchor="w")
        self.head_title.pack(fill="x")
        self.head_sub = ctk.CTkLabel(brand_text, text="", text_color=UI["muted"],
                                     font=(FONT_UI, 10), anchor="w")
        self.head_sub.pack(fill="x", pady=(4, 0))
        status = ctk.CTkFrame(header_inner, fg_color=UI["input"], corner_radius=14,
                              border_width=1, border_color=UI["border"])
        status.pack(side="right", padx=20, pady=22)
        self.status_dot = ctk.CTkLabel(status, text="●", text_color=UI["cyan"], font=(FONT_UI, 13))
        self.status_dot.pack(side="left", padx=(12, 6), pady=8)
        self.status_text = ctk.CTkLabel(status, text="Kuruluma hazır", text_color=UI["text"],
                                         font=(FONT_UI, 10, "bold"))
        self.status_text.pack(side="left", padx=(0, 13), pady=8)

        # Modül seçici — İZİN | DGS. Aynı Chrome oturumu, farklı motor + farklı Excel.
        switcher = ctk.CTkFrame(shell, fg_color=UI["surface"], corner_radius=16,
                                border_width=1, border_color=UI["border"])
        switcher.pack(fill="x", pady=(12, 0), padx=1)
        ctk.CTkLabel(switcher, text="ÇALIŞMA ALANI", text_color=UI["subtle"],
                     font=(FONT_UI, 9, "bold")).pack(side="left", padx=(16, 13), pady=12)
        self.module = ctk.CTkSegmentedButton(
            switcher, values=["İZİN", "DGS"], command=self._switch_module, height=34, corner_radius=14,
            font=(FONT_UI, 11, "bold"), fg_color=UI["input"], selected_color=UI["primary"],
            selected_hover_color=UI["primary_hover"], unselected_color=UI["input"],
            unselected_hover_color=UI["card_hover"], text_color=UI["text"])
        self.module.pack(side="left", pady=9)
        self.module_hint = ctk.CTkLabel(switcher, text="", text_color=UI["muted"],
                                        font=(FONT_UI, 9), anchor="w")
        self.module_hint.pack(side="left", padx=14)
        # Resume klasörü — operatör "kayıtlarım nerede?" diye sormasın; done dosyaları BURAYA yazılır.
        ctk.CTkLabel(switcher, text="●  Kayıtlar güvenle saklanıyor", fg_color=UI["input"],
                     text_color=UI["success"], corner_radius=12, height=28,
                     font=(FONT_UI, 9, "bold")).pack(side="right", padx=14)
        # Sürüm damgası: "yeni exe mi çalışıyor?" sorusu buradan yanıtlanır (bkz. SURUM).
        ctk.CTkLabel(switcher, text=f"sürüm {SURUM}", text_color=UI["subtle"],
                     font=(FONT_UI, 9)).pack(side="right", padx=(0, 12))
        # Log penceresi tuşunun İKİNCİ kopyası — bu şerit pencerenin EN ÜSTÜNDE ve `fill="x"` ile
        # paketli: pencere ne kadar küçülürse küçülsün kırpılamaz. Alt taraftaki büyük tuş kırpılsa
        # bile operatör akışa buradan ulaşır (Windows "log yok" şikâyetine karşı çift emniyet).
        self._ust_log_btn = ctk.CTkButton(switcher, text="⧉  Canlı log", command=self._log_penceresi_ac,
                                          width=120, height=28, corner_radius=12, fg_color=UI["input"],
                                          hover_color=UI["card_hover"], text_color=UI["cyan"],
                                          font=(FONT_UI, 9, "bold"))
        self._ust_log_btn.pack(side="right")

        # İş akışı şeridi (modüle göre yeniden çizilir)
        flow = ctk.CTkFrame(shell, fg_color=UI["surface"], corner_radius=16,
                            border_width=1, border_color=UI["border"])
        flow.pack(fill="x", pady=(10, 14), padx=1)
        ctk.CTkLabel(flow, text="İŞ AKIŞI", text_color=UI["subtle"],
                     font=(FONT_UI, 9, "bold")).pack(side="left", padx=(16, 13), pady=12)
        self.flow_cells = ctk.CTkFrame(flow, fg_color="transparent")
        self.flow_cells.pack(side="left", pady=8)

        # ------------------------------------------------------------------
        # GÖVDE — SOL: kurulum kartları (kaydırılabilir) · SAĞ: çalıştır (sabit) + canlı log
        #
        # Neden kaydırma: modüllerin kart sayısı farklı (İzin 4, DGS 5) ve pencere küçültülebiliyor.
        # Sabit ızgarada kartlar KIRPILIYORDU. Kaydırma ile hiçbir koşulda içerik kesilmez;
        # "Başlat" ve log ise SAĞ sütunda sabit → hiç aranmaz, hep görünür.
        # ------------------------------------------------------------------
        body = ctk.CTkFrame(shell, fg_color="transparent")
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=11, minsize=560)     # kurulum
        body.columnconfigure(1, weight=9, minsize=440)      # çalıştır + log
        body.rowconfigure(0, weight=1)

        setup = ctk.CTkFrame(body, fg_color="transparent")
        setup.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        setup.rowconfigure(0, weight=1)
        setup.columnconfigure(0, weight=1)
        self.izin_view = ctk.CTkScrollableFrame(setup, fg_color="transparent",
                                                scrollbar_button_color=UI["card_hover"],
                                                scrollbar_button_hover_color=UI["primary"])
        self.dgs_view = ctk.CTkScrollableFrame(setup, fg_color="transparent",
                                               scrollbar_button_color=UI["card_hover"],
                                               scrollbar_button_hover_color=UI["primary"])
        for view in (self.izin_view, self.dgs_view):
            view.grid(row=0, column=0, sticky="nsew")

        self.right = right = ctk.CTkFrame(body, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        # 🔴 SATIR AĞIRLIKLARI TERSİNE ÇEVRİLDİ (Windows'ta log görünmüyordu — bkz. _build_log):
        # Tk grid, alan yetmediğinde açığın TAMAMINI ağırlıklı satırdan kısar. Eskiden ağırlık
        # LOG'daydı → dar ekranda log kırpılıp yok oluyordu. Artık ağırlık ÇALIŞTIR kartında:
        # yer daralınca onun ALTI kırpılır (BAŞLAT tuşu kartın en üstünde, hep görünür), log
        # kartı ise ağırlıksız olduğu için doğal boyunu KORUR ve asla kaybolmaz.
        # minsize TABANI: kart ezilip BAŞLAT tuşu kaybolmasın. DGS kartında tuşun ÜSTÜNDE bir de
        # uyarı satırı var (kart başlığı ~60 + uyarı ~2 satır + tuş 44 + pay) → 130 yetmiyordu.
        right.rowconfigure(0, weight=1, minsize=int(180 * getattr(self, "_olcek", 1.0)))
        right.rowconfigure(1, weight=0)                     # log kartı: sabit boy, dokunulmaz

        # Çalıştır sahnesi — her modülün kendi kartı (üst üste; modül seçicisiyle değişir)
        self.run_stage = ctk.CTkFrame(right, fg_color="transparent")
        self.run_stage.grid(row=0, column=0, sticky="nsew")
        self.run_stage.rowconfigure(0, weight=1)
        self.run_stage.columnconfigure(0, weight=1)

        # Log KARTI — PAYLAŞILAN (sahne dışında; modül değişince akış kesilmez).
        self._build_log(right)

        content = self.izin_view                            # İZİN kurulum kartları buraya dikey yığılır
        content.columnconfigure(0, weight=1)

        # 1) PARK
        f1, b1 = self._card(content, "01", "Teknopark seçimi", "İşlem yapılacak portalı belirleyin")
        f1.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        row = ctk.CTkFrame(b1, fg_color="transparent")
        row.pack(fill="x")
        self.park = ctk.CTkComboBox(row, values=AUTO_PARKS, state="readonly", width=210, height=40,
                                    corner_radius=14, fg_color=UI["input"], border_color=UI["border"],
                                    button_color=UI["primary"], button_hover_color=UI["primary_hover"],
                                    text_color=UI["text"], dropdown_fg_color=UI["surface"],
                                    dropdown_hover_color=UI["card_hover"], command=lambda _value: self._on_park_change())
        # --- YEDEK: self.park.set("YTP")  (combobox artık ortak adı gösteriyor)
        self.park.set(park_display(data.PARKS["YTP"]))
        self.park.pack(side="left")
        self.mode_lbl = ctk.CTkLabel(row, text="", text_color=UI["cyan"],
                                     font=("Helvetica", 10, "bold"), anchor="w")
        self.mode_lbl.pack(side="left", padx=12)
        self.park_chip = ctk.CTkLabel(row, text="", fg_color=UI["accent_soft"],
                                      text_color=UI["accent_text"], height=28,
                                      corner_radius=14, font=("Helvetica", 9, "bold"))
        self.park_chip.pack(side="right")
        self.mode_desc = ctk.CTkLabel(b1, text="", text_color=UI["muted"], wraplength=560,
                                      justify="left", anchor="w", font=("Helvetica", 9))
        self.mode_desc.pack(fill="x", pady=(11, 0))

        # 3) EXCEL
        f3, b3 = self._card(content, "03", "İzin Excel'i", "Personel ve izin günlerini içeren kaynak dosya")
        f3.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        row3 = ctk.CTkFrame(b3, fg_color="transparent")
        row3.pack(fill="x")
        ctk.CTkEntry(row3, textvariable=self.excel_path, height=40, corner_radius=14, fg_color=UI["input"],
                     border_color=UI["border"], text_color=UI["text"]).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(row3, text="DOSYA SEÇ", command=self._pick_excel, width=112, height=40, corner_radius=14,
                      fg_color=UI["card_hover"], hover_color=UI["border"], text_color=UI["text"],
                      font=("Helvetica", 10, "bold")).pack(side="left", padx=(8, 0))
        ctk.CTkButton(b3, text="Excel formatını görüntüle  →", command=self._info_excel, height=28,
                      corner_radius=14, fg_color="transparent", hover_color=UI["card_hover"], text_color=UI["cyan"],
                      anchor="w", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=(9, 0))

        # 4) BELGELER
        self.f4, b4 = self._card(content, "04", "İzin belgeleri", "PDF kuralları seçtiğiniz portala göre değişir")
        self.f4.grid(row=3, column=0, sticky="ew")
        row4 = ctk.CTkFrame(b4, fg_color="transparent")
        row4.pack(fill="x")
        self.belge_entry = ctk.CTkEntry(row4, textvariable=self.belge_path, height=40, corner_radius=14,
                                         fg_color=UI["input"], border_color=UI["border"], text_color=UI["text"])
        self.belge_entry.pack(side="left", fill="x", expand=True)
        self.belge_btn = ctk.CTkButton(row4, text="KLASÖR SEÇ", command=self._pick_belge, width=120, height=40,
                                        corner_radius=14, fg_color=UI["card_hover"], hover_color=UI["border"],
                                        text_color=UI["text"], font=("Helvetica", 10, "bold"))
        self.belge_btn.pack(side="left", padx=(8, 0))
        row4b = ctk.CTkFrame(b4, fg_color="transparent")
        row4b.pack(fill="x", pady=(6, 0))
        ctk.CTkButton(row4b, text="Belgeleri denetle", command=self._check_belge, height=28, corner_radius=14,
                      fg_color="transparent", hover_color=UI["card_hover"], text_color=UI["cyan"],
                      font=("Helvetica", 10, "bold")).pack(side="left")
        self.belge_info_btn = ctk.CTkButton(row4b, text="İsimlendirme kuralı", command=self._info_belge,
                                             height=28, corner_radius=14, fg_color="transparent",
                                             hover_color=UI["card_hover"], text_color=UI["cyan"],
                                             font=("Helvetica", 10, "bold"))
        self.belge_info_btn.pack(side="left", padx=(10, 0))
        self.belge_lbl = ctk.CTkLabel(b4, text="", text_color=UI["muted"], wraplength=560,
                                      justify="left", anchor="w", font=("Helvetica", 9))
        self.belge_lbl.pack(fill="x", pady=(4, 0))

        # 2) CHROME + LOGIN
        f2, b2 = self._card(content, "02", "Portal bağlantısı", "Chrome'u açın, giriş yapın ve bağlantıyı doğrulayın")
        f2.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        row2 = ctk.CTkFrame(b2, fg_color="transparent")
        row2.pack(fill="x")
        ctk.CTkButton(row2, text="Chrome'u aç", command=self._open_chrome, width=130, height=40, corner_radius=14,
                      fg_color=UI["card_hover"], hover_color=UI["border"], text_color=UI["text"],
                      font=("Helvetica", 10, "bold")).pack(side="left")
        ctk.CTkButton(row2, text="Login'i kontrol et", command=self._check_login, width=150, height=40,
                      corner_radius=14, fg_color="transparent", hover_color=UI["card_hover"], text_color=UI["cyan"],
                      font=("Helvetica", 10, "bold")).pack(side="left", padx=(8, 0))
        self.login_lbl = ctk.CTkLabel(b2, text="● Login henüz kontrol edilmedi", text_color=UI["warning"],
                                      font=("Helvetica", 9, "bold"), anchor="w")
        self.login_lbl.pack(fill="x", pady=(12, 3))
        ctk.CTkLabel(b2, text="Şifre uygulama tarafından görülmez. Açılan Chrome penceresinde giriş ve gerekiyorsa "
                              "Cloudflare doğrulamasını tamamlayın.", text_color=UI["muted"], wraplength=590,
                     justify="left", anchor="w", font=("Helvetica", 9)).pack(fill="x")

        # 5) ÇALIŞTIR — sağ sütunda, sahnenin dışında: her zaman görünür, kaydırma gerektirmez
        self.izin_run, b5 = self._card(self.run_stage, "05", "Çalıştır",
                                       "Emin değilsen önce 'Güvenli deneme' — hiçbir şey kaydetmez")
        self.izin_run.grid(row=0, column=0, sticky="nsew")
        self.start_btn = ctk.CTkButton(b5, text="OTOMASYONU BAŞLAT", command=lambda: self._start(plan=False),
                                       height=44, corner_radius=16, fg_color=UI["primary"],
                                       hover_color=UI["primary_hover"], text_color="white",
                                       font=("Helvetica", 11, "bold"))
        self.start_btn.pack(fill="x", pady=(0, 8))
        row5 = ctk.CTkFrame(b5, fg_color="transparent")
        row5.pack(fill="x")
        self.preflight_btn = ctk.CTkButton(row5, text="GÜVENLİ DENEME (yazmaz)",
                                           command=lambda: self._start(plan=True),
                                           height=38, corner_radius=14, fg_color=UI["card_hover"],
                                           hover_color=UI["border"], text_color=UI["text"],
                                           font=("Helvetica", 10, "bold"))
        self.preflight_btn.pack(side="left", fill="x", expand=True)
        self.stop_btn = ctk.CTkButton(row5, text="DURDUR", command=self._stop, width=104, height=38,
                                      corner_radius=14, fg_color=UI["danger_soft"],
                                      hover_color=UI["danger_hover"], text_color=UI["danger"],
                                      state="disabled", font=(FONT_UI, 10, "bold"))
        self.stop_btn.pack(side="left", padx=(8, 0))
        ctk.CTkLabel(b5, text="Başlat = giriş yapılır ve onaya gönderilir. Güvenli deneme = sadece ne "
                              "yapılacağını gösterir, portala dokunmaz.",
                     text_color=UI["subtle"], wraplength=380, justify="left", anchor="w",
                     font=("Helvetica", 9)).pack(fill="x", pady=(8, 0))
        adv5 = self._gelismis_alan(b5)
        ctk.CTkCheckBox(adv5, text="Gerçek kayıt (commit)", variable=self.commit, corner_radius=8,
                         border_color=UI["primary"], fg_color=UI["primary"], hover_color=UI["primary_hover"],
                         text_color=UI["text"], font=("Helvetica", 10)).pack(side="left", padx=12, pady=9)
        ctk.CTkCheckBox(adv5, text="Onaya gönder", variable=self.onayla, corner_radius=8,
                         border_color=UI["primary"], fg_color=UI["primary"], hover_color=UI["primary_hover"],
                         text_color=UI["text"], font=("Helvetica", 10)).pack(side="left", padx=8, pady=9)
        self._run_btns += [self.preflight_btn, self.start_btn]
        self._stop_btns.append(self.stop_btn)

        self._build_dgs()

    def _build_log(self, parent):
        """Log KARTI — akışın KENDİSİ ayrı pencerede gösterilir; ana pencerede yalnız bu kart durur.

        🔴 NEDEN AYRI PENCERE (Windows'ta "log görünmüyor" şikâyeti):
        Log paneli sağ sütunun DİBİNDEYDİ. Tk grid alan yetmediğinde açığın tamamını ağırlıklı
        satırdan kısar; dar/yüksek-DPI Windows ekranında "Çalıştır" kartı sütunun tamamını yiyor,
        log gövdesi pencerenin dışında kalıyordu. Önce DPI hesabı (`pencere_plani`), sonra çalıştır
        kartına yükseklik tavanı konarak dengelenmeye çalışıldı — Windows'ta yine görünmedi
        (o tavan kodu, `_sag_sutun_dengele`, artık gereksiz olduğu için SİLİNDİ).

        Kesin çözüm: log ana pencerede HİÇ yer kaplamasın. Bu kart sabit boyludur (ağırlıksız satır →
        asla kırpılmaz) ve tek işi log penceresini açmaktır. Pencerenin boyu ekrana göre hesaplanır
        (`_ust_pencere_plani`), yani kaç ölçekli ekran olursa olsun tamamı görünür.
        """
        card, body = self._card(parent, "LIVE", "Canlı işlem akışı",
                                "Akış ayrı pencerede görünür — koşu başlayınca kendiliğinden açılır")
        card.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        self.log_btn = ctk.CTkButton(body, text="⧉   CANLI LOG'U AÇ", command=self._log_penceresi_ac,
                                     height=44, corner_radius=16, fg_color=UI["primary"],
                                     hover_color=UI["primary_hover"], text_color="white",
                                     font=("Helvetica", 11, "bold"))
        self.log_btn.pack(fill="x")
        # Tek satırlık önizleme: log penceresi kapalıyken de "yaşıyor mu, nerede?" sorusunu yanıtlar.
        # ⚠️ SABİT BOYUTLU KUTU İÇİNDE: çıplak etiket, uzun satırda kendi istediği genişliği büyütüp
        # sağ sütunu (dolayısıyla tüm ızgarayı) her log satırında oynatıyordu. propagate kapalı kutu
        # etiketin ölçüsünü dışarı sızdırmaz — sütun genişliği log metnine göre zıplamaz.
        ozet_kutu = ctk.CTkFrame(body, fg_color="transparent",
                                 height=max(18, self._log_yazi_boy + 8))   # yazı büyüyünce kutu da
        ozet_kutu.pack(fill="x", pady=(9, 0))
        ozet_kutu.pack_propagate(False)
        self.log_ozet = ctk.CTkLabel(ozet_kutu, text="", text_color=UI["muted"],
                                     font=(FONT_MONO, max(9, self._log_yazi_boy - 2)),
                                     anchor="w", justify="left")
        self.log_ozet.pack(fill="both", expand=True)
        # Başarısızları-tekrar tuşu: normalde gizli, koşu başarısız kişiyle bitince belirir. AYNI tuşun
        # bir kopyası log penceresinde de var (bkz. _retry_btnleri_tazele) — operatör hangisine bakıyorsa.
        self.retry_btn = ctk.CTkButton(body, text="", command=self._retry_ac, height=38,
                                       corner_radius=14, fg_color=UI["warning"], hover_color=UI["primary_hover"],
                                       text_color="#1A1206", font=("Helvetica", 11, "bold"))
        self._log(f"Etkin Otomasyon · sürüm {SURUM}\n")
        self._log(f"Hazır. Resume/kayıt klasörü: {DATA_DIR}\n")
        # Ekran/ölçek teşhisi: log görünmeme şikâyetinde bu satır ne olduğunu tek bakışta söyler.
        self._log(f"[EKRAN] {getattr(self, '_ekran_bilgi', '?')}\n")

    def _ust_pencere_plani(self, ist_w: int, ist_h: int, taban_w: int = 420, taban_h: int = 300,
                           hiza: str = "orta"):
        """CTkToplevel için ekrana SIĞAN (geometry, min_w, min_h) — üçü de MANTIKSAL birimde.

        🔴 CTk `geometry()` ve `minsize()` verilen değeri pencere ölçeğiyle ÇARPAR (ayrıntı:
        `pencere_plani`). Sabit "780x700" bu yüzden %150 ölçekli Windows'ta 1170x1050 FİZİKSEL
        piksele çıkıyor; 1080p ekranda pencerenin altı (kapat/kaydet tuşları) görev çubuğunun
        altında kalıyordu — ana penceredeki log sorununun aynısı, alt pencerelerde.

        Burada istenen boy önce fiziksel çalışma alanına kırpılır, sonra ölçeğe BÖLÜNÜP mantıksal
        değere çevrilir. x/y ölçeklenmez (CTk onlara dokunmuyor) → ham piksel verilir.

        🔴 KONUM ANA PENCEREYE GÖRE: `winfo_screenwidth()` yalnız BİRİNCİL ekranı bilir. Konumu ona
        göre hesaplarsak, ana pencere ikinci monitördeyken alt pencere BAŞKA EKRANDA açılır ve
        operatör onu hiç görmez (log penceresinde bu tam olarak asıl şikâyetin tekrarı olurdu).
        Bu yüzden x/y ana pencerenin köşesinden türetilir → alt pencere hep aynı ekranda kalır.
        """
        olcek = getattr(self, "_olcek", 1.0) or 1.0
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        alan_w = max(360, sw - int(60 * olcek))                 # kenar payı
        alan_h = max(260, sh - int(120 * olcek))                # görev/menü çubuğu + başlık payı
        gen = max(1, int(min(ist_w, alan_w / olcek)))
        yuk = max(1, int(min(ist_h, alan_h / olcek)))
        fiz_w, fiz_h = int(gen * olcek), int(yuk * olcek)
        # Ana pencerenin köşesi ve ölçüsü — konumun dayanağı (yukarıdaki "çok monitör" notu).
        px, py = self.root.winfo_rootx(), self.root.winfo_rooty()
        pw, ph = self.root.winfo_width(), self.root.winfo_height()
        if pw <= 1 or ph <= 1:                                  # pencere henüz çizilmediyse ekrana göre
            px, py, pw, ph = 0, 0, sw, sh
        # hiza="sol": ana pencerenin SOL kenarına yaslanır → sağ sütun (BAŞLAT/DURDUR ve log kartı)
        # açıkta kalır; operatör koşuyu izlerken durdurmak için pencere sürüklemek zorunda kalmaz.
        x = px if hiza == "sol" else px + max(0, (pw - fiz_w) // 2)
        y = py + max(0, (ph - fiz_h) // 2)
        return f"{gen}x{yuk}+{x}+{y}", int(min(taban_w, gen)), int(min(taban_h, yuk))

    def _log_penceresi_ac(self, odak: bool = True, one_getir: bool = None):
        """Canlı akış penceresi — ana pencerede log paneli YOK, akış burada okunur (bkz. _build_log).

        odak=False      → klavye odağını ÇALMAZ (operatör o sırada Chrome'da login/Cloudflare ile
                          uğraşıyor olabilir). Kendiliğinden açılışlarda hep böyle çağrılır.
        one_getir=True  → zaten açık pencereyi öne alır (odak vermeden). Koşu BAŞLARKEN istenir.
        one_getir=False → açık pencereye hiç dokunma. İZİN tekrar-deneme ZİNCİRİNDE şart: her kişi
                          ayrı koşu, yoksa pencere her kişide simge durumundan fırlardı.
        Varsayılan: one_getir = odak.
        """
        if one_getir is None:
            one_getir = odak
        if odak:
            self._log_win_kapatildi = False                 # kullanıcı istedi → kilidi kaldır
        elif getattr(self, "_log_win_kapatildi", False):
            return                                          # kapatılmış pencereyi kendiliğinden açma
        win = getattr(self, "_log_win", None)
        if win is not None and win.winfo_exists():          # zaten açık → ikinci pencere AÇMA
            if one_getir:
                try:
                    win.deiconify()
                    win.lift()
                    if odak:
                        win.focus_force()
                except Exception:  # noqa
                    pass
            return

        win = ctk.CTkToplevel(self.root, fg_color=UI["bg"])
        win.title("Canlı işlem akışı — Etkin Otomasyon")
        geo, min_w, min_h = self._ust_pencere_plani(1000, 660, 460, 320, hiza="sol")
        win.geometry(geo)
        win.minsize(min_w, min_h)
        # transient KULLANILMIYOR: operatör bu pencereyi ikinci ekrana taşıyabilsin, ana pencereyle
        # birlikte küçülmesin. Chrome'un arkasında kaybolmaması için "En üstte tut" anahtarı var.

        # ⚠️ GRID (pack DEĞİL): pack, metin kutusuna sırayla yer verdiği için pencere asgari boya
        # inince DİPTEKİ tuş 42px yerine ~12px'e eziliyordu. Grid'de ağırlık metin kutusunda:
        # daralan tek şey o; başlık şeridi ve tuş doğal boylarını korur (ana penceredeki ders).
        win.columnconfigure(0, weight=1)
        win.rowconfigure(0, weight=0)                       # başlık şeridi
        win.rowconfigure(1, weight=1, minsize=80)           # akış (daralan taraf)
        win.rowconfigure(2, weight=0)                       # "kaydedilemedi" tuşu

        ust = ctk.CTkFrame(win, fg_color=UI["surface"], corner_radius=16)
        ust.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 0))
        ctk.CTkLabel(ust, text="CANLI İŞLEM AKIŞI", text_color=UI["text"],
                     font=(FONT_DISPLAY, 13, "bold")).pack(side="left", padx=16, pady=12)
        self._log_ustte = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(ust, text="En üstte tut", variable=self._log_ustte, corner_radius=6,
                        command=self._log_win_ustte_uygula, border_color=UI["primary"],
                        fg_color=UI["primary"], hover_color=UI["primary_hover"],
                        text_color=UI["muted"], font=("Helvetica", 10)).pack(side="right", padx=(0, 16))
        ctk.CTkButton(ust, text="Panoya kopyala", command=self._log_panoya, width=130, height=32,
                      corner_radius=12, fg_color=UI["card"], hover_color=UI["card_hover"],
                      text_color=UI["text"], font=("Helvetica", 10, "bold")).pack(side="right", padx=(0, 12))
        # Yazı boyu — Windows'ta 10 punto okunmuyordu (bkz. LOG_YAZI_VARSAYILAN). Seçim kalıcı.
        boy = ctk.CTkFrame(ust, fg_color="transparent")
        boy.pack(side="right", padx=(0, 12))
        ctk.CTkButton(boy, text="A−", width=38, height=32, corner_radius=12, fg_color=UI["card"],
                      hover_color=UI["card_hover"], text_color=UI["text"], font=("Helvetica", 12, "bold"),
                      command=lambda: self._log_yazi_degistir(-1)).pack(side="left")
        self._log_boy_lbl = ctk.CTkLabel(boy, text=str(self._log_yazi_boy), text_color=UI["subtle"],
                                         width=26, font=("Helvetica", 10))
        self._log_boy_lbl.pack(side="left")
        ctk.CTkButton(boy, text="A+", width=38, height=32, corner_radius=12, fg_color=UI["card"],
                      hover_color=UI["card_hover"], text_color=UI["text"], font=("Helvetica", 12, "bold"),
                      command=lambda: self._log_yazi_degistir(1)).pack(side="left")

        txt = ctk.CTkTextbox(win, corner_radius=14, border_width=1, border_color=UI["border"],
                             fg_color=UI["input"], text_color=UI["muted"],
                             font=(FONT_MONO, self._log_yazi_boy),      # kullanıcının seçtiği boy
                             scrollbar_button_color=UI["card_hover"],
                             scrollbar_button_hover_color=UI["primary"])
        txt.grid(row=1, column=0, sticky="nsew", padx=14, pady=14)
        for ad, renk in _LOG_RENK.items():
            txt._textbox.tag_configure(ad, foreground=renk)
        # "Kaydedilemedi" tuşu — log penceresinin DİBİNDE. Koşuyu buradan izleyen operatör tekrar
        # denemek için ana pencereye dönmek zorunda kalmasın. Normalde gizli (_retry_btnleri_tazele).
        self._log_retry_btn = ctk.CTkButton(win, text="", command=self._log_win_retry, height=42,
                                            corner_radius=14, fg_color=UI["warning"],
                                            hover_color=UI["primary_hover"], text_color="#1A1206",
                                            font=("Helvetica", 11, "bold"))
        try:
            for parca, tg in self._log_gecmis:              # o ana kadarki akışı (renkleriyle) dök
                txt.insert("end", parca, tg)
            txt.see("end")
        except Exception:  # noqa
            pass
        self._log_win, self._log_win_txt = win, txt
        win.protocol("WM_DELETE_WINDOW", self._log_penceresi_kapat)
        self._retry_btnleri_tazele()                        # açık bir başarısız-koşu varsa tuş hemen çıksın
        # CTkToplevel bazen ana pencerenin ALTINDA açılıyor → gecikmeli lift. Pencere bu arada
        # kapatılmış olabilir, ikisi de _sessizce ile sarılı.
        win.after(180, lambda: self._sessizce(win.lift))
        if odak:
            win.after(200, lambda: self._sessizce(win.focus_force))

    def _acilista_log_penceresi(self, deneme: int = 0):
        """Windows açılışı: log penceresini ana pencere ÇİZİLDİKTEN sonra aç, ana pencereyi önde bırak.

        Ana pencere daha çizilmemişken açarsak `_ust_pencere_plani` onun köşesini/ölçüsünü okuyamaz
        (winfo_width()<=1) ve pencere birincil ekrana göre konumlanır — çok monitörde yanlış ekran.
        Bu yüzden çizilene kadar 250ms'lik aralıklarla, en fazla ~3 sn beklenir.

        Sonunda ana pencere öne alınır: operatör kuruluma (park/Chrome/Excel) ondan başlıyor. Log
        penceresi açık ve görev çubuğunda durur; koşu başlar başlamaz `_spawn` onu zaten öne alır.
        """
        if getattr(self, "_log_win", None) is not None or getattr(self, "_log_win_kapatildi", False):
            return                                          # zaten açık ya da kullanıcı kapatmış
        if self.root.winfo_width() <= 1 and deneme < 12:
            self.root.after(250, lambda: self._acilista_log_penceresi(deneme + 1))
            return
        self._log_penceresi_ac(odak=False)
        # Ana pencere önde VE odakta kalsın: yeni Toplevel'e pencere yöneticisi kendiliğinden odak
        # verebiliyor (Windows'ta tipik). _log_penceresi_ac'ın 180ms'lik lift'inden SONRA çalışsın.
        self.root.after(260, lambda: (self._sessizce(self.root.lift),
                                      self._sessizce(self.root.focus_force)))

    @staticmethod
    def _sessizce(fn):
        """Pencere kullanıcı tarafından kapatılmış olabilir — after() geri çağrıları patlamasın."""
        try:
            fn()
        except Exception:  # noqa
            pass

    def _log_penceresi_kapat(self):
        """Log penceresi kapandı: canlı yazma hedeflerini temizle (yoksa yok edilmiş widget'a yazarız)."""
        win = getattr(self, "_log_win", None)
        self._log_win = self._log_win_txt = self._log_retry_btn = None
        # Operatör pencereyi BİLEREK kapattı → koşu başlangıçları onu zorla geri açmasın. İZİN
        # tekrar-deneme zinciri kişi başına bir koşu açıyor; bayrak olmasa pencere her kişide
        # yeniden ekrana fırlardı. Kilidi yalnız kullanıcı tuşa basınca (odak=True) kalkar.
        self._log_win_kapatildi = True
        if win is not None:
            self._sessizce(win.destroy)

    def _modal_oncesi(self):
        """Modal (grab_set / messagebox) açmadan ÖNCE log penceresinin 'en üstte' kilidini kaldır.

        🔴 NEDEN: '-topmost' pencere, grab_set'li modal'ın ÜSTÜNDE kalıyor. Modal tüm tıklamaları
        yuttuğu için operatör ne modal'ı görebiliyor ne de log penceresine tıklayabiliyordu →
        uygulama DONMUŞ gibi. Özellikle log penceresindeki 'TEKRAR DENE' tuşunda: onay kutusu
        pencerenin arkasında açılıyordu. Kutucuğu da temizliyoruz ki arayüz yalan söylemesin.
        """
        if getattr(self, "_log_ustte", None) is not None and self._log_ustte.get():
            self._log_ustte.set(False)
            self._log_win_ustte_uygula()

    def _log_win_ustte_uygula(self):
        """'En üstte tut' — Chrome otomasyonu öne geçtiğinde log penceresi arkada kaybolmasın."""
        win = getattr(self, "_log_win", None)
        if win is None or not win.winfo_exists():
            return
        self._sessizce(lambda: win.attributes("-topmost", bool(self._log_ustte.get())))

    def _log_panoya(self):
        """Bellekteki tüm akışı panoya kopyala (hata paylaşırken ekran görüntüsü uğraşı olmasın)."""
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append("".join(p for p, _ in self._log_gecmis))
            self._set_status("Log panoya kopyalandı", "info")
        except Exception:  # noqa
            pass

    def _log_win_retry(self):
        """Log penceresindeki tuş: bağlam varsa DOĞRUDAN tekrar-deneme (onay kutusu yine çıkar),
        yoksa 'kaydedilemeyen kişiler' penceresini açıp sebepleri gösterir."""
        b = self._retry_baglami
        denenebilir = [h["ad"] for h in self._son_hatalar if self._hata_yorumla(h["mesaj"])[1]]
        if b and denenebilir:
            (self._izin_retry if b.get("modul") == "izin" else self._dgs_retry)(denenebilir, None)
        else:
            self._retry_ac()

    # ---------- modül geçişi ----------
    _FLOW = {
        "İZİN": [("1", "Parkı seç"), ("2", "Portal bağlantısı"), ("3", "İzin Excel'i"),
                 ("4", "Belgeler"), ("5", "Çalıştır")],
        "DGS": [("1", "Parkı seç"), ("2", "Portal bağlantısı"), ("3", "DGS Excel'i"),
                ("4", "Mod ve kapsam"), ("5", "İstisna kişiler"), ("6", "DRY-RUN → gerçek kayıt")],
    }
    _BASLIK = {
        "İZİN": ("İzin Operasyon Merkezi",
                 "ETKİN PROJE  •  Yıllık izin girişi  •  Belge doğrulama  •  Tek akışta onay",
                 "Yıllık izin girişi + belge yükleme + onay."),
        "DGS": ("DGS Operasyon Merkezi",
                "ETKİN PROJE  •  Gelir vergisi istisnası  •  Puantaj girişi  •  SGK Gün = 30",
                "Aylık puantaj (DGS) girişi + onay. Excel'i ve park kodları İZİN'den FARKLI."),
    }

    def _switch_module(self, secim: str):
        dgs = secim == "DGS"
        self._retry_btn_gizle()   # tekrar-dene tuşu DGS'e özel; modül değişince kaldır (log paylaşımlı)
        # grid_remove/grid — kaydırılabilir çerçevede tkraise güvenilmez (iç canvas önde kalabiliyor)
        gizlenecek = (self.izin_view, self.izin_run) if dgs else (self.dgs_view, self.dgs_run)
        gosterilecek = (self.dgs_view, self.dgs_run) if dgs else (self.izin_view, self.izin_run)
        for w in gizlenecek:
            w.grid_remove()
        for w in gosterilecek:
            w.grid()
        baslik, altyazi, ipucu = self._BASLIK[secim]
        self.head_title.configure(text=baslik)
        self.head_sub.configure(text=altyazi)
        self.module_hint.configure(text=ipucu)
        self.root.title(f"Etkin Otomasyon — {secim}")
        if self.module.get() != secim:
            self.module.set(secim)
        for cell in self.flow_cells.winfo_children():
            cell.destroy()
        for i, (no, etiket) in enumerate(self._FLOW[secim]):
            self._make_step(self.flow_cells, no, etiket, active=(i == 0))
        (self._on_dgs_park_change if dgs else self._on_park_change)()

    # ---------- yardımcılar ----------
    def _cur_park(self):
        # combobox artık ORTAK ADI gösteriyor → içerideki İZİN koduna çevir (dönen Park nesnesi eskiyle BİREBİR aynı)
        # --- YEDEK: return data.PARKS[self.park.get()]
        return data.PARKS[IZIN_DISPLAY_TO_CODE[self.park.get()]]

    def _log(self, s: str, tag: str = None):
        """Akışa bir parça yaz: bellek geçmişi + açık log penceresi + önizleme + dosya.

        tag verilirse renk ZORLANIR (koşu sonu kırmızı özeti böyle basılır); verilmezse metinden
        sezilir. Ana pencerede artık log kutusu YOK — hedefler: `_log_gecmis`, log penceresi,
        `log_ozet` tek satırlık önizleme ve gui_canli_log.txt.
        """
        if tag is None:
            if s.lstrip().startswith("$"):
                tag = "command"
            elif "⚠" in s:
                tag = "warning"
            elif "!!" in s or "hata" in s.lower():
                tag = "danger"
            elif "✓" in s or "bitti" in s.lower():
                tag = "success"
        # Bellekteki geçmiş — log penceresi kapalıyken de birikir, açılınca oraya dökülür.
        # Tavan: 8 saatlik koşuda bellek şişmesin (tam kayıt zaten gui_canli_log.txt'de).
        self._log_gecmis.append((s, tag))
        if len(self._log_gecmis) > LOG_TAVAN:
            del self._log_gecmis[:len(self._log_gecmis) - LOG_TAVAN]
        win_txt = getattr(self, "_log_win_txt", None)      # log penceresi açıksa canlı yaz
        if win_txt is not None:
            try:
                if win_txt.winfo_exists():
                    win_txt.insert("end", s, tag)
                    win_txt.see("end")
                else:
                    self._log_win_txt = None
            except Exception:  # noqa
                self._log_win_txt = None
        ozet = getattr(self, "log_ozet", None)             # ana penceredeki tek satırlık önizleme
        if ozet is not None:
            # Son ANLAMLI satır: ayraç/boşluk satırları atlanır, yoksa önizlemede "─────" görünür.
            metin = next((x.strip() for x in reversed(s.splitlines())
                          if any(c.isalnum() for c in x)), None)
            if metin:
                kisa = metin if len(metin) <= 74 else metin[:73].rstrip() + "…"
                self._sessizce(lambda: ozet.configure(text=kisa,
                                                      text_color=_LOG_RENK.get(tag, UI["muted"])))
        if self._log_dosyasi:
            try:
                self._log_dosyasi.write(s)
                self._log_dosyasi.flush()
            except OSError:
                self._log_dosyasi = None

    def _drain_log(self):
        try:
            while True:
                line = self.out_q.get_nowait()
                self._log(line)
                self._ilerleme_guncelle(line)   # canlı sayaç: "=== [X/Y] AD ===" → durum çubuğu
        except queue.Empty:
            pass
        self.root.after(120, self._drain_log)

    def _ilerleme_guncelle(self, line: str):
        """Motorun '=== [X/Y] AD (proje: …) ===' satırından canlı ilerleme sayacı → durum çubuğu.
        Salt görüntü: 'İşleniyor: 13/38 · AD' — kullanıcı kaçıncı kişide olduğunu anında görür."""
        m = _ILERLEME_RE.search(line)
        if not m:
            return
        cur, tot = m.group(1), m.group(2)
        ad = m.group(3).split("(proje")[0].strip()
        self._set_status(f"İşleniyor: {cur}/{tot} · {ad}", "running")

    def _on_park_change(self):
        p = self._cur_park()
        short, desc = MODE_TR.get(p.onay_pdf, ("?", ""))
        self.mode_lbl.configure(text=f"MOD: {short.upper()}")
        # --- YEDEK: text=f"{p.code} PORTALI"  (chip'te de ortak ad; sekmeler arası kod farkı görünmesin)
        self.park_chip.configure(text=f"{park_display(p)} PORTALI")
        extra = f"  ·  Evrak tipi: “{p.evrak_tipi}”" if p.evrak_tipi else ""
        self.mode_desc.configure(text=f"{p.ad}  ·  {p.portal_url}\n{desc}{extra}")
        # belge bölümü modu
        if p.onay_pdf == "":
            self.belge_btn.configure(state="disabled")
            self.belge_info_btn.configure(state="disabled")
            self.belge_entry.configure(state="disabled")
            self.belge_lbl.configure(text="Bu park belge istemez — bu adımı güvenle atlayabilirsin.", text_color=UI["success"])
        elif p.onay_pdf == "ortak":
            self.belge_btn.configure(state="normal", text="Tek PDF Seç…", command=self._pick_belge)
            self.belge_info_btn.configure(state="normal")
            self.belge_entry.configure(state="normal")
            self.belge_lbl.configure(text="Firma tek toplu PDF yolladıysa yalnız o dosyayı seç.", text_color=UI["muted"])
        else:  # per_person
            self.belge_btn.configure(state="normal", text="Klasör Seç…", command=self._pick_belge)
            self.belge_info_btn.configure(state="normal")
            self.belge_entry.configure(state="normal")
            # ⚠ Uyarı BURADA da yazıyor: diyalog başlığını Windows kırpabiliyor, o zaman tek kalan bu.
            self.belge_lbl.configure(
                text="Her kişinin PDF'inin bulunduğu KLASÖRÜ seç. Klasör seçici dosyaları "
                     "listelemez — PDF görmemen normal, klasörün üstüne gelip seç yeter.",
                text_color=UI["muted"])
        # Belge seçimi YALNIZ park gerçekten değiştiğinde sıfırlanır. _switch_module de burayı
        # çağırıyor; koşulsuz temizlik, İZİN↔DGS sekmesine gidip gelen operatörün seçtiği klasörü
        # sessizce uçuruyordu.
        if getattr(self, "_son_park_kodu", None) != p.code:
            self._son_park_kodu = p.code
            self.belge_path.set("")
        self._set_status(f"{p.code} için kurulum hazır", "info")

    # ---------- Chrome / login ----------
    def _open_chrome(self):
        p = self._cur_park()
        ok, msg = launch_chrome(p.portal_url)
        self._log(("🌐 " if ok else "⚠️ ") + msg + "\n")
        self.login_lbl.configure(text="● Chrome açıldı — giriş yapıp bağlantıyı kontrol et", text_color=UI["warning"])
        self._set_status("Portal bağlantısı bekleniyor", "warning" if ok else "danger")

    def _probe_login(self) -> dict:
        """Login/park probe — İZİN ve DGS'in ORTAK kullandığı salt-okur kontrol. Portala bir şey yazmaz.
        ⚠ Döndürdüğü 'park' İZİN kodudur (YTP/İYTE/ULUTEK); DGS tarafı 'portal' alanından kendi koduna çevirir."""
        try:
            r = subprocess.run(izin_frozen.worker_cmd("logincheck") + [CDP_URL],
                               capture_output=True, text=True, timeout=40,
                               encoding="utf-8", errors="replace",
                               stdin=subprocess.DEVNULL, cwd=DATA_DIR)
            line = (r.stdout or "").strip().splitlines()[-1] if r.stdout.strip() else "{}"
            return json.loads(line)
        except Exception as e:  # noqa
            return {"ok": False, "reason": f"kontrol hatası: {str(e)[:100]}", "park": None}

    def _check_login(self):
        self.login_lbl.configure(text="● Bağlantı kontrol ediliyor…", text_color=UI["warning"])
        self._set_status("Portal doğrulanıyor", "running")
        self.root.update_idletasks()

        def worker():
            res = self._probe_login()
            self.root.after(0, lambda: self._login_result(res))

        threading.Thread(target=worker, daemon=True).start()

    def _login_result(self, res: dict):
        if res.get("ok"):
            self.login_lbl.configure(text=f"✓ {res.get('reason','Login OK')}", text_color=UI["success"])
            self._set_status("Portal bağlantısı hazır", "success")
            # açık portal seçili parkla uyuşuyor mu?
            # combobox artık ORTAK ADI gösteriyor → karşılaştırmayı İZİN KODU üzerinden yap (davranış eskiyle aynı)
            # --- YEDEK: if res.get("park") and res["park"] != self.park.get():
            if res.get("park") and res["park"] != self._cur_park().code:
                self._log(f"ℹ️ Açık portal {res['park']}, seçili park {self._cur_park().code}. "
                          f"Seçili parkı açık portala göre değiştir veya doğru portalı aç.\n")
        else:
            self.login_lbl.configure(text=f"● {res.get('reason','Login yok')}", text_color=UI["danger"])
            self._set_status("Portal bağlantısı kurulamadı", "danger")
        self._log(f"[LOGIN] {res.get('reason','')}\n")

    # ---------- kalıcı ayarlar ----------
    @staticmethod
    def _ayar_oku() -> dict:
        try:
            with open(AYAR_DOSYASI, encoding="utf-8") as f:
                d = json.load(f)
            return d if isinstance(d, dict) else {}
        except Exception:  # noqa — dosya yok / bozuk / okunamıyor: varsayılanlarla devam
            return {}

    def _ayar_yaz(self, **kv):
        """Tercihi diske yaz. Başarısız olursa SESSİZCE geç — ayar kaydı işi durdurmaz."""
        self._ayarlar.update(kv)
        try:
            with open(AYAR_DOSYASI, "w", encoding="utf-8") as f:
                json.dump(self._ayarlar, f, ensure_ascii=False, indent=1)
        except OSError:
            pass

    def _log_yazi_degistir(self, adim: int):
        """Log penceresindeki A− / A+ — yazı boyunu değiştirir ve kalıcı olarak saklar."""
        yeni = max(LOG_YAZI_ALT, min(LOG_YAZI_UST, self._log_yazi_boy + adim))
        if yeni == self._log_yazi_boy:
            return
        self._log_yazi_boy = yeni
        txt = getattr(self, "_log_win_txt", None)
        if txt is not None:
            self._sessizce(lambda: txt.configure(font=(FONT_MONO, yeni)))
        etiket = getattr(self, "_log_boy_lbl", None)
        if etiket is not None:
            self._sessizce(lambda: etiket.configure(text=str(yeni)))
        self._ayar_yaz(log_yazi_boy=yeni)

    # ---------- seçiciler ----------
    def _baslangic_klasoru(self, anahtar: str, *adaylar) -> str:
        """Dosya diyaloğu NEREDEN açılsın: en son kullanılan klasör → adaylar → ev klasörü.

        🔴 NEDEN VAR (Windows'ta "hiçbir dosya görünmüyor, klasörler boş" şikâyeti):
        `initialdir` diskte OLMAYAN bir klasörü gösterirse Windows'un yerel diyaloğu boş açılıyor.
        Eski kod yolları mac'e göre sabitlemişti (`~/Desktop`, `~/Desktop/İzin Belgeleri`); Windows'ta
        Masaüstü çoğu zaman OneDrive altına taşındığı için `~/Desktop` HİÇ YOKTU.
        Burada her aday `isdir` ile sınanır; hiçbiri yoksa ev klasörüne düşeriz — diyalog her koşulda
        gerçek bir yerde açılır.
        """
        son = self._son_klasor.get(anahtar)
        return izin_frozen._ilk_var_olan(son, *adaylar) or os.path.expanduser("~")

    def _klasoru_hatirla(self, anahtar: str, secim: str, dosya_mi: bool):
        """Seçim sonrası: bir dahaki sefere aynı yerden başla (dosya seçildiyse onun klasörü)."""
        yol = os.path.dirname(secim) if dosya_mi else secim
        if yol and os.path.isdir(yol):
            self._son_klasor[anahtar] = yol

    def _pick_excel(self):
        self._modal_oncesi()      # topmost log penceresi dosya diyaloğunun ÖNÜNE geçmesin
        f = filedialog.askopenfilename(
            parent=self.root,
            title="İzin Excel'i seç",
            filetypes=EXCEL_TIPLERI,
            initialdir=self._baslangic_klasoru("izin_excel", izin_frozen.indirilenler(),
                                               izin_frozen.masaustu()))
        if f:
            self.excel_path.set(f)
            self._klasoru_hatirla("izin_excel", f, dosya_mi=True)
            self._set_status("İzin Excel'i seçildi", "info")

    def _pick_belge(self):
        # 🔴 topmost log penceresi Windows'ta YEREL dosya diyaloğunu tamamen örtebiliyor; diyalog ana
        # pencereyi de kilitlediği için uygulama DONMUŞ görünüyor (bkz. _modal_oncesi).
        self._modal_oncesi()
        p = self._cur_park()
        masa = izin_frozen.masaustu()
        if p.onay_pdf == "ortak":
            # ORTAK mod: park için TEK toplu PDF → dosya seçici.
            f = filedialog.askopenfilename(
                parent=self.root,
                title=f"{p.code} — ortak izin belgesi (tek PDF dosyası seç)",
                filetypes=PDF_TIPLERI,
                initialdir=self._baslangic_klasoru("belge", os.path.join(masa, "İzin Belgeleri"), masa,
                                                   izin_frozen.indirilenler()))
        else:
            # PER_PERSON mod: kişi başına ayrı PDF → KLASÖR seçici.
            # ⚠ Klasör seçici tasarımı gereği DOSYA LİSTELEMEZ; kullanıcı bunu "klasörler boş"
            #   sanabiliyor. Başlıkta açıkça yazıyoruz ki PDF aramaya çalışmasın.
            f = filedialog.askdirectory(
                parent=self.root,
                title=f"{p.code} — PDF'lerin BULUNDUĞU KLASÖRÜ seç (dosyalar listelenmez)",
                initialdir=self._baslangic_klasoru("belge", os.path.join(masa, "İzin Belgeleri"), masa))
        if f:
            self.belge_path.set(f)
            self._klasoru_hatirla("belge", f, dosya_mi=(p.onay_pdf == "ortak"))
            self._belge_secim_ozeti(p, f)
            self._set_status("Belge kaynağı seçildi", "info")

    def _belge_secim_ozeti(self, park, secim: str):
        """Seçimin HEMEN ardından "ne seçtim, içinde ne var" de — Excel beklemeden.

        🔴 NEDEN: per_person modda klasör seçici kullanılıyor ve Windows'un klasör seçicisi tasarımı
        gereği HİÇBİR DOSYA GÖSTERMEZ. Operatör PDF'lerini göremeyince "klasörler boş, yanlış yer
        seçtim galiba" diye tereddüt ediyor. Ham PDF sayısı bu soruyu anında kapatır (Excel'e,
        eşleştirmeye gerek yok; ayrıntılı denetim yine "Belgeleri denetle" tuşunda).
        """
        try:
            if park.onay_pdf == "ortak":
                self.belge_lbl.configure(text=f"Seçilen belge: {os.path.basename(secim)}",
                                         text_color=UI["success"])
                return
            n = len(izin_belge._pdfs(secim))          # ham sayım — isim eşleştirmesi YAPMAZ
            ad = os.path.basename(secim.rstrip(os.sep)) or secim
            if n:
                self.belge_lbl.configure(
                    text=f"“{ad}” klasöründe {n} PDF bulundu. (Kimin hangi PDF'i aldığını görmek "
                         f"için “Belgeleri denetle”.)", text_color=UI["success"])
            else:
                self.belge_lbl.configure(
                    text=f"“{ad}” klasöründe hiç PDF yok — PDF'lerin DOĞRUDAN içinde olduğu "
                         f"klasörü seç (üst klasörü değil).", text_color=UI["warning"])
        except Exception:  # noqa — özet salt bilgi; okunamıyorsa akışı bozma
            pass

    # ---------- belge denetimi (portalsız) ----------
    def _check_belge(self):
        p = self._cur_park()
        xl = self.excel_path.get().strip()
        if not xl or not os.path.exists(os.path.expanduser(xl)):
            messagebox.showwarning("Excel yok", "Önce geçerli bir İzin Excel'i seç (3. adım).")
            return
        if p.onay_pdf == "":
            self.belge_lbl.configure(text="Bu park belge istemez.", text_color=UI["success"])
            self._set_status("Belge adımı gerekmiyor", "success")
            return
        try:
            by_park, _ = data.read_izin_v2(os.path.expanduser(xl), strict=True)
        except data.DataError as e:
            messagebox.showerror("Excel hatası", f"Excel doğrulanamadı:\n\n{str(e)[:1500]}")
            return
        people = by_park.get(p.code, [])
        if not people:
            self.belge_lbl.configure(text=f"{p.code}: Excel'de bu parkta kişi yok.", text_color=UI["warning"])
            self._set_status("Excel'de seçili park için kişi yok", "warning")
            return
        sel = self.belge_path.get().strip()
        folder = ortak = None
        if p.onay_pdf == "ortak":
            ortak = os.path.expanduser(sel) if sel else None
            r = izin_belge.readiness(people, None, mode="ortak", ortak_belge=ortak)
        else:
            folder = os.path.expanduser(sel) if sel else izin_belge.find_belge_klasor(p)
            r = izin_belge.readiness(people, folder, mode="per_person")
        es, top = r["eslesen"], r["toplam"]
        miss = [f"{m.ad}" for m in r["missing"]][:6]
        nonpdf = r["nonpdf"]
        if es == top and not nonpdf:
            self.belge_lbl.configure(text=f"✓ {es}/{top} belge hazır → onaya gidebilir.", text_color=UI["success"])
            self._set_status("Belgeler doğrulandı", "success")
        else:
            msg = f"⚠ {es}/{top} hazır."
            if miss:
                msg += "  Eksik: " + ", ".join(miss) + ("…" if len(r["missing"]) > 6 else "")
            if nonpdf:
                msg += f"  PDF-dışı: {len(nonpdf)} dosya (yalnız PDF kabul edilir)."
            self.belge_lbl.configure(text=msg, text_color=UI["danger"])
            self._set_status("Belge kontrolü gerekli", "danger")
        self._log(f"[BELGE] {p.code}: {es}/{top} hazır. klasör={folder or ortak}\n")

    # ---------- info popup'lar (ÖĞRETİCİ) ----------
    def _info_excel(self):
        txt = (
            "İZİN EXCEL'İ NASIL OLMALI?\n"
            "──────────────────────────\n"
            "Firma tek bir Excel yollar. Başlık satırı otomatik bulunur; şu KOLONLAR olmalı\n"
            "(başlık isimleri şöyle geçmeli):\n\n"
            "   • T.C.        → 11 haneli T.C. kimlik no (checksum doğrulanır)\n"
            "   • Ad-Soyad    → personelin adı soyadı\n"
            "   • Tarih       → izin günü, GG.AA.YYYY (her gün AYRI satır)\n"
            "   • Gün         → 1  (tam gün)  veya  0,5  (yarım gün)\n"
            "   • TGB         → park kodu:  TPI · BV · İYTE · YTP · ULUTEK · ARI · ODTÜ\n\n"
            "KURALLAR\n"
            "   • Bir kişinin BİRDEN ÇOK izin günü varsa her gün için 1 satır açılır.\n"
            "   • Aynı T.C. farklı isimle yazılırsa uyarı verir; T.C. esastır (kızlık/evlilik sorunu yaşanmaz).\n"
            "   • Tüm satırlar aynı aya ait olmalı (dönem otomatik bulunur).\n"
            "   • Bozuk satır varsa PORTALA HİÇ DOKUNULMADAN durur ve hataları listeler.\n\n"
            "ÖRNEK\n"
            # Örnek satırlar TEMSİLİ olmalı: gerçek T.C. yazma (yardım metni herkese görünür + repoda durur).
            "   T.C.          Ad-Soyad              Tarih        Gün   TGB\n"
            "   12345678901   Ahmet Yılmaz          04.06.2026   1     YTP\n"
            "   12345678901   Ahmet Yılmaz          05.06.2026   1     YTP\n"
            "   12345678902   Ayşe Demir            17.06.2026   0,5   YTP\n"
            "   12345678903   Zeynep Ustaoğlu Kaya  05.06.2026   1     YTP\n\n"
            "Not: ARI ve ODTÜ Teknoera sistemidir (Excel ile toplu yüklenir) → bu otomasyon kapsamı dışı."
        )
        self._popup("Excel formatı", txt)

    def _info_belge(self):
        p = self._cur_park()
        if p.onay_pdf == "ortak":
            txt = (
                f"{p.code} — ORTAK BELGE MODU\n"
                "──────────────────────────\n"
                "Firma, tüm çalışanları listeleyen TEK antetli/kaşeli yazı (tek PDF) yollar.\n\n"
                "   • Sadece 1 PDF seç — aynı belge o parktaki HERKESİN kaydına yüklenir.\n"
                "   • Yalnız PDF kabul edilir (Excel/Word reddedilir).\n"
                "   • Klasörde birden çok PDF varsa hangisi olduğu belirsizdir → tek dosyayı elle seç.\n"
            )
        elif p.onay_pdf == "per_person":
            txt = (
                f"{p.code} — KİŞİ-BAŞI BELGE MODU\n"
                "──────────────────────────\n"
                "Her personelin KENDİ izin formu (PDF) klasörde durur. İsimlendirme:\n\n"
                "   • Dosya adı = kişinin adı soyadı:   «Ad Soyad.pdf»\n"
                "        örn.  Abdullah Kızıl.pdf\n"
                "   • Bir kişide BİRDEN ÇOK belge varsa:  «Ad Soyad (2).pdf», «Ad Soyad (3).pdf»\n"
                "        örn.  Mehmet Delibaş.pdf · Mehmet Delibaş (2).pdf · Mehmet Delibaş (3).pdf\n"
                "   • Dosya adında T.C. de geçebilir — o da eşleşir.\n"
                "   • Portal isim yerine T.C. ile eşler → evlilik/kızlık soyadı farkı SORUN OLMAZ\n"
                "        (portal 'EZGİ ÇİFTÇİ', dosya 'Ezgi Ustaoğlu Çiftçi.pdf' → yine eşleşir).\n"
                "   • YALNIZ PDF. Excel/Word gelirse önce PDF'e çevir.\n\n"
                "Klasör kuralı (varsayılan):\n"
                f"   {os.path.join(izin_frozen.masaustu(), 'İzin Belgeleri', p.code + ' İzin Belgeleri')}\n"
                "   (farklı yerdeyse 'Klasör Seç…' ile göster.)"
            )
        else:
            txt = f"{p.code} — PDF'siz\nBu park onay için belge istemez; 4. adımı atla."
        self._popup(f"{p.code} belge isimlendirme", txt)

    def _popup(self, title, text):
        self._modal_oncesi()              # topmost log penceresi bu pencerenin önüne geçmesin
        win = ctk.CTkToplevel(self.root, fg_color=UI["bg"])
        win.title(title)
        geo, mw, mh = self._ust_pencere_plani(680, 500, 480, 340)   # bkz. _ust_pencere_plani
        win.geometry(geo)
        win.minsize(mw, mh)
        win.transient(self.root)
        head = ctk.CTkFrame(win, fg_color=UI["surface"], corner_radius=18)
        head.pack(fill="x", padx=16, pady=(16, 0))
        ctk.CTkLabel(head, text=title, text_color=UI["text"], font=(FONT_DISPLAY, 14, "bold"),
                     anchor="w").pack(fill="x", padx=18, pady=14)
        t = ctk.CTkTextbox(win, corner_radius=18, border_width=1, border_color=UI["border"],
                           fg_color=UI["input"], text_color=UI["muted"], font=(FONT_MONO, 11),
                           scrollbar_button_color=UI["card_hover"], scrollbar_button_hover_color=UI["primary"])
        t.pack(fill="both", expand=True, padx=16, pady=16)
        t.insert("1.0", text)
        t.configure(state="disabled")
        ctk.CTkButton(win, text="KAPAT", command=win.destroy, width=110, height=40, corner_radius=16,
                      fg_color=UI["primary"], hover_color=UI["primary_hover"], text_color="white",
                      font=("Helvetica", 10, "bold")).pack(pady=(0, 16))

    # ======================================================================
    # DGS MODÜLÜ
    # ======================================================================
    def _build_dgs(self):
        v = self.dgs_view
        v.columnconfigure(0, weight=1)

        # 01) PARK — combobox İzin sekmesiyle AYNI adı gösterir (içeride DGS kodu kullanılır; Excel/resume değişmez)
        # --- YEDEK (eski alt başlık): "DGS park kodları izin kodlarından farklıdır"
        d1, b1 = self._card(v, "01", "Teknopark seçimi", "İzin sekmesiyle aynı adlar — Excel'deki kod farkı için “Kod farkı?”")
        d1.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        r1 = ctk.CTkFrame(b1, fg_color="transparent")
        r1.pack(fill="x")
        self.dgs_park_cb = ctk.CTkComboBox(
            r1, values=DGS_PARKS, state="readonly", width=210, height=40, corner_radius=14,
            fg_color=UI["input"], border_color=UI["border"], button_color=UI["primary"],
            button_hover_color=UI["primary_hover"], text_color=UI["text"], dropdown_fg_color=UI["surface"],
            dropdown_hover_color=UI["card_hover"], command=lambda _v: self._on_dgs_park_change())
        # --- YEDEK: self.dgs_park_cb.set("TPI")  (combobox artık ortak adı gösteriyor)
        self.dgs_park_cb.set(park_display(dgs_park.PARKS["TPI"]))
        self.dgs_park_cb.pack(side="left")
        ctk.CTkButton(r1, text="Kod farkı?", command=self._info_dgs_kod, width=104, height=40,
                      corner_radius=14, fg_color="transparent", hover_color=UI["card_hover"],
                      text_color=UI["cyan"], font=("Helvetica", 10, "bold")).pack(side="left", padx=(8, 0))
        self.dgs_park_chip = ctk.CTkLabel(r1, text="", fg_color=UI["accent_soft"],
                                          text_color=UI["accent_text"], height=28,
                                          corner_radius=14, font=("Helvetica", 9, "bold"))
        self.dgs_park_chip.pack(side="right")
        self.dgs_park_desc = ctk.CTkLabel(b1, text="", text_color=UI["muted"], wraplength=560,
                                          justify="left", anchor="w", font=("Helvetica", 9))
        self.dgs_park_desc.pack(fill="x", pady=(11, 0))

        # 02) PORTAL — Chrome + login (izin tarafıyla AYNI probe; park eşleşmesi portal URL'sinden)
        d2, b2 = self._card(v, "02", "Portal bağlantısı", "Chrome'u açın, giriş yapın ve bağlantıyı doğrulayın")
        d2.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        r2 = ctk.CTkFrame(b2, fg_color="transparent")
        r2.pack(fill="x")
        ctk.CTkButton(r2, text="Chrome'u aç", command=self._dgs_open_chrome, width=130, height=40,
                      corner_radius=14, fg_color=UI["card_hover"], hover_color=UI["border"],
                      text_color=UI["text"], font=("Helvetica", 10, "bold")).pack(side="left")
        ctk.CTkButton(r2, text="Login'i kontrol et", command=self._dgs_check_login, width=150, height=40,
                      corner_radius=14, fg_color="transparent", hover_color=UI["card_hover"],
                      text_color=UI["cyan"], font=("Helvetica", 10, "bold")).pack(side="left", padx=(8, 0))
        self.dgs_login_lbl = ctk.CTkLabel(b2, text="● Login henüz kontrol edilmedi", text_color=UI["warning"],
                                          font=("Helvetica", 9, "bold"), anchor="w")
        self.dgs_login_lbl.pack(fill="x", pady=(12, 3))
        ctk.CTkLabel(b2, text="Cloudflare doğrulaması İNSAN adımıdır — uygulama geçmeyi denemez. Her park "
                              "değişiminde tekrar çıkabilir; Chrome'da geçip login'i yeniden kontrol edin.",
                     text_color=UI["muted"], wraplength=590, justify="left", anchor="w",
                     font=("Helvetica", 9)).pack(fill="x")

        # 03) DGS EXCEL — İZİN Excel'i DEĞİL
        d3, b3 = self._card(v, "03", "DGS Excel'i", "Puantaj kaynağı — izin Excel'inden AYRI dosya")
        d3.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        r3 = ctk.CTkFrame(b3, fg_color="transparent")
        r3.pack(fill="x")
        ctk.CTkEntry(r3, textvariable=self.dgs_excel, height=40, corner_radius=14, fg_color=UI["input"],
                     border_color=UI["border"], text_color=UI["text"]).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(r3, text="DOSYA SEÇ", command=self._pick_dgs_excel, width=112, height=40,
                      corner_radius=14, fg_color=UI["card_hover"], hover_color=UI["border"],
                      text_color=UI["text"], font=("Helvetica", 10, "bold")).pack(side="left", padx=(8, 0))
        ctk.CTkButton(b3, text="Excel formatını görüntüle  →", command=self._info_dgs_excel, height=28,
                      corner_radius=14, fg_color="transparent", hover_color=UI["card_hover"],
                      text_color=UI["cyan"], anchor="w",
                      font=("Helvetica", 10, "bold")).pack(anchor="w", pady=(9, 0))
        ctk.CTkLabel(b3, text="⚠ Bu, izin Excel'i DEĞİL: teknokent puantaj dosyası (ör. “06-TEKNOKENTLER — "
                              "HAZİRAN 2026 - FİNAL_9SAAT.xlsx”). Sayfa adı dönemin ay adıdır.",
                     text_color=UI["warning"], wraplength=560, justify="left", anchor="w",
                     font=("Helvetica", 9)).pack(fill="x", pady=(6, 0))

        # 04) MOD + KAPSAM
        d4, b4 = self._card(v, "04", "Mod ve kapsam", "Ne yapılacak ve kimlere uygulanacak")
        d4.grid(row=3, column=0, sticky="ew", pady=(0, 12))
        self.dgs_mod = ctk.CTkSegmentedButton(
            b4, values=list(DGS_MODLAR), command=lambda _v: self._on_dgs_mod_change(), height=36,
            corner_radius=14, font=("Helvetica", 10, "bold"), fg_color=UI["input"],
            selected_color=UI["primary"], selected_hover_color=UI["primary_hover"],
            unselected_color=UI["input"], unselected_hover_color=UI["card_hover"], text_color=UI["text"])
        self.dgs_mod.set("Giriş + Onay")
        self.dgs_mod.pack(fill="x")
        self.dgs_mod_desc = ctk.CTkLabel(b4, text="", text_color=UI["cyan"], wraplength=560,
                                         justify="left", anchor="w", font=("Helvetica", 9))
        self.dgs_mod_desc.pack(fill="x", pady=(8, 0))
        ctk.CTkLabel(b4, text="Kapsam: bu parktaki HERKES (Ar-Ge + Destek). Daha önce girilenler otomatik "
                              "atlanır, istisna listesindekilere dokunulmaz.",
                     text_color=UI["muted"], wraplength=560, justify="left", anchor="w",
                     font=("Helvetica", 9)).pack(fill="x", pady=(8, 0))
        # Tek kişi / limit / Destek anahtarı = test-ve-tamir araçları → operatörden GİZLİ (Gelişmiş)
        adv4 = self._gelismis_alan(b4)
        r4 = ctk.CTkFrame(adv4, fg_color="transparent")
        r4.pack(fill="x", padx=10, pady=(8, 0))
        ctk.CTkLabel(r4, text="Tek kişi:", text_color=UI["muted"],
                     font=("Helvetica", 9, "bold")).pack(side="left")
        ctk.CTkEntry(r4, textvariable=self.dgs_person, height=32, corner_radius=12, fg_color=UI["card"],
                     border_color=UI["border"], text_color=UI["text"],
                     placeholder_text="AD SOYAD (boş = herkes)").pack(side="left", fill="x", expand=True, padx=(6, 10))
        ctk.CTkLabel(r4, text="Limit:", text_color=UI["muted"],
                     font=("Helvetica", 9, "bold")).pack(side="left")
        ctk.CTkEntry(r4, textvariable=self.dgs_limit, height=32, width=64, corner_radius=12,
                     fg_color=UI["card"], border_color=UI["border"], text_color=UI["text"],
                     placeholder_text="—").pack(side="left", padx=(6, 0))
        self.dgs_destek_cb = ctk.CTkCheckBox(
            adv4, text="Destek personelini de işle (varsayılan AÇIK — Mayıs'ta Destek'çiler de girildi)",
            variable=self.dgs_destek, corner_radius=8, border_color=UI["primary"], fg_color=UI["primary"],
            hover_color=UI["primary_hover"], text_color=UI["text"], font=("Helvetica", 10))
        self.dgs_destek_cb.pack(anchor="w", padx=10, pady=8)

        # 05) İSTİSNA KİŞİLER — otomasyonun DOKUNMAYACAĞI kişiler
        d5, b5i = self._card(v, "05", "İstisna kişiler", "Otomasyonun DOKUNMAYACAĞI kişiler")
        d5.grid(row=4, column=0, sticky="ew")
        r5i = ctk.CTkFrame(b5i, fg_color="transparent")
        r5i.pack(fill="x")
        ctk.CTkButton(r5i, text="KİŞİ SEÇ / DÜZENLE", command=self._istisna_ac, width=180, height=40,
                      corner_radius=14, fg_color=UI["card_hover"], hover_color=UI["border"],
                      text_color=UI["text"], font=("Helvetica", 10, "bold")).pack(side="left")
        ctk.CTkButton(r5i, text="Temizle", command=self._istisna_temizle, width=90, height=40,
                      corner_radius=14, fg_color="transparent", hover_color=UI["card_hover"],
                      text_color=UI["cyan"], font=("Helvetica", 10, "bold")).pack(side="left", padx=(8, 0))
        self.istisna_chip = ctk.CTkLabel(r5i, text="", fg_color=UI["accent_soft"],
                                         text_color=UI["accent_text"], height=28,
                                         corner_radius=14, font=("Helvetica", 9, "bold"))
        self.istisna_chip.pack(side="right")
        self.istisna_lbl = ctk.CTkLabel(b5i, text="", text_color=UI["muted"], wraplength=560,
                                        justify="left", anchor="w", font=("Helvetica", 9))
        self.istisna_lbl.pack(fill="x", pady=(10, 0))

        # 06) ÇALIŞTIR — sağ sütunda, sahnenin dışında (hep görünür).
        # "commit" kavramı operatörden GİZLİ: BAŞLAT = gerçek kayıt · GÜVENLİ DENEME = hiçbir şey yazmaz.
        self.dgs_run, b5 = self._card(self.run_stage, "06", "Çalıştır",
                                      "Emin değilsen önce 'Güvenli deneme' — hiçbir şey kaydetmez")
        self.dgs_run.grid(row=0, column=0, sticky="nsew")
        self.dgs_uyari = ctk.CTkLabel(b5, text="", text_color=UI["warning"], wraplength=380,
                                      justify="left", anchor="w", font=("Helvetica", 9, "bold"))
        self.dgs_uyari.pack(fill="x", pady=(0, 8))
        self.dgs_start_btn = ctk.CTkButton(b5, text="DGS'İ BAŞLAT",
                                           command=lambda: self._dgs_start(gercek=True), height=44,
                                           corner_radius=16, fg_color=UI["primary"],
                                           hover_color=UI["primary_hover"], text_color="white",
                                           font=("Helvetica", 11, "bold"))
        self.dgs_start_btn.pack(fill="x", pady=(0, 8))
        rowd = ctk.CTkFrame(b5, fg_color="transparent")
        rowd.pack(fill="x")
        self.dgs_deneme_btn = ctk.CTkButton(rowd, text="GÜVENLİ DENEME (yazmaz)",
                                            command=lambda: self._dgs_start(gercek=False),
                                            height=38, corner_radius=14, fg_color=UI["card_hover"],
                                            hover_color=UI["border"], text_color=UI["text"],
                                            font=("Helvetica", 10, "bold"))
        self.dgs_deneme_btn.pack(side="left", fill="x", expand=True)
        self.dgs_stop_btn = ctk.CTkButton(rowd, text="DURDUR", command=self._stop, width=104, height=38,
                                          corner_radius=14, fg_color=UI["danger_soft"],
                                          hover_color=UI["danger_hover"], text_color=UI["danger"],
                                          state="disabled", font=(FONT_UI, 10, "bold"))
        self.dgs_stop_btn.pack(side="left", padx=(8, 0))
        ctk.CTkLabel(b5, text="Durdurursan kayıp olmaz — kaldığı yerden devam eder, kimse iki kez girilmez.",
                     text_color=UI["subtle"], wraplength=380, justify="left", anchor="w",
                     font=("Helvetica", 9)).pack(fill="x", pady=(8, 0))
        adv6 = self._gelismis_alan(b5)
        self.dgs_onayla_cb = ctk.CTkCheckBox(
            adv6, text="Onaya gönder (kapatırsan yalnız TASLAK açılır)", variable=self.dgs_onayla,
            corner_radius=8, border_color=UI["primary"], fg_color=UI["primary"],
            hover_color=UI["primary_hover"], text_color=UI["text"], font=("Helvetica", 10))
        self.dgs_onayla_cb.pack(anchor="w", padx=10, pady=8)
        self._run_btns += [self.dgs_start_btn, self.dgs_deneme_btn]
        self._stop_btns.append(self.dgs_stop_btn)
        self._on_dgs_mod_change()
        self._on_dgs_park_change()         # istisna/resume göstergelerini ilk değerlerine kur

    # ---------- DGS: park / mod ----------
    def _cur_dgs_park(self):
        # combobox artık ORTAK ADI gösteriyor → içerideki DGS koduna çevir (dönen Park nesnesi eskiyle BİREBİR aynı)
        # --- YEDEK: return dgs_park.PARKS[self.dgs_park_cb.get()]
        return dgs_park.PARKS[DGS_DISPLAY_TO_CODE[self.dgs_park_cb.get()]]

    def _on_dgs_park_change(self):
        p = self._cur_dgs_park()
        ay = dgs_donem_ay()
        n = dgs_done_sayisi(p.code, ay)
        # --- YEDEK: text=f"{p.code} PORTALI"  (chip'te de ortak ad; sekmeler arası kod farkı görünmesin)
        self.dgs_park_chip.configure(text=f"{park_display(p)} PORTALI")
        resume = (f"Resume: {n} kişi zaten girilmiş → TEKRAR GİRİLMEZ." if n
                  else "Resume dosyası yok — bu parkta sıfırdan başlar.")
        self.dgs_park_desc.configure(text=f"{p.ad}  ·  {p.portal_url}\nDönem: {ay} · {resume}")
        self._istisna_oku(p.code)          # istisnalar PARK BAŞINA tutulur → park değişince yeniden yükle
        self._istisna_yenile()
        self._set_status(f"DGS · {p.code} için kurulum hazır", "info")

    def _on_dgs_mod_change(self):
        mod, aciklama = DGS_MODLAR[self.dgs_mod.get()]
        self.dgs_mod_desc.configure(text=aciklama)
        salt_okur = mod == "kontrol"
        giris = mod == "giris"
        for w, aktif in ((self.dgs_onayla_cb, giris), (self.dgs_destek_cb, giris)):
            w.configure(state="normal" if aktif else "disabled")
        # kontrol zaten salt-okur → "deneme" diye ayrı bir hâli yok
        self.dgs_deneme_btn.configure(state="disabled" if salt_okur else "normal")
        if salt_okur:
            self.dgs_uyari.configure(text="SALT-OKUR — bu mod portala hiçbir şey yazmaz, sadece rapor okur.",
                                     text_color=UI["success"])
        else:
            self.dgs_uyari.configure(text="BAŞLAT = gerçek kayıt (onaya gönderme geri alınamaz). "
                                          "Güvenli deneme = portala dokunmaz.",
                                     text_color=UI["cyan"])

    # ---------- DGS: Chrome / login ----------
    def _dgs_open_chrome(self):
        p = self._cur_dgs_park()
        ok, msg = launch_chrome(p.portal_url)
        self._log(("🌐 " if ok else "⚠️ ") + f"{p.code}: " + msg + "\n")
        self.dgs_login_lbl.configure(text="● Chrome açıldı — giriş yapıp bağlantıyı kontrol et",
                                     text_color=UI["warning"])
        self._set_status("Portal bağlantısı bekleniyor", "warning" if ok else "danger")

    def _dgs_check_login(self):
        self.dgs_login_lbl.configure(text="● Bağlantı kontrol ediliyor…", text_color=UI["warning"])
        self._set_status("Portal doğrulanıyor", "running")
        self.root.update_idletasks()

        def worker():
            res = self._probe_login()
            self.root.after(0, lambda: self._dgs_login_result(res))

        threading.Thread(target=worker, daemon=True).start()

    def _dgs_login_result(self, res: dict):
        # ⚠ probe İZİN kodunu döndürür (YTP/İYTE/ULUTEK) → DGS koduna PORTAL ÜZERİNDEN çevir,
        #   yoksa her Yıldız/İzmir/Ulutek login'inde sahte "park uyuşmazlığı" uyarısı basardık.
        acik = DGS_BY_PORTAL.get(_domain(res.get("portal") or ""))
        secili = self._cur_dgs_park().code
        if res.get("ok"):
            self.dgs_login_lbl.configure(text=f"✓ Login OK — açık portal: {acik or res.get('park')}",
                                         text_color=UI["success"])
            self._set_status("Portal bağlantısı hazır", "success")
            if acik and acik != secili:
                self._log(f"ℹ️ Açık portal {acik}, seçili park {secili}. Doğru portalı aç ya da parkı değiştir.\n")
        else:
            self.dgs_login_lbl.configure(text=f"● {res.get('reason', 'Login yok')}", text_color=UI["danger"])
            self._set_status("Portal bağlantısı kurulamadı", "danger")
        self._log(f"[LOGIN] {res.get('reason', '')}\n")

    # ======================================================================
    # İSTİSNA KİŞİLER — otomasyonun DOKUNMAYACAĞI kişiler
    # ======================================================================
    def _istisna_dosyasi(self, park_code: str) -> str:
        return os.path.join(DATA_DIR, f"dgs_istisna_{park_code}_{dgs_donem_ay()}.txt")

    def _istisna_oku(self, park_code: str):
        """Diskten yükle — uygulama kapanıp açılsa da istisnalar kaybolmasın."""
        self.istisna = []
        try:
            with open(self._istisna_dosyasi(park_code), encoding="utf-8") as f:
                for satir in f:
                    satir = satir.split("#", 1)[0].strip()
                    if not satir:
                        continue
                    tc, ayrac, ad = satir.partition("|")
                    tc, ad = tc.strip(), ad.strip()
                    if not ayrac:                       # tek alan → T.C. mi ad mı?
                        tc, ad = (tc, "") if _tc_rakam(tc) else ("", tc)
                    self.istisna.append({"ad": ad or tc, "tc": tc if ad else "", "girildi": False})
        except OSError:
            pass

    def _istisna_kaydet(self, park_code: str):
        """Diske yaz — motor bu dosyayı `--exclude-file` ile okuyor (tek gerçek kaynak)."""
        yol = self._istisna_dosyasi(park_code)
        if not self.istisna:
            if os.path.exists(yol):
                os.remove(yol)
            self._log(f"[İSTİSNA] {park_code}: istisna listesi temizlendi.\n")
            return
        with open(yol, "w", encoding="utf-8") as f:
            f.write(f"# İSTİSNA KİŞİLER — {park_code} / {dgs_donem_ay()}\n"
                    "# Otomasyon bu kişilere DOKUNMAZ: kayıt açılmaz, onaya gönderilmez.\n"
                    "# Biçim: T.C.|AD SOYAD   (elle de düzenleyebilirsin)\n")
            for k in self.istisna:
                f.write(f"{k['tc']}|{k['ad']}\n")
        self._log(f"[İSTİSNA] {park_code}: {len(self.istisna)} kişi kaydedildi → {os.path.basename(yol)}\n")

    def _istisna_yenile(self):
        n = len(self.istisna)
        self.istisna_chip.configure(text=f"{n} KİŞİ" if n else "İSTİSNA YOK")
        if not n:
            self.istisna_lbl.configure(text="İstisna yok — Excel'deki tüm (Ar-Ge) kişiler işlenecek.",
                                       text_color=UI["muted"])
        else:
            adlar = ", ".join(k["ad"] for k in self.istisna[:4])
            if n > 4:
                adlar += f"  … (+{n - 4} kişi daha)"
            self.istisna_lbl.configure(text=f"⛔ Otomasyon DOKUNMAYACAK: {adlar}", text_color=UI["warning"])

    def _istisna_temizle(self):
        if not self.istisna:
            return
        p = self._cur_dgs_park()
        if not messagebox.askyesno(
                "İstisnaları temizle",
                f"{p.code} parkındaki {len(self.istisna)} istisna kişi silinecek.\n"
                "Bu kişiler bundan sonra otomasyona DAHİL olacak. Emin misin?"):
            return
        self.istisna = []
        self._istisna_kaydet(p.code)
        self._istisna_yenile()

    def _dgs_kisileri_yukle(self, park_code: str, xl: str):
        """Kişileri MOTORUN kendi okuyucusuyla al (`dgs ... liste`) → GUI ve motor BİREBİR aynı
        isimleri/T.C.'leri görür. Ayrı bir Excel parser yazsaydık isimler kayar, istisna tutmazdı."""
        cmd = izin_frozen.worker_cmd("dgs") + ["--park", park_code, "liste", "--excel", xl]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=120,
                               encoding="utf-8", errors="replace",
                               stdin=subprocess.DEVNULL, cwd=DATA_DIR)
        except Exception as e:  # noqa
            return [], f"Liste alınamadı: {e}"
        for satir in (r.stdout or "").splitlines():
            if satir.startswith("<<<LISTE>>>"):
                try:
                    return json.loads(satir[len("<<<LISTE>>>"):])["kisiler"], None
                except Exception as e:  # noqa
                    return [], f"Liste çözümlenemedi: {e}"
        detay = ((r.stderr or "").strip() or (r.stdout or "").strip())[-600:]
        return [], f"Motor kişi listesi döndürmedi.\n\n{detay}"

    def _istisna_ac(self):
        p = self._cur_dgs_park()
        xl = os.path.expanduser(self.dgs_excel.get().strip())
        if not xl or not os.path.exists(xl):
            messagebox.showwarning("DGS Excel'i yok",
                                   "İstisna listesi Excel'den okunuyor.\n\nÖnce 3. adımda DGS Excel'ini seç.")
            return
        self._set_status(f"{p.code}: Excel okunuyor…", "running")
        self.root.update_idletasks()
        kisiler, hata = self._dgs_kisileri_yukle(p.code, xl)
        if hata:
            self._set_status("Kişi listesi alınamadı", "danger")
            messagebox.showerror("Kişi listesi alınamadı", hata)
            return
        if not kisiler:
            self._set_status(f"{p.code}: Excel'de kişi yok", "warning")
            messagebox.showwarning("Kişi yok",
                                   f"Excel'de {p.code} parkında kişi bulunamadı.\n\n"
                                   "Doğru DGS Excel'i ve doğru park seçili mi? (DGS kodu = Excel F kolonu)")
            return
        self._set_status(f"{p.code}: {len(kisiler)} kişi listelendi", "info")
        self._istisna_popup(p, kisiler)

    def _istisna_popup(self, park, kisiler: "list[dict]"):
        secili = {_ist_key(k) for k in self.istisna}

        self._modal_oncesi()          # topmost log penceresi modal'ın önüne geçmesin
        win = ctk.CTkToplevel(self.root, fg_color=UI["bg"])
        win.title(f"İstisna Kişiler — {park.code}")
        geo, mw, mh = self._ust_pencere_plani(780, 700, 560, 440)   # bkz. _ust_pencere_plani
        win.geometry(geo)
        win.minsize(mw, mh)
        win.transient(self.root)

        head = ctk.CTkFrame(win, fg_color=UI["surface"], corner_radius=18)
        head.pack(fill="x", padx=16, pady=(16, 0))
        ctk.CTkLabel(head, text=f"İSTİSNA KİŞİLER — {park.code} · {dgs_donem_ay()}", text_color=UI["text"],
                     font=("Helvetica", 14, "bold"), anchor="w").pack(fill="x", padx=18, pady=(14, 2))
        ctk.CTkLabel(head, text="İşaretlediğin kişilere otomasyon DOKUNMAZ: kaydı açılmaz, onaya gönderilmez.",
                     text_color=UI["muted"], font=("Helvetica", 10), anchor="w").pack(fill="x", padx=18, pady=(0, 14))

        ara = tk.StringVar()
        srow = ctk.CTkFrame(win, fg_color="transparent")
        srow.pack(fill="x", padx=16, pady=(12, 8))
        ctk.CTkEntry(srow, textvariable=ara, height=40, corner_radius=14, fg_color=UI["input"],
                     border_color=UI["border"], text_color=UI["text"],
                     placeholder_text="Ara:  ad soyad  veya  T.C. …").pack(side="left", fill="x", expand=True)
        sayac = ctk.CTkLabel(srow, text="", text_color=UI["cyan"], font=("Helvetica", 10, "bold"))
        sayac.pack(side="left", padx=12)

        liste = ctk.CTkScrollableFrame(win, fg_color=UI["input"], corner_radius=16,
                                       scrollbar_button_color=UI["card_hover"],
                                       scrollbar_button_hover_color=UI["primary"])
        liste.pack(fill="both", expand=True, padx=16)

        # Satırlar BİR KEZ kurulur; arama yalnız pack/pack_forget yapar (her tuşta 100+ widget
        # yeniden yaratmak arayüzü kilitliyordu).
        durum, satirlar = {}, []
        for k in kisiler:
            anahtar = _ist_key(k)
            var = tk.BooleanVar(value=anahtar in secili)
            durum[anahtar] = var
            row = ctk.CTkFrame(liste, fg_color="transparent")
            ctk.CTkCheckBox(row, text=k["ad"], variable=var, corner_radius=6, checkbox_width=20,
                            checkbox_height=20, border_color=UI["primary"], fg_color=UI["primary"],
                            hover_color=UI["primary_hover"], text_color=UI["text"],
                            font=("Helvetica", 11)).pack(side="left", padx=(6, 0), pady=4)
            if k.get("girildi"):
                ctk.CTkLabel(row, text="✓ zaten girilmiş", text_color=UI["warning"],
                             font=("Helvetica", 9, "bold")).pack(side="right", padx=8)
            if not k.get("arge"):
                ctk.CTkLabel(row, text="DESTEK", text_color=UI["subtle"],
                             font=("Helvetica", 9, "bold")).pack(side="right", padx=8)
            tc = _tc_rakam(k.get("tc"))
            if len(tc) >= 4:                                  # T.C. MASKELİ göster (ekranda tam no durmasın)
                ctk.CTkLabel(row, text=f"{tc[:2]}•••••••{tc[-2:]}", text_color=UI["subtle"],
                             font=("Menlo", 10)).pack(side="right", padx=8)
            satirlar.append((row, _fold_tr(k["ad"]), tc))

        def filtrele(*_):
            q = _fold_tr(ara.get().strip())
            for row, _adf, _tc in satirlar:               # önce hepsini kaldır → sıra korunsun
                row.pack_forget()
            n = 0
            for row, adf, tc in satirlar:
                if not q or q in adf or q in tc:
                    row.pack(fill="x", pady=1)
                    n += 1
            sayac.configure(text=f"{n} / {len(satirlar)}")

        ara.trace_add("write", filtrele)
        filtrele()

        def kaydet():
            self.istisna = [{"ad": k["ad"], "tc": _tc_rakam(k.get("tc")), "girildi": bool(k.get("girildi"))}
                            for k in kisiler if durum[_ist_key(k)].get()]
            self._istisna_kaydet(park.code)
            self._istisna_yenile()
            win.destroy()
            # DÜRÜSTLÜK: istisna, PORTALDA ZATEN AÇILMIŞ kaydı geri almaz — sadece bundan sonrasını korur.
            girilmis = [k["ad"] for k in self.istisna if k["girildi"]]
            if girilmis:
                messagebox.showwarning(
                    "Bu kişiler zaten girilmiş",
                    "Aşağıdaki kişiler DAHA ÖNCE girilmiş (resume dosyasında kayıtlı):\n\n  • "
                    + "\n  • ".join(girilmis[:10]) + ("\n  …" if len(girilmis) > 10 else "")
                    + "\n\nİstisnaya almak portaldaki MEVCUT kaydı SİLMEZ — bundan sonra dokunulmayacak,\n"
                      "ama açılmış taslakları portaldan elle kaldırman gerekir.")

        foot = ctk.CTkFrame(win, fg_color="transparent")
        foot.pack(fill="x", padx=16, pady=16)
        ctk.CTkButton(foot, text="KAYDET", command=kaydet, width=130, height=42, corner_radius=16,
                      fg_color=UI["primary"], hover_color=UI["primary_hover"], text_color="white",
                      font=("Helvetica", 10, "bold")).pack(side="right")
        ctk.CTkButton(foot, text="İPTAL", command=win.destroy, width=110, height=42, corner_radius=16,
                      fg_color=UI["card_hover"], hover_color=UI["border"], text_color=UI["text"],
                      font=("Helvetica", 10, "bold")).pack(side="right", padx=(0, 8))
        ctk.CTkLabel(foot, text="T.C. maskeli gösterilir. Eşleştirme T.C. ile yapılır → soyadı değişse de tutar.",
                     text_color=UI["subtle"], font=("Helvetica", 9)).pack(side="left")

        win.after(180, win.lift)          # CTkToplevel bazen ana pencerenin ALTINDA açılıyor

    # ---------- DGS: seçici + öğretici popup'lar ----------
    def _pick_dgs_excel(self):
        self._modal_oncesi()      # topmost log penceresi dosya diyaloğunun ÖNÜNE geçmesin
        f = filedialog.askopenfilename(
            parent=self.root,
            title="DGS Excel'i seç (teknokent puantaj dosyası)",
            filetypes=EXCEL_TIPLERI,
            initialdir=self._baslangic_klasoru("dgs_excel", izin_frozen.indirilenler(),
                                               izin_frozen.masaustu()))
        if f:
            self.dgs_excel.set(f)
            self._klasoru_hatirla("dgs_excel", f, dosya_mi=True)
            self._set_status("DGS Excel'i seçildi", "info")

    def _info_dgs_kod(self):
        self._popup("DGS park kodları", (
            "DİKKAT: AYNI PORTAL, FARKLI KOD\n"
            "──────────────────────────────\n"
            "İzin modülü ile DGS modülü aynı portalları kullanır ama park KODLARI farklıdır.\n"
            "DGS kodu = Excel'in “Lokasyonu SGK” (F) kolonundaki değerdir; resume dosya adlarında\n"
            "da o geçer (ör. dgs_done_Yıldız_Haziran.txt).\n\n"
            "   PORTAL                 İZİN kodu      DGS kodu\n"
            "   ─────────────────────  ───────────    ────────\n"
            "   Teknopark İstanbul     TPI            TPI\n"
            "   Bilişim Vadisi         BV             BV\n"
            "   Teknopark İzmir        İYTE           TPIz     ←\n"
            "   Yıldız Teknopark       YTP            Yıldız   ←\n"
            "   Bursa Ulutek           ULUTEK         Ulutek   ←\n\n"
            "Artık her iki sekmede de aynı TEKNOKENT ADINI görürsün (ör. “Teknopark İzmir”); yukarıdaki\n"
            "kodlar yalnız DOSYALARDA/Excel'de geçer. DGS Excel'inin “Lokasyonu SGK” (F) kolonu ve resume\n"
            "dosya adları hâlâ DGS kodunu kullanır (ör. dgs_done_Yıldız_Haziran.txt). Yanlış portal açıksa\n"
            "motor oraya YAZMAZ — lokasyon filtresi tutmaz, hiçbir kişi işlenmez."
        ))

    def _info_dgs_excel(self):
        self._popup("DGS Excel formatı", (
            "DGS EXCEL'İ NASIL OLMALI?\n"
            "─────────────────────────\n"
            "Teknokentlerin ortak puantaj dosyası (ör. “06-TEKNOKENTLER - HAZİRAN 2026 - FİNAL_9SAAT.xlsx”).\n"
            "Bu format HEP AYNI gelir — kolonlar YERİNE GÖRE okunur:\n\n"
            "   A → Ad Soyad\n"
            "   B → T.C. kimlik no        (kimlik anahtarı; evlilik/kızlık soyadı sorunu yaşanmaz)\n"
            "   E → Bölüm\n"
            "   F → Lokasyonu SGK         ← PARK KODU buradan gelir (TPI · BV · TPIz · Yıldız · Ulutek)\n"
            "   H → Ar-Ge / Destek        ← Destek'çiler VARSAYILAN OLARAK İŞLENMEZ (4. adımdaki kutu)\n"
            "   N → Eksik puantaj (saat)\n"
            "   W → Proje adı\n\n"
            "SAYFA (sheet): dönemin ay adı — ör. “Haziran”. Otomatik seçilir.\n"
            "DÖNEM: bugünden bir ÖNCEKİ ay. Otomatik türetilir.\n\n"
            "⚠ Bu dosya İZİN Excel'i DEĞİLDİR. İzin Excel'i ayrı bir dosyadır (T.C./Ad-Soyad/Tarih/Gün/TGB).\n"
            "   Karıştırırsan motor kişileri bulamaz ve hiçbir şey yazmadan durur."
        ))

    # ---------- DGS: çalıştır ----------
    def _dgs_build_cmd(self) -> list:
        p = self._cur_dgs_park()
        mod, _ = DGS_MODLAR[self.dgs_mod.get()]
        cmd = izin_frozen.worker_cmd("dgs") + ["--park", p.code, mod]
        if mod != "kontrol":                       # kontrol motoru --excel KABUL ETMEZ
            cmd += ["--excel", os.path.expanduser(self.dgs_excel.get().strip())]
        if mod == "giris":
            if self.dgs_onayla.get():
                cmd.append("--onayla")
            if self.dgs_destek.get():
                cmd.append("--include-destek")
            kisi = self.dgs_person.get().strip()
            if kisi:
                cmd += ["--person", kisi]
        if mod in ("giris", "onay"):
            limit = self.dgs_limit.get().strip()
            if limit:
                cmd += ["--limit", limit]
            if self.dgs_commit.get():
                cmd.append("--commit")
            # İSTİSNA: motor bu dosyadaki kişileri hiç görmez (giris→Excel'den, onay→portal listesinden düşer)
            yol = self._istisna_dosyasi(p.code)
            if self.istisna and os.path.exists(yol):
                cmd += ["--exclude-file", yol]
        return cmd

    def _dgs_start(self, gercek: bool):
        if self.proc is not None:
            messagebox.showinfo("Çalışıyor", "Bir işlem zaten sürüyor. Önce bitmesini bekle veya Durdur.")
            return
        p = self._cur_dgs_park()
        mod, _ = DGS_MODLAR[self.dgs_mod.get()]
        # commit'i buton belirler: BAŞLAT=gerçek, GÜVENLİ DENEME=yazmaz (operatör "commit" bilmek zorunda değil)
        self.dgs_commit.set(gercek)

        limit = self.dgs_limit.get().strip()
        if limit and not limit.isdigit():
            messagebox.showwarning("Geçersiz limit", "Limit bir SAYI olmalı (ör. 1). Boş bırakırsan herkes işlenir.")
            return
        if mod != "kontrol":
            xl = os.path.expanduser(self.dgs_excel.get().strip())
            if not xl or not os.path.exists(xl):
                messagebox.showwarning("DGS Excel'i yok",
                                       "Önce geçerli bir DGS Excel'i seç (3. adım).\n\n"
                                       "DİKKAT: bu İZİN Excel'i DEĞİL — teknokent puantaj dosyası.")
                return
        if mod != "kontrol" and self.dgs_commit.get():
            ne = ("GİRİŞ + ONAYA GÖNDERME" if (mod == "giris" and self.dgs_onayla.get())
                  else "ONAYA GÖNDERME" if mod == "onay" else "TASLAK KAYIT")
            grup = "Ar-Ge + Destek" if self.dgs_destek.get() else "yalnız Ar-Ge"
            kapsam = (f"tek kişi: {self.dgs_person.get().strip()}" if self.dgs_person.get().strip()
                      else f"{limit} kişi" if limit else f"TÜM kalan kişiler ({grup})")
            haric = (f"\nİstisna: {len(self.istisna)} kişiye DOKUNULMAYACAK." if self.istisna
                     else "\nİstisna: yok.")
            if not messagebox.askyesno(
                    "Gerçek kayıt onayı",
                    f"{p.code} ({p.ad}) portalında GERÇEK {ne} yapılacak.\n"
                    f"Kapsam: {kapsam}.{haric}\n\n"
                    "Onaya gönderme GERİ ALINAMAZ. Devam edilsin mi?"):
                return

        self._log("[LOGIN] kontrol ediliyor…\n")
        self.root.update_idletasks()
        res = self._probe_login()
        self._dgs_login_result(res)
        if not res.get("ok"):
            messagebox.showwarning("Login algılanmadı",
                                   f"{res.get('reason', '')}\n\nÖnce Chrome'da login ol (Cloudflare'i geç), "
                                   "sonra tekrar Başlat.")
            return
        acik = DGS_BY_PORTAL.get(_domain(res.get("portal") or ""))
        if acik and acik != p.code:
            if not messagebox.askyesno(
                    "Park uyuşmazlığı",
                    f"Açık portal: {acik}.  Seçili park: {p.code}.\n\n"
                    "Motor yanlış portala YAZMAZ (lokasyon filtresi tutmaz, kimse işlenmez).\n"
                    "Yine de devam edilsin mi?"):
                return

        # Tekrar-dene teklifi YALNIZ gerçek DGS-giriş koşusundan sonra anlamlı (dry-run/onay/kontrol'de değil).
        if mod == "giris" and self.dgs_commit.get():
            self._retry_baglami = {
                "modul": "dgs",
                "park": p.code,
                "excel": os.path.expanduser(self.dgs_excel.get().strip()),
                "onayla": bool(self.dgs_onayla.get()),
                "destek": bool(self.dgs_destek.get()),
                # TAM giriş mi? (tek kişi/limit YOKSA) → giriş sonrası otomatik KAPANIŞ yalnız tam koşuda
                "tam": not (self.dgs_person.get().strip() or self.dgs_limit.get().strip()),
            }
        else:
            self._retry_baglami = None

        self._spawn(self._dgs_build_cmd(),
                    "SGK kontrolü çalışıyor" if mod == "kontrol" else "DGS otomasyonu çalışıyor")

    # ---------- çalıştır ----------
    def _build_cmd(self, plan: bool):
        p = self._cur_park()
        xl = os.path.expanduser(self.excel_path.get().strip())
        cmd = izin_frozen.worker_cmd("orchestrator") + [
               "--excel", xl, "--parklar", p.code, "--cdp", CDP_URL]
        if plan:
            cmd.append("--plan")
            return cmd
        if p.onay_pdf == "ortak":
            sel = self.belge_path.get().strip()
            if sel:
                cmd += ["--ortak-belge", os.path.expanduser(sel)]
        elif p.onay_pdf == "per_person":
            sel = self.belge_path.get().strip()
            if sel:
                cmd += ["--belge-klasor", os.path.expanduser(sel)]
        if self.commit.get():
            cmd.append("--commit")
        if self.onayla.get():
            cmd.append("--onayla")
        return cmd

    def _start(self, plan: bool):
        if self.proc is not None:
            messagebox.showinfo("Çalışıyor", "Bir işlem zaten sürüyor. Önce bitmesini bekle veya Durdur.")
            return
        xl = self.excel_path.get().strip()
        if not xl or not os.path.exists(os.path.expanduser(xl)):
            messagebox.showwarning("Excel yok", "Önce geçerli bir İzin Excel'i seç (3. adım).")
            return
        if not plan:
            if self.commit.get() and self.onayla.get():
                if not messagebox.askyesno(
                        "Onay",
                        f"{self.park.get()} parkında GERÇEK giriş + ONAYA GÖNDERME yapılacak.\n"
                        "Onaya gönderme GERİ ALINAMAZ. Devam edilsin mi?"):
                    return
            # login algıla (probe) — Başlat'ta kullanıcı login'i anlaşılır
            self._log("[LOGIN] kontrol ediliyor…\n")
            self.root.update_idletasks()
            res = self._probe_login()
            self._login_result(res)
            if not res.get("ok"):
                messagebox.showwarning("Login algılanmadı",
                                       f"{res.get('reason','')}\n\nÖnce Chrome'da login ol, sonra tekrar Başlat.")
                return
            # combobox artık ORTAK ADI gösteriyor → karşılaştırmayı İZİN KODU üzerinden yap (davranış eskiyle aynı)
            # --- YEDEK: if res.get("park") and res["park"] != self.park.get():
            if res.get("park") and res["park"] != self._cur_park().code:
                if not messagebox.askyesno("Park uyuşmazlığı",
                                           f"Açık portal {res['park']}, seçili park {self._cur_park().code}.\n"
                                           "Yine de seçili parkla devam? (Motor yanlış portala işlem yapmaz, durur.)"):
                    return

        # Tekrar-dene bağlamı: yalnız GERÇEK koşuda anlamlı (güvenli denemede yazan bir şey yok).
        # Komut ŞİMDİ dondurulur — kullanıcı popup'a basana kadar Excel/belge alanını değiştirmiş
        # olabilir; retry, biten koşunun ayarlarıyla koşmalı.
        self._retry_baglami = None if plan else {"modul": "izin", "park": self._cur_park().code,
                                                 "cmd": self._build_cmd(False)}
        self._spawn(self._build_cmd(plan), "Ön uçuş çalışıyor" if plan else "Otomasyon çalışıyor")

    def _spawn(self, cmd: list, calisiyor_mesaji: str):
        """Motoru ALT-SÜREÇ olarak koş — İZİN ve DGS ortak. Çıktısı canlı log'a akar."""
        # Akış ana pencerede DEĞİL (bkz. _build_log) → koşu başlarken log penceresini aç ve ÖNE al.
        # odak=False: operatör o an Chrome'da login/Cloudflare ile uğraşıyor olabilir, klavyeyi çalma.
        # one_getir: yalnız YENİ bir koşuda. İZİN tekrar-deneme zinciri kişi başına bir _spawn açıyor
        # (_izin_retry_aktif dolu) — orada öne almazsak pencere her kişide ekrana fırlamamış olur.
        self._sessizce(lambda: self._log_penceresi_ac(odak=False,
                                                      one_getir=self._izin_retry_aktif is None))
        self._log("\n$ " + " ".join(f'"{c}"' if " " in c else c for c in cmd) + "\n")
        self._son_hatalar = []               # bu koşunun başarısızları taze toplanır (run bitince değerlendirilir)
        self._son_manuel = []                # kapanış sonu MANUEL giriş gereken kişiler [{"ad","sebep"}]
        self._retry_btn_gizle()              # önceki koşunun tekrar-dene tuşu varsa kaldır
        for b in self._run_btns:
            b.configure(state="disabled")
        for b in self._stop_btns:
            b.configure(state="normal")
        self._set_status(calisiyor_mesaji, "running")
        env = os.environ.copy()
        env["DGS_CDP"] = CDP_URL
        env["PYTHONUNBUFFERED"] = "1"

        def worker():
            try:
                # stdin=DEVNULL ŞART: motorlar sonunda input("ENTER ile kapat...") çağırıyor —
                # stdin verilirse alt-süreç orada ASILI KALIR (EOFError yakalanıyor, TTY beklemez).
                # cwd=DATA_DIR ŞART: resume/done dosyaları göreli yolla yazılıyor; donmuş exe'de
                # SCRIPT_DIR geçici _MEIPASS olduğu için oraya yazılan resume UÇAR → mükerrer kayıt.
                # encoding ŞART: worker UTF-8 yazıyor (izin_app._repair_stdio), ama text=True sistem
                # kodlamasıyla çözer → Türkçe Windows'ta cp1254; "Ş"nin 0x9E baytı orada TANIMSIZ →
                # UnicodeDecodeError → okuma thread'i ölür, run yarıda "Çalıştırma hatası" ile düşer.
                self.proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                             stdin=subprocess.DEVNULL, text=True, bufsize=1,
                                             encoding="utf-8", errors="replace",
                                             env=env, cwd=DATA_DIR)
                for line in self.proc.stdout:
                    if "DeprecationWarning" in line or "trace-deprecation" in line or "url.parse()" in line:
                        continue
                    self.out_q.put(line)
                    m = _BASARISIZ_RE.search(line)   # "!! BAŞARISIZ: <ad> — <sebep>" → tekrar-dene için topla
                    if m:
                        self._son_hatalar.append({"ad": m.group(1).strip(), "mesaj": m.group(2).strip()})
                    mm = _MANUEL_RE.search(line)      # kapanış sonu MANUEL listesi (JSON) → popup için sakla
                    if mm:
                        try:
                            self._son_manuel = json.loads(mm.group(1)).get("kisiler", [])
                        except Exception:  # noqa
                            pass
                self.proc.wait()
                self.out_q.put(f"\n=== İşlem bitti (çıkış kodu {self.proc.returncode}) ===\n")
            except Exception as e:  # noqa
                self.out_q.put(f"\n!! Çalıştırma hatası: {e}\n")
            finally:
                self.proc = None
                self.root.after(0, self._run_done)

        threading.Thread(target=worker, daemon=True).start()

    def _run_done(self):
        # ── İZİN TEKRAR-DENEME ZİNCİRİ: her kişi AYRI koşu (motorun --person'u tek ad alıyor) ──
        # Biten kişinin sonucu yazılır, kuyrukta kişi varsa sıradaki başlar; bitince özet gösterilir.
        if self._izin_retry_aktif:
            ad = self._izin_retry_aktif
            self._izin_retry_aktif = None
            hata = next((h for h in self._son_hatalar
                         if _fold_tr(h["ad"]) in _fold_tr(ad) or _fold_tr(ad) in _fold_tr(h["ad"])), None)
            self._izin_retry_sonuc.append({"ad": ad, "ok": hata is None,
                                           "mesaj": hata["mesaj"] if hata else ""})
            self._log(("   ✓ " if hata is None else "   ✗ ") + f"{ad}\n")
            if self._izin_retry_kuyruk:
                self._izin_retry_ilerle()
                return
            for b in self._run_btns:
                b.configure(state="normal")
            for b in self._stop_btns:
                b.configure(state="disabled")
            self._izin_retry_ozet()
            return

        # ── OTOMATİK KAPANIŞ: gerçek DGS-giriş bittiyse SGK raporuyla doğrula + eksikleri retry ──
        # Rapor-tabanlı kapanış, koşunun kendi bildirdiği hatalardan ÜSTÜN: sessizce Gün<30 kalanı
        # (ör. giriş "başarılı" dedi ama SGK'da eksik) da yakalar. _retry_baglami yalnız gerçek giriş-
        # commit'te set edilir; _kapanis_asamasi ise biten koşunun kapanışın kendisi olduğunu işaretler.
        # Yalnız TAM giriş (person/limit YOK) sonrası otomatik kapanış — kısmi girişte tüm parkı
        # geri-alınamaz retry etmesin (o durumda aşağıdaki eski failure-popup devrede kalır).
        if self._retry_baglami and self._retry_baglami.get("tam") and not self._kapanis_asamasi:
            rb = self._retry_baglami
            self._retry_baglami = None
            self._kapanis_asamasi = True
            self._log("\n[KAPANIŞ] ▶ Giriş bitti → SGK raporuyla kapanış kontrolü + eksik-retry başlıyor…\n")
            kcmd = izin_frozen.worker_cmd("dgs") + ["--park", rb["park"], "kapanis", "--excel", rb["excel"]]
            yol = self._istisna_dosyasi(rb["park"])              # istisna kişiler kapanış-retry'sinden de hariç
            if self.istisna and os.path.exists(yol):
                kcmd += ["--exclude-file", yol]
            # ANA-GİRİŞTE KALICI HATA (mükerrer riski / uzaktan-çalışma / #N/A …) alanları kapanış RETRY ETMESİN
            # → skip-file ile geçir (kapanış onları hiç denemez, doğrudan MANUEL'e yazar). _spawn _son_hatalar'ı
            # sıfırladığı için LİSTEYİ ŞİMDİ topla.
            skip = [{"ad": h["ad"], "sebep": self._hata_yorumla(h["mesaj"])[0]}
                    for h in self._son_hatalar if not self._hata_yorumla(h["mesaj"])[1]]
            if skip:
                skip_yol = os.path.join(DATA_DIR, f"dgs_kapanis_skip_{rb['park']}_{dgs_donem_ay()}.json")
                try:
                    with open(skip_yol, "w", encoding="utf-8") as f:
                        json.dump(skip, f, ensure_ascii=False)
                    kcmd += ["--skip-file", skip_yol]
                    self._log(f"[KAPANIŞ] Ana-girişte kalıcı-hata alan {len(skip)} kişi retry EDİLMEYECEK "
                              f"(mükerrer koruması) → doğrudan MANUEL.\n")
                except Exception:  # noqa
                    pass
            self._spawn(kcmd, "Kapanış: kontrol + eksik-retry çalışıyor")   # buton/durum'u _spawn yönetir
            return

        bu_kapanis = self._kapanis_asamasi          # biten koşu kapanış mıydı? (manuel-popup kararı için)
        self._kapanis_asamasi = False
        for b in self._run_btns:
            b.configure(state="normal")
        for b in self._stop_btns:
            b.configure(state="disabled")
        self._on_dgs_park_change()          # resume sayacı tazelensin (kaç kişi girildi?) — status'u da yazar
        # KAPANIŞ bittiyse: otomasyonun tamamlayamadığı (MANUEL giriş gereken) kişileri popup'ta göster.
        if bu_kapanis and self._son_manuel:
            self._retry_btn_gizle()
            self._basarisiz_ozet_logla("MANUEL GİRİŞ GEREKEN KİŞİLER", list(self._son_manuel))
            self._manuel_popup(list(self._son_manuel))
            self._set_status(f"Kapanış bitti — {len(self._son_manuel)} kişi MANUEL giriş gerektiriyor", "warning")
        # Koşu başarısız kişiyle bittiyse: "kaydedilemeyen kişiler" penceresi — HANGİ MODÜL OLURSA OLSUN.
        # (Eskiden yalnız _retry_baglami dolu olan kısmi DGS koşusunda açılıyordu; İZİN koşusundan sonra
        # hiç açılmıyordu. Bağlam yoksa pencere sebepleri gösterir, "Tekrar Dene" tuşu çıkmaz.)
        elif self._son_hatalar:
            toplam = len(self._son_hatalar)
            denenebilir = sum(1 for h in self._son_hatalar if self._hata_yorumla(h["mesaj"])[1])
            self._basarisiz_ozet_logla("KAYDEDİLEMEYEN KİŞİLER", list(self._son_hatalar))
            self._retry_btn_goster(toplam, denenebilir)
            self._set_status(f"{toplam} kişi kaydedilemedi — tekrar denenebilir", "warning")
            self._retry_popup(list(self._son_hatalar))
        else:
            self._retry_btn_gizle()
            self._set_status("İşlem tamamlandı" + (" — kapanış temiz ✓" if bu_kapanis else ""), "success")

    def _manuel_popup(self, kisiler: "list[dict]"):
        """Kapanış sonu: otomasyonun tamamlayamadığı kişileri 'manuel giriş gerekli' diye göster."""
        self._modal_oncesi()              # messagebox topmost log penceresinin arkasında kalmasın
        satirlar = "\n".join(f"   •  {k.get('ad','?')}\n        → {k.get('sebep','')}" for k in kisiler)
        messagebox.showwarning(
            "Manuel giriş gerekli",
            f"Otomasyon bu {len(kisiler)} kişiyi TAMAMLAYAMADI — portalda ELLE girilmesi gerekiyor:\n\n"
            f"{satirlar}\n\n"
            "Tipik sebepler: uzaktan-çalışma istisnası · PDKS çakışması · Excel'de #N/A proje.\n"
            "Bu kişilere gir → puantajı elle tamamla → onaya gönder.")

    # ---------- başarısız kişiler: sınıflandır + tekrar dene ----------
    def _retry_btn_gizle(self):
        self._retry_durum = None
        self._retry_btnleri_tazele()

    def _retry_btn_goster(self, toplam: int, denenebilir: int):
        """Kalıcı 'kaydedilemedi' tuşunu göster. denenebilir=0 ise 'düzeltilmeli' rengi/metni."""
        self._retry_durum = {"toplam": toplam, "denenebilir": denenebilir}
        self._retry_btnleri_tazele()

    def _retry_btnleri_tazele(self):
        """'Kaydedilemedi' tuşunu İKİ yerde birden kur: ana penceredeki log kartı + log penceresi.

        Neden iki kopya: akış artık ayrı pencerede (bkz. _build_log). Operatör koşuyu orada izliyor;
        tekrar denemek için ana pencereyi aramak zorunda kalmasın. Durum tek yerde (`_retry_durum`)
        tutulur, iki tuş da ondan beslenir — log penceresi sonradan açılsa bile tuş doğru görünür.
        """
        d = getattr(self, "_retry_durum", None)
        metin = renk = yazi = None
        if d and d["denenebilir"] > 0:
            metin = f"🔄  {d['toplam']} kişi kaydedilemedi — {d['denenebilir']} kişiyi TEKRAR DENE"
            renk, yazi = UI["warning"], "#1A1206"
        elif d:
            metin = f"⚠  {d['toplam']} kişi kaydedilemedi — düzeltilmeli (detayı gör)"
            renk, yazi = UI["danger"], "white"
        hedefler = (
            (getattr(self, "retry_btn", None), dict(fill="x", pady=(10, 0))),
            (getattr(self, "_log_retry_btn", None), None),   # None → grid (log penceresi grid kullanıyor)
        )
        for btn, yerlesim in hedefler:
            if btn is None:
                continue
            try:
                if not btn.winfo_exists():
                    continue
                if d:
                    btn.configure(text=metin, fg_color=renk, text_color=yazi)
                    if yerlesim is None:
                        btn.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 14))
                    else:
                        btn.pack(**yerlesim)
                else:
                    btn.grid_remove() if yerlesim is None else btn.pack_forget()
            except Exception:  # noqa
                pass

    def _basarisiz_ozet_logla(self, baslik: str, kayitlar: "list[dict]"):
        """Koşu bitiminde 'kim kaydedilemedi?' sorusunu LOG'un içinde, KIRMIZI ve toplu olarak yanıtla.

        Neden: akış yüzlerce satır; tek tek '!! BAŞARISIZ' satırları yukarıda kaybolup gidiyor.
        Operatörün run sonunda tek bakışta göreceği liste burada basılır — önce isimler yan yana,
        sonra kişi başına sebep + 'tekrar denenebilir mi' işareti.
        kayitlar: [{"ad","mesaj"}] (motorun ham sebebi, yorumlanır) | [{"ad","sebep"}] (hazır metin).
        """
        if not kayitlar:
            return
        cizgi = "─" * 60
        satirlar = ["", cizgi,
                    f"!!  {baslik} — {len(kayitlar)} KİŞİ",
                    "!!  " + "  ·  ".join(k.get("ad", "?") for k in kayitlar),
                    cizgi]
        for i, k in enumerate(kayitlar, 1):
            if "mesaj" in k:
                aciklama, tekrar = self._hata_yorumla(k.get("mesaj", ""))
            else:
                aciklama, tekrar = (k.get("sebep", "") or "sebep bildirilmedi"), False
            satirlar.append(f"!!  {i:>2}. {k.get('ad', '?')}")
            satirlar.append(f"        {'🔄 tekrar denenebilir' if tekrar else '🔧 önce düzeltilmeli'}"
                            f"  —  {aciklama}")
        satirlar += [cizgi, ""]
        self._log("\n".join(satirlar) + "\n", tag="danger")

    def _retry_ac(self):
        """Kalıcı tuş → 'Kaydedilemeyen kişiler' penceresini yeniden aç (son koşunun başarısızlarıyla)."""
        if self._son_hatalar:
            self._retry_popup(list(self._son_hatalar))

    def _hata_yorumla(self, mesaj: str) -> "tuple[str, bool]":
        """Motorun ham sebep metnini (kullanıcı-dostu açıklama, tekrar_denenebilir?) ikilisine çevirir.

        Ayrım kritik: VERİ hatası (Excel/kaynak düzeltilmeden retry BOŞUNA döner) ≠ GEÇİCİ hata
        (portal timing/doğrulama — retry çözer). Bilinmeyen → güvenli tarafta 'tekrar denenebilir'."""
        m = _fold_tr(mesaj)   # türkçe-duyarsız, küçük harf (ör. "öneri çıkmadı" → "oneri cikmadi")
        # ——— VERİ HATASI: önce düzelt, retry işe yaramaz ———
        if "#n/a" in m or "oneri cikmadi" in m or "autocomplete" in m or "sonuc yok" in m:
            return ("Proje bilgisi eksik/geçersiz (Excel'de #N/A) — önce projeyi düzelt", False)
        if "kimlik dogrulamasi basarisiz" in m or ("kimlik" in m and "eslesmedi" in m):
            return ("Kişi kimliği eşleşmedi — isim/TC kontrol et", False)
        if "proje" in m and ("secilemedi" in m or "bulunamadi" in m):
            return ("Proje seçilemedi — Excel'deki proje adını kontrol et", False)
        if "haritada yok" in m and "varsay" not in m:
            return ("Puantaj/mesai kaydı bulunamadı", False)
        # ——— MÜKERRER-RİSKLİ / ÇÖZÜLEMEZ: RETRY ETME (yeniden girmek mükerrer yaratır / eklenecek gün yok) ———
        # "başarı mesajı görülemedi — kayıt gitmiş olabilir" → yeniden girmek MÜKERRER riski; "en az 1 gün
        # seçimi" → uzaktan-çalışan, tüm günler PDKS, seçilecek gün yok. İkisi de ELLE girilmeli.
        if "mukerrer" in m or "gitmis olabilir" in m or "basari ile kayit" in m or "en az 1 gun" in m:
            return ("Kayıt belirsiz/çözülemez (uzaktan-çalışma ya da tüm günler PDKS) — ELLE gir; "
                    "yeniden deneme MÜKERRER yaratabilir, önce portal listesini kontrol et", False)
        # ——— PORTAL ÇAKIŞMASI: o tarih(ler)de zaten kayıt var (yanlış-pencere) → portal girişi REDDETTİ ———
        # Genelde ZARARSIZ: kişi çoğu zaman zaten Gün=30 (girişe ihtiyacı yok), portal mükerreri engelliyor.
        # Ama motoru takar (SweetAlert overlay tıklamayı yer). Retry bazen geçer; geçmezse SGK 'kontrol'ü
        # ile Gün=30 mu diye bak — muhtemelen tamamdır. (Kullanıcı 2026-07-15 canlı: bu popup takıyordu.)
        if ("pdks" in m and "mevcut" in m) or ("reddedildi" in m and "portal" in m) or "kayit mevcut" in m:
            return ("Portalda o tarihte zaten kayıt var — çakışma; retry geçmezse SGK kontrol et (çoğunlukla zaten Gün=30)", True)
        # ——— GEÇİCİ: portal timing/doğrulama; retry genelde çözer ———
        if "taze dialog gelmedi" in m or "stale" in m or "damga" in m:
            return ("Portal doğrulaması takıldı (geçici) — tekrar denemede geçebilir", True)
        if "timeout" in m or "zaman asimi" in m or "baglant" in m or "yanit vermedi" in m or "erisilemedi" in m:
            return ("Portal yanıt vermedi (geçici) — tekrar denemede geçebilir", True)
        # ——— BİLİNMEYEN: retry teklif et, ham sebebi kısalt (ss:… ekran-görüntüsü ekini at) ———
        kisa = mesaj.split("(ss:")[0].split("(ss")[0].strip()
        return (kisa[:110] or "Bilinmeyen hata", True)

    def _retry_popup(self, hatalar: "list[dict]"):
        """Run bitiminde: kaydedilemeyen kişileri sebebiyle listele, geçici olanlar için 'Tekrar Dene' sun."""
        yorumlu = []
        for h in hatalar:
            aciklama, tekrar = self._hata_yorumla(h["mesaj"])
            yorumlu.append({"ad": h["ad"], "aciklama": aciklama, "tekrar": tekrar})
        denenebilir = [y["ad"] for y in yorumlu if y["tekrar"]]
        duzeltilecek = sum(1 for y in yorumlu if not y["tekrar"])

        self._modal_oncesi()          # topmost log penceresi modal'ın önüne geçmesin
        win = ctk.CTkToplevel(self.root, fg_color=UI["bg"])
        win.title("Kaydedilemeyen kişiler")
        geo, mw, mh = self._ust_pencere_plani(660, 540, 500, 380)   # bkz. _ust_pencere_plani
        win.geometry(geo)
        win.minsize(mw, mh)
        win.transient(self.root)
        try:
            win.grab_set()
        except Exception:  # noqa
            pass

        head = ctk.CTkFrame(win, fg_color=UI["surface"], corner_radius=18)
        head.pack(fill="x", padx=16, pady=(16, 0))
        ctk.CTkLabel(head, text=f"⚠  {len(hatalar)} kişi kaydedilemedi",
                     text_color=UI["warning"], font=("Helvetica", 15, "bold"),
                     anchor="w").pack(fill="x", padx=18, pady=(12, 2))
        ozet = f"🔄 {len(denenebilir)} kişi tekrar denenebilir"
        if duzeltilecek:
            ozet += f"     🔧 {duzeltilecek} kişi önce düzeltilmeli"
        ctk.CTkLabel(head, text=ozet, text_color=UI["muted"], font=("Helvetica", 11),
                     anchor="w").pack(fill="x", padx=18, pady=(0, 12))

        lst = ctk.CTkScrollableFrame(win, fg_color=UI["input"], corner_radius=16)
        lst.pack(fill="both", expand=True, padx=16, pady=16)
        for y in yorumlu:
            row = ctk.CTkFrame(lst, fg_color="transparent")
            row.pack(fill="x", padx=6, pady=5)
            renk = UI["success"] if y["tekrar"] else UI["warning"]
            ctk.CTkLabel(row, text=("🔄" if y["tekrar"] else "🔧"), font=("Helvetica", 14),
                         width=30).pack(side="left")
            col = ctk.CTkFrame(row, fg_color="transparent")
            col.pack(side="left", fill="x", expand=True)
            ctk.CTkLabel(col, text=y["ad"], text_color=UI["text"], font=("Helvetica", 12, "bold"),
                         anchor="w").pack(fill="x")
            ctk.CTkLabel(col, text=y["aciklama"], text_color=renk, font=("Helvetica", 10),
                         anchor="w", justify="left", wraplength=520).pack(fill="x")

        btnrow = ctk.CTkFrame(win, fg_color="transparent")
        btnrow.pack(fill="x", padx=16, pady=(0, 16))
        ctk.CTkButton(btnrow, text="Kapat", command=win.destroy, width=104, height=40, corner_radius=16,
                      fg_color=UI["card"], hover_color=UI["card_hover"], text_color=UI["text"],
                      font=("Helvetica", 10, "bold")).pack(side="right")
        # "Tekrar Dene" YALNIZ bağlam varken çıkar (hangi park/Excel/ayarla koşulacağı bilinmeli).
        # Bağlam modülü söyler: DGS tek koşuda --file listesiyle, İZİN kişi kişi zincirle dener.
        b = self._retry_baglami
        if denenebilir and b:
            calistir = self._izin_retry if b.get("modul") == "izin" else self._dgs_retry
            ctk.CTkButton(btnrow, text=f"🔄  Tekrar Dene ({len(denenebilir)} kişi)",
                          command=lambda: calistir(denenebilir, win),
                          width=220, height=40, corner_radius=16,
                          fg_color=UI["primary"], hover_color=UI["primary_hover"], text_color="white",
                          font=("Helvetica", 11, "bold")).pack(side="right", padx=(0, 10))

    # ---------- İZİN: tekrar deneme (kişi kişi zincir) ----------
    def _izin_retry(self, adlar: "list[str]", win=None):
        """Seçilen kişileri TEK TEK yeniden koş.

        Neden zincir: izin motorunun `--person` argümanı TEK ad/T.C. alıyor (liste dosyası yok).
        Parkın tamamını yeniden koşmak ise başarısız olmayanları da riske atardı — bu yüzden yalnız
        seçilen adlar, sırayla, kendi koşularında denenir. İzin motoru başarısız kişide "KAYDEDİLMEDİ"
        diyerek çıkıyor (izin_poc.py:800/807) → tekrar denemek mükerrer yaratmaz.
        """
        self._modal_oncesi()              # onay kutusu topmost log penceresinin arkasında kalmasın
        b = self._retry_baglami
        if not b:
            self._sessizce(lambda: win and win.destroy())   # win=None → çağrı log penceresinden geldi
            return
        if self.proc is not None:
            messagebox.showinfo("Çalışıyor", "Bir işlem zaten sürüyor. Önce bitmesini bekle veya Durdur.")
            return
        liste = "\n".join(f"• {a}" for a in adlar)
        if not messagebox.askyesno(
                "Tekrar deneme onayı",
                f"{b['park']} parkında {len(adlar)} kişi TEKRAR denenecek (gerçek kayıt):\n\n{liste}\n\n"
                "Kişiler sırayla, ayrı ayrı koşulur. Devam edilsin mi?"):
            return
        self._sessizce(lambda: win and win.destroy())
        self._izin_retry_kuyruk = list(adlar)
        self._izin_retry_sonuc = []
        self._log(f"\n🔄 TEKRAR DENEME: {len(adlar)} kişi → {', '.join(adlar)}\n")
        self._izin_retry_ilerle()

    def _izin_retry_ilerle(self):
        """Kuyruktaki sıradaki kişiyi koş (koşu bitince _run_done zinciri sürdürür)."""
        ad = self._izin_retry_kuyruk.pop(0)
        self._izin_retry_aktif = ad
        kalan = len(self._izin_retry_kuyruk)
        cmd = list(self._retry_baglami["cmd"]) + ["--person", ad]
        self._spawn(cmd, f"Tekrar deneniyor: {ad}" + (f" (+{kalan} kişi sırada)" if kalan else ""))

    def _izin_retry_ozet(self):
        """Zincir bitti: kaç kişi kurtuldu, kim hâlâ kaydedilemedi → sonuç penceresi."""
        sonuc = self._izin_retry_sonuc
        self._izin_retry_sonuc = []
        olan = [s for s in sonuc if s["ok"]]
        kalan = [{"ad": s["ad"], "mesaj": s["mesaj"]} for s in sonuc if not s["ok"]]
        self._log(f"\n🔄 Tekrar deneme bitti: {len(olan)} kurtarıldı, {len(kalan)} hâlâ kaydedilemedi.\n")
        self._basarisiz_ozet_logla("HÂLÂ KAYDEDİLEMEYEN KİŞİLER", list(kalan))
        self._on_dgs_park_change()                     # resume sayaçları tazelensin
        if not kalan:
            self._modal_oncesi()          # aşağıdaki showinfo topmost pencerenin arkasında kalmasın
            self._retry_btn_gizle()
            self._son_hatalar = []
            self._set_status(f"Tekrar deneme tamam — {len(olan)} kişi kaydedildi", "success")
            messagebox.showinfo("Tekrar deneme bitti",
                                f"Denenen {len(sonuc)} kişinin hepsi kaydedildi ✓")
            return
        # Hâlâ kalanlar: aynı pencereyi yeniden sun (tekrar denenebilir olanlar için tuş yine çıkar)
        self._son_hatalar = kalan
        denenebilir = sum(1 for k in kalan if self._hata_yorumla(k["mesaj"])[1])
        self._retry_btn_goster(len(kalan), denenebilir)
        self._set_status(f"{len(olan)} kurtarıldı · {len(kalan)} kişi hâlâ kaydedilemedi", "warning")
        self._retry_popup(list(kalan))

    def _dgs_retry(self, adlar: "list[str]", win=None):
        """Başarısız kişileri geçici bir --file listesiyle yeniden koş (--file done-skip'i atlar → tam bu kişiler).

        win=None → çağrı log penceresindeki tuştan geldi (kapatılacak bir liste penceresi yok).
        """
        self._modal_oncesi()              # onay kutusu topmost log penceresinin arkasında kalmasın
        b = self._retry_baglami
        if not b:
            self._sessizce(lambda: win and win.destroy())
            return
        if self.proc is not None:   # iki koşu aynı Chrome oturumunu paylaşır → mükerrer kayıt riski
            messagebox.showinfo("Çalışıyor", "Bir işlem zaten sürüyor. Önce bitmesini bekle veya Durdur.")
            return
        liste = "\n".join(f"• {a}" for a in adlar)
        if not messagebox.askyesno(
                "Tekrar deneme onayı",
                f"{b['park']} portalında {len(adlar)} kişi TEKRAR denenecek "
                f"(gerçek kayıt{' + onaya gönderme' if b['onayla'] else ''} — GERİ ALINAMAZ):\n\n"
                f"{liste}\n\nDevam edilsin mi?"):
            return
        self._sessizce(lambda: win and win.destroy())
        yol = os.path.join(DATA_DIR, f"dgs_retry_{b['park']}_{dgs_donem_ay()}.txt")
        try:
            with open(yol, "w", encoding="utf-8") as f:
                f.write("\n".join(adlar) + "\n")
        except OSError as e:
            messagebox.showerror("Yazılamadı", f"Tekrar-dene listesi yazılamadı:\n{e}")
            return
        cmd = izin_frozen.worker_cmd("dgs") + ["--park", b["park"], "giris",
              "--excel", b["excel"], "--file", yol, "--commit"]
        if b["onayla"]:
            cmd.append("--onayla")
        if b["destek"]:
            cmd.append("--include-destek")
        iyol = self._istisna_dosyasi(b["park"])   # istisna kişiye retry'de de dokunma (tutarlılık)
        if self.istisna and os.path.exists(iyol):
            cmd += ["--exclude-file", iyol]
        self._log(f"\n🔄 TEKRAR DENEME: {len(adlar)} kişi → {', '.join(adlar)}\n")
        self._spawn(cmd, "Başarısızlar tekrar deneniyor")   # _retry_baglami duruyor → zincirleme retry olur

    def _stop(self):
        # Zinciri de kes: kuyruk boşaltılmazsa biten koşunun ardından SIRADAKİ kişi kendiliğinden başlar.
        self._izin_retry_kuyruk = []
        if self.proc:
            try:
                self.proc.terminate()
                self._log("\n■ Durduruldu (kaldığı yerden resume ile devam edilebilir).\n")
                self._set_status("İşlem kullanıcı tarafından durduruldu", "warning")
            except Exception as e:  # noqa
                self._log(f"Durdurma hatası: {e}\n")


def main():
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)   # yüksek DPI'da bulanık ölçeklemeyi engelle
        except Exception:
            pass
        try:
            # Taskbar bu uygulamayı KENDİ ikonuyla göstersin (yoksa Windows onu jenerik grup ikonuna
            # bindirir → yanlış/soluk ikon). Pencere OLUŞMADAN ÖNCE ayarlanmalı. Kimlik mac bundle ile aynı.
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("com.fev.izin")
        except Exception:
            pass
    root = ctk.CTk()
    IzinGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
