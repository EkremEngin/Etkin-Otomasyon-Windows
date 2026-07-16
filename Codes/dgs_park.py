#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PARK MODÜLÜ — her teknopark için AYRI script + ortak motor (modüler exe mimarisi)
================================================================================
Kullanıcı kararı (2026-07-14): "her site birbirinden az bişey farklı, her teknoparkınki ayrı olmalı,
modüler çalışan bir exe olmalı."

YAPI:
    dgs_tpi.py · dgs_bv.py · dgs_tpiz.py · dgs_yildiz.py · dgs_ulutek.py   ← PARK SCRIPTLERİ (çalıştırdığın)
        └── her biri: o parkın portal/kod/port ayarları + O PARKA ÖZEL davranış farkları (overrides)
    dgs_park.py (bu dosya)                                                 ← park kayıt defteri + çalıştırıcı
    dgs_poc.py / dgs_onaya.py / dgs_rapor_kontrol.py                       ← ORTAK MOTOR (kanıtlanmış)

NEDEN MOTOR ORTAK: park farkları bugüne kadar hep AYAR farkı çıktı (portal URL, dönem, lokasyon kodu),
davranış farkı değil — Mayıs'ta aynı motor 5 parkta 126/126 onay aldı. Bugünkü menü bug'ı (açık PERSONEL
dropdown'ı 'Onaya Gönder'in üstüne binip gerçek-click'i yiyordu) TPI'de yakalandı ama TÜM parkları
etkiliyordu → tek yerde düzeltildi, hepsi kurtuldu. Park'a ÖZEL bir davranış gerekirse o parkın kendi
scriptindeki `overrides()` içine yazılır — diğer parklar ETKİLENMEZ.

KULLANIM (her park scripti aynı arayüz):
    python3 dgs_tpi.py --excel "<xlsx>"                      # DRY-RUN (yazmaz; UI/dönem/grid doğrular)
    python3 dgs_tpi.py --excel "<xlsx>" --limit 1 --commit   # tek kişi TASLAK (test)
    python3 dgs_tpi.py --excel "<xlsx>" --onayla --commit    # BULK giriş+onay (üretim)
    python3 dgs_tpi.py onay    --excel "<xlsx>" --commit     # safety-net: kalan taslakları onaya gönder
    python3 dgs_tpi.py kontrol --excel "<xlsx>"              # SGK Gün=30 kapanış kontrolü (salt-okur)

Dönem OTOMATİK (bugünden bir önceki ay) → --donem "HAZİRAN 2026" ile override edilir.
Excel sayfa adı dönemden türetilir ("Haziran") → --sheet ile override edilir.
`--no-schedule` park scriptlerinde VARSAYILAN AÇIK (kanıtlanmış reçete); `--schedule` ile kapatılır.
CDP portu: DGS_CDP env veya park.cdp (varsayılan 9222 — paralel park için farklı port ver).
"""
from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass


@dataclass(frozen=True)
class Park:
    code: str          # DGS Excel "Lokasyonu SGK" (F kolonu) DEĞERİ — done/resume dosya adlarında da bu kullanılır
    ad: str
    portal_url: str
    cdp: str = "http://localhost:9222"
    notlar: str = ""


# Excel F kolonundaki kodlarla BİREBİR (canlı doğrulandı: TPI · BV · TPIz · Yıldız · Ulutek · ARI · ODTÜ).
# ARI + ODTÜ = Teknoera (Arge Portal DEĞİL, Excel ile toplu yükleniyor) → Playwright otomasyonu YOK.
PARKS: dict[str, Park] = {
    "TPI": Park("TPI", "Teknopark İstanbul", "https://argeportal.teknoparkistanbul.com.tr/"),
    "BV": Park("BV", "Bilişim Vadisi", "https://argeportal.bilisimvadisi.com.tr/"),
    "TPIz": Park("TPIz", "Teknopark İzmir", "https://argeportal.teknoparkizmir.com.tr/"),
    "Yıldız": Park("Yıldız", "Yıldız Teknopark (Davutpaşa)", "https://argeportal.yildizteknopark.com.tr/"),
    "Ulutek": Park("Ulutek", "Ulutek Teknopark (Bursa)", "https://argeportal.ulutek.com.tr/"),
}

MODLAR = ("giris", "onay", "kontrol", "liste", "kapanis")

# Her motorun GERÇEKTEN kabul ettiği bayraklar (argparse "unrecognized arguments" ile patlamasın diye
# park runner ortak argv'yi moda göre SÜZER). Değer alan bayraklar → _DEGERLI.
_KABUL = {
    "giris": {"--excel", "--sheet", "--donem", "--person", "--file", "--limit", "--commit",
              "--dump-names", "--assume-std", "--include-destek", "--onayla", "--no-schedule", "--lokasyon"},
    "onay": {"--excel", "--sheet", "--donem", "--lokasyon", "--limit", "--commit", "--dry-run"},
    "kontrol": {"--lokasyon", "--sheet", "--donem", "--write-fix", "--threshold"},
    "liste": set(),                      # motora gitmez; dgs_park kendi işler
    "kapanis": {"--excel", "--sheet", "--donem"},   # orkestrasyon; alt-çağrılara (kontrol/giris) geçer
}
_DEGERLI = {"--excel", "--sheet", "--donem", "--lokasyon", "--person", "--file", "--limit",
            "--dump-names", "--threshold"}
_ORTAK = {"--help"}     # her argparse'ın kabul ettiği; süzgeçte yenmesin (yoksa `--help` yerine "--excel şart" hatası gelir)


def _filtrele(argv: list[str], mod: str) -> list[str]:
    """argv'den o modun TANIMADIĞI bayrakları (ve varsa değerlerini) at."""
    kabul, out, i = _KABUL[mod] | _ORTAK, [], 0
    while i < len(argv):
        a = argv[i]
        if a.startswith("--"):
            deger = argv[i + 1] if (a in _DEGERLI and i + 1 < len(argv)) else None
            if a in kabul:
                out.append(a)
                if deger is not None:
                    out.append(deger)
            i += 2 if deger is not None else 1
        else:
            i += 1   # başıboş konumsal argüman → yut (motorlar konumsal almıyor)
    return out


def run(park: Park, overrides=None):
    """Park scriptlerinin çağırdığı tek giriş noktası. Modu seçer, ortak motoru O PARKIN ayarlarıyla koşar."""
    import dgs_poc as D

    argv = sys.argv[1:]
    mod = "giris"
    if argv and argv[0] in MODLAR:
        mod, argv = argv[0], argv[1:]

    # --- istisna kişiler: motorların TANIMADIĞI bayrak → _filtrele'ye gitmeden burada sökülür ---
    istisna_yolu, argv = _pop_deger(argv, "--exclude-file")

    # --- parkı motora bağla (portal + CDP) ---
    D.CONFIG["portal_url"] = park.portal_url
    D.CONFIG["cdp_url"] = os.environ.get("DGS_CDP", park.cdp)

    # --- dönem/sheet: verilmediyse otomatik türet (dgs_onaya/kontrol'e de aynı değeri geçir) ---
    donem = _arg_value(argv, "--donem")
    label, _regex, ay_no = D.donem_label_and_regex(donem)
    if "--sheet" not in argv:
        argv += ["--sheet", D.TR_AYLAR_TITLE[ay_no]]
    if "--donem" not in argv:
        argv += ["--donem", label]
    sheet = _arg_value(argv, "--sheet")

    # --- lokasyon bu park (kullanıcı elle veremesin — yanlış portala yazma riski) ---
    argv += ["--lokasyon", park.code]

    # --- park'a ÖZEL davranış farkları (varsa) ---
    if overrides:
        overrides(D)

    print(f"[PARK] {park.code} — {park.ad} | {park.portal_url} | dönem={label} | mod={mod} "
          f"| CDP={D.CONFIG['cdp_url']}", flush=True)

    if mod == "liste":                       # PORTALA BAĞLANMAZ — GUI'nin istisna seçicisini besler
        _liste(D, park, _arg_value(argv, "--excel"), sheet)
        return

    if mod == "kapanis":                     # KAPANIŞ DÖNGÜSÜ — kontrol + eksikleri tekrar koş
        _kapanis(D, park, argv, sheet, label, istisna_yolu)   # istisna_yolu: yukarıda --exclude-file'dan söküldü
        return

    istisna = _istisna_yukle(istisna_yolu)
    if istisna:
        _istisna_uygula(D, mod, istisna)

    if mod == "giris":
        # kanıtlanmış reçete: --no-schedule (40-sayfa Personel Listesi okuması çoğu parkta boş dönüyor;
        # grid temiz-gün==09:00 çapraz kontrolü anomaliyi zaten yakalıyor). --schedule ile kapatılır.
        istek_schedule = "--schedule" in argv
        argv = _filtrele(argv, "giris")
        if not istek_schedule and "--no-schedule" not in argv:
            argv += ["--no-schedule"]
        # --include-destek motora YETMİYOR: dgs_poc bulk hedef listesini HER KOŞULDA is_ar_ge ile süzer
        # (satır ~1391; bayrak yalnız --person/--file yolunda işlemeye izin verir) → Destek'çiler bulk'ta
        # sessizce dışarıda kalıyordu (Mayıs'ta 14 Destek --person ile TEK TEK girildi). Bulk'a dahil etmek
        # için is_ar_ge burada herkese-True yapılır; done-skip AYNEN çalışır (bypass yalnız --person/--file).
        if "--include-destek" in argv:
            D.is_ar_ge = lambda p: True
            print("[DESTEK] Destek personeli lokasyon-bulk hedefine DAHİL "
                  "(motor varsayılanı yalnız Ar-Ge'ydi).", flush=True)
        _dispatch(D.main, argv)

    elif mod == "onay":
        import dgs_onaya as O
        _dispatch(O.main, _filtrele(argv, "onay"))

    elif mod == "kontrol":
        import dgs_rapor_kontrol as K
        _dispatch(K.main, _filtrele(argv, "kontrol"))


def _dispatch(fn, argv: list[str]):
    sys.argv = [sys.argv[0]] + argv
    fn()


def _arg_value(argv: list[str], flag: str) -> str | None:
    return argv[argv.index(flag) + 1] if flag in argv and argv.index(flag) + 1 < len(argv) else None


def _pop_deger(argv: list[str], flag: str) -> tuple[str | None, list[str]]:
    """Değerli bayrağı argv'den SÖKÜP al (motorlar tanımadığı için _filtrele'ye hiç ulaşmamalı)."""
    if flag not in argv:
        return None, argv
    i = argv.index(flag)
    deger = argv[i + 1] if i + 1 < len(argv) else None
    del argv[i: i + 2]
    return deger, argv


# ==========================================================================
# İSTİSNA KİŞİLER — otomasyonun DOKUNMAYACAĞI kişiler
# ==========================================================================
# Dosya biçimi (GUI yazar, elle de düzenlenebilir); satır başına bir kişi:
#     12345678901|AHMET YILMAZ      ← önerilen (hem T.C. hem ad)
#     12345678901                   ← yalnız T.C.  (DİKKAT: 'onay' modu portalda ada göre eşler → ad da yaz)
#     AHMET YILMAZ                  ← yalnız ad
#     # ile başlayan satırlar yorumdur
def _tc_norm(s) -> str:
    return "".join(ch for ch in str(s or "") if ch.isdigit())


def _istisna_yukle(yol: str | None):
    if not yol:
        return None
    if not os.path.exists(yol):
        _hata(f"istisna dosyası bulunamadı: {yol}")
    tcs, adlar = set(), []
    with open(yol, encoding="utf-8") as f:
        for satir in f:
            satir = satir.split("#", 1)[0].strip()
            if not satir:
                continue
            tc, _, ad = satir.partition("|")
            tc, ad = tc.strip(), ad.strip()
            if not ad and not _tc_norm(tc):      # tek alan var ve rakam değil → isimdir
                ad, tc = tc, ""
            if _tc_norm(tc):
                tcs.add(_tc_norm(tc))
            if ad:
                adlar.append(ad)
    return (tcs, adlar) if (tcs or adlar) else None


def _istisna_uygula(D, mod: str, istisna):
    """PRIME scriptlere (dgs_poc/dgs_onaya) DOKUNMADAN, onların KİŞİ KAYNAĞINI sarmalar:

        giris → dgs_poc.read_excel       : istisna kişiler hiç `targets`a girmez → KAYIT AÇILMAZ.
        onay  → dgs_onaya.read_list_rows : portal listesinde hiç GÖRÜNMEZLER → onaya gönderilmez.
                (portal listesi T.C. taşımıyor → burada eşleşme ADLA yapılır; GUI dosyaya adı da yazar.)
        kontrol: salt-okur, portala yazmıyor → filtre gereksiz.

    Neden sarmalama: dgs_poc/dgs_onaya PRIME (bkz. PRIME_SCRIPTS.md) — canlı kanıtlanmış, dokunulmaz.
    """
    import dgs_onaya as O                     # fold(): portal listesiyle AYNI Türkçe-duyarsız normalize
    tcs, adlar = istisna
    adlar_fold = {O.fold(a) for a in adlar if a}

    def istisna_mi(tc, ad) -> bool:
        t = _tc_norm(tc)
        if t and t in tcs:                    # T.C. esas (evlilik/kızlık soyadı farkına dayanıklı)
            return True
        a = O.fold(ad or "")
        return bool(a) and a in adlar_fold

    print(f"[İSTİSNA] {len(tcs)} T.C. + {len(adlar)} ad yüklendi → bu kişilere DOKUNULMAYACAK.", flush=True)

    if mod == "giris":
        orijinal = D.read_excel

        def read_excel_istisnasiz(path, sheet="Mayıs"):
            allp = orijinal(path, sheet)
            atlanan = [k for k, p in allp.items() if istisna_mi(p.tc, p.ad_soyad)]
            for k in atlanan:
                print(f"[İSTİSNA] ATLANDI: {allp[k].ad_soyad}", flush=True)
                del allp[k]
            print(f"[İSTİSNA] {len(atlanan)} kişi Excel'den düşürüldü → {len(allp)} kişi kaldı.", flush=True)
            return allp

        D.read_excel = read_excel_istisnasiz

    elif mod == "onay":
        orijinal = O.read_list_rows

        def read_list_rows_istisnasiz(page):
            rows = orijinal(page)
            kalan = [r for r in rows if r.get("ad_fold") not in adlar_fold]
            if len(kalan) != len(rows):
                print(f"[İSTİSNA] portal listesinden {len(rows) - len(kalan)} satır gizlendi "
                      f"→ onaya GÖNDERİLMEYECEK.", flush=True)
            return kalan

        O.read_list_rows = read_list_rows_istisnasiz

    else:                                     # kontrol
        print("[İSTİSNA] kontrol modu SALT-OKUR — portala yazılmıyor, filtre gerekmez.", flush=True)


def _liste(D, park: Park, excel: str | None, sheet: str):
    """GUI'nin 'İstisna Kişiler' seçicisini besler: bu parkın Excel'deki kişilerini JSON dök.

    Motorun KENDİ read_excel'ini kullanır → GUI ile motor BİREBİR aynı isimleri/T.C.'leri görür
    (ayrı bir Excel parser yazsaydık isimler kayabilir, istisna tutmayabilirdi). Portala BAĞLANMAZ.
    """
    if not excel:
        _hata("liste modu --excel ister.")
    allp = D.read_excel(excel, sheet)
    done_file = f"dgs_done_{park.code}_{sheet}.txt"          # cwd'ye göre (GUI cwd=data_dir verir)
    done = set()
    if os.path.exists(done_file):
        with open(done_file, encoding="utf-8") as f:
            done = {s.strip() for s in f if s.strip()}
    kisiler = sorted(
        ({"ad": p.ad_soyad, "tc": p.tc, "bolum": p.bolum, "arge": D.is_ar_ge(p),
          "eksik": p.eksik_puantaj, "girildi": p.ad_soyad in done}
         for p in allp.values() if p.lokasyon.upper() == park.code.upper()),
        key=lambda k: k["ad"])
    # Marker: read_excel kendi log satırını da basıyor → GUI JSON'u bu işaretle güvenle ayıklar.
    print("<<<LISTE>>>" + json.dumps({"park": park.code, "sheet": sheet, "kisiler": kisiler},
                                     ensure_ascii=False), flush=True)


def _kapanis(D, park: Park, argv: list[str], sheet: str, donem: str, istisna_yolu: str | None = None):
    """KAPANIŞ DÖNGÜSÜ (2026-07-15) — giriş sonrası kapanış kontrolü + otomatik tekrar-koşu.

    Akış: SGK raporunu 'kontrol' ile çek → ÇALIŞAN + Gün<30 (kontrol <<<EKSIK>>>'i; ayrılanlar
    zaten hariç) → o kişileri --file ile TEKRAR koş (giriş+onay) → tekrar kontrol, temiz olana dek
    (max 3 tur). 3 turda düzelmeyen ısrarcıları (grid/yanlış-pencere sorunlusu) 'ELLE BAK' raporlar.

    ⚠ GERÇEK giriş+onay (geri alınamaz). Rapor-adı ('REMZI TIRE', ASCII) → Excel tam-adı ('REMZİ TİRE')
    fold ile eşlenir: --file exact-upper eşlediğinden fold-eşleme YAPILMAZSA Türkçe İ/I yüzünden atlanır."""
    import subprocess
    import json as _json
    import re as _re
    try:
        import izin_frozen
        base = izin_frozen.worker_cmd("dgs") + ["--park", park.code]
    except Exception:
        base = [sys.executable, os.path.abspath(__file__), "--park", park.code]

    xl = _arg_value(argv, "--excel")
    ortak = ["--sheet", sheet] + (["--donem", donem] if donem else [])

    # rapor-adı (ASCII-fold) → bu PARKIN Excel'indeki TAM ad (yalnız bu park; çapraz-park çakışması olmasın)
    fold_map = {}
    if xl:
        try:
            allp = D.read_excel(os.path.expanduser(xl), sheet)
            fold_map = {D._fold_tr(k): k for k in allp
                        if allp[k].lokasyon.upper() == park.code.upper()}
        except Exception as e:                                          # noqa
            print(f"[KAPANIŞ] UYARI: Excel okunamadı ({e}) → rapor-adı eşlemesi yapılamayacak.", flush=True)

    # İSTİSNA kişiler (varsa): otomasyon onlara DOKUNMASIN → fix hedefinden çıkarılır (giriş bilinçli atlamıştı)
    istisna_fold = set()
    if istisna_yolu and os.path.exists(istisna_yolu):
        try:
            with open(istisna_yolu, encoding="utf-8") as f:
                for satir in f:
                    satir = satir.strip()
                    if satir and not satir.startswith("#"):
                        istisna_fold.add(D._fold_tr(satir.split("|")[-1].strip()))  # "TC|AD" ya da "AD"
            if istisna_fold:
                print(f"[KAPANIŞ] İstisna: {len(istisna_fold)} kişi kapanış-retry'sinden de HARİÇ.", flush=True)
        except Exception:                                               # noqa
            pass

    # ANA-GİRİŞTE KALICI HATA ALANLAR (GUI --skip-file ile verir): kapanış bunları HİÇ retry ETMESİN.
    # Neden: "başarı mesajı görülemedi — KAYIT GİTMİŞ OLABİLİR" uyarısında yeniden girmek MÜKERRER riski
    # yaratır (kural: önce listeyi kontrol et). O yüzden doğrudan MANUEL'e yazılır, hiç denenmez.
    skip_giris = {}
    skip_yolu = _arg_value(argv, "--skip-file")
    if skip_yolu and os.path.exists(skip_yolu):
        try:
            for it in _json.load(open(skip_yolu, encoding="utf-8")):
                skip_giris[D._fold_tr(it["ad"])] = {"ad": it["ad"], "sebep": it.get("sebep", "ana girişte kalıcı hata")}
            if skip_giris:
                print(f"[KAPANIŞ] Ana-giriş kalıcı-hatalı {len(skip_giris)} kişi RETRY EDİLMEYECEK "
                      f"(mükerrer koruması) → MANUEL.", flush=True)
        except Exception:                                               # noqa
            pass

    # KALICI hata = retry BOŞUNA (uzaktan-çalışma "İstenilen toplam uyuşmuyor" / pdks çakışması / #N/A /
    # proje / kimlik) → o kişiyi 1 kez dene, sonra BIRAK ve MANUEL'e yaz (3 tur boşuna dönme).
    _BAS_RE = _re.compile(r"!!\s*BAŞARISIZ:\s*(.+?)\s*[—–]\s*(.+)")
    def _kalici(reason: str) -> bool:
        r = D._fold_tr(reason)   # Türkçe-güvenli fold (İ→i + combining nokta temizle); anahtarlar ASCII-fold biçiminde
        return any(k in r for k in (
            "istenilen toplam", "pdks", "#n/a", "oneri cikmad", "proje", "kimlik", "eslesme bulunam",
            # ⚠ MÜKERRER-RİSKLİ / çözülemez (uzaktan-çalışma, tüm günler pdks): retry ETME, hemen MANUEL.
            "mukerrer", "gitmis olabilir", "basari ile kayit", "en az 1 gun"))

    MAX_TUR = 3
    vazgec, ad_map, son_eksik = dict(skip_giris), {}, []   # vazgec: fold→{ad,sebep} (kalıcı; ana-giriş skip'i dahil)
    for tur in range(1, MAX_TUR + 1):
        print(f"\n[KAPANIŞ] ===== Tur {tur}/{MAX_TUR}: SGK raporu kontrol ediliyor =====", flush=True)
        r = subprocess.run(base + ["kontrol"] + ortak, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", stdin=subprocess.DEVNULL)
        sys.stdout.write(r.stdout)
        if r.returncode != 0 and r.stderr:
            sys.stdout.write(r.stderr)
        m = _re.search(r"<<<EKSIK>>>(.*)", r.stdout)
        if not m:
            print("[KAPANIŞ] ⚠️ Rapor üretilemedi / <<<EKSIK>>> okunamadı → döngü durdu.", flush=True)
            break
        eksik = [nm for nm in _json.loads(m.group(1)).get("kisiler", []) if D._fold_tr(nm) not in istisna_fold]
        for nm in eksik:
            ad_map.setdefault(D._fold_tr(nm), fold_map.get(D._fold_tr(nm), nm))
        son_eksik = eksik
        if not eksik:
            print(f"\n[KAPANIŞ] ✅ TEMİZ — çalışan herkes Gün=30. Kapanış tamam (tur {tur}).", flush=True)
            break
        denenecek = [nm for nm in eksik if D._fold_tr(nm) not in vazgec]   # kalıcı-hatalıları tekrar deneme
        if not denenecek:
            print(f"\n[KAPANIŞ] Kalan {len(eksik)} kişi KALICI hata → tekrar denenmeyecek (MANUEL gerekli).", flush=True)
            break
        hedefler, eslesmeyen = [], []
        for nm in denenecek:
            exact = fold_map.get(D._fold_tr(nm)) if fold_map else nm
            (hedefler.append(exact) if exact else eslesmeyen.append(nm))
        for nm in eslesmeyen:                       # Excel'de eşleşmeyen → manuel (retry edilemez)
            vazgec[D._fold_tr(nm)] = {"ad": nm, "sebep": "Excel'de eşleşmedi (isim/lokasyon kontrol et)"}
        print(f"\n[KAPANIŞ] {len(hedefler)} çalışan-eksik → tekrar giriliyor: {', '.join(denenecek)}", flush=True)
        if eslesmeyen:
            print(f"[KAPANIŞ]    (Excel'de eşleşmeyen, MANUEL: {', '.join(eslesmeyen)})", flush=True)
        if not hedefler:
            break
        retry_file = f"dgs_kapanis_retry_{park.code}_{sheet}.txt"
        with open(retry_file, "w", encoding="utf-8") as f:
            f.write("\n".join(hedefler) + "\n")
        print(f"[KAPANIŞ] --file {retry_file} → giriş+onay koşuluyor…\n", flush=True)
        gr = subprocess.run(base + ["giris", "--file", retry_file, "--commit", "--onayla", "--include-destek"]
                            + (["--excel", xl] if xl else []) + ortak,
                            capture_output=True, text=True, encoding="utf-8", errors="replace",
                            stdin=subprocess.DEVNULL)
        sys.stdout.write(gr.stdout)
        for mm in _BAS_RE.finditer(gr.stdout):      # kalıcı-hata verenleri işaretle → bir daha deneme
            ad, reason = mm.group(1).strip(), mm.group(2).strip()
            if _kalici(reason):
                sebep = reason.split("(ss")[0].strip()[:90]
                vazgec[D._fold_tr(ad)] = {"ad": ad, "sebep": sebep}
                print(f"[KAPANIŞ]    '{ad}' KALICI hata → tekrar denenmeyecek: {sebep}", flush=True)

    # MANUEL = son kontroldeki HÂLÂ EKSİK kişiler (kaydolmuş skip-kişiler eksik değildir → dahil olmaz).
    # Sebep: kalıcı-hata (vazgec) sınıflandırması varsa onunla, yoksa geçici-tekrarlayan.
    kisiler = []
    for nm in son_eksik:
        f = D._fold_tr(nm)
        kisiler.append(vazgec.get(f, {"ad": ad_map.get(f, nm),
                                       "sebep": "3 turda düzelmedi (portal geçici hatası tekrarladı)"}))
    print(f"<<<MANUEL>>>{_json.dumps({'lokasyon': park.code, 'kisiler': kisiler}, ensure_ascii=False)}", flush=True)
    print()
    if kisiler:
        print(f"[KAPANIŞ] ⚠️ MANUEL GİRİŞ GEREKLİ ({len(kisiler)} kişi) — otomasyon tamamlayamadı:", flush=True)
        for k in kisiler:
            print(f"     • {k['ad']} — {k['sebep']}", flush=True)
        print(f"[KAPANIŞ]    Bu kişileri portalda ELLE gir. (Fix listesi: dgs --park {park.code} kontrol --write-fix)", flush=True)
    else:
        print("[KAPANIŞ] ✅ Kapanış tamam — manuel giriş gereken kimse yok.", flush=True)


# ==========================================================================
# TEK GİRİŞ NOKTASI (donmuş exe + GUI)
# ==========================================================================
# Donmuş exe'nin İÇİNDE .py dosyası YOK → `python3 dgs_tpi.py` gibi çağrı yapılamaz. Bu yüzden exe
# kendini `dgs --park <KOD> ...` token'ıyla çağırır (izin_app.py dispatcher'ı buraya yönlendirir).
# Park modülü import EDİLİR — doğrudan run(PARKS[kod]) DEĞİL — ki o parkın `overrides()`'ı KORUNSUN
# (park-başına-script tasarımı bozulmaz; bir parka özel davranış eklenirse exe'de de geçerli olur).
PARK_MODULLERI = {
    "TPI": "dgs_tpi",
    "BV": "dgs_bv",
    "TPIz": "dgs_tpiz",
    "Yıldız": "dgs_yildiz",
    "Ulutek": "dgs_ulutek",
}

_KULLANIM = (
    "Kullanım: dgs --park <{parklar}> [giris|onay|kontrol|liste] [bayraklar]\n"
    "  dgs --park TPI giris   --excel '<xlsx>' --onayla --commit   # BULK giriş+onay\n"
    "  dgs --park TPI onay    --excel '<xlsx>' --commit            # kalan taslakları onaya gönder\n"
    "  dgs --park TPI kontrol                                      # SGK Gün=30 (--excel ALMAZ)\n"
    "  dgs --park TPI liste   --excel '<xlsx>'                     # Excel'deki kişileri JSON dök (portalsız)\n"
    "Her modda: --exclude-file <dosya>  → İSTİSNA KİŞİLER (otomasyon bunlara DOKUNMAZ)\n"
    "DİKKAT: DGS park kodları İZİN kodlarından FARKLI → İYTE=TPIz · YTP=Yıldız · ULUTEK=Ulutek"
)


def _hata(msg: str):
    print(f"HATA: {msg}\n{_KULLANIM.format(parklar='|'.join(PARK_MODULLERI))}", file=sys.stderr)
    sys.exit(2)


def main():
    """`dgs --park <KOD> [mod] ...` → o parkın modülünü yükle, ortak motoru onun ayarlarıyla koş."""
    import importlib

    argv = ["--help" if a == "-h" else a for a in sys.argv[1:]]   # -h süzgeçte yenmesin
    if "--park" not in argv:
        _hata("--park <KOD> şart.")
    i = argv.index("--park")
    if i + 1 >= len(argv):
        _hata("--park bayrağının değeri verilmemiş.")
    kod = argv[i + 1]
    if kod not in PARK_MODULLERI:
        _hata(f"bilinmeyen park {kod!r}.")
    del argv[i:i + 2]                          # --park'ı düş (motorlar bu bayrağı tanımıyor)
    sys.argv = [sys.argv[0]] + argv

    park_mod = importlib.import_module(PARK_MODULLERI[kod])
    run(park_mod.PARK, park_mod.overrides)


if __name__ == "__main__":
    main()
