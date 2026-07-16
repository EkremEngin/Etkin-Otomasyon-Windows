#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SGK RAPOR KONTROL — DGS girişlerini "Gün Detaylı Rapor" ile doğrula (park-agnostik).
====================================================================================
KULLANICI AKIŞI (manuel, video): PERSONEL > PDKS > "Sgk Çalışan Bildirgesi Gün Detaylı
Raporu" → Dönem = bir önceki ay (ör. Haziran'dayız → MAYIS 2026) → "Rapor Hazırla" →
açılan rapor görüntüleyicide kaydet(disket)▼ butonunun üstüne gel → "Adobe Acrobat" (PDF) →
inen raporu açıp DGS girişlerinde eksik/girilmemiş/hatalı giren kişi var mı kontrol et.

Bu script aynı raporu üretir ama PDF yerine raporun CSV export'unu çeker (AYNI VERİ, makine-
okunur — PDF'i parse etmekten çok daha güvenilir) ve otomatik analiz eder:
  • "Gelir Vergi İstisnası Gün" < 30 olan kişileri bulur (= o ay tam istisnaya ulaşmamış).
  • dgs_done_<lokasyon>_<sheet>.txt setiyle ÇAPRAZ-REFERANS:
      - BİZİM (done'daki) kişi Gün<30          → EKSİK/HATALI giriş → DÜZELTİLMELİ.
      - bizim done'da olup raporda HİÇ görünmeyen → GİRİLMEMİŞ (kayıt yok) → DÜZELTİLMELİ.
      - bizim-olmayan Gün<30                    → kapsam dışı (Destek/ayrılan), yalnız BİLGİ.
  • --write-fix ile bizim-eksikleri fix_<lokasyon>.txt'e yazar (dgs_sil + dgs_poc --file için).

KRİTER (Plan §6): Gün<30. ('Toplam<Gereken' DEĞİL — izinli kişide Toplam<Gereken NORMALDİR,
ama Gün yine 30 olur çünkü izin günü de istisnaya sayılır. O yüzden ölçüt yalnız Gün.)

GÜVENLİK: salt-OKUR. Hiçbir kayıt değiştirmez/silmez/göndermez (yalnız rapor üretir + CSV indirir).
İndirme = portal kendi export'u (window.open URL'i fetch). Sayfa RELOAD YOK (Cloudflare).

Kullanım:
  DGS_CDP=http://localhost:9223 python3 dgs_rapor_kontrol.py --lokasyon TPIz
  DGS_CDP=http://localhost:9223 python3 dgs_rapor_kontrol.py --lokasyon TPIz --write-fix
  (Dönem otomatik bir önceki ay; override: --donem "MAYIS 2026")
"""
from __future__ import annotations
import argparse
import csv
import datetime
import io
import json
import os
import re
import sys
import tempfile

import dgs_poc as D
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

LOG = "[RAPOR]"
TR_AYLAR = {1: "OCAK", 2: "ŞUBAT", 3: "MART", 4: "NİSAN", 5: "MAYIS", 6: "HAZİRAN",
            7: "TEMMUZ", 8: "AĞUSTOS", 9: "EYLÜL", 10: "EKİM", 11: "KASIM", 12: "ARALIK"}
REPORT_ITEM = "Sgk Çalışan Bildirgesi Gün Detaylı Raporu"


def log(*a):
    print(LOG, *a, flush=True)


def prev_month_label(today: datetime.date | None = None) -> str:
    """Bugünden bir önceki takvim ayı → 'MAYIS 2026'."""
    today = today or datetime.date.today()
    y, m = today.year, today.month
    m -= 1
    if m == 0:
        m, y = 12, y - 1
    return f"{TR_AYLAR[m]} {y}"


def fold(s: str) -> str:
    """Türkçe-duyarsız normalize (isim eşleştirme; İ↔I, Ç↔C ...)."""
    s = (s or "").replace("̇", "")
    tr = {'ç': 'c', 'Ç': 'c', 'ğ': 'g', 'Ğ': 'g', 'ı': 'i', 'İ': 'i', 'I': 'i',
          'ö': 'o', 'Ö': 'o', 'ş': 's', 'Ş': 's', 'ü': 'u', 'Ü': 'u'}
    return " ".join("".join(tr.get(c, c) for c in s).lower().split())


def open_sgk_report(page, donem_label: str):
    """PERSONEL → PDKS (hover) → SGK rapor → Dönem(etiket) set → Rapor Hazırla → rapor gelene dek bekle."""
    # temiz slate: açık dialogları kapat (önceki form/liste menüyü bloklayabilir)
    for _ in range(6):
        n = page.evaluate(r"""()=>{const b=[...document.querySelectorAll('div.ui-dialog-titlebar')].filter(x=>x.offsetParent!==null);
              b.forEach(x=>{const c=x.querySelector('.fa-times,.ui-dialog-titlebar-close');if(c)c.click();});return b.length;}""")
        page.wait_for_timeout(300)
        if n == 0:
            break
    # menü: PERSONEL → PDKS hover → rapor öğesi (jQuery delegated → gerçek click/hover)
    last = None
    for _ in range(3):
        try:
            page.get_by_text("PERSONEL", exact=True).first.click(timeout=5000)
            page.wait_for_timeout(600)
            page.get_by_text("PDKS", exact=True).filter(visible=True).first.hover(timeout=4000)
            page.wait_for_timeout(900)
            page.get_by_text(REPORT_ITEM, exact=True).filter(visible=True).first.click(timeout=5000)
            break
        except Exception as e:
            last = e
            page.wait_for_timeout(500)
    else:
        raise D.VerifyError(f"SGK rapor menüsü açılamadı: {last}")
    # Dönem select gelene dek bekle
    page.wait_for_selector("#Donem_Id, select[name='Donem_Id']", timeout=20_000)
    D.wait_for_form_idle(page)
    # Dönem'i ETİKETE göre seç (park-bağımsız; Donem_Id park-park değişir)
    val = page.evaluate(
        r"""(lbl)=>{const s=document.querySelector('#Donem_Id, select[name="Donem_Id"]'); if(!s)return '__YOK__';
            const o=[...s.options].find(o=>o.text.replace(/\s+/g,' ').trim().toUpperCase()===lbl.toUpperCase());
            return o?o.value:'__BULUNAMADI__';}""", donem_label)
    if val in ("__YOK__", "__BULUNAMADI__"):
        opts = page.evaluate(
            r"""()=>{const s=document.querySelector('#Donem_Id, select[name="Donem_Id"]');
                return s?[...s.options].map(o=>o.text.trim()).filter(Boolean):[];}""")
        raise D.VerifyError(f"Dönem etiketi {donem_label!r} bulunamadı. Mevcut: {opts}")
    page.select_option("#Donem_Id", value=val)
    page.wait_for_timeout(500)
    # Rapor Hazırla (F3) — gerçek click (fallback evaluate)
    try:
        page.locator("a:has-text('Rapor Hazırla'), button:has-text('Rapor Hazırla')").filter(visible=True).first.click(timeout=5000)
    except Exception:
        page.evaluate(r"""()=>{const b=[...document.querySelectorAll('a,button,input')].find(x=>x.offsetParent!==null
              && /Rapor Hazırla/i.test(x.textContent||x.value||'')); if(b)b.click();}""")
    # rapor üretimi: "Gelir Vergi" + "İstisna" body'de görünene dek bekle (loading)
    for _ in range(25):
        page.wait_for_timeout(1000)
        if page.evaluate(r"""()=>{const t=document.body.innerText; return /Gelir Vergi/i.test(t) && /İstisna/i.test(t);}"""):
            break
    page.wait_for_timeout(1500)


def fetch_report_csv(page) -> str:
    """Rapor görüntüleyicinin CSV export URL'ini (export_csv input onclick window.open) çek + fetch et.
    PDF (Adobe Acrobat) yerine CSV: aynı 'Gelir Vergi İstisnası Gün' verisi, makine-okunur."""
    url = page.evaluate(
        r"""()=>{const el=document.querySelector("input[name='export_csv'], input[id*='export_csv']");
            if(!el)return null; const oc=el.getAttribute('onclick')||'';
            const m=oc.match(/window\.open\('([^']+)'/); return m?m[1].replace(/&amp;/g,'&'):null;}""")
    if not url:
        # alternatif: herhangi bir export linkinde CSV uzantısı/format=csv
        url = page.evaluate(
            r"""()=>{const a=[...document.querySelectorAll('a,input')].map(e=>e.getAttribute('onclick')||e.href||'')
                .find(s=>/Format=CSV|\.csv/i.test(s)); if(!a)return null;
                const m=a.match(/'([^']*(?:Format=CSV|\.csv)[^']*)'/i)||a.match(/(https?:\/\/[^'"\s]+)/);
                return m?m[1].replace(/&amp;/g,'&'):null;}""")
    if not url:
        raise D.VerifyError("CSV export URL'i bulunamadı (rapor görüntüleyici farklı olabilir — ekran görüntüsü al).")
    txt = page.evaluate(
        r"""async (u)=>{const r=await fetch(u,{credentials:'include'}); const buf=await r.arrayBuffer();
            return new TextDecoder('windows-1254').decode(buf);}""", url)
    return txt


_TIME = re.compile(r'^\d{1,2}:\d{2}$')
_NUM = re.compile(r'^\d{1,3},\d{2}$')   # XX,XX (gün sayısı)
_DATE = re.compile(r'^\d{2}-\d{2}-\d{4}$')   # GG-AA-YYYY (işe başlangıç / ayrılış tarihi)


def _num(s):
    try:
        return float(s.replace(',', '.'))
    except Exception:
        return None


def parse_report(txt: str) -> list[tuple[str, float | None, str, str]]:
    """KOLON-KAYMASI TOLERANSLI parse: her TC ('*'li) satırında SON İKİ 'XX,XX' =
    (Gelir Vergi İstisnası Gün, Sigorta Primi Gün). Günlük grid HH:MM olduğu için XX,XX
    yalnız özet blokta bulunur → satırın tümünden toplamak güvenli.
    Dönüş: [(ad, gv_gun, sig_str, ayrilis_tarihi), ...].

    🔴 2026-07-15 FIX (yedek: prime/dgs_rapor_kontrol_2026-07-15.py): eski sürüm 'ilk HH:MM
    kolonundan ÖNCEKİ' XX,XX'e bakıyordu. Tam-ay çalışanın Toplam Süre'si '198:44' (3 haneli →
    _TIME saymaz) sorun değildi; ama 100 SAATTEN AZ çalışanın Toplam'ı '54:10'/'45:00' (2 haneli
    → _TIME SAYAR) ilk-HH:MM'i öne çekiyor, Gün kolonu yakalanamıyor, satır gv=None ile SESSİZCE
    DÜŞÜYORDU. Tam da en riskli kişiler (az-saatli/eksik girilen — ör. REMZİ TİRE Gün=6,02) kaçıyordu."""
    rows = list(csv.reader(io.StringIO(txt), delimiter=';'))
    out = []
    for r in rows:
        cells = [c.strip() for c in r]
        if not cells or '*' not in cells[0]:        # TC (maskeli) satırı değil → atla
            continue
        ad = next((c for c in cells[1:12] if c and not _NUM.match(c) and not re.match(r'^\d', c)), '?')
        xx = [c for c in cells if _NUM.match(c)]     # TÜM XX,XX (özet blok) → son iki = (GV Gün, Sig Gün)
        gv = xx[-2] if len(xx) >= 2 else None
        sig = xx[-1] if len(xx) >= 1 else '?'
        dates = [c for c in cells if _DATE.match(c)]  # 1. = işe başlangıç, 2. (varsa) = AYRILIŞ
        ayrilis = dates[1] if len(dates) >= 2 else ''  # boş → hâlâ çalışıyor (eksikse DÜZELTİLMELİ)
        out.append((ad, _num(gv) if gv else None, sig, ayrilis))
    return out


def main():
    ap = argparse.ArgumentParser(description="SGK Gün Detaylı Rapor ile DGS doğrula (salt-okur)")
    ap.add_argument("--lokasyon", default="TPIz")
    ap.add_argument("--sheet", default="Mayıs")
    ap.add_argument("--donem", default=None, help="Dönem etiketi (ör. 'MAYIS 2026'); yoksa oto önceki ay")
    ap.add_argument("--write-fix", action="store_true", help="Bizim-eksikleri fix_<lokasyon>.txt'e yaz")
    ap.add_argument("--threshold", type=float, default=29.995, help="Gün eşiği (altı = eksik)")
    args = ap.parse_args()
    label = args.donem or prev_month_label()

    done_file = f"dgs_done_{args.lokasyon}_{args.sheet}.txt"
    done_names = [l.strip() for l in open(done_file, encoding="utf-8")] if os.path.exists(done_file) else []
    done_names = [n for n in done_names if n]
    fold2orig = {fold(n): n for n in done_names}
    our = set(fold2orig)
    log(f"Dönem: {label} | lokasyon: {args.lokasyon} | done(bizim girdiğimiz): {len(our)}")

    with sync_playwright() as pw:
        _b, page = D.attach_browser(pw)
        D.assert_logged_in(page)
        log("Rapor üretiliyor (PERSONEL > PDKS > SGK Gün Detaylı Raporu > Rapor Hazırla)...")
        open_sgk_report(page, label)
        try:
            txt = fetch_report_csv(page)
        except D.VerifyError as e:
            ts = int(__import__("time").time())
            page.screenshot(path=f"sgk_rapor_{ts}.png", full_page=True)
            log(f"HATA: {e} (ss: sgk_rapor_{ts}.png)")
            sys.exit(2)
        # gettempdir(): Windows'ta "/tmp" C:\tmp'ye çözülür ve o klasör yoktur → FileNotFoundError.
        csv_path = os.path.join(tempfile.gettempdir(), f"sgk_{args.lokasyon}_{args.sheet}.csv")
        open(csv_path, "w", encoding="utf-8").write(txt)
        log(f"CSV indirildi: {csv_path}")

    people = parse_report(txt)
    if not people:
        log("Rapor boş/parse edilemedi — CSV'yi elle incele.")
        sys.exit(2)
    in_report = {fold(ad) for ad, _, _, _ in people}
    short = [(ad, gv, sig, ay) for ad, gv, sig, ay in people if gv is not None and gv < args.threshold]
    # Gün<30 ama AY İÇİNDE AYRILMIŞ (ayrılış tarihi dolu) → o kişi için Gün<30 NORMAL, DÜZELTİLMEZ.
    ayrilan_short = [p for p in short if p[3]]
    calisan_short = [p for p in short if not p[3]]        # çalışıyor + eksik → gerçek düzeltilecek

    # ÇAPRAZ-REFERANS (yalnız ÇALIŞAN-eksikler üzerinden — ayrılanı asla düzeltme)
    our_short = [(ad, gv, sig, ay) for ad, gv, sig, ay in calisan_short if fold(ad) in our]
    other_short = [(ad, gv, sig, ay) for ad, gv, sig, ay in calisan_short if fold(ad) not in our]
    our_missing = sorted(fold2orig[f] for f in our if f not in in_report)

    print()
    log(f"==== RAPOR: {len(people)} kişi | Gün=30: {len(people)-len(short)} | "
        f"Gün<30: {len(short)} (çalışan-eksik {len(calisan_short)} + ayrılan {len(ayrilan_short)}) ====")
    print()
    log(f">>> BİZİM ÇALIŞAN-EKSİK ({len(our_short)}) — Gün<30, DÜZELTİLMELİ:")
    for ad, gv, sig, ay in sorted(our_short, key=lambda x: x[1] if x[1] is not None else 0):
        log(f"     {ad[:34]:34} | GV Gün={gv:>6}")
    if not our_short:
        log("     (yok — bizim girdiğimiz çalışan herkes Gün=30 ✓)")
    print()
    log(f">>> BİZİM GİRİLMEMİŞ ({len(our_missing)}) — done'da var ama raporda YOK:")
    for n in our_missing:
        log(f"     {n}")
    if not our_missing:
        log("     (yok — done'daki herkes raporda görünüyor ✓)")
    print()
    if other_short:
        log(f">>> ÇALIŞAN-EKSİK, bizim-OLMAYAN ({len(other_short)}) — Destek/başka giren, yalnız bilgi:")
        for ad, gv, sig, ay in other_short[:40]:
            log(f"     {ad[:34]:34} | GV Gün={gv:>6}")
        print()
    if ayrilan_short:
        log(f">>> AYRILANLAR ({len(ayrilan_short)}) — Gün<30 ama ay içinde AYRILMIŞ → NORMAL, DOKUNMA:")
        for ad, gv, sig, ay in ayrilan_short:
            log(f"     {ad[:34]:34} | GV Gün={gv:>6} | ayrılış={ay}")
        print()

    # DÜZELTİLECEK = bu lokasyonda ÇALIŞAN + eksik HERKES (bizim done'da olsun/olmasın — REMZİ gibi
    # 'girmeye çalıştık ama düştü' de dahil, çünkü hepsi bizim girmemiz gereken kişiler) + done'da olup
    # raporda görünmeyenler. Bizim-olanlar için done'daki tam ad; diğerleri için rapor-adı
    # (kapanış-orkestrasyonu Excel'e karşı fold ile eşleyip tam adı bulur).
    fix_targets = sorted(set(
        [fold2orig.get(fold(ad), ad) for ad, _, _, _ in calisan_short] + our_missing))
    # MAKİNE-OKUNUR — kapanış-orkestrasyonu (dgs_park kapanis) düzeltilecekleri BU satırdan okur.
    print(f"<<<EKSIK>>>{json.dumps({'lokasyon': args.lokasyon, 'kisiler': fix_targets}, ensure_ascii=False)}")
    print()
    if fix_targets:
        log(f"==== DÜZELTİLECEK (çalışan-eksik): {len(fix_targets)} kişi ====")
        for n in fix_targets:
            log(f"     {n}")
        if args.write_fix:
            fix_file = f"fix_{args.lokasyon}.txt"
            with open(fix_file, "w", encoding="utf-8") as f:
                f.write("\n".join(fix_targets) + "\n")
            log(f"\n>> {fix_file} yazıldı. DÜZELTME AKIŞI:")
            log(f"   1) python3 dgs_sil.py        --file {fix_file} --commit   # onaylı kaydı sil (dönem oto)")
            log(f"   2) (resume'dan çıkar) dgs_done/dgs_onaya_done'dan bu isimleri sil")
            log(f"   3) python3 dgs_poc.py        --excel <XL> --lokasyon {args.lokasyon} --file {fix_file} --no-schedule --commit")
            log(f"   4) python3 dgs_onaya.py      --excel <XL> --lokasyon {args.lokasyon} --commit")
            log(f"   5) python3 dgs_rapor_kontrol.py --lokasyon {args.lokasyon}   # tekrar kontrol → hepsi Gün=30 olana dek")
        else:
            log("   (fix dosyası yazmak için --write-fix ekle)")
    else:
        log("==== ✅ TEMİZ: bizim girdiğimiz ÇALIŞAN herkes Gün=30. Düzeltme gerekmez. ====")


if __name__ == "__main__":
    main()
