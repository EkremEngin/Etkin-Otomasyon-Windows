#!/usr/bin/env python3
"""Tam-dışarıda çalışan kişi: form tam yüklenene kadar dene (flaky load), İstenilen=0:00
(her iş günü 9s dışarıda + arife 4:30, otomatik tikli) doğrula, KAYDET, teyit et.
Kullanım: DGS_CDP=http://localhost:9222 python3 fill_external.py "AD SOYAD" "PROJE" [--commit]
"""
import os, sys, json
import dgs_poc as D
from playwright.sync_api import sync_playwright

ad = sys.argv[1] if len(sys.argv) > 1 else "MERVE GÜLVEREN SARIZ"
proje = sys.argv[2] if len(sys.argv) > 2 else "ESCALATE - Sıfır Emisyonlu HDV'leri"
commit = "--commit" in sys.argv
D.CONFIG["dry_run"] = not commit

def istenilen(page):
    return page.evaluate(r"""()=>{const c=document.querySelector('.BuFordaIstenilen');const i=c?c.querySelector('input'):null;return i?i.value.trim():'?';}""")

with sync_playwright() as pw:
    browser, page = D.attach_browser(pw)
    D.assert_logged_in(page)
    good = False
    for attempt in range(1, 11):
        try:
            D.open_dgs_form(page)
            D.list_set_donem_and_yeni(page)
            D.ensure_card_fresh(page)
            D.ensure_gorevlendirme_diger(page)
            D.select_personel(page, ad)           # grid gelmezse raise → retry
            D.ensure_calisma_turu(page)
            D.set_proje(page, proje)
            D.tick_clean_days(page)               # tiksiz temiz günleri tikle (güvenli/idempotent)
            # tam yüklenme + tatmin sinyali: ~21 satır VE İstenilen=0:00
            ist = "?"
            for _ in range(8):
                page.wait_for_timeout(700)
                rows = D.read_grid_rows(page)
                ist = istenilen(page)
                if len(rows) >= 18 and ist == "0:00":
                    good = True; break
            print(f"deneme {attempt}: satır={len(rows)} İstenilen={ist} -> {'TAM ✓' if good else 'kısmi, tekrar'}")
            if good:
                break
        except Exception as e:
            print(f"deneme {attempt}: {str(e)[:80]}")
        page.wait_for_timeout(600)

    if not good:
        print("!! Tam yükleme sağlanamadı — KAYDEDİLMEDİ."); sys.exit(2)

    # kimlik teyidi
    selval = page.evaluate(r"""()=>{const e=[...document.querySelectorAll('input[name="string_Personel_Id"]')].filter(x=>x.offsetParent!==null).pop(); return e?e.value:'';}""")
    if D._fold_tr(ad) not in D._fold_tr(selval):
        print(f"!! KİMLİK uyuşmadı: form={selval!r} hedef={ad!r} — KAYDEDİLMEDİ."); sys.exit(3)
    rows = D.read_grid_rows(page)
    ticked = [r for r in rows if r["checked"]]
    total = sum(D.to_min(r["toplam"]) for r in ticked if r["toplam"])
    print(f"Kayıt öncesi: kimlik={selval[:30]!r} | tikli gün={len(ticked)} | toplam={D.to_hhmm(total)} | İstenilen={istenilen(page)}")

    # ROBUST KAYDET: gerçek Playwright click (evaluate-click flaky) + mesajı uzun poll + retry
    if not commit:
        res = "DRY-RUN (kaydedilmedi)"
    else:
        saved = False
        for s in range(4):
            try:
                page.locator("a:has-text('Kaydet (F3)'), button:has-text('Kaydet (F3)')").filter(visible=True).first.click(timeout=4000)
            except Exception:
                page.evaluate(r"""()=>{const b=[...document.querySelectorAll('a,button')].find(x=>x.offsetParent!==null && /Kaydet/i.test(x.textContent) && /F3/.test(x.textContent)); if(b)b.click();}""")
            for _ in range(10):
                page.wait_for_timeout(600)
                if "başarı ile kayıt" in page.inner_text("body").lower():
                    saved = True; break
            if saved:
                break
            print(f"  Kaydet denemesi {s+1}: mesaj yok, tekrar...")
            page.wait_for_timeout(800)
        res = "TASLAK kaydedildi" if saved else "KAYDEDİLEMEDİ (başarı mesajı yok)"
    print(f"KAYDET sonucu: {res}")
    if res == "TASLAK kaydedildi" and commit:
        with open("dgs_done_BV_Mayıs.txt", "a", encoding="utf-8") as f:
            f.write(ad + "\n")
        print(f">> {ad} dgs_done_BV'ye eklendi ✓")
