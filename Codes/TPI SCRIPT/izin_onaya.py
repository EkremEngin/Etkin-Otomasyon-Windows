#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
İzin TASLAK'larını ONAYA GÖNDER — izin_poc.py'nin girdiği taslakları onaya gönderir.
=====================================================================================
Akış (canlı doğrulandı 2026-06-17): liste → taslağa ÇİFT TIKLA → "Onaya Gönder" → "Emin misiniz? Evet"
→ kayıt "Yönetici Şirket Tarafından Onaylanmış" olur. E-İMZA/ŞİFRE ÇIKMIYOR (operatör adımı; bu portalda
direkt onaylanıyor).

YÖNTEM (sağlam): Liste DURUM filtresini SADECE "Değerlendirmeye Gönderilmemiş"e ayarla → listede yalnız
taslaklar kalır. Onaya gönderdikçe o kayıt listeden düşer → hep "ilk taslağı onayla, listeyi tazele" yeter
(isim-arama / sayfalama gerekmez). GÜVENLİK: sadece `izin_done_<lok>_<sheet>.txt`'deki isimleri onaylar;
listede bizim-olmayan taslak çıkarsa DOKUNMAZ (atlar, uyarır). Onay sonrası kaydın listeden düştüğünü
DOĞRULAR. Resume: `izin_onaya_done_<lok>_<sheet>.txt`. GERİ ALINMASI ZOR (resmi onay) — kullanıcı istedi.

Kullanım:
  python3 izin_onaya.py --limit 1     # tek taslak (test)
  python3 izin_onaya.py               # izin_done'daki tüm taslaklar
"""
from __future__ import annotations
import argparse
import os
import sys
import time

from playwright.sync_api import sync_playwright

from izin_data import fold

CDP = os.environ.get("DGS_CDP", "http://localhost:9222")  # paralel park için: DGS_CDP ile port ver (9222 BV / 9223 TPIz)
DONEM_LABEL = "MAYIS 2026"   # dönem ETİKETİ — park-bazlı Donem_Id değişir → etiketle seç (Plan §5.2)
LOG = "[ONAY]"

# JS tarafı Türkçe-fold (Python fold ile aynı kural)
_JS_FOLD = r"""const fold=s=>{const tr={'ç':'c','Ç':'c','ğ':'g','Ğ':'g','ı':'i','İ':'i','I':'i','ö':'o','Ö':'o','ş':'s','Ş':'s','ü':'u','Ü':'u'};
  return s.split('').map(c=>tr[c]||c).join('').toLowerCase().replace(/\s+/g,' ').trim();};"""


def log(*a):
    print(LOG, *a, flush=True)


def set_donem_by_label(page, label: str):
    """Liste filtresinde Dönem'i ETİKETE göre seç (Playwright select_option) + value DOĞRULA. Park-bağımsız:
    Donem_Id değeri parka göre değişir ama etiket ('MAYIS 2026') sabit. Ham `value=` bazı combobox'ta tutmaz."""
    val = page.evaluate(
        r"""(lbl)=>{const s=[...document.querySelectorAll('#Donem_Id, select[name="Donem_Id"]')].filter(e=>e.offsetParent!==null).pop();
            if(!s)return '__YOK__'; const o=[...s.options].find(o=>(o.text||'').trim()===lbl); return o?o.value:'__BULUNAMADI__';}""",
        label)
    if val in ("__YOK__", "__BULUNAMADI__"):
        opts = page.evaluate(r"""()=>{const s=[...document.querySelectorAll('#Donem_Id, select[name="Donem_Id"]')].filter(e=>e.offsetParent!==null).pop();
            return s?[...s.options].map(o=>(o.text||'').trim()).filter(Boolean):[];}""")
        raise RuntimeError(f"Dönem etiketi {label!r} bulunamadı (val={val}). Mevcut: {opts}")
    page.locator('select#Donem_Id:visible, select[name="Donem_Id"]:visible').last.select_option(value=val, timeout=6000)
    page.wait_for_timeout(400)
    got = page.evaluate(r"""()=>{const s=[...document.querySelectorAll('#Donem_Id, select[name="Donem_Id"]')].filter(e=>e.offsetParent!==null).pop();
        return s?s.value:'';}""")
    if got != val:
        raise RuntimeError(f"Dönem set edilemedi (value={got!r} != {val!r}, label={label!r}).")


def wait_ajax(page, timeout=20000, settle=80):
    """Yükleme bitene kadar bekle — GÖSTERGE: görünür .form_loading_panel/.pReload.loading kaybolması.
    🔴 KRİTİK (2026-06-18 Ulutek): jQuery.active bu portalda KALICI bağlantılar yüzünden HİÇ 0 olmuyor (sabit 5)
    → eski 'jQuery.active==0' koşulu her çağrıda 20s timeout yiyordu. KALDIRILDI (panel + poll yeterli)."""
    try:
        page.wait_for_function(
            r"""()=>![...document.querySelectorAll('.form_loading_panel,.pReload.loading')].some(e=>e.offsetParent!==null)""",
            timeout=timeout)
    except Exception:
        pass
    page.wait_for_timeout(settle)


def main():
    ap = argparse.ArgumentParser(description="İzin taslaklarını onaya gönder (sadece izin_done'dakiler)")
    ap.add_argument("--lokasyon", default="TPI")
    ap.add_argument("--sheet", default="Mayıs")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    done_file = f"izin_done_{args.lokasyon}_{args.sheet}.txt"
    if not os.path.exists(done_file):
        log(f"HATA: {done_file} yok."); sys.exit(1)
    our = set(fold(l.strip()) for l in open(done_file, encoding="utf-8") if l.strip())
    onay_file = f"izin_onaya_done_{args.lokasyon}_{args.sheet}.txt"
    onaylanan = set(fold(l) for l in open(onay_file, encoding="utf-8").read().splitlines()) if os.path.exists(onay_file) else set()
    hedef = len([x for x in our if x not in onaylanan])
    limit = args.limit if args.limit else 10**9
    log(f"{len(our)} taslak (izin_done); {len(onaylanan)} zaten onaya gönderilmiş; {min(hedef, limit)} işlenecek.")

    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp(CDP)
        page = next((p for p in browser.contexts[0].pages if "argeportal" in (p.url or "")), None)
        if page is None:
            log("Portal sekmesi yok."); sys.exit(1)

        def close_all():
            page.evaluate(r"""()=>{[...document.querySelectorAll('a,button')].filter(b=>b.offsetParent!==null
                  && /^(TAMAM|OK|Kapat)$/.test(b.textContent.trim())).forEach(b=>b.click());
                  document.querySelectorAll('.sweet-alert .confirm,.sweet-alert button.confirm').forEach(b=>b.click());
                  document.querySelectorAll('.sweet-overlay,.sweet-alert,.ui-widget-overlay').forEach(e=>e.style.display='none');}""")
            page.wait_for_timeout(400)
            page.evaluate(r"""()=>{[...document.querySelectorAll('div.ui-dialog-titlebar')].filter(b=>b.offsetParent!==null)
                  .forEach(b=>{const x=b.querySelector('.fa-times,.ui-dialog-titlebar-close'); if(x)x.click();});}""")
            page.wait_for_timeout(700)

        def close_form():
            # popup'ları kapat + SADECE form penceresini kapat (LİSTE açık kalmalı, yoksa sonraki listele boşa düşer)
            page.evaluate(r"""()=>{[...document.querySelectorAll('a,button')].filter(b=>b.offsetParent!==null
                  && /^(TAMAM|OK|Kapat)$/.test(b.textContent.trim())).forEach(b=>b.click());}""")
            page.wait_for_timeout(300)
            page.evaluate(r"""()=>{[...document.querySelectorAll('div.ui-dialog-titlebar')].filter(b=>b.offsetParent!==null
                  && /Personel İzin Bildirim Formu/.test(b.textContent) && !/Listesi/.test(b.textContent))
                  .forEach(b=>{const x=b.querySelector('.fa-times,.ui-dialog-titlebar-close'); if(x)x.click();});}""")
            page.wait_for_timeout(500)

        def list_open():
            return page.evaluate(r"""()=>[...document.querySelectorAll('#Donem_Id')].some(e=>e.offsetParent!==null)""")

        def open_list():
            close_all()
            for _ in range(3):
                try:
                    page.get_by_text("PERSONEL", exact=True).first.click(timeout=5000)
                    page.wait_for_timeout(600)
                    page.get_by_text("İzin Bildirim Formu", exact=True).filter(visible=True).first.click(timeout=5000)
                    page.wait_for_timeout(1200)
                    break
                except Exception:
                    page.wait_for_timeout(500)
            page.wait_for_selector('select#Donem_Id', timeout=20000)
            try:
                page.locator(".ui-dialog-titlebar").filter(has_text="Listesi").first.click(timeout=2000)
            except Exception:
                pass
            page.wait_for_timeout(400)
            # Dönem = MAYIS 2026 — ETİKETE göre (park-bazlı Donem_Id değişir → label ile select_option)
            set_donem_by_label(page, DONEM_LABEL)
            # DURUM filtresi: SADECE "Değerlendirmeye Gönderilmemiş" (OnayDurumu_Id_5) açık kalsın
            page.evaluate(r"""()=>{document.querySelectorAll('input[type=checkbox][name^="OnayDurumu_Id"], input[type=checkbox][id^="OnayDurumu_Id"]').forEach(cb=>{
                  const keep = (cb.name==='OnayDurumu_Id_5'||cb.id==='OnayDurumu_Id_5');
                  if(cb.checked!==keep){ cb.click(); }});}""")
            page.wait_for_timeout(300)

        def listele():
            page.evaluate(r"""()=>{const b=[...document.querySelectorAll('a,button')].find(x=>x.offsetParent!==null
                  && /Listele/.test(x.textContent)); if(b)b.click();}""")
            wait_ajax(page)   # listeleme AJAX'ı bitene kadar bekle (loading-fix)

        # İlk "Gönderilmemiş" taslağı aç (bizim setten, failed olmayan). Adı döndür; yoksa null.
        FIND_FIRST = r"""(args)=>{const [our, failed]=args; %s
          for(const tr of document.querySelectorAll('tr')){
            const t=tr.innerText.replace(/\s+/g,' ').trim();
            if(!/MAYIS 2026/.test(t)||!/Değerlendirmeye Gönderilmemiş/.test(t))continue;
            const m=t.match(/MAYIS 2026\s+(.+?)\s+\d+\*+\d+/); if(!m)continue;
            const nf=fold(m[1]);
            if(failed.includes(nf))continue;
            if(!our.includes(nf)) continue;                 // GÜVENLİK: sadece bizim taslaklar
            tr.dispatchEvent(new MouseEvent('dblclick',{bubbles:true,view:window})); return m[1].trim();}
          return null;}""" % _JS_FOLD

        # Bir ismin hâlâ "Gönderilmemiş" satırı var mı? (onay doğrulaması)
        STILL_DRAFT = r"""(args)=>{const [nm]=args; %s const nf=fold(nm);
          for(const tr of document.querySelectorAll('tr')){const t=tr.innerText.replace(/\s+/g,' ').trim();
            if(!/MAYIS 2026/.test(t)||!/Değerlendirmeye Gönderilmemiş/.test(t))continue;
            const m=t.match(/MAYIS 2026\s+(.+?)\s+\d+\*+\d+/); if(m && fold(m[1])===nf) return true;}
          return false;}""" % _JS_FOLD

        # bizim-olmayan kalan taslak sayısı (uyarı için)
        FOREIGN_LEFT = r"""(args)=>{const [our]=args; %s let n=0;
          for(const tr of document.querySelectorAll('tr')){const t=tr.innerText.replace(/\s+/g,' ').trim();
            if(!/MAYIS 2026/.test(t)||!/Değerlendirmeye Gönderilmemiş/.test(t))continue;
            const m=t.match(/MAYIS 2026\s+(.+?)\s+\d+\*+\d+/); if(m && !our.includes(fold(m[1]))) n++;}
          return n;}""" % _JS_FOLD

        open_list()
        ok = 0
        failed = set()
        skipped_onay = [x for x in our if x in onaylanan]
        # zaten onaya gönderilmişleri başta resume'a yazılı say
        for _iter in range(len(our) + 25):
            if ok >= limit:
                break
            if not list_open():        # liste kapandıysa yeniden aç (filtre+dönem ile)
                open_list()
            listele()
            name = page.evaluate(FIND_FIRST, [list(our), list(failed)])
            if not name:
                foreign = page.evaluate(FOREIGN_LEFT, [list(our)])
                log(f"Bizim gönderilmemiş taslak kalmadı. (bizim-olmayan kalan taslak: {foreign})")
                break
            wait_ajax(page)   # form açıldı — loading bitene kadar bekle (acele etme)
            # 🔴 Onaya Gönder GERÇEK Playwright click ŞART (evaluate-click jQuery handler'ı tetiklemiyor — 2026-06-18 Yıldız izin Kaydet gibi)
            try:
                page.locator("a:has-text('Onaya Gönder'), button:has-text('Onaya Gönder')").filter(visible=True).first.click(timeout=5000)
            except Exception:
                page.evaluate(r"""()=>{const b=[...document.querySelectorAll('a,button')].find(x=>x.offsetParent!==null
                      && /Onaya Gönder/i.test(x.textContent) && !/Geri/.test(x.textContent)); if(b)b.click();}""")
            wait_ajax(page)   # Onaya Gönder → "Emin misiniz?" onay istemi yüklenirken bekle
            # 🔴 Onay istemi "Evet" — GERÇEK Playwright click ŞART. Dialog tipi PORTAL-BAĞIMLI:
            # BV/TPIz = Bootstrap modal (#modalYes / .modal .color-green), TPI = SweetAlert (.confirm).
            # evaluate-click jQuery handler'ı GÜVENİLMEZ tetikliyor → gerçek locator click. (2026-06-18 BV)
            confirmed = False
            for sel in ("#modalYes", ".modal.in button.color-green", ".modal.show button.color-green",
                        ".modal.in button:has-text('Evet')", ".sweet-alert button.confirm", ".sweet-alert a.confirm"):
                try:
                    page.locator(sel).filter(visible=True).first.click(timeout=2500)
                    confirmed = True; break
                except Exception:
                    continue
            if not confirmed:
                for nm in ("Evet", "Eminim", "Onayla"):
                    try:
                        page.get_by_role("button", name=nm, exact=True).filter(visible=True).first.click(timeout=1500)
                        confirmed = True; break
                    except Exception:
                        continue
            # 🔴 DOĞRULAMA = form'da "Kayıt Başarı ile Onaylandı" mesajı (GÜVENİLİR sinyal).
            # relist/STILL_DRAFT KULLANMA: form aç/kapa Listesi'nin dönem filtresini sıfırlıyor →
            # yanlış-negatif ("GİTMEDİ" der ama aslında onaylanmış; 2026-06-18 BV'de tüm liste böyle raporlandı).
            ok_msg = False
            for _ in range(8):
                page.wait_for_timeout(500)
                if page.evaluate(r"""()=>/Ba[şs]ar[ıi].{0,8}Onayland[ıi]/i.test(document.body.innerText)"""):
                    ok_msg = True; break
            page.evaluate(r"""()=>{document.querySelectorAll('.modal.in .close,.modal.show .close,.sweet-alert .confirm,.sweet-alert button.confirm').forEach(b=>{try{b.click()}catch(e){}}); document.querySelectorAll('.sweet-overlay,.sweet-alert,.modal-backdrop').forEach(e=>e.style.display='none');}""")  # kalan overlay/modal temizle (menü bloklamasın)
            close_form()
            if not ok_msg:
                failed.add(fold(name))
                log(f"!! {name} → onay başarı mesajı görülmedi (atlandı; dilekçe/UI farkı? elle bak)")
                continue
            ok += 1
            with open(onay_file, "a", encoding="utf-8") as f:
                f.write(name + "\n")
            log(f"[{ok}] {name} → ONAYLANDI ✓ (Kayıt Başarı ile Onaylandı)")

        log(f"\n==== ÖZET: bu koşuda {ok} kişi onaya gönderildi ====")
        if failed:
            log(f"  !! ONAYA GİTMEYEN ({len(failed)}): elle bak.")
        try:
            input(f"{LOG} >> Bitti. ENTER...")
        except EOFError:
            pass


if __name__ == "__main__":
    main()
