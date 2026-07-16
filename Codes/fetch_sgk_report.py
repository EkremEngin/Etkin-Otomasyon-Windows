#!/usr/bin/env python3
"""SGK Çalışan Bildirgesi Gün Detaylı Raporu'nu üret + oku → herkeste Gelir Vergi İstisnası Gün=30 mu?
PERSONEL > PDKS > rapor → Dönem=MAYIS 2026 → Rapor Hazırla → tabloyu parse et, Gün<30 olanları bas.
Kullanım: DGS_CDP=http://localhost:9222 python3 fetch_sgk_report.py
"""
import os, sys, json
import dgs_poc as D
from playwright.sync_api import sync_playwright

DONEM_TEXT = "MAYIS 2026"

with sync_playwright() as pw:
    _b, page = D.attach_browser(pw)
    D.assert_logged_in(page)
    page.evaluate(r"""()=>{[...document.querySelectorAll('div.ui-dialog-titlebar')].filter(x=>x.offsetParent!==null).forEach(x=>{const c=x.querySelector('.fa-times,.ui-dialog-titlebar-close');if(c)c.click();});}""")
    page.wait_for_timeout(400)
    # PERSONEL → PDKS (hover) → rapor
    page.get_by_text("PERSONEL", exact=True).first.click(timeout=5000); page.wait_for_timeout(600)
    page.get_by_text("PDKS", exact=True).filter(visible=True).first.hover(timeout=4000); page.wait_for_timeout(900)
    page.get_by_text("Sgk Çalışan Bildirgesi Gün Detaylı Raporu", exact=True).filter(visible=True).first.click(timeout=5000)
    page.wait_for_timeout(2500)
    # Dönem = MAYIS 2026 (etikete göre value bul)
    val = page.evaluate(r"""(txt)=>{const s=document.querySelector('#Donem_Id'); if(!s)return null;
        const o=[...s.options].find(o=>o.text.replace(/\s+/g,' ').trim().toUpperCase().includes(txt)); return o?o.value:null;}""", DONEM_TEXT)
    if not val:
        print("Dönem MAYIS 2026 bulunamadı"); sys.exit(1)
    page.select_option("#Donem_Id", value=val)
    print(f"Dönem set: MAYIS 2026 (value={val})")
    page.wait_for_timeout(600)
    # Rapor Hazırla (F3) — gerçek click
    try:
        page.locator("a:has-text('Rapor Hazırla'), button:has-text('Rapor Hazırla')").filter(visible=True).first.click(timeout=5000)
    except Exception:
        page.evaluate(r"""()=>{const b=[...document.querySelectorAll('a,button')].find(x=>x.offsetParent!==null && /Rapor Hazırla/i.test(x.textContent)); if(b)b.click();}""")
    # rapor üretimi — loading bitene kadar bekle
    for _ in range(20):
        page.wait_for_timeout(1000)
        done = page.evaluate(r"""()=>{const t=document.body.innerText; return /Gelir Vergi/i.test(t) && /İstisna/i.test(t);}""")
        if done: break
    page.wait_for_timeout(1500)
    # tabloyu bul + dök
    data = page.evaluate(r"""()=>{
        const tbls=[...document.querySelectorAll('table')].filter(t=>/Gelir Vergi|Personel/i.test(t.innerText));
        let best=null, mx=0;
        for(const t of tbls){const n=t.querySelectorAll('tr').length; if(n>mx){mx=n;best=t;}}
        if(!best) return null;
        return [...best.querySelectorAll('tr')].map(tr=>[...tr.children].map(td=>(td.textContent||'').replace(/\s+/g,' ').trim()));
    }""")
    if not data:
        print("Rapor tablosu bulunamadı"); page.screenshot(path="sgk_norows.png"); sys.exit(2)
    print(f"Tablo: {len(data)} satır")
    # başlık satırı: 'Personel' ve 'Gelir Vergi' içeren
    hdr_i = next((i for i,r in enumerate(data) if any("Personel" in c for c in r) and any("Gelir Vergi" in c for c in r)), 0)
    hdr = data[hdr_i]
    def col(key):
        return next((i for i,c in enumerate(hdr) if key.lower() in c.lower()), None)
    ci_pers = col("Personel"); ci_gv = col("Gelir Vergi"); ci_sig = col("Sigorta Primi"); ci_ist = col("İstisnalı Çalış"); ci_tam = col("Tam İstisna")
    print(f"kolonlar: Personel={ci_pers} GelirVergi={ci_gv} Sigorta={ci_sig} İstisnalı={ci_ist} TamGereken={ci_tam}")
    print("\n==== TÜM KİŞİLER (Gelir Vergi İstisnası Gün) ====")
    problems=[]
    for r in data[hdr_i+1:]:
        if ci_pers is None or ci_pers >= len(r): continue
        ad=r[ci_pers]
        if not ad or "Personel" in ad: continue
        gv = r[ci_gv] if ci_gv is not None and ci_gv < len(r) else "?"
        sig = r[ci_sig] if ci_sig is not None and ci_sig < len(r) else "?"
        ist = r[ci_ist] if ci_ist is not None and ci_ist < len(r) else "?"
        try: short = float(gv.replace(",", ".")) < 29.995
        except: short = False
        mark = "  <<< EKSİK" if short else ""
        print(f"  {ad[:30]:30} | GV={gv:>7} | Sig={sig:>7} | İst={ist:>8}{mark}")
        if short: problems.append((ad, gv, sig, ist))
    print(f"\n==== EKSİK (Gün<30): {len(problems)} kişi ====")
    for ad,gv,sig,ist in problems: print(f"  {ad} | GV={gv} Sig={sig} İst={ist}")
