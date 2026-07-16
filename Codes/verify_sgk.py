#!/usr/bin/env python3
"""SGK raporunu YENİDEN üret + CSV çek + KOLON-KAYMASI TOLERANSLI parse → Gün<30 olanları bas.
GV/Sigorta = günlük saat (HH:MM) kolonlarından ÖNCEKİ son iki 'XX,XX' değeri (her iki sayfa düzeninde çalışır).
Kullanım: DGS_CDP=http://localhost:9222 python3 verify_sgk.py
"""
import os, sys, csv, re, io
import dgs_poc as D
from playwright.sync_api import sync_playwright

with sync_playwright() as pw:
    _b, page = D.attach_browser(pw)
    D.assert_logged_in(page)
    page.evaluate(r"""()=>{[...document.querySelectorAll('div.ui-dialog-titlebar')].filter(x=>x.offsetParent!==null).forEach(x=>{const c=x.querySelector('.fa-times,.ui-dialog-titlebar-close');if(c)c.click();});}""")
    page.wait_for_timeout(400)
    page.get_by_text("PERSONEL", exact=True).first.click(timeout=5000); page.wait_for_timeout(600)
    page.get_by_text("PDKS", exact=True).filter(visible=True).first.hover(timeout=4000); page.wait_for_timeout(900)
    page.get_by_text("Sgk Çalışan Bildirgesi Gün Detaylı Raporu", exact=True).filter(visible=True).first.click(timeout=5000)
    page.wait_for_timeout(2500)
    val = page.evaluate(r"""()=>{const s=document.querySelector('#Donem_Id'); if(!s)return null;
        const o=[...s.options].find(o=>o.text.replace(/\s+/g,' ').trim().toUpperCase().includes('MAYIS 2026')); return o?o.value:null;}""")
    page.select_option("#Donem_Id", value=val)
    page.wait_for_timeout(500)
    try:
        page.locator("a:has-text('Rapor Hazırla'), button:has-text('Rapor Hazırla')").filter(visible=True).first.click(timeout=5000)
    except Exception:
        page.evaluate(r"""()=>{const b=[...document.querySelectorAll('a,button')].find(x=>x.offsetParent!==null && /Rapor Hazırla/i.test(x.textContent)); if(b)b.click();}""")
    for _ in range(20):
        page.wait_for_timeout(1000)
        if page.evaluate(r"""()=>/Gelir Vergi/i.test(document.body.innerText)"""): break
    page.wait_for_timeout(1500)
    url = page.evaluate(r"""()=>{const el=document.querySelector("input[name='export_csv']"); if(!el)return null; const oc=el.getAttribute('onclick')||''; const m=oc.match(/window\.open\('([^']+)'/); return m?m[1].replace(/&amp;/g,'&'):null;}""")
    if not url:
        print("CSV URL yok"); sys.exit(1)
    txt = page.evaluate(r"""async (u)=>{const r=await fetch(u,{credentials:'include'}); const buf=await r.arrayBuffer(); return new TextDecoder('windows-1254').decode(buf);}""", url)
    open("/tmp/sgk_final.csv", "w", encoding="utf-8").write(txt)

rows = list(csv.reader(io.StringIO(txt), delimiter=';'))
TIME = re.compile(r'^\d{1,2}:\d{2}$')
NUM = re.compile(r'^\d{1,3},\d{2}$')   # XX,XX
def num(s):
    try: return float(s.replace(',', '.'))
    except: return None
people, probs = [], []
for r in rows:
    if not r or '*' not in r[0]: continue   # TC satırı değil
    # ad: TC'den sonraki ilk metin hücresi (sayı/tarih/boş değil)
    ad = next((c.strip() for c in r[1:12] if c.strip() and not NUM.match(c.strip())
               and not re.match(r'^\d', c.strip())), '?')
    # ilk HH:MM kolonu
    ti = next((i for i,c in enumerate(r) if TIME.match(c.strip())), len(r))
    xx = [c.strip() for c in r[:ti] if NUM.match(c.strip())]
    gv = xx[-2] if len(xx) >= 2 else '?'
    sig = xx[-1] if len(xx) >= 1 else '?'
    people.append((ad, gv, sig))
    g = num(gv)
    if g is not None and g < 29.995:
        probs.append((ad, gv, sig))

print(f"TOPLAM {len(people)} kişi")
print(f"\n=== Gün<30 (EKSİK) ===")
for ad, gv, sig in probs:
    print(f"  {ad[:36]:36} | GV={gv:>6} Sig={sig:>6}")
print(f"\nEKSİK: {len(probs)} | TAM(30): {len(people)-len(probs)}")
