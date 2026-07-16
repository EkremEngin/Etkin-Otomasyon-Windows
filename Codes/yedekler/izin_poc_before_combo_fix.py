#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
İZİN (Yıllık İzin) Otomasyon — verify-or-halt sürümü
====================================================
Teknopark ARGE Portalı "İzin Bildirim Formu"na, firmanın gönderdiği detay izin tablosuna göre
her Ar-Ge personelinin yıllık izin günlerini TASLAK olarak giren betik. DGS betiğinin (`dgs_poc.py`)
altyapısını yeniden kullanır; veri katmanı `izin_data.py`'den gelir.

TASARIM İLKESİ — "yanlış veri yazmaktansa DUR" (DGS ile aynı):
  Her adım kendi sonucunu DOĞRULAR. Uymazsa kişi KAYDEDİLMEZ (FAILED), sonrakine temiz geçilir.
  Form "Kaydet (F3)" YALNIZCA tüm kontroller geçerse tıklanır. **ONAYA GÖNDER ASLA YAPILMAZ** — taslak
  bırakılır; onaya gönderme + (gerekiyorsa) dilekçe + e-imza İNSAN adımıdır.

ÇALIŞMA MODELİ (ATTENDED): Cloudflare + login SEN bir kez elle. Betik 9222-CDP ile bağlanır, reload YOK.

CANLI ÇIKARILAN İZİN FORMU DOM (2026-06-17, bkz DEVAM_REHBERI §0.8):
  - Menü: PERSONEL → "İzin Bildirim Formu" → liste (#Donem_Id=151 MAYIS, Yeni = .newButton).
  - Form: IzinTuruTanimi_Id=1 (Ücretli Yıllık İzin, default). string_Personel_Id autocomplete (DGS ile aynı).
  - Grid kolonları: [2]Tarih [9]Toplam [10]PdksCakisma [11]ResmiTatilMi [12]Açıklama, checkbox module_update_*,
    ✎ = span.edit-row. Disabled'ın GERÇEK sebebi AÇIKLAMA metnindedir (kolon bayrağı yanıltıcı):
      ""               → TEMİZ, izin girilebilir
      "Pdks Kaydı..."  → kişi o gün İŞTE (izin değil) → atla + UYUŞMAZLIK flag (Excel izin diyorsa)
      "Yıllık izin..." → izin ZATEN girilmiş → atla (kullanıcı: "Excel'de ne ise o")
      "...resmi tatil" → tatil (26 Kurban arifesi dahil) → otomatik sayılır → atla
  - Tam gün: checkbox JS .click() → +9h. Yarım gün: ✎ → dialog (Başlangıç/Bitiş select'leri, dialog'unki
    ekran x<500) → Bitiş 13:30 (09:00–13:30=4,5h) → "Kaydet" → satırı tikle. Dialog kapat: span.edit-gird-close.
  - Toplam: #IzinGunu / #IzinSaati tik atınca otomatik güncellenir → verify anchor.
  - "PDKS-izin çakışması" YOK (kişi izin günü kart basmaz → gün temiz gelir).

Kullanım:
  python3 izin_poc.py --person "ENES CANSU"          # tek kişi DRY-RUN
  python3 izin_poc.py --person "ENES CANSU" --commit  # tek kişi TASLAK
  python3 izin_poc.py --limit 3 --commit              # ilk 3 kişi TASLAK (döngü testi)
  python3 izin_poc.py --commit                        # tüm TPI Ar-Ge izinli (76)
"""
from __future__ import annotations
import argparse
import os
import sys
import time

from playwright.sync_api import sync_playwright, Page, TimeoutError as PWTimeout

from izin_data import build_izin_targets, IzinPerson, fold

CONFIG = {
    "cdp_url": os.environ.get("DGS_CDP", "http://localhost:9222"),  # paralel park için: DGS_CDP ile port ver (9222 BV / 9223 TPIz)
    "portal_url": "https://argeportal.teknoparkistanbul.com.tr/",
    "donem_text": "MAYIS 2026",           # dönem ETİKETİ — park-bazlı Donem_Id değişir → etiketle seç
    "ay_regex": r"^\d{2}\.05\.2026$",      # grid gün filtresi (dönem değişince güncelle)
    "prefer_start": "09:00",               # izin günü başlangıcı
    "full_day_min_default": 540,           # 9h; grid'den okunamazsa yedek
    "izin_turu_value": "1",                # Ücretli Yıllık İzin
    "izin_path": "Mayıs 2026 İzin Günleri Detayı v2.xlsx",
    "meta_path": "05-TEKNOKENTLER - MAYIS 2026 - FİNAL_9SAAT (1).xlsx",
    "dry_run": True,                       # True iken Kaydet (F3) ASLA tıklanmaz
}

LOG = "[İZİN]"


def log(*a):
    print(LOG, *a, flush=True)


class VerifyError(Exception):
    """Doğrulama başarısız → kişi KAYDEDİLMEDEN atlanır."""


class CloudflareHalt(Exception):
    """Cloudflare/oturum gerekli → insan müdahalesi."""


# ----------------------------------------------------------------------------
# Zaman yardımcıları
# ----------------------------------------------------------------------------
def to_min(hhmm: str) -> int:
    h, m = hhmm.strip().split(":")
    return int(h) * 60 + int(m)


def to_hhmm(mins: int) -> str:
    return f"{mins // 60:02d}:{mins % 60:02d}"


# ----------------------------------------------------------------------------
# Playwright bağlanma / oturum
# ----------------------------------------------------------------------------
def attach_browser(pw):
    log(f"Chrome'a bağlanılıyor: {CONFIG['cdp_url']}")
    browser = pw.chromium.connect_over_cdp(CONFIG["cdp_url"])
    ctx = browser.contexts[0] if browser.contexts else browser.new_context()
    page = next((p for p in ctx.pages if "argeportal" in (p.url or "")), None)
    if page is None:
        page = ctx.new_page()
        page.goto(CONFIG["portal_url"])
    page.bring_to_front()
    return browser, page


def assert_logged_in(page: Page):
    title = (page.title() or "")
    body = ""
    try:
        body = page.inner_text("body")[:600]
    except Exception:
        pass
    if ("Güvenlik doğrulaması" in body or "Just a moment" in title or "Bir dakika" in title
            or "Gerçek kişi" in body or "GİRİŞ" in title.upper()):
        raise CloudflareHalt("Cloudflare/giriş ekranı. Elle geç, dashboard açıkken tekrar koş.")
    if page.locator("text=PERSONEL").count() == 0:
        raise CloudflareHalt("Dashboard görülemedi (PERSONEL menüsü yok). Giriş yap, tekrar koş.")


def _open_menu_item(page: Page, item_text: str):
    """PERSONEL menüsünden öğe aç — GERÇEK click (jQuery delegated; JS click navigasyonu tetiklemez).
    Her denemede ÖNCE lingering SweetAlert overlay'i temizle: izin Kaydet'in onay/başarı popup'ı kapatılmazsa
    `.sweet-overlay` pointer-event'leri yutup PERSONEL click'ini bloklar (CANLI BULGU 2026-06-18 Yıldız)."""
    last = None
    for _ in range(3):
        try:
            page.evaluate(
                r"""() => {document.querySelectorAll('.sweet-alert .confirm, .sweet-alert button.confirm').forEach(b=>b.click());
                    document.querySelectorAll('.sweet-overlay, .sweet-alert, .ui-widget-overlay').forEach(e=>e.style.display='none');}""")
            page.get_by_text("PERSONEL", exact=True).first.click(timeout=5000)
            page.wait_for_timeout(600)
            page.get_by_text(item_text, exact=True).filter(visible=True).first.click(timeout=5000)
            page.wait_for_timeout(1000)
            return
        except Exception as e:
            last = e
            page.wait_for_timeout(500)
    raise VerifyError(f"Menü öğesi açılamadı: {item_text} ({last})")


def wait_for_form_idle(page: Page, timeout: int = 20000, settle: int = 80):
    """Yükleme bitene kadar bekle — GÖSTERGE: görünür .form_loading_panel/.pReload.loading kaybolması.
    🔴 KRİTİK (2026-06-18 Ulutek canlı): jQuery.active bu portalda KALICI bağlantılar (SignalR/poll) yüzünden
    HİÇ 0 olmuyor (sabit 5) → eski 'jQuery.active==0' koşulu HER çağrıda 20s timeout yiyordu (kişi başı 8-9 dk!).
    KALDIRILDI; panel + downstream doğrulamalar (İzinSaati==hesap, satır toplamı, poll) yeterli."""
    try:
        page.wait_for_function(
            r"""() => ![...document.querySelectorAll('.form_loading_panel, .pReload.loading')]
                  .some(e=>e.offsetParent!==null)""",
            timeout=timeout)
    except PWTimeout:
        pass
    page.wait_for_timeout(settle)


# ----------------------------------------------------------------------------
# İzin formu açma / dönem / Yeni
# ----------------------------------------------------------------------------
def _dismiss_menu(page: Page):
    """PERSONEL üst menüsü öğe tıklandıktan sonra AÇIK kalıyor (formun solunu/Personel alanını kapatıp
    seçimi+grid'i bozuyor — aralıklı 'grid boş' hatasının sebebi). Menü-dışı nötr tık ile kapat:
    önce liste penceresinin başlık çubuğu, olmazsa boş dashboard noktası."""
    try:
        page.locator(".ui-dialog-titlebar").filter(has_text="Listesi").first.click(timeout=2500)
    except Exception:
        try:
            w = int(page.evaluate("() => window.innerWidth"))
            h = int(page.evaluate("() => window.innerHeight"))
            page.mouse.click(int(w * 0.82), int(h * 0.85))
        except Exception:
            pass
    page.wait_for_timeout(400)


def open_izin_form(page: Page):
    # TEMİZ SLATE: önceki run/diagnostic'ten kalan TÜM pencereleri + SweetAlert overlay'leri kapat. Yoksa eski
    # GİZLİ liste pencerelerinin Donem_Id'si "var" sanılır (wait_for_selector hidden'ı da yakalar) / overlay menüyü bloklar.
    for _ in range(10):
        n = page.evaluate(
            r"""() => {document.querySelectorAll('.sweet-alert .confirm, .sweet-alert button.confirm').forEach(b=>b.click());
                document.querySelectorAll('.sweet-overlay, .sweet-alert, .ui-widget-overlay').forEach(e=>e.style.display='none');
                const bars=[...document.querySelectorAll('div.ui-dialog-titlebar')].filter(x=>x.offsetParent!==null);
                bars.forEach(b=>{const x=b.querySelector('.fa-times,.ui-dialog-titlebar-close'); if(x)x.click();});
                return bars.length;}""")
        page.wait_for_timeout(300)
        if n == 0:
            break
    _open_menu_item(page, "İzin Bildirim Formu")
    # GÖRÜNÜR Donem_Id bekle (wait_for_selector gizli stale select'leri de yakalar → wait_for_function görünür şart)
    page.wait_for_function(
        r"""() => [...document.querySelectorAll('#Donem_Id, select[name="Donem_Id"]')].some(e=>e.offsetParent!==null)""",
        timeout=30_000)
    _dismiss_menu(page)              # menü açık kalıyor → kapat (yoksa grid boş gelebilir)
    page.wait_for_timeout(500)


def _set_donem(page: Page) -> str:
    """Liste penceresinde Dönem'i ETİKETE göre seç (park-bazlı Donem_Id değeri değişir → etiketle bul).
    Playwright select_option (kullanıcının elle seçimiyle aynı; ham `value=` bazı combobox'larda tutmuyor —
    DEVAM_REHBERI §0.9) + value DOĞRULA. Uymazsa VerifyError. (Plan §5.2 park-bağımsızlık)"""
    label = CONFIG["donem_text"]
    val = page.evaluate(
        r"""(lbl) => {
            const sel=[...document.querySelectorAll('#Donem_Id, select[name="Donem_Id"]')].filter(e=>e.offsetParent!==null).pop();
            if(!sel) return '__YOK__';
            const o=[...sel.options].find(o=>(o.text||'').trim()===lbl);
            return o ? o.value : '__BULUNAMADI__';
        }""", label)
    if val in ("__YOK__", "__BULUNAMADI__"):
        opts = page.evaluate(
            r"""()=>{const s=[...document.querySelectorAll('#Donem_Id, select[name="Donem_Id"]')].filter(e=>e.offsetParent!==null).pop();
                return s?[...s.options].map(o=>(o.text||'').trim()).filter(Boolean):[];}""")
        raise VerifyError(f"Dönem etiketi {label!r} bulunamadı (val={val}). Mevcut dönemler: {opts}")
    page.locator('select#Donem_Id:visible, select[name="Donem_Id"]:visible').last.select_option(value=val, timeout=6000)
    page.wait_for_timeout(400)
    got = page.evaluate(
        r"""()=>{const s=[...document.querySelectorAll('#Donem_Id, select[name="Donem_Id"]')].filter(e=>e.offsetParent!==null).pop();
            return s?s.value:'';}""")
    if got != val:
        raise VerifyError(f"Dönem set edilemedi (value={got!r} != {val!r}, label={label!r}).")
    return val


def list_set_donem_and_yeni(page: Page):
    """Liste ekranında Dönem (ETİKET 'MAYIS 2026', park-bağımsız) + 'Yeni (F2)' (.newButton) → temiz giriş formu."""
    _set_donem(page)            # etikete göre seçer + içeride DOĞRULAR (uymazsa VerifyError)
    page.wait_for_timeout(300)
    clicked = page.evaluate(
        r"""() => {
            const b=[...document.querySelectorAll('a,button')].find(x=>x.offsetParent!==null
              && /Yeni\s*\(F2\)/.test(x.textContent) && /newButton/.test(x.className));
            if(!b) return false; b.click(); return true;
        }""")
    if not clicked:
        raise VerifyError("Liste 'Yeni (F2)' butonu bulunamadı.")
    page.wait_for_selector('input[name="string_Personel_Id"]', timeout=20_000)
    page.wait_for_timeout(700)


# İKİ "string_Personel_Id" input var (form + liste filtresi). Daima FORM dialog'undakini hedefle:
# liste başlığı "...Formu Listesi", form başlığı "...Bildirim Formu" (Listesi YOK).
_FIND_FORM = r"""
  const _dlg=[...document.querySelectorAll('.ui-dialog')].filter(d=>{
    const t=(d.querySelector('.ui-dialog-title')||{}).textContent||'';
    return d.offsetParent!==null && t.includes('Personel İzin Bildirim Formu') && !t.includes('Listesi');
  }).pop();
"""


def ensure_card_fresh(page: Page):
    """Form temiz mi (FORM dialog'undaki personel boş) + İzin Türü = Ücretli Yıllık İzin (1)."""
    import json
    s = json.loads(page.evaluate(
        r"""() => {
          %s
          if(!_dlg) return JSON.stringify({pv:'__yok__'});
          const p=_dlg.querySelector('input[name="string_Personel_Id"]');
          const it=_dlg.querySelector('#IzinTuruTanimi_Id, select[name="IzinTuruTanimi_Id"]');
          return JSON.stringify({pv: p?p.value.trim():'__yok__', izinTuru: it?it.value:'?'});
        }""" % _FIND_FORM))
    if s["pv"] == "__yok__":
        raise VerifyError("Form açılmadı (form dialog/personel input yok).")
    if s["pv"] != "":
        raise VerifyError(f"Form temiz değil (personel dolu: {s['pv']!r}).")
    if s["izinTuru"] not in (CONFIG["izin_turu_value"], "?"):
        raise VerifyError(f"İzin Türü beklenmedik (={s['izinTuru']}, 1=Ücretli Yıllık İzin olmalı).")


# ----------------------------------------------------------------------------
# Personel seçimi (autocomplete, çakışmaya dayanıklı)
# ----------------------------------------------------------------------------
def _autocomplete_search(page: Page, text: str) -> bool:
    """FORM dialog'undaki personel input'una yaz + o input'un jQuery autocomplete'ini tetikle."""
    return page.evaluate(
        r"""(txt) => {
          %s
          const el = _dlg ? _dlg.querySelector('input[name="string_Personel_Id"]') : null;
          if(!el) return false;
          el.focus(); el.value=txt;
          if(window.jQuery){ try{ jQuery(el).autocomplete('search', txt); }catch(e){} }
          el.dispatchEvent(new Event('input',{bubbles:true}));
          el.dispatchEvent(new Event('keyup',{bubbles:true}));
          return true;
        }""" % _FIND_FORM, text)


_PICK_LI_JS = r"""(toks) => {
  const fold=s=>{const tr={'ç':'c','Ç':'c','ğ':'g','Ğ':'g','ı':'i','İ':'i','I':'i','ö':'o','Ö':'o','ş':'s','Ş':'s','ü':'u','Ü':'u'};
    return s.split('').map(c=>tr[c]||c).join('').toLowerCase();};
  const li=[...document.querySelectorAll('ul.ui-autocomplete li, ul.ui-menu li')]
    .filter(x=>x.offsetParent!==null && toks.every(t=>fold(x.textContent).includes(t)))[0];
  if(!li) return false;
  const tgt=li.querySelector('a,div')||li;
  for(const t of ['mouseover','mouseenter','mousemove','mousedown','mouseup','click'])
    tgt.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:window}));
  return true;
}"""

_LI_VISIBLE_JS = r"""(toks) => {
  const fold=s=>{const tr={'ç':'c','Ç':'c','ğ':'g','Ğ':'g','ı':'i','İ':'i','I':'i','ö':'o','Ö':'o','ş':'s','Ş':'s','ü':'u','Ü':'u'};
    return s.split('').map(c=>tr[c]||c).join('').toLowerCase();};
  const lis=[...document.querySelectorAll('ul.ui-autocomplete li, ul.ui-menu li')].filter(li=>li.offsetParent!==null);
  return lis.some(li=>toks.every(t=>fold(li.textContent).includes(t)));
}"""


def _clear_personel(page: Page):
    page.evaluate(
        r"""() => { %s if(!_dlg) return;
          const el=_dlg.querySelector('input[name="string_Personel_Id"]');
          if(el){ el.value=''; el.dispatchEvent(new Event('input',{bubbles:true})); }
          const h=_dlg.querySelector('input[type=hidden][name="Personel_Id"]'); if(h) h.value='';
        }""" % _FIND_FORM)


def select_personel(page: Page, portal_ad: str):
    """Personel seç + gizli Personel_Id + grid'i DOĞRULA. ROBUST: autocomplete widget bağlanana kadar bekle
    (reopen sonrası geç bağlanıyor → erken seçim grid'i boş bırakıyor), sonra TÜM seçimi (ara→öneri→tıkla→
    gizli-id→grid) grid dolana kadar 3 kez dene; denemeler arası alanı temizle. Soyadla ara (boşluklu tam-ad
    araması grid callback'ini tetiklemiyor), token-eşleşmesiyle çakışan soyadlarda doğru kişiyi seç."""
    import json
    tokens = fold(portal_ad).split()
    surname = portal_ad.split()[-1]
    # autocomplete widget'ı (grid-yükleme callback'i onunla gelir) bağlanana kadar bekle
    try:
        page.wait_for_function(
            r"""() => { %s const el=_dlg?_dlg.querySelector('input[name="string_Personel_Id"]'):null;
                  return !!(el && window.jQuery && jQuery(el).data('ui-autocomplete')); }""" % _FIND_FORM,
            timeout=10000)
    except PWTimeout:
        pass
    last_err = None
    for attempt in range(3):
        term = surname if attempt < 2 else portal_ad        # 3. denemede tam ad
        if not _autocomplete_search(page, term):
            raise VerifyError("Form personel input bulunamadı.")
        try:
            page.wait_for_function(_LI_VISIBLE_JS, arg=tokens, timeout=6000)
        except PWTimeout:
            last_err = f"'{term}' önerisi çıkmadı"
            _clear_personel(page); page.wait_for_timeout(800); continue
        if not page.evaluate(_PICK_LI_JS, tokens):
            last_err = "öneri tıklanamadı"
            _clear_personel(page); page.wait_for_timeout(800); continue
        page.wait_for_timeout(1200)
        s = json.loads(page.evaluate(
            r"""() => { %s if(!_dlg) return JSON.stringify({err:'form yok'});
              const el=_dlg.querySelector('input[name="string_Personel_Id"]');
              const h=_dlg.querySelector('input[type=hidden][name="Personel_Id"]');
              return JSON.stringify({val: el?el.value:'', pid: h?h.value:''});
            }""" % _FIND_FORM))
        if s.get("err"):
            raise VerifyError("Form dialog bulunamadı (personel teyidi).")
        if not s["pid"] or s["pid"] == "0":
            last_err = "gizli Personel_Id set olmadı"
            _clear_personel(page); page.wait_for_timeout(1000); continue
        if not all(t in fold(s["val"]) for t in tokens):
            raise VerifyError(f"Yanlış/eksik personel: beklenen '{portal_ad}', input='{s['val']}'")
        # grid dolsun
        wait_for_form_idle(page)
        try:
            page.wait_for_function(
                r"""(re)=>[...document.querySelectorAll('tr')].some(tr=>tr.children[2]
                      && new RegExp(re).test((tr.children[2].textContent||'').trim()))""",
                arg=CONFIG["ay_regex"], timeout=12000)
        except PWTimeout:
            last_err = "seçim tuttu ama grid boş"
            log(f"  (deneme {attempt+1}: grid boş, tekrar)")
            _clear_personel(page); page.wait_for_timeout(1500); continue
        wait_for_form_idle(page)
        log(f"  Personel seçildi: {s['val']} (Personel_Id={s['pid']})")
        return
    raise VerifyError(f"Personel seçilemedi/grid boş: {portal_ad} ({last_err})")


# ----------------------------------------------------------------------------
# Grid okuma + sınıflandırma
# ----------------------------------------------------------------------------
def read_grid(page: Page) -> dict[str, dict]:
    """{tarih: {toplam, ack, dis, durum}}. durum: clean|izin_var|pdks|tatil|other."""
    import json
    rows = json.loads(page.evaluate(
        r"""(re) => {
          const rx=new RegExp(re); const out=[];
          document.querySelectorAll('tr').forEach(tr=>{const td=tr.children; if(td.length<13) return;
            const t=(td[2]?.textContent||'').trim(); if(!rx.test(t)) return;
            const cb=tr.querySelector('input[type=checkbox][name^="module_update_"]');
            out.push({tarih:t, toplam:(td[9]?.textContent||'').trim(), ack:(td[12]?.textContent||'').trim(),
              dis: cb?cb.disabled:true, checked: cb?cb.checked:false});});
          return JSON.stringify(out);
        }""", CONFIG["ay_regex"]))
    grid = {}
    for r in rows:
        ack = r["ack"]
        low = fold(ack)
        if not r["dis"]:
            durum = "clean"
        elif "yillik izin kaydi" in low:
            durum = "izin_var"
        elif "resmi tatil" in low:
            durum = "tatil"
        elif "pdks kaydi" in low:
            durum = "pdks"
        else:
            durum = "other"
        r["durum"] = durum
        grid[r["tarih"]] = r
    return grid


def read_full_day_min(page: Page, grid: dict) -> int:
    """Temiz bir günün varsayılan toplamından günlük tam-süreyi (dk) türet (TPI=9h=540)."""
    for r in grid.values():
        if r["durum"] == "clean" and r["toplam"]:
            try:
                return to_min(r["toplam"])
            except Exception:
                pass
    return CONFIG["full_day_min_default"]


def read_totals(page: Page) -> tuple[str, str]:
    import json
    s = json.loads(page.evaluate(
        r"""() => JSON.stringify({g: (document.querySelector('#IzinGunu, input[name="IzinGunu"]')||{}).value || '',
                                  s: (document.querySelector('#IzinSaati, input[name="IzinSaati"]')||{}).value || ''})"""))
    return s["g"].strip(), s["s"].strip()


# ----------------------------------------------------------------------------
# Gün işleme: tam gün tikle / yarım gün ✎ ile düzenle
# ----------------------------------------------------------------------------
def tick_day(page: Page, tarih: str):
    page.evaluate(
        r"""(d) => {const r=[...document.querySelectorAll('tr')].find(tr=>tr.children[2]&&tr.children[2].textContent.trim()===d);
              const cb=r&&r.querySelector('input[type=checkbox][name^="module_update_"]');
              if(cb&&!cb.disabled&&!cb.checked)cb.click();}""", tarih)
    chk = page.evaluate(
        r"""(d) => {const r=[...document.querySelectorAll('tr')].find(tr=>tr.children[2]&&tr.children[2].textContent.trim()===d);
              const cb=r&&r.querySelector('input[type=checkbox][name^="module_update_"]'); return cb?cb.checked:false;}""", tarih)
    if not chk:
        raise VerifyError(f"{tarih}: tiklenemedi.")


def _open_edit(page: Page, tarih: str) -> bool:
    ok = page.evaluate(
        r"""(d) => {const r=[...document.querySelectorAll('tr')].find(tr=>tr.children[2]&&tr.children[2].textContent.trim()===d);
              if(!r) return false; const e=r.querySelector('span.edit-row'); if(!e) return false; e.click(); return true;}""",
        tarih)
    if ok:
        wait_for_form_idle(page)   # ✎ dialog ~2-3 sn yükleme → bekle (loading-fix, TEST'te gözlemle)
    return ok


def _dialog_open(page: Page) -> bool:
    """Yarım-gün dialog'u açık mı? Dialog'a ÖZGÜ select'lerin (Bitis_Saat/Baslangic_Saat, ALT ÇİZGİLİ)
    varlığıyla tespit edilir — başlık metnine ('Izin Talep Formu Bilgileri', düz I) güvenmek İ/I tuzağı."""
    return page.evaluate(
        r"""() => [...document.querySelectorAll('select[name="Bitis_Saat"], select[name="Baslangic_Saat"]')]
              .some(e=>e.offsetParent!==null)""")


def _set_dialog_time(page: Page, start: int, end: int):
    """Yarım-gün dialog'unun saat select'lerini set + DOĞRULA. CANLI BULGU: izin dialog'u DGS'ten FARKLI —
    select adları ALT ÇİZGİLİ: Baslangic_Saat/Baslangic_Dakika/Bitis_Saat/Bitis_Dakika (form panelininki
    BaslangicSaat/BitisSaat, alt çizgisiz, x~623). Bu adlar dialog'a özgü olduğundan x filtresine gerek yok."""
    import json
    res = page.evaluate(
        r"""({sh,sm,eh,em}) => {
          const setL=(name,v)=>{const el=[...document.querySelectorAll(`select[name="${name}"]`)]
              .filter(e=>e.offsetParent!==null).pop();
            if(!el) return null;
            el.value=String(v).padStart(2,'0'); el.dispatchEvent(new Event('change',{bubbles:true}));
            if(window.jQuery) jQuery(el).trigger('change'); return el.value;};
          return JSON.stringify({bs:setL('Baslangic_Saat',sh), bd:setL('Baslangic_Dakika',sm),
                                 es:setL('Bitis_Saat',eh), ed:setL('Bitis_Dakika',em)});
        }""", {"sh": start // 60, "sm": start % 60, "eh": end // 60, "em": end % 60})
    r = json.loads(res)
    want = {"bs": f"{start//60:02d}", "bd": f"{start%60:02d}", "es": f"{end//60:02d}", "ed": f"{end%60:02d}"}
    if r != want:
        raise VerifyError(f"Dialog saat set edilemedi: oldu={r} beklenen={want}")


def _dialog_save(page: Page):
    """Dialog 'Kaydet' (F3 İÇERMEYEN) → kapanana kadar dene. Her tıklamadan sonra loading beklenir
    (save AJAX bitmeden 'kapandı mı' bakma — loading-fix, TEST'te gözlemle)."""
    wait_for_form_idle(page)
    for _ in range(3):
        page.evaluate(
            r"""() => {const b=[...document.querySelectorAll('button,a')].filter(x=>x.offsetParent!==null
                  && /Kaydet/i.test(x.textContent) && !/F3/.test(x.textContent)).pop();
                if(b) b.click();}""")
        wait_for_form_idle(page)
        if not _dialog_open(page):
            return
    if _dialog_open(page):
        raise VerifyError("Dialog Kaydet tutmadı (dialog hâlâ açık).")


def set_half_day(page: Page, tarih: str, half_min: int):
    """Yarım gün: ✎ → dialog → 09:00–(09:00+half) → Kaydet → satır toplam doğrula → tikle."""
    if not _open_edit(page, tarih):
        raise VerifyError(f"{tarih}: ✎ ikonu bulunamadı.")
    page.wait_for_timeout(600)
    if not _dialog_open(page):
        raise VerifyError(f"{tarih}: yarım-gün dialogu açılmadı.")
    start = to_min(CONFIG["prefer_start"])
    end = start + half_min
    _set_dialog_time(page, start, end)
    _dialog_save(page)
    # satır toplam == half olmalı
    row = read_grid(page).get(tarih)
    if not row or row["toplam"] != to_hhmm(half_min):
        raise VerifyError(f"{tarih}: yarım gün sonrası satır toplam {row['toplam'] if row else '?'} != {to_hhmm(half_min)}")
    tick_day(page, tarih)
    log(f"  {tarih}: yarım gün ({to_hhmm(start)}–{to_hhmm(end)}={to_hhmm(half_min)}) ✓")


# ----------------------------------------------------------------------------
# Reopen (sonraki kişi için temiz form) + Kaydet
# ----------------------------------------------------------------------------
def reopen_fresh_form(page: Page):
    """Sonraki kişi: SADECE form penceresini kapat (liste 'Listesi' penceresi kalsın), listeyi yeniden kullan."""
    page.evaluate(
        r"""() => {const bar=[...document.querySelectorAll('div.ui-dialog-titlebar')]
              .filter(b=>b.offsetParent!==null && /Personel İzin Bildirim Formu/.test(b.textContent)
                       && !/Listesi/.test(b.textContent))[0];
            if(bar){const x=bar.querySelector('.fa-times, .ui-dialog-titlebar-close'); if(x)x.click();}}""")
    page.wait_for_timeout(700)
    has_list = page.evaluate(
        r"""()=>[...document.querySelectorAll('#Donem_Id, select[name="Donem_Id"]')].some(e=>e.offsetParent!==null)""")
    if not has_list:
        open_izin_form(page)           # liste de kapandıysa menüden temiz aç
    list_set_donem_and_yeni(page)
    ensure_card_fresh(page)


def save_draft(page: Page) -> str:
    """Beyan (Sozlesme) checkbox'ını işaretle (kaydetmek İÇİN ZORUNLU — 'belirtilen şartları onaylamalısınız')
    → Form 'Kaydet (F3)' → TASLAK. dry_run'da tıklamaz. ONAYA GÖNDER ASLA."""
    if CONFIG["dry_run"]:
        return "DRY-RUN (kaydedilmedi)"
    # 1) zorunlu beyan/şartlar checkbox'ı (Begüm'ün de elle işaretlediği) — FORM dialog'unda
    chk = page.evaluate(
        r"""() => {
          %s
          const cb=(_dlg||document).querySelector('#Sozlesme, input[name="Sozlesme"]');
          if(!cb) return 'yok';
          if(!cb.checked){ cb.click(); if(window.jQuery) jQuery(cb).trigger('change'); }
          return cb.checked ? 'checked' : 'fail';
        }""" % _FIND_FORM)
    if chk != "checked":
        raise VerifyError(f"Beyan (Sozlesme) kutusu işaretlenemedi (={chk}).")
    page.wait_for_timeout(200)
    # 2) Kaydet (F3) — 🔴 GERÇEK Playwright click ŞART (CANLI BULGU 2026-06-18 Yıldız): evaluate-click jQuery
    #    save handler'ını GÜVENİLMEZ tetikliyor — bazen HİÇ kaydetmiyor (veri durur, mesaj yok), bazen takılı
    #    SweetAlert overlay'i bırakıp sonraki menüyü blokluyor. Gerçek click sağlam kaydediyor (Onaya Gönder tuzağı gibi).
    clicked = False
    for force in (False, True):
        try:
            page.locator("a:has-text('Kaydet (F3)'), button:has-text('Kaydet (F3)')").filter(visible=True).last.click(timeout=4000, force=force)
            clicked = True
            break
        except Exception:
            page.wait_for_timeout(300)
    if not clicked:                       # son çare: evaluate-click
        page.evaluate(
            r"""() => {const b=[...document.querySelectorAll('a,button')].find(x=>x.offsetParent!==null
                  && /Kaydet/i.test(x.textContent) && /F3/.test(x.textContent)); if(b)b.click();}""")
    # olası onay SweetAlert'i (Evet/Eminim/Tamam/Onayla) + 'başarı' poll (~5s)
    ok = False
    for _ in range(25):
        page.wait_for_timeout(200)
        page.evaluate(
            r"""() => {const s=[...document.querySelectorAll('.sweet-alert')].filter(e=>e.offsetParent!==null)[0];
                if(!s) return; if(/emin|onayla|kaydetmek/i.test(s.innerText)){
                  const c=[...s.querySelectorAll('button,a')].find(x=>/^(Evet|Eminim|Tamam|Onayla)$/i.test(x.textContent.trim())); if(c)c.click();}}""")
        body = page.inner_text("body").lower()
        if "başarı" in body or "kayıt altına" in body or "kaydedildi" in body:
            ok = True
            break
    # açık kalan SweetAlert/overlay'i KAPAT (sonraki kişinin menüsünü bloklamasın — Yıldız'da kritik)
    page.evaluate(
        r"""() => {document.querySelectorAll('.sweet-alert .confirm, .sweet-alert button.confirm').forEach(b=>b.click());
            document.querySelectorAll('.sweet-overlay, .sweet-alert').forEach(e=>e.style.display='none');}""")
    page.wait_for_timeout(300)
    if ok:
        return "TASLAK kaydedildi"
    warn = page.evaluate(
        r"""() => {const el=[...document.querySelectorAll('.sweet-alert, .modal, .bootbox, .jconfirm, [class*=alert], [class*=popup]')]
              .filter(e=>e.offsetParent!==null).map(e=>e.innerText.replace(/\s+/g,' ').trim()).filter(Boolean); return el[0]||'';}""")
    raise VerifyError(f"Kaydet sonrası başarı mesajı yok. Ekrandaki uyarı: {warn[:140]!r}")


# ----------------------------------------------------------------------------
# Tek kişi
# ----------------------------------------------------------------------------
def process_one_person(page: Page, person: IzinPerson, first: bool) -> dict:
    if first:
        open_izin_form(page)
        list_set_donem_and_yeni(page)
        ensure_card_fresh(page)
    else:
        reopen_fresh_form(page)

    select_personel(page, person.portal_ad)
    grid = read_grid(page)
    full_min = read_full_day_min(page, grid)
    half_min = full_min // 2

    girilen, atlanan_izin, atlanan_tatil, flag = [], [], [], []
    exp_min = 0
    for tarih, gun in person.gunler:
        row = grid.get(tarih)
        if row is None:
            flag.append(f"{tarih}: grid'de yok (hafta sonu/dönem dışı?)")
            continue
        durum = row["durum"]
        if durum == "izin_var":
            atlanan_izin.append(tarih)                      # zaten girilmiş → "Excel'de ne ise o"
            continue
        if durum == "tatil":
            atlanan_tatil.append(tarih)                     # resmi tatil (26 arife dahil) → otomatik sayılır
            continue
        if durum in ("pdks", "other"):
            flag.append(f"{tarih}: '{row['ack']}' (Excel izin diyor ama portal kapalı) → ELLE BAK")
            continue
        # clean → gir
        if gun == 0.5:
            set_half_day(page, tarih, half_min)
            exp_min += half_min
        else:
            tick_day(page, tarih)
            exp_min += full_min
            log(f"  {tarih}: tam gün ✓")
        girilen.append((tarih, gun))

    # toplam doğrula (girilen kadar)
    g_val, s_val = read_totals(page)
    saat_ok = True
    try:
        saat_ok = (to_min(s_val) == exp_min) if (s_val and exp_min) else (exp_min == 0)
    except Exception:
        saat_ok = False
    if girilen and not saat_ok:
        raise VerifyError(f"İzin toplamı uyuşmuyor: portal IzinSaati={s_val!r} (gün={g_val!r}) hesap={to_hhmm(exp_min)}")

    if not girilen:
        # girilecek yeni gün yok (hepsi zaten girili/tatil ya da flag) → KAYDETME
        msg = "YENİ GİRİŞ YOK"
        if flag:
            msg += " | FLAG: " + "; ".join(flag)
        return {"ad": person.portal_ad, "ok": not flag, "kaydedildi": False, "mesaj": msg,
                "girilen": girilen, "atlanan_izin": atlanan_izin, "atlanan_tatil": atlanan_tatil, "flag": flag}

    msg = save_draft(page)
    durum_ek = ""
    if flag:
        durum_ek = " | ⚠ FLAG: " + "; ".join(flag)
    return {"ad": person.portal_ad, "ok": True, "kaydedildi": "DRY-RUN" not in msg,
            "mesaj": f"{msg} (gün={g_val}, saat={s_val}){durum_ek}",
            "girilen": girilen, "atlanan_izin": atlanan_izin, "atlanan_tatil": atlanan_tatil, "flag": flag}


# ----------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="İzin Otomasyon (verify-or-halt; TASLAK; Onaya Gönder YOK)")
    ap.add_argument("--izin", default=CONFIG["izin_path"])
    ap.add_argument("--meta", default=CONFIG["meta_path"])
    ap.add_argument("--lokasyon", default="TPI")
    ap.add_argument("--person", default=None, help="Sadece tek kişi (portal/Excel adıyla, kısmi olabilir)")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--commit", action="store_true", help="TASLAK kaydet (Onaya Gönder YİNE YOK)")
    ap.add_argument("--include-destek", action="store_true", help="Ar-Ge olmayan (Destek) izinli kişileri de işle (4691 %10 Destek)")
    args = ap.parse_args()
    CONFIG["dry_run"] = not args.commit

    targets, unresolved = build_izin_targets(args.izin, args.meta, args.lokasyon, only_arge=not args.include_destek)
    if args.person:
        key = fold(args.person)
        targets = [p for p in targets if key in fold(p.portal_ad) or key in fold(p.detay_ad)]
        if not targets:
            log(f"HATA: '{args.person}' bulunamadı (lokasyon={args.lokasyon}).")
            sys.exit(1)
    if args.limit:
        targets = targets[: args.limit]

    log(f"Hedef: {len(targets)} kişi (lokasyon={args.lokasyon}, Ar-Ge, izinli) | commit={args.commit} (False=DRY-RUN)")
    if unresolved:
        log(f"UYARI: {len(unresolved)} kişi meta'sız (atlandı): {[u['ad'] for u in unresolved]}")
    if not targets:
        sys.exit(1)

    done_file = f"izin_done_{args.lokasyon}_Mayıs.txt"
    done = set(open(done_file, encoding="utf-8").read().splitlines()) if os.path.exists(done_file) else set()
    if done:
        log(f"{len(done)} kişi zaten işlenmiş (atlanacak).")

    results = []
    with sync_playwright() as pw:
        browser, page = attach_browser(pw)
        try:
            assert_logged_in(page)
            first = True
            for i, person in enumerate(targets, 1):
                if person.portal_ad in done:
                    continue
                log(f"\n=== [{i}/{len(targets)}] {person.portal_ad} | Excel: {person.gunler} ===")
                try:
                    assert_logged_in(page)
                    r = process_one_person(page, person, first)
                    first = False
                except CloudflareHalt as e:
                    log(f"!! DURDU (Cloudflare/oturum): {e}")
                    break
                except VerifyError as e:
                    ts = int(time.time()); shot = f"izin_verify_{ts}.png"
                    try: page.screenshot(path=shot, full_page=True)
                    except Exception: pass
                    r = {"ad": person.portal_ad, "ok": False, "kaydedildi": False,
                         "mesaj": f"DOĞRULAMA HATASI (KAYDEDİLMEDİ): {e} (ss:{shot})", "flag": []}
                    first = False
                except Exception as e:
                    ts = int(time.time()); shot = f"izin_hata_{ts}.png"
                    try: page.screenshot(path=shot, full_page=True)
                    except Exception: pass
                    r = {"ad": person.portal_ad, "ok": False, "kaydedildi": False,
                         "mesaj": f"HATA (KAYDEDİLMEDİ): {e} (ss:{shot})", "flag": []}
                    first = False
                results.append(r)
                log(f"  -> {r['mesaj']}")
                if r.get("kaydedildi") and args.commit:
                    with open(done_file, "a", encoding="utf-8") as f:
                        f.write(person.portal_ad + "\n")

            ok = sum(1 for r in results if r["ok"])
            kaydli = sum(1 for r in results if r.get("kaydedildi"))
            log(f"\n==== ÖZET: {ok}/{len(results)} sorunsuz | {kaydli} TASLAK kaydedildi ====")
            for r in results:
                if not r["ok"]:
                    log(f"  !! BAŞARISIZ: {r['ad']} — {r['mesaj']}")
                elif r.get("flag"):
                    log(f"  ⚠  İNCELE: {r['ad']} — {'; '.join(r['flag'])}")
            log("HATIRLATMA: Hepsi TASLAK. 'Onaya Gönder' + (gerekiyorsa) dilekçe/e-imza SENDE.")
            try:
                input(f"{LOG} >> Bitti. ENTER ile kapat...")
            except EOFError:
                pass
        except CloudflareHalt as e:
            log(f"GENEL DURUŞ (Cloudflare/oturum): {e}")
        except Exception as e:
            log(f"GENEL HATA: {e}")
            raise


if __name__ == "__main__":
    main()
