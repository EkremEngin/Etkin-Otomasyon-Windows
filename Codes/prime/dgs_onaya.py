#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DGS TASLAK'larını ONAYA GÖNDER — dgs_poc.py'nin kaydettiği taslakları onaya gönderir.
=====================================================================================
Akış (kullanıcı canlı gösterdi 2026-06-18):
  PERSONEL → Dışarıda Geçirilen Süreler Formu → "Dışarıda Geçirilen Süreler Listesi"
  → Filtre → DÖNEM seç (ZORUNLU; yoksa liste boş döner!) → "Listele (F8)"
  → kayda ÇİFT TIKLA → ~3 sn form açılır → "Onaya Gönder" → (varsa "Evet/Emin misiniz")
  → kayıt "Firma Yetkilisi E-imzası Bekleniyor" statüsüne düşer.
  E-İMZA = İNSAN ADIMI (script atmaz). 'Onaya Gönder' tıklanamıyorsa SIRADAKİ kişiye geçer.

DÖNEM otomatik: bugünden bir ÖNCEKİ takvim ayı (ör. Haziran'dayız → MAYIS 2026). İleri dönemler
için de kendiliğinden doğru ayı seçer. Listede o etiket yoksa, dropdown'daki SON aydan bir öncekine düşer.

GÜVENLİK (verify-or-halt):
  - Yalnız dgs_done_<lok>_<sheet>.txt'deki isimleri işler; listede bizim-olmayan kayda DOKUNMAZ.
  - Onaya Gönder sonrası kaydın taslak listesinden DÜŞTÜĞÜNÜ doğrular; düşmediyse "GİTMEDİ" der, atlar.
  - VARSAYILAN DRY-RUN: sadece listeler + eşleştirir + statü dağılımı raporlar (HİÇBİR gönderme yok).
  - --commit ile gerçek 'Onaya Gönder'. GERİ ALINMASI ZOR (resmi onay akışı) — kullanıcı istedi.
  - Resume: dgs_onaya_done_<lok>_<sheet>.txt. Sayfa RELOAD YOK (Cloudflare).

Kullanım:
  python3 dgs_onaya.py --excel ... --dry-run            # önizleme (varsayılan)
  python3 dgs_onaya.py --excel ... --commit --limit 1   # GERÇEK: 1 kişi (test)
  python3 dgs_onaya.py --excel ... --commit             # hepsi
"""
from __future__ import annotations
import argparse
import datetime
import os
import re
import sys

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

import dgs_poc as D

LOG = "[ONAY-DGS]"
TR_AYLAR = {1: "OCAK", 2: "ŞUBAT", 3: "MART", 4: "NİSAN", 5: "MAYIS", 6: "HAZİRAN",
            7: "TEMMUZ", 8: "AĞUSTOS", 9: "EYLÜL", 10: "EKİM", 11: "KASIM", 12: "ARALIK"}


def log(*a):
    print(LOG, *a, flush=True)


def fold(s: str) -> str:
    """Türkçe-duyarsız normalize (isim eşleştirme için)."""
    s = (s or "").replace("̇", "")
    tr = {'ç': 'c', 'Ç': 'c', 'ğ': 'g', 'Ğ': 'g', 'ı': 'i', 'İ': 'i', 'I': 'i',
          'ö': 'o', 'Ö': 'o', 'ş': 's', 'Ş': 's', 'ü': 'u', 'Ü': 'u'}
    s = "".join(tr.get(c, c) for c in s)
    return " ".join(s.lower().split())


def prev_month_label(today: datetime.date | None = None) -> str:
    """Bugünden bir önceki takvim ayı → 'MAYIS 2026' gibi."""
    today = today or datetime.date.today()
    y, m = today.year, today.month
    m -= 1
    if m == 0:
        m = 12
        y -= 1
    return f"{TR_AYLAR[m]} {y}"


# ---- DOM yardımcıları (Listesi penceresi scope'lu) ----
LIST_TITLE = "Dışarıda Geçirilen Süreler Listesi"
FORM_TITLE = "Dışarıda Geçirilen Süreler Formu"


def close_all_dialogs(page):
    for _ in range(10):
        n = page.evaluate(
            r"""()=>{const bars=[...document.querySelectorAll('div.ui-dialog-titlebar')].filter(b=>b.offsetParent!==null);
              bars.forEach(b=>{const x=b.querySelector('.fa-times,.ui-dialog-titlebar-close'); if(x)x.click();}); return bars.length;}""")
        page.wait_for_timeout(350)
        if n == 0:
            break


def close_form_only(page):
    """SADECE form penceresini kapat (LİSTE açık kalsın)."""
    page.evaluate(
        r"""(ft)=>{const bar=[...document.querySelectorAll('div.ui-dialog-titlebar')]
              .filter(b=>b.offsetParent!==null && new RegExp(ft).test(b.textContent) && !/Listesi/.test(b.textContent))[0];
            if(bar){const x=bar.querySelector('.fa-times,.ui-dialog-titlebar-close'); if(x)x.click();}}""", FORM_TITLE)
    page.wait_for_timeout(700)


def open_list(page):
    close_all_dialogs(page)
    D._open_menu_item(page, FORM_TITLE)
    # 'select görünür' beklemek kırılgan (Filtre paneli kapalıyken Donem_Id gizli) → Listesi penceresini bekle
    page.wait_for_function(
        r"""()=>[...document.querySelectorAll('div.ui-dialog-titlebar')].some(b=>b.offsetParent!==null && /Dışarıda Geçirilen Süreler Listesi/.test(b.textContent))""",
        timeout=30_000)
    page.wait_for_timeout(1200)
    ensure_filter_open(page)


def _donem_visible(page) -> bool:
    return page.evaluate(
        r"""()=>{const dlg=[...document.querySelectorAll('div.ui-dialog')].filter(d=>d.offsetParent!==null && /Dışarıda Geçirilen Süreler Listesi/.test(d.querySelector('.ui-dialog-titlebar')?.textContent||''))[0];
            if(!dlg)return false; const s=dlg.querySelector('select[name="Donem_Id"], select#Donem_Id'); return !!(s && s.offsetParent!==null);}""")


def ensure_filter_open(page):
    """Filtre paneli (Tablo Filtre Seçenekleri) — Dönem select görünür değilse 'Filtre'ye bas + bekle."""
    if _donem_visible(page):
        return
    page.evaluate(
        r"""()=>{const dlg=[...document.querySelectorAll('div.ui-dialog')].filter(d=>d.offsetParent!==null && /Listesi/.test(d.querySelector('.ui-dialog-titlebar')?.textContent||''))[0];
          const b=[...dlg.querySelectorAll('a,button')].find(x=>x.offsetParent!==null && /^Filtre$/.test(x.textContent.trim())); if(b)b.click();}""")
    for _ in range(10):
        page.wait_for_timeout(500)
        if _donem_visible(page):
            return
    raise D.VerifyError("Filtre paneli açılamadı (Dönem select görünmedi).")


def donem_options(page) -> list[dict]:
    return page.evaluate(
        r"""()=>{const dlg=[...document.querySelectorAll('div.ui-dialog')].filter(d=>d.offsetParent!==null && /Listesi/.test(d.querySelector('.ui-dialog-titlebar')?.textContent||''))[0];
            const sel=dlg?dlg.querySelector('select[name="Donem_Id"], select#Donem_Id'):null;
            if(!sel)return []; return [...sel.options].map(o=>({value:o.value,text:o.text.trim()})).filter(o=>o.value&&o.value!=='0');}""")


def set_donem(page, label: str) -> str:
    """Listesi filtresinde Dönem'i seç (Playwright select_option — kullanıcının elle seçimiyle aynı).
    Etiket yoksa dropdown'daki SON aydan bir öncekisine düşer (ileri dönem güvenliği)."""
    opts = donem_options(page)
    val = next((o["value"] for o in opts if o["text"] == label), None)
    chosen = label
    if not val:
        try:
            ordered = sorted([o for o in opts if o["value"].isdigit()], key=lambda o: int(o["value"]))
            if len(ordered) >= 2:
                val, chosen = ordered[-2]["value"], ordered[-2]["text"]
                log(f"  UYARI: '{label}' dönem listesinde yok → '{chosen}' seçildi (son-aydan-önceki).")
        except Exception:
            pass
    if not val:
        raise D.VerifyError(f"Dönem ayarlanamadı: '{label}' yok. Mevcut: {[o['text'] for o in opts]}")
    page.locator('div.ui-dialog select[name="Donem_Id"]').last.select_option(value=val, timeout=6000)
    page.wait_for_timeout(500)
    got = page.evaluate(
        r"""()=>{const dlg=[...document.querySelectorAll('div.ui-dialog')].filter(d=>d.offsetParent!==null && /Listesi/.test(d.querySelector('.ui-dialog-titlebar')?.textContent||''))[0];
            const s=dlg.querySelector('select[name="Donem_Id"], select#Donem_Id'); return s?s.value:'';}""")
    if got != val:
        raise D.VerifyError(f"Dönem set edilemedi (value={got!r} != {val!r}).")
    return chosen


def set_status_filter(page, only_draft: bool):
    """Durumlar: only_draft ise yalnız 'Değerlendirmeye Gönderilmemiş' (Id_5); değilse HEPSİ açık."""
    page.evaluate(
        r"""(draft)=>{document.querySelectorAll('input[type=checkbox][id^="OnayDurumu_Id"]').forEach(cb=>{
              const want = draft ? (cb.id==='OnayDurumu_Id_5') : true;
              if(cb.checked!==want) cb.click();});}""", only_draft)
    page.wait_for_timeout(300)


def footer_total(page) -> str:
    """'Toplam N kayıt' alt-bilgisini döndür (gerçek toplam; sayfa başı 50 gösterse de)."""
    return page.evaluate(
        r"""()=>{const dlg=[...document.querySelectorAll('div.ui-dialog')].filter(d=>d.offsetParent!==null && /Dışarıda Geçirilen Süreler Listesi/.test(d.querySelector('.ui-dialog-titlebar')?.textContent||''))[0];
            if(!dlg)return ''; const m=(dlg.innerText.match(/Toplam[^\n]*kayıt[^\n]*/i)||[]); return m[0]||'';}""")


def listele(page):
    """Filtre panelindeki 'Listele (F8)' butonuna bas — GERÇEK Playwright click (canlı çalışan yol)."""
    try:
        page.locator("a:has-text('Listele (F8)')").filter(visible=True).first.click(timeout=4000)
    except Exception:
        clicked = page.evaluate(
            r"""()=>{const dlg=[...document.querySelectorAll('div.ui-dialog')].filter(d=>d.offsetParent!==null && /Listesi/.test(d.querySelector('.ui-dialog-titlebar')?.textContent||''))[0];
                if(!dlg)return false;
                const b=[...dlg.querySelectorAll('a,button')].find(x=>x.offsetParent!==null
                  && (/helper_kayitlistele/.test(x.getAttribute('onclick')||'') || (/Listele/.test(x.textContent)&&/listButton/.test(x.className))));
                if(!b)return false; b.click(); return true;}""")
        if not clicked:
            page.keyboard.press("F8")
    D.wait_for_form_idle(page)   # listeleme AJAX'ı bitene kadar bekle (loading-fix; TEST'te gözlemle)


_NAME_RE = re.compile(r"^(.+?)\s+[\d\*]{3,}")


def read_list_rows(page) -> list[dict]:
    """Sonuç grid'ini oku — flexigrid GÖVDESİ (.flexigrid .bDiv). Gizli id kolonları nedeniyle:
       id=td[0], Personel=td[6], Onay Durumu=td[12], OnayDurumu_Id=td[13], Toplam=td[15].
    İsim TC'den arındırılır + fold'lanır."""
    raw = page.evaluate(
        r"""()=>{const dlg=[...document.querySelectorAll('div.ui-dialog')].filter(d=>d.offsetParent!==null && /Dışarıda Geçirilen Süreler Listesi/.test(d.querySelector('.ui-dialog-titlebar')?.textContent||''))[0];
            if(!dlg)return [];
            const body=dlg.querySelector('.flexigrid .bDiv'); if(!body)return [];
            return [...body.querySelectorAll('tr')]
              .filter(tr=>tr.children.length>=13 && /FEV TR/.test(tr.children[4]?.textContent||''))
              .map(tr=>({id:(tr.children[0]?.textContent||'').trim(),
                         pers:(tr.children[6]?.textContent||'').replace(/\s+/g,' ').trim(),
                         onay:(tr.children[12]?.textContent||'').replace(/\s+/g,' ').trim(),
                         onayId:(tr.children[13]?.textContent||'').trim(),
                         total:(tr.children[15]?.textContent||'').trim()}));}""")
    out = []
    for r in raw:
        m = _NAME_RE.match(r["pers"])
        ad = (m.group(1) if m else r["pers"]).strip()
        out.append({"id": r["id"], "pers_raw": r["pers"], "ad": ad, "ad_fold": fold(ad),
                    "onay": r["onay"], "onayId": r["onayId"], "total": r["total"]})
    return out


def dblclick_row(page, ad_fold: str) -> bool:
    """İsme göre flexigrid GÖVDE satırını çift-tıkla (Personel=td[6], TC'den arındırılmış fold eşleşmesi)."""
    return page.evaluate(
        r"""(args)=>{const [target] = args;
            const fold=s=>{const tr={'ç':'c','Ç':'c','ğ':'g','Ğ':'g','ı':'i','İ':'i','I':'i','ö':'o','Ö':'o','ş':'s','Ş':'s','ü':'u','Ü':'u'};
              return s.replace(/̇/g,'').split('').map(c=>tr[c]||c).join('').toLowerCase().replace(/\s+/g,' ').trim();};
            const dlg=[...document.querySelectorAll('div.ui-dialog')].filter(d=>d.offsetParent!==null && /Dışarıda Geçirilen Süreler Listesi/.test(d.querySelector('.ui-dialog-titlebar')?.textContent||''))[0];
            const body=dlg?dlg.querySelector('.flexigrid .bDiv'):null; if(!body)return false;
            for(const tr of body.querySelectorAll('tr')){
              if(tr.children.length<13)continue;
              const nm=fold((tr.children[6]?.textContent||'').replace(/\s+[\d\*]{3,}.*$/,''));
              if(nm===target){ tr.scrollIntoView({block:'center'});
                for(const t of ['mouseover','mousedown','mouseup','click','dblclick'])
                  tr.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:window}));
                return true; }}
            return false;}""", [ad_fold])


def wait_form_open(page, timeout_ms: int = 7000) -> bool:
    """Form penceresi + 'Onaya Gönder' butonu görünür olana dek bekle (~3sn açılıyor)."""
    try:
        page.wait_for_function(
            r"""()=>{const bar=[...document.querySelectorAll('div.ui-dialog-titlebar')].some(b=>b.offsetParent!==null && /Dışarıda Geçirilen Süreler Formu/.test(b.textContent));
                const btn=[...document.querySelectorAll('a,button')].some(x=>x.offsetParent!==null && /Onaya Gönder/i.test(x.textContent));
                return bar && btn;}""", timeout=timeout_ms)
        page.wait_for_timeout(600)
        return True
    except PWTimeout:
        return False


def form_personel_matches(page, ad_fold: str) -> bool:
    val = page.evaluate(
        r"""()=>{const e=[...document.querySelectorAll('input[name="string_Personel_Id"]')].filter(x=>x.offsetParent!==null).pop();
            return e?e.value:'';}""")
    return ad_fold in fold(val or "")


def form_istenilen(page) -> str:
    return page.evaluate(
        r"""()=>{const c=document.querySelector('.BuFordaIstenilen'); const i=c?c.querySelector('input'):null;
            return i?i.value.trim():'?';}""")


def onaya_gonder_clickable(page) -> bool:
    return page.evaluate(
        r"""()=>{const b=[...document.querySelectorAll('a,button')].find(x=>x.offsetParent!==null && /Onaya Gönder/i.test(x.textContent));
            if(!b)return false;
            const dis=b.disabled || /disabled/.test(b.className) || b.getAttribute('aria-disabled')==='true';
            return !dis;}""")


def form_status(page) -> str:
    """Form'daki 'Onay Durumu' select'inin metni (doğrulama anchor'ı)."""
    return page.evaluate(
        r"""()=>{const dlg=[...document.querySelectorAll('div.ui-dialog')].filter(d=>d.offsetParent!==null && /Dışarıda Geçirilen Süreler Formu/.test(d.querySelector('.ui-dialog-titlebar')?.textContent||''))[0];
            if(!dlg)return ''; const sels=[...dlg.querySelectorAll('select')].map(s=>(s.selectedOptions[0]?.text||'').trim())
              .filter(t=>/Gönderil|Onay|Beklen|Reddedil|Değerlendir/.test(t)); return sels.join(' | ');}""")


def click_onaya_gonder(page) -> None:
    """Onaya Gönder + onay istemi — GERÇEK Playwright click (evaluate-click jQuery handler'ı
    GÜVENİLMEZ tetikliyor). Onay istemi PORTAL-BAĞIMLI: BV/TPIz Bootstrap modal (#modalYes /
    .modal button.color-green), TPI SweetAlert/native. (2026-06-18 BV izin onayıyla aynı tuzak.)"""
    page.locator("a:has-text('Onaya Gönder'), button:has-text('Onaya Gönder')").filter(visible=True).first.click(timeout=5000)
    D.wait_for_form_idle(page)   # Onaya Gönder → onay istemi yüklenirken bekle (loading-fix)
    # onay istemi "Evet" — sırayla GERÇEK locator-click (modal yoksa sessizce geç; form_status doğrular).
    # native confirm() ise zaten page.on("dialog", accept) yakalar → bu döngü boşa düşer, sorun değil.
    for sel in ("#modalYes", ".modal.in button.color-green", ".modal.show button.color-green",
                ".modal.in button:has-text('Evet')", ".sweet-alert button.confirm", ".sweet-alert a.confirm"):
        try:
            page.locator(sel).filter(visible=True).first.click(timeout=2000)
            break
        except Exception:
            continue
    else:
        for nm in ("Evet", "Eminim", "Onayla"):
            try:
                page.get_by_role("button", name=nm, exact=True).filter(visible=True).first.click(timeout=1200)
                break
            except Exception:
                continue
    D.wait_for_form_idle(page)   # onay AJAX'ı bitene kadar bekle


def main():
    ap = argparse.ArgumentParser(description="DGS taslaklarını ONAYA GÖNDER (sadece dgs_done'dakiler)")
    ap.add_argument("--excel", required=True)
    ap.add_argument("--sheet", default="Mayıs")
    ap.add_argument("--lokasyon", default="TPI")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--commit", action="store_true", help="GERÇEK Onaya Gönder (yoksa DRY-RUN)")
    ap.add_argument("--dry-run", action="store_true", help="(varsayılan) sadece önizleme; gönderme yok")
    ap.add_argument("--donem", default=None, help="Dönem etiketi override (ör. 'MAYIS 2026'); yoksa oto önceki ay")
    args = ap.parse_args()
    dry = not args.commit

    done_file = f"dgs_done_{args.lokasyon}_{args.sheet}.txt"
    if not os.path.exists(done_file):
        log(f"HATA: {done_file} yok."); sys.exit(1)
    our = set(fold(l.strip()) for l in open(done_file, encoding="utf-8") if l.strip())
    onay_file = f"dgs_onaya_done_{args.lokasyon}_{args.sheet}.txt"
    already = set(fold(l) for l in open(onay_file, encoding="utf-8").read().splitlines()) if os.path.exists(onay_file) else set()
    label = args.donem or prev_month_label()
    limit = args.limit if args.limit else 10**9
    log(f"Dönem: {label} | done={len(our)} | zaten-onaya-gönderilmiş={len(already)} | "
        f"MOD={'DRY-RUN (gönderme YOK)' if dry else 'COMMIT (GERÇEK)'}")

    with sync_playwright() as pw:
        _browser, page = D.attach_browser(pw)
        page.on("dialog", lambda d: d.accept())   # native confirm() çıkarsa kabul et (submit'i iptal etme)
        D.assert_logged_in(page)

        # ---- DRY-RUN: tüm statülerle listele, done setiyle eşleştir, statü dağılımı raporla ----
        if dry:
            open_list(page)
            set_donem(page, label)
            set_status_filter(page, only_draft=False)
            listele(page)
            rows = read_list_rows(page)
            log(f"Listede bu sayfada {len(rows)} kayıt | {footer_total(page)} (dönem={label}, tüm statüler).")
            ours = [r for r in rows if r["ad_fold"] in our]
            yabanci = [r for r in rows if r["ad_fold"] not in our]
            from collections import Counter
            dist = Counter(r["onay"] for r in ours)
            log(f"Bizim done'dan eşleşen: {len(ours)} | listede bizim-olmayan: {len(yabanci)}")
            log("Eşleşenlerin statü dağılımı:")
            for st, c in dist.most_common():
                log(f"    {c:>3}  {st}")
            eksik = sorted(n for n in our if n not in {r['ad_fold'] for r in rows})
            if eksik:
                log(f"Listede HİÇ görünmeyen done-kişi: {len(eksik)} (ilk 10: {eksik[:10]})")
            log("Örnek 5 eşleşen:")
            for r in ours[:5]:
                log(f"    {r['pers_raw'][:32]:<32} | {r['onay']}")
            log(">> DRY-RUN bitti. Göndermek için --commit ekle.")
            try: input(f"{LOG} ENTER ile kapat...")
            except EOFError: pass
            return

        # ---- COMMIT: taslak (Id_5) filtrele → ilk bizim-kişiyi onaya gönder → tazele → tekrar ----
        ok = 0
        failed = set()
        open_list(page)
        for _iter in range(len(our) + 40):
            if ok >= limit:
                break
            # HER TUR baştan: dönem + draft filtre + listele (form aç/kapa filtreyi sıfırlıyor — CANLI BULGU)
            if not page.evaluate(r"""()=>[...document.querySelectorAll('div.ui-dialog-titlebar')].some(b=>b.offsetParent!==null && /Dışarıda Geçirilen Süreler Listesi/.test(b.textContent))"""):
                open_list(page)
            ensure_filter_open(page)
            set_donem(page, label)
            set_status_filter(page, only_draft=True)   # yalnız "Değerlendirmeye Gönderilmemiş"
            listele(page)
            rows = read_list_rows(page)
            # işlenecek ilk satır: bizim setten, daha önce onaya gitmemiş, bu koşuda failed olmamış
            target = next((r for r in rows if r["ad_fold"] in our
                           and r["ad_fold"] not in already and r["ad_fold"] not in failed), None)
            if not target:
                yabanci = [r for r in rows if r["ad_fold"] not in our]
                log(f"Bizim gönderilecek taslak kalmadı. (listede bizim-olmayan taslak: {len(yabanci)})")
                break
            ad = target["ad"]
            # GÜVENLİK 1: liste satırının Toplam'ı (td[15]) dolu mu — BOŞ kayıt onaya gitmesin
            if not target["total"] or target["total"] in ("0:00", "00:00"):
                failed.add(target["ad_fold"]); log(f"!! {ad}: liste toplamı boş/0 ({target['total']!r}) — atlandı"); continue
            if not dblclick_row(page, target["ad_fold"]):
                failed.add(target["ad_fold"]); log(f"!! {ad}: satır çift-tıklanamadı, atlandı"); continue
            if not wait_form_open(page):
                failed.add(target["ad_fold"]); log(f"!! {ad}: form ~7sn'de açılmadı, atlandı"); close_form_only(page); continue
            D.wait_for_form_idle(page)   # form içeriği tam yüklensin (status/personel okumadan önce) — loading-fix
            # GÜVENLİK 2: form gerçekten bu kişinin mi?
            if not form_personel_matches(page, target["ad_fold"]):
                failed.add(target["ad_fold"]); log(f"!! {ad}: form personeli eşleşmedi (GÜVENLİK), atlandı"); close_form_only(page); continue
            # GÜVENLİK 3: form statüsü gerçekten taslak mı?
            st_before = form_status(page)
            if "Gönderilmemiş" not in st_before:
                failed.add(target["ad_fold"]); log(f"!! {ad}: form statüsü taslak değil ({st_before!r}) — atlandı"); close_form_only(page); continue
            # 'Onaya Gönder' tıklanabilir mi? (kullanıcı kuralı: değilse sıradakine geç)
            if not onaya_gonder_clickable(page):
                failed.add(target["ad_fold"]); log(f"!! {ad}: 'Onaya Gönder' tıklanamıyor (pasif), atlandı"); close_form_only(page); continue
            click_onaya_gonder(page)
            # DOĞRULAMA: form statüsü taslaktan çıkana kadar ~6sn poll (async güncelleme → false-negative önle)
            st_after = ""
            for _ in range(6):
                st_after = form_status(page)
                if st_after and "Gönderilmemiş" not in st_after:
                    break
                page.wait_for_timeout(1000)
            close_form_only(page)
            if not st_after or "Gönderilmemiş" in st_after:
                failed.add(target["ad_fold"]); log(f"!! {ad}: onaya GİTMEDİ (statü={st_after!r}), atlandı")
                continue
            ok += 1
            with open(onay_file, "a", encoding="utf-8") as f:
                f.write(ad + "\n")
            already.add(target["ad_fold"])
            log(f"[{ok}] {ad} → GÖNDERİLDİ ✓ (Toplam={target['total']}, statü={st_after})")

        log(f"\n==== ÖZET: bu koşuda {ok} kişi onaya gönderildi ====")
        if failed:
            log(f"  !! GÖNDERİLEMEYEN/atlanan ({len(failed)}): elle bak.")
        log("HATIRLATMA: Bu portalda Onaya Gönder kaydı doğrudan 'Onaylanmış'a geçiriyor (e-imza çıkmadı). "
            "Yine de SGK raporundan teyit önerilir.")
        try: input(f"{LOG} >> Bitti. ENTER ile kapat...")
        except EOFError: pass


if __name__ == "__main__":
    main()
