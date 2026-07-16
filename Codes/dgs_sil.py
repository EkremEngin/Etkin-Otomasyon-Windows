#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DGS kaydı SİL — hatalı/eksik (onaya gönderilmiş) DGS kayıtlarını siler; sonra yeniden girilir.
=============================================================================================
Akış (kullanıcı canlı gösterdi 2026-06-18):
  PERSONEL → "Dışarıda Geçirilen Süreler Listesi" → Dönem seç → personel FİLTRE (kimlik doğrulamalı)
  → Listele → kişinin kaydına ÇİFT TIKLA → form → "Sil (F4)" → onay → kayıt listeden DÜŞTÜ doğrula.

GÜVENLİK (verify-or-halt):
  - Yalnız --person / --file ile verilen isimler işlenir.
  - Form açıldıktan sonra personel KİMLİĞİ doğrulanır (yanlış kişiyi ASLA silmez).
  - Silme sonrası kaydın gerçekten gittiği DOĞRULANIR (gitmezse 'GİTMEDİ' der, atlar).
  - VARSAYILAN DRY-RUN: kaydı bulur, kimliği doğrular, Sil butonunun durumunu raporlar — SİLMEZ.
  - --commit ile GERÇEK siler. GERİ ALINMASI ZOR — kullanıcı istedi (hatalı kayıtları yeniden gireceğiz).

ANİMASYON/LOADING: HER pop-up/işlem sınırında D.wait_for_form_idle (jQuery.active>0 + yükleme paneli +
  settle) çağrılır — portal pop-up açarken 2-3 sn loading dönüyor; beklemeden müdahale 'tutmadı' sanılıyor.

Kullanım:
  python3 dgs_sil.py --person "AD SOYAD"            # dry-run (göster, silmez)
  python3 dgs_sil.py --person "AD SOYAD" --commit   # GERÇEK sil
  python3 dgs_sil.py --file silinecekler.txt --commit
"""
from __future__ import annotations
import argparse
import os
import sys

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

import dgs_poc as D
import dgs_onaya as O

LOG = "[SİL]"
def log(*a): print(LOG, *a, flush=True)


def filter_to_person(page, name: str) -> bool:
    """Listesi personel filtresine kişiyi yaz — TAM-AD fold BİREBİR eşleşme (yanlış kişi seçme).
    Her denemede loading beklenir."""
    intended = D._fold_tr(name)
    parts = name.split()
    cands = [name]
    if len(parts) >= 3: cands.append(" ".join(parts[:2]))
    cands.append(parts[0])
    if len(parts) > 1 and parts[-1] not in cands: cands.append(parts[-1])
    for q in cands:
        page.evaluate(
            r"""({txt})=>{const el=[...document.querySelectorAll('input[name="string_Personel_Id"]')].filter(e=>e.offsetParent!==null).pop();
                if(el){el.focus();el.value=txt; if(window.jQuery){try{jQuery(el).autocomplete('search',txt);}catch(e){}}
                  el.dispatchEvent(new Event('input',{bubbles:true})); el.dispatchEvent(new Event('keyup',{bubbles:true}));}}""",
            {"txt": q})
        page.wait_for_timeout(1300)   # autocomplete önerisi
        picked = page.evaluate(
            r"""(args)=>{const [target]=args; %s
                const lis=[...document.querySelectorAll('ul.ui-autocomplete li, ul.ui-menu li')].filter(li=>li.offsetParent!==null);
                for(const li of lis){ const nm=_fold((li.textContent||'').replace(/\s+[\d\*]{3,}.*$/,''));
                  if(nm===target){ const t=li.querySelector('a,div')||li;
                    for(const e of ['mouseover','mouseenter','mousemove','mousedown','mouseup','click'])
                      t.dispatchEvent(new MouseEvent(e,{bubbles:true,cancelable:true,view:window}));
                    return true; } }
                return false;}""" % D._PERS_FOLD_JS, [intended])
        if picked:
            D.wait_for_form_idle(page)   # filtre seçimi sonrası AJAX
            return True
    return False


def form_sil_button(page):
    """Form dialogundaki 'Sil (F4)' butonu: {var, enabled} döndürür (tıklamadan)."""
    return page.evaluate(
        r"""()=>{const dlg=[...document.querySelectorAll('div.ui-dialog')].filter(d=>d.offsetParent!==null && /Dışarıda Geçirilen Süreler Formu/.test(d.querySelector('.ui-dialog-titlebar')?.textContent||''))[0];
            if(!dlg) return {var:false};
            const b=[...dlg.querySelectorAll('a,button')].find(x=>x.offsetParent!==null && /Sil/.test(x.textContent) && /F4/.test(x.textContent));
            if(!b) return {var:false};
            const dis=b.disabled || /disabled/.test(b.className) || b.getAttribute('aria-disabled')==='true';
            return {var:true, enabled:!dis};}""")


def click_sil(page):
    """'Sil (F4)' GERÇEK click + onay. Her adımda loading beklenir. (native confirm → dialog handler accept)."""
    page.locator("a:has-text('Sil (F4)'), button:has-text('Sil (F4)')").filter(visible=True).first.click(timeout=5000)
    D.wait_for_form_idle(page)   # Sil → onay popup'ı / silme AJAX'ı yüklenirken bekle
    # onay kutusu (DOM): Evet/Eminim/Tamam/Sil/Onayla  (native confirm ise dialog handler zaten accept eder)
    for _ in range(3):
        clicked = page.evaluate(
            r"""()=>{const b=[...document.querySelectorAll('a,button,.sweet-alert .confirm,.swal2-confirm')].find(x=>x.offsetParent!==null && /^(Evet|Eminim|Tamam|Onayla|Sil)$/i.test((x.textContent||'').trim())); if(b){b.click(); return true;} return false;}""")
        D.wait_for_form_idle(page)
        if clicked:
            break
        page.wait_for_timeout(400)


def process_one(page, name: str, dry: bool) -> str:
    """Tek kişinin DGS kaydını sil. Dönüş: 'silindi' / 'dry' / 'yok' / 'hata: ...'."""
    # 1) liste + dönem + personel filtre + TÜM statüler + Listele  (hepsi loading-beklemeli)
    O.open_list(page); D.wait_for_form_idle(page)
    O.ensure_filter_open(page)
    O.set_donem(page, _donem_label())
    if not filter_to_person(page, name):
        return f"hata: personel filtresine '{name}' yazılamadı (autocomplete eşleşmedi)"
    O.set_status_filter(page, only_draft=False)   # TÜM statüler (kayıt Onaylanmış olabilir)
    O.listele(page); D.wait_for_form_idle(page)
    rows = [r for r in O.read_list_rows(page) if r["ad_fold"] == D._fold_tr(name)]
    if not rows:
        return "yok (listede bu kişinin kaydı yok — zaten silinmiş/girilmemiş olabilir)"
    if len(rows) > 1:
        log(f"  NOT: {name} için {len(rows)} kayıt var; ilki işlenecek (statü={rows[0]['onay']}).")
    target = rows[0]
    log(f"  bulundu: {target['pers_raw'][:30]} | statü={target['onay']} | toplam={target['total']}")
    # 2) kayda çift-tıkla → form aç (loading beklemeli)
    if not O.dblclick_row(page, target["ad_fold"]):
        return "hata: kayıt satırı çift-tıklanamadı"
    D.wait_for_form_idle(page)
    if not O.wait_form_open(page):
        return "hata: form ~7sn'de açılmadı"
    D.wait_for_form_idle(page)
    # 3) GÜVENLİK: form gerçekten bu kişinin mi?
    if not O.form_personel_matches(page, target["ad_fold"]):
        O.close_form_only(page)
        return "hata: form personeli EŞLEŞMEDİ (GÜVENLİK — silmedim)"
    # 4) Sil butonu durumu
    sb = form_sil_button(page)
    if not sb.get("var"):
        O.close_form_only(page)
        return "hata: 'Sil (F4)' butonu formda yok"
    if not sb.get("enabled"):
        O.close_form_only(page)
        return "hata: 'Sil (F4)' pasif (silinemez)"
    if dry:
        O.close_form_only(page)
        return "dry (kayıt + Sil butonu HAZIR; --commit ile silinir)"
    # 5) GERÇEK SİL + onay (loading-beklemeli)
    click_sil(page)
    # 6) DOĞRULA: kayıt listeden düştü mü? (personel filtre + tüm statü + Listele yeniden)
    O.open_list(page); D.wait_for_form_idle(page)
    O.set_donem(page, _donem_label()); O.ensure_filter_open(page)
    if not filter_to_person(page, name):
        return "silindi? (doğrulama filtresi yazılamadı — elle teyit et)"
    O.set_status_filter(page, only_draft=False)
    O.listele(page); D.wait_for_form_idle(page)
    still = [r for r in O.read_list_rows(page) if r["ad_fold"] == D._fold_tr(name)]
    if still:
        return f"GİTMEDİ (hâlâ {len(still)} kayıt var, statü={still[0]['onay']})"
    return "silindi ✓"


def _donem_label():
    return O.prev_month_label()


def main():
    ap = argparse.ArgumentParser(description="DGS kaydı SİL (yalnız verilen isimler; verify-or-halt; loading-beklemeli)")
    ap.add_argument("--person", default=None)
    ap.add_argument("--file", default=None, help="Her satırda bir ad-soyad")
    ap.add_argument("--commit", action="store_true", help="GERÇEK sil (yoksa DRY-RUN)")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()
    dry = not args.commit

    names = []
    if args.person: names = [args.person.strip()]
    elif args.file:
        names = [l.strip() for l in open(args.file, encoding="utf-8") if l.strip()]
    else:
        log("HATA: --person veya --file ver."); sys.exit(1)

    log(f"{len(names)} kişi | dönem={_donem_label()} | MOD={'DRY-RUN (silmez)' if dry else 'COMMIT (GERÇEK SİLER)'}")
    with sync_playwright() as pw:
        _b, page = D.attach_browser(pw)
        page.on("dialog", lambda d: d.accept())   # native confirm → kabul (silme onayı)
        D.assert_logged_in(page)
        ok = 0; fail = []
        for i, nm in enumerate(names, 1):
            log(f"[{i}/{len(names)}] {nm}")
            try:
                res = process_one(page, nm, dry)
            except Exception as e:
                res = f"hata: {str(e)[:100]}"
            log(f"    -> {res}")
            if res.startswith("silindi") or res.startswith("dry"):
                ok += 1
            else:
                fail.append((nm, res))
        log(f"\n==== ÖZET: {ok}/{len(names)} {'hazır (dry)' if dry else 'silindi'} ====")
        if fail:
            log(f"  !! İşlenemeyen ({len(fail)}):")
            for nm, r in fail: log(f"     {nm}: {r}")
        try: input(f"{LOG} >> Bitti. ENTER ile kapat...")
        except EOFError: pass


if __name__ == "__main__":
    main()
