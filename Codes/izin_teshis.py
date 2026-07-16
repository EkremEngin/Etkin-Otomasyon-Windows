#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
İZİN TEŞHİS — personel-seçimi başarısızlıklarını INSTRUMENTED olarak yeniden üret + NE OLDUĞUNU yakala.
=====================================================================================================
Neden: bulk'ta bazı kişiler "autocomplete önerisi çıkmadı" / "Kaydet mesajı yok" ile düşüyor (2026-07-08 TPI:
8 kişi). Bu sistemsel mi (server o kişiyi döndürmüyor / widget bağlanmıyor) yoksa timing/fatigue mi belli
değildi — sadece hata mesajı + screenshot vardı. Bu araç, motorun (izin_poc) form-açma + autocomplete'ini
AYNEN kullanır ama başarısızlık anında ZENGİN teşhis toplar:

  • autocomplete AJAX'ının SERVER YANITI (personel arama XHR gövdesi) — server o kişiyi döndürüyor mu?
  • dropdown DOM'u (ul.ui-autocomplete li) — hiç öğe var mı, eşleşen var mı?
  • jQuery UI autocomplete WIDGET bağlı mı (bağlı değilse form JS init olmamış = fatigue/timing güçlü aday)
  • console hataları + takılı loading overlay
  • kısa-ad (soyad) VE tam-ad aramasını ayrı dener (hangi terim tutuyor?)

Her kişi için bir VERDICT üretir (sistemsel/timing ayrımı) + `izin_teshis_<park>_<ay>.jsonl` log yazar
(makine-okunur; sonradan analiz/örüntü için) + okunur konsol özeti. Sonda örüntü özeti (N/N aynı mı?).

GÜVENLİK: DRY — Kaydet (F3) ASLA tıklanmaz (dry_run=True). Form açar + arar + okur; kayıt GİRMEZ/GÖNDERMEZ.
  Açık sekmeyi kullanır (izin_otomasyon gibi); park açık portaldan algılanır (--lokasyon ile zorlanabilir).
  izin_poc/izin_otomasyon'a DOKUNMAZ — sadece import edip reuse eder.

Kullanım:
  DGS_CDP=http://localhost:9222 python3 izin_teshis.py --excel "<XL>"                 # done'da olmayan (başarısız) herkes
  ...  --person "Beyhannur Çeviral"     # tek kişi
  ...  --lokasyon TPI  --limit 3        # ilk 3 başarısız
  Harness/arka planda:  < /dev/null  ekle (bitişteki input() asılmasın).
"""
from __future__ import annotations
import argparse
import datetime
import json
import os
import sys

from playwright.sync_api import sync_playwright

import izin_poc
import izin_data_v2 as data
from izin_data_v2 import read_izin_v2, PARKS, resolve_park, fold
from izin_otomasyon import park_from_url, _domain, configure_engine


def log(*a):
    print("[TEŞHİS]", *a, flush=True)


# Arama sonrası form/dropdown/widget durumunu tek seferde dök (izin_poc._FIND_FORM ile form-dialog bulunur)
_CAPTURE_JS = (r"""() => {
  %s
  const el = _dlg ? _dlg.querySelector('input[name="string_Personel_Id"]') : null;
  const lis = [...document.querySelectorAll('ul.ui-autocomplete li, ul.ui-menu li')];
  const visible = lis.filter(x=>x.offsetParent!==null).map(x=>x.textContent.trim());
  const all = lis.map(x=>x.textContent.trim());
  const overlay = [...document.querySelectorAll('.form_loading_panel, .pReload.loading, .loading, .blockUI')]
                    .some(e=>e.offsetParent!==null);
  const widget = !!(el && window.jQuery && jQuery(el).data('ui-autocomplete'));
  const ac_ul = document.querySelector('ul.ui-autocomplete');
  return JSON.stringify({
    input_found: !!el,
    input_value: el ? el.value : null,
    widget_bound: widget,
    dropdown_present: !!(ac_ul && ac_ul.offsetParent!==null),
    visible_items: visible,
    all_li: all,
    loading_overlay: overlay
  });
}""") % izin_poc._FIND_FORM


def summarize_net(responses):
    """XHR/fetch yanıtlarını özetle; personel/autocomplete gibi görünenlerin GÖVDESİNİ de oku (server ne döndü?)."""
    out = []
    for r in responses:
        try:
            rt = r.request.resource_type
        except Exception:
            rt = "?"
        if rt not in ("xhr", "fetch"):
            continue
        url = r.url
        try:
            status = r.status
        except Exception:
            status = None
        body = ""
        if any(k in url.lower() for k in ("personel", "autocomplete", "search", "arama")):
            try:
                body = r.text()[:1500]
            except Exception:
                body = "(gövde okunamadı)"
        out.append({"url": url, "status": status, "body": body})
    return out


def verdict(cap, net_sum):
    """Yakalanan sinyallerden sistemsel/timing ayrımı için okunur teşhis."""
    if not cap.get("input_found"):
        return "🔴 FORM AÇILMADI (personel input yok) → form-açma/menü sorunu"
    if not cap.get("widget_bound"):
        return "🟠 AUTOCOMPLETE WIDGET BAĞLANMAMIŞ → form JS init olmadı (fatigue/timing GÜÇLÜ aday)"
    xhr = net_sum
    if not xhr:
        return "🟠 AJAX ATILMADI → widget var ama arama server'a gitmedi (tetikleme/timing)"
    bodied = [n for n in xhr if n["body"]]
    empty = [n for n in bodied if n["body"].strip() in ("[]", "{}", "")]
    if bodied and empty and not cap.get("all_li"):
        return "🔴 SERVER BOŞ DÖNDÜ ([]/{}) → o terimde personel bulunamadı (SİSTEMSEL: isim/kodlama/kapsam)"
    if cap.get("all_li") and not cap.get("visible_items"):
        return "🟠 SERVER DÖNDÜ ama dropdown'da EŞLEŞEN token yok → isim-eşleştirme/görünürlük"
    if not cap.get("all_li"):
        return "🟠 DROPDOWN BOŞ (server yanıtı belirsiz) → render/timing; AJAX gövdesine bak"
    return "🟢 DROPDOWN DOLDU → bu denemede geçti (INTERMITTENT = timing/fatigue kanıtı)"


def _purge_dialogs(page):
    """open_izin_form ÖNCESİ temiz-slate: bekleyen LİSTE/dialog/overlay'i kapat. Checker/probe (veya önceki kişi)
    İzin LİSTE dialog'unu açık bırakmış olabilir → form o bayat pencere üstüne biner, temiz açılmaz (yanlış 'form yok')."""
    try:
        page.evaluate(r"""()=>{
          document.querySelectorAll('.sweet-overlay,.sweet-alert,.ui-widget-overlay,.modal-backdrop').forEach(e=>{try{e.remove()}catch(_){}});
          [...document.querySelectorAll('div.ui-dialog-titlebar')].filter(b=>b.offsetParent!==null)
            .forEach(b=>{const x=b.querySelector('.fa-times,.ui-dialog-titlebar-close'); if(x)x.click();});
        }""")
        page.wait_for_timeout(700)
    except Exception:
        pass


def run(page, targets, park, meta, jsonl_path):
    net, console_msgs = [], []
    page.on("response", lambda r: net.append(r))
    page.on("console", lambda m: console_msgs.append(f"{m.type}: {m.text}"[:300]))

    results = []
    for i, p in enumerate(targets, 1):
        surname = p.ad.split()[-1]
        entry = {"ad": p.ad, "tc": p.tc, "surname": surname,
                 "ts": datetime.datetime.now().isoformat(timespec="seconds")}
        log(f"\n[{i}/{len(targets)}] {p.ad}  (soyad='{surname}')")
        try:
            _purge_dialogs(page)           # temiz-slate: bayat liste/dialog kapat (form temiz açılsın)
            izin_poc.open_izin_form(page)
        except Exception as e:
            entry["form_error"] = str(e)[:200]
            entry["verdict"] = f"🔴 FORM AÇILAMADI: {str(e)[:120]}"
            _emit(entry, jsonl_path); results.append(entry); continue

        searches = {}
        for tag, term in (("soyad", surname), ("tam_ad", p.ad)):
            net.clear(); console_msgs.clear()
            try:
                izin_poc._autocomplete_search(page, term)
            except Exception as e:
                searches[tag] = {"error": str(e)[:150]}
                continue
            page.wait_for_timeout(3500)     # AJAX + dropdown render için bekle (acele etme)
            try:
                cap = json.loads(page.evaluate(_CAPTURE_JS))
            except Exception as e:
                cap = {"capture_error": str(e)[:150]}
            searches[tag] = {"capture": cap, "network": summarize_net(list(net)),
                             "console": [c for c in console_msgs if "error" in c.lower() or "warn" in c.lower()][:8]}

        shot = f"izin_teshis_{fold(p.ad).replace(' ', '_')}.png"
        try:
            page.screenshot(path=shot, full_page=True)
        except Exception:
            shot = ""
        entry["searches"] = searches
        entry["screenshot"] = shot
        cap_soyad = searches.get("soyad", {}).get("capture", {})
        net_soyad = searches.get("soyad", {}).get("network", [])
        entry["verdict"] = verdict(cap_soyad, net_soyad)
        _emit(entry, jsonl_path); results.append(entry)
    return results


def _emit(entry, jsonl_path):
    """JSONL log'a yaz + okunur konsol özeti."""
    with open(jsonl_path, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")
    print(f"    VERDICT: {entry.get('verdict')}")
    for tag in ("soyad", "tam_ad"):
        s = entry.get("searches", {}).get(tag)
        if not s:
            continue
        if "error" in s:
            print(f"      {tag:6s}: HATA {s['error']}"); continue
        cap = s.get("capture", {})
        net = s.get("network", [])
        xhr_info = "; ".join(f"{n['status']} {os.path.basename(n['url'].split('?')[0])}"
                             + (f" body={n['body'][:60]!r}" if n['body'] else "") for n in net[:3]) or "XHR yok"
        print(f"      {tag:6s}: widget={cap.get('widget_bound')} dropdown={len(cap.get('all_li',[]))}öğe "
              f"eşleşen={len(cap.get('visible_items',[]))} overlay={cap.get('loading_overlay')} | {xhr_info}")
        if s.get("console"):
            print(f"              console: {s['console'][:2]}")


def main():
    ap = argparse.ArgumentParser(description="İzin personel-seçimi başarısızlıklarını instrumented teşhis (dry, salt-gözlem)")
    ap.add_argument("--excel", required=True)
    ap.add_argument("--lokasyon", default=None, help="Parkı zorla; yoksa açık sekmeden algılanır")
    ap.add_argument("--person", default=None, help="Tek kişiyi teşhis et (tam ad)")
    ap.add_argument("--limit", type=int, default=0, help="En çok kaç kişi (0=hepsi)")
    ap.add_argument("--cdp", default=os.environ.get("DGS_CDP", "http://localhost:9222"))
    args = ap.parse_args()

    try:
        by_park, meta = read_izin_v2(os.path.expanduser(args.excel), strict=True)
    except data.DataError as e:
        print(f"\n❌ VERİ HATASI:\n{e}\n", file=sys.stderr); sys.exit(2)

    forced = resolve_park(args.lokasyon) if args.lokasyon else None
    if args.lokasyon and forced is None:
        log(f"HATA: bilinmeyen park {args.lokasyon!r}"); sys.exit(1)

    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp(args.cdp)
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        page = next((p for p in ctx.pages if "argeportal" in (p.url or "")), None)
        if page is None:
            log("HATA: açık 'argeportal' sekmesi yok."); sys.exit(1)
        active = park_from_url(page.url)
        if active is None:
            log(f"HATA: açık sekme ({_domain(page.url)}) bilinen portal değil."); sys.exit(1)
        if forced is not None and forced.code != active.code:
            log(f"🔴 GÜVENLİK: --lokasyon {forced.code} ama açık portal {active.code} → DURDU."); sys.exit(3)

        configure_engine(izin_poc, active, meta, commit=False, cdp=args.cdp)   # DRY (kaydetmez)

        people = by_park.get(active.code, [])
        done_file = f"izin_done_{active.label}_{meta['ay_key']}.txt"
        done = set()
        if os.path.exists(done_file):
            done = set(fold(l) for l in open(done_file, encoding="utf-8").read().splitlines() if l.strip())
        if args.person:
            pf = fold(args.person)
            targets = [p for p in people if fold(p.ad) == pf]
            if not targets:
                log(f"HATA: '{args.person}' {active.code} listesinde yok."); sys.exit(1)
        else:
            targets = [p for p in people if fold(p.ad) not in done]   # done'da olmayan = başarısız/girilmemiş
        if args.limit:
            targets = targets[:args.limit]
        if not targets:
            log(f"{active.code}: teşhis edilecek (done-dışı) kişi yok — hepsi girilmiş görünüyor."); sys.exit(0)

        jsonl = f"izin_teshis_{active.label}_{meta['ay_key']}.jsonl"
        log(f"Park {active.code} | teşhis edilecek {len(targets)} kişi (done-dışı) | dönem {meta['donem_label']} | DRY")
        log(f"Log: {jsonl}")
        results = run(page, targets, active, meta, jsonl)

        # ÖRÜNTÜ ÖZETİ
        print("\n" + "=" * 84)
        print(f"TEŞHİS ÖZETİ — {active.code}: {len(results)} kişi")
        print("=" * 84)
        from collections import Counter
        vc = Counter(r.get("verdict", "?").split("→")[0].strip() for r in results)
        for v, n in vc.most_common():
            print(f"  {n:2d}×  {v}")
        widgets = [r["searches"]["soyad"]["capture"].get("widget_bound")
                   for r in results if r.get("searches", {}).get("soyad", {}).get("capture")]
        if widgets and all(w is False for w in widgets):
            print("\n  🟠 ÖRÜNTÜ: HERKESTE widget bağlı DEĞİL → form JS init/timing sistemsel (fatigue çok olası).")
        elif widgets and all(w is True for w in widgets):
            print("\n  🔎 ÖRÜNTÜ: widget herkeste bağlı → sorun AJAX/server yanıtında; gövdelere bak (jsonl).")
        print(f"\n  Ayrıntı (makine-okunur): {jsonl}")
        print("=" * 84)

    try:
        input("\n>> Teşhis bitti (DRY — hiçbir şey kaydedilmedi). ENTER ile kapat...")
    except EOFError:
        pass


if __name__ == "__main__":
    main()
