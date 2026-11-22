#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
✅ KAYDEDİLDİ — "Yıldız / Çoklu İzin Yükleme" scripti. CANLI KANITLI (2026-06-19): 23/23 onaylandı, 0 hata.
   Donmuş yedek: prime/izin_onaya_dosyali_yildiz.py (PRIME_SCRIPTS.md). Bozulursa oradan geri yükle.
🌟 YILDIZ varyantı — İzin TASLAK'larını ONAYA GÖNDER + her kişiye KENDİ belge(ler)ini yükle.
=====================================================================================
Türetildiği taban: izin_onaya_dosyali.py (canlı doğrulanmış blueimp dosya-yükleme akışı). Tek FARK:
tek global `--dosya` yerine, her kişinin belge(ler)i `--dosya-klasor`'dan DGS ADIYLA çözülür ve
kişide kaç belge varsa HEPSİ sırayla yüklenir (örn. SENANUR IŞIK → 3 dosya = 3 upload). Geri kalan
her şey (liste, çift-tık, Onaya Gönder, Evet, doğrulama, resume) tabandakiyle BİREBİR aynıdır.

Belge eşleştirme: klasördeki dosyalar "AD SOYAD.pdf" veya "AD SOYAD (n).pdf" formatında (bu repo'da
DGS adlarına göre yeniden adlandırıldı). Kişinin portal-adı fold() ile dosya adına eşlenir; (n) sırasına
göre 1→2→3 yüklenir. Klasörde belgesi olmayan kişi onaya GÖNDERİLMEZ (atlanır, uyarılır — elle yükle).

Akış (canlı doğrulandı 2026-06-17): liste → taslağa ÇİFT TIKLA → [belgeleri yükle] → "Onaya Gönder"
→ "Emin misiniz? Evet" → kayıt "Yönetici Şirket Tarafından Onaylanmış" olur. E-İMZA/ŞİFRE ÇIKMIYOR.

YÖNTEM (sağlam): Liste DURUM filtresini SADECE "Değerlendirmeye Gönderilmemiş"e ayarla → listede yalnız
taslaklar kalır. Onaya gönderdikçe o kayıt listeden düşer → hep "ilk taslağı onayla, listeyi tazele" yeter
(isim-arama / sayfalama gerekmez). GÜVENLİK: sadece `izin_done_<lok>_<sheet>.txt`'deki isimleri onaylar;
listede bizim-olmayan taslak çıkarsa DOKUNMAZ (atlar, uyarır). Onay sonrası kaydın listeden düştüğünü
DOĞRULAR. Resume: `izin_onaya_done_<lok>_<sheet>.txt`. GERİ ALINMASI ZOR (resmi onay) — kullanıcı istedi.

✅ CANLI KANITLI (2026-06-19, port 9224): çok-dosya döngüsü 23/23 kişide çalıştı (Senanur 3 belge dahil),
   hepsi "Yönetici Şirket Tarafından Onaylanmış", 0 hata. Çözülen 2 Yıldız-bug'ı:
   (a) evrak tipi etiketi="Ücretli Yıllık İzin Formu" + JS'te 'İ'.toLowerCase()≠'i' → fold ile eşlenir;
   (b) portal SADECE PDF kabul (.xlsm/.xlsx → "Dosya tipi desteklenmemektedir") → Excel formlar
       LibreOffice ile 1-sayfa PDF'e çevrilip yüklendi (orijinaller _xlsx_orijinal/'de).

Orkestratör --belge-listesi ile ön kontrolde kişi başına doğrulanan PDF yollarını iletir.

Bayraklar: --dosya-klasor (kişi-başı belge klasörü), --person (tek kişi/retry, whitelist korur),
           --sadece-yukle (onaya GÖNDERMEDEN test), PROBE_UPLOAD=1 env (evrak-tipi/upload-fail modal dump).
Kullanım (lokasyon default=Yıldız):
  DGS_CDP=http://localhost:9224 python3 izin_onaya_dosyali_yildiz.py --person "<kişi>" --sadece-yukle --limit 1  # güvenli test
  DGS_CDP=http://localhost:9224 python3 izin_onaya_dosyali_yildiz.py                                            # tüm Yıldız
"""
from __future__ import annotations
import argparse
import json
import os
import re
import sys
import unicodedata
from collections import defaultdict
from urllib.parse import urlsplit

from playwright.sync_api import sync_playwright

from izin_belge import portal_filename, resolve_person_files
from izin_data_v2 import fold, resolve_park

CDP = os.environ.get("DGS_CDP", "http://localhost:9222")  # paralel park için: DGS_CDP ile port ver (9222 BV / 9223 TPIz)
DONEM_LABEL = "MAYIS 2026"   # dönem ETİKETİ — park-bazlı Donem_Id değişir → etiketle seç (Plan §5.2)
LOG = "[ONAY]"

# JS tarafı Türkçe-fold (Python fold ile aynı kural)
_JS_FOLD = r"""const fold=s=>{const tr={'ç':'c','Ç':'c','ğ':'g','Ğ':'g','ı':'i','İ':'i','I':'i','ö':'o','Ö':'o','ş':'s','Ş':'s','ü':'u','Ü':'u','â':'a','Â':'a','î':'i','Î':'i','û':'u','Û':'u'};
  return s.split('').map(c=>tr[c]||c).join('').toLowerCase().replace(/\s+/g,' ').trim();};"""


def log(*a):
    print(LOG, *a, flush=True)


def _mask_key(tc):
    if re.fullmatch(r"[0-9]{11}", tc):
        return tc[:2] + tc[-2:]
    if re.fullmatch(r"[0-9]{2}\*+(?:[0-9]{2})?", tc):
        return tc.replace("*", "")
    return ""


def saved_portal_people(path, park_code, donem):
    """Yalnız bu park ve dönemde başarıyla kaydedilmiş taslakların kişi kimliğini al."""
    found = defaultdict(set)
    if os.path.isfile(path):
        with open(path, encoding="utf-8") as log_file:
            for line in log_file:
                try:
                    record = json.loads(line)
                except ValueError:  # yarım kalmış günlük satırı kimlik kanıtı değildir
                    continue
                if not isinstance(record, dict) or not (record.get("park") == park_code
                        and record.get("donem") == donem and record.get("commit") is True
                        and record.get("kaydedildi") is True and record.get("durum") == "girildi"):
                    continue
                tc, name = str(record.get("tc", "")), fold(record.get("ad"))
                portal = record.get("portal_person")
                if not isinstance(portal, dict) or not re.fullmatch(r"[0-9]{11}", tc):
                    continue
                portal_name, pid = fold(portal.get("ad")), str(portal.get("id", ""))
                if portal_name and re.fullmatch(r"[1-9][0-9]*", pid):
                    found[(tc, name)].add((portal_name, pid))
    return found


def load_file_plan(path, done_pairs, person="", saved_people=None):
    """Ön kontroldeki dosyaları yalnız bu koşunun kaydedilmiş taslaklarına bağla."""
    with open(path, encoding="utf-8") as source:
        rows = json.load(source)
    if not isinstance(rows, list):
        raise ValueError("Belge listesi geçersiz.")

    done = {(tc.strip(), fold(ad)) for tc, ad in done_pairs}
    same_name = defaultdict(set)
    for tc, ad in done:
        if tc:
            same_name[ad].add(tc)
    files_used = set()
    entries = {}
    for row in rows:
        if not isinstance(row, dict):
            raise ValueError("Belge listesinde geçersiz satır.")
        tc, ad, files = row.get("tc"), row.get("ad"), row.get("files")
        if not isinstance(tc, str) or not _mask_key(tc) or not fold(ad):
            raise ValueError("Belge listesinde geçersiz kimlik.")
        identity = (tc, fold(ad))
        if identity in entries or not isinstance(files, list) or not files:
            raise ValueError(f"{ad} için belge listesi eksik veya yinelenmiş.")
        row = {"tc": tc, "ad": ad, "files": files}
        for file in files:
            if not (isinstance(file, str) and os.path.isabs(file) and
                    file.lower().endswith(".pdf") and os.path.isfile(file)):
                raise ValueError(f"{ad} için listedeki PDF okunamadı.")
            real = os.path.normcase(os.path.realpath(file))
            if real in files_used:
                raise ValueError(f"{ad} için aynı PDF birden fazla kez eşleşmiş.")
            files_used.add(real)
        portal_name = identity[1]
        known = (saved_people or {}).get(identity, set())
        if known and identity in done:
            if len(known) != 1:
                raise ValueError(f"{ad} için taslak günlüğünde çelişen portal kimlikleri var.")
            verified_name, pid = next(iter(known))
            # Sadece maske ve tek ortak kelimeyle farklı bir soyada belge bağlama.
            if verified_name != portal_name and len(set(verified_name.split()) & set(portal_name.split())) >= 2:
                portal_name = verified_name
                row = dict(row, portal_id=pid)
        entries[identity] = (row, portal_name)
        same_name[portal_name].add(tc)
        same_name[identity[1]].add(tc)

    result = {}
    for (tc, name), (row, portal_name) in entries.items():
        if person and person.strip() != tc and fold(person) != name:
            continue
        if (tc, name) not in done and ("", name) not in done:
            continue
        if (tc, name) not in done and len(same_name[name]) != 1:
            raise ValueError(f"{row['ad']} için eski devam kaydının kimliği belirsiz.")
        mask = _mask_key(tc)
        for other in same_name[portal_name] - {tc}:
            other_mask = _mask_key(other)
            if other_mask.startswith(mask) or mask.startswith(other_mask):
                raise ValueError(f"{row['ad']} için portalda ayırt edilemeyen kimlikler var.")
        key = f"{portal_name}\t{mask}"
        if key in result:
            raise ValueError(f"{row['ad']} için aynı portal kaydına bağlı birden fazla belge planı var.")
        result[key] = row
    return result


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


def pending_plan_files(page, paths):
    """Geçerli her portal satırında tek bir dosya adı gerekir. Yalnız eksik yolları döndür."""
    expected = {}
    for path in paths:
        name = portal_filename(path)
        if name in expected:
            raise ValueError(f"Seçilen PDF'ler portalda aynı dosya adıyla görünecek: {name}")
        expected[name] = path

    rows = page.evaluate(r"""()=>{
        const modal=[...document.querySelectorAll('.modal.in,.modal.show,.ui-dialog')].find(el=>
            el.offsetParent!==null && /Dosya Y[üu]kleme/i.test(el.innerText));
        if(!modal)return null;
        const docs=[];
        for(const row of modal.querySelectorAll('tr')){
            const text=row.innerText;
            const names=[...row.querySelectorAll('td,a,span')]
                .map(el=>el.textContent.trim()).filter(value=>/\.pdf$/i.test(value));
            if(/Hata|Ba[şs]ar[ıi]s[ıi]z/i.test(text))continue;
            if(/\d\s*[KM]B\b/i.test(text))docs.push([...new Set(names)]);
            else if(names.length)throw new Error('Yüklenen dosyanın durumu okunamadı.');
        }
        return docs;
    }""")
    if rows is None:
        raise RuntimeError("Dosya yükleme penceresi okunamadı.")
    seen = set()
    for row in rows:
        names = {unicodedata.normalize("NFC", value).strip() for value in row}
        if len(names) != 1:
            raise ValueError("Yüklü satırda tek bir PDF adı doğrulanamadı.")
        name = names.pop()
        if name not in expected or name in seen:
            raise ValueError(f"Yüklü PDF beklenmiyor veya birden fazla kez yüklenmiş: {name}")
        seen.add(name)
    return [path for name, path in expected.items() if name not in seen]


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
    ap.add_argument("--lokasyon", default="Yıldız")
    ap.add_argument("--sheet", default="Mayıs")
    ap.add_argument("--donem", default=DONEM_LABEL, help="Dönem etiketi (ör. 'HAZİRAN 2026'); grid+dönem-filtresi bununla eşleşir. Park-agnostik (İYTE/YTP).")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--dosya-klasor", default="/Users/ekoreiz/Downloads/Yıldız İzin Formları",
                    help="Her kişinin belge(ler)inin DGS adıyla (AD SOYAD[ (n)].ext) durduğu klasör. Kişiye ait TÜM belgeler sırayla yüklenir.")
    ap.add_argument("--belge-listesi", help="Ön kontrolde doğrulanan kişi ve PDF yollarının JSON listesi.")
    ap.add_argument("--dosya", default="",
                    help="(Opsiyonel) FALLBACK tek dosya — klasörde kişinin belgesi YOKSA bu yüklenir. Boş='' → fallback yok.")
    ap.add_argument("--sadece-yukle", action="store_true",
                    help="🧪 TEST: belgeleri yükle ama ONAYA GÖNDERME (geri-alınamaz resmi onayı atla). Resume'a YAZMAZ; taslak kalır.")
    ap.add_argument("--person", default="",
                    help="Sadece bu kişiyi işle (izin_done içinde OLMALI — whitelist korunur). Tek-kişi test/retry için.")
    args = ap.parse_args()
    donem = args.donem   # dönem etiketi — JS desenlerine + set_donem'e parametrik geçer (Mayıs varsayılan → davranış korunur)
    if not args.belge_listesi and args.dosya_klasor and not os.path.isdir(args.dosya_klasor):
        log(f"UYARI: --dosya-klasor bulunamadı: {args.dosya_klasor!r} (kişi-başı belge çözülemeyecek; sadece --dosya fallback varsa yüklenir)")

    done_file = f"izin_done_{args.lokasyon}_{args.sheet}.txt"
    if not os.path.exists(done_file):
        log(f"HATA: {done_file} yok."); sys.exit(1)
    # done-dosyası satırları "T.C.\tAd" (yeni) veya düz "Ad" (eski) — ikisini de destekle.
    pairs = []                                        # (tc|"", ad)
    for l in open(done_file, encoding="utf-8"):
        l = l.strip()
        if not l:
            continue
        parts = l.split("\t")
        pairs.append((parts[0] if len(parts) > 1 else "", parts[-1]))
    file_plan = None
    if args.belge_listesi:
        try:
            park = resolve_park(args.lokasyon)
            if not park or not park.onay_pdf or not park.onay_dogrulandi:
                ap.error("Seçilen parkta belge yükleme desteği doğrulanmadı.")
            history = saved_portal_people(f"izin_runlog_{park.label}_{args.sheet}.jsonl", park.code, donem)
            file_plan = load_file_plan(args.belge_listesi, pairs, args.person, history)
        except (OSError, ValueError) as exc:
            ap.error(str(exc))
        pairs = [(row["tc"], row["ad"]) for row in file_plan.values()]
        if not pairs:
            log("Belge listesinde bu koşu için kaydedilmiş taslak yok.")
            return
    elif args.person:                                 # tek-kişi: whitelist'i o kişiye daralt (güvenlik korunur)
        pf = fold(args.person)
        keep = [(tc, ad) for (tc, ad) in pairs
                if pf == fold(ad) or pf in fold(ad) or args.person.strip() == str(tc).strip()]
        if not keep:
            log(f"HATA: --person {args.person!r} (fold={pf!r}) {done_file} içinde yok — whitelist dışı, çıkılıyor."); sys.exit(1)
        pairs = keep
        log(f"--person aktif: SADECE {args.person!r} işlenecek ({len(pairs)} kayıt).")
    our = set(fold(ad) for _, ad in pairs)            # folded isim (isim eşleşmesi + geriye uyum)
    our_tc = set()                                    # maskeli-T.C. anahtarı (ilk2+son2) — kızlık/evlilik-soyadından BAĞIMSIZ kimlik
    name_by_tc = {}                                   # maskeli-T.C. → Excel/DGS adı (grid portal-adı farklıysa belgeyi bununla çöz)
    for tc, ad in pairs:
        tc = str(tc).strip()
        if tc.isdigit() and len(tc) == 11:
            key = tc[:2] + tc[-2:]
            our_tc.add(key)
            name_by_tc[key] = ad
    onay_file = f"izin_onaya_done_{args.lokasyon}_{args.sheet}.txt"
    onaylanan = set(fold(l) for l in open(onay_file, encoding="utf-8").read().splitlines()) if os.path.exists(onay_file) else set()
    hedef = len([x for x in our if x not in onaylanan])
    limit = args.limit if args.limit else 10**9
    log(f"{len(our)} taslak (izin_done); {len(onaylanan)} zaten onaya gönderilmiş; {min(hedef, limit)} işlenecek.")

    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp(CDP)
        if file_plan is not None:
            park = resolve_park(args.lokasyon)
            host = urlsplit(park.portal_url).hostname if park else None
            tabs = [p for p in browser.contexts[0].pages if host and urlsplit(p.url).hostname == host]
            if len(tabs) != 1:
                ap.error("Belge yüklemek için seçilen parkın tek bir portal sekmesi açık olmalı.")
            page = tabs[0]
        else:
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
            # Dönem = donem (etiket, park/ay-agnostik) — ETİKETE göre (park-bazlı Donem_Id değişir → label ile select_option)
            set_donem_by_label(page, donem)
            # DURUM filtresi: TÜM durumlar AÇIK. FIND_FIRST yine SADECE 'Gönderilmemiş'i (taslak) bulur; ama onay
            # DOĞRULAMASI kişinin GERÇEK OnayDurumu'nu (Bekleyen/Gönderilmiş/Onaylanmış) okuyabilsin diye hepsi görünmeli.
            page.evaluate(r"""()=>{document.querySelectorAll('input[type=checkbox][name^="OnayDurumu_Id"], input[type=checkbox][id^="OnayDurumu_Id"]').forEach(cb=>{ if(!cb.checked){ cb.click(); } });}""")
            page.wait_for_timeout(300)

        def listele():
            page.evaluate(r"""()=>{const b=[...document.querySelectorAll('a,button')].find(x=>x.offsetParent!==null
                  && /Listele/.test(x.textContent)); if(b)b.click();}""")
            wait_ajax(page)   # listeleme AJAX'ı bitene kadar bekle (loading-fix)

        # İlk "Gönderilmemiş" taslağı aç (bizim setten, failed olmayan). Adı döndür; yoksa null.
        # Ay etiketi (donem) grid satırında görünür → JS regex'lerine GÖMÜLÜR (park/ay-agnostik; Mayıs'ta orijinalle bit-bit aynı).
        FIND_FIRST = (r"""(args)=>{const [our, failed, ourTc, planned]=args; %s
          const rows=[];
          for(const tr of document.querySelectorAll('tr')){
            const t=tr.innerText.replace(/\s+/g,' ').trim();
            const m=t.match(/%s\s+(.+?)\s+(\d{2})\*+(\d{2})/);
            if(!m || !tr.getClientRects().length || tr.querySelector('tr'))continue;
            rows.push({tr, t, name:m[1].trim(), nf:fold(m[1]), mask:m[2]+m[3]});}
          for(const row of rows){
            if(!row.t.includes('Değerlendirmeye Gönderilmemiş'))continue;
            const key=row.nf+'\t'+row.mask;
            if(failed.includes(key))continue;
            let planKey=null;
            if(planned!==null){
              const prefixKey=row.nf+'\t'+row.mask.slice(0,2);
              planKey=planned.includes(key)?key:planned.includes(prefixKey)?prefixKey:null;
              if(!planKey)continue;
              // Bir görünür kimlik iki satıra uyuyorsa çalışan ve taslak ayırt edilemez.
              const mask=planKey===key?row.mask:row.mask.slice(0,2);
              if(rows.filter(r=>r.nf===row.nf && r.mask.startsWith(mask)).length!==1)continue;
            }else if(!our.includes(row.nf) && !ourTc.includes(row.mask))continue;
            row.tr.dispatchEvent(new MouseEvent('dblclick',{bubbles:true,view:window}));
            return [row.name,row.mask,planKey];}
          return null;}""") % (_JS_FOLD, re.escape(donem))

        # Bir ismin hâlâ "Gönderilmemiş" satırı var mı? (onay doğrulaması)
        STILL_DRAFT = (r"""(args)=>{const [nm]=args; %s const nf=fold(nm);
          for(const tr of document.querySelectorAll('tr')){const t=tr.innerText.replace(/\s+/g,' ').trim();
            if(!/%s/.test(t)||!/Değerlendirmeye Gönderilmemiş/.test(t))continue;
            const m=t.match(/%s\s+(.+?)\s+\d+\*+\d+/); if(m && fold(m[1])===nf) return true;}
          return false;}""") % (_JS_FOLD, donem, donem)

        # bizim-olmayan kalan taslak sayısı (uyarı için)
        FOREIGN_LEFT = (r"""(args)=>{const [our]=args; %s let n=0;
          for(const tr of document.querySelectorAll('tr')){const t=tr.innerText.replace(/\s+/g,' ').trim();
            if(!/%s/.test(t)||!/Değerlendirmeye Gönderilmemiş/.test(t))continue;
            const m=t.match(/%s\s+(.+?)\s+\d+\*+\d+/); if(m && !our.includes(fold(m[1]))) n++;}
          return n;}""") % (_JS_FOLD, donem, donem)

        # Kişinin listedeki GERÇEK OnayDurumu'nu (string) döndür — onay doğrulaması için (body-text banner GÜVENİLMEZ).
        # 'Değerlendirmeye Gönderilmemiş'=taslak; 'Değerlendirme Bekleyen'/'...Gönderilmiş'/'...Onaylanmış'=onaya gitti.
        STATUS_OF = (r"""([nf,mask])=>{const STAT=['Yönetici Şirket Tarafından Onaylanmış','Değerlendirmeye Gönderilmemiş','Değerlendirme Bekleyen','Değerlendirmeye Gönderilmiş','Onaylanmış','İmzalanmış','Reddedilmiş']; %s
          for(const tr of document.querySelectorAll('tr')){const t=tr.innerText.replace(/\s+/g,' ').trim();
            if(!/%s/.test(t))continue; const m=t.match(/%s\s+(.+?)\s+(\d{2})\*+(\d{2})/); if(!m)continue;
            if(fold(m[1])===nf && m[2]+m[3]===mask){for(const s of STAT)if(t.includes(s))return s; return '(?)';}}
          return null;}""") % (_JS_FOLD, donem, donem)

        def reset_list_filter():
            # close_form dönem/durum filtresini sıfırlayabilir → dönem + TÜM durum yeniden uygula (Listele AYRI çağrılır)
            try: set_donem_by_label(page, donem)
            except Exception: pass
            page.evaluate(r"""()=>{document.querySelectorAll('input[type=checkbox][name^="OnayDurumu_Id"],input[type=checkbox][id^="OnayDurumu_Id"]').forEach(cb=>{if(!cb.checked)cb.click();});}""")
            page.wait_for_timeout(200)

        open_list()
        ok = 0
        failed = set()
        processed = set()
        processed_plan_keys = set()
        planned_keys = list(file_plan) if file_plan is not None else None
        # zaten onaya gönderilmişleri başta resume'a yazılı say
        for _iter in range(len(pairs) + 25):
            if ok >= limit:
                break
            if not list_open():        # liste kapandıysa yeniden aç (filtre+dönem ile)
                open_list()
            else:
                reset_list_filter()    # liste açık ama filtre sıfırlanmış olabilir → dönem+TÜM durum yenile
            listele()
            res = page.evaluate(FIND_FIRST, [list(our), list(failed | processed), list(our_tc), planned_keys])
            if not res:
                if file_plan is not None:
                    log("Görünen listede belge planıyla adı ve kimliği eşleşen başka bir taslak bulunamadı.")
                else:
                    foreign = page.evaluate(FOREIGN_LEFT, [list(our)])
                    log(f"Bizim gönderilmemiş taslak kalmadı. (bizim-olmayan kalan taslak: {foreign})")
                break
            name, tcKey, plan_key = res
            row_key = f"{fold(name)}\t{tcKey}"
            if file_plan is not None:
                belge_ad = file_plan[plan_key]["ad"]
                dosyalar = file_plan[plan_key]["files"]
            else:
                belge_ad = name_by_tc.get(tcKey, name)
                dosyalar = resolve_person_files(args.dosya_klasor, belge_ad)
                if not dosyalar and fold(belge_ad) != fold(name):
                    dosyalar = resolve_person_files(args.dosya_klasor, name)
                if not dosyalar and args.dosya:
                    dosyalar = [args.dosya]
            if fold(belge_ad) != fold(name):
                log(f"   ℹ️ İSİM UYUŞMAZLIĞI: portal='{name}' ↔ Excel='{belge_ad}' (T.C.…{tcKey}) — T.C. ile eşlendi.")
            wait_ajax(page)   # form açıldı — loading bitene kadar bekle (acele etme)
            if file_plan is not None and file_plan[plan_key].get("portal_id"):
                form = page.get_by_role("dialog", name=re.compile(r"Personel İzin Bildirim Formu(?! Listesi)"))
                field = form.locator("input[type=hidden][name='Personel_Id']")
                if field.count() != 1 or field.input_value() != file_plan[plan_key]["portal_id"]:
                    failed.add(row_key)
                    log(f"!! {name}: açık formun personel ID'si kaydedilen taslakla doğrulanamadı.")
                    close_form()
                    continue
            # 🔴 ÖNCEKİ kişiden kalan confirm/alert modal'ı DOM'dan KALDIR — birikince "Dosya Yükle"/Evet click'ini
            #   bloklar + duplicate #modalYes → yanlış-pozitif onay (2026-06-19 bulk2: ilk 10 'Dosya Yükle' timeout;
            #   MEHMET KORAL diag: stale 'başarı' banner). Form ve "Dosya Yükleme Modülü"ne DOKUNMA.
            page.evaluate(r"""()=>{[...document.querySelectorAll('.modal,.sweet-alert,.modal-backdrop,.sweet-overlay')]
                .filter(m=>!/Personel İzin Bildirim Formu|Dosya Y[üu]kleme Mod/i.test(m.innerText||''))
                .forEach(m=>m.remove());}""")
            page.wait_for_timeout(200)
            # 🔴 DİLEKÇE YÜKLE (TPIz: izin onayı için 'yıllık izin formu' PDF ZORUNLU — yoksa "dosya yükle bölümünden
            # yüklemelisiniz" hatası, taslak kalır). Zaten geçerli dosya varsa tekrar YÜKLEME. Max 1 MB.
            # 🔴 DOĞRU SIRA (2026-06-19 ANIL/ANIL AKSOY ile CANLI doğrulandı): gizli #gosb-fileupload-button'a
            # DOĞRUDAN set_input_files yapmak "Dosya Yüklemesi Başarısız" verir (sunucuya EvrakTipi gitmediği için;
            # boyut bile yanlış: 761.55 vs gerçek 760.51 KB). ŞART = blueimp jQuery File Upload sırası:
            #   1) "Dosya Yükle" → "Dosya Yükleme Modülü" modalı
            #   2) "+ Dosya Ekle" (label.fileinput-button) → GİZLİ "Evrak Tipi" paneli (.upload-evrak-tipi-panel) açılır
            #      (native chooser DEĞİL — sadece paneli gösterir)
            #   3) "yıllık izin formu" radyosunu (name=EvrakTipi) SEÇ → #fileupload-aciklama + #gosb-fileupload-button GÖRÜNÜR olur
            #   4) ŞİMDİ set_input_files → blueimp autoUpload, EvrakTipi seçili olduğu için POST KABUL edilir → başarı satırı.
            # Radyoyu ETİKETLE seç (value park-bazlı: TPIz=54; "value HARDCODE ETME" kuralı).
            # 🌟 YILDIZ: tek global dosya yerine KİŞİNİN KENDİ belge(ler)ini klasörden çöz → HEPSİNİ sırayla yükle.
            #   (örn. SENANUR IŞIK → 3 belge = 3 ayrı upload). Tek-dosyada davranış tabandakiyle aynıdır (exp=1).
            if not dosyalar:
                failed.add(row_key); log(f"!! {name} → klasörde belge bulunamadı, onaya GÖNDERİLMEDİ (elle yükle)"); close_form(); continue
            # modaldaki GEÇERLİ (KB/MB var, Hata/Başarısız YOK) satır SAYISI — yükleme doğrulaması buradan (boolean değil sayım)
            count_rows = lambda: page.evaluate(r"""()=>{const m=[...document.querySelectorAll('.modal.in,.ui-dialog')].find(x=>x.offsetParent!==null && /Dosya Y[üu]kleme/i.test(x.innerText));
                return m ? [...m.querySelectorAll('tr')].filter(tr=>/KB|MB/.test(tr.innerText) && !/Hata|Ba[şs]ar[ıi]s[ıi]z/i.test(tr.innerText)).length : 0;}""")
            exp = len(dosyalar)
            try:
                page.locator("a:has-text('Dosya Yükle'), button:has-text('Dosya Yükle')").filter(visible=True).first.click(timeout=4000)
                page.wait_for_timeout(1200)
                have = count_rows()
                if file_plan is not None:
                    pending = pending_plan_files(page, dosyalar)
                elif have >= exp:
                    pending = []
                elif have:
                    raise RuntimeError(f"{have}/{exp} belge mevcut, eksikler dosya adıyla doğrulanamadı.")
                else:
                    pending = dosyalar
                if not pending:
                    log(f"  {name}: {have} belge zaten yüklü → yükleme atlandı")
                else:
                    for i, fpath in enumerate(pending, exp - len(pending) + 1):
                        if file_plan is not None and fpath not in pending_plan_files(page, dosyalar):
                            continue
                        sz_kb = os.path.getsize(fpath) / 1024
                        if sz_kb > 1024:
                            log(f"     ⚠️ {os.path.basename(fpath)} {sz_kb:.0f} KB > 1 MB — portal reddedebilir")
                        before = count_rows()
                        # 2) Evrak Tipi paneli kapalıysa "+ Dosya Ekle" ile aç (her belge için TEKRAR; panel yüklemeden sonra kapanır)
                        panel_open = page.evaluate(r"""()=>{const p=document.querySelector('.upload-evrak-tipi-panel'); return !!(p && p.offsetParent!==null);}""")
                        if not panel_open:
                            page.locator("label.fileinput-button:has-text('Dosya Ekle')").filter(visible=True).first.click(timeout=4000, no_wait_after=True)
                            page.wait_for_timeout(700)
                        # 3) Evrak tipi radyosunu seç — Türkçe-İ GÜVENLİ (fold). Yıldız etiketi "Ücretli Yıllık İzin Formu";
                        #    JS regex /izin/i "İzin"i TUTMAZ ('İ'.toLowerCase()≠'i') → fold ile eşle; tek seçenek varsa direkt al.
                        radio_ok = page.evaluate(r"""()=>{%s
                          const panel=document.querySelector('.upload-evrak-tipi-panel'); const scope=panel||document;
                          const radios=[...scope.querySelectorAll('#GerekliEvrakTanimi_Id input[type=radio], input[type=radio][name=EvrakTipi]')];
                          let pick=radios.find(r=>{const li=r.closest('li')||r.parentElement; return fold(li?li.innerText:'').includes('yillik izin');});
                          if(!pick && radios.length===1) pick=radios[0];
                          if(!pick) return false;
                          const li=pick.closest('li'); const lb=(li||pick.parentElement).querySelector('label');
                          pick.checked=true; (lb||li||pick).click(); pick.dispatchEvent(new Event('change',{bubbles:true}));
                          return true;}""" % _JS_FOLD)
                        if not radio_ok:
                            if os.environ.get("PROBE_UPLOAD"):
                                diag = page.evaluate(r"""()=>{const vis=e=>e&&e.offsetParent!==null;
                                  const modal=[...document.querySelectorAll('.modal,.ui-dialog')].find(x=>vis(x)&&/Dosya Y[üu]kleme/i.test(x.innerText||''));
                                  const panel=document.querySelector('.upload-evrak-tipi-panel');
                                  return JSON.stringify({
                                    modalText:(modal?modal.innerText:'(modal yok)').replace(/\n{2,}/g,'\n').slice(0,900),
                                    panel: panel?{gorunur:vis(panel),html:panel.outerHTML.slice(0,1400)}:'(.upload-evrak-tipi-panel YOK)',
                                    radios:[...(modal||document).querySelectorAll('input[type=radio]')].map(r=>({name:r.name,value:r.value,gorunur:vis(r),label:((r.closest('label')||r.parentElement)?.innerText||'').trim().slice(0,50)})),
                                    selects:[...(modal||document).querySelectorAll('select')].filter(vis).map(s=>({name:s.name,id:s.id,opts:[...s.options].map(o=>(o.text||'').trim()).filter(Boolean).slice(0,20)})),
                                    evrak:[...document.querySelectorAll('[id*=Evrak],[name*=Evrak],[class*=evrak]')].filter(vis).map(e=>({tag:e.tagName,id:e.id,name:e.getAttribute('name'),txt:(e.innerText||'').trim().slice(0,40)})).slice(0,12),
                                    buttons: modal?[...modal.querySelectorAll('a,button,label')].filter(vis).map(b=>(b.innerText||'').trim()).filter(Boolean).slice(0,25):[]
                                  },null,2);}""")
                                log("🔬 PROBE evrak-tipi dump:\n" + diag)
                            raise RuntimeError(f"'yıllık izin formu' evrak tipi bulunamadı (belge {i}/{exp})")
                        # 4) input görünür olunca dosyayı set et → autoUpload; satır artana kadar bekle (max ~6 sn)
                        page.wait_for_selector("#gosb-fileupload-button", state="visible", timeout=6000)
                        page.wait_for_timeout(400)
                        page.locator("#gosb-fileupload-button").set_input_files(fpath)
                        ok_row = False
                        for _w in range(15):
                            page.wait_for_timeout(400)
                            if count_rows() > before:
                                ok_row = True; break
                        if not ok_row:
                            if os.environ.get("PROBE_UPLOAD"):
                                d2 = page.evaluate(r"""()=>{const vis=e=>e&&e.offsetParent!==null;
                                  const modal=[...document.querySelectorAll('.modal,.ui-dialog')].find(x=>vis(x)&&/Dosya Y[üu]kleme/i.test(x.innerText||''));
                                  return JSON.stringify({
                                    fileInputs:[...document.querySelectorAll('input[type=file]')].map(f=>({id:f.id,accept:f.accept,multiple:f.multiple})),
                                    rows: modal?[...modal.querySelectorAll('tr')].map(tr=>tr.innerText.replace(/\s+/g,' ').trim()).filter(Boolean).slice(0,12):[],
                                    alerts:[...document.querySelectorAll('.sweet-alert,.toast,.message,.alert,.card-message,.validation-summary-errors')].filter(vis).map(a=>(a.innerText||'').trim().slice(0,120)).slice(0,6),
                                    modalTail:(modal?modal.innerText:'').replace(/\n{2,}/g,'\n').slice(-400)
                                  },null,2);}""")
                                log("🔬 PROBE upload-fail dump:\n" + d2)
                            raise RuntimeError(f"belge {i}/{exp} '{os.path.basename(fpath)}' yüklenemedi (satır artmadı/Hata)")
                        log(f"     ✓ {os.path.basename(fpath)} ({i}/{exp})")
                final = count_rows()
                if file_plan is not None and pending_plan_files(page, dosyalar):
                    raise RuntimeError("Yüklenen PDF'ler kişi için seçilen dosyaların tamamıyla eşleşmedi.")
                page.evaluate(r"""()=>{const m=[...document.querySelectorAll('.modal.in,.ui-dialog')].find(x=>x.offsetParent!==null && /Dosya Y[üu]kleme/i.test(x.innerText));
                    if(m){const c=m.querySelector('.close,.ui-dialog-titlebar-close,.fa-times,button.close'); if(c)c.click();}}""")
                page.wait_for_timeout(900)
                if final < exp:
                    failed.add(row_key); log(f"!! {name} → {final}/{exp} belge yüklendi (EKSİK), onaya gönderilmedi; elle bak"); close_form(); continue
                log(f"  {name}: {final}/{exp} belge yüklendi ✓")
            except Exception as e:
                failed.add(row_key); log(f"!! {name} → Dosya Yükle hatası ({str(e)[:60]}), atlandı"); close_form(); continue
            wait_ajax(page)
            # 🧪 SADECE-YÜKLE test modu: belgeler yüklendi → ONAYA GÖNDERME (geri-alınamaz onayı atla, taslak kalsın)
            if args.sadece_yukle:
                ok += 1
                processed.add(row_key)
                processed_plan_keys.add(plan_key)
                log(f"[{ok}] {name} → SADECE-YÜKLE ✓ ({len(dosyalar)} belge yüklü, ONAYA GÖNDERİLMEDİ); modalda gözle/temizle")
                close_form()
                continue
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
            # kalan onay overlay/modal'ı temizle (menü/Listele bloklamasın), sonra formu kapat
            page.evaluate(r"""()=>{document.querySelectorAll('.modal.in .close,.modal.show .close,.sweet-alert .confirm,.sweet-alert button.confirm').forEach(b=>{try{b.click()}catch(e){}}); document.querySelectorAll('.sweet-overlay,.sweet-alert,.modal-backdrop').forEach(e=>e.style.display='none');}""")
            close_form()
            # 🔴 DOĞRULAMA = listeden GERÇEK OnayDurumu (2026-06-19; eski body-text "Kayıt Başarı" banner'ı GÜVENİLMEZDİ:
            #   bayat metin=yanlış-POZİTİF — dosyasız onay taslak kalsa bile banner çıkabiliyordu; kısa poll=yanlış-NEGATİF).
            #   close_form filtreyi sıfırlayabilir → her okumada reset_list_filter+Listele; satır güncellenene dek kısa poll.
            #   BAŞARI = durum 'Gönderilmemiş' DEĞİL (yani Bekleyen/Gönderilmiş/Onaylanmış'a geçmiş). Self-verifying — ayrı check_status gerekmez.
            durum = None
            sent_states = {"Yönetici Şirket Tarafından Onaylanmış", "Değerlendirme Bekleyen",
                           "Değerlendirmeye Gönderilmiş", "Onaylanmış", "İmzalanmış"}
            for _v in range(6):
                if not list_open():
                    open_list()
                else:
                    reset_list_filter()
                listele()
                durum = page.evaluate(STATUS_OF, [fold(name), tcKey])
                if durum in sent_states:
                    break
                page.wait_for_timeout(600)
            if durum not in sent_states:
                failed.add(row_key)
                log(f"!! {name} → ONAYA GİTMEDİ (durum: {durum or 'okunamadı/taslak'}); elle bak")
                continue
            ok += 1
            processed.add(row_key)
            processed_plan_keys.add(plan_key)
            with open(onay_file, "a", encoding="utf-8") as f:
                f.write(belge_ad + "\n")   # Excel/DGS adı → resume+onaylanan `our` (Excel isimleri) ile eşleşsin
            log(f"[{ok}] {name} → ONAYA GÖNDERİLDİ ✓ (gerçek durum: {durum})")

        log(f"\n==== ÖZET: bu koşuda {ok} kişi onaya gönderildi ====")
        if file_plan is not None and ok < limit:
            pending = [row["ad"] for key, row in file_plan.items()
                       if key not in processed_plan_keys and fold(row["ad"]) not in onaylanan]
            if pending:
                log("  Bu koşuda işlem doğrulanamayan çalışanlar: " + ", ".join(pending))
        if failed:
            log(f"  !! ONAYA GİTMEYEN ({len(failed)}): elle bak.")
        try:
            input(f"{LOG} >> Bitti. ENTER...")
        except EOFError:
            pass


if __name__ == "__main__":
    main()
