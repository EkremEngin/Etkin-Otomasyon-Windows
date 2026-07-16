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

Belge eşleştirme: klasördeki dosyalar "AD SOYAD.ext" veya "AD SOYAD (n).ext" formatında (bu repo'da
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

Bayraklar: --dosya-klasor (kişi-başı belge klasörü), --person (tek kişi/retry, whitelist korur),
           --sadece-yukle (onaya GÖNDERMEDEN test), PROBE_UPLOAD=1 env (evrak-tipi/upload-fail modal dump).
Kullanım (lokasyon default=Yıldız):
  DGS_CDP=http://localhost:9224 python3 izin_onaya_dosyali_yildiz.py --person "<kişi>" --sadece-yukle --limit 1  # güvenli test
  DGS_CDP=http://localhost:9224 python3 izin_onaya_dosyali_yildiz.py                                            # tüm Yıldız
"""
from __future__ import annotations
import argparse
import os
import re
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


_NUM_SUFFIX = re.compile(r"\s*\((\d+)\)\s*$")   # " (1)", " (2)" ... son eki


def resolve_person_files(folder: str, name: str) -> list[str]:
    """Klasörde bu KİŞİYE ait belge(ler)i isimle çöz.
    Dosya adı 'AD SOYAD.ext' veya 'AD SOYAD (n).ext' (DGS adıyla adlandırılmış). fold() ile eşleşir
    (büyük/küçük + Türkçe diakritik bağımsız). (n) numarasına göre sıralı (1→2→3) tam yol listesi döner.
    Eşleşme yoksa boş liste."""
    if not folder or not os.path.isdir(folder):
        return []
    nf = fold(name)
    found: list[tuple[int, str]] = []
    for f in sorted(os.listdir(folder)):
        if f.startswith(".") or f.startswith("_"):
            continue
        p = os.path.join(folder, f)
        if not os.path.isfile(p):
            continue
        stem = os.path.splitext(f)[0]
        m = _NUM_SUFFIX.search(stem)
        num = int(m.group(1)) if m else 0
        core = _NUM_SUFFIX.sub("", stem)
        if fold(core) == nf:
            found.append((num, p))
    found.sort()
    return [p for _, p in found]


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
    ap.add_argument("--lokasyon", default="Yıldız")
    ap.add_argument("--sheet", default="Mayıs")
    ap.add_argument("--donem", default=DONEM_LABEL, help="Dönem etiketi (ör. 'HAZİRAN 2026'); grid+dönem-filtresi bununla eşleşir. Park-agnostik (İYTE/YTP).")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--dosya-klasor", default="/Users/ekoreiz/Downloads/Yıldız İzin Formları",
                    help="Her kişinin belge(ler)inin DGS adıyla (AD SOYAD[ (n)].ext) durduğu klasör. Kişiye ait TÜM belgeler sırayla yüklenir.")
    ap.add_argument("--dosya", default="",
                    help="(Opsiyonel) FALLBACK tek dosya — klasörde kişinin belgesi YOKSA bu yüklenir. Boş='' → fallback yok.")
    ap.add_argument("--sadece-yukle", action="store_true",
                    help="🧪 TEST: belgeleri yükle ama ONAYA GÖNDERME (geri-alınamaz resmi onayı atla). Resume'a YAZMAZ; taslak kalır.")
    ap.add_argument("--person", default="",
                    help="Sadece bu kişiyi işle (izin_done içinde OLMALI — whitelist korunur). Tek-kişi test/retry için.")
    args = ap.parse_args()
    donem = args.donem   # dönem etiketi — JS desenlerine + set_donem'e parametrik geçer (Mayıs varsayılan → davranış korunur)
    if args.dosya_klasor and not os.path.isdir(args.dosya_klasor):
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
    if args.person:                                   # tek-kişi: whitelist'i o kişiye daralt (güvenlik korunur)
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
        FIND_FIRST = (r"""(args)=>{const [our, failed, ourTc]=args; %s
          for(const tr of document.querySelectorAll('tr')){
            const t=tr.innerText.replace(/\s+/g,' ').trim();
            if(!/%s/.test(t)||!/Değerlendirmeye Gönderilmemiş/.test(t))continue;
            const m=t.match(/%s\s+(.+?)\s+(\d+)\*+(\d+)/); if(!m)continue;
            const nf=fold(m[1]); const tcKey=m[2]+m[3];
            if(failed.includes(nf)||failed.includes(tcKey))continue;
            if(!our.includes(nf) && !ourTc.includes(tcKey)) continue;   // GÜVENLİK: bizim (isim VEYA maskeli-T.C. — kızlık/evlilik farkında da tutar)
            tr.dispatchEvent(new MouseEvent('dblclick',{bubbles:true,view:window})); return [m[1].trim(), tcKey];}
          return null;}""") % (_JS_FOLD, donem, donem)

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
        STATUS_OF = (r"""(nf)=>{const STAT=['Yönetici Şirket Tarafından Onaylanmış','Değerlendirmeye Gönderilmemiş','Değerlendirme Bekleyen','Değerlendirmeye Gönderilmiş','Onaylanmış','İmzalanmış','Reddedilmiş']; %s
          for(const tr of document.querySelectorAll('tr')){const t=tr.innerText.replace(/\s+/g,' ').trim();
            if(!/%s/.test(t))continue; const m=t.match(/%s\s+(.+?)\s+\d+\*+\d+/); if(!m)continue;
            if(fold(m[1])===nf){for(const s of STAT)if(t.includes(s))return s; return '(?)';}}
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
        skipped_onay = [x for x in our if x in onaylanan]
        # zaten onaya gönderilmişleri başta resume'a yazılı say
        for _iter in range(len(our) + 25):
            if ok >= limit:
                break
            if not list_open():        # liste kapandıysa yeniden aç (filtre+dönem ile)
                open_list()
            else:
                reset_list_filter()    # liste açık ama filtre sıfırlanmış olabilir → dönem+TÜM durum yenile
            listele()
            res = page.evaluate(FIND_FIRST, [list(our), list(failed), list(our_tc)])
            if not res:
                foreign = page.evaluate(FOREIGN_LEFT, [list(our)])
                log(f"Bizim gönderilmemiş taslak kalmadı. (bizim-olmayan kalan taslak: {foreign})")
                break
            name, tcKey = res[0], res[1]                    # name=grid(portal) adı; tcKey=maskeli T.C. ilk2+son2
            belge_ad = name_by_tc.get(tcKey, name)          # belgeyi Excel/DGS adıyla çöz (portal adı kızlık/evlilikte farklı olabilir)
            if fold(belge_ad) != fold(name):
                log(f"   ℹ️ İSİM UYUŞMAZLIĞI: portal='{name}' ↔ Excel='{belge_ad}' (T.C.…{tcKey}) — T.C. ile eşlendi.")
            wait_ajax(page)   # form açıldı — loading bitene kadar bekle (acele etme)
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
            dosyalar = resolve_person_files(args.dosya_klasor, belge_ad)
            if not dosyalar and fold(belge_ad) != fold(name):
                dosyalar = resolve_person_files(args.dosya_klasor, name)   # yedek: belge portal-adıyla adlandırılmış olabilir
            if not dosyalar and args.dosya:
                dosyalar = [args.dosya]                       # fallback: global tek dosya (verildiyse)
            if not dosyalar:
                failed.add(fold(name)); log(f"!! {name} → klasörde belge bulunamadı, onaya GÖNDERİLMEDİ (elle yükle)"); close_form(); continue
            # modaldaki GEÇERLİ (KB/MB var, Hata/Başarısız YOK) satır SAYISI — yükleme doğrulaması buradan (boolean değil sayım)
            count_rows = lambda: page.evaluate(r"""()=>{const m=[...document.querySelectorAll('.modal.in,.ui-dialog')].find(x=>x.offsetParent!==null && /Dosya Y[üu]kleme/i.test(x.innerText));
                return m ? [...m.querySelectorAll('tr')].filter(tr=>/KB|MB/.test(tr.innerText) && !/Hata|Ba[şs]ar[ıi]s[ıi]z/i.test(tr.innerText)).length : 0;}""")
            exp = len(dosyalar)
            try:
                page.locator("a:has-text('Dosya Yükle'), button:has-text('Dosya Yükle')").filter(visible=True).first.click(timeout=4000)
                page.wait_for_timeout(1200)
                have = count_rows()
                if have >= exp:
                    log(f"  {name}: {have} belge zaten yüklü (≥{exp} beklenen) → yükleme atlandı")
                elif have > 0:
                    # kısmi durum (0<have<exp): hangi belge eksik bilinmiyor → çift-yükleme riski; güvenli taraf = elle
                    failed.add(fold(name)); log(f"!! {name} → modalda {have}/{exp} belge var (KISMİ), karışmasın diye ATLANDI; elle tamamla"); close_form(); continue
                else:
                    # have == 0 → kişinin TÜM belgelerini sırayla yükle. Her biri: "+ Dosya Ekle" → 'yıllık izin formu' → set_input_files
                    for i, fpath in enumerate(dosyalar, 1):
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
                page.evaluate(r"""()=>{const m=[...document.querySelectorAll('.modal.in,.ui-dialog')].find(x=>x.offsetParent!==null && /Dosya Y[üu]kleme/i.test(x.innerText));
                    if(m){const c=m.querySelector('.close,.ui-dialog-titlebar-close,.fa-times,button.close'); if(c)c.click();}}""")
                page.wait_for_timeout(900)
                if final < exp:
                    failed.add(fold(name)); log(f"!! {name} → {final}/{exp} belge yüklendi (EKSİK), onaya gönderilmedi; elle bak"); close_form(); continue
                log(f"  {name}: {final}/{exp} belge yüklendi ✓")
            except Exception as e:
                failed.add(fold(name)); log(f"!! {name} → Dosya Yükle hatası ({str(e)[:60]}), atlandı"); close_form(); continue
            wait_ajax(page)
            # 🧪 SADECE-YÜKLE test modu: belgeler yüklendi → ONAYA GÖNDERME (geri-alınamaz onayı atla, taslak kalsın)
            if args.sadece_yukle:
                ok += 1
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
            for _v in range(6):
                if not list_open():
                    open_list()
                else:
                    reset_list_filter()
                listele()
                durum = page.evaluate(STATUS_OF, fold(name))
                if durum and "Gönderilmemiş" not in durum:
                    break
                page.wait_for_timeout(600)
            if not (durum and "Gönderilmemiş" not in durum):
                failed.add(fold(name))
                log(f"!! {name} → ONAYA GİTMEDİ (durum: {durum or 'okunamadı/taslak'}); elle bak")
                continue
            ok += 1
            with open(onay_file, "a", encoding="utf-8") as f:
                f.write(belge_ad + "\n")   # Excel/DGS adı → resume+onaylanan `our` (Excel isimleri) ile eşleşsin
            log(f"[{ok}] {name} → ONAYA GÖNDERİLDİ ✓ (gerçek durum: {durum})")

        log(f"\n==== ÖZET: bu koşuda {ok} kişi onaya gönderildi ====")
        if failed:
            log(f"  !! ONAYA GİTMEYEN ({len(failed)}): elle bak.")
        try:
            input(f"{LOG} >> Bitti. ENTER...")
        except EOFError:
            pass


if __name__ == "__main__":
    main()
