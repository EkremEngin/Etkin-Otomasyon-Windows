#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
İZİN OTOMASYON — tek-dosya format + %100 mutabakat (giriş + onaya gönder)
=========================================================================
Firmadan gelen yeni standart izin Excel'ini (`izin_data_v2`) okur; kanıtlı giriş motorunu
(`izin_poc`) ve onay motorunu (`izin_onaya`) **DEĞİŞTİRMEDEN import ederek** her parka uygular.

"KİŞİ ATLAMA" BUG'INA KARŞI ÜÇ KATMAN:
  1) Veri katmanı: isim-eşleştirme YOK (eski atlama kaynağı). Her satırda T.C.; park F kolonunda.
     Bozuk satır varsa portala DOKUNULMADAN durur (izin_data_v2.DataError).
  2) Portal-park kilidi: açık Chrome sekmesinin URL'si işlenen parkın portalıyla eşleşmezse DURUR
     (yanlış portala giriş imkânsız — eski sistemde bu kontrol yoktu).
  3) MUTABAKAT: park bitince beklenen kişi == (girildi + zaten-girili + tatil/izin-var + AÇIK-flag)
     eşitliği DENETLENİR. Biri bile "sebepsiz" düşerse GÜRÜLTÜLÜ ALARM verir. Sessiz atlama imkânsız.

ÇALIŞMA MODELİ (ATTENDED — eski sistemle aynı):
  Chrome'u `--remote-debugging-port` ile aç, işlemek istediğin PARKIN portalına gir, login + Cloudflare'i
  bir kez ELLE geç, sonra bu betiği çalıştır. Betik açık sekmenin URL'sinden parkı otomatik algılar.

Kullanım:
  python3 izin_otomasyon.py --excel "~/Downloads/Haziran 2026 Yıllık İzinler.xlsx" --plan
      → portalsız ön-uçuş: hangi park kaç kişi, kim girilecek (hiçbir şeye dokunmaz).
  python3 izin_otomasyon.py --excel "<dosya>"                 # açık portalın parkı, DRY-RUN (kaydetmez)
  python3 izin_otomasyon.py --excel "<dosya>" --commit        # TASLAK gir + mutabakat
  python3 izin_otomasyon.py --excel "<dosya>" --commit --onayla   # gir + onaya gönder (PDF'siz parklar)
  python3 izin_otomasyon.py --excel "<dosya>" --park TPI --person "AHMET" --commit   # tek kişi / retry
"""
from __future__ import annotations
import argparse
import datetime
import json
import os
import subprocess
import sys
import time

import izin_data_v2 as data
import izin_belge
from izin_data_v2 import Park, PARKS, resolve_park, fold

LOG = "[İZİN-OTO]"


def log(*a):
    print(LOG, *a, flush=True)


def _runlog(park, meta, commit, entry):
    """Koşuyu kişi-bazında JSONL'e yaz (kalıcı iz — 'ne döndü' sonradan analiz edilsin). Best-effort, akışı BOZMAZ."""
    try:
        rec = {"ts": datetime.datetime.now().isoformat(timespec="seconds"),
               "park": park.code, "donem": meta.get("donem_label"), "commit": commit}
        rec.update(entry)
        with open(f"izin_runlog_{park.label}_{meta['ay_key']}.jsonl", "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False, default=str) + "\n")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Portal ↔ park eşleştirme (güvenlik kilidi)
# ---------------------------------------------------------------------------
def _domain(url: str) -> str:
    u = (url or "").split("//", 1)[-1]
    return u.split("/", 1)[0].lower()


def park_from_url(url: str) -> Park | None:
    """Açık sekmenin URL'sinden parkı bul (portal domain eşleşmesi)."""
    d = _domain(url)
    if not d:
        return None
    for p in PARKS.values():
        if _domain(p.portal_url) and _domain(p.portal_url) in d:
            return p
    return None


# ---------------------------------------------------------------------------
# izin_poc motorunu bu park + bu dönem için ayarla (CONFIG'i çalışma-anında yaz)
# ---------------------------------------------------------------------------
def configure_engine(izin_poc, park: Park, meta: dict, commit: bool, cdp: str):
    izin_poc.CONFIG["donem_text"] = meta["donem_label"]     # "HAZİRAN 2026"
    izin_poc.CONFIG["ay_regex"] = meta["ay_regex"]          # ^\d{2}\.06\.2026$
    izin_poc.CONFIG["portal_url"] = park.portal_url
    izin_poc.CONFIG["dry_run"] = not commit
    if cdp:
        izin_poc.CONFIG["cdp_url"] = cdp


# ---------------------------------------------------------------------------
# Tek parkın girişi + MUTABAKAT
# ---------------------------------------------------------------------------
def run_park_entry(izin_poc, page, park: Park, targets, meta: dict, commit: bool,
                   limit: int, only_person: str | None):
    """Parkın hedeflerini gir (TASLAK). Dönüş: (results, done_names_now, recon)."""
    ay_key = meta["ay_key"]
    done_file = f"izin_done_{park.label}_{ay_key}.txt"
    done = set()
    if os.path.exists(done_file):
        # "TC\tAd" veya düz "Ad" satırları — ikisini de destekle
        for line in open(done_file, encoding="utf-8"):
            line = line.strip()
            if not line:
                continue
            done.add(line.split("\t")[-1])              # ad kısmı (izin_onaya uyumu)

    work = list(targets)
    if only_person:
        key = fold(only_person)
        work = [p for p in work if key in fold(p.ad) or key in p.tc]
        if not work:
            log(f"HATA: '{only_person}' {park.code} listesinde yok.")
            return [], [], None
    if limit:
        work = work[:limit]

    log(f"\n{'='*70}\n### PARK: {park.code} ({park.ad}) — {len(targets)} kişi | dönem {meta['donem_label']} "
        f"| commit={commit} (False=DRY-RUN)\n{'='*70}")
    if done:
        log(f"{len(done)} kişi zaten girilmiş (resume, atlanacak).")
    _runlog(park, meta, commit, {"olay": "run_start", "hedef": len(work),
                                 "resume_atlanacak": len(done), "toplam_park_kisi": len(targets)})

    results = []
    done_now = []
    first_done_before = bool(done)  # form ilk açılışını yönetmek için
    first = True
    for i, person in enumerate(work, 1):
        if person.ad in done:
            results.append({"tc": person.tc, "ad": person.ad, "durum": "zaten_girili",
                            "ok": True, "kaydedildi": False, "flag": [], "mesaj": "resume: zaten girilmiş"})
            _runlog(park, meta, commit, {"olay": "kisi", "i": i, "ad": person.ad, "tc": person.tc,
                                         "durum": "zaten_girili", "mesaj": "resume: atlandı"})
            continue
        log(f"\n--- [{i}/{len(work)}] {person.ad}  (T.C. {person.tc}) | Excel: {person.gunler} ---")
        try:
            izin_poc.assert_logged_in(page)
            r = izin_poc.process_one_person(page, person, first)
            first = False
        except izin_poc.CloudflareHalt as e:
            log(f"!! DURDU (Cloudflare/oturum): {e}")
            results.append({"tc": person.tc, "ad": person.ad, "durum": "halt",
                            "ok": False, "kaydedildi": False, "flag": [], "mesaj": f"CloudflareHalt: {e}"})
            break
        except izin_poc.VerifyError as e:
            first = False
            shot = _shot(page, "izin_verify")
            r = {"ad": person.ad, "ok": False, "kaydedildi": False,
                 "mesaj": f"DOĞRULAMA HATASI (KAYDEDİLMEDİ): {e} (ss:{shot})", "flag": []}
        except Exception as e:  # noqa
            first = False
            shot = _shot(page, "izin_hata")
            r = {"ad": person.ad, "ok": False, "kaydedildi": False,
                 "mesaj": f"HATA (KAYDEDİLMEDİ): {e} (ss:{shot})", "flag": []}
        r.setdefault("tc", person.tc)
        r["durum"] = _classify(r)
        results.append(r)
        log(f"  -> {r['mesaj']}")
        _runlog(park, meta, commit, {"olay": "kisi", "i": i, "ad": person.ad, "tc": person.tc,
                                     "durum": r["durum"], "kaydedildi": bool(r.get("kaydedildi")),
                                     "mesaj": r.get("mesaj"), "girilen": r.get("girilen"),
                                     "flag": r.get("flag"), "excel_gunler": person.gunler})
        if r.get("kaydedildi") and commit:
            with open(done_file, "a", encoding="utf-8") as f:
                f.write(person.ad + "\n")     # isim-only (izin_onaya + dosyali uploader whitelist uyumu)
            done_now.append(person.ad)

    recon = reconcile(work, results, park)
    if recon:
        _runlog(park, meta, commit, {"olay": "run_end",
                                     "mutabakat": {k: len(recon[k]) for k in
                                                   ("girildi", "dry", "zaten", "yeni_yok", "flag", "flagli",
                                                    "basarisiz", "halt", "kayip") if isinstance(recon.get(k), list)},
                                     "tam_mutabakat": recon.get("tam_mutabakat")})
    return results, done_now, recon


def _classify(r: dict) -> str:
    if not r.get("ok"):
        return "BAŞARISIZ"
    if r.get("kaydedildi"):
        return "girildi"
    if r.get("girilen"):        # gün(ler) işlendi ama kaydedilmedi = DRY-RUN önizleme (commit'te kaydolacak)
        return "dry_girilecek"
    if r.get("flag"):
        return "FLAG"
    return "yeni_giris_yok"      # girilecek yeni gün yok (hepsi zaten-girili/tatil)


def _shot(page, prefix):
    ts = int(time.time())
    path = f"{prefix}_{ts}.png"
    try:
        page.screenshot(path=path, full_page=True)
    except Exception:
        pass
    return path


def reconcile(work, results, park: Park) -> dict:
    """MUTABAKAT: her hedef kişi bir sonuç kovasına düşmeli. 'Kayıp' (hiç sonuç üretmeyen) = ALARM."""
    by_tc = {r["tc"]: r for r in results}
    girildi = [r for r in results if r["durum"] == "girildi"]
    dry = [r for r in results if r["durum"] == "dry_girilecek"]
    zaten = [r for r in results if r["durum"] == "zaten_girili"]
    yeni_yok = [r for r in results if r["durum"] == "yeni_giris_yok"]
    flag = [r for r in results if r["durum"] == "FLAG"]
    basarisiz = [r for r in results if r["durum"] == "BAŞARISIZ"]
    halt = [r for r in results if r["durum"] == "halt"]
    kayip = [p for p in work if p.tc not in by_tc]       # sonucu OLMAYAN hedef = sessiz atlama adayı
    flagli = [r for r in results if r.get("flag")]       # kovadan bağımsız, flag'i olan herkes (uyarı için)
    return {
        "park": park.code, "beklenen": len(work),
        "girildi": girildi, "dry": dry, "zaten": zaten, "yeni_yok": yeni_yok,
        "flag": flag, "flagli": flagli, "basarisiz": basarisiz, "halt": halt, "kayip": kayip,
        "tam_mutabakat": (len(kayip) == 0 and len(basarisiz) == 0 and len(halt) == 0 and len(flagli) == 0),
    }


def print_recon(recon: dict):
    if not recon:
        return
    log(f"\n{'─'*70}")
    log(f"MUTABAKAT — {recon['park']}: beklenen {recon['beklenen']} kişi")
    if recon["girildi"]:
        log(f"   ✅ girildi (kaydedildi): {len(recon['girildi'])}")
    if recon["dry"]:
        log(f"   👁️  girilecek (DRY-RUN): {len(recon['dry'])}  (gün(ler) işlendi, --commit yok → kaydedilmedi; commit'te girilecek)")
    log(f"   ⏩ zaten girili       : {len(recon['zaten'])}")
    log(f"   ➖ yeni giriş yok      : {len(recon['yeni_yok'])}  (hepsi tatil/izin-var; kaydedilecek yeni gün yok)")
    if recon["flagli"]:
        log(f"   ⚠️  FLAG (elle bak)    : {len(recon['flagli'])}")
        for r in recon["flagli"]:
            log(f"        - {r['ad']} (T.C.{r['tc']}): {'; '.join(r.get('flag', []))}")
    if recon["basarisiz"]:
        log(f"   ❌ BAŞARISIZ          : {len(recon['basarisiz'])}")
        for r in recon["basarisiz"]:
            log(f"        - {r['ad']} (T.C.{r['tc']}): {r['mesaj']}")
    if recon["halt"]:
        log(f"   🛑 DURDURULDU         : {len(recon['halt'])} (Cloudflare/oturum)")
    if recon["kayip"]:
        log(f"   🔴🔴 KAYIP (sonuç YOK — SESSİZ ATLAMA ALARMI): {len(recon['kayip'])}")
        for p in recon["kayip"]:
            log(f"        - {p.ad} (T.C.{p.tc})  ← bu kişi hiç işlenmedi, SEBEBİNİ BUL")
    toplam = (len(recon['girildi']) + len(recon['dry']) + len(recon['zaten']) + len(recon['yeni_yok'])
              + len(recon['flag']) + len(recon['basarisiz']) + len(recon['halt']) + len(recon['kayip']))
    log(f"   ── toplam hesaplanan: {toplam} / beklenen {recon['beklenen']} "
        f"→ {'✅ TAM MUTABAKAT' if toplam == recon['beklenen'] and not recon['kayip'] else '🔴 EKSİK — İNCELE'}")
    if recon["tam_mutabakat"]:
        log(f"   ✅✅ {recon['park']}: herkes hesaba katıldı, sessiz atlama YOK.")


# ---------------------------------------------------------------------------
# Onaya gönder (park-bağımlı PDF)
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def parse_parklar(s: str | None) -> "list[Park] | None":
    """'TPI,BV,ULUTEK' → sıralı Park listesi (yalnız otomasyon parkları). None → filtre yok.
    Bilinmeyen/Teknoera kod → hata. Kullanıcı: 'hangi parklar koşacak, sırayla işaretle'."""
    if not s:
        return None
    out = []
    for tok in s.split(","):
        tok = tok.strip()
        if not tok:
            continue
        pk = resolve_park(tok)
        if pk is None:
            raise data.DataError(f"--parklar: bilinmeyen park {tok!r}. Geçerli: {', '.join(c for c in PARKS if PARKS[c].otomasyon)}")
        if not pk.otomasyon:
            raise data.DataError(f"--parklar: {pk.code} Teknoera (excel upload) — otomasyon kapsamı dışı, seçilemez.")
        if pk.code not in [p.code for p in out]:
            out.append(pk)
    return out


def run_approval_pdfsiz(page, park: Park, meta: dict, limit: int, people=None):
    """PDF'siz parklar (TPI/BV/ULUTEK): kanıtlı izin_onaya akışı — AYNI page üzerinde (inline).
    KİMLİK: onay da maskeli-T.C. ile eşleşir (isim evlilik/kızlık farkında taslağı bulamıyordu — 2026-07-09)."""
    ay_key = meta["ay_key"]
    done_file = f"izin_done_{park.label}_{ay_key}.txt"
    if not os.path.exists(done_file):
        log(f"[ONAY] {park.code}: {done_file} yok — önce giriş (--commit) gerekli. Atlanıyor.")
        return
    import izin_onaya
    our = set(fold(line.strip().split("\t")[-1]) for line in open(done_file, encoding="utf-8") if line.strip())
    our_tc = set()                                   # maskeli-T.C. anahtarı (ilk2+son2) — isimden bağımsız kimlik
    for p in (people or []):
        tc = str(getattr(p, "tc", "") or "").strip()
        if tc.isdigit() and len(tc) == 11:
            our_tc.add(tc[:2] + tc[-2:])
    onay_file = f"izin_onaya_done_{park.label}_{ay_key}.txt"
    onaylanan = set()
    if os.path.exists(onay_file):
        onaylanan = set(fold(l) for l in open(onay_file, encoding="utf-8").read().splitlines() if l.strip())
    log(f"\n[ONAY] {park.code}: {len(our)} taslak, {len(onaylanan)} zaten gönderilmiş → onaya gönderiliyor (PDF'siz).")
    ok, failed = izin_onaya.approve_drafts(page, meta["donem_label"], our, onay_file,
                                           limit=(limit or 10**9), already=onaylanan, log=log, our_tc=our_tc)
    log(f"[ONAY] {park.code}: bu koşuda {ok} onaya gönderildi; {len(failed)} gitmedi.")
    _runlog(park, meta, True, {"olay": "onay", "gonderildi": ok, "gitmedi": len(failed),
                               "gitmedi_kisiler": sorted(failed)})


def run_approval_pdfli(park: Park, meta: dict, people, belge_klasor: str | None, cdp: str, limit: int,
                       ortak_belge: str | None = None):
    """PDF'li parklar: belgeleri çöz → HEPSİ hazırsa kanıtlı yükleyiciyi (subprocess) çalıştır.
    İKİ MOD (park.onay_pdf):
      • per_person (YTP/ULUTEK) → `--dosya-klasor <folder>`: yükleyici her kişinin belgesini adıyla çözer.
      • ortak      (İYTE)       → `--dosya <ortak.pdf>` + boş `--dosya-klasor`: kişi-başı çözüm KAPALI,
        yükleyicinin global-tek-dosya fallback'i herkese aynı belgeyi yükler. (Klasörü geçmiyoruz ki
        oraya bırakılmış bir "Ad Soyad.pdf" sessizce birine farklı belge yüklemesin.)
    ADAPTİF: belge eksikse onaya GÖNDERMEZ (giriş taslak kalır), net rapor verir — PDF'li parkları beklemeden
    PDF'siz parklar koşulabilsin diye. Yükleyici AYRI process (kendi CDP oturumu) → eş-zamanlı bağlantı yok."""
    ay_key = meta["ay_key"]
    done_file = f"izin_done_{park.label}_{ay_key}.txt"
    if not os.path.exists(done_file):
        log(f"[ONAY] {park.code}: {done_file} yok — önce giriş (--commit) gerekli. Atlanıyor.")
        return
    folder = izin_belge.find_belge_klasor(park, base=belge_klasor)
    r = izin_belge.report(park, people, folder=folder, ortak_belge=ortak_belge)
    if r["eslesen"] != r["toplam"] or r["nonpdf"]:
        log(f"[ONAY] {park.code}: belgeler TAM DEĞİL ({r['eslesen']}/{r['toplam']}"
            f"{', PDF-dışı var' if r['nonpdf'] else ''}{', ' + r['hata'] if r['hata'] else ''}) "
            f"→ onaya GÖNDERİLMEDİ (adaptif). Giriş TASLAK kaldı; eksiği tamamlayıp tekrar `--onayla` çalıştır.")
        return
    script = os.path.join(SCRIPT_DIR, "izin_onaya_dosyali_yildiz.py")
    cmd = [sys.executable, script, "--lokasyon", park.label, "--sheet", ay_key,
           "--donem", meta["donem_label"]]
    if r["mod"] == izin_belge.ORTAK:
        cmd += ["--dosya", r["ortak"], "--dosya-klasor", ""]
        nasil = f"ORTAK belge → {r['toplam']} kişinin hepsine '{os.path.basename(r['ortak'])}'"
    else:
        cmd += ["--dosya-klasor", folder]
        nasil = f"{r['toplam']} kişi-başı belge hazır ✓"
    if limit:
        cmd += ["--limit", str(limit)]
    env = os.environ.copy()
    env["DGS_CDP"] = cdp
    log(f"\n[ONAY] {park.code}: {nasil} → yükle+onaya gönder (makinalı tüfek):")
    log("   $ " + " ".join(f'"{c}"' if (" " in c or not c) else c for c in cmd))
    subprocess.run(cmd, stdin=subprocess.DEVNULL, env=env, cwd=SCRIPT_DIR)


# ---------------------------------------------------------------------------
# Plan modu (portalsız ön-uçuş — offline test edilebilir)
# ---------------------------------------------------------------------------
def run_plan(excel: str, parklar, belge_klasor: str | None, detay: bool = True,
             ortak_belge: str | None = None):
    by_park, meta = data.read_izin_v2(excel, strict=True)
    print("=" * 74)
    print(f"ÖN-UÇUŞ PLANI — {os.path.basename(os.path.expanduser(excel))}")
    print(f"Dönem: {meta['donem_label']} | done-dosya eki: _{meta['ay_key']}.txt")
    print("=" * 74)
    # sıralı kapsam: --parklar verildiyse o sıra, yoksa tüm otomasyon parkları
    if parklar:
        order = [p.code for p in parklar if p.code in by_park]
    else:
        order = [c for c in by_park if PARKS[c].otomasyon]
    print(f"Koşulacak parklar (sıra): {' → '.join(order) if order else '(yok)'}")
    for code in order:
        pk = PARKS[code]
        people = by_park[code]
        if pk.onay_pdf:
            folder = izin_belge.find_belge_klasor(pk, base=belge_klasor)
            rd = izin_belge.readiness(people, folder, mode=pk.onay_pdf, ortak_belge=ortak_belge)
            hazir = rd["eslesen"] == rd["toplam"] and not rd["nonpdf"]
            if pk.onay_pdf == izin_belge.ORTAK:
                durum = (f"ORTAK belge '{os.path.basename(rd['ortak'])}' HAZIR ✓" if hazir
                         else f"ORTAK belge ⚠ {rd['hata'] or 'PDF-dışı dosya var'}")
            else:
                durum = f"kişi-başı belge {rd['eslesen']}/{rd['toplam']}" + (" HAZIR ✓" if hazir else
                         (" ⚠ klasör YOK" if not folder else " ⚠ EKSİK"))
            pdf = f"  [ONAY: {durum}]"
        else:
            pdf = "  [ONAY: PDF'siz, direkt]"
        print(f"\n### {pk.code} — {len(people)} kişi{pdf}  → {pk.portal_url}")
        if detay:
            for p in people:
                print(f"    {p.ad:32s} T.C.{p.tc} | {p.toplam_gun:>4}g | {[d for d,_ in p.gunler]}")
    if parklar:
        atlanan = [c for c in by_park if PARKS[c].otomasyon and c not in order]
        if atlanan:
            n = sum(len(by_park[c]) for c in atlanan)
            print(f"\n⏭️  Seçilmedi → ATLANACAK: {', '.join(atlanan)} — {n} kişi")
    tekno = [c for c in by_park if not PARKS[c].otomasyon]
    if tekno:
        n = sum(len(by_park[c]) for c in tekno)
        print(f"⏭️  Teknoera (kapsam dışı, excel upload): {', '.join(tekno)} — {n} kişi")
    print("\n" + "=" * 74)
    print("Seçilen parklardaki HERKES işlenecek — sessiz atlama yok. PDF parkları belge hazırsa onaya gider,")
    print("değilse giriş TASLAK kalır (adaptif). Canlı çalıştırmak için --plan'ı kaldır.")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="İzin Otomasyon v2 — tek-dosya format, %100 mutabakat, adaptif")
    ap.add_argument("--excel", required=True, help="İzin Excel yolu (yeni tek-dosya format)")
    ap.add_argument("--parklar", default=None,
                    help="Koşulacak parklar, sıralı (ör. 'TPI,BV,ULUTEK'). Yoksa açık portalın parkı işlenir. "
                         "Seçilmeyen parklar ATLANIR (PDF gerektirenleri beklemek istemiyorsan seçme).")
    ap.add_argument("--park", default=None, help="Tek parkı zorla/doğrula (yoksa açık portaldan otomatik algılanır)")
    ap.add_argument("--belge-klasor", default=None,
                    help="PDF'li park için belge klasörü. Yoksa Masaüstü konvansiyonu (~/Desktop/İzin Belgeleri/<Park> İzin Belgeleri) aranır.")
    ap.add_argument("--ortak-belge", default=None,
                    help="ORTAK modlu parkta (İYTE) toplu belgenin yolu — klasör keşfini baypas eder. "
                         "Aynı PDF o parktaki HERKESİN kaydına yüklenir.")
    ap.add_argument("--commit", action="store_true", help="TASLAK kaydet (yoksa DRY-RUN)")
    ap.add_argument("--onayla", action="store_true", help="Girişten sonra onaya gönder (PDF'li parkta belgeler hazırsa yükle+gönder)")
    ap.add_argument("--person", default=None, help="Sadece tek kişi (ad veya T.C. ile) — retry için")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--cdp", default=None, help="CDP adresi (ör. http://localhost:9223); yoksa DGS_CDP/9222")
    ap.add_argument("--plan", action="store_true", help="Portalsız ön-uçuş planı (hiçbir şeye dokunmaz)")
    args = ap.parse_args()

    excel = os.path.expanduser(args.excel)
    if not os.path.exists(excel):
        log(f"HATA: Excel bulunamadı: {excel}"); sys.exit(1)
    try:
        parklar = parse_parklar(args.parklar)
    except data.DataError as e:
        log(f"❌ {e}"); sys.exit(2)

    # --- Veri: bozuksa BURADA durur (portala dokunmadan) ---
    try:
        by_park, meta = data.read_izin_v2(excel, strict=True)
    except data.DataError as e:
        log(f"❌ VERİ HATASI — portala dokunulmadı:\n{e}"); sys.exit(2)
    log(f"Veri OK: {meta['toplam_kisi']} kişi, dönem {meta['donem_label']}. Otomasyon parkları: "
        f"{', '.join(c for c in by_park if PARKS[c].otomasyon)}")

    if args.plan:
        run_plan(excel, parklar, args.belge_klasor, ortak_belge=args.ortak_belge)
        return

    # --- Canlı: Playwright motorunu şimdi import et (plan modunda gereksiz) ---
    import izin_poc
    from playwright.sync_api import sync_playwright

    active = None
    targets = []
    cdp = args.cdp or izin_poc.CONFIG.get("cdp_url")
    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp(cdp)
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
        page = next((p for p in ctx.pages if "argeportal" in (p.url or "")), None)
        if page is None:
            log("HATA: Açık bir Arge Portal sekmesi yok. Chrome'da işlemek istediğin PARKIN portalına "
                "gir, login + Cloudflare'i geç, tekrar çalıştır."); sys.exit(1)
        page.bring_to_front()

        # PARK KİLİDİ: açık portalın parkı (yanlış portala giriş imkânsız)
        active = park_from_url(page.url)
        if active is None:
            log(f"HATA: Açık sekme ({_domain(page.url)}) bilinen bir Arge Portal değil. Doğru portala gir."); sys.exit(1)
        if args.park:
            forced = resolve_park(args.park)
            if forced is None or forced.code != active.code:
                log(f"🔴 PARK UYUŞMAZLIĞI: --park={args.park} ({forced.code if forced else '?'}) ama açık portal "
                    f"{active.code} ({active.portal_url}). Yanlış portala giriş önlendi — DURDU."); sys.exit(3)
        if parklar and active.code not in [p.code for p in parklar]:
            log(f"🔴 SEÇİM DIŞI: açık portal {active.code} ama --parklar={args.parklar}. "
                f"Ya listeye ekle ya da seçili bir parkın portalına geç — DURDU (yanlışlıkla işlem yok)."); sys.exit(3)
        if not active.otomasyon:
            log(f"HATA: {active.code} Teknoera (excel upload) — Playwright kapsamı dışı."); sys.exit(1)

        targets = by_park.get(active.code, [])
        if not targets:
            log(f"{active.code} için bu Excel'de kişi yok."); sys.exit(0)

        configure_engine(izin_poc, active, meta, args.commit, cdp)
        try:
            izin_poc.assert_logged_in(page)
        except izin_poc.CloudflareHalt as e:
            log(f"🛑 {e}"); sys.exit(1)

        results, done_now, recon = run_park_entry(izin_poc, page, active, targets, meta,
                                                   args.commit, args.limit, args.person)
        print_recon(recon)

        # PDF'siz onay AYNI oturumda (inline). PDF'li onay oturum kapandıktan SONRA (subprocess, eş-zamanlı bağlantı yok).
        if args.onayla and not args.commit:
            log("[ONAY] --onayla verildi ama --commit yok → DRY-RUN'da onay yapılmaz.")
        elif args.onayla and args.commit and not active.onay_pdf:
            run_approval_pdfsiz(page, active, meta, args.limit, targets)

    # oturum kapandı → PDF'li park onayı (belge çöz + kanıtlı yükleyici subprocess)
    if args.onayla and args.commit and active and active.otomasyon and active.onay_pdf:
        run_approval_pdfli(active, meta, targets, args.belge_klasor, cdp, args.limit,
                           ortak_belge=args.ortak_belge)

    log("\nHATIRLATMA: Giriş TASLAK. PDF'siz parklar --onayla ile onaya gitti; PDF'li parklar belge hazırsa gitti, "
        "değilse taslak kaldı (adaptif).")
    try:
        input(f"{LOG} >> Bitti. ENTER ile kapat...")
    except EOFError:
        pass


if __name__ == "__main__":
    main()
