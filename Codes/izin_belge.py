#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
İzin belgesi çözücü — kişi ↔ PDF eşleştirme + klasör keşfi (adaptif onay için).
==============================================================================
Onayı PDF gerektiren parklarda belgeyi bulur ve kişiyle eşleştirir. İKİ BELGE MODU var (`Park.onay_pdf`):

  • "per_person" (YTP, ULUTEK) — HER personelin KENDİ izin formu. Kişi başına 1+ dosya.
  • "ortak"      (İYTE)        — firma park için TEK antetli/kaşeli toplu belge yollar (herkesi listeler,
                                 tek imza/kaşe). Portal kişi başına dosya istediği için AYNI PDF her kişinin
                                 kaydına yüklenir. Klasörde TAM 1 PDF beklenir (0 → eksik, >1 → belirsiz).

Kaynak iki yoldan gelir (ikisi de sadece bir PATH — exe'nin "dosya yükle" arayüzü de aynı path'i verecek):
  1) MASAÜSTÜ KONVANSİYONU (otomatik keşif): `~/Desktop/İzin Belgeleri/<Park> İzin Belgeleri/`
     ör. `~/Desktop/İzin Belgeleri/İYTE İzin Belgeleri/`  ·  `.../YTP İzin Belgeleri/`
  2) Açık verilen: `--belge-klasor /yol/...` (per_person)  ·  `--ortak-belge /yol/belge.pdf` (ortak)
  → exe UI karşılığı: per_person parkta "klasör seç", ortak parkta "tek belge seç".

EŞLEŞTİRME — per_person (iki güvenli yol, ikisi de yanlış-pozitif vermez):
  a) Dosya adında kişinin 11 haneli T.C.'si geçiyorsa → o kişinin belgesi (EN GÜÇLÜ; T.C. tekil).
  b) Dosya adı (çekirdeği) kişinin adıyla TAM fold-eşleşiyorsa → o kişinin belgesi.
     Çekirdek = uzantı + " (n)" tekrar-eki + varsa sondaki T.C. çıkarıldıktan sonrası ("AD SOYAD (2).pdf" → "AD SOYAD").
Bir kişinin >1 belgesi olabilir (n=1,2,3 sırasıyla). Portal SADECE PDF kabul eder (memory) → PDF-dışı UYARILIR.

Bu modül PORTALDAN BAĞIMSIZDIR (Playwright yok) → tek başına test edilebilir:
  python3 izin_belge.py "~/Downloads/Haziran 2026 Yıllık İzinler.xlsx" --park İYTE
  python3 izin_belge.py "<xlsx>" --park YTP --belge-klasor "/yol/YTP Belgeleri"
  python3 izin_belge.py "<xlsx>" --park İYTE --ortak-belge "/yol/Antetli Liste.pdf"
"""
from __future__ import annotations
import argparse
import os
import re
import sys

from izin_data_v2 import fold, PARKS, resolve_park, read_izin_v2

_NUM_SUFFIX = re.compile(r"\s*\((\d+)\)\s*$")          # " (1)", " (2)" tekrar eki
_TC_IN_NAME = re.compile(r"(?<!\d)(\d{11})(?!\d)")     # dosya adında 11 haneli T.C.
_PDF_EXT = {".pdf"}
DEFAULT_BASE = os.path.expanduser("~/Desktop/İzin Belgeleri")
LOG = "[BELGE]"

PER_PERSON = "per_person"   # kişi başına kendi belgesi
ORTAK = "ortak"             # park için tek toplu belge (aynı PDF herkese)


def _clean_stem(stem: str) -> str:
    """Dosya adı çekirdeği: sondaki '(n)' tekrar-eki + varsa sondaki T.C. çıkar."""
    s = _NUM_SUFFIX.sub("", stem).strip()
    s = _TC_IN_NAME.sub("", s).strip(" -_")
    return s


def _list_files(folder: str) -> list[str]:
    if not folder or not os.path.isdir(folder):
        return []
    out = []
    for f in sorted(os.listdir(folder)):
        if f.startswith(".") or f.startswith("_"):
            continue
        if os.path.isfile(os.path.join(folder, f)):
            out.append(f)
    return out


def resolve_person_files(folder: str, ad: str, tc: str = "") -> list[str]:
    """Bu kişiye ait belge(ler)in tam yolu (sıralı). T.C. dosya adında ise onunla, yoksa tam-isim ile.
    Eşleşme yoksa boş liste. (yildiz `resolve_person_files` mantığının T.C.-farkındalıklı üst-kümesi.)"""
    if not folder or not os.path.isdir(folder):
        return []
    nf = fold(ad)
    tc = str(tc).strip()
    found: list[tuple[int, str]] = []
    for f in _list_files(folder):
        stem, ext = os.path.splitext(f)
        p = os.path.join(folder, f)
        m = _NUM_SUFFIX.search(stem)
        num = int(m.group(1)) if m else 0
        tcs = _TC_IN_NAME.findall(stem)
        if tc and tc in tcs:                          # (a) T.C. eşleşmesi — en güçlü
            found.append((num, p)); continue
        if fold(_clean_stem(stem)) == nf:             # (b) tam isim eşleşmesi
            found.append((num, p))
    found.sort()
    return [p for _, p in found]


def _pdfs(folder: str | None) -> list[str]:
    return [f for f in _list_files(folder) if os.path.splitext(f)[1].lower() in _PDF_EXT] if folder else []


def resolve_ortak_belge(folder: str | None, explicit: str | None = None) -> tuple[str | None, str]:
    """ORTAK mod: parkın TEK toplu belgesini çöz → (yol, hata). Hata boşsa yol geçerli.
    `explicit` (--ortak-belge) verilirse klasör hiç okunmaz — exe'nin "tek belge seç" düğmesinin karşılığı.
    Klasörden çözerken TAM 1 PDF şart: 0 → eksik, >1 → BELİRSİZ (sessizce yanlışını seçmektense DUR)."""
    if explicit:
        p = os.path.expanduser(explicit)
        if not os.path.isfile(p):
            return None, f"ortak belge bulunamadı: {p}"
        if os.path.splitext(p)[1].lower() not in _PDF_EXT:
            return None, f"ortak belge PDF değil (portal reddeder): {os.path.basename(p)}"
        return p, ""
    if not folder or not os.path.isdir(folder):
        return None, "belge klasörü yok"
    pdfs = _pdfs(folder)
    if not pdfs:
        return None, "klasörde PDF yok"
    if len(pdfs) > 1:
        return None, (f"klasörde {len(pdfs)} PDF var, hangisi ortak belge belirsiz "
                      f"({', '.join(pdfs[:4])}{'…' if len(pdfs) > 4 else ''}) → --ortak-belge ile seç")
    return os.path.join(folder, pdfs[0]), ""


def find_belge_klasor(park, base: str | None = None) -> str | None:
    """Masaüstü konvansiyonuyla parkın belge klasörünü keşfet. Bulamazsa None.
    `base` altındaki alt-klasörleri fold ile eşler (yazım/diakritik toleranslı)."""
    base = base or DEFAULT_BASE
    base = os.path.expanduser(base)
    # base'in kendisi de farklı yazılmış olabilir → üst dizinde fold ile ara
    if not os.path.isdir(base):
        parent = os.path.dirname(base)
        target = fold(os.path.basename(base))
        if os.path.isdir(parent):
            for e in os.listdir(parent):
                if os.path.isdir(os.path.join(parent, e)) and fold(e) == target:
                    base = os.path.join(parent, e); break
    if not os.path.isdir(base):
        return None
    wanted = {fold(x) for x in (
        f"{park.code} İzin Belgeleri", f"{park.label} İzin Belgeleri", f"{park.ad} İzin Belgeleri",
        f"{park.code} İzin Formları", f"{park.label} İzin Formları",
        park.code, park.label, park.ad)}
    for e in sorted(os.listdir(base)):
        full = os.path.join(base, e)
        if os.path.isdir(full) and fold(e) in wanted:
            return full
    # alt-klasör yoksa base'in kendisi PDF içeriyorsa onu kullan (düz klasör senaryosu)
    if any(os.path.splitext(f)[1].lower() in _PDF_EXT for f in _list_files(base)):
        return base
    return None


def readiness(people, folder: str | None, mode: str = PER_PERSON, ortak_belge: str | None = None) -> dict:
    """Kişi listesi + klasör → {matched{tc:[dosya]}, missing[person], extra[dosya], nonpdf[dosya], folder, mod, ortak, hata}.
    mode=ORTAK → tek toplu belge TÜM kişilere eşleşir (çözülemezse herkes 'missing', sebep `hata`da)."""
    all_files = _list_files(folder) if folder else []
    nonpdf = [f for f in all_files if os.path.splitext(f)[1].lower() not in _PDF_EXT]

    if mode == ORTAK:
        shared, hata = resolve_ortak_belge(folder, ortak_belge)
        return {"matched": ({p.tc: [shared] for p in people} if shared else {}),
                "missing": ([] if shared else list(people)), "extra": [], "nonpdf": nonpdf,
                "folder": folder, "toplam": len(people), "eslesen": (len(people) if shared else 0),
                "mod": ORTAK, "ortak": shared, "hata": hata}

    matched, missing = {}, []
    used = set()
    for p in people:
        files = resolve_person_files(folder, p.ad, p.tc) if folder else []
        if files:
            matched[p.tc] = files
            used.update(files)
        else:
            missing.append(p)
    extra = [f for f in all_files if os.path.join(folder, f) not in used]
    return {"matched": matched, "missing": missing, "extra": extra, "nonpdf": nonpdf,
            "folder": folder, "toplam": len(people), "eslesen": len(matched),
            "mod": PER_PERSON, "ortak": None, "hata": ""}


def report(park, people, base: str | None = None, folder: str | None = None,
           ortak_belge: str | None = None) -> dict:
    """Parkın belge hazırlık raporunu bas + readiness dict döndür. Mod `park.onay_pdf`'ten gelir."""
    mode = park.onay_pdf or PER_PERSON
    folder = folder or find_belge_klasor(park, base)
    r = readiness(people, folder, mode=mode, ortak_belge=ortak_belge)
    etiket = "ORTAK tek belge" if mode == ORTAK else "kişi-başı belge"
    print(f"\n{LOG} {park.code} ({park.ad}) — {r['toplam']} kişi · mod: {etiket}")

    if mode == ORTAK:
        if r["ortak"]:
            print(f"   📄 Ortak belge: {r['ortak']}")
            print(f"   ✅ {r['toplam']} kişinin hepsine bu belge yüklenecek.")
        else:
            print(f"   ❌ Ortak belge çözülemedi: {r['hata']}")
            print(f"   → Bu park onaya GÖNDERİLMEZ (adaptif atlama); giriş TASLAK kalır.")
        if r["nonpdf"]:
            print(f"   ⚠️  PDF-DIŞI dosya ({len(r['nonpdf'])}) — portal reddeder, PDF'e çevir: {r['nonpdf'][:6]}")
        if r["ortak"] and not r["nonpdf"]:
            print(f"   🎯 {park.code}: ortak belge hazır → onaya gönderilebilir.")
        return r

    if not folder:
        print(f"   ❌ Belge klasörü bulunamadı. Aranan konvansiyon: "
              f"'{base or DEFAULT_BASE}/{park.code} İzin Belgeleri/' (ya da --belge-klasor ile ver).")
        print(f"   → Bu park onaya GÖNDERİLMEZ (adaptif atlama); giriş TASLAK kalır.")
        return r
    print(f"   📁 Klasör: {folder}")
    print(f"   ✅ Belgesi eşleşen: {r['eslesen']}/{r['toplam']}")
    if r["missing"]:
        print(f"   ⚠️  Belgesi EKSİK ({len(r['missing'])}) — onaya gönderilmez, giriş taslak kalır:")
        for p in r["missing"]:
            print(f"        - {p.ad} (T.C.{p.tc})")
    if r["nonpdf"]:
        print(f"   ⚠️  PDF-DIŞI dosya ({len(r['nonpdf'])}) — portal reddeder, PDF'e çevir: {r['nonpdf'][:6]}")
    if r["extra"]:
        ex = [f for f in r["extra"] if os.path.splitext(f)[1].lower() in _PDF_EXT]
        if ex:
            print(f"   ℹ️  Kimseyle eşleşmeyen PDF ({len(ex)}) — isim yazımını kontrol et: {ex[:6]}")
    if r["eslesen"] == r["toplam"] and not r["nonpdf"]:
        print(f"   🎯 {park.code}: TÜM belgeler hazır → onaya gönderilebilir.")
    return r


def _main():
    ap = argparse.ArgumentParser(description="İzin belgesi çözücü — kişi↔PDF eşleştirme raporu (portalsız)")
    ap.add_argument("excel", help="İzin Excel yolu")
    ap.add_argument("--park", required=True, help="PDF'li park (İYTE/YTP/ULUTEK)")
    ap.add_argument("--belge-klasor", default=None, help="Belge klasörü (yoksa Masaüstü konvansiyonu aranır)")
    ap.add_argument("--ortak-belge", default=None,
                    help="ORTAK modlu parkta (İYTE) toplu belgenin yolu — klasör keşfini baypas eder")
    ap.add_argument("--base", default=None, help="Konvansiyon kök dizini (varsayılan ~/Desktop/İzin Belgeleri)")
    args = ap.parse_args()
    pk = resolve_park(args.park)
    if pk is None:
        print(f"Bilinmeyen park: {args.park}", file=sys.stderr); sys.exit(1)
    if not pk.onay_pdf:
        print(f"{pk.code} onay için PDF gerektirmez (PDF'siz direkt onay).", file=sys.stderr); sys.exit(0)
    if args.ortak_belge and pk.onay_pdf != ORTAK:
        print(f"{pk.code} modu '{pk.onay_pdf}' — --ortak-belge sadece ORTAK parkta geçerli.", file=sys.stderr); sys.exit(1)
    by_park, meta = read_izin_v2(os.path.expanduser(args.excel), strict=True)
    people = by_park.get(pk.code, [])
    if not people:
        print(f"{pk.code} için Excel'de kişi yok.", file=sys.stderr); sys.exit(0)
    r = report(pk, people, base=args.base, folder=args.belge_klasor, ortak_belge=args.ortak_belge)
    sys.exit(0 if r["eslesen"] == r["toplam"] and not r["nonpdf"] else 4)


if __name__ == "__main__":
    _main()
