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

Masaüstündeki `İzin Belgeleri/<park> İzin Belgeleri` klasörü otomatik keşfedilir.
`--ortak-belge` tek bir PDF seçer. `--kisi-belge-klasoru` ile kullanıcı seçtiği klasör ve alt
klasörlerindeki PDF'leri kişi bazında eşleştirebilir. Dosya veya alt klasör adındaki tam ad-soyad
ya da T.C. kişiyi belirler.
Kişi yalnız klasör adından tanınıyorsa dosya veya alt klasör adında 'izin' bulunmalıdır.
Birden fazla kişiye işaret eden PDF yüklenmez. Aynı kişinin ayrı izin formları sırayla yüklenir.

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
import unicodedata
from types import SimpleNamespace

from izin_data_v2 import fold, resolve_park, read_izin_v2
import izin_frozen                                     # Masaüstü'nü platformdan bağımsız bulmak için

_NUM_SUFFIX = re.compile(r"\s*\((\d+)\)\s*$")          # " (1)", " (2)" tekrar eki
_TC_IN_NAME = re.compile(r"(?<!\d)(\d{11})(?!\d)")     # dosya adında 11 haneli T.C.
_PDF_EXT = {".pdf"}
_LEAVE = re.compile(r"\b(?:izin(?:ler|leri|li)?|izni|istirahat|is goremezlik)\b")
_ONBOARDING = re.compile(
    r"\b(?:diploma(?:si)?|kimlik|kimligi|nufus|ozgecmis|ikametgah|sozlesme(?:si)?|"
    r"bordro(?:su)?|adli sicil|ise giris|isten ayrilis)\b"
)
# Kimlik numarası yoksa tanıdık belge/tarih sözcükleri ad-soyada eşlik edebilir.
# Örneğin 'Ayça Ayşe Güneş' dosyasını yalnızca 'Ayşe Güneş' izinde diye ona atamamalıyız.
_DOC_WORDS = frozenset("""
    izin izinli izni izinler izinleri yillik ucretli ucretsiz mazeret idari hastalik rapor raporu
    istirahat is goremezlik form formlar formu formlari dilekce dilekcesi dilekceleri talep talebi
    belge belgesi belgeler belgeleri onay onayi onayli imza imzali tarama taranmis scan signed
    ek ekim ocak subat mart nisan mayis haziran temmuz agustos eylul kasim aralik ay ayi tarih tarihli
    personel personeli calisan tc tckn no numarali gun gunluk kopya kopyasi ve ile icin
    annual leave request application form pdf
""".split())
# 🔴 expanduser("~/Desktop") DEĞİL: Windows'ta Masaüstü OneDrive'a yönlendirilmiş olabilir ve o yol
# hiç var olmaz → belge klasörü otomatik keşfi sessizce başarısız olurdu (bkz. izin_frozen.masaustu).
DEFAULT_BASE = os.path.join(izin_frozen.masaustu(), "İzin Belgeleri")
LOG = "[BELGE]"

PER_PERSON = "per_person"   # kişi başına kendi belgesi
ORTAK = "ortak"             # park için tek toplu belge (aynı PDF herkese)


def person_key(tc: str, ad: str) -> str:
    """Tam T.C. tekildir. Maskeli veya eski kayıtta kimliği adla birlikte tut."""
    tc = str(tc or "").strip()
    return tc if _TC_IN_NAME.fullmatch(tc) else f"{tc}|{fold(ad)}"


def _list_files(folder: str | None, recursive: bool = False) -> list[str]:
    if not folder or not os.path.isdir(folder):
        return []
    files = []
    for root, dirs, names in os.walk(folder):
        dirs[:] = sorted(d for d in dirs if not d.startswith((".", "_"))) if recursive else []
        for name in sorted(names):
            if name.startswith((".", "_", "~$")):
                continue
            path = os.path.join(root, name)
            if os.path.isfile(path) and not os.path.islink(path):
                files.append(os.path.relpath(path, folder))
    return files


def _pdfs(folder: str | None, recursive: bool = True) -> list[str]:
    return [f for f in _list_files(folder, recursive) if os.path.splitext(f)[1].lower() in _PDF_EXT]


def _words(text: str) -> str:
    # macOS dosya adları ayrık Unicode ile gelebilir. Noktalama ve alt çizgi sözcük sınırıdır.
    value = fold(unicodedata.normalize("NFC", text))
    return " ".join(re.findall(r"[^\W_]+", value))


def _owner(parts, people_by_key, names):
    """(kişi anahtarı, belirsiz mi). T.C. ile isim çelişiyorsa dosyayı tahminle atama."""
    words = [f" {_words(part)} " for part in parts]
    by_name = {key for key, name in names.items() if any(name in part for part in words)}
    ids = {tc for part in parts for tc in _TC_IN_NAME.findall(part)}
    if not ids:
        if len(by_name) != 1:
            return None, len(by_name) > 1
        key = next(iter(by_name))
        for part in words:
            if names[key] not in part:
                continue
            rest = part.replace(names[key], " ", 1).split()
            if all(w.isdigit() or w in _DOC_WORDS for w in rest):
                return key, False
        return None, False
    if len(ids) != 1:
        return None, bool(by_name or ids.intersection(people_by_key))
    tc = next(iter(ids))
    if tc in people_by_key:
        # Aynı adı taşıyan iki kişide tam T.C. ayırt eder. Başka bir ad geçiyorsa çelişki vardır.
        if any(names[key] != names.get(tc) for key in by_name):
            return None, True
        return tc, False
    if len(by_name) == 1:
        key = next(iter(by_name))
        mask = str(people_by_key[key].tc or "").strip()
        if re.fullmatch(r"\d{2}\*+", mask) and tc.startswith(mask[:2]):
            return key, False
        if re.fullmatch(r"[0-9*]{11}", mask) and "*" in mask:
            if all(c == "*" or c == digit for c, digit in zip(mask, tc)):
                return key, False
    return None, bool(by_name)


def portal_filename(path: str) -> str:
    """Portal yüklenen dosyanın klasörünü saklamaz ve adı farklı Unicode biçiminde gösterebilir."""
    return unicodedata.normalize("NFC", os.path.basename(path)).strip()


def _file_order(path: str):
    suffix = _NUM_SUFFIX.search(os.path.splitext(os.path.basename(path))[0])
    return (int(suffix[1]) if suffix else 0, path)


def resolve_ortak_belge(folder: str | None, explicit: str | None = None) -> tuple[str | None, str]:
    """ORTAK mod: parkın TEK toplu belgesini çöz → (yol, hata). Hata boşsa yol geçerli.
    `explicit` (--ortak-belge) verilirse klasör hiç okunmaz — exe'nin "tek belge seç" düğmesinin karşılığı.
    Klasörden çözerken TAM 1 PDF şart: 0 → eksik, >1 → BELİRSİZ (sessizce yanlışını seçmektense DUR)."""
    if explicit:
        p = os.path.abspath(os.path.expanduser(explicit))
        if not os.path.isfile(p):
            return None, f"ortak belge bulunamadı: {p}"
        if os.path.splitext(p)[1].lower() not in _PDF_EXT:
            return None, f"ortak belge PDF değil (portal reddeder): {os.path.basename(p)}"
        return p, ""
    if not folder or not os.path.isdir(folder):
        return None, "belge klasörü yok"
    pdfs = _pdfs(folder, recursive=False)
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


def readiness(people, folder: str | None, mode: str = PER_PERSON, ortak_belge: str | None = None,
              files=None) -> dict:
    """Kişi listesi + klasör → {matched{tc:[dosya]}, missing[person], extra[dosya], nonpdf[dosya], folder, mod, ortak, hata}.
    mode=ORTAK → tek toplu belge TÜM kişilere eşleşir (çözülemezse herkes 'missing', sebep `hata`da).
    files: kullanıcının TEK TEK seçtiği dosyalar (Windows'ta klasör penceresi PDF göstermiyor → çoklu dosya
    seçimi). Verilirse klasör taranmaz; yalnız bu dosyalar, DOSYA ADINDAKİ isim/T.C. ile eşleştirilir."""
    folder = os.path.abspath(os.path.expanduser(folder)) if folder else None
    if mode == ORTAK:
        shared, hata = resolve_ortak_belge(folder, ortak_belge)
        # Elle seçilen tek PDF bağımsızdır. Klasörden keşifte eski tek-belge kuralını koru.
        nonpdf = [] if ortak_belge else [f for f in _list_files(folder)
                                          if os.path.splitext(f)[1].lower() not in _PDF_EXT]
        matched = {person_key(p.tc, p.ad): [shared] for p in people} if shared else {}
        return {"matched": matched, "missing": ([] if shared else list(people)), "extra": [],
                "nonpdf": nonpdf, "ambiguous": [], "folder": folder, "toplam": len(people),
                "eslesen": len(matched), "mod": ORTAK, "ortak": shared, "hata": hata,
                "hazir": bool(shared and len(matched) == len(people) and not nonpdf)}

    people_by_key = {person_key(p.tc, p.ad): p for p in people}
    names = {key: f" {_words(p.ad)} " for key, p in people_by_key.items() if _words(p.ad)}
    matched, extra, ambiguous, nonpdf = {}, [], [], []
    if files is not None:
        # (gösterilen ad, tam yol, eşleşme bileşenleri): yalnız dosya adı — klasör adı kişiye işaret etmez.
        entries = []
        for f in files:
            full = os.path.abspath(os.path.expanduser(str(f)))
            if os.path.isfile(full):
                entries.append((os.path.basename(full), full, [os.path.splitext(os.path.basename(full))[0]]))
        try:
            folder = os.path.commonpath([os.path.dirname(e[1]) for e in entries]) if entries else None
        except ValueError:                       # Windows: farklı sürücüler ortak köke sahip değil
            folder = os.path.dirname(entries[0][1])
    else:
        # Tüm ad bir dosya veya klasör bileşeninde geçmeli (iki klasör bir adı oluşturamaz).
        entries = [(rel, os.path.join(folder, rel), os.path.splitext(rel)[0].split(os.sep))
                   for rel in _list_files(folder, recursive=True)]
    pdf_stems = {os.path.normcase(os.path.splitext(rel)[0]) for rel, _full, _p in entries
                 if os.path.splitext(rel)[1].lower() in _PDF_EXT}
    for rel, full, parts in entries:
        stem, ext = os.path.splitext(rel)
        owner, conflict = _owner(parts, people_by_key, names)
        if owner:
            file_words = f" {_words(parts[-1])} "
            has_name = owner in names and names[owner] in file_words
            file_identity = has_name or _TC_IN_NAME.search(parts[-1])
            # Personel klasöründe kimlik, diploma gibi izinle ilgisiz PDF'ler de bulunabilir.
            leave_path = any(_LEAVE.search(_words(part)) for part in parts)
            if _ONBOARDING.search(file_words) or not (file_identity or leave_path):
                owner = None
        if ext.lower() not in _PDF_EXT:
            if owner and os.path.normcase(stem) not in pdf_stems:
                nonpdf.append(rel)
        elif conflict:
            ambiguous.append(rel)
        elif owner:
            matched.setdefault(owner, []).append(full)
        else:
            extra.append(rel)
    same_basename = []
    for paths in matched.values():
        paths.sort(key=_file_order)
        seen_names = set()
        for path in paths:
            name = portal_filename(path)
            if name in seen_names:
                same_basename.append(os.path.relpath(path, folder))
            seen_names.add(name)
    missing = [p for p in people if person_key(p.tc, p.ad) not in matched]
    hata = ""
    if files is not None and not entries:
        hata = "Seçilen PDF dosyaları bulunamadı."
    elif not folder or not os.path.isdir(folder):
        hata = "Belge klasörü bulunamadı."
    elif same_basename:
        hata = f"Aynı kişi için dosya adları aynı. PDF'leri farklı adlandırın: {', '.join(same_basename[:4])}"
    elif ambiguous:
        hata = f"Kişisi belirsiz PDF ({len(ambiguous)}): {', '.join(ambiguous[:4])}"
    return {"matched": matched, "missing": missing, "extra": extra, "nonpdf": nonpdf,
            "ambiguous": ambiguous, "folder": folder, "toplam": len(people), "eslesen": len(matched),
            "mod": PER_PERSON, "ortak": None, "hata": hata, "hazir": not (hata or missing or nonpdf)}


def resolve_person_files(folder: str, ad: str, tc: str = "") -> list[str]:
    """Tek kişilik denetim. Toplu yüklemede tüm kişileri birlikte readiness ile eşleştir."""
    r = readiness([SimpleNamespace(ad=ad, tc=tc)], folder)
    return r["matched"].get(person_key(tc, ad), [])


def dosya_listesi_oku(yol: str | None):
    """GUI'nin yazdığı seçili-PDF listesi (UTF-8, satır başına bir yol). Yol verilmezse None."""
    if not yol:
        return None
    with open(yol, encoding="utf-8-sig") as f:
        return [s.strip() for s in f if s.strip()]


def dosya_listesi_yaz(dosyalar, yol: str) -> str:
    with open(yol, "w", encoding="utf-8") as f:
        f.write("\n".join(dosyalar) + "\n")
    return yol


def source_readiness(park, people, *, base=None, ortak_belge=None, kisi_klasoru=None, kisi_dosyalari=None):
    """Plan, ekran ve yükleyici için aynı belge kaynağını çöz."""
    if kisi_dosyalari is not None:
        if ortak_belge or kisi_klasoru is not None or not park.belge_klasoru_destekli:
            raise ValueError("Seçilen kişisel PDF'ler ortak PDF veya klasörle birlikte kullanılamaz.")
        return readiness(people, None, mode=PER_PERSON, files=kisi_dosyalari)
    if kisi_klasoru is not None:
        if ortak_belge or not park.belge_klasoru_destekli:
            raise ValueError("Bu park veya ortak PDF seçimi kişi klasörüyle birlikte kullanılamaz.")
        return readiness(people, kisi_klasoru, mode=PER_PERSON)
    folder = None if ortak_belge and park.onay_pdf == ORTAK else find_belge_klasor(park, base)
    return readiness(people, folder, mode=park.onay_pdf or PER_PERSON, ortak_belge=ortak_belge)


def report(park, people, base: str | None = None, folder: str | None = None,
           ortak_belge: str | None = None, kisi_klasoru: str | None = None, kisi_dosyalari=None) -> dict:
    """Parkın seçilen kaynağına göre belge hazırlık raporunu bas."""
    if folder and kisi_klasoru is None and kisi_dosyalari is None:
        r = readiness(people, folder, mode=park.onay_pdf or PER_PERSON, ortak_belge=ortak_belge)
    else:
        r = source_readiness(park, people, base=base, ortak_belge=ortak_belge, kisi_klasoru=kisi_klasoru,
                             kisi_dosyalari=kisi_dosyalari)
    mode, folder = r["mod"], r["folder"]
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
    if r["hata"]:
        print(f"   ⚠️  {r['hata']}")
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
    if r["hazir"]:
        if park.onay_dogrulandi:
            print(f"   🎯 {park.code}: TÜM belgeler hazır → onaya gönderilebilir.")
        else:
            print(f"   {park.code}: yerel PDF eşleşmeleri tamam. Portal girişi yalnız taslaktır.")
    return r


def _main():
    ap = argparse.ArgumentParser(description="İzin belgesi çözücü — kişi↔PDF eşleştirme raporu (portalsız)")
    ap.add_argument("excel", help="İzin Excel yolu")
    ap.add_argument("--park", required=True, help="Belge denetimi yapılan park (İYTE/YTP/ULUTEK/DIJITALPARK)")
    ap.add_argument("--belge-klasor", default=None, help="Belge klasörü (yoksa Masaüstü konvansiyonu aranır)")
    src = ap.add_mutually_exclusive_group()
    src.add_argument("--kisi-belge-klasoru", help="Kişi bazında izin formlarının aranacağı klasör ve alt klasörleri")
    src.add_argument("--kisi-belge-listesi", help="Tek tek seçilmiş kişisel PDF'lerin listesi (satır başına bir yol)")
    src.add_argument("--ortak-belge", default=None,
                    help="ORTAK modlu parkta (İYTE) toplu belgenin yolu — klasör keşfini baypas eder")
    ap.add_argument("--base", default=None, help="Konvansiyon kök dizini (varsayılan ~/Desktop/İzin Belgeleri)")
    args = ap.parse_args()
    pk = resolve_park(args.park)
    if pk is None:
        print(f"Bilinmeyen park: {args.park}", file=sys.stderr); sys.exit(1)
    if not pk.belge_klasoru_destekli:
        print(f"{pk.code} onay için PDF gerektirmez (PDF'siz direkt onay).", file=sys.stderr); sys.exit(0)
    if args.ortak_belge and pk.onay_pdf != ORTAK:
        print(f"{pk.code} modu '{pk.onay_pdf}' — --ortak-belge sadece ORTAK parkta geçerli.", file=sys.stderr); sys.exit(1)
    by_park, meta = read_izin_v2(os.path.expanduser(args.excel), strict=True, default_park=pk.code)
    people = by_park.get(pk.code, [])
    if not people:
        print(f"{pk.code} için Excel'de kişi yok.", file=sys.stderr); sys.exit(0)
    r = report(pk, people, base=args.base, folder=args.belge_klasor, ortak_belge=args.ortak_belge,
               kisi_klasoru=args.kisi_belge_klasoru, kisi_dosyalari=dosya_listesi_oku(args.kisi_belge_listesi))
    sys.exit(0 if r["hazir"] else 4)


if __name__ == "__main__":
    _main()
