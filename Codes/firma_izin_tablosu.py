#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Firma çalışan tablosu → günlük izin satırları (izin v2 biçimi).

Bazı firmalar (ör. bir Dijitalpark firması) izni gün gün değil, kişi başına ARALIK olarak yollar:

    Sıra | Ad Soyad | Kimlik Numarası | ARGE Portal Personel Tipi | İŞE GİRİŞ | İŞTEN ÇIKIŞ |
    (Başlangıç Tarihi | Bitiş Tarihi (İşe Başlama) | İzin Kullanımı Gün Sayısı | İzin Türü) × N | İzin Toplam Gün

Kurallar (tahmin yok, tutmazsa DUR):
  • Bitiş = işe DÖNÜŞ günü → aralığa dahil DEĞİL.
  • Aralık iş günlerine açılır: hafta sonu ve resmi tatil atlanır, arife yarım gün (0,5) sayılır.
  • Açılan toplam Excel'deki gün sayısına eşit olmalı. Tek günlük aralıkta 0,5 yazıyorsa o gün yarım gündür.
  • İzin günü işe girişten önce / işten çıkıştan sonra olamaz; seçilen dönemin dışında kalan günler atılır.

Portalsız kontrol:  python3 firma_izin_tablosu.py "<xlsx>"
v2 dosyası üret:    python3 firma_izin_tablosu.py "<xlsx>" --park DIJITALPARK --out izin_v2.xlsx
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import sys
from dataclasses import dataclass, field

from openpyxl import Workbook, load_workbook

import izin_data_v2 as v2

# Resmi tatiller. Sabit günler her yıl; dini bayramlar yıl yıl eklenir. Listede olmayan bir bayram aralığa
# düşerse açılan gün sayısı Excel'i tutmaz → dosya DURUR (sessizce yanlış gün girilmez).
_SABIT_TATIL = {(1, 1), (4, 23), (5, 1), (5, 19), (7, 15), (8, 30), (10, 29)}
_SABIT_ARIFE = {(10, 28)}                       # Cumhuriyet Bayramı arifesi: öğleden sonra tatil
_DINI_TATIL = {
    2026: {dt.date(2026, 3, 20), dt.date(2026, 3, 21), dt.date(2026, 3, 22),
           dt.date(2026, 5, 27), dt.date(2026, 5, 28), dt.date(2026, 5, 29), dt.date(2026, 5, 30)},
}
_DINI_ARIFE = {2026: {dt.date(2026, 3, 19), dt.date(2026, 5, 26)}}


class TabloHatasi(ValueError):
    """Firma tablosunda sert hata; portala dokunulmadan durulur."""


@dataclass
class IzinGunu:
    tarih: dt.date
    gun: float          # 1.0 | 0.5
    tur: str            # Excel'deki izin türü metni (Yıllık İzin, Sağlık İzni, ...)


@dataclass
class Calisan:
    satir: int
    tc: str
    ad: str
    tip: str
    giris: dt.date | None
    cikis: dt.date | None
    izinler: list[IzinGunu] = field(default_factory=list)
    beyan_toplam: float = 0.0


def hak_edilmis(tur: str) -> bool:
    """Yalnız hak edilmiş (yıllık) izin portala izin olarak girilir ve teşvik kapsamındadır.
    Sağlık, doğum günü vb. izin günlerine ne izin ne DGS girilir; gün boş kalır (kullanıcı kuralı 2026-11-24)."""
    return "yillik" in v2.fold(tur)


def bir_yil_doldu(giris: dt.date | None, tarih: dt.date) -> bool:
    """4857 s. Kanun / Yıllık Ücretli İzin Yönetmeliği md. 9: 1 yılını doldurmamış çalışana yıllık ücretli izin
    verilemez (portal da kaydı reddediyor — 2026-11-24 canlı). Giriş tarihi bilinmiyorsa engel sayılmaz."""
    if giris is None:
        return True
    try:
        yildonumu = giris.replace(year=giris.year + 1)
    except ValueError:                       # 29 Şubat girişi
        yildonumu = giris.replace(year=giris.year + 1, day=28)
    return tarih >= yildonumu


def izin_girilir(k: "Calisan", g: "IzinGunu") -> bool:
    """Portala izin olarak girilecek gün mü: hak edilmiş (yıllık) VE çalışan 1 yılını doldurmuş."""
    return hak_edilmis(g.tur) and bir_yil_doldu(k.giris, g.tarih)


def izin_girilmeme_sebebi(k: "Calisan", g: "IzinGunu") -> str:
    if not hak_edilmis(g.tur):
        return g.tur
    return f"{g.tur} — 1 yıl dolmamış (giriş {k.giris:%d.%m.%Y})"


def gun_agirligi(gun: dt.date) -> float:
    """İş günü 1, arife 0,5, hafta sonu/tatil 0."""
    if gun.weekday() >= 5:
        return 0.0
    if (gun.month, gun.day) in _SABIT_TATIL or gun in _DINI_TATIL.get(gun.year, set()):
        return 0.0
    if (gun.month, gun.day) in _SABIT_ARIFE or gun in _DINI_ARIFE.get(gun.year, set()):
        return 0.5
    return 1.0


def araligi_ac(bas: dt.date, donus: dt.date, beyan: float) -> list[tuple[dt.date, float]]:
    """[bas, donus) aralığını izin günlerine aç; toplam beyanı tutmazsa TabloHatasi."""
    if donus <= bas:
        raise TabloHatasi(f"bitiş ({donus:%d.%m.%Y}) başlangıçtan ({bas:%d.%m.%Y}) sonra olmalı")
    gunler = []
    g = bas
    while g < donus:
        w = gun_agirligi(g)
        if w:
            gunler.append((g, w))
        g += dt.timedelta(days=1)
    toplam = sum(w for _, w in gunler)
    if abs(toplam - beyan) < 1e-9:
        return gunler
    if len(gunler) == 1 and gunler[0][1] == 1.0 and abs(beyan - 0.5) < 1e-9:
        return [(gunler[0][0], 0.5)]                     # tek gün, yarım izin
    raise TabloHatasi(f"{bas:%d.%m.%Y}–{donus:%d.%m.%Y} aralığı {toplam:g} iş günü ediyor, "
                      f"Excel {beyan:g} gün diyor (hangi günün yarım olduğu belirsiz)")


def _tarih(v) -> dt.date | None:
    if v in (None, ""):
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise TabloHatasi(f"tarih okunamadı: {s!r}")


def _sayi(v) -> float:
    if v in (None, ""):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return float(str(v).strip().replace(",", "."))


def _baslik_bul(rows) -> tuple[int, dict]:
    """Başlık satırını ve kolonları bul. İzin grupları (başlangıç, bitiş, gün, tür) sırayla eşlenir."""
    for i, r in enumerate(rows[:15]):
        f = [v2.fold(str(c)) if c is not None else "" for c in r]
        try:
            ad = f.index("ad soyad")
        except ValueError:
            continue
        tc = next((j for j, c in enumerate(f) if c in ("kimlik numarasi", "t.c. kimlik", "tc kimlik", "t.c.")), None)
        if tc is None:
            continue
        bul = lambda *ks: next((j for j, c in enumerate(f) if any(c.startswith(k) for k in ks)), None)
        gruplar = []
        for j, c in enumerate(f):
            if c.startswith("baslangic tarihi") and j + 3 < len(f):
                if not (f[j + 1].startswith("bitis tarihi") and f[j + 2].startswith("izin kullanimi")
                        and f[j + 3].startswith("izin turu")):
                    raise TabloHatasi(f"başlık satırı {i + 1}: izin grubu (Başlangıç, Bitiş, Gün Sayısı, "
                                      f"İzin Türü) sırası beklenenden farklı (kolon {j + 1})")
                gruplar.append(j)
        if not gruplar:
            continue
        return i, {"ad": ad, "tc": tc, "tip": bul("arge portal"), "giris": bul("ise giris"),
                   "cikis": bul("isten cikis"), "toplam": bul("izin toplam"), "gruplar": gruplar}
    raise TabloHatasi("Firma çalışan tablosu başlığı bulunamadı ('Ad Soyad', 'Kimlik Numarası', "
                      "'Başlangıç Tarihi / Bitiş Tarihi / İzin Kullanımı Gün Sayısı / İzin Türü').")


def firma_tablosu_mu(path: str) -> bool:
    """Dosya bu biçimde mi? (GUI/okuyucu otomatik seçimi için; hata fırlatmaz.)"""
    try:
        wb = load_workbook(os.path.expanduser(path), data_only=True, read_only=True)
    except Exception:
        return False
    try:
        for ws in wb:
            try:
                _baslik_bul(list(ws.iter_rows(values_only=True, max_row=15)))
                return True
            except TabloHatasi:
                continue
        return False
    finally:
        wb.close()


def donem_tahmin(path: str) -> tuple[int, int]:
    """İzin başlangıç tarihlerinin en sık (yıl, ay)'ı; izin yoksa bugünden bir önceki ay."""
    wb = load_workbook(os.path.expanduser(path), data_only=True, read_only=True)
    sayac: dict[tuple[int, int], int] = {}
    try:
        for ws in wb:
            rows = list(ws.iter_rows(values_only=True))
            try:
                hdr, ci = _baslik_bul(rows)
            except TabloHatasi:
                continue
            for r in rows[hdr + 1:]:
                for j in ci["gruplar"]:
                    try:
                        g = _tarih(r[j]) if j < len(r) else None
                    except TabloHatasi:
                        g = None
                    if g:
                        sayac[(g.year, g.month)] = sayac.get((g.year, g.month), 0) + 1
            break
    finally:
        wb.close()
    if sayac:
        return max(sayac, key=sayac.get)
    onceki = dt.date.today().replace(day=1) - dt.timedelta(days=1)
    return onceki.year, onceki.month


def oku(path: str, yil: int, ay: int) -> list[Calisan]:
    """Tabloyu oku + doğrula. Hatalar toplanıp tek TabloHatasi ile verilir (portala dokunulmaz)."""
    wb = load_workbook(os.path.expanduser(path), data_only=True, read_only=True)
    try:
        bulunan = None
        for ws in wb:
            rows = list(ws.iter_rows(values_only=True))
            try:
                bulunan = (rows, *_baslik_bul(rows))
                break
            except TabloHatasi:
                continue
    finally:
        wb.close()
    if not bulunan:
        raise TabloHatasi("Firma çalışan tablosu başlığı bulunamadı.")
    rows, hdr, ci = bulunan
    ay_bas = dt.date(yil, ay, 1)
    ay_son = (dt.date(yil + (ay == 12), ay % 12 + 1, 1))
    hatalar, kisiler, gorulen_tc = [], [], set()
    for n, r in enumerate(rows[hdr + 1:], start=hdr + 2):
        cell = lambda j: r[j] if j is not None and j < len(r) else None
        ad = str(cell(ci["ad"]) or "").strip()
        if not ad:
            continue
        tc = str(cell(ci["tc"]) or "").strip()
        if tc.endswith(".0"):
            tc = tc[:-2]
        try:
            if not v2.tc_valid(tc):
                raise TabloHatasi("T.C. geçersiz (11 hane + resmi kontrol)")
            if tc in gorulen_tc:
                raise TabloHatasi("aynı T.C. ikinci kez var")
            gorulen_tc.add(tc)
            k = Calisan(n, tc, ad, str(cell(ci["tip"]) or "").strip(), _tarih(cell(ci["giris"])),
                        _tarih(cell(ci["cikis"])), beyan_toplam=_sayi(cell(ci["toplam"])))
            acilan = 0.0
            for j in ci["gruplar"]:
                bas, bit, gun, tur = (cell(j), cell(j + 1), cell(j + 2), cell(j + 3))
                if all(v in (None, "") for v in (bas, bit, gun, tur)):
                    continue
                bas, bit, gun = _tarih(bas), _tarih(bit), _sayi(gun)
                if not (bas and bit and gun > 0):
                    raise TabloHatasi(f"kolon {j + 1}: izin grubu eksik (başlangıç/bitiş/gün)")
                tur = str(tur or "").strip() or "Yıllık İzin"
                for g, w in araligi_ac(bas, bit, gun):
                    acilan += w
                    if k.giris and g < k.giris:
                        raise TabloHatasi(f"{g:%d.%m.%Y} izni işe girişten ({k.giris:%d.%m.%Y}) önce")
                    if k.cikis and g > k.cikis:
                        raise TabloHatasi(f"{g:%d.%m.%Y} izni işten çıkıştan ({k.cikis:%d.%m.%Y}) sonra")
                    if ay_bas <= g < ay_son:
                        if any(x.tarih == g for x in k.izinler):
                            raise TabloHatasi(f"{g:%d.%m.%Y} iki izin aralığında birden var")
                        k.izinler.append(IzinGunu(g, w, tur))
            if ci["toplam"] is not None and abs(acilan - k.beyan_toplam) > 1e-9:
                raise TabloHatasi(f"izin grupları {acilan:g} gün, 'İzin Toplam Gün' {k.beyan_toplam:g}")
            kisiler.append(k)
        except TabloHatasi as e:
            hatalar.append(f"  satır {n} ({ad}): {e}")
    if hatalar:
        raise TabloHatasi("Firma tablosunda düzeltilmesi gereken satırlar var; portala dokunulmadı:\n"
                          + "\n".join(hatalar))
    if not kisiler:
        raise TabloHatasi("Tabloda çalışan bulunamadı.")
    return kisiler


def v2_yaz(kisiler: list[Calisan], park_code: str, out: str, turler: set[str] | None = None) -> int:
    """İzin günlerini izin_data_v2'nin okuduğu günlük biçime yaz. turler verilirse yalnız o türler."""
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.title = "Izinler"
    ws.append(["T.C.", "Ad-Soyad", "Tarih", "Gün Sayısı", "TGB", "İzin Türü"])
    n = 0
    for k in kisiler:
        for g in sorted(k.izinler, key=lambda x: x.tarih):
            if turler is not None and v2.fold(g.tur) not in {v2.fold(t) for t in turler}:
                continue
            ws.append([k.tc, k.ad, g.tarih.strftime("%d.%m.%Y"), g.gun, park_code, g.tur])
            n += 1
    wb.save(out)
    wb.close()
    return n


def ar_ge_sinifi(tip: str) -> str:
    """Portal personel tipinden DGS sınıfı. Bilinmeyen tip tahmin edilmez → TabloHatasi."""
    f = v2.fold(tip)
    if "destek" in f:
        return "Destek"
    if "yazilim" in f or "arastirmaci" in f or "arge" in f.replace("-", ""):
        return "Ar-Ge"
    raise TabloHatasi(f"personel tipi {tip!r} Ar-Ge mi Destek mi bilinmiyor")


def dgs_yaz(kisiler: list[Calisan], lokasyon: str, proje: str, ay_adi: str, out: str,
            tip_duzelt: dict[str, str] | None = None) -> int:
    """Ortak DGS puantaj biçimini yaz (dgs_poc.read_excel 'ortak' başlıkları). tip_duzelt: T.C. → tip."""
    if not proje.strip():
        raise TabloHatasi("DGS için proje adı gerekli")
    wb = Workbook()
    ws = wb.worksheets[0]
    ws.title = ay_adi
    ws.append(["AD SOYAD", "TC Kimlik", "BÖLÜM", "Lokasyonu SGK", "Ar-Ge/ Destek/ K.Dışı",
               "Eksik Puantaj (Saat)", "Proje Adı"])
    hatalar = []
    for k in kisiler:
        try:
            sinif = ar_ge_sinifi((tip_duzelt or {}).get(k.tc, k.tip))
        except TabloHatasi as e:
            hatalar.append(f"  satır {k.satir} ({k.ad}): {e}")
            continue
        ws.append([k.ad, k.tc, "", lokasyon, sinif, "", proje.strip()])
    if hatalar:
        wb.close()
        raise TabloHatasi("DGS sınıfı belirlenemeyen çalışanlar var:\n" + "\n".join(hatalar))
    wb.save(out)
    wb.close()
    return len(kisiler)


def main():
    ap = argparse.ArgumentParser(description="Firma çalışan tablosu → günlük izin (izin v2) dönüştürücü")
    ap.add_argument("excel")
    ap.add_argument("--donem", default=None, help="'EKİM 2026'; yoksa bugünden bir önceki ay")
    ap.add_argument("--park", default=None, help="v2 TGB kodu (ör. DIJITALPARK); --out ile gerekli")
    ap.add_argument("--out", default=None, help="Yazılacak izin v2 xlsx yolu")
    ap.add_argument("--tur", action="append", default=None,
                    help="Yalnız bu izin türünü yaz (tekrarlanabilir, ör. --tur 'Yıllık İzin')")
    ap.add_argument("--dgs-out", default=None, help="Ortak DGS puantaj xlsx yolu (dgs_poc --excel girdisi)")
    ap.add_argument("--dgs-lokasyon", default=None, help="DGS park kodu (ör. Dijitalpark); --dgs-out ile gerekli")
    ap.add_argument("--proje", default=None, help="DGS proje adı (portalda tek aktif proje)")
    ap.add_argument("--tip", action="append", default=[],
                    help="Excel tipi tanınmayan kişi için 'T.C.=Yazılım Personeli' (portaldaki tip)")
    args = ap.parse_args()

    if args.donem:
        m = re.fullmatch(r"\s*(\S+)\s+(\d{4})\s*", args.donem)
        ay = next((k for k, v in v2._TURKCE_AY.items() if m and v2.fold(v) == v2.fold(m[1])), None)
        if not (m and ay):
            ap.error("--donem 'EKİM 2026' biçiminde olmalı")
        yil = int(m[2])
    else:
        bugun = dt.date.today().replace(day=1) - dt.timedelta(days=1)
        yil, ay = bugun.year, bugun.month
    try:
        kisiler = oku(args.excel, yil, ay)
    except TabloHatasi as e:
        print(f"HATA: {e}")
        sys.exit(2)

    izinli = [k for k in kisiler if k.izinler]
    print(f"Dönem {v2._TURKCE_AY[ay]} {yil}: {len(kisiler)} çalışan, {len(izinli)} izinli, "
          f"{sum(len(k.izinler) for k in kisiler)} izin günü satırı")
    turler: dict[str, float] = {}
    for k in izinli:
        for g in k.izinler:
            turler[g.tur] = turler.get(g.tur, 0) + g.gun
        gunler = ", ".join(f"{g.tarih:%d.%m}{'(½)' if g.gun == 0.5 else ''} {g.tur}" for g in k.izinler)
        print(f"  {k.ad:<28} {sum(g.gun for g in k.izinler):>4g} gün | {gunler}")
    print("Türlere göre:", ", ".join(f"{t}={g:g}" for t, g in turler.items()))
    for k in kisiler:
        if k.cikis:
            print(f"  ÇIKIŞ: {k.ad} {k.cikis:%d.%m.%Y}")
        if k.giris and (k.giris.year, k.giris.month) == (yil, ay):
            print(f"  GİRİŞ: {k.ad} {k.giris:%d.%m.%Y}")

    if args.out:
        park = v2.resolve_park(args.park)
        if park is None:
            ap.error("--out için geçerli --park gerekli (ör. DIJITALPARK)")
        n = v2_yaz(kisiler, park.code, args.out, set(args.tur) if args.tur else None)
        print(f"Yazıldı: {args.out} ({n} satır, park={park.code})")
    if args.dgs_out:
        if not (args.dgs_lokasyon and args.proje):
            ap.error("--dgs-out için --dgs-lokasyon ve --proje gerekli")
        tip_duzelt = dict(t.split("=", 1) for t in args.tip)
        try:
            n = dgs_yaz(kisiler, args.dgs_lokasyon, args.proje, v2._AY_KEY[ay], args.dgs_out, tip_duzelt)
        except TabloHatasi as e:
            print(f"HATA: {e}")
            sys.exit(2)
        print(f"DGS yazıldı: {args.dgs_out} ({n} kişi, lokasyon={args.dgs_lokasyon}, sayfa={v2._AY_KEY[ay]})")


if __name__ == "__main__":
    main()
