#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BİLİŞİM VADİSİ (BV) — DGS giriş + onay scripti.
===============================================
    python3 dgs_bv.py --excel "<xlsx>"                      # DRY-RUN (yazmaz)
    python3 dgs_bv.py --excel "<xlsx>" --limit 1 --commit   # tek kişi TASLAK (test)
    python3 dgs_bv.py --excel "<xlsx>" --onayla --commit    # BULK giriş+onay
    python3 dgs_bv.py onay    --excel "<xlsx>" --commit     # kalan taslakları onaya gönder
    python3 dgs_bv.py kontrol --excel "<xlsx>"              # SGK Gün=30 kapanış kontrolü

BİLİNEN PARK NOTU (Mayıs 2026): BV'de ay boyu TGB DIŞINDA çalışan (PDKS damgası olmayan) kişiler çıkmıştı
(MERVE GÜLVEREN SARIZ 139:30, MİRAC GÜL 90:02). Formları açılınca grid tüm günleri temiz + 09:00 dışarıda
OTOMATİK tikli gösteriyor, ama grid AJAX'ı bu kişilerde çok flaky → yarım yükleniyor, verify-or-halt DOĞRU
durduruyor. Çözüm: `fill_external.py` (tam-yükleme gate'i + brute-force retry). Genel scriptle boğuşma.
"""
import dgs_park

PARK = dgs_park.PARKS["BV"]


def overrides(D):
    """BV'ye ÖZEL davranış farkları buraya (diğer parklar etkilenmez).
    2026-07-14 itibarıyla motor-farkı YOK — tam-dışarıda çalışanlar için ayrı araç: fill_external.py."""
    pass


if __name__ == "__main__":
    dgs_park.run(PARK, overrides)
