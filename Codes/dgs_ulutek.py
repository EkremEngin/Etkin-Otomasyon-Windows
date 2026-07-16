#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ULUTEK TEKNOPARK (Bursa) — DGS giriş + onay scripti.
====================================================
    python3 dgs_ulutek.py --excel "<xlsx>"                      # DRY-RUN (yazmaz)
    python3 dgs_ulutek.py --excel "<xlsx>" --limit 1 --commit   # tek kişi TASLAK (test)
    python3 dgs_ulutek.py --excel "<xlsx>" --onayla --commit    # BULK giriş+onay
    python3 dgs_ulutek.py onay    --excel "<xlsx>" --commit     # kalan taslakları onaya gönder
    python3 dgs_ulutek.py kontrol --excel "<xlsx>"              # SGK Gün=30 kapanış kontrolü

CANLI DURUM (Mayıs 2026): DGS 9/9 giriş + 9/9 onay ("Yönetici Şirket Tarafından Onaylanmış").
DERS (Mayıs): bu parkta izinler firma İK'sınca ZATEN girilmişti → DGS'ten önce izin durumunu kontrol et.
NOT: Ulutek'te İZİN onayı zorunlu dilekçe (PDF) ister (DGS değil) — o ayrı akış: izin_otomasyon.py.
"""
import dgs_park

PARK = dgs_park.PARKS["Ulutek"]


def overrides(D):
    """Ulutek'e ÖZEL davranış farkları buraya (diğer parklar etkilenmez).
    2026-07-14 itibarıyla DGS tarafında motor-farkı YOK — standart akış çalışıyor."""
    pass


if __name__ == "__main__":
    dgs_park.run(PARK, overrides)
