#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TEKNOPARK İZMİR (TPIz) — DGS giriş + onay scripti.
==================================================
    python3 dgs_tpiz.py --excel "<xlsx>"                      # DRY-RUN (yazmaz)
    python3 dgs_tpiz.py --excel "<xlsx>" --limit 1 --commit   # tek kişi TASLAK (test)
    python3 dgs_tpiz.py --excel "<xlsx>" --onayla --commit    # BULK giriş+onay
    python3 dgs_tpiz.py onay    --excel "<xlsx>" --commit     # kalan taslakları onaya gönder
    python3 dgs_tpiz.py kontrol --excel "<xlsx>"              # SGK Gün=30 kapanış kontrolü

⚠️ Excel lokasyon kodu = "TPIz" (izin tarafındaki park kodu "İYTE" — AYNI portal, farklı etiket).
CANLI DURUM (Mayıs 2026): DGS 37/37 giriş + 37/37 onay + SGK raporu temiz (hepsi Gün=30), tek kod
değişikliği gerekmeden. `dgs_rapor_kontrol` bu parkta CSV-fetch ile ÇALIŞIYOR (Yıldız'da 401 veriyordu).
NOT: TPIz'de İZİN onayı zorunlu PDF ister (DGS değil) — o ayrı akış: izin_onaya_dosyali*.py.
"""
import dgs_park

PARK = dgs_park.PARKS["TPIz"]


def overrides(D):
    """TPIz'e ÖZEL davranış farkları buraya (diğer parklar etkilenmez).
    2026-07-14 itibarıyla DGS tarafında motor-farkı YOK — standart akış çalışıyor."""
    pass


if __name__ == "__main__":
    dgs_park.run(PARK, overrides)
