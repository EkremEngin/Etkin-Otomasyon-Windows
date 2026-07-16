#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TEKNOPARK İSTANBUL (TPI) — DGS giriş + onay scripti.
====================================================
    python3 dgs_tpi.py --excel "<xlsx>"                      # DRY-RUN (yazmaz)
    python3 dgs_tpi.py --excel "<xlsx>" --limit 1 --commit   # tek kişi TASLAK (test)
    python3 dgs_tpi.py --excel "<xlsx>" --onayla --commit    # BULK giriş+onay
    python3 dgs_tpi.py onay    --excel "<xlsx>" --commit     # kalan taslakları onaya gönder
    python3 dgs_tpi.py kontrol --excel "<xlsx>"              # SGK Gün=30 kapanış kontrolü

CANLI DURUM (HAZİRAN 2026, 2026-07-14): doğrulama turu 3/3 "Yönetici Şirket Tarafından Onaylanmış"
(ABDULBAKİ ENES BEĞEÇARSLAN 140:44 · ALPHAN YALAZ 126:30 · ARDA DENİZ BOSTAN 153:00).
NOT: portalda bize AİT OLMAYAN taslaklar olabiliyor (ör. SEHER ÖZDEMİR 18:00) — motor yalnız
dgs_done_TPI_<ay>.txt'deki kişilere dokunur, onlara ASLA dokunmaz.
"""
import dgs_park

PARK = dgs_park.PARKS["TPI"]


def overrides(D):
    """TPI'ye ÖZEL davranış farkları buraya (diğer parklar etkilenmez).
    2026-07-14 itibarıyla TPI'de motor-farkı YOK — standart akış çalışıyor."""
    pass


if __name__ == "__main__":
    dgs_park.run(PARK, overrides)
