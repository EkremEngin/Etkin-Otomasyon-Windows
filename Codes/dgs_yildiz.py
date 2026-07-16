#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
YILDIZ TEKNOPARK (Davutpaşa) — DGS giriş + onay scripti.
========================================================
    python3 dgs_yildiz.py --excel "<xlsx>"                      # DRY-RUN (yazmaz)
    python3 dgs_yildiz.py --excel "<xlsx>" --limit 1 --commit   # tek kişi TASLAK (test)
    python3 dgs_yildiz.py --excel "<xlsx>" --onayla --commit    # BULK giriş+onay
    python3 dgs_yildiz.py onay    --excel "<xlsx>" --commit     # kalan taslakları onaya gönder
    python3 dgs_yildiz.py kontrol --excel "<xlsx>"              # SGK Gün=30 kapanış kontrolü

⚠️ Excel lokasyon kodu = "Yıldız" (YTP DEĞİL — YTP izin tarafının park kodu).
CANLI DURUM (HAZİRAN 2026, 2026-07-14): doğrulama turu 3/3 "Yönetici Şirket Tarafından Onaylanmış"
(ABDULLAH KIZIL 142:15 · ALEYNA YILDIRIM 167:01 · AYÇA AYDINLI 157:49).
BİLİNEN PARK NOTU: `dgs_rapor_kontrol` Yıldız'da CSV-fetch'te 401 dönmüştü (endpoint/auth park-bazlı) →
SGK teyidi elle (PERSONEL > PDKS > SGK Gün Detaylı Rapor). Giriş/onay motoru sorunsuz.
"""
import dgs_park

PARK = dgs_park.PARKS["Yıldız"]


def overrides(D):
    """Yıldız'a ÖZEL davranış farkları buraya (diğer parklar etkilenmez).
    2026-07-14 itibarıyla giriş/onay motorunda Yıldız-farkı YOK — standart akış çalışıyor."""
    pass


if __name__ == "__main__":
    dgs_park.run(PARK, overrides)
