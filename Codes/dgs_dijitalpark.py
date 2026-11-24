"""Dijitalpark Teknokent DGS ayarları; giriş ortak doğrulamalı motoru kullanır."""
import dgs_park

PARK = dgs_park.PARKS["Dijitalpark"]


def overrides(motor):
    """Park için doğrulanmış farklı bir form davranışı bulunursa burada tanımlanır."""
    # 2026-11-24 Ekim dönemi canlı: personel seçildikten sonra gün grid'i 12 sn'den geç geliyor
    # (sonradan 22 gün doluyor) → motorun varsayılan beklemesi "grid boş" deyip kişiyi atlıyordu.
    motor.CONFIG["grid_timeout_ms"] = 40000
    # Aynı yavaşlık ✎ gün dialogunda: STALE-guard doğru veriyi 8 sn içinde göremeyip kişiyi durduruyordu.
    # Kontrolün kendisi aynı (tarih + damga + toplam eşleşmeli); yalnız bekleme uzadı.
    motor.CONFIG["dialog_timeout_ms"] = 30000
    # Yavaş dialogda damga tablosu bir an boş, aralıklar önceki günün kalıyor (toplam eşitse eski kontrol
    # geçiyordu → portal 'pdks kaydı mevcuttur' reddi). PDKS gününde o günün damgaları şart.
    motor.CONFIG["pdks_damga_kontrol"] = True
    # Kilitli arife satırında içeride < 4:30 ise eksik dışarıda tamamlanır (SGK raporu GV Gün 29,8x çıkıyordu).
    motor.CONFIG["arife_kilitli_tamamla"] = True


if __name__ == "__main__":
    dgs_park.run(PARK, overrides)
