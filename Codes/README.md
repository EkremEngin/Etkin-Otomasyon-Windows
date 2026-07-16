# DGS Otomasyon — PoC (tek personel)

Teknopark ARGE Portalı'nda **Dışarıda Geçirilen Süreler (DGS)** girişini, firmanın Excel
puantajına göre otomatik dolduran kanıt-niteliğinde betik.

> **Model: ATTENDED (gözetimli).** Cloudflare doğrulamasını ve girişi **sen** yaparsın;
> betik senin açtığın Chrome oturumuna bağlanır. **"Onaya gönder" ve e-imza her zaman sende** —
> betik bunlara asla dokunmaz. Varsayılan **DRY-RUN**: sunucuya hiçbir şey yazılmaz.

## 1) Kurulum
```bash
cd "dgs-otomasyon"
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```
(CDP ile gerçek Chrome'a bağlandığımız için `playwright install` gerekmez.)

## 2) Chrome'u uzaktan-hata-ayıklama portuyla aç
Tüm Chrome pencerelerini kapat, sonra (macOS):
```bash
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
  --remote-debugging-port=9222 \
  --user-data-dir="$HOME/dgs-chrome-profile"
```
Açılan Chrome'da `https://argeportal.teknoparkistanbul.com.tr/` adresine git,
**Cloudflare'i geç ve giriş yap.** Dashboard'u görene kadar bekle.

## 3) Çalıştırma — KADEMELİ ilerle (resmi veri!)
Excel yolu kısaltması: `XL="$HOME/Downloads/05-TEKNOKENTLER - MAYIS 2026 - FİNAL_9SAAT.xlsx"`

**a) Tek kişi DRY-RUN** (hiçbir şey kaydetmez — akışı doğrula):
```bash
python dgs_poc.py --excel "$XL" --person "AHMET FURKAN FAİDECİ"
```
**b) Tek kişi TASLAK kaydet** (Onaya Gönder YOK):
```bash
python dgs_poc.py --excel "$XL" --person "AHMET FURKAN FAİDECİ" --commit
```
**c) İlk 3 kişiyle batch döngüsünü doğrula** (çoklu-kişi `start_new_entry` ilk kez burada test edilir):
```bash
python dgs_poc.py --excel "$XL" --limit 3 --commit
```
**d) Tüm TPI Ar-Ge (116 kişi) batch:**
```bash
python dgs_poc.py --excel "$XL" --commit
```
- `--lokasyon TPI` varsayılan = bu portaldaki kişiler (Excel "Lokasyonu SGK"). Diğer teknoparklar için `--lokasyon ARI` vb. (o portala girip).
- Hepsi **TASLAK**; biten kişiler `dgs_done_TPI_Mayıs.txt`'e yazılır → tekrar koşarsan kaldığı yerden devam (resume).
- Destek/K.Dışı otomatik atlanır. Hata olan kişi atlanır + ekran görüntüsü alınır + özet sonunda listelenir.

## Ayarlar (`dgs_poc.py` → `CONFIG`)
- `donem_value` — dönem (MAYIS 2026 = `151`).
- `daily_target_min` — **5 gün/45 saat → 540** (9:00); **6 gün/45 saat → 450** (7:30) + `include_saturday=True`.
- `prefer_start` — PDKS günlerinde dışarıda süreyi tercihen bu saatten itibaren boş slota koyar.
- `sel` — CSS seçicileri (gerekirse ayarla).

## Mantık (özet)
1. Dönem seç → Yeni → Personel + Proje + (çalışma türü 10766 default).
2. Grid her iş gününü otomatik 9 saate açar.
3. **PDKS günü:** ✎ aç → kesinleşmiş içeride aralıkları oku → `9:00 − içeride` kadar
   süreyi **çakışmayan** boş slota yaz → dialog Kaydet.
4. **PDKS'siz gün:** sadece tikle.
5. DRY-RUN değilse form Kaydet (taslak). Onaya gönder + e-imza insanda.

## Bilinen sınırlar (PoC)
- Eski tip ASP.NET MVC + jQuery SPA; bazı seçiciler ilk çalıştırmada ufak ayar isteyebilir
  (hata anında `dgs_hata_*.png` ekran görüntüsü alınır).
- Görevlendirme Türü (Lisansüstü/Arge/Diğer) şimdilik portal varsayılanında bırakılır —
  doğru değeri birlikte teyit edip `CONFIG`'e ekleyeceğiz.
- 6 gün çalışan / ay-içi giriş-çıkış / kısmi izin gibi kenar durumlar tam-ölçek aşamasında eklenecek.
