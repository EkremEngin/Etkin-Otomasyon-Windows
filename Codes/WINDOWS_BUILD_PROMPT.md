# WINDOWS BUILD PROMPT — Claude'a bu dosyayı ver

> Bu dosyayı Windows makinede Claude Code'a **olduğu gibi yapıştır** (veya "WINDOWS_BUILD_PROMPT.md'yi
> oku ve uygula" de). Kendi kendine yeter — mac'teki konuşmayı bilmene gerek yok.

---

## GÖREV

Bu projeyi Windows'ta `.exe` olarak derle ve **portala hiç dokunmadan** doğrula.

Kod 2026-07-16'da çift-platform hale getirildi ve **macOS'ta doğrulandı**, ama **Windows'ta HİÇ
çalıştırılmadı**. Senin işin: derlenip çalıştığını kanıtlamak, çıkan Windows'a özgü sorunları düzeltmek.

**Yeni özellik ekleme. Refactor etme. Sadece Windows'ta çalışmasını sağla.**

---

## PROJE NE

**Etkin Otomasyon** — Teknopark portallarında (Teknopark İstanbul, Bilişim Vadisi, Teknopark İzmir,
Yıldız, Ulutek) iki işi otomatikleştirir:
- **DGS**: Ar-Ge personelinin aylık gün/saat bildirimi (giriş + onaya gönderme + SGK raporuyla doğrulama)
- **İZİN**: yıllık izin girişi + belge yükleme + onay

Python + Playwright. **Kendi tarayıcısını indirmez** — kullanıcının Chrome'una `connect_over_cdp`
(localhost:9222) ile bağlanır. GUI: customtkinter/Tkinter. Son kullanıcı Python bilmiyor → tek `.exe`.

**Mimari:** `izin_app.py` tek giriş noktası; ilk argümana (token) göre yönlendirir.
Token yok → GUI. `paths` → teşhis. `dgs --park <KOD>` → DGS motoru. `orchestrator`/`uploader`/`logincheck`
→ izin motorları. Donmuş exe **kendini** token'la çağırır (içinde `.py` dosyası yok) — `izin_frozen.py`.

⚠ Devir dokümanları (`EXE_DEVIR.md`, `EXE_BUILD.md`, `PRIME_SCRIPTS.md`) **depoda yok** — içlerinde
çalışan adları geçtiği için alınmadı. Bu prompt kendi kendine yeter; gerekirse kullanıcıdan iste.

---

## ÖNKOŞULLAR

- **Python 3.12+** — python.org'dan (tkinter dahil gelir; Microsoft Store sürümünü KULLANMA, tkinter sorunlu)
- **Google Chrome kurulu olmalı** (uygulama ona bağlanıyor)
- `playwright install` **GEREKMEZ** (tarayıcı indirmiyoruz)

```powershell
cd Codes                                          # ← tüm kod burada
python -m pip install pyinstaller -r requirements.txt
```

⚠ `pip install pyinstaller playwright openpyxl` **YETMEZ** — spec `collect_all('customtkinter')` çağırıyor.
Mutlaka `-r requirements.txt` kullan.

---

## BUILD

```powershell
cd Codes                                          # spec ve assets/ buraya göreli — başka yerden çalışmaz
pyinstaller izin_app.spec
```

Çıktı: `Codes\dist\IzinOtomasyonu.exe` (tek dosya, `console=False`).
PyInstaller **çapraz-derleme yapmaz** — bu yüzden Windows'ta derliyoruz.

### 🔴 exe'yi nereye koyacaksın (build'den SONRA)

`Codes\dist\` **kullanım yeri DEĞİL**, sadece build çıktısı. exe'yi oradan **bir üst klasöre** (repo kökü,
`Codes`'un yanına) kopyala ve oradan çalıştır.

**Neden:** uygulama resume/done dosyalarını **kendi durduğu klasöre** yazar (`izin_frozen.data_dir()`).
`dist\` içinden çalıştırırsan kökteki `dgs_done_*.txt` / `izin_done_*.txt` dosyalarını **göremez** →
girilmiş kişileri atlamaz → **portala MÜKERRER KAYIT**. Kural: *exe nerede duruyorsa resume orada.*

macOS'taki düzen aynen şöyle (Windows'ta da böyle olmalı):
```
proje-kökü/
├── Etkin Otomasyon.app / IzinOtomasyonu.exe   ← operatör buna tıklar, resume BUNUN yanına yazılır
├── Codes/                                      ← kod + build (dist/ dahil)
├── dgs_done_*.txt · izin_done_*.txt            ← resume; exe ile AYNI klasör olmalı
└── *.xlsx                                      ← operatörün seçtiği Excel'ler
```
Teşhis: `IzinOtomasyonu.exe paths` → `data_dir` resume dosyalarının klasörünü göstermeli.

---

## DOĞRULAMA — sırayla, portala DOKUNMADAN

> Aşağıdaki komutlar **`Codes\` klasöründen** çalıştırılır (build çıktısı `Codes\dist\` altında).
> Bunlar sadece derleme testi — bitince exe'yi yukarıdaki gibi bir üst klasöre taşımayı unutma.

### 1) Dispatcher + kalıcı klasör
```powershell
dist\IzinOtomasyonu.exe paths
```
JSON basmalı. **Şunlar şart:**
- `"frozen": true`
- `"data_dir_gecici_mi": false`  ← **KRİTİK.** `true` ise resume dosyaları geçici `_MEIPASS`'e yazılıyor
  demektir; çıkışta silinir → ikinci koşu kimseyi atlamaz → **MÜKERRER PORTAL KAYDI**. Böyleyse DUR ve düzelt.

macOS referans çıktısı:
```json
{ "frozen": true,
  "data_dir": "/Users/.../dgs-otomasyon",
  "script_dir": "/var/folders/.../T/_MEIcABCXT",   ← geçici olması NORMAL
  "data_dir_gecici_mi": false }
```

### 2) Gerçek Excel okuma + Türkçe karakter (salt-okur, portala dokunmaz)
```powershell
dist\IzinOtomasyonu.exe dgs --park TPI liste --sheet "Mayıs" --excel "C:\...\05-TEKNOKENTLER - MAYIS 2026 - FİNAL_9SAAT (1).xlsx"
```
`<<<LISTE>>>{...}` satırı basmalı, çıkış kodu 0.

**macOS referansı (birebir aynı çıkmalı):** TPI **129** kişi · BV **44** · Yıldız **34**.

⚠ **Türkçe karakter kontrolü:** çıktıdaki isimlerde `Ğ İ Ş Ç Ö Ü` harfleri **bozulmamış** olmalı.
`�` veya `Ä°`/`Åž` gibi bir şey görürsen kodlama hatası var (aşağıdaki cp1254 notuna bak).
Çıktı **gerçek çalışan verisi** içerir (ad + T.C.) — **loglama, paylaşma, dosyaya yazma.** Sadece bak ve
"bozulma var/yok" diye raporla.

Bu test aynı anda şunları kanıtlar: PyInstaller'ın `importlib`-yüklü park modüllerini bundle etmesi,
openpyxl, ve UTF-8 çözümlemesi.

### 3) GUI
`dist\IzinOtomasyonu.exe`'yi çift tıkla → pencere açılmalı (boş/siyah değil, kartlar görünür).
Log kutusunda **"Hazır. Resume/kayıt klasörü: ..."** yazar — o yol done `.txt`'lerin bulunduğu klasör olmalı.

### 4) Chrome bulunuyor mu
GUI'de **"Chrome'u aç"** → debug portlu Chrome açılmalı. Açılmazsa `izin_gui.py` içindeki `find_chrome()`
Windows dalına bak (Program Files / Program Files (x86) / LocalAppData).

**BURAYA KADAR. Login/Cloudflare/gerçek koşu YAPMA** (aşağıdaki "YAPMA" bölümü).

---

## WINDOWS'TA BEKLENEN RİSKLER (ilk kez koşuyor)

1. **Playwright node driver** — `collect_all('playwright')` Windows'ta `node.exe` toplamalı. Test 2 bunu
   kanıtlar; `ModuleNotFoundError`/driver hatası çıkarsa `--collect-all playwright` veya hook'a bak.
2. **İkon** — spec `icon='assets/etkn-app-icon.png'`. PyInstaller Windows'ta `.ico` ister ama Pillow ile
   PNG'yi kendisi çevirir (Pillow `requirements.txt`'te var). Patlarsa PNG'yi `.ico`'ya çevirip spec'i güncelle.
3. **Konsol penceresi fırlarsa** — `console=False` ve alt-süreç exe'nin kendisi olduğu için fırlamamalı.
   Fırlarsa `subprocess` çağrılarına `creationflags=subprocess.CREATE_NO_WINDOW` ekle (`izin_gui.py`).
4. **Defender / SmartScreen** — imzasız exe'yi uyarır → "Yine de çalıştır". Normal, hata değil.
5. **`data_dir`** — kural: **"exe nerede duruyorsa resume dosyaları orada."** exe'yi `dist\` içinde bırakma;
   `dgs_done_*.txt` / `izin_done_*.txt` dosyalarının yanına (proje köküne) kopyala. `ETKN_DATA_DIR` env ezer.

---

## WINDOWS İÇİN ZATEN YAPILAN DÜZELTMELER (tekrar etme, geri alma)

2026-07-16'da mac'te yapıldı ve doğrulandı:
- `dgs_rapor_kontrol.py` — `/tmp/sgk_*.csv` → `tempfile.gettempdir()`.
  (Windows'ta `/tmp` = `C:\tmp`, yoktur → `FileNotFoundError` → DGS kapanış kontrolü çökerdi.)
- `izin_gui.py` ×3 (`_probe_login`, `_dgs_kisileri_yukle`, ana worker `Popen`) — `text=True`'ya
  **`encoding="utf-8", errors="replace"`** eklendi. `text=True` sistem kodlamasını kullanır: mac'te UTF-8
  (çalışır), **Türkçe Windows'ta cp1254** → "Ş"nin `0x9E` baytı cp1254'te TANIMSIZ → `UnicodeDecodeError`
  → canlı-log thread'i ölür, koşu "Çalıştırma hatası" ile yarıda düşer. **Bu yüzden yeni `subprocess`
  çağrısı yazarsan `encoding="utf-8"` vermeyi UNUTMA.**
- `izin_otomasyon.py:346` — uploader `cwd=SCRIPT_DIR` ile açılıyordu → donmuşta `_MEIPASS`
  → `cwd=izin_frozen.data_dir()` oldu. Ölü `SCRIPT_DIR` sabiti silindi.
- `EXE_BUILD.md` — pip satırı `-r requirements.txt`.

Zaten çapraz-platform olan yerler (dokunma): `find_chrome()` win dalı, `izin_app._repair_stdio()`
(Windows'ta `console=False` → `sys.stdout is None` çökmesini onarır), spec'teki `if sys.platform == "darwin"`
guard'ı, `izin_frozen.data_dir()`, tüm `open()`'larda açık `encoding="utf-8"`.

---

## YAPMA

- **PRIME scriptleri değiştirme:** `dgs_poc.py`, `dgs_onaya.py`, `dgs_rapor_kontrol.py`,
  `izin_onaya_dosyali.py`, `izin_onaya_dosyali_yildiz.py` (bkz. `PRIME_SCRIPTS.md`). Bunlar canlı portalda
  doğrulanmış, DONMUŞ motorlar. Kullanıcı **açıkça istemedikçe** dokunma. Windows için zorunlu bir
  düzeltme gerekiyorsa önce kullanıcıya sor.
- **Gerçek portal koşusu yapma.** Login + Cloudflare **insan** gerektirir; portala yazmak resmi kayıt
  üretir ve geri alması zordur. İlk gerçek koşu **gözetimli**, **tek park**, **az kişi** ile yapılır.
- Excel/done `.txt` dosyalarını düzenleme — operasyonel veri.
- Test için `--commit` bayrağını kullanma; `liste` ve `paths` yeterli.

---

## KODU NEREDEN ALACAKSIN

```powershell
git clone https://github.com/EkremEngin/Etkin_Otomasyon.git
cd Etkin_Otomasyon
```
(private depo — `gh auth login` veya GitHub kimliği gerekir)

Depoda **yalnız `Codes/` klasörü** vardır — tüm kod, spec ve assets orada. Repo kökü (yani `Codes`'un
üstü) build'den sonra exe'nin ve resume dosyalarının duracağı yerdir; şimdilik boş olması normal.

### 🔴 Depoda OLMAYAN, ama lazım olan: test Excel'i

Depoya **bilerek yalnızca kod** konuldu. Çalışan verisi, portal şifreleri, loglar, ekran görüntüleri
ve done/resume dosyaları **KVKK gereği depoya alınmadı** — `.gitignore` beyaz-liste mantığında yazıldı
(varsayılan: her şey hariç).

Doğrulama testi #2 için Excel'i kullanıcı sana **ayrıca** verecek (USB/güvenli aktarım):
`05-TEKNOKENTLER - MAYIS 2026 - FİNAL_9SAAT (1).xlsx`

⚠ **Bu Excel ~300 çalışanın adı + T.C. kimlik numarasını içerir.** Depoya ekleme, commit'leme,
bir yere kopyalama, içeriğini loglama/yapıştırma. Test bitince olduğu yerde bırak.

### Depoda olmayan dokümanlar

`EXE_DEVIR.md`, `EXE_BUILD.md`, `PRIME_SCRIPTS.md`, `CANLI_TEST_PROMPT.md` mac'te kaldı (içlerinde
çalışan adları geçiyor). İhtiyacın olursa kullanıcıdan iste. Bu prompt kendi kendine yeter —
PRIME dosya listesi ve build adımları yukarıda zaten var.

### Gerçek koşu yapılacaksa

`izin_done_*` / `dgs_done_*` resume dosyaları depoda YOK. Windows'ta **gerçek** koşu yapılacaksa
kullanıcıdan istenmeli — yoksa otomasyon zaten girilmiş kişileri tekrar girer → **mükerrer portal kaydı**.
(Derleme + doğrulama testleri için gerekmez.)

---

## BİTİRİNCE RAPORLA

- `pyinstaller` çıkış kodu + hata/uyarı var mıydı
- Test 1 JSON'u (özellikle `data_dir_gecici_mi`)
- Test 2 kişi sayıları (TPI/BV/Yıldız) — mac'le (129/44/34) uyuşuyor mu, Türkçe bozuldu mu
- GUI açıldı mı, Chrome bulundu mu
- Windows'a özgü düzelttiğin her şey (dosya + satır + neden)
