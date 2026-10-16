# ActivaterWindows

> **Terminal-first Microsoft activator.** Windows, Office ve ESU aktivasyonu — tek satır, kalıcı yöntemler, MAS gücüyle.

```powershell
irm https://raw.githubusercontent.com/tarikelertarnak/ActivaterWindows/main/ActivaterWindows.ps1 | iex
```

<p align="center">
  <img src="https://img.shields.io/badge/platform-Windows%2010%2F11-blue?style=flat-square" />
  <img src="https://img.shields.io/badge/version-2.1-orange?style=flat-square" />
  <img src="https://img.shields.io/badge/size-%7E30KB-success?style=flat-square" />
  <img src="https://img.shields.io/badge/engine-auto--tracked%20%2B%20SHA256-red?style=flat-square" />
  <img src="https://img.shields.io/badge/license-MIT-green?style=flat-square" />
</p>

---

## Neden v2.0 (Terminal Edition)?

v1.0 tek bir PyInstaller `.exe` idi (19 MB). v2.0 altyapı sıfırdan değişti:

| | v1.0 (exe) | v2.0 (terminal) |
|---|---|---|
| Boyut | 19 MB PE dosyası | **~25 KB** düz metin |
| Antivirus algılama | Yüksek (paketlenmiş exe + lisans API çağrıları = klasik heuristic tuzağı) | **Çok düşük** (paketleme yok, imzalı uzak kaynak + SHA256 doğrulama, her satır okunabilir) |
| Başlatma hızı | PyInstaller unpacking | Anında |
| MAS özellikleri | Gömülü, sürüm donmuş (v3.12) | Otomatik izlenen upstream + 2 yedek kademe — **tüm özellikler** (Troubleshoot, sürüm değiştirme, $OEM$ dahil) |
| Kurulum | Yok | Yok |

Antivirus notu: düz metin script + bütünlük doğrulaması, imzalı üç ayrı kaynak ve MAS'ın kendi resmi indirme yöntemini birebir kullanır — paketlenmiş binary olmadığı için reputation tabanlı taramalarda exe'ye göre çok daha az işaretlenir. Yine de 3. parti antivirus script indirmeyi engellerse script bunu algılar ve haber verir.

---

## Kullanım

### Tek Satır (Önerilen)

PowerShell'i aç, yapıştır:

```powershell
irm https://raw.githubusercontent.com/tarikelertarnak/ActivaterWindows/main/ActivaterWindows.ps1 | iex
```

Menü gelir. UAC izni gerektiğinde sorar.

### Yerel Dosya

1. `ActivaterWindows.cmd` + `ActivaterWindows.ps1` indir (aynı klasörde dursunlar)
2. `ActivaterWindows.cmd`'ye çift tıkla

### Doğrudan Yöntem (argümanlı)

```powershell
irm https://raw.githubusercontent.com/tarikelertarnak/ActivaterWindows/main/ActivaterWindows.ps1 | iex   # interaktif menü
.\ActivaterWindows.ps1 /HWID       # yöntem adıyla, menü olmadan
.\ActivaterWindows.ps1 -Probe      # indirme + hash doğrulama testi (hiçbir şey çalıştırmaz)
.\ActivaterWindows.ps1 -SelfTest   # yerleşik kontroller (ağ/kurulum/yetki gerektirmez)
```

---

## Menü

| # | Özellik | Süre |
|---|---|---|
| 1 | Windows 10/11 etkinleştirme | Kalıcı |
| 2 | Office etkinleştirme (LTSC/365/C2R) | Kalıcı |
| 3 | Windows + ESU + Office (tek geçiş) | Kalıcı |
| 4 | Windows 10 ESU | Kalıcı |
| 5 | Aktivasyon durumu kontrolü | — |
| 6 | Office LTSC Pro Plus 2024 kurulumu (dil + bileşen seçimi) | — |
| 7 | **HEPSINI ETKINLESTIR** — *kurulum yok.* Mevcut olan her şeyi kalıcı etkinleştirir (Windows + Office + Win10'da ESU) ve durum raporu verir | Kalıcı |
| 8 | **HEPSINI YAP** — Office eksikse kurar, sonra [7] akışını çalıştırır, en son tam MAS menüsünü teklif eder | Kalıcı |
| A / T / E | Arayüz dili: Otomatik / Türkçe / English | — |

> Menü yalnızca **ne yapacağını** yazar, **hangi yöntemle** yaptığını yazmaz. Yöntem adları arayüzde hiç görünmez.
> Online KMS (180 gün) kaldırıldı — kalıcı yöntemler var. İhtiyaç halinde **8**'deki tam MAS menüsünden erişilir.

### Görünmez etkinleştirme

Etkinleştirme arka planda, **hiçbir pencere açılmadan** çalışır. Motorun kendi çıktısı geçici bir günlüğe alınır; yalnızca sonuç satırları süzülüp gösterilir:

```
  Office is permanently activated.
  Windows 11 Pro is LICENSED.
```

Süzgeç, yöntem/marka adlarını ve cmd komut yankılarını kullanıcıya göstermez. UAC penceresi (yerinde açıksa) görünür olmaya devam eder — bu bir güvenlik sınırıdır ve gizlenemez.


### Office dili seçimi

Kurulumda 83 Office dil kodu **alfabetik numaralı liste** halinde gösterilir:

- `tr-tr` gibi **kodu** yaz → direkt kabul
- **Numara** yaz (ör. `76`) → "76) tr-tr — onaylıyor musun? Enter = evet, 0 = geri" onayı
- **Boş Enter** → otomatik algılanan dil kullanılır

### Office bileşen seçimi (3 mod)

| Mod | Ne kurar |
|---|---|
| **1 — Önerilen (hafif)** | Word, Excel, PowerPoint (Access, Outlook, Publisher, OneNote, OneDrive, Skype hariç) |
| **2 — Tam paket** | Tüm bileşenler |
| **3 — Özel Seçim** | 9 bileşeni tek tek işaretle |

Özel Seçim renkli işaret-kaldır listesi sunar:

```
  1) [X] Word          4) [ ] Outlook         7) [ ] OneNote
  2) [X] Excel         5) [ ] Access          8) [ ] OneDrive
  3) [X] PowerPoint    6) [ ] Publisher       9) [ ] Skype for Business
```

- Numaraları virgülle yaz, aç/kapa: `1,3,5`
- `a` = tümü · `n` = hiçbiri · `r` = önerilen · **Enter** = onayla
- Yeşil = kurulacak, gri = atlanacak. En az bir bileşen şart.


### İndirme göstergesi

Tüm indirmelerde gerçek yüzde çubugu + saniye sayacı:

```
  [############------------------]  40%    3s
```


### Güvenlik Mimarisi

Motor üç kademeli olarak çözülür; **her kademede SHA256 doğrulaması zorunludur**:

| Kademe | Kaynak | Neden |
|---|---|---|
| 1 | Upstream'in kendi bootstrap'ı (`get.activated.win`) → güncel commit + hash oradan okunur | Uygulama güncellenmese bile motor güncel kalır |
| 2 | Gövdeye gömülü pin (commit + hash) | Upstream formatı/argüman sırası değişirse çalışmaya devam eder |
| 3 | Son doğrulanmış yerel kopya (`%LOCALAPPDATA%\ActivaterWindows`) | Ağ/antivirus tamamen engellese bile doğrulanmış kopya kullanılır |

- **SHA256 doğrulama:** Hash uyuşmayan indirme **asla çalıştırılmaz** (bozulmuş/oynanmış içerik reddedilir).
- **3 ayna kaynak:** raw.githubusercontent → Azure DevOps → git.activated.win (rastgele sırayla denenir).
- **Argüman sırasından bağımsız:** yöntem argümanları tek geçişte gönderilir; motor hangi sırayla işlerse işlesin hepsi uygulanır.
- **Geçici dosya:** Çalıştırılan cmd `%TEMP%`'e yazılır, iş bitince silinir.
- **Yerleşik kontroller:** `-SelfTest` ağ/yetki gerektirmeden filtre, dil paritesi ve marka sızıntısını doğrular.

> Not: Bu kademeler *kaynak bulma* dayanıklılığıdır. Motor (MAS) GPL-3.0 lisanslıdır ve kaynak koda atıf README/disclaimer'da korunur — arayüzde gösterilmemesi lisansı etkilemez.


---

## Eski Sürüm (GUI exe)

v1.0 Tkinter GUI'si (`builder/`, `resources/`)legacy olarak korunur: [Releases](https://github.com/tarikelertarnak/ActivaterWindows/releases) sayfasından `ActivaterWindows.exe` indirilebilir. Yeni geliştirme terminal sürümündedir.

---

## Teknik Detay

| Bileşen | Kaynak |
|---|---|
| Aktivasyon motoru | [Microsoft Activation Scripts](https://github.com/massgravel/Microsoft-Activation-Scripts) (GPL-3.0); güncel commit upstream'den izlenir, yedek olarak `05c4f881` pini |
| Office Deployment Tool | [aka.ms/odt](https://aka.ms/odt) (kurulumda ihtiyaç anında indirilir) |
| Office LTSC Pro Plus 2024 | PerpetualVL2024 kanalı, x64 |

## Proje Yapısı

```
ActivaterWindows/
├── ActivaterWindows.ps1      # Terminal çekirdeği (tek dosya, ~25 KB)
├── ActivaterWindows.cmd      # Çift tık başlatıcısı
├── builder/                  # (legacy) v1.0 GUI exe kaynak kodu
├── resources/                # (legacy) ODT + MAS v3.12 gömülü kopyalar
└── mas-src/                  # (referans) MAS kaynak anlık görüntüsü
```

## Sık Sorulanlar

**Antivirus uyarısı verdi, ne yapmalıyım?** Windows Defender genellikle bırakır. 3. parti antivirus script indirmeyi engellerse: klasörü antivirus istisnasına ekle veya `irm … | iex` yöntemini kullan (dosya diske yazılmaz, only doğrulanıp cmd'ye geçilir).

**Offline çalışır mı?** Kısmen. İlk başarılı çalıştırmadan sonra motor `%LOCALAPPDATA%\ActivaterWindows`'a doğrulanmış olarak önbelleğe alınır; ağ yoksa bu kopya kullanılır. Hiç önbellek yoksa ve ağ da yoksa etkinleştirme yapılamaz (bütünlük doğrulaması gerektiği için). v1.0 GUI exe gömülü MAS ile her zaman offline çalışır.

---

## Disclaimer

This project is a wrapper around the community-developed **Microsoft Activation Scripts (MAS)** (GPL-3.0, [massgrave/Microsoft-Activation-Scripts](https://github.com/massgravel/Microsoft-Activation-Scripts)). Using unlicensed software is subject to your country's laws and Microsoft's terms of use. Educational and personal use only; usage responsibility rests entirely with the user.

**Made by:** [Tarik Eler](https://github.com/tarikelertarnak)

## License

MIT — see [LICENSE](LICENSE). MAS is GPL-3.0 and credited above.
