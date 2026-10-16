# ActivaterWindows

> **Terminal-first Microsoft activator.** Windows, Office ve ESU aktivasyonu — tek satır, kalıcı yöntemler, MAS gücüyle.

```powershell
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; irm https://raw.githubusercontent.com/tarikelertarnak/ActivaterWindows/main/ActivaterWindows.ps1 | iex
```

<p align="center">
  <img src="https://img.shields.io/badge/platform-Windows%2010%2F11-blue?style=flat-square" />
  <img src="https://img.shields.io/badge/version-2.1.3-orange?style=flat-square" />
  <img src="https://img.shields.io/badge/size-%7E44KB-success?style=flat-square" />
  <img src="https://img.shields.io/badge/engine-auto--tracked%20%2B%20SHA256-red?style=flat-square" />
  <img src="https://img.shields.io/badge/deps-none%20required-lightgrey?style=flat-square" />
  <img src="https://img.shields.io/badge/license-MIT-green?style=flat-square" />
  <a href="https://github.com/tarikelertarnak/ActivaterWindows/releases/latest"><img src="https://img.shields.io/github/v/release/tarikelertarnak/ActivaterWindows?style=flat-square&label=release" /></a>
</p>

---

## Neden v2.0 (Terminal Edition)?

v1.0 tek bir PyInstaller `.exe` idi (19 MB). v2.0 altyapı sıfırdan değişti:

| | v1.0 (exe) | v2.0 (terminal) |
|---|---|---|
| Boyut | 19 MB PE dosyası | **~44 KB** düz metin |
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
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; irm https://raw.githubusercontent.com/tarikelertarnak/ActivaterWindows/main/ActivaterWindows.ps1 | iex
```

Menü gelir. UAC izni gerektiğinde sorar.

### Yerel Dosya

1. `ActivaterWindows.cmd` + `ActivaterWindows.ps1` indir (aynı klasörde dursunlar)
2. `ActivaterWindows.cmd`'ye çift tıkla

### Doğrudan Yöntem (argümanlı)

```powershell
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; irm https://raw.githubusercontent.com/tarikelertarnak/ActivaterWindows/main/ActivaterWindows.ps1 | iex   # interaktif menü
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

### Her temiz Windows'ta çalışır

Ek bir şey kurmaya gerek yoktur:

| | |
|---|---|
| Gereken | Windows 10/11 (PowerShell 5.1 **sistemle gelir**), internet, yönetici yetkisi (UAC) |
| Kurulan | Hiçbir şey. Ne .NET paketi, ne çalışma zamanı, ne `choco`, ne derleyici |
| Yazılan | Geçici bir `.cmd` + günlüğü; işlem bitince silinir. Kalıcı dosya **yok** |
| Eski Windows | PowerShell 2.0 / Windows 7-8 desteklenmez (TLS 1.2 zorunlu) |

Komut `irm | iex` ile çalıştığı için `ExecutionPolicy` ayarı da gerekmez. TLS 1.2 komutun içinde açıkça zorlanır; eski sürüm/varsayılanlar yüzünden `irm` bağlantı hatası alan makinelerde bu satır sorunu çözer.

### Görünmez etkinleştirme

Etkinleştirme arka planda, **hiçbir pencere açılmadan** çalışır. Motorun kendi çıktısı geçici bir günlüğe alınır; yalnızca sonuç satırları süzülüp gösterilir:

```
  Windows 11 Pro is LICENSED.
  Office is permanently activated.
```

Süzgeç yalnızca **gizlilik** içindir, güvenlik sınırı değildir: yöntem/marka adları, kaynak adresleri, lisans metinleri ve cmd komut yankıları düşürülür, sonuc/uyarı satırları korunur.

**Zaman sınırı:** motorun bazı akışları sonda bir tuşa bekler (`choice /c 0 /n`). Gizli pencerede o tuş bekleyemez, yani süreç sonsuza kadar asılı kalırdı — bu yüzden gizli çalıştırmada **900 sn** üst sınırı uygulanır: süre dolarsa süreç kapatılır, günlüğe yazılan sonuçlar yine gösterilir ve kullanıcıya "işlem tamamlanmamış olabilir" uyarısı verilir. Filtre hiç satır bırakmazsa da "sonuç satırı yok" mesajı basılır; böylece **boş ekran ile başarısızlık birbirine karışmaz**.

UAC penceresi görünür kalır; bu bir güvenlik sınırıdır ve gizlenemez. UAC onaylanmadan hiçbir şey çalışmaz.


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

Motor üç kademeli olarak çözülür. İlk iki kademe **bilinen SHA256** ile doğrulanır; üçüncü kademe, kaydedildiği anda alınan hash ile yeniden doğrulanır:

| Kademe | Kaynak | Doğrulama | Neden |
|---|---|---|---|
| 1 | Upstream'in kendi bootstrap'ı (`get.activated.win`) → o anki commit + hash | Bootstrap'tan okunan SHA256 | Uygulama güncellenmese bile motor güncel kalır |
| 2 | Gövdeye gömülü pin (commit + hash) | Gömülü SHA256 | Upstream formatı/argüman sırası değişirse çalışmaya devam eder |
| 3 | Yerel önbellek (`%LOCALAPPDATA%\ActivaterWindows`) | Kayıt anındaki SHA256 | Ağ/antivirus tamamen engellese bile çalışır |

- **SHA256 doğrulama:** İlk iki kademede hash uyuşmayan içerik **asla çalıştırılmaz** — bozulmuş/oynanmış indirme reddedilir ve alt kademeye düşülür.
- **3 ayna kaynak:** raw.githubusercontent → Azure DevOps → git.activated.win (rastgele sırayla denenir).
- **Argüman sırasından bağımsız:** yöntem argümanları tek geçişte gönderilir; motor hangi sırayla işlerse işlesin hepsi uygulanır.
- **Geçici dosya:** Çalıştırılan cmd geçici klasöre yazılır (yönetici ise `%SystemRoot%\Temp`, değilse `%TEMP%`), iş bitince silinir.
- **Yerleşik kontroller:** `-SelfTest` ağ/yetki gerektirmeden filtre, dil paritesi ve marka sızıntısını doğrular; `-Probe` motoru indirip hash'ini doğrular.

> **Dürüst sınır (3. kademe):** Yerel önbellek kendi yanındaki `.sha256` dosyasıyla doğrulanır; bu **bozulmayı ve sonradan değiştirilmeyi** yakalar, ama aynı kullanıcı olarak çalışan bir saldırganın iki dosyayı birlikte değiştirmesini engellemez. Zaten yönetici yetkisi olan bir aktör bu kapsamın dışındadır. Bu kademe bir dayanıklılık yedeğidir, bir güven sınırı değildir.

> Not: Motor (MAS) GPL-3.0 lisanslıdır ve kaynak koda atıf README/disclaimer'da korunur — arayüzde gösterilmemesi lisansı etkilemez.


---

## Eski Sürüm (GUI exe — arşiv)

v1.0 Tkinter GUI'si (`builder/`, `resources/`) legacy olarak korunur ama **artık geliştirilmiyor** ve
paketlenmiş `.exe` olduğu için antivirüs heuristic'lerini tetikleyebilir. Yalnızca tarihsel arşivdir:
[v1.0.0 (arşiv)](https://github.com/tarikelertarnak/ActivaterWindows/releases/tag/v1.0.0).

**Kullanman gereken sürüm:** [v2.1.3 — Terminal Edition](https://github.com/tarikelertarnak/ActivaterWindows/releases/tag/v2.1.3)
(yukarıdaki tek satır). `ActivaterWindows.exe` indirmene gerek yok.

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
