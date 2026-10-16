# ActivaterWindows

> **Terminal-first Microsoft activator.** Windows, Office ve ESU aktivasyonu — tek satır, kalıcı yöntemler, MAS gücüyle.

```powershell
irm https://raw.githubusercontent.com/tarikelertarnak/ActivaterWindows/main/ActivaterWindows.ps1 | iex
```

<p align="center">
  <img src="https://img.shields.io/badge/platform-Windows%2010%2F11-blue?style=flat-square" />
  <img src="https://img.shields.io/badge/version-2.0-orange?style=flat-square" />
  <img src="https://img.shields.io/badge/size-%7E25KB-success?style=flat-square" />
  <img src="https://img.shields.io/badge/MAS-pinned%20%2B%20SHA256-red?style=flat-square" />
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
| MAS özellikleri | Gömülü, sürüm donmuş (v3.12) | Pinned commit'ten doğrulanmış indirme — **tüm MAS özellikleri** (Troubleshoot, sürüm değiştirme, $OEM$ dahil) |
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
.\ActivaterWindows.ps1 /HWID      # Windows HWID, menü olmadan
.\ActivaterWindows.ps1 /Ohook     # Office Ohook
.\ActivaterWindows.ps1 -Probe     # indirme + hash doğrulama testi (hiçbir şey çalıştırmaz)
```

---

## Menü

| # | Özellik | Yöntem | Süre |
|---|---|---|---|
| 1 | Windows 10/11 etkinleştirme | **HWID** (dijital lisans) | Kalıcı |
| 2 | Office etkinleştirme (LTSC/365/C2R) | **Ohook** | Kalıcı |
| 3 | Windows + ESU + Office | **TSforge** | Kalıcı |
| 4 | Windows 10 ESU | **TSforge** | Kalıcı |
| 5 | Aktivasyon durumu | MAS Check | — |
| 6 | **TAM MAS menüsü** | Tüm özellikler: Troubleshoot, Windows/Office sürüm değiştirme, $OEM$ klasörü, changelog | — |
| 7 | Office LTSC Pro Plus 2024 kurulumu | ODT + otomatik Ohook teklifi | — |
| 8 | **HEPSINI YAP** | Ofis yoksa kurar → Windows + ESU + Office TSforge ile tek geçişte → durum raporu | Kalıcı |
| L | Arayuz dili (Otomatik / Turkce / English) | — | — |

> Online KMS (180 gün) menüden kaldırıldı — kalıcı TSforge varken gereksiz. MAS'ın Tam Menüsü'nden (6) erişilebilir.

### Güvenlik Mimarisi

- **Pinned commit:** MAS AIO her zaman `05c4f881…` commit'inden indirilir — sürüm sürprizi yok.
- **SHA256 doğrulama:** İndirilen her script, bilinen hash ile karşılaştırılır; uyuşmazsa **çalıştırılmaz**.
- **3 ayna kaynak:** raw.githubusercontent → Azure DevOps → git.activated.win (rastgele sırayla denenir).
- **Geçici dosya:** Çalıştırılan cmd `%TEMP%`'e yazılır, iş bitince silinir.

---

## Eski Sürüm (GUI exe)

v1.0 Tkinter GUI'si (`builder/`, `resources/`)legacy olarak korunur: [Releases](https://github.com/tarikelertarnak/ActivaterWindows/releases) sayfasından `ActivaterWindows.exe` indirilebilir. Yeni geliştirme terminal sürümündedir.

---

## Teknik Detay

| Bileşen | Kaynak |
|---|---|
| Aktivasyon motoru | [Microsoft Activation Scripts](https://github.com/massgravel/Microsoft-Activation-Scripts) (GPL-3.0), pinned `05c4f881` |
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

**Offline çalışır mı?** Hayır — aktivasyon scripti pinned kaynaktan indirilir (bütünlük için). v1.0 GUI exe gömülü MAS ile offline çalışır.

---

## Disclaimer

This project is a wrapper around the community-developed **Microsoft Activation Scripts (MAS)** (GPL-3.0, [massgrave/Microsoft-Activation-Scripts](https://github.com/massgravel/Microsoft-Activation-Scripts)). Using unlicensed software is subject to your country's laws and Microsoft's terms of use. Educational and personal use only; usage responsibility rests entirely with the user.

**Made by:** [Tarik Eler](https://github.com/tarikelertarnak)

## License

MIT — see [LICENSE](LICENSE). MAS is GPL-3.0 and credited above.
