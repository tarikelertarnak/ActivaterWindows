# ActivaterWindows

> **Microsoft Office LTSC Professional Plus 2024'ü tek tıkla kurun; Windows, Office ve ESU aktivasyonlarını tek arayüzden yönetin.**

`ActivaterWindows`, Office Deployment Tool'u (ODT) sararak kurulumu otomatikleştiren ve **Microsoft Activation Scripts (MAS v3.12)** ile aktivasyonu tek tıkla yapan kullanıcı dostu bir **portable Windows masaüstü uygulamasıdır**. Tek `.exe` dosyası — kurulum gerekmez, silindiğinde sıfır kalıntı.

<p align="center">
  <img src="https://img.shields.io/badge/platform-Windows%2010%2F11-blue?style=flat-square" />
  <img src="https://img.shields.io/badge/version-1.0.0-orange?style=flat-square" />
  <img src="https://img.shields.io/badge/python-3.8+-3776AB?style=flat-square&logo=python&logoColor=white" />
  <img src="https://img.shields.io/badge/UI-Tkinter-orange?style=flat-square" />
  <img src="https://img.shields.io/badge/license-MIT-green?style=flat-square" />
  <img src="https://img.shields.io/badge/ODT-16.0.20228-brightgreen?style=flat-square" />
  <img src="https://img.shields.io/badge/MAS-3.12-red?style=flat-square" />
</p>

---

## Ozellikler

| Ozellik | Aciklama |
|---|---|
| **Portable** | Tek `.exe`, kurulum gerekmez, her bilgisayarda calisir |
| **Sifir Kalıntı** | Silindiginde temp dosyalari, loglar, registry temizlenir |
| **Tek Tıkla Kurulum** | Office LTSC Professional Plus 2024 - istedigin dilde tek tıkla |
| **Ozel Konfigurasyon** | Kendi `Configuration.xml` dosyani kullan |
| **38+ Dil Desteği** | tr-tr, en-us, de-de, fr-fr ve daha fazlasi |
| **Aydınlık/Karanlık Tema** | Modern, sistem temasi otomatik algilama |
| **Gomulu Aktivasyon (MAS)** | Windows / Office / ESU aktivasyonu script internetten inmez, exe icinde gomuludur |
| **Aktivasyon Tipi Secimi** | Windows 11 Pro / for Workstations / 11 / 10 Pro / 10 + Office — 16 yontem |
| **Aktivasyon Durum Kontrolu** | Tek tıkla mevcut aktivasyon durumunu gosterir |
| **Otomatik ODT Indirme** | Eger setup.exe yoksa ODT'yi otomatik indirir |
| **Internet Kontrolu** | Kurulumdan once internet baglantisini dogrular |

---

## Kullanim

### Portable Kullanim (Onerilen)

1. `ActivaterWindows.exe` dosyasini indirin
2. Calistirin - hepsi bu kadar! **Kurulum gerekmez.**
3. Gerekli dosyalar otomatik olarak indirilir ve kurulur
4. Silmek icin `ActivaterWindows.exe`'yi silmeniz yeterli — **sifir kalıntı**

### Kaynaktan Calistirma

```powershell
git clone https://github.com/tarikelertarnak/ActivaterWindows.git
cd ActivaterWindows
python builder/office_setup_wizard.py
```

### Derleme (Standalone .exe)

```powershell
powershell -ExecutionPolicy Bypass -File builder/build.ps1
```

Derlenen yurutebilir dosya `dist/` klasorunde olusur.

---

## Kullanim Adimlari

### Sekme 1: Office Kurulumu
1. **Kurulum Modu** sec: *Otomatik* (LTSC 2024) veya *Ozel Configuration*
2. Kurulum dilini sec (orn. `tr-tr`, `en-us`)
3. **Kurulumu Baslat** butonuna bas
4. Office indirilir, kurulur ve (istersen) Ohook ile etkinlestirilir

### Sekme 2: Aktivasyon
1. **Aktivasyon Turu** sec (combobox):
   - `Windows 11 Pro — Kalıcı (HWID)`
   - `Windows 11 Pro for Workstations — Kalıcı (HWID)`
   - `Windows 11 — Kalıcı (HWID)`
   - `Windows 10 Pro — Kalıcı (HWID)`
   - `Windows 10 Home / Windows 10 — Kalıcı (HWID)`
   - `Office LTSC/365 — Kalıcı (Ohook)`
   - `Windows 7-11 + Server — Kalıcı (TSforge)`
   - ve diger secenekler (ESU, Online KMS, kaldirma)
2. **Aktifleştir** butonuna bas
3. Istersen **Aktivasyon Durumunu Kontrol Et** ile dogrula

### Temizlik

Uygulama kapatildiginda olusturulan tum gecici dosyalar otomatik olarak temizlenir:
- `%TEMP%\ActivaterWindows\` klasoru
- `%TEMP%\oi_debug.log` debug logu
- `%TEMP%\office_config.xml` konfigurasyon
- `%TEMP%\MAS_*.cmd` / `%TEMP%\CheckAct_*.cmd` aktivasyon scriptleri (kullanim sonrasi silinir)

`ActivaterWindows.exe`'yi sildiginizde bilgisayarda hicbir kalıntı kalmaz.

---

## Teknik Detaylar

| Bilesen | Surum | Kaynak |
|---|---|---|
| Office Deployment Tool | 16.0.20228.20124 | [aka.ms/odt](https://aka.ms/odt) |
| Microsoft Activation Scripts | v3.12 | [massgrave.dev](https://massgrave.dev) |
| Office LTSC Pro Plus 2024 | Volume License | PerpetualVL2024 kanali |

### Desteklenen Uygulamalar

Word, Excel, PowerPoint, Outlook, OneNote, Access, Publisher, Skype for Business, OneDrive, Visio, Project

### Desteklenen Diller

38+ dil: Turkce, English, Deutsch, Francais, Espanol, Italiano, Portugues, Nederlands, Polski, Rus, Ukraynaca, Arapca, Hintce, Bahasa, Tayca, Vietnamca, Cince, Japonca, Korece ve daha fazlasi.

### Aktivasyon Yontemleri (MAS v3.12)

| Yontem | Kapsam | Sure |
|---|---|---|
| HWID | Windows 10/11 | Kalici (dijital lisans) |
| TSforge | Windows 7-11 + Server | Kalici |
| ESU | Windows 10 (ESU) | Kalici |
| Ohook | Office LTSC/365 | Kalici |
| Online KMS | Windows / Office | 180 gun (oto-yenileme gorevi) |

---

## Proje Yapisi

```
ActivaterWindows/
├── ActivaterWindows.exe                # Derlenmis portable uygulama
├── builder/
│   ├── build.ps1                   # .exe derleme betigi
│   ├── office_setup_wizard.py      # Ana uygulama
│   ├── office_setup_wizard.spec    # PyInstaller spec (MAS scripti gomulur)
│   ├── office_setup_wizard.ico     # Uygulama simgesi
│   ├── version_info.txt            # Windows surum bilgisi
│   └── icons/                      # Tema simgeleri
│       ├── light/                  # Aydinlik tema ikonlari
│       └── dark/                   # Karanlik tema ikonlari
├── resources/
│   ├── setup.exe                   # Office Deployment Tool
│   ├── Configuration.xml           # Varsayilan konfigurasyon
│   ├── configuration-Office365-x64.xml
│   └── mas/
│       ├── MAS_AIO.cmd             # Gomulu aktivasyon scripti (v3.12)
│       └── Check_Activation_Status.cmd
├── LICENSE                         # MIT Lisans
└── README.md
```

---

## Sorumluluk Reddi

Bu arac resmi Microsoft araclari (ODT) uzerine kuruludur. Aktivasyon ozelligi, topluluk tarafindan gelistirilen **Microsoft Activation Scripts (MAS)** projesini kullanir (GPL-3.0, [massgrave/Microsoft-Activation-Scripts](https://github.com/massgravel/Microsoft-Activation-Scripts)). Lissansiz yazilim kullanimi ulkenizin yasalarina ve Microsoft kullanim kosullarina tabidir. Bu proje yalnizca egitim ve kisisel kullanim amaciyla paylasilmaktadir; kullanim sorumlulugu tamamen kullanicidir.

---

## Katkida Bulun

Hata bildirimleri ve gelistirme onerileri icin [Issues](https://github.com/tarikelertarnak/ActivaterWindows/issues) sayfasini kullanabilirsin.

**Yapan:** [Tariq Eler](https://github.com/tarikelertarnak)

---

## Lisans

MIT License - Detaylar icin [LICENSE](LICENSE) dosyasina bakin. Gomulu MAS scripti GPL-3.0 lisansli [Microsoft Activation Scripts](https://github.com/massgravel/Microsoft-Activation-Scripts) projesinden alinmistir ve kaynak olarak belirtilir; aktivasyon sekmesi bu scriptin kilifidir.