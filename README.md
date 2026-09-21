# Office Setup Wizard

> **Microsoft Office LTSC Professional Plus 2024'ü tek tıkla kurun ve etkinleştirin.**

`Office Setup Wizard`, Office Deployment Tool'u (ODT) sararak kurulumu otomatikleştiren, kullanıcı dostu bir **portable Windows masaüstü uygulamasıdır**. Karmaşık XML yapılandırmalarıyla uğraşmadan, grafik arayüzden birkaç tıklamayla Office'i kurar.

<p align="center">
  <img src="https://img.shields.io/badge/platform-Windows%2010%2F11-blue?style=flat-square" />
  <img src="https://img.shields.io/badge/version-2.0.0-orange?style=flat-square" />
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
| **Tek Tıkla Kurulum** | Office LTSC Professional Plus 2024 - istedigin dilde tek tıkla kurulum ozel konfigurasyon |
| **Ozel Konfigurasyon** | Kendi `Configuration.xml` dosyani kullan (ileri duzey kullanicilar icin) |
| **38+ Dil Desteği** | tr-tr, en-us, de-de, fr-fr ve daha fazlasi |
| **Aydınlık/Karanlık Tema** | Modern, tercihine gore degisen arayuz + sistem temasi otomatik algilama |
| **Otomatik Aktivasyon** | Kurulum sonrasi MAS v3.12 ile Ohook aktivasyonu |
| **Detaylı Gunluk** | Kurulum surecini adim adim izleme ve hata kaydi |
| **Otomatik ODT Indirme** | Eger setup.exe yoksa ODT'yi otomatik indirir |
| **Internet Kontrolu** | Kurulumdan once internet baglantisini dogrular |
| **Hakkinda Paneli** | Surum, MAS ve ODT bilgilerini gosterir |
| **Portable** | Tek .exe, kurulum gerekmez, her bilgisayarda calisir |

---

## Kullanim

### On Kosullar
- **Windows 10 / 11** (x64)
- Internet baglantisi (Office dosyalarini indirmek icin)

### Portable Kullanim (Onerilen)

1. `OfficeSetupWizard.exe` dosyasini indirin
2. Calistirin - hepsi bu kadar! **Kurulum gerekmez.**
3. Gerekli dosyalar otomatik olarak indirilir ve kurulur

### Kaynaktan Calistirma

```powershell
git clone https://github.com/tarikelertarnak/OfficeSetupWizard.git
cd OfficeSetupWizard
python builder/office_setup_wizard.py
```

### Derleme (Standalone .exe)

```powershell
powershell -ExecutionPolicy Bypass -File builder/build.ps1
```

Derlenen yurutebilir dosya `dist/` klasorunde olusur.

---

## Kullanim Adimlari

1. **Kurulum Modu** sec: *Otomatik* (LTSC 2024) veya *Ozel Configuration*
2. Kurulum dilini sec (orn. `tr-tr`, `en-us`)
3. **Kurulumu Baslat** butonuna bas
4. Office indirilir, kurulur ve (istersen) etkinlestirilir

### Ileri Duzey Secenekler

- **Konfigurasyon Duzenleyici**: Dil, mimari (32/64-bit), kanal, uygulama secimleri, guncelleme tercihleri
- **Ozel XML**: config.office.com/deploymentsettings adresinden kendi XML dosyani olustur
- **Tema Secimi**: Gece / Aydinlik / Otomatik (sistem temasi)
- **Masaustu Kisayollari**: Kurulum sonrasi otomatik olusturma

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

38+ dil: Turkce, English, Deutsch, Francais, Espanol, Italiano, Portugues, Nederlands, Polski, Русский, Українська, العربية, हिन्दी, Bahasa, ไทย, Tieng Viet, 中文, 日本語, 한국어 ve daha fazlasi.

---

## Proje Yapisi

```
OfficeSetupWizard/
├── OfficeSetupWizard.exe          # Derlenmis portable uygulama
├── builder/
│   ├── build.ps1                  # .exe derleme betigi
│   ├── office_setup_wizard.py     # Ana uygulama (~2200 satir)
│   ├── office_setup_wizard.spec   # PyInstaller spec
│   ├── office_setup_wizard.ico    # Uygulama simgesi
│   ├── version_info.txt           # Windows surum bilgisi
│   └── icons/                     # Tema simgeleri
│       ├── light/                 # Aydinlik tema ikonlari
│       └── dark/                  # Karanlik tema ikonlari
└── resources/
    ├── setup.exe                  # Office Deployment Tool
    ├── Configuration.xml          # Varsayilan konfigurasyon
    └── configuration-Office365-x64.xml
```

---

## Sorumluluk Reddi

- Bu arac resmi Microsoft araclari (ODT) uzerine kuruludur.
- Aktivasyon ozelligi, topluluk tarafindan gelistirilen **Microsoft Activation Scripts (MAS)** projesini kullanir.
- Lissansiz yazilim kullanimi ulkenizin yasalarina ve Microsoft kullanim kosullarina tabidir. Bu proje yalnizca egitim ve kisisel kullanim amaciyla paylasilmaktadir; kullanim sorumlulugu tamamen kullanicidir.

---

## Katkida Bulun

Hata bildirimleri ve gelistirme onerileri icin [Issues](https://github.com/tarikelertarnak/OfficeSetupWizard/issues) sayfasini kullanabilirsin.

**Yapan:** [Tariq Eler](https://github.com/TARIKELER-TARNAK)

---

## Lisans

MIT License - Detaylar icin [LICENSE](LICENSE) dosyasina bakin.
