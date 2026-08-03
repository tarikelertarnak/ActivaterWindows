# ⚙️ Office Setup Wizard

> **Microsoft Office LTSC Professional Plus 2024'ü tek tıkla kurun ve etkinleştirin.**

`Office Setup Wizard`, Office Deployment Tool'u (ODT) sararak kurulumu otomatikleştiren, kullanıcı dostu bir **Windows masaüstü uygulamasıdır**. Karmaşık XML yapılandırmalarıyla uğraşmadan, grafik arayüzden birkaç tıklamayla Office'i kurar.

<p align="center">
  <img src="https://img.shields.io/badge/platform-Windows-blue?style=flat-square" />
  <img src="https://img.shields.io/badge/language-Python-3776AB?style=flat-square&logo=python&logoColor=white" />
  <img src="https://img.shields.io/badge/UI-Tkinter-orange?style=flat-square" />
  <img src="https://img.shields.io/badge/license-MIT-green?style=flat-square" />
</p>

---

## ✨ Özellikler

| Özellik | Açıklama |
|---|---|
| ⚡ **Otomatik Kurulum** | Office LTSC Professional Plus 2024 - istediğin dilde tek tıkla kurulum |
| 📄 **Özel Yapılandırma** | Kendi `Configuration.xml` dosyanı kullan (ileri düzey kullanıcılar) |
| 🌍 **Dil Seçimi** | `tr-tr`, `en-us` vb. herhangi bir dil koduyla kurulum |
| 🎨 **Aydınlık/Karanlık Tema** | Modern, tercihine göre değişen arayüz |
| 🧾 **Etkinleştirme** | Kurulum sonrası isteğe bağlı aktivasyon (MAS) |
| 📋 **Ayrıntılı Günlük** | Kurulum sürecini adım adım izleme ve hata kaydı |
| 📥 **Otomatik İndirme** | Gerekli ODT dosyalarını ve yapılandırmayı kendi indirir |

---

## 🚀 Kullanım

### Ön Koşullar
- **Windows 10 / 11** (x64)
- **Python 3.8+** (kaynaktan çalıştırmak için)
- İnternet bağlantısı (Office dosyalarını indirmek için)

### Çalıştırma (Kaynaktan)

```powershell
git clone https://github.com/TARIKTR1099/OfficeSetupWizard.git
cd OfficeSetupWizard
python builder/office_setup_wizard.py
```

### Derleme (Standalone .exe)

```powershell
powershell -ExecutionPolicy Bypass -File builder/build.ps1
```

Derlenen yürütülebilir dosya `dist/` klasöründe oluşur.

---

## 🖱️ Kullanım Adımları

1. **Kurulum Modu** seç: *Otomatik* (LTSC 2024) veya *Özel Configuration*
2. Kurulum dilini seç (ör. `tr-tr`, `en-us`)
3. **Kurulumu Başlat** butonuna bas
4. Office indirilir, kurulur ve (istersen) etkinleştirilir ✅

---

## ⚠️ Sorumluluk Reddi

- Bu araç resmi Microsoft araçları (ODT) üzerine kuruludur.
- Aktivasyon özelliği, topluluk tarafından geliştirilen **Microsoft Activation Scripts (MAS)** projesini kullanır.
- Lisanssız yazılım kullanımı ülkenizin yasalarına ve Microsoft kullanım koşullarına tabidir. Bu proje yalnızca eğitim ve kişisel kullanım amacıyla paylaşılmaktadır; kullanım sorumluluğu tamamen kullanıcıya aittir.

---

## 📦 Proje Yapısı

```
OfficeSetupWizard/
├── builder/
│   ├── build.ps1            # .exe derleme betiği
│   ├── office_setup_wizard.py  # Ana uygulama
│   ├── office_setup_wizard.spec # PyInstaller spec
│   └── icons/               # Uygulama simgeleri
└── resources/
    ├── Configuration.xml    # Örnek yapılandırma
    └── ...
```

---

## 🤝 Katkıda Bulun

Hata bildirimleri ve geliştirme önerileri için [Issues](https://github.com/TARIKTR1099/OfficeSetupWizard/issues) sayfasını kullanabilirsin.

**Yapan:** [Tarık Eler](https://github.com/TARIKTR1099) · ⭐ Beğendiysen yıldızlamayı unutma!
