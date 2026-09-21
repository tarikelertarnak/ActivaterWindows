# ActivaterWindows — Progress Log

## 2026-09-21 — v1.0.0: ActivaterWindows isim + 16 aktivasyon yöntemi + koyu sekme

### Tamamlanan (kod)
- `APP_TITLE`/`APP_VERSION` → **ActivaterWindows / 1.0.0**; User-Agent güncellendi.
- Tüm kalıntılar temizlendi (grep 0): mutex `Local\ActivaterWindows_SingleInstance`, temp `%TEMP%\ActivaterWindows`, cleanup desenleri, about_info, res_missing_where.
- `ACT_METHODS` 16 yönteme genişletildi:
  - Windows 11 Pro, 11 Pro for Workstations, 11 Home, 11 → `/HWID`
  - Windows 10 Pro, 10 Home, 10 → `/HWID`
  - Ohook `/Ohook`, TSforge `/Z-Windows`, ESU `/Z-ESU`
  - KMS all/windows/office, hepsi `/Z-WindowsESUOffice`, Ohook kaldır, KMS kaldır
- tr/en: `act_win11pro/work/home/win11/win10pro/home/win10` etiket + desc girişleri eklendi; `act_run` → "Aktifleştir"/"Activate", `act_frame` → "Windows / Office Aktivasyonu".
- Koyu tema sekme renkleri: `TNotebook` bg `#202124`, `TNotebook.Tab` bg `#3c4043`, seçili `#8ab4f8` (`_apply_theme`).
- Spec `name='ActivaterWindows'`; build.ps1 → `ActivaterWindows.exe` v1.0.0; version_info.txt (1,0,0,0).
- .gitignore: `/ActivaterWindows.exe`, eski exe'ler çıkarıldı/düzenlendi.

### Doğrulama (smoke)
- Import: APP_TITLE=ActivaterWindows, APP_VERSION=1.0.0, 16 yöntem — tamam.
- GUI smoke: başlık "ActivaterWindows v1.0.0", sekmeler ['Office Kurulumu','Aktivasyon'], combobox 16 değer hepsi çevrili (ham anahtar yok), desc'ler tüm ürünlerde yüklü — tamam.
- Tab stilleri: bg #202124 / #3c4043 / seçili #8ab4f8 — doğrulandı.

### Build
- `python -m PyInstaller --noconfirm --clean office_setup_wizard.spec` → `ActivaterWindows.exe` 18.9 MB (LastWrite 20:24).
- MAS_AIO.cmd gömülü: True. Runtime: exe çalıştı, Responding True, sonra kapatıldı.
- Eski `ActivatePro.exe` kökten silindi (kilitli olduğu için process kapatılarak).

### GitHub (ENGELLİ)
- Hedef: `tarikelertarnak/OfficeSetupWizard` sil → `ActivaterWindows` repo oluştur → push.
- gh token (fine-grained, Contents RO + Issues/PR RW) admin işlemlerde 403: delete, rename (PATCH), createRepository (GraphQL), POST user/repos (REST) — hepsi denendi, hepsi "Resource not accessible by personal access token".
- BrowserOS neo (port 9014) kapalı; `neo_*` MCP tool'ları bu oturuma yüklü değil; Chrome kurulu değil (sadece Edge), bu yüzden chrome-devtools da bağlanamıyor.
- Çözüm yolu: yeni session'da BrowserOS neo açıkken repo işlemleri (web UI + push için token scopes'u Contents RW lazım) — kullanıcıya bildirilecek.

## 2026-09-21 — (önceki) v3.0.0: ActivatePro rename + embedded MAS activation tab
- (Eski kayıt — ActivatePro adıyla yapılmıştı, ActivaterWindows'a evrildi. Detaylar: commit dcaaf5c + eski progress.md arşivi.)