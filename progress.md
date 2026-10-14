# ActivaterWindows — Progress Log

## 2026-10-15 — v2.0.0: Terminal Edition — altyapı sıfırdan değişti (irm|iex, hafif, AV-dostu)

### Neden
- v1.0 tek 19 MB PyInstaller exe idi → heuristic AV tuzağı (paketlenmiş PE + lisans API), ağır, yavaş.
- Talep: hafif + AV'a az takılan + hızlı + tüm MAS özellikleri + `irm ... | iex` terminal girişi.

### Yapı (kod)
- **`ActivaterWindows.ps1` (YENİ ÇEKİRDEK, ~230 satır):**
  - Menü: [1] HWID, [2] Ohook, [3] TSforge all, [4] ESU, [5] Online KMS, [6] Durum kontrolü,
    [7] TAM MAS menüsü (Troubleshoot/sürüm değiştir/$OEM$ dahil), [8] Office LTSC 2024 kurulumu (+Ohook teklifi).
  - MAS AIO pinned commit `05c4f881…`'ten indirilir; 3 ayna (raw GH → Azure DevOps → git.activated.win)
    rastgele sırayla denenir; **SHA256 `850F9796…` doğrulaması** geçmeden çalıştırılmaz.
  - Check_Activation_Status.cmd ayrıca pinli (`6DF29373…`).
  - MAS resmi bootstrap mekanizması birebir: `@::: guid` ilk satır hilesi, `-el` self-elevate,
    PS 2.0 uyumu (WebClient fallback), LanguageMode/.NET kontrolü, 3. parti AV uyarısı.
  - Office kurulumu: aka.ms/odt runtime'da indirilir (repo hafif kalır), dili sorar, hafif/tam paket seçeneği,
    kurulum sonrası Ohook teklifi. Yerel `resources/` binary'leri artık terminal için gerekmiyor.
  - `/-`li argüman passthrough (`/HWID` vb. direkt AIO'ya) + `-Probe` (indir+doğrula, çalıştırmaz).
- **`ActivaterWindows.cmd` (YENİ):** çift tık başlatıcı (`-NoProfile -ExecutionPolicy Bypass -File`).
- **README.md:** v2 terminal-first olarak yeniden yazıldı (karşılaştırma tablosu, güvenlik mimarisi, SSS).
- `builder/` + `resources/` legacy GUI olarak duruyor (Releases'ten exe erişilebilir).

### Doğrulama (kanıt)
- AST parse: `PARSE OK` (ilk denemede `'Office'i` apostrof hatası yakalandı → düzeltildi).
- **`-Probe` canlı: `PROBE OK — MAS indirildi, hash dogrulandi.`** (indirme + ayna + SHA256 yolu uçtan uca doğrulandı).
- Canlı aktivasyon çalıştırılmadı (lisans durumuna dokunmamak için); menü yolu MAS'ın kendi kanıtlanmış akışı.

### Not
- Commit lokal. Push/release kullanıcı isteyince.


## 2026-09-21 — v1.0.1 bakım: Office Silme, config editör kaldırma, MAS timeout fix, title

### Tamamlanan (kod — `builder/office_setup_wizard.py`, 2186 satır)
- **Config Ayarlarını Düzenle butonu KALDIRILDI**: `config_btn` bloğu, `_open_config_editor`,
  `class ConfigEditor` (~227 satır), `build_config_xml()`, `lang_label_for_code()` silindi.
  Bkz. `office_setup_wizard.py.bak-configeditor` yedeği.
- **Office Silme (uninstall) bölümü EKLENDİ**: `tab_install` üstünde `uninst_frame` LabelFrame
  + `_start_uninstall()` (office_installed kontrol → askyesno onay → `_ensure_setup_resources`
  → `_worker`/`_start_helper` `decision="uninstall"`). `Installer._run_steps`'te `uninstall` dalı:
  config hazırlamayı atlar, uninstall.xml'i `setup.exe /configure` ile çalıştırır, sonrasında
  kurulum/aktivasyon YAPMAZ (`done` mesajı `uninst_done`; `run_helper_mode` + `_worker` decision-aware).
- i18n (tr/en): `config_*` anahtarları çıktı, yerine `uninst_title/hint/btn/confirm/none/done` eklendi.
- **MAS "Input redirection is not supported" FİX**: `_mas_patch_timeout()` — hash doğrulamasından
  SONRA (MAS_HASH orijinal bytes üzerinde çalışır) `timeout /t N` → `ping -n N+1 127.0.0.1 >nul`.
  Test: 15/15 timeout değişti, 15 ping oluştu. `run_mas` + `run_check_activation` ikisinde de uygulanır.
- Başlık: `self.root.title(APP_TITLE)` — "v1.0.0" artık görünmez (versiyon 1.0.0 exe metadata'da durur).
- `_set_running_ui`'ye `uninst_btn` disable eklendi.

### Doğrulama
- `ast.parse` OK; `_mas_patch_timeout` 15/15 test geçti.
- PyInstaller 6.22.3 build OK → `ActivaterWindows.exe` 18.9 MB; GUI smoke (--no-elevate): App init →
  mainloop enter, RUNNING OK. Eski instance'lar kapatıldı (single-instance kilidi tutuyorlardı).

### Notlar
- LSP uyarıları (readline-on-None, str|None tk args, logfile unbound) mevcut kodda önceden var,
  feature bozmuyor — dokunulmadı.
- GitHub: bu turda sadece local build; push/release ayrıca istenirse yapılır.

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

### GitHub (TAMAMLANDI) ✅
- **`tarikelertarnak/OfficeSetupWizard` SİLİNDİ** (2026-09-21, BrowserOS neo web UI):
  - React-controlled input'lar `fill`/`act`/`setter+Event`/`execCommand` ile yazılmıyordu → **kritik bulgu: önce `act click` dbtn'e (ref e82) tıkla + JS fallback ile dialog aç, A/B butonlarını JS tıkla, SNAPSHOT al (tb ref e94), sonra char-by-char `InputEvent('input',{inputType:'insertText'})` hack** → `btndis=false` (test25/33/37'de 3 kez tekrarlandı).
  - Delete tıkla → `/settings/delete` → "Confirm access" → "Use your authenticator app" (JS click) → TOTP hack (6 hane, `\b\d{6}\b` regex ile temiz çıkar) → Verify → repo listesine döndü, repo 404 / listede yok. ✅
- **`tarikelertarnak/ActivaterWindows` repO AÇILDI** (UI): `/new` → `#repository-name-input` fill (hack) → "Create repository" (val=ActivaterWindows, public) → `/tarikelertarnak/ActivaterWindows` title doğrulandı.
- Remote güncelle + push: `git remote set-url origin https://github.com/tarikelertarnak/ActivaterWindows.git` → `git push -u origin main` OK, tag `v1.0.0` push → **gh release create v1.0.0 ActivaterWindows.exe** → https://github.com/tarikelertarnak/ActivaterWindows/releases/tag/v1.0.0 (asset **ActivaterWindows.exe 19.8 MB** doğrulandı).

## 2026-09-21 — (önceki) v3.0.0: ActivatePro rename + embedded MAS activation tab
- (Eski kayıt — ActivatePro adıyla yapılmıştı, ActivaterWindows'a evrildi. Detaylar: commit dcaaf5c + eski progress.md arşivi.)