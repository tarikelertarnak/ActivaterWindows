# ActivaterWindows — Progress Log

## 2026-10-16 — v2.0.5: Tam MAS menusu [7] icine entegre edildi, menu sadelesti

### Degisiklikler
- **TAM MAS MENUSU ayri satir olarak KALDIRILDI** (eski [6]).
- **[7] HEPSINI ETKINLESTIR** (eski [8] HEPSINI YAP): tek tikli tam akis —
  Office kurulu degilse sorar/kurar (dil + bilesen secimi) → `TSforge /Z-WindowsESUOffice`
  (Windows + ESU + Office, kalici) → durum raporu → **sonda tam MAS menusu teklifi**
  (Troubleshoot, surum degistir, $OEM$ — hepsi bu akisin icinden). Boylece eski [6] ozelligi
  kaybolmadi, [7]'ye entegre oldu.
- Menu: [1-4] aktivasyon, [5] durum, [6] Office kur, [7] HEPSINI ETKINLESTIR, [0], [L].
- Yeni string'ler: `all_masmenu`, `all_done` (TR/EN). Eski `m8` anahtari silindi.

### Dogrulama
- Parse: `PARSE OK`.
- i18n: `KEYS OK: 52 anahtar x 2 dil`, `I18N TEST OK`.
- Menu render (TR+EN): 7 satir dogru, m8 kalmadi, `MENU RENDER OK`.

## 2026-10-16 — v2.0.4: Office bilesen secimi (Ozel Secim), renkli arayuz

### Degisiklikler
- **Office kurulum modu 3 secenek oldu:** [1] Onerilen (hafif), [2] Tam paket, **[3] Ozel Secim**.
- **Ozel Secim (`Select-OfficeApps`):** 9 bilesen (Word, Excel, PowerPoint, Outlook, Access, Publisher,
  OneNote, OneDrive, Skype for Business) numarali tuket-sec listesi. Komutlar: `1,3,5` toggle,
  `a`=tumu, `n`=hicbiri, `r`=onerilen, Enter=onayla. Yesil [X]=kurulacak, gri [ ]=atlanacak.
  En az bir bilesen zorunlu; secilen → ODT `ExcludeApp` listesine cevrilir.
- **Renklendirme:** mod menusu renkli, app picker renkli, secim sonucu yesil ozet satiri.
- **Tam TR/EN cevirisi** yeni arayuz icin (app_head, app_hint, app_sel, app_done, mode_c...).
- `sel12` → `Secim (1/2/3)`; recommended exclude listesi sabit `$OfficeAppsRecommended`'dan turetiliyor.

### Dogrulama
- Parse: `PARSE OK`.
- Picker testi (Read-Host mock'lu): toggle (5,7 ac + 1 kapat → Word cikti), `a`→9, `n`+`2`→tek Excel,
  exclude uretimi 8 dogru → `PICKER TEST OK`.
- i18n: `KEYS OK: 51 anahtar x 2 dil`, `I18N TEST OK`.
- Bulunan hata: `, $chosen` donusu array'i cift sarmaliyordu (Count=1 gorunuyordu) → duzeltildi.

## 2026-10-16 — v2.0.3: gizli kurulum penceresi, indirme barı, dil listesi, banner

### Degisiklikler
- **Baslik:** "ActivaterWindows — Terminal" → "ActivaterWindows"; site linki `https://tarikelertarnak.pages.dev/` banner'a eklendi.
- **Gizli Office kurulumu:** ODT ayiklama + `setup.exe /configure` + elevated (`powershell.exe -Verb RunAs`)
  hepsi `-WindowStyle Hidden` ile calisiyor. Kullanici artik o setup.exe penceresini gormez.
- **Indirme cubugu (gercek, chunked):** WebClient event handler'lari senkron cagrida ATESLENMIYOR
  (kanit: testte `FIRED COUNT: 0`) → `Invoke-Download` 64KB'lik chunk dongusuyle `HttpWebRequest`
  stream'ini okuyup `\r` ile gercek yuzde + saniye basiyor. Yuzde degismedikce tekrar basmiyor (jitter yok).
- **Dil listesi:** 83 Office dil kodu alfabetik numarali listede (5 kolon). Kod yaz ya da numara sec;
  numarada "N) kod — onayliyor musun? Enter=evet 0=geri" onayi. Bos Enter = otomatik.
  Yeni `Resolve-OfficeLang` + `Show-LangList`.
- **Banner:** basitlestirildi (36 karakter genislik), site + repo linki.

### Dogrulama
- Parse: `PARSE OK`.
- i18n: `I18N TEST OK` (TR/EN anahtar setleri esit).
- Bar/liste testi: `LIST COUNT: 83`, `SORT OK`, `tr-tr numarasi: 76` (index math dogru),
  `HASH OK`, `SIZE: 3.4 MB`, `DL+LIST TEST OK` — chunked bar canli render edildi.
- Event kaniti: WebClient sync event `FIRED COUNT: 0` (neden chunked'a gecildigi).

## 2026-10-16 — v2.0.2: arayuz dili (tr/en), HEPSINI YAP, KMS kaldirildi

### Degisiklikler
- **Arayuz dili secenegi** menu en ustunde: `[A] Otomatik (su an: tr-tr)  [T] Turkce  [E] English`
  + `[L]` tusu ile menu icinden degistirilebilir. Varsayilan: Otomatik (sistem kulturu; `tr*` → tr, digeri → en).
  Iki dil icin tam ceviri sozlugu (42 anahtar x 2 dil). Tum menu/alt metin/mesajlar T() ile cevrili.
- **[8] HEPSINI YAP** (tek tik "her sey"): Ofis kurulu degilse sorup kurar (TSforge zaten Office'i aktive eder),
  ardindan `/Z-WindowsESUOffice` ile Windows + ESU + Office kalici aktivasyon, en son durum raporu.
- **Online KMS (180 gun) menuden KALDIRILDI** — kalici TSforge + Ohook + HWID varken gereksiz.
  MAS'in kendi tam menusunden (6) erisilebilir.
- Menu yeniden numaralandi (KMS cikti): [1-4] aktivasyon, [5] durum, [6] TAM MAS, [7] Office kur, [8] hepsini, [0] cik.
- Office kurulum dili default'u arayuz diliyle eslesiyor (tr → tr-tr, en → en-us, auto → sistem).

### Dogrulama
- Parse: `PARSE OK`.
- i18n testi: TR/EN render, 42 anahtar karsilastirmasi (fark 0), auto→tr, Test-OfficeC2R ve
  Default-OfficeLang calisiyor → `I18N TEST OK`.

## 2026-10-16 — v2.0.1: ilk test turu düzeltmeleri (banner, dil, ODT indirme)

### Düzeltmeler (Kai'nin test raporu)
- **Banner'da sürüm numarası kaldırıldı:** "ActivaterWindows v2.0 — Terminal" → "ActivaterWindows — Terminal".
- **Dil seçimi otomatik + Enter:** Sistem kültür adı (`(Get-Culture).Name`, örn. `tr-tr`) default olarak
  prompt'ta gösteriliyor; boş Enter artık onu kullanıyor. "degistirmek istemiyorsan Enter'a bas" ipucu eklendi.
- **ODT indirme hatalı kaynaktı → düzeltildi:**
  - `aka.ms/odt` artık doğrudan `.exe`'ye değil **dokumantasyon sayfasına** gidiyordu (HTML → "dosya çok küçük" hatası).
  - `confirmation.aspx?id=49117` **404** döndü (Microsoft kaldırmış).
  - Yeni kaynak: `details.aspx?id=49117` sayfasındaki gerçek link regex'le çekiliyor
    (test edildi: `download.microsoft.com/download/6c1eeb25-.../officedeploymenttool_20326-20112.exe`).
  - **Yerel kopya önce:** `resources/officedeploymenttool*.exe` varsa internete dokunmadan oradan kullanılıyor
    (bu sayede offline çalışır).
- **Boyut eşiği `5MB` → `2MB`:** Gerçek ODT 3.4 MB çıktı; 5MB eşiği sahte "yanlış içerik" hatası üretirdi.

### Doğrulama
- Parse: `PARSE OK`.
- ODT uçtan uca testi: `LINK: download.microsoft.com/...officedeploymenttool_20326-20112.exe` →
  `SIZE: 3.4 MB` → `EXTRACT OK: setup.exe 7026 KB` → `LOCAL OK: officedeploymenttool_20131-20090.exe` → `ODT TEST OK`.

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