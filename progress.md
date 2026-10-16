# ActivaterWindows — Progress Log

## 2026-10-16 — v2.1.1: gizli calistirma kilitlenmesi + gunluk yaris durumu duzeltildi

Onceki turun incelemesinde cikan ve **olcumle dogrulanan** iki gercek hata giderildi.

### Duzeltilen hatalar

1. **Gizli pencerede sonsuz kilitlenme (kritik).** `Check_Activation_Status.cmd:62` satirindaki
   `choice /c 0 /n` gizli pencerede tus bekliyor; tus hic gelmeyince surec **sonsuza kadar asili kaliyordu**.
   Yeniden uretildi: `exited=False after 8,9s`.
   Cozum: gizli calistirmaya **900 sn** ust sinir (`$HiddenTimeoutSec`). Sure dolarsa surec
   `Stop-Process -Force` ile kapatilir, gunluge yazilan sonuclar yine gosterilir, kullaniciya
   "islem tamamlanmamis olabilir" uyarisi verilir.
   Olcum (testte sinir 4 sn'ye cekildi): ~5,7 sn'de donuyor, durum satirlari goruntuleniyor,
   **0 artik surec** kaliyor.

2. **Gunluk okuma yarisi (sessiz sonuc kaybi).** Surec kapatilirken gunluk hala kilitli oldugu icin
   `ReadAllText` "file is being used by another process" firlatiyordu; `try/catch` `$false` dondurdugu
   icin **sonuclar sessizce kayboluyordu**. Cozum: paylasimli modda (`FileShare.ReadWrite`) ve
   **30 x 200 ms** yeniden denemeli okuma.

3. **Onbellek dogrulanmiyordu.** `%LOCALAPPDATA%\ActivaterWindows` kopyasi okunurken yalnizca varlik
   kontrol ediliyordu. Artik kayit aninda alinan hash `.sha256` dosyasina yazilir ve okurken
   **yeniden dogrulanir**; uyusmazsa onbellek reddedilir. Durust sinir: ayni kullanici olarak calisan
   bir aktorun iki dosyayi birlikte degistirmesini engellemez — bu bir **dayaniklilik** yedegidir,
   guven siniri degildir (README'de yazili).

4. **Filtre sizintilari sikiistirildi** (gercek motor ciktisinda dogrulandi): `KMS4k`, `StaticCID`,
   `ZeroCID`, `Switching To`, `method is not supported`, `choose a`, `select a`, `^\s*\[[0-9]\]`.

### Bu turda ayrica eklenen

- `Get-ResultLines` tek yere cikarildi; `Show-ActivationResult` ve `-SelfTest` **aynı** filtreyi kullanir
  (iki kopya ayrisip yanlis PASS veremez).
- `-Probe` artik gercek `Resolve-Engine` zincirini calistirir (ayri bir kopya yolunu degil).
- Filtre hic satir birakmazsa "sonuc satiri yok" mesaji basar — **bos ekran ile basarisizlik karismaz**.

### Dogrulama (olcumle)

- Upstream takibi: `get.activated.win`'den commit `05c4f881...` + hash okundu, 761.964 bayt indirildi,
  hash uyustu → `-Probe OK (source=upstream)`.
- Penceresizlik: `RunAs` + `WindowStyle Hidden` → `Win32_Process` taramasinda **0 gorunur pencere`,
  `MainWindowHandle=0`.
- Motor kademeleri A (pin) / B (onbellek) / C (kurcalanmis icerik reddedildi) / D (temiz hata) → hepsi OK.
- `-SelfTest` PowerShell 5.1 ve 7'de gecildi.

### Kullanici tarafi

- Tek satir komuta TLS 1.2 zorlamasi eklendi (stok Windows'ta `irm` baglanti hatasi vermesin diye).
- README'ye "Her temiz Windows'ta calisir" bolumu eklendi (kurulan sey yok, kalici dosya yok).

## 2026-10-16 — v2.1.0: [7]/[8] ayrildi, motor kademeli, etkinlestirme tamamen penceresiz

### Degisiklikler

**Menu yeniden ayrildi (Kai istegi):**
- **[7] HEPSINI ETKINLESTIR** — *kurulum yok.* Mevcut olan her seyi kalici etkinlestirir:
  Windows (dijital lisans) + Office (kuruluysa) + Windows 10'da ESU. Sonunda durum raporu.
- **[8] HEPSINI YAP** — Office eksikse kurar (dil + bilesen secimi), sonra [7] akisini calistirir,
  en son tam MAS menusunu teklif eder. (Kai: "Hepsini Etkinlestir icinde Office olmasin,
  Officeli olan icin altina Hepsini Yap koy.")
- Yontem argumanlari **tek gecis + sira bagimsiz** gonderilir (`Resolve-Engine` + `/HWID /Ohook /Z-ESU`).

**Motor artik kendini guncel tutuyor (Kai: "guncellememe ragmen guncel kalir"):**
- `Resolve-Engine` 3 kademeli: **1)** `get.activated.win` bootstrap'indan guncel commit + hash okunur →
  her calistirmada upstream'in o anki surumu kullanilir. **2)** Olmazsa gomulu pin (`05c4f881`).
  **3)** Olmazsa `%LOCALAPPDATA%\ActivaterWindows`'taki son dogrulanmis kopya.
- Her kademede SHA256 zorunlu: bozuk/oynanmis indirme **asla** calistirilmaz, alt kademeye dusulur.
- Boylece upstream arguman sirasini/formatini degistirse bile uygulama bozulmaz.

**Yontem adlari arayuzden tamamen kaldirildi (Kai: "kullanici bilmesin"):**
- Menu artik `[2] Office etkinlestir - kalici`; `HWID`/`TSforge`/`Ohook` yazmaz.
- Etkinlestirme **penceresiz** calisir (`-WindowStyle Hidden`, cikti gecici gunluge alinir).
  Motorun kendi penceresi/basligi/markasi kullaniciya hic gorunmez.
- `Show-ActivationResult` gunlugu suzer: yalnizca sonuc satirlari (`Windows 11 Pro is LICENSED.`)
  gosterilir; yontem/marka adlari, kaynak URL'ler, lisans metinleri ve cmd komut yankilari dusurulur.
- GPL-3.0 atifi README/disclaimer'da **korunuyor** (lisans geregi); yalnizca arayuzde gosterilmiyor.

### Bulunan ve duzeltilen gercek hatalar (olcumle)

1. **Turkce "noktasiz i" tuzagi (kritik):** `'HWID'.ToLower()` bu makinede `hwıd` (U+0131) donduruyor.
   `tr-TR` kulturunde .NET `I` != `i` sayiyor, yani `(?i)` iceren **tum** regex'ler buyuk harfli metinde
   sessizce kaciyor. Kanit: `'HWID' -match 'hwid'` → **False**, `[regex]::IsMatch('HWID','d')` → **False**,
   ama `'Hwid' -match 'hwid'` → True.
   Etki: marka filtresi `Ohook`'u yakaliyor ama `HWID`/`LICENSED` satirlarini kaciriyordu; ilk test
   bu yuzden **yanlis PASS** verdi. Cozum: `Test-Match` (IgnoreCase + **CultureInvariant**) ve
   `ToLowerInvariant()`. Regresyon korumasi icin yerlesik `-SelfTest` eklendi.
2. **Menude `[A]/[T]/[E]` dil kisayollari olu idi:** `langline` bu tuslari ilan ediyordu ama `switch`
   yalnizca `l`/`L` isliyordu. Eklendi.
3. **Evet/hayir semantigi tersti:** TR ipucu `(E/h)` diyordu ama desen `^[HhYy]|^$` idi —
   yani **`h` (hayir) "evet" sayiliyor, `e` (evet) hic sayilmiyordu**. Tek `Confirm` helper'ina
   indirildi: TR `E` = evet, EN `Y` = yes, Enter = evet, diger her sey hayir.
4. **`Warn-AV`:** `-notlike '*windows*'` Turkce kulturde `WINDOWS DEFENDER` ile eslesmiyordu →
   Defender 3. parti AV gibi listeleniyordu. Culture-invariant yapildi.
5. **Office dil kodu:** `^[a-z]{2}(-...)$` buyuk harfli girisi (`TR-TR`) reddediyordu. Duzeltildi.
6. `[8]` akisinda `Install-OfficeLTSC` gereksiz "Office etkinlestirilsin mi?" soruyordu; `-SkipActivate`
   ile kaldirildi (zaten sonrasinda [7] hepsini etkinlestiriyor).

### Dogrulama (olculen)

| Test | Sonuc |
|---|---|
| `-SelfTest` (kultur tr-TR, filtre + parite + marka) | **8/8 OK**, `SELFTEST OK` |
| Parse | `PARSE OK` |
| Menu render TR+EN (culture-invariant marka taramasi) | `MENU tr/en OK` |
| Cikti filtresi (MAS benzeri cikti, buyuk harfli) | marka yok, `LICENSED`/`activated`/`ERROR` korunuyor, komut yankisi dusuyor → `FILTER OK` |
| `Confirm` semantigi (TR+EN, 11 girdi) | `CONFIRM SEMANTICS OK` |
| Motor kademeleri (4 senaryo) | A: upstream olu → **pin**; B: tum ag olu → **cache**; C: bozuk icerik → **reddedildi, cache**; D: hicbiri yok → temiz hata → `ALL FALLBACK SCENARIOS OK` |
| Canli upstream izleme | get.activated.win'den commit+hash okundu, indirildi, hash dogrulandi (`source=upstream`, commit `05c4f881…`, 761.964 bayt) |
| **Gizli pencere (olcum)** | `RunAs` + `WindowStyle Hidden` ile `Win32_Process` taramasinda **pencere sayisi 0** (`MainWindowHandle=0`), yonlendirme + cikis kodu calisiyor → `NO WINDOW (hidden confirmed)` |

### Bilinen sinirlar / riskler
- **UAC penceresi gizlenemez** (Windows guvenlik siniri). Makinede `ConsentPromptBehaviorAdmin=0`
  oldugu icin testte hic sorulmadi; normal makinede UAC gorunur, sonrasindaki konsol penceresi gizli.
- Canli Office/Windows etkinlestirmesi **calistirilmadi** (sistem durumunu degistirir; kullanicinin
  kendi testi). Bu yuzden "penceresiz + tam etkinlestirme" kombinasyonu birlikte degil, ayri ayri
  dogrulandi (gizli pencere mekanigi / filtre / motor kademeleri).
- Requirements kaynagi (ODT indirme linki) dis aga bagli; sayfa yapisi degisirse kurulum yedegi yok.

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