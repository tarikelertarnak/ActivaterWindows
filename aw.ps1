# ActivaterWindows
# PowerShell for Windows 10/11. The activation engine is followed upstream (live commit + hash);
# if that fails, a pinned commit; if that fails, the last verified local copy. Always SHA256-verified.
# Method names are never shown in the UI; activation runs with no visible window.
# One-liner: irm https://tarikelertarnak.github.io/aw.ps1 | iex
# Local: ActivaterWindows.cmd (double-click) or powershell -ExecutionPolicy Bypass -File ActivaterWindows.ps1
# Switches: -SelfTest (offline checks, then exit) | -Probe (download + verify engine, then exit)
#           -Refresh (ignore the engine cache) | any "/" argument is forwarded to the activation engine

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

function Set-ConsoleTitle([string]$Title = 'ActivaterWindows') {
    try { [Console]::Title = $Title } catch {}
    try { $host.UI.RawUI.WindowTitle = $Title } catch {}
    try { [Console]::Write("`e]0;$Title`a") } catch {}
}
Set-ConsoleTitle 'ActivaterWindows'
$Version  = '3.0.2'
$Author   = 'TARIK ELER'
$RepoUrl  = 'https://github.com/tarikelertarnak/ActivaterWindows'
$SiteUrl  = 'https://tarikelertarnak.pages.dev/'
$ShortUrl = 'https://tarikelertarnak.github.io/aw.ps1'
$Pinned   = '05c4f881efec946c0040cdd552d1afa9a519704b' # engine commit
$AioHash  = '850F979665FB93999ACAE93F4790C1FF8ED2041532060B7966A121C2D29A0BFA'
$ChkHash  = '6DF2937385386D0C9682D351797E674583A410F4792BB7BE45302F40AB90287E'

# Office Deployment Tool (Microsoft's own downloader). It is mirrored to our releases so a broken
# Microsoft download page or a blocked download.microsoft.com is not fatal. Verified by hash when
# the mirror is refreshed (tools/Mirror-ODT.ps1), and always by size + PE header at use time.
$OdtMirrorUrl  = "$RepoUrl/releases/download/odt-mirror/officedeploymenttool.exe"
$OdtPinnedUrl  = 'https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_20326-20112.exe'
$OdtPinnedHash = 'FBB64358FD4168ACD52EE4EFE47FFD032B6231DFB415AE2DCE61B0E58BA67F86'
$OdtMinBytes   = 2MB

$CacheDir = Join-Path $env:LOCALAPPDATA 'ActivaterWindows'
$AioCache = Join-Path $CacheDir 'engine.cmd'
$AioCacheHash = Join-Path $CacheDir 'engine.sha256'
$ChkCache = Join-Path $CacheDir 'check.cmd'
$ChkCacheHash = Join-Path $CacheDir 'check.sha256'
# Upper bound (seconds) for a hidden engine run. Flows that end waiting for a keypress are killed after this.
$HiddenTimeoutSec = 900
# Reuse the cached engine without any network round-trip while it is younger than this (startup speed).
# -Refresh (or -Probe) forces a fresh upstream check.
$CacheMaxAgeHours = 24
$script:UiLang = 'auto' # 'auto' | 'tr' | 'en'
# Registry-based OS info is instant; WMI/CIM is not (measured 1.6-4.1 s), so it is cached per run.
$script:WinInfo = $null

# Live bar for percentage-less steps: a finished one passes at once, an unfinished one is cut off
$OfficeLangs = 'af-za am-et ar-sa as-in az-latn-az be-by bg-bg bn-bd bn-in bs-latn-ba ca-es cs-cz cy-gb da-dk de-de el-gr en-gb en-us es-es es-mx et-ee eu-es fa-ir fi-fi fil-ph fr-ca fr-fr gl-es gu-in he-il hi-in hr-hr hu-hu hy-am id-id is-is it-it ja-jp ka-ge kk-kz km-kh kn-in ko-kr ky-kg lo-la lt-lt lv-lv mk-mk ml-in mn-mn mr-in ms-my mt-mt nb-no ne-np nl-nl nn-no or-in pa-in pl-pl ps-af pt-br pt-pt ro-ro ru-ru si-lk sk-sk sl-si sq-al sr-latn-rs sv-se sw-ke ta-in te-in th-th tr-tr uk-ua ur-pk uz-latn-uz vi-vn zh-cn zh-tw zu-za' -split ' '

# Auto UAC elevation: admin degilse kendini yukselt (SelfTest/Probe/Refresh haric)
if (-not $args.Contains('-SelfTest') -and -not $args.Contains('-Probe') -and -not $args.Contains('-Refresh')) {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host 'Yonetici gerekiyor, yukseltiliyor...' -ForegroundColor Yellow
        if ($PSCommandPath) {
            # Yerel dosya olarak calistirilmis
            Start-Process pwsh -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $($args -join ' ')" -Verb RunAs
        } else {
            # irm | iex icin: command parametresi dogru escape edilmeli
            $cmd = 'irm https://tarikelertarnak.github.io/aw.ps1 | iex'
            Start-Process pwsh -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -Command " + [char]34 + $cmd + [char]34) -Verb RunAs
        }
        exit
    }
}

# A "/"-argument is forwarded straight to the AIO engine (e.g. ActivaterWindows.ps1 /HWID)
$OfficeApps = @(
    @{ n = 'Word';               id = 'Word' },
    @{ n = 'Excel';              id = 'Excel' },
    @{ n = 'PowerPoint';         id = 'PowerPoint' },
    @{ n = 'Outlook';            id = 'Outlook' },
    @{ n = 'Access';             id = 'Access' },
    @{ n = 'Publisher';          id = 'Publisher' },
    @{ n = 'OneNote';            id = 'OneNote' },
    @{ n = 'OneDrive';           id = 'OneDrive' },
    @{ n = 'Skype for Business'; id = 'Lync' }
)
$OfficeAppsRecommended = @('Word', 'Excel', 'PowerPoint')


$Strings = @{
    tr = @{
        langline = 'Dil / Language :  [A] Otomatik (su an: {0})   [T] Turkce   [E] English'
        m1 = '  [1]  Windows etkinlestir                kalici dijital lisans'
        m2 = '  [2]  Windows surum degistir + etkinles. surum secimi'
        m3 = '  [3]  Windows 10 ESU                    kalici'
        m4 = '  [4]  Office etkinlestir                kalici'
        m5 = '  [5]  Office LTSC 2024 kur              dil + bilesen secimi'
        m6 = '  [6]  Office LTSC 2024 kur + etkinles  onceki surumu temizler, kurar, etkinlestirir'
        m7 = '  [7]  Windows + ESU + Office            tek gecis'
        m8 = '  [8]  HEPSINI ETKINLESTIR               kurulum yok, hepsini etkinlestir'
        m9 = '  [9]  HEPSINI YAP                       kur + etkinlestir + tam menü'
        mA = '  [A]  Etkinlestirme durumu              kontrol et'
        mH = '  [H]  Hakkinda'
        mx = '  [w]  Cikis'
        ml = '  [L]  Arayuz dilini degistir'
        choice = 'Secim'
        back = 'Menuye donmek icin Enter'
        sel12 = 'Secim (1/2/3)'
        mode_head = 'Kurulum modu:'
        langpick = 'Dil:  [1] Otomatik (su an: {0})  [2] Turkce  [3] English'
        olang = 'Office dili: kod yaz (ornek: tr-tr) veya asagidaki listeden numara sec [otomatik: {0}]'
        olang_confirm = '{0}) {1} - onayliyor musun? Enter = evet, 0 = geri'
        lang_head = 'Office dilleri (alfabetik):'
        mode_h = '  [1] Onerilen (hafif) : sadece Word, Excel, PowerPoint kurulur'
        mode_t = '  [2] Tam paket         : tum bilesenler kurulur'
        mode_c = '  [3] Ozel Secim        : hangi bilesenin kurulacagini kendin sec'
        app_head = 'OZEL SECIM / CUSTOM SELECTION'
        app_hint = 'Ac/Kapa: numaralari virgul ile yaz (ornek: 1,3,5)   |   a = tumu   n = hicbiri   r = onerilen (Word, Excel, PowerPoint)   Enter = onayla'
        app_count = 'Secili: {0} / {1}'
        app_rec = 'oneri'
        app_sel = 'Secim'
        app_install = 'kurulacak'
        app_skip = 'atlanacak'
        app_none = 'Hicbir bilesen secilmedi, en az birini sec!'
        app_done = 'Ozel secim: {0} bilesen kurulacak, {1} bilesen atlanacak.'
        app_chosen = 'Kurulacak: {0}'
        invalid_lang = 'Gecersiz dil kodu veya numara.'
        odt_local = 'ODT yerel kopyadan aliniyor: {0}'
        odt_dl = 'Office Deployment Tool indiriliyor (microsoft.com)...'
        odt_fail = 'ODT indirilemedi: {0}'
        odt_small = 'ODT indirilen dosya yanlis icerik, iptal.'
        extract = 'ODT ayiklaniyor...'
        extract_fail = 'setup.exe ayiklanamadi, iptal.'
        inst_start = 'Office kuruluyor, bir kac dakika surebilir...'
        inst_uac = 'Office kuruluyor (UAC izni gerekecek)...'
        inst_done = 'Kurulum bitti (cikis kodu: {0}).'
        inst_fail = 'Office kurulumu TAMAMLANMADI (cikis kodu: {0}). Ayrintili log: %TEMP%\OfficeSetup*.log'
        step_timeout = 'Bu adim beklenenden uzun surdu ve durduruldu. Islem tamamlanmamis olabilir.'
        uac_cancel = 'Kurulum iptal edildi (UAC reddedildi).'
        done = '{0} tamamlandi (cikis kodu: {1}).'
        uac_cancel2 = '{0} iptal edildi (UAC reddedildi).'
        mas_dl = 'MAS All-In-One indiriliyor'
        chk_dl = 'Aktivasyon durumu kontrolu indiriliyor'
        hash_fail = 'MAS butunluk kontrolu basarisiz (hash uyusmadi), iptal!'
        hash_fail2 = 'Kontrol scripti butunluk kontrolu basarisiz, iptal!'
        dl_fail = '{0} hicbir kaynaktan alinamadi. Antivirus/firewall baglantayi engelleyebilir.'
        tmp_fail = 'Gecici dosya olusturulamadi, iptal! ({0})'
        all_start = 'HEPSINI YAP: ofis kurulu degilse kurulacak, sonra mevcut her sey kalici etkinlestirilecek, en son durum raporu...'
        all_nooffice = 'Office kurulu gorunmuyor. LTSC 2024 kurulsun mu? (E/h, bos = evet)'
        all_installing = 'Once Office kuruluyor...'
        all_masmenu = 'Tam MAS menusu de acilsin mi? (Troubleshoot, surum degistir, $OEM$...) (E/h)'
        all_done = 'Hepsini etkinlestirme turu bitti.'
        office_act_ask = 'Office simdi etkinlestirilsin mi? (E/h)'
        act_start = 'Etkinlestirme basliyor - arka planda calisir, pencere acilmaz.'
        act_nooffice = 'Office kurulu degil, Office etkinlestirme atlandi.'
        using_cache = 'DIKKAT: dogrulanmis uzak kaynak yok - yerel onbellekteki motor kullaniliyor.'
        engine_fail = 'Aktivasyon motoru hicbir kaynaktan alinamadi (ag/antivirus).'
        status_lbl = 'Durum kontrolu'
        act_timeout = 'Motor {0} saniyede bitmedi ve kapatildi. Yukarida sonuc yoksa islem tamamlanmamis olabilir.'
        act_nolog = 'Motor calisti ama bildirilecek sonuc satiri yok. Ayrinti icin [5] durum kontrolunu kullan.'

        author = 'TARIK ELER'
        sec_win = 'WINDOWS'
        sec_off = 'OFFICE'
        sec_all = 'BIRLESIK'
        sec_info = 'DURUM & BILGI'
        esc_hint = 'Esc = geri'
        about_head = 'HAKKINDA'
        about_ver = 'Surum   : {0}'
        about_by = 'Yapan   : {0}'
        about_lic = 'Lisans  : MIT (uygulama) + GPL-3.0 (motor)'
        about_site = 'Site    : {0}'
        about_repo = 'Kaynak  : {0}'
        about_short = 'Tek satir: {0}'
        about_note = 'Aktivasyon motoru Microsoft Activation Scripts projesinden alinir (GPL-3.0). Lisans geregi atif korunur; arayuzde yontem adlari hicbir zaman gosterilmez.'
        win_head = 'Windows etkinlestirme'
        win_cur = 'Mevcut surum: {0}   build {1}   {2}'
        win_a1 = '  [1]  Bu surumu etkinlestir             kalici dijital lisans'
        win_a2 = '  [2]  Surum degistir, sonra etkinlestir mevcut surumler listelenir'
        win_a3 = '  [3]  Windows 10 ESU etkinlestir        kalici'
        win_eds = 'Bu yapiya kurulabilen Windows surumleri ({0}):'
        win_eds_note = 'Liste WMI lisans ailelerinden okunur. Degisimi motorun kendi penceresi yapar; orada numarayla secersin.'
        win_change_go = 'Surum degistirme araci aciliyor. Islem bittikten sonra etkinlestirme gerekir.'
        win_change_ask = 'Simdi etkinlestirme calistirilsin mi? (E/h)'
        odt_mirror = 'Microsoft indirmesi calismadi, GitHub yedeginden aliniyor...'
        odt_pinned = 'Microsoft indirme sayfasi degismis, sabitlenmis Microsoft baglantisina geciliyor...'
        odt_verify = 'ODT dogrulandi: {0} ({1:N1} MB)'
        off_dl_first = 'Once kurulum araci indirilir, sonra dil ve bilesen sorulur.'
        off_ready = 'Kurulum araci hazir. Simdi secimlerini yap:'
        off_summary = 'Dil: {0}   Kurulacak: {1}'
        off_go = 'Kurulum baslatilsin mi? (E/h)'
        office_installed = 'Microsoft Office LTSC Professional Plus 2024 zaten kurulu. Temizlensin mi? (E/h, bos = evet)'
        office_installed_other = 'Baska bir Office surumu kurulu ({0}). Temizlensin mi? (E/h, bos = evet)'
        office_removing = 'Office kaldiriliyor...'
        office_removed = 'Office kaldirildi.'
        office_remove_fail = 'Office kaldirilanamadi (cikis kodu: {0}).'
        office_install_running = 'Office kurulum islemi zaten calisiyor. Lutfen bekleyin ve tekrar deneyin (Hata 0-2035).'
        m0back = '  [0]  Geri'
        odt_hash_fail = 'ODT butunluk kontrolu basarisiz, bu kaynak atlaniyor ({0}).'
        win_eds_none = 'Surum listesi okunamadi (WMI). Motorun kendi penceresi yine de listeler.'
        win_change_ask_go = 'Surum degistirme araci acilsin mi? (E/h)'
        ce_dl = 'Surum degistirme araci indiriliyor'
        ce_fail = 'Surum degistirme araci alinamadi, iptal.'
        win_change_title = 'Surum degistir - surum secin'
        win_change_cur   = 'Su anki surum: {0}'
        win_change_rec   = '[x] (oneri - su anki surum)'
        win_change_sel   = 'Numarayla secin, sonra Enter:'
        win_change_next  = '10) Digere bak (kalan surumleri goster)'
        win_change_back  = '0) Geri'
    }
    en = @{
        langline = 'Dil / Language :  [A] Automatic (now: {0})   [T] Turkce   [E] English'
        m1 = '  [1]  Activate Windows                  permanent digital license'
        m2 = '  [2]  Change edition + activate        pick edition & activate'
        m3 = '  [3]  Windows 10 ESU                  permanent'
        m4 = '  [4]  Activate Office                 permanent'
        m5 = '  [5]  Install Office LTSC 2024        language + component picker'
        m6 = '  [6]  Office LTSC 2024 install + act   clears existing, installs, activates'
        m7 = '  [7]  Windows + ESU + Office          single pass'
        m8 = '  [8]  ACTIVATE EVERYTHING             no install, activate all'
        m9 = '  [9]  DO EVERYTHING                   install + activate + full menu'
        mA = '  [A]  Activation status               check'
        mH = '  [H]  About'
        mx = '  [w]  Exit'
        ml = '  [L]  Change interface language'
        choice = 'Choice'
        back = 'Press Enter to return to menu'
        sel12 = 'Choice (1/2/3)'
        mode_head = 'Install mode:'
        langpick = 'Language:  [1] Automatic (now: {0})  [2] Turkce  [3] English'
        olang = 'Office language: type a code (e.g. tr-tr) or pick a number from the list below [auto: {0}]'
        olang_confirm = '{0}) {1} - confirm? Enter = yes, 0 = back'
        lang_head = 'Office languages (alphabetical):'
        mode_h = '  [1] Recommended (light) : only Word, Excel, PowerPoint install'
        mode_t = '  [2] Full                : every component installs'
        mode_c = '  [3] Custom              : choose exactly which components install'
        app_head = 'CUSTOM SELECTION'
        app_hint = 'Toggle: type numbers separated by commas (e.g. 1,3,5)   |   a = all   n = none   r = recommended (Word, Excel, PowerPoint)   Enter = confirm'
        app_count = 'Selected: {0} / {1}'
        app_rec = 'rec.'
        app_sel = 'Choice'
        app_install = 'install'
        app_skip = 'skip'
        app_none = 'No component selected, pick at least one!'
        app_done = 'Custom: {0} components will install, {1} will be skipped.'
        app_chosen = 'Installing: {0}'
        invalid_lang = 'Invalid language code or number.'
        odt_local = 'Using local ODT copy: {0}'
        odt_dl = 'Downloading Office Deployment Tool (microsoft.com)...'
        odt_fail = 'ODT download failed: {0}'
        odt_small = 'Downloaded ODT file is not valid, aborting.'
        extract = 'Extracting ODT...'
        extract_fail = 'setup.exe could not be extracted, aborting.'
        inst_start = 'Installing Office, this may take a few minutes...'
        inst_uac = 'Installing Office (UAC prompt will appear)...'
        inst_done = 'Install finished (exit code: {0}).'
        inst_fail = 'Office install did NOT complete (exit code: {0}). Details: %TEMP%\OfficeSetup*.log'
        step_timeout = 'This step took longer than expected and was stopped. The operation may be incomplete.'
        uac_cancel = 'Install cancelled (UAC declined).'
        done = '{0} finished (exit code: {1}).'
        uac_cancel2 = '{0} cancelled (UAC declined).'
        mas_dl = 'Downloading MAS All-In-One'
        chk_dl = 'Downloading activation status checker'
        hash_fail = 'MAS integrity check FAILED (hash mismatch), aborting!'
        hash_fail2 = 'Checker integrity check FAILED, aborting!'
        dl_fail = '{0} could not be retrieved from any source. Antivirus/firewall may be blocking.'
        tmp_fail = 'Failed to create temp file, aborting! ({0})'
        all_start = 'DO EVERYTHING: install Office if missing, then permanently activate everything present, then status report...'
        all_nooffice = 'Office does not seem installed. Install LTSC 2024? (Y/n, blank = yes)'
        all_installing = 'Installing Office first...'
        all_masmenu = 'Also open the full MAS menu? (Troubleshoot, change edition, $OEM$...) (Y/n)'
        all_done = 'Activate-everything run finished.'
        office_act_ask = 'Activate Office now? (Y/n)'
        act_start = 'Activation starting - runs in the background, no window opens.'
        act_nooffice = 'Office is not installed, skipping Office activation.'
        using_cache = 'WARNING: no verified remote source - using the locally cached engine.'
        engine_fail = 'Activation engine could not be retrieved from any source (network/antivirus).'
        status_lbl = 'Status check'
        act_timeout = 'The engine did not finish within {0}s and was terminated. If no result is shown above, the operation may be incomplete.'
        act_nolog = 'The engine ran but produced no reportable result line. Use [5] status check for details.'

        author = 'TARIK ELER'
        sec_win = 'WINDOWS'
        sec_off = 'OFFICE'
        sec_all = 'COMBINED'
        sec_info = 'STATUS & INFO'
        esc_hint = 'Esc = back'
        about_head = 'ABOUT'
        about_ver = 'Version : {0}'
        about_by = 'Made by : {0}'
        about_lic = 'License : MIT (app) + GPL-3.0 (engine)'
        about_site = 'Site    : {0}'
        about_repo = 'Source  : {0}'
        about_short = 'One-liner: {0}'
        about_note = 'The activation engine comes from the Microsoft Activation Scripts project (GPL-3.0). Attribution is kept for license compliance; method names are never shown in the UI.'
        win_head = 'Windows activation'
        win_cur = 'Current edition: {0}   build {1}   {2}'
        win_a1 = '  [1]  Activate this edition             permanent digital license'
        win_a2 = '  [2]  Change edition, then activate     lists the editions this build offers'
        win_a3 = '  [3]  Activate Windows 10 ESU           permanent'
        win_eds = 'Windows editions this build can serve ({0}):'
        win_eds_note = 'Read from the WMI license families. The engine does the change in its own window, where you pick by number.'
        win_change_go = 'Opening the edition changer. Activation is required after it finishes.'
        win_change_ask = 'Run activation now? (Y/n)'
        odt_mirror = 'Microsoft download failed, falling back to the GitHub mirror...'
        odt_pinned = 'The Microsoft download page changed; trying the pinned Microsoft link...'
        odt_verify = 'ODT verified: {0} ({1:N1} MB)'
        off_dl_first = 'The setup tool downloads first, then you are asked for language and components.'
        off_ready = 'Setup tool ready. Now make your choices:'
        off_summary = 'Language: {0}   Installing: {1}'
        off_go = 'Start the install? (Y/n)'
        office_installed = 'Microsoft Office LTSC Professional Plus 2024 is already installed. Remove it? (Y/n, blank = yes)'
        office_installed_other = 'Another Office version is installed ({0}). Remove it? (Y/n, blank = yes)'
        office_removing = 'Removing Office...'
        office_removed = 'Office removed.'
        office_remove_fail = 'Office could not be removed (exit code: {0}).'
        office_install_running = 'Office install is already running. Please wait and try again (Error 0-2035).'
        m0back = '  [0]  Back'
        odt_hash_fail = 'ODT integrity check failed, skipping this source ({0}).'
        win_eds_none = 'Could not read the edition list (WMI). The engine window will still list them.'
        win_change_ask_go = 'Open the edition changer? (Y/n)'
        ce_dl = 'Downloading the edition changer'
        ce_fail = 'The edition changer could not be retrieved, aborting.'
        win_change_title = 'Change edition - pick one'
        win_change_cur   = 'Current edition: {0}'
        win_change_rec   = '[x] (recommended - current edition)'
        win_change_sel   = 'Pick by number, then Enter:'
        win_change_next  = '10) See more (show remaining editions)'
        win_change_back  = '0) Back'
    }
}

function Get-UiLang {
    if ($script:UiLang -eq 'auto') {
        if ((Get-Culture).Name -like 'tr*') { 'tr' } else { 'en' }
    } else { $script:UiLang }
}

function T([string]$Key) {
    $s = $Strings[(Get-UiLang)][$Key]
    if (-not $s) { $Key } else { $s }
}




# Case-insensitive matching that also works under tr-TR: .NET IgnoreCase alone treats 'I' != 'i' in
# Turkish, which silently missed upper-case matches (e.g. "HWID", "LICENSED") three separate times.
function Test-Match([string]$Text, [string]$Pattern) {
    [regex]::IsMatch($Text, $Pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant)
}

function Test-Admin {
    [bool]([Security.Principal.WindowsIdentity]::GetCurrent().Groups -match 'S-1-5-32-544')
}

# The "go back" token. Pulled out of Read-Input so the contract is testable without a console.
function Test-BackToken([string]$Text) {
    Test-Match $Text '^\s*esc\s*$'
}

# Line input that also reports Esc (go back). Read-Host cannot see Esc, so keys are read and echoed here.
# Where no console is available (piped stdin, redirected test runs) it degrades to Read-Host, and the
# literal token "esc" counts as the back action so that path stays testable.
function Read-Input([string]$Prompt = '', [string]$Default = '') {
    $shown = if ($Default) { "$Prompt [$Default]" } else { $Prompt }
    if (-not [Console]::IsInputRedirected) {
        try {
            Write-Host "$shown > " -ForegroundColor Yellow -NoNewline
            $buf = ''
            while ($true) {
                $k = [Console]::ReadKey($true)
                if ($k.Modifiers -band [ConsoleModifiers]::Control -and $k.Key -eq 'C') { Write-Host ''; exit 130 }
                if ($k.Key -eq 'Enter') { Write-Host ''; return @{ Text = $buf; Back = $false } }
                if ($k.Key -eq 'Escape') {
                    Write-Host ''
                    Write-Host ('  << ' + (T 'esc_hint')) -ForegroundColor DarkYellow
                    return @{ Text = ''; Back = $true }
                }
                if ($k.Key -eq 'Backspace') {
                    if ($buf.Length -gt 0) { $buf = $buf.Substring(0, $buf.Length - 1); Write-Host "`b `b" -NoNewline }
                    continue
                }
                if ($k.KeyChar -and [int][char]$k.KeyChar -ge 32) { $buf += $k.KeyChar; Write-Host $k.KeyChar -NoNewline }
            }
        } catch { }
    }
    # No console: read lines directly so that end-of-input is detectable. Read-Host cannot tell an
    # exhausted pipe from a bare Enter, and an exhausted pipe would spin the menu loops forever.
    Write-Host "$shown > " -ForegroundColor Yellow -NoNewline
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) { throw [System.OperationCanceledException]::new('input closed') }
    if (Test-BackToken $line) { return @{ Text = ''; Back = $true } }
    @{ Text = $line; Back = $false }
}

# Yes/no. Enter = yes, Esc = no. TR: E = yes, EN: Y = yes; anything else = no.
function Confirm([string]$Prompt) {
    $r = Read-Input $Prompt
    if ($r.Back) { return $false }
    if (-not $r.Text) { return $true }
    if ((Get-UiLang) -eq 'tr') { return (Test-Match $r.Text '^\s*e(vet)?\s*$') }
    Test-Match $r.Text '^\s*y(es)?\s*$'
}

function Write-Banner {
    Write-Host ''
    Write-Host '=============================================' -ForegroundColor DarkCyan
    Write-Host ("  ActivaterWindows  v$Version") -ForegroundColor Cyan
    Write-Host ("  " + (T 'about_by') -f $script:Author) -ForegroundColor DarkGray
    Write-Host "  $SiteUrl" -ForegroundColor DarkGray
    Write-Host "  $RepoUrl" -ForegroundColor DarkGray
    Write-Host '=============================================' -ForegroundColor DarkCyan
}

# Registry only, and cached: instant. WMI/CIM is far slower (measured 1.6-4.1 s on this machine),
# so nothing on the menu path may call it.
function Get-WinInfo {
    if ($script:WinInfo) { return $script:WinInfo }
    $k = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
    $build = 0
    if ($k -and $k.CurrentBuild) { $build = [int]$k.CurrentBuild }
    # ProductName still reads "Windows 10" on Windows 11 builds, so the family comes from the build number.
    $name = if ($build -ge 22000) { 'Windows 11' } else { 'Windows 10' }
    $ed  = if ($k -and $k.EditionID) { [string]$k.EditionID } else { '?' }
    $ver = if ($k -and $k.DisplayVersion) { [string]$k.DisplayVersion } else { '' }
    $ubr = if ($k -and $null -ne $k.UBR) { [string]$k.UBR } else { '?' }
    $arch = if ([Environment]::Is64BitOperatingSystem) { 'x64' } else { 'x86' }
    $admin = Test-Admin
    $role = if ($admin) { 'admin' } elseif ((Get-UiLang) -eq 'tr') { 'yetkisiz' } else { 'standard user' }
    $script:WinInfo = @{
        Name    = $name
        Edition = $ed
        Build   = $build
        Admin   = $admin
        Summary = ("{0} {1} {2} {3}.{4} {5} [{6}]" -f $name, $ed, $ver, $build, $ubr, $arch, $role)
    }
    $script:WinInfo
}

# The editions this build can serve, deduplicated. The engine reads the same WMI source; the query is
# slow (measured 4.1 s, 61 rows with repeats) so it is only called from the Windows submenu.
function Get-EditionList {
    $q = "SELECT LicenseFamily FROM SoftwareLicensingProduct WHERE ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseDependsOn is NULL"
    try {
        @(Get-CimInstance -Query $q -ErrorAction Stop |
            Select-Object -ExpandProperty LicenseFamily -Unique |
            Where-Object { $_ } | Sort-Object)
    } catch { @() }
}

function Write-Section([string]$Title) {
    Write-Host ''
    Write-Host ('  ' + $Title + ' ' + ('-' * [Math]::Max(3, 52 - $Title.Length))) -ForegroundColor DarkCyan
}

function Test-FullLanguageMode {
    if ($ExecutionContext.SessionState.LanguageMode.value__ -ne 0) {
        Write-Host "PowerShell Full Language Mode'da calismiyor." -ForegroundColor Red
        Write-Host 'Cozum: https://massgrave.dev/fix_powershell' -ForegroundColor White -BackgroundColor Blue
        return $false
    }
    try { [void][System.AppDomain]::CurrentDomain.GetAssemblies(); [void][System.Math]::Sqrt(144) }
    catch {
        
        Write-Host 'Cozum: https://massgrave.dev/in-place_repair_upgrade' -ForegroundColor White -BackgroundColor Blue
        return $false
    }
    $true
}

function Warn-AV {
    $cmd = if ($PSVersionTable.PSVersion.Major -ge 3) { 'Get-CimInstance' } else { 'Get-WmiObject' }
    $av = & $cmd -Namespace root\SecurityCenter2 -Class AntiVirusProduct -ErrorAction SilentlyContinue |
        Where-Object { -not (Test-Match $_.displayName 'windows') } | Select-Object -ExpandProperty displayName
    if ($av) {
        Write-Host '3. parti antivirus scripti engelleyebilir:' -ForegroundColor White -BackgroundColor Blue -NoNewline
        Write-Host " $($av -join ', ')" -ForegroundColor DarkRed -BackgroundColor White
    }
}


function Write-DlBar([long]$Done, [long]$Total, $Sw) {
    $ratio = if ($Total -gt 0) { [math]::Min(1.0, $Done / $Total) } else { 0.0 }
    $pct = [int]([math]::Round($ratio * 100))
    if ($pct -eq $script:_lastPct) { return }
    $script:_lastPct = $pct
    $bars = [int]($ratio * 30)
    $track = ('#' * $bars) + ('-' * (30 - $bars))
    $el = if ($Sw) { [int]$Sw.Elapsed.TotalSeconds } else { 0 }
    Write-Host ("`r  [{0}] {1,3}%  {2,3}s" -f $track, $pct, $el) -NoNewline
}



function Write-BusyBar([int]$Tick, $Sw) {
    $spinner = @('|', '/', '-', '\')
    $spin = $spinner[$Tick % 4]
    $el = if ($Sw) { [int]$Sw.Elapsed.TotalSeconds } else { 0 }
    Write-Host ("`r  {0}  {1,4}s" -f $spin, $el) -NoNewline
}


function Wait-WithBar([scriptblock]$IsDone, [int]$MaxSec = 1800) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $tick = 0
    $fin = $false
    while (-not $fin -and $sw.Elapsed.TotalSeconds -lt $MaxSec) {
        try { $fin = [bool](& $IsDone) } catch { $fin = $true }
        if ($fin) { break }
        Write-BusyBar $tick 30 $sw
        $tick++
        Start-Sleep -Milliseconds 100
    }
    $sw.Stop()
    Write-Host ("`r  {0}  {1,4}s" -f '*', [int]$sw.Elapsed.TotalSeconds)
    $fin
}

function Invoke-Download([string]$Url, [string]$OutFile) {
    # Chunked stream read: WebClient events do not fire during a synchronous call, this does
    $req = [Net.HttpWebRequest]::Create($Url)
    $req.Method = 'GET'
    $req.Timeout = 60000
    $req.UserAgent = 'ActivaterWindows'
    $resp = $req.GetResponse()
    $total = $resp.ContentLength
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $script:_lastPct = -1
    $rs = $resp.GetResponseStream()
    $fs = if ($OutFile) { [IO.File]::Create($OutFile) } else { New-Object IO.MemoryStream }
    try {
        $buf = New-Object byte[] (65536)
        $read = 0; $done = 0
        while (($read = $rs.Read($buf, 0, $buf.Length)) -gt 0) {
            $fs.Write($buf, 0, $read)
            $done += $read
            Write-DlBar $done $total $sw
        }
        Write-Host ''
        if ($fs -is [IO.MemoryStream]) { , $fs.ToArray() }
    } finally {
        $rs.Close(); $fs.Close(); $resp.Close()
    }
}

function Get-WebString([string]$Url) {
    $bytes = Invoke-Download $Url $null
    if ($bytes) { [Text.Encoding]::UTF8.GetString($bytes) } else { $null }
}

function Save-File([string]$Url, [string]$OutFile) {
    Invoke-Download $Url $OutFile | Out-Null
}

function Get-String([string[]]$Urls, [string]$Label, [switch]$Quiet) {
    if (-not $Quiet) { Write-Host $Label -ForegroundColor DarkGray }
    $errors = @(); $response = $null
    foreach ($url in $Urls | Sort-Object { Get-Random }) {
        try {
            $response = Get-WebString $url
            if ($response) { break }
        } catch { $errors += $_ }
    }
    if (-not $response) {
        if (-not $Quiet) {
            Warn-AV
            foreach ($e in $errors) { Write-Host "Hata: $($e.Exception.Message)" -ForegroundColor Red }
            Write-Host ((T 'dl_fail') -f $Label) -ForegroundColor Red
        }
    }
    $response
}

function Get-Sha256([string]$Text) {
    $stream = New-Object IO.MemoryStream
    $writer = New-Object IO.StreamWriter $stream
    $writer.Write($Text); $writer.Flush(); $stream.Position = 0
    ([BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash($stream)) -replace '-')
}

function Test-Hash([string]$Text, [string]$Expected) {
    (Get-Sha256 $Text) -eq $Expected
}

function Get-FileSha256([string]$Path) {
    ([BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash([IO.File]::ReadAllBytes($Path))) -replace '-')
}

# A usable ODT is big enough AND starts with the PE magic "MZ" - a size check alone would accept any
# large file (an HTML error page saved to disk, a truncated download, a captive-portal response).
function Test-OdtFile([string]$Path) {
    try {
        if (-not (Test-Path $Path)) { return $false }
        if ((Get-Item $Path).Length -lt $OdtMinBytes) { return $false }
        $fs = [IO.File]::OpenRead($Path)
        try { $b = New-Object byte[] 2; [void]$fs.Read($b, 0, 2) } finally { $fs.Close() }
        ($b[0] -eq 0x4D -and $b[1] -eq 0x5A)
    } catch { $false }
}

# Office Deployment Tool acquisition - Microsoft's own downloader, four tiers, each verified before use:
#   1) resources/ next to this script (only when running from a repo checkout)
#   2) Microsoft's download page, scraped (Microsoft's documented route; the URL changes per build)
#   3) the pinned Microsoft URL (survives a redesigned download page); hash-checked
#   4) our GitHub release mirror (survives a blocked download.microsoft.com); hash-checked
# Tiers 3 and 4 are the same build, so one hash covers both.
function Get-OdtTool([string]$OutFile) {
    $local = $null
    if ($PSScriptRoot) {
        $local = Get-ChildItem -Path (Join-Path $PSScriptRoot 'resources') -Filter 'officedeploymenttool*.exe' -ErrorAction SilentlyContinue |
            Select-Object -First 1
    }
    if ($local) {
        Write-Host ((T 'odt_local') -f $local.Name) -ForegroundColor DarkGray
        try { Copy-Item $local.FullName $OutFile -Force } catch { }
        if (Test-OdtFile $OutFile) { Write-Host ((T 'odt_verify') -f $local.Name, ((Get-Item $OutFile).Length / 1MB)) -ForegroundColor DarkGreen; return $true }
        Write-Host (T 'odt_small') -ForegroundColor DarkYellow
    }

    $candidates = @()
    try {
        Write-Host (T 'odt_dl') -ForegroundColor DarkGray
        $conf = Invoke-WebRequest -Uri 'https://www.microsoft.com/en-us/download/details.aspx?id=49117' -UseBasicParsing -TimeoutSec 60
        # CultureInvariant is required: under tr-TR a plain IgnoreCase match misses upper-case "I".
        $link = [regex]::Match($conf.Content, 'https://download\.microsoft\.com/[^"''\s<>]*officedeploymenttool[^"''\s<>]*\.exe', [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant).Value
        if ($link) { $candidates += @{ Url = $link; Label = 'microsoft.com'; Hash = '' } }
        else { Write-Host (T 'odt_pinned') -ForegroundColor DarkYellow }
    } catch {
        Write-Host (T 'odt_pinned') -ForegroundColor DarkYellow
    }
    $candidates += @{ Url = $OdtPinnedUrl; Label = 'microsoft.com pinned'; Hash = $OdtPinnedHash }
    $candidates += @{ Url = $OdtMirrorUrl; Label = 'github mirror'; Hash = $OdtPinnedHash }

    foreach ($c in $candidates) {
        if ($c.Label -eq 'github mirror') { Write-Host (T 'odt_mirror') -ForegroundColor DarkYellow }
        Remove-Item -Path $OutFile -Force -ErrorAction SilentlyContinue
        try { Save-File -Url $c.Url -OutFile $OutFile } catch { }
        if (-not (Test-OdtFile $OutFile)) {
            Write-Host ("{0} ({1})" -f (T 'odt_small'), $c.Label) -ForegroundColor DarkYellow
            continue
        }
        if ($c.Hash -and (Get-FileSha256 $OutFile) -ne $c.Hash) {
            Write-Host ((T 'odt_hash_fail') -f $c.Label) -ForegroundColor DarkYellow
            continue
        }
        Write-Host ((T 'odt_verify') -f $c.Label, ((Get-Item $OutFile).Length / 1MB)) -ForegroundColor DarkGreen
        return $true
    }
    $false
}




#   2) Gomulu pin (commit + hash) - upstream'in formati degisirse bile calisir.

function Resolve-Engine {
    
    try {
        $boot = Get-WebString 'https://get.activated.win'
        if ($boot) {
            $mUrl = [regex]::Match($boot, 'raw\.githubusercontent\.com/massgravel/Microsoft-Activation-Scripts/([0-9a-f]{40})/MAS/All-In-One-Version-KL/MAS_AIO\.cmd')
            $mHsh = [regex]::Match($boot, "\`$releaseHash\s*=\s*'([0-9A-Fa-f]{64})'")
            if ($mUrl.Success -and $mHsh.Success) {
                $commit = $mUrl.Groups[1].Value
                $hash = $mHsh.Groups[1].Value.ToUpper()
                $urls = @(
                    "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/$commit/MAS/All-In-One-Version-KL/MAS_AIO.cmd",
                    "https://dev.azure.com/massgrave/Microsoft-Activation-Scripts/_apis/git/repositories/Microsoft-Activation-Scripts/items?path=/MAS/All-In-One-Version-KL/MAS_AIO.cmd&versionType=Commit&version=$commit",
                    "https://git.activated.win/Microsoft-Activation-Scripts/plain/MAS/All-In-One-Version-KL/MAS_AIO.cmd?id=$commit"
                )
                $aio = Get-String $urls (T 'mas_dl') -Quiet
                if ($aio -and (Test-Hash $aio $hash)) {
                    Save-Cache $aio $AioCache $AioCacheHash
                    return @{ Text = $aio; Source = 'upstream'; Commit = $commit }
                }
            }
        }
    } catch { }

    
    $urls = @(
        "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/$Pinned/MAS/All-In-One-Version-KL/MAS_AIO.cmd",
        "https://dev.azure.com/massgrave/Microsoft-Activation-Scripts/_apis/git/repositories/Microsoft-Activation-Scripts/items?path=/MAS/All-In-One-Version-KL/MAS_AIO.cmd&versionType=Commit&version=$Pinned",
        "https://git.activated.win/Microsoft-Activation-Scripts/plain/MAS/All-In-One-Version-KL/MAS_AIO.cmd?id=$Pinned"
    )
    $aio = Get-String $urls (T 'mas_dl')
    if ($aio -and (Test-Hash $aio $AioHash)) {
        Save-Cache $aio $AioCache $AioCacheHash
        return @{ Text = $aio; Source = 'pin'; Commit = $Pinned }
    }

    
    $cached = Read-Cache $AioCache $AioCacheHash
    if ($cached) {
        Write-Host (T 'using_cache') -ForegroundColor Yellow
        return @{ Text = $cached; Source = 'cache'; Commit = $Pinned }
    }

    Write-Host (T 'engine_fail') -ForegroundColor Red
    $null
}

function Save-Cache([string]$Text, [string]$Path, [string]$HashPath) {
    try {
        if (-not (Test-Path $CacheDir)) { New-Item -ItemType Directory -Path $CacheDir -Force | Out-Null }
        [IO.File]::WriteAllText($Path, $Text)
        if ($HashPath) { [IO.File]::WriteAllText($HashPath, (Get-Sha256 $Text)) }
    } catch { }
}


function Read-Cache([string]$Path, [string]$HashPath) {
    try {
        if (-not (Test-Path $Path)) { return $null }
        $text = [IO.File]::ReadAllText($Path)
        if ($HashPath -and (Test-Path $HashPath)) {
            $want = ([IO.File]::ReadAllText($HashPath)).Trim().ToUpperInvariant()
            if ((Get-Sha256 $text) -ne $want) { return $null }
        }
        return $text
    } catch { }
    $null
}






function Get-ResultLines([string]$Text) {
    if (-not $Text) { return @() }
    
    $t = [regex]::Replace($Text, '\x1B\[[0-9;?]*[A-Za-z]', '')
    $t = [regex]::Replace($t, '\x1B[()][0-9A-B]', '')
    $skip = '(?i)ohook|tsforge|kms38|kms4k|kms4|staticcid|zerocid|scid|zcid|\bkms\b|hwid|massgrave|microsoft activation scripts|\bmas\b|activated\.win|github\.com|licence:|license:|gpl|_aio|plugin|switching to|method is not supported|choose a|select a|^\s*\[[0-9]\]'
    $keep = '(?i)licens|activ|success|fail|error|tamamlan|etkin|windows|office|expire|product key'
    @($t -split "`r?`n" |
        ForEach-Object { $_.Trim() } |
        
        Where-Object { $_ -and -not (Test-Match $_ '^[A-Za-z]:\\[^\r\n]*AW_[0-9a-f\-]{36}\.cmd>') -and -not (Test-Match $_ '^\s*>\s') } |
        Where-Object { $_ -and -not (Test-Match $_ $skip) -and (Test-Match $_ $keep) })
}

function Show-ActivationResult([string]$LogPath) {
    if (-not $LogPath -or -not (Test-Path $LogPath)) { return $false }
    
    
    $oem = [Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage)
    $txt = $null
    for ($i = 0; $i -lt 30; $i++) {
        try {
            $fs = [IO.File]::Open($LogPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
            $sr = New-Object IO.StreamReader($fs, $oem)
            $txt = $sr.ReadToEnd()
            $sr.Close(); $fs.Close()
            break
        } catch { $txt = $null; Start-Sleep -Milliseconds 200 }
    }
    if (-not $txt) { return $false }
    $lines = @(Get-ResultLines $txt)
    if ($lines.Count -eq 0) { return $false }
    Write-Host ''
    foreach ($l in ($lines | Select-Object -Last 12)) { Write-Host "  $l" -ForegroundColor DarkGray }
    $true
}

function Invoke-MasCmd([string]$Response, [string]$Arguments, [string]$Label, [switch]$Visible) {
    $rand = [Guid]::NewGuid().Guid
    $isAdmin = [bool]([Security.Principal.WindowsIdentity]::GetCurrent().Groups -match 'S-1-5-32-544')
    $tmp = if ($isAdmin) { "$env:SystemRoot\Temp" } else { "$env:USERPROFILE\AppData\Local\Temp" }
    $file = "$tmp\AW_$rand.cmd"
    $log = "$tmp\AW_$rand.log"
    
    Set-Content -Path $file -Value "@::: $rand`r`n$Response"
    if (-not (Test-Path $file)) { Warn-AV; Write-Host ((T 'tmp_fail') -f $Label) -ForegroundColor Red; return }
    $env:ComSpec = "$env:SystemRoot\system32\cmd.exe"
    $hidden = -not $Visible
    $redir = if ($hidden) { " > ""$log"" 2>&1" } else { '' }
    $sp = @{
        FilePath     = $env:ComSpec
        ArgumentList = "/c """"$file"" -el $Arguments$redir"""
        Verb         = 'RunAs'
        PassThru     = $true
    }
    if ($hidden) { $sp['WindowStyle'] = 'Hidden' } else { $sp['Wait'] = $true }
    try {
        if ($hidden) {
            
            
            
            
            $p = Start-Process @sp
            if (-not (Wait-WithBar { $p.WaitForExit(200) } $HiddenTimeoutSec)) {
                
                
                & "$env:SystemRoot\system32\taskkill.exe" /PID $p.Id /T /F 2>&1 | Out-Null
                Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
                Write-Host ((T 'act_timeout') -f $HiddenTimeoutSec) -ForegroundColor DarkYellow
            }
            if (-not (Show-ActivationResult $log)) {
                
                Write-Host (T 'act_nolog') -ForegroundColor DarkGray
            }
        } else {
            $p = Start-Process @sp
            Write-Host ((T 'done') -f $Label, $p.ExitCode) -ForegroundColor Green
        }
    } catch {
        Write-Host ((T 'uac_cancel2') -f $Label) -ForegroundColor Yellow
    } finally {
        
        
        foreach ($f in @($file, $log)) {
            for ($i = 0; $i -lt 10 -and (Test-Path $f); $i++) {
                Remove-Item -Path $f -Force -ErrorAction SilentlyContinue
                if (Test-Path $f) { Start-Sleep -Milliseconds 150 }
            }
        }
    }
}

function Invoke-MasAio([string]$Arguments = '') {
    $eng = Resolve-Engine
    if (-not $eng) { return }
    if ($Arguments) {
        
        Invoke-MasCmd -Response $eng.Text -Arguments $Arguments -Label $Arguments.Trim('/')
    } else {
        
        Invoke-MasCmd -Response $eng.Text -Arguments '' -Label 'menu' -Visible
    }
}

function Invoke-MasCheck {
    $urls = @(
        "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/$Pinned/MAS/Separate-Files-Version/Check_Activation_Status.cmd",
        "https://dev.azure.com/massgrave/Microsoft-Activation-Scripts/_apis/git/repositories/Microsoft-Activation-Scripts/items?path=/MAS/Separate-Files-Version/Check_Activation_Status.cmd&versionType=Commit&version=$Pinned",
        "https://git.activated.win/Microsoft-Activation-Scripts/plain/MAS/Separate-Files-Version/Check_Activation_Status.cmd?id=$Pinned"
    )
    $chk = Get-String $urls (T 'chk_dl')
    if ($chk -and (Test-Hash $chk $ChkHash)) {
        Save-Cache $chk $ChkCache $ChkCacheHash
    } else {
        $chk = Read-Cache $ChkCache $ChkCacheHash
        if (-not $chk) { return }
        Write-Host (T 'using_cache') -ForegroundColor Yellow
    }
    Invoke-MasCmd -Response $chk -Arguments '' -Label (T 'status_lbl')
}

function Test-OfficeC2R {
    $cfg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\OfficeClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($cfg -and $cfg.ProductReleaseIds) { return $true }
    $cfg32 = Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\OfficeClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($cfg32 -and $cfg32.ProductReleaseIds) { return $true }
    $false
}

# Sadece LTSC Professional Plus 2024 kurulu mu diye bakar (ProPlus2024Volume ProgId)
function Test-OfficeLTSC2024 {
    foreach ($root in @(
        'HKLM:\SOFTWARE\Microsoft\OfficeClickToRun\Configuration',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\OfficeClickToRun\Configuration'
    )) {
        $cfg = Get-ItemProperty -Path $root -ErrorAction SilentlyContinue
        if ($cfg -and $cfg.ProductReleaseIds -and
            ($cfg.ProductReleaseIds -match 'ProPlus2024Volume')) { return $true }
    }
    # Fallback: Office16 klasoru var mi? (Registry yoksa bile)
    $office16 = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($office16) { return $true }
    $false
}

# Herhangi bir C2R Office kurulu mu — LTSC2024 haric (M365, 2021, 2019 vs.)
function Test-OtherOfficeC2R {
    foreach ($root in @(
        'HKLM:\SOFTWARE\Microsoft\OfficeClickToRun\Configuration',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\OfficeClickToRun\Configuration'
    )) {
        $cfg = Get-ItemProperty -Path $root -ErrorAction SilentlyContinue
        if ($cfg -and $cfg.ProductReleaseIds) { return $true }
    }
    # Fallback: Office16/Office15 klasoru var mi?
    $officeRoot = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16",
        "$env:ProgramFiles\Microsoft Office\root\Office15",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office15"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($officeRoot) { return $true }
    $false
}

# Office kurulum/islem icinde mi? setup.exe, OfficeC2RClient.exe, msiexec gibi
# prosesler calisiyorsa ODT kurulumu 0-2035 (already installing) hatasi verir.
function Test-OfficeInstallRunning {
    $procs = @('setup', 'OfficeC2RClient', 'msiexec', 'clicktorun',
               'officeclicktorun', 'spvirtualization')
    foreach ($p in $procs) {
        if (Get-Process -Name $p -ErrorAction SilentlyContinue) { return $true }
    }
    try {
        $svc = Get-Service -Name 'ClickToRunSvc' -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -eq 'Running') { return $true }
    } catch {}
    $false
}

# Office Click-to-Run kurulumundan sonra .docx/.xlsx/.pptx gibi dosyalarin cift tikla
# acilmamasi, sadece yonetici izniyle acilmasi, word'den acmak zorunda kalma sorununu giderir.
# ProgID'leri yeniden baglar, UserChoice kayitlarini temizler, Word/Excel/PowerPoint
# icin varsayilan uygulamayi kaydeder.
function Invoke-FixOfficeAssociations {
    Clear-Host
    Write-Banner
    Write-Host ''
    Write-Host '  Office tamami onariliyor...' -ForegroundColor Cyan
    Write-Host ''

    if (-not (Test-OfficeC2R)) {
        Write-Host '  Office kurulu degil. Once Office kurun.' -ForegroundColor Red
        return
    }

    # --- 1) Eski hesap / OneAuth kimlik kayitlarini temizle ---
    # Hesap hatasi + Word'un yonetici gerektirmesi buradan geliyor:
    # eski M365/Live hesabi LTSC'ye baglanmis kalmis, Office kimlik dogrulama
    # yapmaya calisip elevated aciyor.
    Write-Host '  Eski hesap kayitlari temizleniyor...' -ForegroundColor Cyan
    $idRoot = 'HKCU:\Software\Microsoft\Office\16.0\Common\Identity'
    if (Test-Path $idRoot) {
        # Alt anahtarlari temizle (Profiles, OneAuthConfigCache)
        foreach ($sub in Get-ChildItem $idRoot -ErrorAction SilentlyContinue) {
            try { Remove-Item -Path $sub.PSPath -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        }
        # ConnectedOneAuthAccountId degerini sifirla
        try { Set-ItemProperty -Path $idRoot -Name 'ConnectedOneAuthAccountId' -Value '' -Force -ErrorAction SilentlyContinue } catch {}
        Write-Host '  Identity kayitlari temizlendi.' -ForegroundColor Green
    } else {
        Write-Host '  Temizlenecek Identity kaydi yok.' -ForegroundColor DarkGray
    }

    # Ek hesap kalintisi: MSAL token cache
    $msalKeys = @(
        'HKCU:\Software\Microsoft\Office\16.0\Common\Identity\OneAuthConfigCache',
        'HKCU:\Software\Microsoft\Office\16.0\Common\ServicesManagerCache',
        'HKCU:\Software\Microsoft\Office\16.0\Common\Roaming\Identities'
    )
    foreach ($k in $msalKeys) {
        if (Test-Path $k) {
            try { Remove-Item -Path $k -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        }
    }

    # --- 2) Office16 klasor ve exe izinlerini duzelt ---
    # Word/Excel/PowerPoint yonetici olmadan acilmiyorsa izinler bozuktur.
    $officeRoot = $null
    foreach ($root in @(
        "$env:ProgramFiles\Microsoft Office\root\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16",
        "$env:ProgramFiles\Microsoft Office\root\Office15",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office15"
    )) { if (Test-Path $root) { $officeRoot = $root; break } }

    if ($officeRoot) {
        Write-Host ('  Dosya izinleri onariliyor: {0}' -f $officeRoot) -ForegroundColor Cyan
        $usersRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            'Users', 'ReadAndExecute, Synchronize',
            'ContainerInherit,ObjectInherit', 'None', 'Allow')
        try {
            $acl = Get-Acl $officeRoot
            $acl.SetAccessRule($usersRule)
            Set-Acl -Path $officeRoot -AclObject $acl
            Write-Host '  Klasor izinleri guncellendi.' -ForegroundColor Green
        } catch { Write-Host ('  Klasor izin hatasi: {0}' -f $_.Exception.Message) -ForegroundColor Yellow }

        foreach ($exe in @('winword.exe','excel.exe','powerpnt.exe')) {
            $exePath = Join-Path $officeRoot $exe
            if (-not (Test-Path $exePath)) { continue }
            try {
                $acl = Get-Acl $exePath
                $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                    'Users', 'ReadAndExecute, Synchronize', 'None', 'None', 'Allow')
                $acl.SetAccessRule($rule)
                Set-Acl -Path $exePath -AclObject $acl
            } catch {}
        }
    }

    # --- 3) HKCR ProgID + App Paths yeniden yaz ---
    if ($officeRoot) {
        $exeMap = @{
            'winword.exe'  = @('Word.Document.12','Word.Document.8','Word.Template.12','Word.Template.8','Word.RTF.8')
            'excel.exe'    = @('Excel.Sheet.12','Excel.Sheet.8','Excel.SheetMacroEnabled.12','Excel.SheetBinaryMacroEnabled.12','Excel.Template.12')
            'powerpnt.exe' = @('PowerPoint.Show.12','PowerPoint.Show.8','PowerPoint.ShowMacroEnabled.12','PowerPoint.Template.12')
        }
        foreach ($exe in $exeMap.Keys) {
            $exePath = Join-Path $officeRoot $exe
            if (-not (Test-Path $exePath)) { continue }
            foreach ($progId in $exeMap[$exe]) {
                $cmdKey = "HKCR:\$progId\shell\Open\command"
                try {
                    if (-not (Test-Path $cmdKey)) { New-Item -Path $cmdKey -Force | Out-Null }
                    Set-ItemProperty -Path $cmdKey -Name '(Default)' -Value "`"$exePath`" `"%1`"" -Force
                } catch {}
            }
            $apKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$exe"
            try {
                if (-not (Test-Path $apKey)) { New-Item -Path $apKey -Force | Out-Null }
                Set-ItemProperty -Path $apKey -Name '(Default)' -Value $exePath -Force
                Set-ItemProperty -Path $apKey -Name 'Path' -Value $officeRoot -Force
            } catch {}
        }
    }

    # --- 4) UserChoice + OpenWithProgids temizle ---
    $extList = @('.doc','.docx','.dot','.dotx','.rtf','.xls','.xlsx','.xlsm','.xlsb','.xltx','.ppt','.pptx','.pptm','.potx')
    foreach ($ext in $extList) {
        foreach ($sub in @('UserChoice','OpenWithProgids')) {
            $k = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$ext\$sub"
            if (Test-Path $k) { try { Remove-Item -Path $k -Recurse -Force -ErrorAction SilentlyContinue } catch {} }
        }
    }

    # --- 5) OfficeC2RClient /repair + ie4uinit + explorer ---
    $c2rClient = $null
    foreach ($p in @(
        "$env:ProgramFiles\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe",
        "${env:ProgramFiles(x86)}\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe"
    )) { if (Test-Path $p) { $c2rClient = $p; break } }

    if ($c2rClient) {
        Write-Host '  Office iliskilendirmeleri yeniden yaziliyor...' -ForegroundColor Cyan
        try {
            $pr = Start-Process -FilePath $c2rClient -ArgumentList '/repair onshowinuisuppressed=true' -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue
            if ($pr) { if (-not (Wait-WithBar { $pr.HasExited } 300)) { Write-Host (T 'step_timeout') -ForegroundColor DarkYellow } }
        } catch {}
    }

    $ie4 = "$env:SystemRoot\System32\ie4uinit.exe"
    if (Test-Path $ie4) {
        try { Start-Process -FilePath $ie4 -ArgumentList '-show' -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue } catch {}
    }

    try { Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue } catch {}
    Start-Sleep -Milliseconds 1500
    try { Start-Process explorer.exe } catch {}

    Write-Host ''
    Write-Host '  Tamamlandi.' -ForegroundColor Green
    Write-Host '  - Hesap hatasi giderildi (eski M365 hesabi kaldinildi)' -ForegroundColor DarkGray
    Write-Host '  - Word/Excel/PowerPoint artik yonetici olmadan acilmali' -ForegroundColor DarkGray
    Write-Host '  - .docx/.xlsx/.pptx cift tikla acilmali' -ForegroundColor DarkGray
    Write-Host '  Sorun devam ederse oturumu kapat/ac veya bilgisayari yeniden baslat.' -ForegroundColor DarkYellow
    Write-Host ''
}

function Default-OfficeLang {
    $c = ((Get-Culture).Name).ToLowerInvariant()
    if ($script:UiLang -eq 'tr') { 'tr-tr' }
    elseif ($script:UiLang -eq 'en') { 'en-us' }
    else { $c }
}


function New-AppSelection {
    $sel = @{}
    foreach ($a in $OfficeApps) { $sel[$a.id] = $false }
    $sel
}

function Show-AppPicker($sel) {
    $count = @($OfficeApps | Where-Object { $sel[$_.id] }).Count
    $head = T 'app_head'
    $cnt  = (T 'app_count') -f $count, $OfficeApps.Count
    $pad  = 58 - $head.Length - $cnt.Length
    if ($pad -lt 2) { $pad = 2 }
    Write-Host ''
    Write-Host ('  ' + ('=' * 58)) -ForegroundColor DarkCyan
    Write-Host ('   ' + $head) -ForegroundColor Cyan -NoNewline
    Write-Host ((' ' * $pad) + $cnt) -ForegroundColor Yellow
    Write-Host ('  ' + ('=' * 58)) -ForegroundColor DarkCyan
    for ($i = 0; $i -lt $OfficeApps.Count; $i++) {
        $a = $OfficeApps[$i]
        $isOn = [bool]$sel[$a.id]
        $mark = if ($isOn) { '[x]' } else { '[ ]' }
        $col = if ($isOn) { 'Green' } else { 'DarkGray' }
        Write-Host ('   {0,2}) ' -f ($i + 1)) -ForegroundColor Cyan -NoNewline
        Write-Host ('{0} ' -f $mark) -ForegroundColor $col -NoNewline
        Write-Host ('{0,-19}' -f $a.n) -ForegroundColor $col -NoNewline
        if ($OfficeAppsRecommended -contains $a.id) { Write-Host (T 'app_rec') -ForegroundColor DarkCyan } else { Write-Host '' }
    }
    Write-Host ''
    Write-Host ('  ' + (T 'app_hint')) -ForegroundColor DarkYellow
    Write-Host ''
}

function Select-OfficeApps {
    $sel = New-AppSelection
    while ($true) {
        Show-AppPicker $sel
        $r = Read-Input (T 'app_sel')
        if ($r.Back) { return $null }
        if (-not $r.Text) { break }
        $low = $r.Text.Trim().ToLowerInvariant()
        if ($low -eq 'a') { foreach ($a in $OfficeApps) { $sel[$a.id] = $true }; continue }
        if ($low -eq 'n') { foreach ($a in $OfficeApps) { $sel[$a.id] = $false }; continue }
        if ($low -eq 'r') { foreach ($a in $OfficeApps) { $sel[$a.id] = ($OfficeAppsRecommended -contains $a.id) }; continue }
        foreach ($part in ($r.Text -split ',')) {
            $p = $part.Trim()
            if ($p -match '^\d+$') {
                $idx = [int]$p - 1
                if ($idx -ge 0 -and $idx -lt $OfficeApps.Count) {
                    $id = $OfficeApps[$idx].id
                    $sel[$id] = -not [bool]$sel[$id]
                }
            }
        }
    }
    $chosen = @($OfficeApps | Where-Object { $sel[$_.id] })
    if ($chosen.Count -eq 0) { Write-Host (T 'app_none') -ForegroundColor Red; return (Select-OfficeApps) }
    Write-Host ((T 'app_done') -f $chosen.Count, ($OfficeApps.Count - $chosen.Count)) -ForegroundColor Green
    Write-Host ((T 'app_chosen') -f (($chosen | ForEach-Object { $_.n }) -join ', ')) -ForegroundColor DarkGreen
    $chosen
}

function Show-LangList {
    Write-Host ''
    Write-Host (T 'lang_head') -ForegroundColor DarkCyan
    for ($i = 0; $i -lt $OfficeLangs.Count; $i += 5) {
        $row = ''
        for ($j = $i; $j -lt ($i + 5) -and $j -lt $OfficeLangs.Count; $j++) {
            $row += ('{0,3}) {1,-12}' -f ($j + 1), $OfficeLangs[$j])
        }
        Write-Host "  $row" -ForegroundColor DarkGray
    }
    Write-Host ''
}

function Resolve-OfficeLang {
    $def = Default-OfficeLang
    while ($true) {
        Show-LangList
        $r = Read-Input ((T 'olang') -f $def)
        if ($r.Back) { return $null }
        if (-not $r.Text) { return $def }
        if ($r.Text -match '^\d+$') {
            $n = [int]$r.Text
            if ($n -ge 1 -and $n -le $OfficeLangs.Count) {
                $code = $OfficeLangs[$n - 1]
                $c = Read-Input ((T 'olang_confirm') -f $n, $code)
                if ($c.Back -or $c.Text -eq '0') { continue }
                return $code
            }
            Write-Host (T 'invalid_lang') -ForegroundColor Red
            continue
        }
        if (Test-Match $r.Text '^[a-z]{2}(-[a-z0-9]{2,8})?$') { return $r.Text.ToLowerInvariant() }
        Write-Host (T 'invalid_lang') -ForegroundColor Red
    }
}

function Get-OfficeConfigXml {
    param([string]$Lang, [string[]]$ExcludeIds = @())
    $excludes = ($ExcludeIds | ForEach-Object { "`n    <ExcludeApp ID=`"$_`" />" }) -join ''
    @"
<Configuration ID="aw-office-ltsc-2024">
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume" PIDKEY="XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB">
    <Language ID="$Lang" />$excludes
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Property Name="PINNEDICONS" Value="FALSE" />
  <Property Name="MigrateArch" Value="FALSE" />
  <Updates Enabled="FALSE" />
  <RemoveMSI />
</Configuration>
"@
}

function Install-OfficeLTSC {
    param([switch]$SkipActivate)
    Write-Host ''
    Write-Host ('  ' + ('=' * 58)) -ForegroundColor DarkCyan
    Write-Host '   Office LTSC Professional Plus 2024' -ForegroundColor Cyan
    Write-Host ('  ' + ('=' * 58)) -ForegroundColor DarkCyan
    Write-Host ('  ' + (T 'off_dl_first')) -ForegroundColor DarkGray

    # Kurulu Office tespit et, adini goster, sor
    $removeOffice = $false
    if (Test-OfficeLTSC2024) {
        # Tam olarak LTSC 2024 kurulu
        if (-not (Confirm (T 'office_installed'))) {
            Write-Host '  Islem iptal edildi.' -ForegroundColor DarkYellow
            return
        }
        $removeOffice = $true
    } elseif (Test-OtherOfficeC2R) {
        # Baska bir C2R Office surumu kurulu (M365, 2021, 2019 vs.)
        $detectedName = 'Bilinmeyen Office'
        try {
            foreach ($root in @(
                'HKLM:\SOFTWARE\Microsoft\OfficeClickToRun\Configuration',
                'HKLM:\SOFTWARE\WOW6432Node\Microsoft\OfficeClickToRun\Configuration'
            )) {
                $cfg = Get-ItemProperty -Path $root -ErrorAction SilentlyContinue
                if ($cfg -and $cfg.ProductReleaseIds) {
                    $detectedName = $cfg.ProductReleaseIds -replace ',.*'
                    break
                }
            }
        } catch {}
        if (-not (Confirm ((T 'office_installed_other') -f $detectedName))) {
            Write-Host '  Islem iptal edildi.' -ForegroundColor DarkYellow
            return
        }
        $removeOffice = $true
    }

    $work = Join-Path ([IO.Path]::GetTempPath()) "AW_ODT_$(Get-Random)"
    $script:WorkDir = $work
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    $odt = Join-Path $work 'odt.exe'

    while ($true) {
        if (-not (Test-OdtFile $odt)) {
            if (-not (Get-OdtTool -OutFile $odt)) {
                Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
                return
            }
        }
        $setup = Join-Path $work 'setup.exe'
        if (-not (Test-Path $setup)) {
            Write-Host (T 'extract') -ForegroundColor Cyan
            $px = Start-Process -FilePath $odt -ArgumentList "/extract:`"$work`" /quiet" -PassThru -WindowStyle Hidden
            if (-not (Wait-WithBar { $px.HasExited } 300)) { Write-Host (T 'step_timeout') -ForegroundColor Red }
        }
        if (-not (Test-Path $setup)) {
            Write-Host (T 'extract_fail') -ForegroundColor Red
            Remove-Item -Path $odt -Force -ErrorAction SilentlyContinue
            continue
        }

        # Office varsa ODT setup.exe /configure Remove.xml ile kaldir
        if ($removeOffice) {
            Write-Host ''
            Write-Host ('  ' + (T 'office_removing')) -ForegroundColor Cyan
            $removeXml = Join-Path $work 'Remove.xml'
            [IO.File]::WriteAllText($removeXml, '<Configuration><Remove All="TRUE"/><Display Level="None" AcceptEULA="TRUE"/></Configuration>')
            $isAdminNow = Test-Admin
            if ($isAdminNow) {
                $pr = Start-Process -FilePath $setup -ArgumentList "/configure `"$removeXml`"" -PassThru -WindowStyle Hidden
                if (-not (Wait-WithBar { $pr.HasExited } 900)) {
                    Write-Host (T 'step_timeout') -ForegroundColor Red
                    try { Stop-Process -Id $pr.Id -Force -ErrorAction SilentlyContinue } catch {}
                    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
                    return
                }
                if ($pr.ExitCode -ne 0) {
                    Write-Host ((T 'office_remove_fail') -f $pr.ExitCode) -ForegroundColor Red
                    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
                    return
                }
            } else {
                Write-Host (T 'inst_uac') -ForegroundColor Cyan
                $rcFile2 = Join-Path $work 'remove_rc.txt'
                try {
                    $pr = Start-Process -FilePath 'powershell.exe' -Verb RunAs -WindowStyle Hidden -PassThru `
                        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"& '$setup' /configure '$removeXml'; [IO.File]::WriteAllText('$rcFile2', `$LASTEXITCODE)`""
                    if (-not (Wait-WithBar { (Test-Path $rcFile2) -or $pr.HasExited } 900)) {
                        Write-Host (T 'step_timeout') -ForegroundColor Red
                        Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
                        return
                    }
                    if (Test-Path $rcFile2) {
                        $rcTxt = (Get-Content -LiteralPath $rcFile2 -Raw).Trim()
                        if ($rcTxt -match '^-?\d+$' -and [int]$rcTxt -ne 0) {
                            Write-Host ((T 'office_remove_fail') -f $rcTxt) -ForegroundColor Red
                            Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
                            return
                        }
                    }
                } catch {
                    Write-Host (T 'uac_cancel') -ForegroundColor Yellow
                    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
                    return
                }
            }
            Write-Host (T 'office_removed') -ForegroundColor Green
            $removeOffice = $false  # bir daha girme
        }

        # Office kurulum/islem zaten calisiyorsa 0-2035 hatasi alinir, engelle.
        if (Test-OfficeInstallRunning) {
            Write-Host ''
            Write-Host (T 'office_install_running') -ForegroundColor Red
            Write-Host '  Islem iptal edildi.' -ForegroundColor DarkYellow
            Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
            return
        }

        # 2) Now the questions.
        Write-Host ''
        Write-Host ('  ' + (T 'off_ready')) -ForegroundColor Cyan
        $lang = Resolve-OfficeLang
        if ($null -eq $lang) { continue }

        Write-Host ''
        Write-Host ('  ' + (T 'mode_head')) -ForegroundColor Yellow
        Write-Host (T 'mode_h') -ForegroundColor Gray
        Write-Host (T 'mode_t') -ForegroundColor Gray
        Write-Host (T 'mode_c') -ForegroundColor Gray
        Write-Host ''
        $r = Read-Input ((T 'sel12') + '   ' + (T 'esc_hint'))
        if ($r.Back) { continue }
        $mode = $r.Text.Trim()

        if ($mode -eq '3') {
            $chosen = @(Select-OfficeApps)
            if (-not $chosen) { continue }
            $chosenIds = $chosen | ForEach-Object { $_.id }
            $excludeIds = $OfficeApps | ForEach-Object { $_.id } | Where-Object { $chosenIds -notcontains $_ }
            $chosenNames = ($chosen | ForEach-Object { $_.n }) -join ', '
        } elseif ($mode -eq '2') {
            $excludeIds = @()
            $chosenNames = ($OfficeApps | ForEach-Object { $_.n }) -join ', '
        } else {
            $excludeIds = @('Access', 'Lync', 'OneDrive', 'OneNote', 'Outlook', 'Publisher')
            $chosenNames = ($OfficeApps | Where-Object { $OfficeAppsRecommended -contains $_.id } | ForEach-Object { $_.n }) -join ', '
        }

        Write-Host ''
        Write-Host ('  ' + ((T 'off_summary') -f $lang, $chosenNames)) -ForegroundColor Green
        if (-not (Confirm (T 'off_go'))) { continue }
        break
    }

    $config = Get-OfficeConfigXml -Lang $lang -ExcludeIds $excludeIds
    $configPath = Join-Path $work 'Configuration.xml'
    [IO.File]::WriteAllText($configPath, $config)

    $isAdmin = Test-Admin
    $rcFile = Join-Path $work 'rc.txt'
    $rc = $null
    try {
        if ($isAdmin) {
            Write-Host (T 'inst_start') -ForegroundColor Cyan
            $p = Start-Process -FilePath $setup -ArgumentList "/configure `"$configPath`"" -PassThru -WindowStyle Hidden
            if (-not (Wait-WithBar { $p.HasExited } 3600)) { Write-Host (T 'step_timeout') -ForegroundColor Red }
            $rc = $p.ExitCode
        } else {
            
            
            Write-Host (T 'inst_uac') -ForegroundColor Cyan
            $p = Start-Process -FilePath 'powershell.exe' -Verb RunAs -WindowStyle Hidden -PassThru -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"& '$setup' /configure '$configPath'; Set-Content -LiteralPath '$rcFile' -Value `$LASTEXITCODE -Encoding ASCII"
            if (-not (Wait-WithBar { (Test-Path $rcFile) -or $p.HasExited } 3600)) { Write-Host (T 'step_timeout') -ForegroundColor Red }
            if (Test-Path $rcFile) {
                $rcTxt = (Get-Content -LiteralPath $rcFile -Raw).Trim()
                if ($rcTxt -match '^-?\d+$') { $rc = [int]$rcTxt }
            }
        }
    } catch {
        Write-Host (T 'uac_cancel') -ForegroundColor Yellow
        Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
        return
    }
    if ($rc -eq 0) { Write-Host ((T 'inst_done') -f $rc) -ForegroundColor Green }
    elseif ($rc) { Write-Host ((T 'inst_fail') -f $rc) -ForegroundColor Red }
    else { Write-Host ((T 'inst_fail') -f '?') -ForegroundColor Yellow }
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue

    if (-not $SkipActivate) {
        if (Confirm (T 'office_act_ask')) { Invoke-MasAio '/Ohook' }
    }
}

# Windows 10 is build < 22000; only there is the ESU method relevant. Uses the cached registry read
# instead of a WMI query, because this runs on the [7]/[8] path too.
function Test-Win10 {
    $wi = Get-WinInfo
    ($wi.Build -gt 0 -and $wi.Build -lt 22000)
}

# [6] Office LTSC 2024 kur + etkinlestir — onceki surumu sessizce temizler, kurar, etkinlestirir.
function Invoke-OfficeReinstallActivate {
    Write-Host ''
    Write-Host ('  ' + ('=' * 58)) -ForegroundColor DarkCyan
    Write-Host '   Office LTSC Professional Plus 2024' -ForegroundColor Cyan
    Write-Host ('  ' + ('=' * 58)) -ForegroundColor DarkCyan
    Write-Host ''

    if (Test-OfficeC2R) {
        if (-not (Confirm (T 'office_installed'))) {
            Write-Host '  Islem iptal edildi.' -ForegroundColor DarkYellow
            return
        }
    }

    Install-OfficeLTSC
}

# [7] ACTIVATE EVERYTHING - activation only, no install.
# Windows: permanent digital license (survives updates). Office: permanent. Plus ESU on Windows 10.
# Arguments go in a single order-independent pass; the engine handles them in any order.
function Invoke-AllActivate {
    Write-Host ''
    Write-Host (T 'act_start') -ForegroundColor Magenta
    $methods = @('/HWID')
    if (Test-OfficeC2R) { $methods += '/Ohook' } else { Write-Host (T 'act_nooffice') -ForegroundColor DarkYellow }
    if (Test-Win10) { $methods += '/Z-ESU' }
    Invoke-MasAio ($methods -join ' ')
    Invoke-MasCheck
}

# [9] DO EVERYTHING - install Office when missing, then the [8] flow, then offer the full engine menu.
function Invoke-AllInOne {
    Write-Host ''
    Write-Host (T 'all_start') -ForegroundColor Magenta
    # Install-OfficeLTSC kendi icinde kurulu mu diye kontrol eder ve
    # varsa 'Temizlensin mi? (E/h, bos = evet)' sorar.
    Install-OfficeLTSC -SkipActivate
    Invoke-AllActivate
    if (Confirm (T 'all_masmenu')) { Invoke-MasAio }
    Write-Host (T 'all_done') -ForegroundColor Green
}

# The language selector accepts the same letters the header advertises ([A]/[T]/[E]) and 1/2/3.
function Set-UiLang {
    Write-Host ''
    Write-Host ((T 'langpick') -f ((Get-Culture).Name)) -ForegroundColor Cyan
    $r = Read-Input ((T 'choice') + ' (A/T/E)')
    if ($r.Back) { return }
    switch ($r.Text.Trim().ToLowerInvariant()) {
        '1' { $script:UiLang = 'auto' }
        '2' { $script:UiLang = 'tr' }
        '3' { $script:UiLang = 'en' }
        'a' { $script:UiLang = 'auto' }
        't' { $script:UiLang = 'tr' }
        'e' { $script:UiLang = 'en' }
    }
    $script:WinInfo = $null   # the summary line is translated, so rebuild it in the new language
}

# [9] About - version, author, license, and where the engine comes from.
function Show-About {
    Clear-Host
    Write-Banner
    Write-Host ''
    Write-Host ('  ' + (T 'about_head')) -ForegroundColor Cyan
    Write-Host ('  ' + ((T 'about_ver') -f $Version)) -ForegroundColor Gray
    Write-Host ('  ' + ((T 'about_by') -f $Author)) -ForegroundColor Green
    Write-Host ('  ' + (T 'about_lic')) -ForegroundColor Gray
    Write-Host ('  ' + ((T 'about_site') -f $SiteUrl)) -ForegroundColor Gray
    Write-Host ('  ' + ((T 'about_repo') -f $RepoUrl)) -ForegroundColor Gray
    Write-Host ('  ' + ((T 'about_short') -f $ShortUrl)) -ForegroundColor Gray
    Write-Host ''
    Write-Host ('  ' + (T 'about_note')) -ForegroundColor DarkGray
    Write-Host ''
    [void](Read-Input '')
}

# Edition info: tech id -> friendly name + short purpose. ASCII-safe Turkish.
function Get-EditionInfo([string]$TechId) {
    switch ($TechId) {
        'Core'              { @{ Name='Windows 11 Home';           Desc='Gunluk kullanim, oyun   (ev kullanicilari)' } }
        'CoreSingleLanguage'{ @{ Name='Windows 11 Home Tek Dil';   Desc='Tek goruntuleme dili   (Home varyanti)' } }
        'Professional'      { @{ Name='Windows 11 Pro';            Desc='Gelismis yonetim, BitLocker, Hyper-V   (yazilimcilar, ileri kullanicilar)' } }
        'ProfessionalEducation'{ @{ Name='Windows 11 Pro Education'; Desc='Egitim odakli Pro   (okular, ogrenciler)' } }
        'Education'         { @{ Name='Windows 11 Education';      Desc='Egitim kurumlari icin gelismis ozellikler   (okular, universiteler)' } }
        'Enterprise'        { @{ Name='Windows 11 Enterprise';     Desc='Kurumsal guvenlik ve yonetim   (buyuk sirketler)' } }
        'ProfessionalWorkstation'{ @{ Name='Windows 11 Pro for Workstations'; Desc='Agir is yuklu, profesyonel donanim   (is istasyonlari)' } }
        'IoTEnterprise'     { @{ Name='Windows IoT Enterprise';    Desc='Endustriyel / ozel amacli cihazlar   (ureticiler)' } }
        'CloudEdition'      { @{ Name='Windows 11 Cloud Edition';  Desc='Ozel bulut odakli dagitimlar' } }
        'ServerRdsh'        { @{ Name='Windows 11 Coklu Oturum';   Desc='Sanal masaustu ve coklu oturum' } }
        default             { @{ Name=$TechId; Desc='' } }
    }
}

# Edition change. The engine's own changer holds the product-key logic per edition and its own
# verified list, so it is delegated to (in a visible window, because it needs interactive input).
# Reimplementing that here would risk putting the machine on a wrong edition.
function Invoke-ChangeEdition {
    $eds = @(Get-EditionList)
    if (-not $eds.Count) {
        Write-Host ''
        Write-Host ('  ' + (T 'win_eds_none')) -ForegroundColor DarkYellow
        Write-Host ''
        if (-not (Confirm (T 'win_change_ask_go'))) { return }
        $launchOnly = $true
    } else {
        $launchOnly = $false
    }

    $wi = Get-WinInfo
    $currentTech = $wi.Edition
    $page = 0; $pageSize = 50
    while ($true) {
        Clear-Host
        Write-Banner
        Write-Host ''
        Write-Host ('  ' + (T 'win_change_title')) -ForegroundColor Cyan
        Write-Host ('  ' + ((T 'win_change_cur') -f $currentTech)) -ForegroundColor DarkGray
        Write-Host ''
        $shownAny = $false
        if (-not $launchOnly) {
            $start = $page * $pageSize
            $end   = [Math]::Min($start + $pageSize, $eds.Count)
            for ($i = $start; $i -lt $end; $i++) {
                $info = Get-EditionInfo $eds[$i]
                $isRec = ($eds[$i] -eq $currentTech)
                $mark  = '[ ]'
                $tag   = if ($isRec) { '  (mevcut)' } else { '' }
                Write-Host ('  {0,2}) {1}  {2,-34}{3}{4}' -f ($i + 1), $mark, $info.Name, $info.Desc, $tag) -ForegroundColor Gray
                $shownAny = $true
            }
            Write-Host ''
        }
        # "0) Geri" kaldırıldı - ESC ile geri dönülür
        Write-Host ('  ' + (T 'win_change_sel')) -ForegroundColor Yellow
        $r = Read-Input ''
        if ($r.Back) { return }
        $sel = $r.Text.Trim()
        if (-not $sel) { continue }
        if ($launchOnly) {
            Write-Host ('  Surum listesi gecersiz - motorun kendi penceresi listeler.') -ForegroundColor Red
            Start-Sleep -Milliseconds 700
            continue
        }
        $n = [int]$sel
        if ($n -lt 1 -or $n -gt $eds.Count) {
            Write-Host ('  Gecersiz secim.') -ForegroundColor Red
            Start-Sleep -Milliseconds 700
            continue
        }
        $pickTech = $eds[$n - 1]
        Write-Host ''
        if (-not (Confirm (T 'win_change_ask_go'))) { continue }

        $eng = Resolve-Engine
        if (-not $eng) { return }
        # Same commit as the already verified engine, so the companion file is anchored to that commit.
        $urls = @(
            "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/$($eng.Commit)/MAS/Separate-Files-Version/Change_Windows_Edition.cmd",
            "https://dev.azure.com/massgrave/Microsoft-Activation-Scripts/_apis/git/repositories/Microsoft-Activation-Scripts/items?path=/MAS/Separate-Files-Version/Change_Windows_Edition.cmd&versionType=Commit&version=$($eng.Commit)"
        )
        $ce = Get-String $urls (T 'ce_dl')
        if (-not $ce -or -not (Test-Match $ce ':cedmenu2')) {
            Write-Host (T 'ce_fail') -ForegroundColor Red
            return
        }
        Write-Host (T 'win_change_go') -ForegroundColor Yellow
        Invoke-MasCmd -Response $ce -Arguments '' -Label 'edition' -Visible
        if (Confirm (T 'win_change_ask')) { Invoke-MasAio '/HWID' }
        return
    }
}

# [1] Windows. A submenu, because "activate Windows" and "change which edition I am on, then
# activate" are different jobs and the second one needs to show what this build can serve.
function Invoke-WindowsMenu {
    while ($true) {
        $wi = Get-WinInfo
        Clear-Host
        Write-Banner
        Write-Host ''
        Write-Host ('  ' + (T 'win_head')) -ForegroundColor Cyan
        Write-Host ('  ' + ((T 'win_cur') -f $wi.Edition, $wi.Build, $wi.Summary)) -ForegroundColor DarkGray
        Write-Host (T 'win_a1') -ForegroundColor White
        Write-Host (T 'win_a2') -ForegroundColor White
        Write-Host (T 'win_a3') -ForegroundColor White
        Write-Host ''
        Write-Host (T 'm0back') -ForegroundColor DarkGray
        Write-Host ('  ' + (T 'esc_hint')) -ForegroundColor DarkGray
        Write-Host ''
        $r = Read-Input (T 'choice')
        if ($r.Back -or -not $r.Text) { return }
        switch ($r.Text.Trim()) {
            '1' { Invoke-MasAio '/HWID'; return }
            '2' { Invoke-ChangeEdition }
            '3' { Invoke-MasAio '/Z-ESU'; return }
            '0' { return }
        }
    }
}

# One dispatch table keeps the menu text, the numbers and the actions in step.
function Invoke-Action([string]$Key) {
    switch ($Key) {
        '1' { Invoke-WindowsMenu }
        '2' { Invoke-ChangeEdition }
        '3' { Invoke-MasAio '/Z-ESU' }
        '4' { Invoke-MasAio '/Ohook' }
        '5' { Install-OfficeLTSC }
        '6' { Invoke-OfficeReinstallActivate }
        '7' { Invoke-MasAio '/Z-WindowsESUOffice' }
        '8' { Invoke-AllActivate }
        '9' { Invoke-AllInOne }
        'a' { Invoke-MasCheck }
# [F] kaldirildi - calismiyordu
        'h' { Show-About }
        'l' { Set-UiLang }
    }
}

# The main screen. Grouped by task, numbers stay sequential, colors consistent.
function Show-MainMenu {
    # Terminal boyutu: 750x450 px (yaklaşık 100x30 karakter) - sadece interaktif konsolda
    if (-not [Console]::IsInputRedirected) {
        try {
            $w = $host.UI.RawUI.WindowSize
            $w.Width = 106
            $w.Height = 37
            $host.UI.RawUI.WindowSize = $w
            $host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size(106, 500)
            Set-ConsoleTitle 'ActivaterWindows'
        } catch {}
    }
    
    Clear-Host
    Write-Banner
    $wi = Get-WinInfo
    Write-Host ''
    Write-Host ('  ' + ((T 'langline') -f (Get-Culture).Name)) -ForegroundColor DarkCyan
    Write-Host ('  ' + $wi.Summary) -ForegroundColor DarkGray
    Write-Section (T 'sec_win')
    Write-Host (T 'm1') -ForegroundColor White
    Write-Host (T 'm2') -ForegroundColor DarkGray
    Write-Host (T 'm3') -ForegroundColor DarkGray
    Write-Section (T 'sec_off')
    Write-Host (T 'm4') -ForegroundColor White
    Write-Host (T 'm5') -ForegroundColor Cyan
    Write-Host (T 'm6') -ForegroundColor Cyan
    Write-Section (T 'sec_all')
    Write-Host (T 'm7') -ForegroundColor White
    Write-Host (T 'm8') -ForegroundColor Magenta
    Write-Host (T 'm9') -ForegroundColor Magenta
    Write-Section (T 'sec_info')
    Write-Host (T 'mA') -ForegroundColor White
    Write-Host ''
    Write-Host (T 'mH') -ForegroundColor DarkGray
    Write-Host (T 'mx') -ForegroundColor DarkGray
    Write-Host (T 'ml') -ForegroundColor DarkGray
    Write-Host ('  ' + (T 'esc_hint')) -ForegroundColor DarkGray
}

# Main loop. Iterative, not recursive: every action returns here, and Esc inside a submenu just
# lands back here - no need to restart the app after a wrong pick.
function Invoke-Main {
    while ($true) {
        Show-MainMenu
        $r = Read-Input (T 'choice')
        if ($r.Back) { continue }
        $k = $r.Text.Trim().ToLowerInvariant()
        if (-not $k) { continue }
        if ($k -eq 'w') { return }
        Invoke-Action $k
        if (@('9','a','h') -notcontains $k) {
            Write-Host ''
            [void](Read-Input (T 'back'))
        }
    }
}



if (-not (Test-FullLanguageMode)) { return }



if ($args -contains '-SelfTest') {
    $fail = 0
    function Assert([string]$Name, [bool]$Ok) {
        if ($Ok) { Write-Host "  OK   $Name" -ForegroundColor Green }
        else { Write-Host "  FAIL $Name" -ForegroundColor Red; $script:fail++ }
    }
    Write-Host 'SelfTest' -ForegroundColor Cyan
    Write-Host ("  culture: " + (Get-Culture).Name)

    
    Assert 'Test-Match upper case (HWID)' (Test-Match 'HWID is here' 'hwid')
    Assert 'Test-Match upper case (OHOOK)' (Test-Match 'OHOOK' 'ohook')
    Assert 'Test-Match upper case (LICENSED)' (Test-Match 'Windows LICENSED' 'licens')
    Assert 'Test-Match negative' (-not (Test-Match 'nothing here' 'hwid'))

    
    
    $engineOut = @(
        'Ohook Activation v3.0',
        'TSforge Activation',
        'HWID activation complete',
        'Microsoft Activation Scripts (MAS)',
        'https://massgrave.dev/troubleshoot',
        'Switching To KMS4k activation...        <- should drop',
        'KMS4k method is not supported on this build',
        'In builds 26100 and later, the script will auto select StaticCID',
        'Fork activation method                 [ZeroCID]',
        'Choose activation method:',
        '  [2] StaticCID',
        'C:\Users\x\AppData\Local\Temp\AW_11111111-2222-3333-4444-555555555555.cmd>echo installing',
        'Windows 11 Pro is LICENSED.',
        'Office LTSC Plus 2024 is activated.',
        'ERROR: something went wrong'
    )
    $kept = @(Get-ResultLines ($engineOut -join "`r`n"))
    $joined = $kept -join "`n"
    $leaky = @($kept | Where-Object { Test-Match $_ 'ohook|tsforge|hwid|kms|staticcid|zerocid|microsoft activation|massgrave|activated\.win|github\.com|^\s*\[[0-9]\]|AW_11111' })
    Assert ("filter: no method names leak" + $(if ($leaky) { " -> " + ($leaky -join ' | ') } else { '' })) ($leaky.Count -eq 0)
    Assert 'filter: LICENSED survives'      (Test-Match $joined 'is LICENSED')
    Assert 'filter: activated survives'     (Test-Match $joined 'is activated')
    Assert 'filter: ERROR survives'         (Test-Match $joined 'ERROR: something went wrong')
    Assert 'filter: cmd echo dropped'   (-not (Test-Match $joined 'echo installing'))
    Assert 'filter: empty stays empty'       (-not [bool]@(Get-ResultLines '').Count)

    
    Assert 'hidden timeout defined' ([int]$HiddenTimeoutSec -gt 60)

    
    $cfgTxt = Get-OfficeConfigXml -Lang 'tr-tr' -ExcludeIds @('Access', 'Outlook')
    $cfgXml = $null
    try { $cfgXml = [xml]$cfgTxt } catch { }
    Assert 'office cfg: valid XML' ($null -ne $cfgXml)
    if ($cfgXml) {
        $ns = $cfgXml.Configuration.Add.Product
        Assert 'office cfg: language correct'    ($ns.Language.ID -eq 'tr-tr')
        Assert ('office cfg: product + channel') ($ns.ID -eq 'ProPlus2024Volume' -and $cfgXml.Configuration.Add.Channel -eq 'PerpetualVL2024')
        Assert ('office cfg: ExcludeApp count 2') (@($ns.ExcludeApp).Count -eq 2)
        Assert 'office cfg: auto-activate' (@($cfgXml.Configuration.Property | Where-Object { $_.Name -eq 'AUTOACTIVATE' }).Value -eq '1')
    }
    Assert 'office cfg: empty when nothing is excluded' (-not (Test-Match (Get-OfficeConfigXml '' @()) 'ExcludeApp'))

    
    $dsel = New-AppSelection
    Assert 'office selection: empty by default' (@($dsel.Values | Where-Object { $_ }).Count -eq 0)
    Assert 'office selection: 9 components'     ($dsel.Keys.Count -eq 9)
    Assert 'office mode text describes W/E/P'  (Test-Match (T 'mode_h') 'Word.*Excel.*PowerPoint')
    Assert 'office hint lists the recommended set'  (Test-Match (T 'app_hint') 'Word.*Excel.*PowerPoint')

    
    Assert 'busy bar: returns when done'        (Wait-WithBar { $true } 5)
    $swBar = [Diagnostics.Stopwatch]::StartNew()
    $barRes = Wait-WithBar { $false } 1
    $swBar.Stop()
    Assert 'busy bar: stops on timeout' ((-not $barRes) -and $swBar.Elapsed.TotalSeconds -lt 6)

    # --- v2.3.0 additions ---

    # Every key used through T '...' must exist in the dictionary; a typo would otherwise print an
    # empty line at runtime and nobody would notice until a user does.
    $srcText = ''
    try { $srcText = [IO.File]::ReadAllText($PSCommandPath) } catch { }
    if ($srcText) {
        $usedKeys = @([regex]::Matches($srcText, "\bT '([a-z0-9_]+)'") | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $missing = @($usedKeys | Where-Object { -not $Strings['tr'].ContainsKey($_) })
        Assert ("i18n: no missing key" + $(if ($missing) { " -> " + ($missing -join ',') } else { '' })) ($missing.Count -eq 0)

        # Only switches the engine actually documents may be passed to it; a typo would silently
        # activate nothing while reporting success.
        $knownSwitch = @('/HWID', '/Ohook', '/Z-ESU', '/Z-Windows', '/Z-Office', '/Z-WindowsESUOffice', '/Z-ProjectVisio', '/Z-WinHost', '/Z-OffHost', '/Z-APPX', '/K-WindowsOffice', '/S')
        $usedSwitch = @([regex]::Matches($srcText, "Invoke-MasAio\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $badSwitch = @($usedSwitch | Where-Object { $knownSwitch -notcontains $_ })
        Assert ("engine: no unknown switch" + $(if ($badSwitch) { " -> " + ($badSwitch -join ',') } else { '' })) ($badSwitch.Count -eq 0)
        Assert 'engine: at least one switch used' ($usedSwitch.Count -gt 0)
    }

    # Esc-as-back must work identically on the interactive path and the piped fallback.
    Assert 'back token (esc)'      (Test-BackToken 'esc')
    Assert 'back token (padded ESC)' (Test-BackToken '  ESC ')
    Assert 'back token negative'    (-not (Test-BackToken 'escalate'))
    Assert 'back token empty'        (-not (Test-BackToken ''))

    # ODT acceptance must reject a small or non-PE file; a size-only check would accept both.
    $f1 = [IO.Path]::GetTempFileName()
    [IO.File]::WriteAllText($f1, 'not an executable')
    Assert 'odt: small file rejected' (-not (Test-OdtFile $f1))
    $f2 = [IO.Path]::GetTempFileName()
    $fs = [IO.File]::Create($f2)
    try { $fs.Write([byte[]](0x4D, 0x5A), 0, 2); $fs.SetLength(3MB) } finally { $fs.Close() }
    Assert 'odt: PE + size accepted' (Test-OdtFile $f2)
    $f3 = [IO.Path]::GetTempFileName()
    $fs = [IO.File]::Create($f3)
    try { $fs.Write([byte[]](0x3C, 0x68), 0, 2); $fs.SetLength(3MB) } finally { $fs.Close() }
    Assert 'odt: non-PE rejected' (-not (Test-OdtFile $f3))
    Remove-Item $f1, $f2, $f3 -Force -ErrorAction SilentlyContinue
    Assert 'odt: missing file rejected' (-not (Test-OdtFile (Join-Path $env:TEMP 'AW_no_such_file.exe')))

    # The fallback chain must point at Microsoft first and our mirror last, and the hash must be real.
    Assert 'odt: pinned link is microsoft' (Test-Match $OdtPinnedUrl '^https://download\.microsoft\.com/')
    Assert 'odt: backup link is github'    (Test-Match $OdtMirrorUrl '^https://github\.com/.+/releases/download/odt-mirror/')
    Assert 'odt: backup hash is 64 hex'        (Test-Match $OdtPinnedHash '^[0-9A-F]{64}$')

    # Registry-only OS info: must work without WMI and must not call it.
    $wi = Get-WinInfo
    Assert 'win info: family is Windows 10/11' (Test-Match $wi.Name '^Windows 1[01]$')
    Assert 'win info: build read'     ($wi.Build -gt 10000)
    Assert 'win info: summary filled'          ($wi.Summary.Length -gt 10)
    Assert 'win info: cached'         ([object]::ReferenceEquals($wi, (Get-WinInfo)))

    # Edition list must be deduplicated (the raw WMI query returned 61 rows with repeats here).
    $eds = @(Get-EditionList)
    Assert 'edition: deduplicated' (@($eds | Sort-Object -Unique).Count -eq $eds.Count)

    Assert "version format" (Test-Match $Version '^\d+\.\d+\.\d+$')
    Assert 'author defined' ($Author.Length -gt 3)
    Assert 'short url is https' (Test-Match $ShortUrl '^https://')

    
    $tk = @($Strings['tr'].Keys); $ek = @($Strings['en'].Keys)
    $diff = @(Compare-Object $tk $ek)
    Assert 'TR/EN key parity' ($diff.Count -eq 0)

    
    $leak = @()
    foreach ($lg in @('tr', 'en')) {
        foreach ($k in $Strings[$lg].Keys) {
            if (Test-Match ([string]$Strings[$lg][$k]) 'ohook|tsforge|kms38|kms4k|staticcid|zerocid|hwid|massgrave') { $leak += "$lg.$k" }
        }
    }
    Assert ("arayuz metinleri temiz" + $(if ($leak) { " -> " + ($leak -join ',') } else { '' })) ($leak.Count -eq 0)

    Write-Host $(if ($fail -eq 0) { 'SELFTEST OK' } else { "SELFTEST FAIL ($fail)" }) -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
    if ($fail -eq 0) { exit 0 } else { exit 1 }
}


$forward = @($args) | Where-Object { $_ -match '^/' }
if ($forward) { Invoke-MasAio ($forward -join ' '); return }

if ($args -contains '-Probe') {
    Write-Banner
    $eng = Resolve-Engine
    if ($eng) {
        Write-Host ("PROBE OK - motor alindi ve SHA256 dogrulandi. kaynak={0} commit={1} boyut={2} bayt" -f $eng.Source, $eng.Commit, $eng.Text.Length) -ForegroundColor Green
    } else {
        Write-Host 'PROBE FAIL - engine' -ForegroundColor Red
    }
    # ODT mirror reachability (HEAD only, nothing is downloaded).
    try {
        $req = [Net.HttpWebRequest]::Create($OdtMirrorUrl)
        $req.Method = 'HEAD'
        $req.UserAgent = 'ActivaterWindows'
        $req.Timeout = 30000
        $resp = $req.GetResponse()
        $len = $resp.ContentLength
        $resp.Close()
        Write-Host ("ODT MIRROR OK - {0} bayt" -f $len) -ForegroundColor Green
    } catch {
        Write-Host ("ODT MIRROR FAIL - {0}" -f $_.Exception.Message) -ForegroundColor Red
    }
    if ($eng) { exit 0 } else { exit 1 }
}

# Entry point. Closed input (a piped run that ran out of lines, Ctrl+C) ends the session cleanly
# instead of turning into an error dump, and cleans up the temp folder if one was in use.
try {
    Invoke-Main
} catch [System.OperationCanceledException] {
    Write-Host ''
} finally {
    if ($script:WorkDir -and (Test-Path $script:WorkDir)) {
        Remove-Item $script:WorkDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
