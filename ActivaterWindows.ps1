# ActivaterWindows
# PowerShell + Windows 10/11. Aktivasyon motoru: masaustu kaynagindan otomatik izlenir (upstream commit+hash),
# olmazsa gomulu pin, o da olmazsa son dogrulanmis yerel kopya. Her durumda SHA256 ile dogrulanir.
# Metot adlari arayuzde gosterilmez; etkinlestirme arka planda (penceresiz) calisir.
# Tek satir: irm https://raw.githubusercontent.com/tarikelertarnak/ActivaterWindows/main/ActivaterWindows.ps1 | iex
# Yerel: ActivaterWindows.cmd (cift tik) veya powershell -ExecutionPolicy Bypass -File ActivaterWindows.ps1
# Parametre: -Probe (indir+dogrula testi, calistirmadan cikar) | diger "/"-li argumanlar dogrudan MAS AIO'ya iletilir.

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
$RepoUrl  = 'https://github.com/tarikelertarnak/ActivaterWindows'
$SiteUrl  = 'https://tarikelertarnak.pages.dev/'
$Pinned   = '05c4f881efec946c0040cdd552d1afa9a519704b' # MAS commit
$AioHash  = '850F979665FB93999ACAE93F4790C1FF8ED2041532060B7966A121C2D29A0BFA'
$ChkHash  = '6DF2937385386D0C9682D351797E674583A410F4792BB7BE45302F40AB90287E'
$CacheDir = Join-Path $env:LOCALAPPDATA 'ActivaterWindows'
$AioCache = Join-Path $CacheDir 'engine.cmd'
$AioCacheHash = Join-Path $CacheDir 'engine.sha256'
$ChkCache = Join-Path $CacheDir 'check.cmd'
$ChkCacheHash = Join-Path $CacheDir 'check.sha256'
# Gizli calisan motorun ust siniri (saniye). Sonda tus bekleyen akislar bu sure sonunda kapatilir.
$HiddenTimeoutSec = 900
$script:UiLang = 'auto' # 'auto' | 'tr' | 'en'

# Office C2R desteklenen diller (alfabetik, numarali liste secimi icin)
$OfficeLangs = 'af-za am-et ar-sa as-in az-latn-az be-by bg-bg bn-bd bn-in bs-latn-ba ca-es cs-cz cy-gb da-dk de-de el-gr en-gb en-us es-es es-mx et-ee eu-es fa-ir fi-fi fil-ph fr-ca fr-fr gl-es gu-in he-il hi-in hr-hr hu-hu hy-am id-id is-is it-it ja-jp ka-ge kk-kz km-kh kn-in ko-kr ky-kg lo-la lt-lt lv-lv mk-mk ml-in mn-mn mr-in ms-my mt-mt nb-no ne-np nl-nl nn-no or-in pa-in pl-pl ps-af pt-br pt-pt ro-ro ru-ru si-lk sk-sk sl-si sq-al sr-latn-rs sv-se sw-ke ta-in te-in th-th tr-tr uk-ua ur-pk uz-latn-uz vi-vn zh-cn zh-tw zu-za' -split ' '

# Office bilesenleri (ozel secim icin) — ODT ExcludeApp ID'leri
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

# --- arayuz dili sozlugu ---
$Strings = @{
    tr = @{
        langline = 'Dil / Language :  [A] Otomatik (su an: {0})   [T] Turkce   [E] English'
        m1 = '  [1] Windows 10/11 etkinlestir - kalici'
        m2 = '  [2] Office etkinlestir        - kalici'
        m3 = '  [3] Windows + ESU + Office    - kalici (tek gecis)'
        m4 = '  [4] Windows 10 ESU            - kalici'
        m5 = '  [5] Aktivasyon durumu - kontrol et'
        m6 = '  [6] Office LTSC Pro Plus 2024 kur (dil + bilesen secimi)'
        m7 = '  [7] HEPSINI ETKINLESTIR - mevcut her seyi kalici etkinlestir (kurulum yok)'
        m8 = '  [8] HEPSINI YAP - eksikse Office kur + her seyi etkinlestir + tam MAS menusu'
        mx = '  [0] Cikis'
        ml = '  [L] Arayuz dilini degistir'
        choice = 'Secim'
        back = 'Menuye donmek icin Enter'
        sel12 = 'Secim (1/2/3)'
        mode_head = 'Kurulum modu:'
        langpick = 'Dil:  [1] Otomatik (su an: {0})  [2] Turkce  [3] English'
        olang = 'Office dili: kod yaz (ornek: tr-tr) veya asagidaki listeden numara sec [otomatik: {0}]'
        olang_confirm = '{0}) {1} — onayliyor musun? Enter = evet, 0 = geri'
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
    }
    en = @{
        langline = 'Dil / Language :  [A] Automatic (now: {0})   [T] Turkce   [E] English'
        m1 = '  [1] Activate Windows 10/11    - permanent'
        m2 = '  [2] Activate Office           - permanent'
        m3 = '  [3] Windows + ESU + Office    - permanent (single pass)'
        m4 = '  [4] Windows 10 ESU            - permanent'
        m5 = '  [5] Check activation status'
        m6 = '  [6] Install Office LTSC Pro Plus 2024 (language + component picker)'
        m7 = '  [7] ACTIVATE EVERYTHING - permanently activate everything present (no install)'
        m8 = '  [8] DO EVERYTHING - install Office if missing + activate all + full MAS menu'
        mx = '  [0] Exit'
        ml = '  [L] Change interface language'
        choice = 'Choice'
        back = 'Press Enter to return to menu'
        sel12 = 'Choice (1/2/3)'
        mode_head = 'Install mode:'
        langpick = 'Language:  [1] Automatic (now: {0})  [2] Turkce  [3] English'
        olang = 'Office language: type a code (e.g. tr-tr) or pick a number from the list below [auto: {0}]'
        olang_confirm = '{0}) {1} — confirm? Enter = yes, 0 = back'
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

# Turkce kulturde 'I' ile 'i' eslesmez: 'I'.ToLower() = 'ı' (noktasiz i). Bu yuzden
# varsayilan -match/-like karsilastirmalari buyuk harfli metinde sessizce kacar (orn. "HWID").
# Metin karsilastirmalarinda daima bu fonksiyonu kullan.
function Test-Match([string]$Text, [string]$Pattern) {
    [regex]::IsMatch($Text, $Pattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant)
}

# Evet/hayir sorusu. TR: E = evet, EN: Y = yes. Enter = evet. Diger her giris = hayir.
function Confirm([string]$Prompt) {
    $ans = Read-Host $Prompt
    if (-not $ans) { return $true }
    if ((Get-UiLang) -eq 'tr') { return (Test-Match $ans '^\s*e(vet)?\s*$') }
    Test-Match $ans '^\s*y(es)?\s*$'
}

function Write-Banner {
    Write-Host ''
    Write-Host '=============================================' -ForegroundColor DarkCyan
    Write-Host '  ActivaterWindows' -ForegroundColor Cyan
    Write-Host "  $SiteUrl" -ForegroundColor DarkGray
    Write-Host "  $RepoUrl" -ForegroundColor DarkGray
    Write-Host '=============================================' -ForegroundColor DarkCyan
}

function Test-FullLanguageMode {
    if ($ExecutionContext.SessionState.LanguageMode.value__ -ne 0) {
        Write-Host "PowerShell Full Language Mode'da calismiyor." -ForegroundColor Red
        Write-Host 'Cozum: https://massgrave.dev/fix_powershell' -ForegroundColor White -BackgroundColor Blue
        return $false
    }
    try { [void][System.AppDomain]::CurrentDomain.GetAssemblies(); [void][System.Math]::Sqrt(144) }
    catch {
        Write-Host "Hata: $($_.Exception.Message)" -ForegroundColor Red
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

# --- indirme: gercek yuzde barı + saniye sayaci (chunked, senkron garantili) ---
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

# Yuzdesi bilinmeyen ama uzun suren adimlar (ODT ayiklama, Office kurulumu) icin kayan cubuk.
# Indirme gibi toplami bilinen isler Write-DlBar'in gercek yuzdesini kullanir; bu onun tamamlayicisi.
function Write-BusyBar([int]$Tick, [int]$Width = 30, $Sw) {
    $pos = $Tick % ($Width * 2)
    if ($pos -ge $Width) { $pos = ($Width * 2) - $pos - 1 }
    $track = ('-' * $pos) + '#' + ('-' * ($Width - $pos - 1))
    $el = if ($Sw) { [int]$Sw.Elapsed.TotalSeconds } else { 0 }
    Write-Host ("`r  [{0}] {1,4}s" -f $track, $el) -NoNewline
}

# Kosul saglanana kadar cubugu canli tutar. $MaxSec asilirsa durur ve $false doner (sonsuz asilma yok).
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
    Write-Host ("`r  [{0}] {1,4}s" -f ('#' * 30), [int]$sw.Elapsed.TotalSeconds)
    $fin
}

function Invoke-Download([string]$Url, [string]$OutFile) {
    # chunked stream okuma — WebClient event'leri senkron cagrida calismaz, bu calisir
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

# Aktivasyon motorunun kaynagi. Sirasi:
#   1) Upstream'in kendi bootstrap'i (get.activated.win) — commit + hash oradan okunur, boylece
#      uygulama guncellenmese bile motor her zaman guncel kalir.
#   2) Gomulu pin (commit + hash) — upstream'in formati degisirse bile calisir.
#   3) Yerel onbellek — ag/upstream tamamen bozulsa bile son dogrulanmis kopya kullanilir.
function Resolve-Engine {
    # 1) upstream bootstrap -> guncel commit + hash
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

    # 2) gomulu pin
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

    # 3) yerel onbellek (son dogrulanmis kopya) — icerik, kaydedilen hash ile yeniden dogrulanir
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

# Yerel kurcalama/bozulmayi yakalar: icerik, kaydedildigi andaki hash ile uyusmazsa onbellek reddedilir.
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

# Aktivasyon ciktisindan kullaniciya gosterilecek satirlar.
# Metot/marka adlari (motorun kendi adi, kaynak adresleri, lisans metinleri) filtrelenir.
# Filtre (tek kaynak): kullaniciya gosterilecek satirlari secer. Sonuc/uyari satirlari kalir;
# metot/marka adlari, yontem secim metinleri, kaynak adresleri, lisans metinleri ve cmd komut
# yankilari dusurulur. `-SelfTest` bu fonksiyonu aynen test eder.
function Get-ResultLines([string]$Text) {
    if (-not $Text) { return @() }
    # ANSI renk kodlarini temizle
    $t = [regex]::Replace($Text, '\x1B\[[0-9;?]*[A-Za-z]', '')
    $t = [regex]::Replace($t, '\x1B[()][0-9A-B]', '')
    $skip = '(?i)ohook|tsforge|kms38|kms4k|kms4|staticcid|zerocid|scid|zcid|\bkms\b|hwid|massgrave|microsoft activation scripts|\bmas\b|activated\.win|github\.com|licence:|license:|gpl|_aio|plugin|switching to|method is not supported|choose a|select a|^\s*\[[0-9]\]'
    $keep = '(?i)licens|activ|success|fail|error|tamamlan|etkin|windows|office|expire|product key'
    @($t -split "`r?`n" |
        ForEach-Object { $_.Trim() } |
        # cmd'nin ekrana bastigi komut satirlari (AW_<guid>.cmd> echo ...) kullaniciya cikmasin
        Where-Object { $_ -and -not (Test-Match $_ '^[A-Za-z]:\\[^\r\n]*AW_[0-9a-f\-]{36}\.cmd>') -and -not (Test-Match $_ '^\s*>\s') } |
        Where-Object { $_ -and -not (Test-Match $_ $skip) -and (Test-Match $_ $keep) })
}

function Show-ActivationResult([string]$LogPath) {
    if (-not $LogPath -or -not (Test-Path $LogPath)) { return $false }
    # Surec zorla sonlandirildiginda gunlugun yazma tutucusu hala kapanmis olmayabilir;
    # o zaman ReadAllText IOException firlatir (içerik aslinda orada). Paylasimli ac + kisa tekrar.
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
    # ilk satirdaki @::: guidi cmd'nin -el ile kendi kendine yeniden calistirmasinda ilk satiri etkisizler
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
            # Motor bazi akislarda sonda tus bekler (or. durum kontrolu `choice /c 0 /n`).
            # Gizli pencerede tus bekleyen surec sonsuza kadar asili kalir; bu yuzden sinirli bekle:
            # cikti zaten gecici gunluge yazilmis olur, sonra surec sonlandirilir.
            # Motorun kendi penceresi gizli oldugu icin tek geri bildirim bu cubuktur.
            $p = Start-Process @sp
            if (-not (Wait-WithBar { $p.WaitForExit(200) } $HiddenTimeoutSec)) {
                # Stop-Process yalnizca cmd.exe'yi oldurur; motor (ve onun timeout/choice cocuklari)
                # yasamaya devam eder -> gizli yetim surecler. /T ile tum agac kapatilir.
                & "$env:SystemRoot\system32\taskkill.exe" /PID $p.Id /T /F 2>&1 | Out-Null
                Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
                Write-Host ((T 'act_timeout') -f $HiddenTimeoutSec) -ForegroundColor DarkYellow
            }
            if (-not (Show-ActivationResult $log)) {
                # filtre hicbir satir birakmadi: "hic cikti yok" ile "basarisiz" ayni gorunmesin
                Write-Host (T 'act_nolog') -ForegroundColor DarkGray
            }
        } else {
            $p = Start-Process @sp
            Write-Host ((T 'done') -f $Label, $p.ExitCode) -ForegroundColor Green
        }
    } catch {
        Write-Host ((T 'uac_cancel2') -f $Label) -ForegroundColor Yellow
    } finally {
        # Surec yeni oldugunde handle'lar hemen birakilmaz; tek denemeli silme sessizce basarisiz olur
        # ve motorun tam kopyasini iceren .cmd gecici klasorde kalir. Kisa yeniden deneme.
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
        # metot cagrilari: pencere acilmaz, sonuc kendi filtremizden gecer
        Invoke-MasCmd -Response $eng.Text -Arguments $Arguments -Label $Arguments.Trim('/')
    } else {
        # tam menu kullanicinin etkilesimi icin gorunur olmak zorunda
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

function Default-OfficeLang {
    $c = ((Get-Culture).Name).ToLowerInvariant()
    if ($script:UiLang -eq 'tr') { 'tr-tr' }
    elseif ($script:UiLang -eq 'en') { 'en-us' }
    else { $c }
}

# Hicbir bilesen onceden secili gelmez; kullanici kendi secer.
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
        $in = Read-Host (T 'app_sel')
        if (-not $in) { break }
        $low = $in.Trim().ToLowerInvariant()
        if ($low -eq 'a') { foreach ($a in $OfficeApps) { $sel[$a.id] = $true }; continue }
        if ($low -eq 'n') { foreach ($a in $OfficeApps) { $sel[$a.id] = $false }; continue }
        if ($low -eq 'r') { foreach ($a in $OfficeApps) { $sel[$a.id] = ($OfficeAppsRecommended -contains $a.id) }; continue }
        foreach ($part in ($in -split ',')) {
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
        $in = Read-Host ((T 'olang') -f $def)
        if (-not $in) { return $def }
        if ($in -match '^\d+$') {
            $n = [int]$in
            if ($n -ge 1 -and $n -le $OfficeLangs.Count) {
                $code = $OfficeLangs[$n - 1]
                $ans = Read-Host ((T 'olang_confirm') -f $n, $code)
                if ($ans -eq '0') { continue }
                return $code
            }
            Write-Host (T 'invalid_lang') -ForegroundColor Red
            continue
        }
        if (Test-Match $in '^[a-z]{2}(-[a-z0-9]{2,8})?$') { return $in.ToLowerInvariant() }
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
    # Dil once, mod sonra: dil yapilandirmanin parcasi, kurulum kapsamindan bagimsiz secilir.
    $lang = Resolve-OfficeLang
    Write-Host ''
    Write-Host ('  ' + (T 'mode_head')) -ForegroundColor Yellow
    Write-Host (T 'mode_h') -ForegroundColor Gray
    Write-Host (T 'mode_t') -ForegroundColor Gray
    Write-Host (T 'mode_c') -ForegroundColor Gray
    Write-Host ''
    $mode = Read-Host (T 'sel12')

    $work = Join-Path ([IO.Path]::GetTempPath()) "AW_ODT_$(Get-Random)"
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    $odt = Join-Path $work 'odt.exe'

    # 1) repo icindeki yerel kopya (script dosyasiyla birlikte calisirken)
    $local = $null
    if ($PSScriptRoot) {
        $local = Get-ChildItem -Path (Join-Path $PSScriptRoot 'resources') -Filter 'officedeploymenttool*.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    if ($local) {
        Write-Host ((T 'odt_local') -f $local.Name) -ForegroundColor DarkGray
        Copy-Item $local.FullName $odt
    } else {
        # 2) MS indirme sayfasindan dogrudan .exe linki (aka.ms/odt artik sadece dokumantasyona gider)
        Write-Host (T 'odt_dl') -ForegroundColor DarkGray
        try {
            $conf = Invoke-WebRequest -Uri 'https://www.microsoft.com/en-us/download/details.aspx?id=49117' -UseBasicParsing
            # CultureInvariant sart: tr-TR'de 'I' != 'i', yani düz 'IgnoreCase' buyuk harfli linki kacirir.
            $link = [regex]::Match($conf.Content, 'https://download\.microsoft\.com/[^"''\s<>]*officedeploymenttool[^"''\s<>]*\.exe', [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant).Value
            if (-not $link) { throw 'download link not found on page' }
            Save-File -Url $link -OutFile $odt
        } catch {
            Write-Host ((T 'odt_fail') -f $_.Exception.Message) -ForegroundColor Red
            Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
            return
        }
    }
    if ((Get-Item $odt).Length -lt 2MB) { Write-Host (T 'odt_small') -ForegroundColor Red; Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue; return }

    # bilesen secimi → ExcludeApp listesi
    if ($mode -eq '3') {
        $chosen = @(Select-OfficeApps)
        $chosenIds = $chosen | ForEach-Object { $_.id }
        $excludeIds = $OfficeApps | ForEach-Object { $_.id } | Where-Object { $chosenIds -notcontains $_ }
    } elseif ($mode -eq '2') {
        $excludeIds = @()
    } else {
        $excludeIds = @('Access', 'Lync', 'OneDrive', 'OneNote', 'Outlook', 'Publisher')
    }

    Write-Host (T 'extract') -ForegroundColor Cyan
    $px = Start-Process -FilePath $odt -ArgumentList "/extract:`"$work`" /quiet" -PassThru -WindowStyle Hidden
    if (-not (Wait-WithBar { $px.HasExited } 300)) { Write-Host (T 'step_timeout') -ForegroundColor Red }
    $setup = Join-Path $work 'setup.exe'
    if (-not (Test-Path $setup)) { Write-Host (T 'extract_fail') -ForegroundColor Red; Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue; return }

    $config = Get-OfficeConfigXml -Lang $lang -ExcludeIds $excludeIds
    $configPath = Join-Path $work 'Configuration.xml'
    [IO.File]::WriteAllText($configPath, $config)

    $isAdmin = [bool]([Security.Principal.WindowsIdentity]::GetCurrent().Groups -match 'S-1-5-32-544')
    $rcFile = Join-Path $work 'rc.txt'
    $rc = $null
    try {
        if ($isAdmin) {
            Write-Host (T 'inst_start') -ForegroundColor Cyan
            $p = Start-Process -FilePath $setup -ArgumentList "/configure `"$configPath`"" -PassThru -WindowStyle Hidden
            if (-not (Wait-WithBar { $p.HasExited } 3600)) { Write-Host (T 'step_timeout') -ForegroundColor Red }
            $rc = $p.ExitCode
        } else {
            # Yukseletilen cocuk surec kendi cikis kodunu yazar. Aksi halde kurulum basarisiz olsa bile
            # 'basarili' gorunurdu (eski hata). Cubuk, cikis kodu dosyasi olusana kadar canli kalir.
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

function Test-Win10 {
    try { [int](Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).BuildNumber -lt 22000 }
    catch { $false }
}

# [7] HEPSINI ETKINLESTIR — sadece etkinlestirme, kurulum yok.
# Windows: kalici dijital lisans (guncellemelerden etkilenmez). Office: kalici. Win10'da ESU da.
# Argumanlar tek geciste ve sira bagimsiz gonderilir; motor hangi sirada olursa olsun hepsini isler.
function Invoke-AllActivate {
    Write-Host ''
    Write-Host (T 'act_start') -ForegroundColor Magenta
    $methods = @('/HWID')
    if (Test-OfficeC2R) { $methods += '/Ohook' } else { Write-Host (T 'act_nooffice') -ForegroundColor DarkYellow }
    if (Test-Win10) { $methods += '/Z-ESU' }
    Invoke-MasAio ($methods -join ' ')
    Invoke-MasCheck
}

# [8] HEPSINI YAP — eksikse Office kur, sonra [7] akisini calistir, en son tam MAS menusunu teklif et.
function Invoke-AllInOne {
    Write-Host ''
    Write-Host (T 'all_start') -ForegroundColor Magenta
    if (-not (Test-OfficeC2R)) {
        if (Confirm (T 'all_nooffice')) { Write-Host (T 'all_installing') -ForegroundColor DarkGray; Install-OfficeLTSC -SkipActivate }
    }
    Invoke-AllActivate
    if (Confirm (T 'all_masmenu')) { Invoke-MasAio }
    Write-Host (T 'all_done') -ForegroundColor Green
}

function Set-UiLang {
    Write-Host ''
    Write-Host ((T 'langpick') -f ((Get-Culture).Name)) -ForegroundColor Cyan
    $c = Read-Host '>>'
    switch ($c) {
        '1' { $script:UiLang = 'auto' }
        '2' { $script:UiLang = 'tr' }
        '3' { $script:UiLang = 'en' }
    }
}

function Show-Menu {
    Clear-Host
    Write-Banner
    Write-Host ''
    Write-Host ((T 'langline') -f ((Get-Culture).Name)) -ForegroundColor DarkCyan
    Write-Host ''
    Write-Host (T 'm1') -ForegroundColor White
    Write-Host (T 'm2') -ForegroundColor White
    Write-Host (T 'm3') -ForegroundColor White
    Write-Host (T 'm4') -ForegroundColor White
    Write-Host (T 'm5') -ForegroundColor White
    Write-Host (T 'm6') -ForegroundColor Cyan
    Write-Host (T 'm7') -ForegroundColor Magenta
    Write-Host (T 'm8') -ForegroundColor Magenta
    Write-Host (T 'mx') -ForegroundColor DarkGray
    Write-Host (T 'ml') -ForegroundColor DarkGray
    Write-Host ''
    $choice = Read-Host (T 'choice')
    switch ($choice) {
        '1' { Invoke-MasAio '/HWID' }
        '2' { Invoke-MasAio '/Ohook' }
        '3' { Invoke-MasAio '/Z-WindowsESUOffice' }
        '4' { Invoke-MasAio '/Z-ESU' }
        '5' { Invoke-MasCheck }
        '6' { Install-OfficeLTSC }
        '7' { Invoke-AllActivate }
        '8' { Invoke-AllInOne }
        '0' { return $true }
        'l' { Set-UiLang }
        'L' { Set-UiLang }
        'a' { $script:UiLang = 'auto' }
        'A' { $script:UiLang = 'auto' }
        't' { $script:UiLang = 'tr' }
        'T' { $script:UiLang = 'tr' }
        'e' { $script:UiLang = 'en' }
        'E' { $script:UiLang = 'en' }
        default { }
    }
    if ($choice -ne '0') { Write-Host ''; Read-Host (T 'back') | Out-Null; Show-Menu }
}

# --- giris ---
if (-not (Test-FullLanguageMode)) { return }

# Yerlesik kontroller: ag/kurulum/yetki gerektirmez. Ince davranislari (ozellikle
# Turkce kulturde buyuk harfli metin eslesmesini) her surumde dogrular.
if ($args -contains '-SelfTest') {
    $fail = 0
    function Assert([string]$Name, [bool]$Ok) {
        if ($Ok) { Write-Host "  OK   $Name" -ForegroundColor Green }
        else { Write-Host "  FAIL $Name" -ForegroundColor Red; $script:fail++ }
    }
    Write-Host 'SelfTest' -ForegroundColor Cyan
    Write-Host ("  culture: " + (Get-Culture).Name)

    # Turkce 'I' != 'i' tuzagi: bu satir duz -match ile gecmez, Test-Match ile gecmeli
    Assert 'Test-Match buyuk harf (HWID)' (Test-Match 'HWID is here' 'hwid')
    Assert 'Test-Match buyuk harf (OHOOK)' (Test-Match 'OHOOK' 'ohook')
    Assert 'Test-Match buyuk harf (LICENSED)' (Test-Match 'Windows LICENSED' 'licens')
    Assert 'Test-Match olumsuz' (-not (Test-Match 'nothing here' 'hwid'))

    # cikti filtresi — GERCEK fonksiyon (Get-ResultLines) test edilir, kopya degil:
    # metot adlari sizmamali, sonuc/uyari satirlari kalmali
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
    Assert ("filtre: metot/yontem adlari duser" + $(if ($leaky) { " -> " + ($leaky -join ' | ') } else { '' })) ($leaky.Count -eq 0)
    Assert 'filtre: LICENSED kalir'      (Test-Match $joined 'is LICENSED')
    Assert 'filtre: activated kalir'     (Test-Match $joined 'is activated')
    Assert 'filtre: ERROR kalir'         (Test-Match $joined 'ERROR: something went wrong')
    Assert 'filtre: cmd yankisi duser'   (-not (Test-Match $joined 'echo installing'))
    Assert 'filtre: bom bos cikti'       (-not [bool]@(Get-ResultLines '').Count)

    # gizli calistirma asilmasin diye zaman asimi tanimli ve makul
    Assert 'gizli zaman asimi tanimli' ([int]$HiddenTimeoutSec -gt 60)

    # Office kurulum yapilandirmasi: ODT gecersiz XML'de kurulumu reddeder -> her surumde dogrula.
    $cfgTxt = Get-OfficeConfigXml -Lang 'tr-tr' -ExcludeIds @('Access', 'Outlook')
    $cfgXml = $null
    try { $cfgXml = [xml]$cfgTxt } catch { }
    Assert 'office cfg: gecerli XML' ($null -ne $cfgXml)
    if ($cfgXml) {
        $ns = $cfgXml.Configuration.Add.Product
        Assert 'office cfg: dil dogru'    ($ns.Language.ID -eq 'tr-tr')
        Assert ('office cfg: urun + kanal') ($ns.ID -eq 'ProPlus2024Volume' -and $cfgXml.Configuration.Add.Channel -eq 'PerpetualVL2024')
        Assert ('office cfg: ExcludeApp sayisi 2') (@($ns.ExcludeApp).Count -eq 2)
        Assert 'office cfg: otomatik etkinlestirme' (@($cfgXml.Configuration.Property | Where-Object { $_.Name -eq 'AUTOACTIVATE' }).Value -eq '1')
    }
    Assert 'office cfg: hic dislama yoksa bos' (-not (Test-Match (Get-OfficeConfigXml '' @()) 'ExcludeApp'))

    # Bilesen secimi: hicbiri onceden secili gelmemeli (regresyon: eskiden Word/Excel/PowerPoint secili geliyordu)
    $dsel = New-AppSelection
    Assert 'office secim: varsayilan bos' (@($dsel.Values | Where-Object { $_ }).Count -eq 0)
    Assert 'office secim: 9 bilesen'     ($dsel.Keys.Count -eq 9)
    Assert 'office mod metni W/E/P der'  (Test-Match (T 'mode_h') 'Word.*Excel.*PowerPoint')
    Assert 'office ipucu onerilen seti'  (Test-Match (T 'app_hint') 'Word.*Excel.*PowerPoint')

    # Yuzdesiz adimlar icin canli cubuk: biteni hemen gecmeli, bitmeyeni zaman asiminda kesmeli
    Assert 'busy bar: biteni bekler'        (Wait-WithBar { $true } 5)
    $swBar = [Diagnostics.Stopwatch]::StartNew()
    $barRes = Wait-WithBar { $false } 1
    $swBar.Stop()
    Assert 'busy bar: zaman asiminda durur' ((-not $barRes) -and $swBar.Elapsed.TotalSeconds -lt 6)

    # TR/EN anahtar paritesi
    $tk = @($Strings['tr'].Keys); $ek = @($Strings['en'].Keys)
    $diff = @(Compare-Object $tk $ek)
    Assert 'TR/EN anahtar paritesi' ($diff.Count -eq 0)

    # arayuz metinlerinde metot adi gecmemeli (her iki dil)
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

# "/"-li arguman varsa dogrudan AIO'ya aktar (orn: ActivaterWindows.ps1 /HWID)
$forward = @($args) | Where-Object { $_ -match '^/' }
if ($forward) { Invoke-MasAio ($forward -join ' '); return }

if ($args -contains '-Probe') {
    Write-Banner
    $eng = Resolve-Engine
    if ($eng) {
        Write-Host ("PROBE OK - motor alindi ve SHA256 dogrulandi. kaynak={0} commit={1} boyut={2} bayt" -f $eng.Source, $eng.Commit, $eng.Text.Length) -ForegroundColor Green
        exit 0
    }
    Write-Host 'PROBE FAIL' -ForegroundColor Red
    exit 1
}

while ($true) { if (Show-Menu) { break } }
