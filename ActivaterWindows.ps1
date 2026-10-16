# ActivaterWindows
# PowerShell + Windows 10/11. Aktivasyon: Microsoft Activation Scripts (pinned commit, SHA256 dogrulamali).
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
$script:UiLang = 'auto' # 'auto' | 'tr' | 'en'

# Office C2R desteklenen diller (alfabetik, numarali liste secimi icin)
$OfficeLangs = 'af-za am-et ar-sa as-in az-latn-az be-by bg-bg bn-bd bn-in bs-latn-ba ca-es cs-cz cy-gb da-dk de-de el-gr en-gb en-us es-es es-mx et-ee eu-es fa-ir fi-fi fil-ph fr-ca fr-fr gl-es gu-in he-il hi-in hr-hr hu-hu hy-am id-id is-is it-it ja-jp ka-ge kk-kz km-kh kn-in ko-kr ky-kg lo-la lt-lt lv-lv mk-mk ml-in mn-mn mr-in ms-my mt-mt nb-no ne-np nl-nl nn-no or-in pa-in pl-pl ps-af pt-br pt-pt ro-ro ru-ru si-lk sk-sk sl-si sq-al sr-latn-rs sv-se sw-ke ta-in te-in th-th tr-tr uk-ua ur-pk uz-latn-uz vi-vn zh-cn zh-tw zu-za' -split ' '

# --- arayuz dili sozlugu ---
$Strings = @{
    tr = @{
        langline = 'Dil / Language :  [A] Otomatik (su an: {0})   [T] Turkce   [E] English'
        m1 = '  [1] Windows 10/11 etkinlestir  - HWID    (kalici)'
        m2 = '  [2] Office etkinlestir         - Ohook   (kalici)'
        m3 = '  [3] Windows + ESU + Office     - TSforge (kalici)'
        m4 = '  [4] Windows 10 ESU             - TSforge (kalici)'
        m5 = '  [5] Aktivasyon durumu - kontrol et'
        m6 = '  [6] TAM MAS MENUSU - tum ozellikler (Troubleshoot, Surum degistir, $OEM$...)'
        m7 = '  [7] Office LTSC Pro Plus 2024 kur (+Ohook)'
        m8 = '  [8] HEPSINI YAP - Windows + Office + Office kurulumu, tek tik'
        mx = '  [0] Cikis'
        ml = '  [L] Arayuz dilini degistir'
        choice = 'Secim'
        back = 'Menuye donmek icin Enter'
        sel12 = 'Secim (1/2)'
        langpick = 'Dil:  [1] Otomatik (su an: {0})  [2] Turkce  [3] English'
        olang = 'Office dili: kod yaz (ornek: tr-tr) veya asagidaki listeden numara sec [otomatik: {0}]'
        olang_confirm = '{0}) {1} — onayliyor musun? Enter = evet, 0 = geri'
        lang_head = 'Office dilleri (alfabetik):'
        mode_h = '  [1] Onerilen (hafif): Word, Excel, PowerPoint, OneDrive disi bilesenler kurulmaz'
        mode_t = '  [2] Tam paket: tum bilesenler'
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
        uac_cancel = 'Kurulum iptal edildi (UAC reddedildi).'
        done = '{0} tamamlandi (cikis kodu: {1}).'
        uac_cancel2 = '{0} iptal edildi (UAC reddedildi).'
        mas_dl = 'MAS All-In-One indiriliyor'
        chk_dl = 'Aktivasyon durumu kontrolu indiriliyor'
        hash_fail = 'MAS butunluk kontrolu basarisiz (hash uyusmadi), iptal!'
        hash_fail2 = 'Kontrol scripti butunluk kontrolu basarisiz, iptal!'
        dl_fail = '{0} hicbir kaynaktan alinamadi. Antivirus/firewall baglantayi engelleyebilir.'
        tmp_fail = 'Gecici dosya olusturulamadi, iptal! ({0})'
        all_start = 'HEPSINI YAP: ofis kurulu degilse kurulacak, sonra Windows + ESU + Office TSforge ile kalici etkinlestirilecek, en son durum raporu...'
        all_nooffice = 'Office kurulu gorunmuyor. LTSC 2024 kurulsun mu? (E/h, bos = evet)'
        all_installing = 'Once Office kuruluyor...'
        ohook_ask = 'Office simdi Ohook ile etkinlestirilsin mi? (E/h)'
    }
    en = @{
        langline = 'Dil / Language :  [A] Automatic (now: {0})   [T] Turkce   [E] English'
        m1 = '  [1] Activate Windows 10/11  - HWID    (permanent)'
        m2 = '  [2] Activate Office         - Ohook   (permanent)'
        m3 = '  [3] Windows + ESU + Office  - TSforge (permanent)'
        m4 = '  [4] Windows 10 ESU          - TSforge (permanent)'
        m5 = '  [5] Check activation status'
        m6 = '  [6] FULL MAS MENU - all features (Troubleshoot, Change Edition, $OEM$...)'
        m7 = '  [7] Install Office LTSC Pro Plus 2024 (+Ohook)'
        m8 = '  [8] DO EVERYTHING - Windows + Office + Office install, one go'
        mx = '  [0] Exit'
        ml = '  [L] Change interface language'
        choice = 'Choice'
        back = 'Press Enter to return to menu'
        sel12 = 'Choice (1/2)'
        langpick = 'Language:  [1] Automatic (now: {0})  [2] Turkce  [3] English'
        olang = 'Office language: type a code (e.g. tr-tr) or pick a number from the list below [auto: {0}]'
        olang_confirm = '{0}) {1} — confirm? Enter = yes, 0 = back'
        lang_head = 'Office languages (alphabetical):'
        mode_h = '  [1] Recommended (light): all components except Word, Excel, PowerPoint, OneDrive'
        mode_t = '  [2] Full: all components'
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
        uac_cancel = 'Install cancelled (UAC declined).'
        done = '{0} finished (exit code: {1}).'
        uac_cancel2 = '{0} cancelled (UAC declined).'
        mas_dl = 'Downloading MAS All-In-One'
        chk_dl = 'Downloading activation status checker'
        hash_fail = 'MAS integrity check FAILED (hash mismatch), aborting!'
        hash_fail2 = 'Checker integrity check FAILED, aborting!'
        dl_fail = '{0} could not be retrieved from any source. Antivirus/firewall may be blocking.'
        tmp_fail = 'Failed to create temp file, aborting! ({0})'
        all_start = 'DO EVERYTHING: install Office if missing, then activate Windows + ESU + Office permanently via TSforge, then status report...'
        all_nooffice = 'Office does not seem installed. Install LTSC 2024? (Y/n, blank = yes)'
        all_installing = 'Installing Office first...'
        ohook_ask = 'Activate Office with Ohook now? (Y/n)'
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
        Where-Object { $_.displayName -notlike '*windows*' } | Select-Object -ExpandProperty displayName
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

function Get-String([string[]]$Urls, [string]$Label) {
    Write-Host $Label -ForegroundColor DarkGray
    $errors = @(); $response = $null
    foreach ($url in $Urls | Sort-Object { Get-Random }) {
        try {
            $response = Get-WebString $url
            if ($response) { break }
        } catch { $errors += $_ }
    }
    if (-not $response) {
        Warn-AV
        foreach ($e in $errors) { Write-Host "Hata: $($e.Exception.Message)" -ForegroundColor Red }
        Write-Host ((T 'dl_fail') -f $Label) -ForegroundColor Red
    }
    $response
}

function Test-Hash([string]$Text, [string]$Expected) {
    $stream = New-Object IO.MemoryStream
    $writer = New-Object IO.StreamWriter $stream
    $writer.Write($Text); $writer.Flush(); $stream.Position = 0
    ([BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash($stream)) -replace '-') -eq $Expected
}

function Invoke-MasCmd([string]$Response, [string]$Arguments, [string]$Label) {
    $rand = [Guid]::NewGuid().Guid
    $isAdmin = [bool]([Security.Principal.WindowsIdentity]::GetCurrent().Groups -match 'S-1-5-32-544')
    $file = if ($isAdmin) { "$env:SystemRoot\Temp\AW_$rand.cmd" } else { "$env:USERPROFILE\AppData\Local\Temp\AW_$rand.cmd" }
    # ilk satirdaki @::: guidi cmd'nin -el ile kendi kendine yeniden calistirmasinda ilk satiri etkisizler
    Set-Content -Path $file -Value "@::: $rand`r`n$Response"
    if (-not (Test-Path $file)) { Warn-AV; Write-Host ((T 'tmp_fail') -f $Label) -ForegroundColor Red; return }
    $env:ComSpec = "$env:SystemRoot\system32\cmd.exe"
    try {
        $p = Start-Process -FilePath $env:ComSpec -ArgumentList "/c """"$file"" -el $Arguments""" -Verb RunAs -Wait -PassThru
        Write-Host ((T 'done') -f $Label, $p.ExitCode) -ForegroundColor Green
    } catch {
        Write-Host ((T 'uac_cancel2') -f $Label) -ForegroundColor Yellow
    } finally {
        Remove-Item -Path $file -ErrorAction SilentlyContinue
    }
}

function Invoke-MasAio([string]$Arguments = '') {
    $urls = @(
        "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/$Pinned/MAS/All-In-One-Version-KL/MAS_AIO.cmd",
        "https://dev.azure.com/massgrave/Microsoft-Activation-Scripts/_apis/git/repositories/Microsoft-Activation-Scripts/items?path=/MAS/All-In-One-Version-KL/MAS_AIO.cmd&versionType=Commit&version=$Pinned",
        "https://git.activated.win/Microsoft-Activation-Scripts/plain/MAS/All-In-One-Version-KL/MAS_AIO.cmd?id=$Pinned"
    )
    $aio = Get-String $urls (T 'mas_dl')
    if (-not $aio) { return }
    if (-not (Test-Hash $aio $AioHash)) { Write-Warning (T 'hash_fail'); return }
    Invoke-MasCmd -Response $aio -Arguments $Arguments -Label $(if ($Arguments) { $Arguments.Trim('/') } else { 'MAS menu' })
}

function Invoke-MasCheck {
    $urls = @(
        "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/$Pinned/MAS/Separate-Files-Version/Check_Activation_Status.cmd",
        "https://dev.azure.com/massgrave/Microsoft-Activation-Scripts/_apis/git/repositories/Microsoft-Activation-Scripts/items?path=/MAS/Separate-Files-Version/Check_Activation_Status.cmd&versionType=Commit&version=$Pinned",
        "https://git.activated.win/Microsoft-Activation-Scripts/plain/MAS/Separate-Files-Version/Check_Activation_Status.cmd?id=$Pinned"
    )
    $chk = Get-String $urls (T 'chk_dl')
    if (-not $chk) { return }
    if (-not (Test-Hash $chk $ChkHash)) { Write-Warning (T 'hash_fail2'); return }
    Invoke-MasCmd -Response $chk -Arguments '' -Label $(if ((Get-UiLang) -eq 'tr') { 'Durum kontrolu' } else { 'Status check' })
}

function Test-OfficeC2R {
    $cfg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\OfficeClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($cfg -and $cfg.ProductReleaseIds) { return $true }
    $cfg32 = Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\OfficeClickToRun\Configuration' -ErrorAction SilentlyContinue
    if ($cfg32 -and $cfg32.ProductReleaseIds) { return $true }
    $false
}

function Default-OfficeLang {
    $c = ((Get-Culture).Name).ToLower()
    if ($script:UiLang -eq 'tr') { 'tr-tr' }
    elseif ($script:UiLang -eq 'en') { 'en-us' }
    else { $c }
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
        if ($in -match '^[a-z]{2}(-[a-z0-9]{2,8})?$') { return $in }
        Write-Host (T 'invalid_lang') -ForegroundColor Red
    }
}

function Install-OfficeLTSC {
    Write-Host ''
    Write-Host '--- Office LTSC Professional Plus 2024 ---' -ForegroundColor Cyan
    Write-Host (T 'mode_h')
    Write-Host (T 'mode_t')
    $mode = Read-Host (T 'sel12')
    $lang = Resolve-OfficeLang

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
            $link = [regex]::Match($conf.Content, 'https://download\.microsoft\.com/[^"''\s<>]*officedeploymenttool[^"''\s<>]*\.exe', 'IgnoreCase').Value
            if (-not $link) { throw 'download link not found on page' }
            Save-File -Url $link -OutFile $odt
        } catch {
            Write-Host ((T 'odt_fail') -f $_.Exception.Message) -ForegroundColor Red
            Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
            return
        }
    }
    if ((Get-Item $odt).Length -lt 2MB) { Write-Host (T 'odt_small') -ForegroundColor Red; Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue; return }

    Write-Host (T 'extract') -ForegroundColor DarkGray
    Start-Process -FilePath $odt -ArgumentList "/extract:`"$work`" /quiet" -Wait -WindowStyle Hidden
    $setup = Join-Path $work 'setup.exe'
    if (-not (Test-Path $setup)) { Write-Host (T 'extract_fail') -ForegroundColor Red; Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue; return }

    $excludes = if ($mode -eq '2') { '' } else { "`n    <ExcludeApp ID=`"Access`" />`n    <ExcludeApp ID=`"Lync`" />`n    <ExcludeApp ID=`"OneDrive`" />`n    <ExcludeApp ID=`"OneNote`" />`n    <ExcludeApp ID=`"Outlook`" />`n    <ExcludeApp ID=`"Publisher`" />" }
    $config = @"
<Configuration ID="aw-office-ltsc-2024">
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume" PIDKEY="XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB">
    <Language ID="$lang" />$excludes
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
    $configPath = Join-Path $work 'Configuration.xml'
    [IO.File]::WriteAllText($configPath, $config)

    $isAdmin = [bool]([Security.Principal.WindowsIdentity]::GetCurrent().Groups -match 'S-1-5-32-544')
    try {
        if ($isAdmin) {
            Write-Host (T 'inst_start') -ForegroundColor DarkGray
            $p = Start-Process -FilePath $setup -ArgumentList "/configure `"$configPath`"" -Wait -PassThru -WindowStyle Hidden
            Write-Host ((T 'inst_done') -f $p.ExitCode) -ForegroundColor Green
        } else {
            Write-Host (T 'inst_uac') -ForegroundColor DarkGray
            Start-Process -FilePath 'powershell.exe' -Verb RunAs -WindowStyle Hidden -Wait -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"& '$setup' /configure '$configPath'; `$RC = `$LASTEXITCODE`""
            Write-Host ((T 'inst_done') -f '0') -ForegroundColor Green
        }
    } catch {
        Write-Host (T 'uac_cancel') -ForegroundColor Yellow
        Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
        return
    }
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue

    $ask = if ((Get-UiLang) -eq 'tr') { '^[HhYy]|^$' } else { '^[Yy]|^$' }
    $act = Read-Host (T 'ohook_ask')
    if ($act -match $ask) { Invoke-MasAio '/Ohook' }
}

function Invoke-AllInOne {
    Write-Host ''
    Write-Host (T 'all_start') -ForegroundColor Magenta
    if (-not (Test-OfficeC2R)) {
        $ask = if ((Get-UiLang) -eq 'tr') { '^[HhYy]|^$' } else { '^[Yy]|^$' }
        $r = Read-Host (T 'all_nooffice')
        if ($r -match $ask) { Write-Host (T 'all_installing') -ForegroundColor DarkGray; Install-OfficeLTSC }
    }
    # TSforge: Windows + ESU + Office hepsi tek geciste, kalici
    Invoke-MasAio '/Z-WindowsESUOffice'
    Invoke-MasCheck
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
    Write-Host (T 'm6') -ForegroundColor Yellow
    Write-Host (T 'm7') -ForegroundColor Cyan
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
        '6' { Invoke-MasAio }
        '7' { Install-OfficeLTSC }
        '8' { Invoke-AllInOne }
        '0' { return $true }
        'l' { Set-UiLang }
        'L' { Set-UiLang }
        default { }
    }
    if ($choice -ne '0') { Write-Host ''; Read-Host (T 'back') | Out-Null; Show-Menu }
}

# --- giris ---
if (-not (Test-FullLanguageMode)) { return }

# "/"-li arguman varsa dogrudan AIO'ya aktar (orn: ActivaterWindows.ps1 /HWID)
$forward = @($args) | Where-Object { $_ -match '^/' }
if ($forward) { Invoke-MasAio ($forward -join ' '); return }

if ($args -contains '-Probe') {
    Write-Banner
    $urls = @(
        "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/$Pinned/MAS/All-In-One-Version-KL/MAS_AIO.cmd",
        "https://dev.azure.com/massgrave/Microsoft-Activation-Scripts/_apis/git/repositories/Microsoft-Activation-Scripts/items?path=/MAS/All-In-One-Version-KL/MAS_AIO.cmd&versionType=Commit&version=$Pinned",
        "https://git.activated.win/Microsoft-Activation-Scripts/plain/MAS/All-In-One-Version-KL/MAS_AIO.cmd?id=$Pinned"
    )
    $aio = Get-String $urls 'PROBE'
    if ($aio -and (Test-Hash $aio $AioHash)) { Write-Host 'PROBE OK — MAS indirildi, hash dogrulandi.' -ForegroundColor Green; exit 0 }
    else { Write-Host 'PROBE FAIL' -ForegroundColor Red; exit 1 }
}

while ($true) { if (Show-Menu) { break } }
