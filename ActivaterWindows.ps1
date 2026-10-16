# ActivaterWindows v2.0 — Terminal Edition
# PowerShell + Windows 10/11. Aktivasyon: Microsoft Activation Scripts (pinned commit, SHA256 dogrulamali).
# Tek satir: irm https://raw.githubusercontent.com/tarikelertarnak/ActivaterWindows/main/ActivaterWindows.ps1 | iex
# Yerel: ActivaterWindows.cmd (cift tik) veya powershell -ExecutionPolicy Bypass -File ActivaterWindows.ps1
# Parametre: -Probe (indir+dogrula testi, calistirmadan cikar) | diger "/"-li argumanlar dogrudan MAS AIO'ya iletilir.

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
$RepoUrl  = 'https://github.com/tarikelertarnak/ActivaterWindows'
$Pinned   = '05c4f881efec946c0040cdd552d1afa9a519704b' # MAS commit
$AioHash  = '850F979665FB93999ACAE93F4790C1FF8ED2041532060B7966A121C2D29A0BFA'
$ChkHash  = '6DF2937385386D0C9682D351797E674583A410F4792BB7BE45302F40AB90287E'

function Write-Banner {
    Write-Host ''
    Write-Host '==========================================================' -ForegroundColor DarkCyan
    Write-Host '  ActivaterWindows — Terminal' -ForegroundColor Cyan
    Write-Host "  $RepoUrl" -ForegroundColor DarkGray
    Write-Host '==========================================================' -ForegroundColor DarkCyan
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

function Get-String([string[]]$Urls, [string]$Label) {
    Write-Progress -Activity "$Label indiriliyor" -Status 'Lutfen bekleyin'
    $errors = @(); $response = $null
    foreach ($url in $Urls | Sort-Object { Get-Random }) {
        try {
            if ($PSVersionTable.PSVersion.Major -ge 3) { $response = Invoke-RestMethod -Uri $url -UseBasicParsing }
            else { $response = (New-Object Net.WebClient).DownloadString($url) }
            break
        } catch { $errors += $_ }
    }
    Write-Progress -Activity "$Label indiriliyor" -Status 'Tamam' -Completed
    if (-not $response) {
        Warn-AV
        foreach ($e in $errors) { Write-Host "Hata: $($e.Exception.Message)" -ForegroundColor Red }
        Write-Host "$Label hicbir kaynaktan alinamadi. Antivirus/firewall baglantayi engelliyor olabilir." -ForegroundColor Red
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
    if (-not (Test-Path $file)) { Warn-AV; Write-Host "Gecici dosya olusturulamadi, iptal! ($Label)" -ForegroundColor Red; return }
    $env:ComSpec = "$env:SystemRoot\system32\cmd.exe"
    try {
        $p = Start-Process -FilePath $env:ComSpec -ArgumentList "/c """"$file"" -el $Arguments""" -Verb RunAs -Wait -PassThru
        Write-Host "$Label tamamlandi. (cikis kodu: $($p.ExitCode))" -ForegroundColor Green
    } catch {
        Write-Host "$Label iptal edildi (UAC reddedildi)." -ForegroundColor Yellow
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
    $aio = Get-String $urls 'MAS All-In-One'
    if (-not $aio) { return }
    if (-not (Test-Hash $aio $AioHash)) {
        Write-Warning "MAS butunluk kontrolu basarisiz (hash uyusmadi), iptal!"
        return
    }
    Invoke-MasCmd -Response $aio -Arguments $Arguments -Label $(if ($Arguments) { $Arguments.Trim('/') } else { 'MAS menusu' })
}

function Invoke-MasCheck {
    $urls = @(
        "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/$Pinned/MAS/Separate-Files-Version/Check_Activation_Status.cmd",
        "https://dev.azure.com/massgrave/Microsoft-Activation-Scripts/_apis/git/repositories/Microsoft-Activation-Scripts/items?path=/MAS/Separate-Files-Version/Check_Activation_Status.cmd&versionType=Commit&version=$Pinned",
        "https://git.activated.win/Microsoft-Activation-Scripts/plain/MAS/Separate-Files-Version/Check_Activation_Status.cmd?id=$Pinned"
    )
    $chk = Get-String $urls 'Aktivasyon durumu kontrolu'
    if (-not $chk) { return }
    if (-not (Test-Hash $chk $ChkHash)) {
        Write-Warning 'Kontrol scripti butunluk kontrolu basarisiz, iptal!'
        return
    }
    Invoke-MasCmd -Response $chk -Arguments '' -Label 'Durum kontrolu'
}

function Install-OfficeLTSC {
    Write-Host ''
    Write-Host '--- Office LTSC Professional Plus 2024 Kurulumu ---' -ForegroundColor Cyan
    Write-Host '[1] Onerilen (hafif): Word, Excel, PowerPoint, OneDrive haric bilesenler kurulmaz'
    Write-Host '[2] Tam paket: tum bilesenler'
    $mode = Read-Host 'Secim (1/2)'
    $sysLang = ((Get-Culture).Name).ToLower()
    $lang = Read-Host "Dil kodu [otomatik secildi: $sysLang — degistirmek istemiyorsan Enter'a bas]"
    if (-not $lang) { $lang = $sysLang }
    if ($lang -notmatch '^[a-z]{2}(-[a-z0-9]{2,8})?$') { Write-Host 'Gecersiz dil kodu.' -ForegroundColor Red; return }

    $work = Join-Path ([IO.Path]::GetTempPath()) "AW_ODT_$(Get-Random)"
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    $odt = Join-Path $work 'odt.exe'

    # 1) repo icindeki yerel kopya (script dosyasiyla birlikte calisirken)
    $local = $null
    if ($PSScriptRoot) {
        $local = Get-ChildItem -Path (Join-Path $PSScriptRoot 'resources') -Filter 'officedeploymenttool*.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    if ($local) {
        Write-Host "ODT yerel kopyadan aliniyor: $($local.Name)" -ForegroundColor DarkGray
        Copy-Item $local.FullName $odt
    } else {
        # 2) MS indirme sayfasindan dogrudan .exe linki (aka.ms/odt artik sadece dokumantasyona gider)
        Write-Host 'Office Deployment Tool indiriliyor (microsoft.com)...' -ForegroundColor DarkGray
        try {
            $conf = Invoke-WebRequest -Uri 'https://www.microsoft.com/en-us/download/details.aspx?id=49117' -UseBasicParsing
            $link = [regex]::Match($conf.Content, 'https://download\.microsoft\.com/[^"''\s<>]*officedeploymenttool[^"''\s<>]*\.exe', 'IgnoreCase').Value
            if (-not $link) { throw 'indirme linki sayfada bulunamadi' }
            if ($PSVersionTable.PSVersion.Major -ge 3) { Invoke-WebRequest -Uri $link -OutFile $odt -UseBasicParsing }
            else { (New-Object Net.WebClient).DownloadFile($link, $odt) }
        } catch {
            Write-Host "ODT indirilemedi: $($_.Exception.Message)" -ForegroundColor Red
            Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
            return
        }
    }
    if ((Get-Item $odt).Length -lt 2MB) { Write-Host 'ODT indirilen dosya yanlis icerik, iptal.' -ForegroundColor Red; Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue; return }

    Write-Host 'ODT ayiklaniyor...' -ForegroundColor DarkGray
    Start-Process -FilePath $odt -ArgumentList "/extract:`"$work`" /quiet" -Wait
    $setup = Join-Path $work 'setup.exe'
    if (-not (Test-Path $setup)) { Write-Host 'setup.exe ayiklanamadi, iptal.' -ForegroundColor Red; Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue; return }

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
            Write-Host 'Office kuruluyor (bir kac dakika surebilir)...' -ForegroundColor DarkGray
            $p = Start-Process -FilePath $setup -ArgumentList "/configure `"$configPath`"" -Wait -PassThru
            Write-Host "Kurulum bitti (cikis kodu: $($p.ExitCode))." -ForegroundColor Green
        } else {
            Write-Host 'Office kuruluyor (UAC izni gerekecek)...' -ForegroundColor DarkGray
            Start-Process -FilePath 'powershell.exe' -Verb RunAs -Wait -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"& '$setup' /configure '$configPath'; `$RC = `$LASTEXITCODE`""
            Write-Host 'Kurulum bitti.' -ForegroundColor Green
        }
    } catch {
        Write-Host 'Kurulum iptal edildi (UAC reddedildi).' -ForegroundColor Yellow
        Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
        return
    }
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue

    $act = Read-Host 'Office''i simdi Ohook ile etkinlestir? (E/h)'
    if ($act -match '^[EeYy]|^$') { Invoke-MasAio '/Ohook' }
}

function Show-Menu {
    Clear-Host
    Write-Banner
    Write-Host ''
    Write-Host '  [1] Windows 10/11 etkinlestir  — HWID   (kalici)' -ForegroundColor White
    Write-Host '  [2] Office etkinlestir         — Ohook  (kalici)' -ForegroundColor White
    Write-Host '  [3] Windows + ESU + Office     — TSforge (kalici)' -ForegroundColor White
    Write-Host '  [4] Windows 10 ESU             — TSforge (kalici)' -ForegroundColor White
    Write-Host '  [5] Windows / Office           — Online KMS (180 gun)' -ForegroundColor White
    Write-Host '  [6] Aktivasyon durumu — kontrol et' -ForegroundColor White
    Write-Host '  [7] TAM MAS MENUSU — tum ozellikler (Troubleshoot, Surum degistir, $OEM$...)' -ForegroundColor Yellow
    Write-Host '  [8] Office LTSC Pro Plus 2024 kur (+Ohook)' -ForegroundColor Cyan
    Write-Host '  [0] Cikis' -ForegroundColor DarkGray
    Write-Host ''
    $choice = Read-Host 'Secim'
    switch ($choice) {
        '1' { Invoke-MasAio '/HWID' }
        '2' { Invoke-MasAio '/Ohook' }
        '3' { Invoke-MasAio '/Z-WindowsESUOffice' }
        '4' { Invoke-MasAio '/Z-ESU' }
        '5' { Invoke-MasAio '/K-WindowsOffice' }
        '6' { Invoke-MasCheck }
        '7' { Invoke-MasAio }
        '8' { Install-OfficeLTSC }
        '0' { return $true }
        default { }
    }
    if ($choice -ne '0') { Write-Host ''; Read-Host 'Menuye donmek icin Enter' | Out-Null; Show-Menu }
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
