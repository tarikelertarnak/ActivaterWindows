# Office16 klasörü var mı? (Registry yoksa bile fix çalışsın)
$officeRoot = @(
    "$env:ProgramFiles\Microsoft Office\root\Office16",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16",
    "$env:ProgramFiles\Microsoft Office\root\Office15",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office15"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $officeRoot) {
    Write-Host 'Office kurulu degil (Office16/Office15 klasoru yok).' -ForegroundColor Red
    return
}

Write-Host "Office bulundu: $officeRoot" -ForegroundColor Green

# 1) Eski hesap kayitlarini temizle
$idRoot = 'HKCU:\Software\Microsoft\Office\16.0\Common\Identity'
if (Test-Path $idRoot) {
    Get-ChildItem $idRoot -EA SilentlyContinue | ForEach-Object { Remove-Item $_.PSPath -Recurse -Force -EA SilentlyContinue }
    Set-ItemProperty -Path $idRoot -Name 'ConnectedOneAuthAccountId' -Value '' -Force -EA SilentlyContinue
    Write-Host 'Identity kayitlari temizlendi' -ForegroundColor Green
} else {
    Write-Host 'Temizlenecek Identity kaydi yok.' -ForegroundColor DarkGray
}

# Ek hesap kalintilari
$msalKeys = @(
    'HKCU:\Software\Microsoft\Office\16.0\Common\Identity\OneAuthConfigCache',
    'HKCU:\Software\Microsoft\Office\16.0\Common\ServicesManagerCache',
    'HKCU:\Software\Microsoft\Office\16.0\Common\Roaming\Identities'
)
foreach ($k in $msalKeys) {
    if (Test-Path $k) { Remove-Item -Path $k -Recurse -Force -EA SilentlyContinue }
}

# 2) Office16 klasor ve exe izinlerini duzelt
Write-Host "Dosya izinleri onariliyor: $officeRoot" -ForegroundColor Cyan
$usersRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    'Users', 'ReadAndExecute, Synchronize',
    'ContainerInherit,ObjectInherit', 'None', 'Allow')
try {
    $acl = Get-Acl $officeRoot
    $acl.SetAccessRule($usersRule)
    Set-Acl -Path $officeRoot -AclObject $acl
    Write-Host 'Klasor izinleri guncellendi.' -ForegroundColor Green
} catch { Write-Host "Klasor hata: $($_.Exception.Message)" -ForegroundColor Yellow }

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
Write-Host 'EXE izinleri guncellendi.' -ForegroundColor Green

# 3) HKCR ProgID + App Paths yeniden yaz
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
        if (-not (Test-Path $cmdKey)) { New-Item -Path $cmdKey -Force | Out-Null }
        Set-ItemProperty -Path $cmdKey -Name '(Default)' -Value "`"$exePath`" `"%1`"" -Force
    }
    $apKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$exe"
    if (-not (Test-Path $apKey)) { New-Item -Path $apKey -Force | Out-Null }
    Set-ItemProperty -Path $apKey -Name '(Default)' -Value $exePath -Force
    Set-ItemProperty -Path $apKey -Name 'Path' -Value $officeRoot -Force
}
Write-Host 'HKCR + App Paths guncellendi.' -ForegroundColor Green

# 4) UserChoice + OpenWithProgids temizle
$exts = @('.doc','.docx','.dot','.dotx','.rtf','.xls','.xlsx','.xlsm','.xlsb','.xltx','.ppt','.pptx','.pptm','.potx')
foreach ($ext in $exts) {
    foreach ($sub in @('UserChoice','OpenWithProgids')) {
        $k = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$ext\$sub"
        if (Test-Path $k) { Remove-Item -Path $k -Recurse -Force -EA SilentlyContinue }
    }
}
Write-Host 'UserChoice temizlendi.' -ForegroundColor DarkGray

# 5) C2RClient repair + ie4uinit + explorer
$c2r = @(
    "$env:ProgramFiles\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe",
    "${env:ProgramFiles(x86)}\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($c2r) {
    Write-Host 'Office repair baslatiliyor...' -ForegroundColor Cyan
    $pr = Start-Process -FilePath $c2r -ArgumentList '/repair onshowinuisuppressed=true' -PassThru -WindowStyle Hidden -EA SilentlyContinue
    if ($pr) { $pr.WaitForExit() }
    Write-Host 'Office repair tamamlandi.' -ForegroundColor Green
} else {
    Write-Host 'OfficeC2RClient.exe bulunamadi (repair atlandi).' -ForegroundColor DarkYellow
}

$ie4 = "$env:SystemRoot\System32\ie4uinit.exe"
if (Test-Path $ie4) { Start-Process -FilePath $ie4 -ArgumentList '-show' -WindowStyle Hidden -Wait -EA SilentlyContinue }

Stop-Process -Name explorer -Force -EA SilentlyContinue
Start-Sleep -Seconds 2
Start-Process explorer.exe

Write-Host ''
Write-Host 'TAMAMLANDI - Word/Excel/PowerPoint artik yonetici olmadan acilmali' -ForegroundColor Green
Write-Host '  - Eski M365 hesabi kimlik kayitlari silindi' -ForegroundColor DarkGray
Write-Host '  - Office16 klasor + exe izinleri duzeltildi' -ForegroundColor DarkGray
Write-Host '  - HKCR ProgID + App Paths yeniden yazildi' -ForegroundColor DarkGray
Write-Host '  - UserChoice temizlendi + C2R repair + ie4uinit + explorer' -ForegroundColor DarkGray
Write-Host 'Sorun devam ederse oturumu kapat/ac veya bilgisayari yeniden baslat.' -ForegroundColor DarkYellow