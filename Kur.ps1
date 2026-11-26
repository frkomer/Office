function Etkinlestir-Office {
    Write-Host ""
    Write-Host "=== ACTIVATION / ETKINLESTIRME ===" -ForegroundColor Cyan
    $key = Read-Host "Enter 25-character Product Key / 25 haneli urun anahtarini girin (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)"
    if ([string]::IsNullOrWhiteSpace($key)) {
        Write-Host "No product key entered / Urun anahtari girilmedi, islem sonlandirildi." -ForegroundColor Yellow
        return
    }
    $osppPath = "C:\Program Files\Microsoft Office\Office16\ospp.vbs"
    if (-not (Test-Path $osppPath)) {
        $osppPath = "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs"
    }
    if (Test-Path $osppPath) {
        Write-Host "Installing product key / Urun anahtari yukleniyor..." -ForegroundColor Cyan
        cscript //nologo $osppPath /inpkey:$key
        Write-Host "Activating Office online / Cevrimici etkinlestiriliyor..." -ForegroundColor Cyan
        cscript //nologo $osppPath /act
    } else {
        Write-Host "ospp.vbs not found / Office kurulu degil veya etkinlestirme araci bulunamadi." -ForegroundColor Red
    }
}

Write-Host "=== OFFICE SETUP & MANAGEMENT TOOL ===" -ForegroundColor Cyan
Write-Host "1 - Download Only      (Sadece indir)" -ForegroundColor Yellow
Write-Host "2 - Download & Install (Sadece indir kur)" -ForegroundColor Yellow
Write-Host "3 - Uninstall Only     (Sadece kaldir)" -ForegroundColor Yellow
Write-Host "4 - Clean Install      (Temiz kurulum - Eskisini kaldir ve kur)" -ForegroundColor Yellow
Write-Host "5 - Activate Only      (Sadece etkinlestir)" -ForegroundColor Yellow
Write-Host ""

$islem = Read-Host "Select Option / Islem Secin (1, 2, 3, 4, 5)"

if ($islem -notin @("1", "2", "3", "4", "5")) {
    Write-Host "Invalid selection / Gecersiz secim yaptiniz. (1, 2, 3, 4, 5)" -ForegroundColor Red
    exit
}

if ($islem -eq "5") {
    Etkinlestir-Office
    exit
}

if ($islem -in @("1", "2", "4")) {
    $surum = Read-Host "Select Version / Surum Secin (2024 / 2021)"
    if ($surum -eq "2024") {
        $productId = "ProPlus2024Volume"
        $channel = "PerpetualVL2024"
    } elseif ($surum -eq "2021") {
        $productId = "ProPlus2021Volume"
        $channel = "PerpetualVL2021"
    } else {
        Write-Host "Invalid version / Gecersiz surum secimi. (2024 / 2021)" -ForegroundColor Red
        exit
    }
}

New-Item -ItemType Directory -Force -Path "C:\Office" | Out-Null
Set-Location "C:\Office"

Write-Host "Preparing Office Deployment Tool (ODT) / Kurulum araci hazirlaniyor..." -ForegroundColor Cyan
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$odtUrls = @(
    "https://microsoft.com",
    "https://microsoft.com",
    "https://microsoft.com"
)

$indirildi = $false
foreach ($url in $odtUrls) {
    try {
        Invoke-WebRequest -Uri $url -OutFile "C:\Office\odt.exe" -UseBasicParsing -UserAgent "Mozilla/5.0" -ErrorAction Stop
        $bytes = [System.IO.File]::ReadAllBytes("C:\Office\odt.exe")
        if ($bytes.Length -gt 1000000 -and $bytes[0] -eq 77 -and $bytes[1] -eq 90) {
            $indirildi = $true
            break
        }
    } catch {}
}

if (-not $indirildi) {
    Write-Host "Trying Winget fallback / Winget ile arac indiriliyor..." -ForegroundColor Yellow
    winget install --id Microsoft.OfficeDeploymentTool --accept-source-agreements --accept-package-agreements --silent | Out-Null
    $wingetSetup = Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -Filter "setup.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($wingetSetup) {
        Copy-Item $wingetSetup.FullName -Destination "C:\Office\setup.exe" -Force
        $indirildi = $true
    }
} else {
    Start-Process -FilePath "C:\Office\odt.exe" -ArgumentList "/quiet /extract:C:\Office" -Wait
}

if (-not (Test-Path "C:\Office\setup.exe")) {
    Write-Host "ERROR: setup.exe could not be downloaded. Check internet connection. / HATA: Kurulum araci indirilemedi." -ForegroundColor Red
    exit
}

$kaldirXml = @"
<Configuration>
  <Remove All="TRUE" />
  <RemoveMSI />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@
$kaldirXml | Out-File -FilePath "C:\Office\remove.xml" -Encoding UTF8

if ($islem -in @("1", "2")) {
    $kurXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$channel" SourcePath="C:\Office" AllowCdnFallback="TRUE">
    <Product ID="$productId">
      <Language ID="tr-tr" />
    </Product>
  </Add>
  <RemoveMSI />
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@
    $kurXml | Out-File -FilePath "C:\Office\configuration.xml" -Encoding UTF8
} elseif ($islem -eq "4") {
    $temizKurXml = @"
<Configuration>
  <Remove All="TRUE" />
  <RemoveMSI />
  <Add OfficeClientEdition="64" Channel="$channel" SourcePath="C:\Office" AllowCdnFallback="TRUE">
    <Product ID="$productId">
      <Language ID="tr-tr" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@
    $temizKurXml | Out-File -FilePath "C:\Office\configuration.xml" -Encoding UTF8
}

if ($islem -eq "1") {
    Write-Host "Downloading Office $surum to C:\Office / Dosyalar indiriliyor..." -ForegroundColor Green
    Start-Process -FilePath "C:\Office\setup.exe" -ArgumentList "/download C:\Office\configuration.xml" -Wait
    Write-Host "Download complete! / Indirme tamamlandi! (C:\Office)" -ForegroundColor Green
} elseif ($islem -eq "2") {
    Write-Host "Step 1/2: Downloading Office $surum / Dosyalar indiriliyor..." -ForegroundColor Green
    Start-Process -FilePath "C:\Office\setup.exe" -ArgumentList "/download C:\Office\configuration.xml" -Wait
    Write-Host "Step 2/2: Installing Office $surum / Kurulum baslatiliyor..." -ForegroundColor Green
    Start-Process -FilePath "C:\Office\setup.exe" -ArgumentList "/configure C:\Office\configuration.xml" -Wait
    Write-Host "Installation complete! / Kurulum tamamlandi!" -ForegroundColor Green
} elseif ($islem -eq "3") {
    Write-Host "Uninstalling all Office versions / Tum surumler arka planda kaldiriliyor..." -ForegroundColor Yellow
    Start-Process -FilePath "C:\Office\setup.exe" -ArgumentList "/configure C:\Office\remove.xml" -Wait
    Write-Host "Uninstall complete! / Kaldirma islemi tamamlandi!" -ForegroundColor Green
} elseif ($islem -eq "4") {
    Write-Host "Step 1/2: Downloading Office $surum / Dosyalar indiriliyor..." -ForegroundColor Green
    Start-Process -FilePath "C:\Office\setup.exe" -ArgumentList "/download C:\Office\configuration.xml" -Wait
    Write-Host "Step 2/2: Removing old & installing Office $surum silently / Eskisi kaldirilip kuruluyor..." -ForegroundColor Green
    Start-Process -FilePath "C:\Office\setup.exe" -ArgumentList "/configure C:\Office\configuration.xml" -Wait
    Write-Host "All tasks completed! / Tum islemler basariyla tamamlandi!" -ForegroundColor Green
}

if ($islem -in @("2", "4")) {
    Write-Host ""
    $etkinlestir = Read-Host "Activate with Product Key / Urun anahtari ile etkinlestir? (E / H)"
    if ($etkinlestir -in @("E", "e", "Y", "y")) {
        Etkinlestir-Office
    }
}
