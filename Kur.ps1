Write-Host "=== OFFICE SETUP & MANAGEMENT TOOL ===" -ForegroundColor Cyan
Write-Host "1 - Download Only      (Sadece indir)" -ForegroundColor Yellow
Write-Host "2 - Download & Install (Sadece indir kur)" -ForegroundColor Yellow
Write-Host "3 - Uninstall Only     (Sadece kaldir)" -ForegroundColor Yellow
Write-Host "4 - Clean Install      (Eskisi varsa kaldir yenisini indir kur)" -ForegroundColor Yellow
Write-Host ""

$islem = Read-Host "Select Option / Islem Secin (1, 2, 3, 4)"

if ($islem -notin @("1", "2", "3", "4")) {
    Write-Host "Invalid selection / Gecersiz secim yaptiniz. (1, 2, 3, 4)" -ForegroundColor Red
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

Write-Host "Finding latest Office Deployment Tool (ODT) / Guncel arac bulunuyor..." -ForegroundColor Cyan
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$sayfa = Invoke-WebRequest -Uri "https://microsoft.com" -UseBasicParsing
$eslesme = [regex]::Match($sayfa.Content, 'https://download\.microsoft\.com/download/[^"''\s]+officedeploymenttool[^"''\s]+\.exe')

if ($eslesme.Success) {
    $odtUrl = $eslesme.Value
} else {
    $odtUrl = "https://microsoft.com"
}

Write-Host "Downloading ODT / Office Dagitim Araci indiriliyor..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $odtUrl -OutFile "C:\Office\odt.exe" -UseBasicParsing
Start-Process -FilePath "C:\Office\odt.exe" -ArgumentList "/quiet /extract:C:\Office" -Wait

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
    Write-Host "=== ACTIVATION / ETKINLESTIRME ===" -ForegroundColor Cyan
    $etkinlestir = Read-Host "Activate with Product Key / Urun anahtari ile etkinlestir? (E / H)"
    if ($etkinlestir -in @("E", "e", "Y", "y")) {
        $key = Read-Host "Enter 25-character Product Key / 25 haneli urun anahtarini girin (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)"
        $osppPath = "C:\Program Files\Microsoft Office\Office16\ospp.vbs"
        if (Test-Path $osppPath) {
            Write-Host "Installing product key / Urun anahtari yukleniyor..." -ForegroundColor Cyan
            cscript //nologo $osppPath /inpkey:$key
            Write-Host "Activating Office online / Cevrimici etkinlestiriliyor..." -ForegroundColor Cyan
            cscript //nologo $osppPath /act
        } else {
            Write-Host "ospp.vbs not found / Etkinlestirme araci bulunamadi." -ForegroundColor Red
        }
    }
}
