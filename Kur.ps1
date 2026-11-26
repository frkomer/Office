Write-Host "=== OFFICE YONETIM VE KURULUM ARACI ===" -ForegroundColor Cyan
Write-Host "1 - Sadece indir" -ForegroundColor Yellow
Write-Host "2 - Sadece indir kur" -ForegroundColor Yellow
Write-Host "3 - Sadece kaldir" -ForegroundColor Yellow
Write-Host "4 - Eskisi varsa kaldir yenisini indir kur (Arka planda)" -ForegroundColor Yellow
Write-Host ""

$islem = Read-Host "Yapmak istediginiz islemi secin (1, 2, 3 veya 4)"

if ($islem -notin @("1", "2", "3", "4")) {
    Write-Host "Gecersiz secim yaptiniz. Lutfen 1, 2, 3 veya 4 yazin." -ForegroundColor Red
    exit
}

if ($islem -in @("1", "2", "4")) {
    $surum = Read-Host "Kurmak veya indirmek istediginiz surumu secin (2024 veya 2021 yazin)"
    if ($surum -eq "2024") {
        $productId = "ProPlus2024Volume"
        $channel = "PerpetualVL2024"
    } elseif ($surum -eq "2021") {
        $productId = "ProPlus2021Volume"
        $channel = "PerpetualVL2021"
    } else {
        Write-Host "Gecersiz surum secimi. Lutfen 2024 veya 2021 yazin." -ForegroundColor Red
        exit
    }
}

New-Item -ItemType Directory -Force -Path "C:\Office" | Out-Null
Set-Location "C:\Office"

Write-Host "Guncel Office Dagitim Araci (ODT) baglantisi bulunuyor..." -ForegroundColor Cyan
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$sayfa = Invoke-WebRequest -Uri "https://microsoft.com" -UseBasicParsing
$eslesme = [regex]::Match($sayfa.Content, 'https://download\.microsoft\.com/download/[^"''\s]+officedeploymenttool[^"''\s]+\.exe')

if ($eslesme.Success) {
    $odtUrl = $eslesme.Value
} else {
    $odtUrl = "https://microsoft.com"
}

Write-Host "Office Dagitim Araci indiriliyor..." -ForegroundColor Cyan
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
    Write-Host "Office $surum dosyalari C:\Office klasorune indiriliyor (Lutfen bekleyin)..." -ForegroundColor Green
    Start-Process -FilePath "C:\Office\setup.exe" -ArgumentList "/download C:\Office\configuration.xml" -Wait
    Write-Host "Indirme tamamlandi! Tum dosyalar C:\Office klasorunde hazir." -ForegroundColor Green
} elseif ($islem -eq "2") {
    Write-Host "Adim 1/2: Office $surum dosyalari C:\Office klasorune indiriliyor..." -ForegroundColor Green
    Start-Process -FilePath "C:\Office\setup.exe" -ArgumentList "/download C:\Office\configuration.xml" -Wait
    Write-Host "Adim 2/2: Kurulum baslatiliyor..." -ForegroundColor Green
    Start-Process -FilePath "C:\Office\setup.exe" -ArgumentList "/configure C:\Office\configuration.xml" -Wait
    Write-Host "Kurulum tamamlandi!" -ForegroundColor Green
} elseif ($islem -eq "3") {
    Write-Host "Bilgisayardaki tum Office surumleri arka planda kaldiriliyor..." -ForegroundColor Yellow
    Start-Process -FilePath "C:\Office\setup.exe" -ArgumentList "/configure C:\Office\remove.xml" -Wait
    Write-Host "Kaldirma islemi tamamlandi!" -ForegroundColor Green
} elseif ($islem -eq "4") {
    Write-Host "Adim 1/2: Office $surum dosyalari C:\Office klasorune indiriliyor..." -ForegroundColor Green
    Start-Process -FilePath "C:\Office\setup.exe" -ArgumentList "/download C:\Office\configuration.xml" -Wait
    Write-Host "Adim 2/2: Eskisi kaldirilip Office $surum arka planda kuruluyor (Lutfen bekleyin)..." -ForegroundColor Green
    Start-Process -FilePath "C:\Office\setup.exe" -ArgumentList "/configure C:\Office\configuration.xml" -Wait
    Write-Host "Tum islemler basariyla tamamlandi!" -ForegroundColor Green
}
