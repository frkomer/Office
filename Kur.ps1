$surum = Read-Host "Kurmak istediginiz surumu secin (2024 veya 2021 yazin)"

if ($surum -eq "2024") {
    $productId = "ProPlus2024Volume"
    $channel = "PerpetualVL2024"
} elseif ($surum -eq "2021") {
    $productId = "ProPlus2021Volume"
    $channel = "PerpetualVL2021"
} else {
    Write-Host "Gecersiz secim yaptiniz. Lutfen 2024 veya 2021 yazin." -ForegroundColor Red
    exit
}

New-Item -ItemType Directory -Force -Path "C:\Office" | Out-Null
Set-Location "C:\Office"

Write-Host "Guncel Office Dagitim Araci (ODT) baglantisi bulunuyor..." -ForegroundColor Cyan
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$sayfa = Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/download/details.aspx?id=49117" -UseBasicParsing
$eslesme = [regex]::Match($sayfa.Content, 'https://download\.microsoft\.com/download/[^"''\s]+officedeploymenttool[^"''\s]+\.exe')

if ($eslesme.Success) {
    $odtUrl = $eslesme.Value
} else {
    $odtUrl = "https://microsoft.com"
}

Write-Host "Office Dagitim Araci indiriliyor..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $odtUrl -OutFile "C:\Office\odt.exe" -UseBasicParsing
Start-Process -FilePath "C:\Office\odt.exe" -ArgumentList "/quiet /extract:C:\Office" -Wait

$xml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$channel">
    <Product ID="$productId">
      <Language ID="tr-tr" />
    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@

$xml | Out-File -FilePath "C:\Office\configuration.xml" -Encoding UTF8

Write-Host "Office $surum kurulumu baslatiliyor..." -ForegroundColor Green
Start-Process -FilePath "C:\Office\setup.exe" -ArgumentList "/configure C:\Office\configuration.xml" -Wait
