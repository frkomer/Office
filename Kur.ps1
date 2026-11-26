$secim = Read-Host "Kurmak istediginiz surumu secin (2024 veya 2021 yazin)"

New-Item -ItemType Directory -Path "C:\Office" -Force | Out-Null
Set-Location -Path "C:\Office"

Write-Host "Office Dagitim Araci indiriliyor..." -ForegroundColor Cyan
Invoke-WebRequest -Uri "https://microsoft.com" -OutFile "odt.exe"
Start-Process -FilePath ".\odt.exe" -ArgumentList "/quiet /extract:C:\Office" -Wait

if ($secim -eq "2021") {
    $channel = "PerpetualVL2021"
    $product = "ProPlus2021Volume"
} else {
    $channel = "PerpetualVL2024"
    $product = "ProPlus2024Volume"
}

@"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$channel">
    <Product ID="$product">
      <Language ID="tr-tr" />
    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@ | Out-File -FilePath "C:\Office\configuration.xml" -Encoding UTF8

Write-Host "Office $secim kurulumu baslatiliyor..." -ForegroundColor Green
.\setup.exe /configure configuration.xml
