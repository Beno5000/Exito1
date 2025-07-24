# ===============================
# Script PowerShell: abrir-con-complemento.ps1
# ===============================

$manifiesto = "manifest.xml"
$destino = "$env:LOCALAPPDATA\Microsoft\Office\Addins"

Write-Host "`n=== Iniciando instalación de complemento personalizado para Excel ==="

# Verificar existencia del manifiesto
if (!(Test-Path $manifiesto)) {
    Write-Host "❌ No se encontró el archivo manifest.xml en esta carpeta." -ForegroundColor Red
    exit
}

# Crear carpeta destino si no existe
if (!(Test-Path $destino)) {
    New-Item -ItemType Directory -Path $destino | Out-Null
    Write-Host "✅ Carpeta de complementos creada."
} else {
    Write-Host "✅ Carpeta de complementos ya existe."
}

# Copiar el manifiesto
Copy-Item -Path $manifiesto -Destination $destino -Force
Write-Host "✅ Manifest.xml copiado a $destino"

# Verificar si WebView2 está instalado
$webview2Path = "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F1E7E6DA-A2DB-4F67-87DE-EE4C2DB4C6D1}"
$webview2Installed = Test-Path $webview2Path

if ($webview2Installed) {
    Write-Host "✅ WebView2 está instalado." -ForegroundColor Green
} else {
    Write-Host "❌ WebView2 NO está instalado. El complemento emergente no funcionará correctamente." -ForegroundColor Red
}

# Abrir Excel
Start-Process "excel.exe"
Write-Host "✅ Excel se está abriendo..."
Write-Host "=== Proceso finalizado ===`n"
