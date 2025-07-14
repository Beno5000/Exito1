# Ruta al manifiesto del complemento
$manifestPath = "D:\Complementos en JavaScripExcel\excel-get-started-with-dev-kit\manifest.xml"

# Carpeta donde Office carga los complementos personalizados
$addinFolder = "$env:LOCALAPPDATA\Microsoft\Office\Addins"

Write-Host "Verificando carpeta de complementos..."

# Crear la carpeta si no existe
if (!(Test-Path -Path $addinFolder)) {
    Write-Host "Carpeta no existe. Creando..."
    New-Item -ItemType Directory -Force -Path $addinFolder | Out-Null
} 
else {
    Write-Host "Carpeta ya existe."
}

# Copiar el manifiesto
Write-Host "Copiando manifest.xml a $addinFolder..."
Copy-Item -Path $manifestPath -Destination $addinFolder -Force

# Registrar el complemento usando npx
Write-Host "Registrando complemento con office-addin-dev-settings..."
npx office-addin-dev-settings register $manifestPath

# Abrir Excel
Write-Host "Abriendo Excel..."
Start-Process "excel.exe"
