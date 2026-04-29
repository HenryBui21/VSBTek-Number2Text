# ---------------------------------------------------------------------------------------
# Script: VSBTek-Installer.ps1
# Author: VSBTek
# Purpose: Giai phap cai dat tu dong toan dien cho Excel Add-in
# ---------------------------------------------------------------------------------------

$AddinName = "VSBTek-Number2Text.xlam"
$AddinPath = "$PSScriptRoot\$AddinName"
$SourceDir  = "$PSScriptRoot\src"

Clear-Host
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "          VSBTek NUMBER TO TEXT - INSTALLER               " -ForegroundColor White -BackgroundColor Blue
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $SourceDir)) {
    Write-Host "[!] Khong tim thay thu muc 'src'!" -ForegroundColor Red
    pause; exit
}

if (Get-Process excel -ErrorAction SilentlyContinue) {
    Write-Host "[!] Dang tat Excel..." -ForegroundColor Yellow
    Stop-Process -Name excel -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

# 1. Cau hinh Registry
Write-Host "[*] Dang cau hinh quyen bao mat Excel..." -ForegroundColor White
$ExcelVersions = @("16.0", "15.0", "14.0")
foreach ($ver in $ExcelVersions) {
    $SecurityPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if (Test-Path $SecurityPath) {
        Set-ItemProperty -Path $SecurityPath -Name "AccessVBOM"  -Value 1 -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $SecurityPath -Name "VBAWarnings" -Value 1 -ErrorAction SilentlyContinue # 1 = Enable all macros
    }
}
Write-Host " -> Da cau hinh xong." -ForegroundColor Green

# 2. Build bang VBScript (chay STA mode - tranh loi COM cua PowerShell)
Write-Host "[*] Dang Build Add-in (qua VBScript)..." -ForegroundColor White

$bas1 = "$SourceDir\modFontConverter.bas"
$bas2 = "$SourceDir\modNumber2TextCore.bas"
$bas3 = "$SourceDir\modNumber2TextEng.bas"
$bas4 = "$SourceDir\modPublicAPI.bas"
$vbsLog = Join-Path $env:TEMP "VSBTek_Build_Result.txt"
$vbsPath = Join-Path $env:TEMP "VSBTek_Build.vbs"

# Xoa log cu neu co
if (Test-Path $vbsLog) { Remove-Item $vbsLog -Force }

$TempGuid = [Guid]::NewGuid().ToString().Substring(0,8)
$TempXlam = Join-Path $env:TEMP "VSBTek_Build_$TempGuid.xlam"

$vbsContent = @"
On Error Resume Next

Dim oXL, oWB, oVBP, oFSO, oFile
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oXL  = CreateObject("Excel.Application")

If Err.Number <> 0 Then
    Set oFile = oFSO.CreateTextFile("$vbsLog", True)
    oFile.Write "ERROR:Excel:" & Err.Number & ":" & Err.Description
    oFile.Close : WScript.Quit 1
End If

' CHAY AN TRONG BACKGROUND
oXL.Visible = False 
oXL.DisplayAlerts = False
oXL.AutomationSecurity = 2 ' msoAutomationSecurityLow

Set oWB  = oXL.Workbooks.Add()

Err.Clear
' B1: IMPORT CODE
Set oVBP = oWB.VBProject
oVBP.VBComponents.Import "$bas1"
oVBP.VBComponents.Import "$bas2"
oVBP.VBComponents.Import "$bas3"
oVBP.VBComponents.Import "$bas4"

If Err.Number <> 0 Then
    Set oFile = oFSO.CreateTextFile("$vbsLog", True)
    oFile.Write "ERROR:Import:" & Err.Number & ":" & Err.Description
    oFile.Close : oWB.Close False : oXL.Quit : WScript.Quit 1
End If

' B2: LUU FILE THANH ADDIN XLAM
Err.Clear
oWB.IsAddin = True
oWB.SaveAs "$TempXlam", 55

If Err.Number <> 0 Then
    Set oFile = oFSO.CreateTextFile("$vbsLog", True)
    oFile.Write "ERROR:Save_AfterImport:" & Err.Number & ":" & Err.Description
    oFile.Close : oWB.Close False : oXL.Quit : WScript.Quit 1
End If

oWB.Close False
oXL.Quit

Set oFile = oFSO.CreateTextFile("$vbsLog", True)
oFile.Write "SUCCESS"
oFile.Close
WScript.Quit 0
"@

Set-Content -Path $vbsPath -Value $vbsContent -Encoding ASCII

$proc = Start-Process -FilePath "wscript.exe" -ArgumentList "`"$vbsPath`"" -Wait -PassThru
Start-Sleep -Seconds 1

# Kiem tra ket qua
if (Test-Path $vbsLog) {
    $result = Get-Content $vbsLog -Raw
    if ($result -like "SUCCESS*") {
        # Di chuyen file xlam
        if (Test-Path $AddinPath) { Remove-Item $AddinPath -Force }
        Move-Item -Path $TempXlam -Destination $AddinPath -Force
        Write-Host " -> Build thanh cong: $AddinName" -ForegroundColor Green
    } else {
        Write-Host "[!] LOI Build: $result" -ForegroundColor Red
        Remove-Item $vbsPath  -Force -ErrorAction SilentlyContinue
        Remove-Item $TempXlam -Force -ErrorAction SilentlyContinue
        pause; exit
    }
} else {
    Write-Host "[!] Khong doc duoc ket qua build." -ForegroundColor Red
    pause; exit
}

Remove-Item $vbsPath -Force -ErrorAction SilentlyContinue
Remove-Item $vbsLog -Force -ErrorAction SilentlyContinue

# 3. Unblock, Trusted Location, Dang ky Add-in
Write-Host "[*] Dang dang ky Add-in..." -ForegroundColor White
Unblock-File -Path $AddinPath -ErrorAction SilentlyContinue

foreach ($ver in $ExcelVersions) {
    $RegBase = "HKCU:\Software\Microsoft\Office\$ver\Excel"
    $TLPath = "$RegBase\Security\Trusted Locations\VSBTekNumber2Text"
    if (-not (Test-Path $TLPath)) { New-Item -Path $TLPath -Force | Out-Null }
    Set-ItemProperty -Path $TLPath -Name "Path"        -Value $PSScriptRoot
    Set-ItemProperty -Path $TLPath -Name "Description" -Value "VSBTek Addin"

    $OptionsPath = "$RegBase\Options"
    if (Test-Path $OptionsPath) {
        $slotFound = $false
        for ($i = 0; $i -lt 15; $i++) {
            $slotName = if ($i -eq 0) { "OPEN" } else { "OPEN$i" }
            $existing = Get-ItemProperty -Path $OptionsPath -Name $slotName -ErrorAction SilentlyContinue
            if (-not $existing) {
                Set-ItemProperty -Path $OptionsPath -Name $slotName -Value "/R `"$AddinPath`""
                $slotFound = $true; break
            } elseif ($existing.$slotName -like "*$AddinName*") {
                $slotFound = $true; break
            }
        }
        if (-not $slotFound) {
            Write-Host "    [!] Khong co slot cho Office $ver" -ForegroundColor Yellow
        }
    }
}
Write-Host " -> Da hoan tat dang ky." -ForegroundColor Green

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "       CHUC MUNG! CAI DAT DA HOAN TAT THANH CONG!        " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "[*] Mo Excel va su dung ham: =VND(1000000)" -ForegroundColor White
Write-Host ""
pause
