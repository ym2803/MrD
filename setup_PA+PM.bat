@echo off
setlocal EnableDelayedExpansion
title Installer & Setup Task Scheduler PA+PM
echo ============================================================
echo   Download, Ekstrak, dan Pasang Task Scheduler + Autostart PA+PM
echo ============================================================
echo.

:: Cek Hak Akses Administrator
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Klik kanan file ini lalu pilih "Run as administrator"!
    pause
    exit /b 1
)

:: Konfigurasi Path & Parameter (SESUAIKAN NAMA EXE KEDUA DI BARIS EXE2)
set "FILE_ID=1GF0LG6QZxjqGTapubDBWB1rHDxv45OW7"
set "TARGET_DIR=D:\MP\PA+PM"
set "RAR_FILE=%TARGET_DIR%\PA+PM.rar"
set "EXE1=%TARGET_DIR%\PrintAuto.exe"
set "EXE2=%TARGET_DIR%\ProdMast.exe"

:: 1. Buat Folder Tujuan D:\MP\PA+PM
echo [1/6] Membuat folder %TARGET_DIR% ...
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

:: 2. Download File RAR dari Google Drive (Bypass konfirmasi file besar >25MB)
echo [2/6] Mengunduh PA+PM.rar dari Google Drive (~72 MB)...
set "PS1=%TEMP%\dl_pa_pm.ps1"
>"%PS1%" echo param([string]$Url,[string]$Out)
>>"%PS1%" echo $ErrorActionPreference='Stop'
>>"%PS1%" echo [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
>>"%PS1%" echo $fs=$null
>>"%PS1%" echo try {
>>"%PS1%" echo $req=[System.Net.HttpWebRequest]::Create($Url)
>>"%PS1%" echo $req.UserAgent='Mozilla/5.0'
>>"%PS1%" echo $res=$req.GetResponse()
>>"%PS1%" echo $total=$res.ContentLength
>>"%PS1%" echo $in=$res.GetResponseStream()
>>"%PS1%" echo $fs=[System.IO.File]::Create($Out)
>>"%PS1%" echo $buf=New-Object byte[] 81920
>>"%PS1%" echo $got=0
>>"%PS1%" echo $last=0
>>"%PS1%" echo $sw=[System.Diagnostics.Stopwatch]::StartNew()
>>"%PS1%" echo while(($n=$in.Read($buf,0,$buf.Length)) -gt 0){
>>"%PS1%" echo $fs.Write($buf,0,$n)
>>"%PS1%" echo $got+=$n
>>"%PS1%" echo if($sw.ElapsedMilliseconds - $last -ge 200){
>>"%PS1%" echo $last=$sw.ElapsedMilliseconds
>>"%PS1%" echo $mb=[math]::Round($got/1MB,1)
>>"%PS1%" echo if($total -gt 0){
>>"%PS1%" echo $pct=[int](($got*100)/$total)
>>"%PS1%" echo $bar=('#'*[int]($pct/5)).PadRight(20,'-')
>>"%PS1%" echo Write-Host -NoNewline ("`r      [$bar] $pct%% - $mb MB   ")
>>"%PS1%" echo } else {
>>"%PS1%" echo Write-Host -NoNewline ("`r      Mengunduh... $mb MB   ")
>>"%PS1%" echo }
>>"%PS1%" echo }
>>"%PS1%" echo }
>>"%PS1%" echo $fs.Close()
>>"%PS1%" echo $in.Close()
>>"%PS1%" echo $mb=[math]::Round($got/1MB,1)
>>"%PS1%" echo Write-Host ("`r      [" + ('#'*20) + "] 100%% - $mb MB   ")
>>"%PS1%" echo } catch {
>>"%PS1%" echo if($fs){$fs.Close()}
>>"%PS1%" echo if(Test-Path $Out){Remove-Item $Out -Force}
>>"%PS1%" echo exit 1
>>"%PS1%" echo }
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Url "https://drive.usercontent.google.com/download?id=%FILE_ID%&export=download&confirm=t" -Out "%RAR_FILE%"
del /f /q "%PS1%" >nul 2>&1

if not exist "%RAR_FILE%" (
    echo [ERROR] Gagal mengunduh file PA+PM.rar. Periksa koneksi internet!
    pause
    exit /b 1
)

:: 3. Ekstrak File RAR ke D:\MP\PA+PM (Otomatis deteksi WinRAR / 7-Zip)
echo [3/6] Mengekstrak PA+PM.rar ke %TARGET_DIR% ...
set "WINRAR="
if exist "%ProgramFiles%\WinRAR\WinRAR.exe" set "WINRAR=%ProgramFiles%\WinRAR\WinRAR.exe"
if exist "%ProgramFiles(x86)%\WinRAR\WinRAR.exe" set "WINRAR=%ProgramFiles(x86)%\WinRAR\WinRAR.exe"
set "SEVENZIP="
if exist "%ProgramFiles%\7-Zip\7z.exe" set "SEVENZIP=%ProgramFiles%\7-Zip\7z.exe"
if exist "%ProgramFiles(x86)%\7-Zip\7z.exe" set "SEVENZIP=%ProgramFiles(x86)%\7-Zip\7z.exe"

if defined WINRAR (
    "%WINRAR%" x -o+ "%RAR_FILE%" "%TARGET_DIR%\"
) else if defined SEVENZIP (
    "%SEVENZIP%" x "%RAR_FILE%" -o"%TARGET_DIR%" -y >nul
) else (
    tar -xf "%RAR_FILE%" -C "%TARGET_DIR%"
)

if %errorlevel% neq 0 (
    echo [ERROR] Gagal mengekstrak PA+PM.rar! Pastikan WinRAR atau 7-Zip terpasang di PC.
    pause
    exit /b 1
)

:: 4. Hapus File RAR setelah berhasil diekstrak
echo [4/6] Menghapus file installer PA+PM.rar ...
del /f /q "%RAR_FILE%" >nul 2>&1

:: 5. Buat Task Scheduler (Mulai 07:00 AM)
echo [5/6] Mendaftarkan jadwal ke Task Scheduler...

:: Task 1: Jalankan PrintAuto.exe setiap hari pukul 07:00 AM
schtasks /create /tn "Shutdown Windows1" /tr "\"%EXE1%\"" /sc daily /st 07:00 /rl highest /f >nul 2>&1

:: Task 2: Jalankan Aplikasi Kedua setiap hari pukul 07:00 AM
schtasks /create /tn "Shutdown Windows2" /tr "\"%EXE2%\"" /sc daily /st 07:00 /rl highest /f >nul 2>&1

:: Hapus Watchdog lama (jika pernah terpasang) karena diganti Autostart
schtasks /delete /tn "MRDIY_PrintAuto_Watchdog" /f >nul 2>&1

:: 6. Install Autostart PrintAuto.exe (terbuka otomatis setiap login Windows)
echo [6/6] Memasang Autostart PrintAuto.exe ...
if not exist "%EXE1%" (
    echo [ERROR] PrintAuto.exe tidak ditemukan di %TARGET_DIR%
    pause
    exit /b 1
)
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v PrintAuto /t REG_SZ /d "\"%EXE1%\"" /f >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Gagal memasang Autostart PrintAuto.exe!
    pause
    exit /b 1
)

echo.
echo [SELESAI] Semua proses berhasil dijalankan!
echo - Folder instalasi : %TARGET_DIR%
echo - Jadwal 07:00 AM  : PrintAuto.exe dan Aplikasi Kedua
echo - Autostart        : PrintAuto.exe terbuka otomatis setiap login Windows
echo.
echo Jendela ini akan tertutup otomatis dalam 10 detik...
timeout /t 10 /nobreak >nul
exit /b 0
