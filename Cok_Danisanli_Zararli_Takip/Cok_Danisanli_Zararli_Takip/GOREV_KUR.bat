@echo off
cd /d "%~dp0"
schtasks /Create /F /SC DAILY /ST 08:30 /TN "ZararliTakipGunluk" /TR "\"%~dp0GUNCELLE.bat\""
pause
