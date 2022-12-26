@echo off
setlocal EnableDelayedExpansion

::net file to test privileges, 1>NUL redirects output, 2>NUL redirects errors
NET FILE 1>NUL 2>NUL
if '%errorlevel%' == '0' ( goto START ) else ( goto getPrivileges ) 

:getPrivileges
if '%1'=='ELEV' ( goto START )

set "batchPath=%~f0"
set "batchArgs=ELEV"

::Add quotes to the batch path, if needed
set "script=%0"
set script=%script:"=%
IF '%0'=='!script!' ( GOTO PathQuotesDone )
    set "batchPath=""%batchPath%"""
:PathQuotesDone

::Add quotes to the arguments, if needed.
:ArgLoop
IF '%1'=='' ( GOTO EndArgLoop ) else ( GOTO AddArg )
    :AddArg
    set "arg=%1"
    set arg=%arg:"=%
    IF '%1'=='!arg!' ( GOTO NoQuotes )
        set "batchArgs=%batchArgs% "%1""
        GOTO QuotesDone
        :NoQuotes
        set "batchArgs=%batchArgs% %1"
    :QuotesDone
    shift
    GOTO ArgLoop
:EndArgLoop

::Create and run the vb script to elevate the batch file
echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\OEgetPrivileges.vbs"
echo UAC.ShellExecute "cmd", "/c ""!batchPath! !batchArgs!""", "", "runas", 1 >> "%temp%\OEgetPrivileges.vbs"
"%temp%\OEgetPrivileges.vbs" 
exit /B

:START
::Remove the elevation tag and set the correct working directory
IF '%1'=='ELEV' ( shift /1 )
cd /d %~dp0

:: .... your code start ....


:: Global options for ospp.vbs
:: https://docs.microsoft.com/en-us/deployoffice/vlactivation/tools-to-manage-volume-activation-of-office


title Office ProPlus vb. EtkinleŸtirme Scripti
echo ============================================================================&
echo #Proje: Sadece Lisans Kodunu girerek otomatik aktivasyon iŸlemi sa§lanr.&
echo ============================================================================&
echo.&
echo #Desteklenen rnler: Office ProPlus vb.&
echo.&
echo.& 
title Office Lisans EtkinleŸtirme Scripti

:secim
set /P c=YklenmiŸ tm lisans anahtarlar silinecektir. Devam etmek istedi§ine emin misin [E/H]?
if /I "%c%" EQU "E" goto :devamet
if /I "%c%" EQU "H" goto :devametme
goto :secim

:devamet
if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"
::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))
echo.&
echo ============================================================================&


set /P c=YklenmiŸ tm lisans anahtarlar silinsin mi [E/H]?
if /I "%c%" EQU "E" goto :lisanssil
if /I "%c%" EQU "H" goto :keygir

:lisanssil
echo Deneme Srm veya YklenmiŸ Tm Lisans Anahtarlar Siliniyor...&
::Office Lisanslarn Sil
::for /f "tokens=8" %b in ('cscript ospp.vbs /dstatus ^| findstr /b /c:"Last 5"') do (cscript ospp.vbs /unpkey:%b)
for /f "tokens=8" %%b in ('cscript ospp.vbs /dstatus ^| findstr /b /c:"Last 5"') do (cscript ospp.vbs /unpkey:%%b)
::@For /F "Tokens=1* Delims=:" %%G In ('^""%__AppDir__%cscript.exe" "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /DStatus 2^> NUL ^| "%__AppDir__%find.exe" "Last 5"^"') Do @For %%I In (%%H) Do @If /I Not "XXXXX" == "%%I" "%__AppDir__%cscript.exe" "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" /UnPKey:%%I

rem/||(
cscript ospp.vbs /unpkey:6MWKP
cscript ospp.vbs /unpkey:BTDRB
cscript ospp.vbs /unpkey:DRTFM
cscript ospp.vbs /unpkey:WFG99
cscript ospp.vbs /unpkey:27GXM
)
:keygir
set /p LicenseKey=Lisans Anahtar Gir:
cscript ospp.vbs /inpkey:%LicenseKey%

cscript ospp.vbs /act | find /i "Product activation successful" && (echo.&echo ************************************************* &echo.&choice /n /c HE /m "Aktivasyon baŸarl...Kapatmak istiyor musunuz? (E/H)" & if errorlevel 2 exit) || (echo Aktivasyon BaŸarsz...! Yeniden baŸlanyor...) &

set /P d=Aktivasyon iŸlemine yeni anahtar ekleyerek devam edilecektir. Emin misin [E/H]?
if /I "%d%" EQU "E" goto :devamet
if /I "%d%" EQU "H" goto :devametme

:devametme
Echo Kurulum ˜Ÿlemi Tamamland...
timeout 5
exit

echo.&
echo ============================================================================&

rem/||(
cscript ospp.vbs /act | find /i "successful" && (echo.&echo ************************************************* &echo.&choice /n /c YN /m "Do you want to restart your PC now [Y,N]?" & if errorlevel 2 exit) || (echo There is an error)
shutdown.exe /r /t 00
)
rem/||(
cscript ospp.vbs /act | find /i "successful" && (echo.&echo ************************************************* &echo.&choice /n /c YN /m "Would you like to visit my blog [Y,N]?" & if errorlevel 2 exit) || (echo There is an error)
explorer "http://website.com"&goto halt
)

:: €OKLU REM SATIRLARI OKUNMAZ
rem/||(
::Office Yolunu Bul
for %a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%a")
if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%a"))

for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))


::Office Lisanslarn Sil
for /f "tokens=8" %b in ('cscript ospp.vbs /dstatus ^| findstr /b /c:"Last 5"') do (cscript ospp.vbs /unpkey:%b)

::2019 Convert Retail to Volume
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
for /f %i in ('dir /b ..\root\Licenses16\ProPlus2019VL_MAK*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"


::2013 Convert Retail to Volume
for /f %x in ('dir /b "..\..\Microsoft Office 15\root\Licenses\"ProPlusVL*.xrm-ms') do cscript ospp.vbs /inslic:"..\..\Microsoft Office 15\root\Licenses\"%x


)
:: €OKLU REM SATIRLARI OKUNMAZ


:: GOTO SATIRLARI OKUNMAZ
goto :start

€ok satrl bir yorum blo§u buraya gidebilir.
| > gibi ”zel karakterler de i‡erebilir.
cscript ospp.vbs /dti <<<<<<<<<<<<<Offline Phone Activation
:start

:: GOTO SATIRLARI OKUNMAZ