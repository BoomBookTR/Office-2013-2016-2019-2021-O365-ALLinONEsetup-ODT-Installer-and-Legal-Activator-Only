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
ECHO Set UAC = CreateObject^("Shell.Application"^) > "%temp%\OEgetPrivileges.vbs"
ECHO UAC.ShellExecute "cmd", "/c ""!batchPath! !batchArgs!""", "", "runas", 1 >> "%temp%\OEgetPrivileges.vbs"
"%temp%\OEgetPrivileges.vbs" 
exit /B

:START
::Remove the elevation tag and set the correct working directory
IF '%1'=='ELEV' ( shift /1 )
cd /d %~dp0

:: .... your code start ....


title Office 2010 ProPlus Volume YÅkleniyor. KAPATMAYIN..!
echo ============================================================================&
echo #Proje: Office TEK TIKLAMA ile KURULUM&
echo ============================================================================&
echo.&
echo #DESTEKLEYEN VERSòYONLAR: Office 2010 ProPlus vb. &
echo.&
echo.& 
echo #KURULUM BòTENE KADAR KAPATMAYIN &
echo.& 
echo #KURULUM BòTENE KADAR KAPATMAYIN &
echo.& 
echo #KURULUM BòTENE KADAR KAPATMAYIN &
echo.& 
echo #KURULUM BòTENE KADAR KAPATMAYIN &
echo.& 
::setup.exe /download mysettings.xml

..\dosyalar\OfflineInstaller\2010\Office\Data\setup.exe /config "..\..\2010volume.xml"

::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))


:convertedilsinmi
set /P c=Retail to Volume iülemi geráekleütirilsin mi [E/H]?
if /I "%c%" EQU "E" goto :convertet
if /I "%c%" EQU "H" goto :secim
goto :convertedilsinmi

:convertet
set batdir=%~dp0
(if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office14")&
(if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office14")& cd %folder%

::CONVERT TO VOLUME
cd /d %batdir%..\dosyalar\OfflineInstaller\2010\Office\Convert_VL
for /f %%x in ('dir /b *.xrm-ms') do cscript %folder%\ospp.vbs /inslic:%%x
::for /f %%x in ('dir /b "%batdir%..\dosyalar\OfflineInstaller\2010\Office\Convert_VL\"*.xrm-ms') do cscript %folder%\ospp.vbs /inslic:"%batdir%..\dosyalar\OfflineInstaller\2010\Office\Convert_VL\"%%x
::CONVERT TO VOLUME


::cd /d %folder%
::cscript ospp.vbs /dstatus


:secim
set /P c=Aktivasyon iülemine devam etmek istedißine emin misin [E/H]?
if /I "%c%" EQU "E" goto :devamet
if /I "%c%" EQU "H" goto :devametme
goto :secim

:devamet
Echo YÅklÅ anahtar ile aktivasyon deneniyor.

:anahtar
if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"

cscript ospp.vbs /act | find /i "Product activation successful" && (echo.&echo ************************************************* &echo.&choice /n /c HE /m "Aktivasyon baüarçlç...Kapatmak istiyor musunuz? (E/H)" & if errorlevel 2 exit) || (echo Aktivasyon Baüarçsçz...! Yeniden baülançyor...) &

set /P c=Aktivasyon iülemine yeni anahtar ekleyerek devam edilecektir. Emin misin [E/H]?
if /I "%c%" EQU "E" goto :anahtaryukle
if /I "%c%" EQU "H" goto :anahtaryukleme

:anahtaryukle
cd %~dp0
start "Office Aktivasyonu Äalçüçyor!" ".\startlisansaktifet.cmd"


:anahtaryukleme
echo Aktivasyon ekranç kapatçlçyor...
timeout 5
exit

:devametme
Echo Kurulum òülemi Tamamlandç...
timeout 5
exit


::for /f "tokens=8" %%b in ('cscript ospp.vbs /dstatus ^| findstr /b /c:"Last 5"') do (cscript ospp.vbs /unpkey:%%b)

