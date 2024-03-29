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

title Office 2013 ProPlus Volume Y�kleniyor. KAPATMAYIN..!
echo ============================================================================&
echo #Proje: Office TEK TIKLAMA ile KURULUM&
echo ============================================================================&
echo.&
echo #DESTEKLEYEN VERS�YONLAR: Office 2013 ProPlus vb. &
echo.&
echo.& 
echo #KURULUM B�TENE KADAR KAPATMAYIN &
echo.& 
echo #KURULUM B�TENE KADAR KAPATMAYIN &
echo.& 
echo #KURULUM B�TENE KADAR KAPATMAYIN &
echo.& 
echo #KURULUM B�TENE KADAR KAPATMAYIN &
echo.& 

::homedir=\ALLinONEsetup\kurulumayar
set homedir="%~dp0"

::filedir=\ALLinONEsetup
::echo %homedir%
set filedir=%homedir%..
::echo %filedir%
set onlinesetupdir=%filedir%\dosyalar\OnlineInstaller\2013
::echo %onlinesetupdir%
set offlinesetupdir=%filedir%\dosyalar\OfflineInstaller\2013
::echo %offlinesetupdir%


:indir-me
::choice /n /c CY /m "Office yaz�l�m� �evrimi�i (C) mi yerel dosyalarla (Y) m� y�klensin [C/Y]?" & if errorlevel 2 goto :cevrimdisi


:cevrimici
cd %onlinesetupdir%
::echo %cd%
if exist Office rd /s /q Office
echo Dosyar�n indirilme s�resi internet h�z�n�za g�re de�i�iklik g�sterecektir. 
echo Dosyalar indirildi�inde kurulum otomatik olarak ba�layacakt�r.

setup.exe /download 2013volume.xml
setup.exe /configure 2013volume.xml
::..\dosyalar\OnlineInstaller\2013\setup.exe /download "..\dosyalar\OnlineInstaller\2013\2013volume.xml"
goto :secim

:cevrimdisi
cd %offlinesetupdir%
::echo %cd%
setup.exe /configure 2013volume.xml
::..\dosyalar\OfflineInstaller\2013\setup.exe /configure "..\dosyalar\OfflineInstaller\2013\2013volume.xml"


:secim
set /P c=Aktivasyon i�lemine devam etmek istedi�ine emin misin [E/H]?
if /I "%c%" EQU "E" goto :devamet
if /I "%c%" EQU "H" goto :devametme
goto :secim

:devamet
Echo Haz�r oldu�unda aktivasyon i�lemine ge�ilecektir.

:anahtar

:: CONVERT TO VOLUME AFTER HERE
if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"
::(if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office15")&
::(if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office15")& cd %folder%


:convertedilsinmi
set /P c=Retail to Volume i�lemi ger�ekle�tirilsin mi [E/H]?
if /I "%c%" EQU "E" goto :convertet
if /I "%c%" EQU "H" goto :etkinlestir
goto :convertedilsinmi


:convertet

::CONVERT TO VOLUME 
for /f %%x in ('dir /b "..\..\Microsoft Office 15\root\Licenses\"ProPlusVL*.xrm-ms') do cscript ospp.vbs /inslic:"..\..\Microsoft Office 15\root\Licenses\"%%x
::CONVERT TO VOLUME


:etkinlestir
Echo Y�kl� anahtar ile aktivasyon deneniyor.

cscript ospp.vbs /act | find /i "Product activation successful" && (echo.&echo ************************************************* &echo.&choice /n /c HE /m "Aktivasyon ba�ar�l�...Kapatmak istiyor musunuz? (E/H)" & if errorlevel 2 exit) || (echo Aktivasyon Ba�ar�s�z...! Yeniden ba�lan�yor...) &

:sec
set /P a=Aktivasyon i�lemine yeni anahtar ekleyerek devam edilecektir. Emin misin [E/H]?
if /I "%a%" EQU "E" goto :anahtaryukle
if /I "%a%" EQU "H" goto :anahtaryukleme
goto :sec

:anahtaryukle
cd %~dp0
start "Office Aktivasyonu �al���yor!" ".\startlisansaktifet.cmd"


:anahtaryukleme
echo Aktivasyon ekran� kapat�l�yor...
exit

:devametme
Echo Kurulum ��lemi Tamamland�...
exit