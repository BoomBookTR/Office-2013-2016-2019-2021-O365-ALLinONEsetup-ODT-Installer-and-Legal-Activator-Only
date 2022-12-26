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

@echo off

setlocal
call :setESC

cls

title Office Lisans Bilgisi Al�n�yor...

if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
if exist "C:\Program Files\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office15"
if exist "C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office15"
if exist "C:\Program Files\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office14"
if exist "C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office14"

::for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles%\Microsoft Office\Office1%%a")
::if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (cd "%ProgramFiles(x86)%\Microsoft Office\Office1%%a"))

cscript //nologo ospp.vbs /dstatus

echo %ESC%[101mB�L�NMES� GEREKEN KODLAR %ESC%[0m
echo -----------------------------------------------------------
echo -----------------------------------------------------------
echo unpkey ile lisans silinir.
echo %ESC%[94mcscript //nologo ospp.vbs /unpkey:%ESC%[0m
echo -----------------------------------------------------------
echo inpkey ile lisans eklenir.
echo %ESC%[94mcscript //nologo ospp.vbs /inpkey:%ESC%[0m
echo -----------------------------------------------------------
echo Y�klenmi� lisans ile etkinle�tirir.
echo %ESC%[94mcscript //nologo ospp.vbs /act%ESC%[0m
echo -----------------------------------------------------------
echo Y�klenmi� lisans bilgilerini g�sterir.
echo %ESC%[94mcscript //nologo ospp.vbs /dstatus%ESC%[0m
echo -----------------------------------------------------------
echo �evrimd��� etkinle�tirme i�in Kurulum Kimli�ini (Installation ID) g�r�nt�ler.
echo %ESC%[94mcscript //nologo ospp.vbs /dinstid%ESC%[0m
echo -----------------------------------------------------------
echo �r�n�, kullan�c� taraf�ndan sa�lanan Onay Kimli�i (Confirmation ID) ile etkinle�tirir.
echo %ESC%[94mcscript //nologo ospp.vbs /actcid:value%ESC%[0m
echo -----------------------------------------------------------
echo Daha fazla bilgi;
echo https://docs.microsoft.com/en-us/deployoffice/vlactivation/tools-to-manage-volume-activation-of-office
echo -----------------------------------------------------------
echo -----------------------------------------------------------
cmd /k

:: BATCH dosyas�n�n kapanmas�n� istemiyorsan dosyan�n sonuna cmd /k ekle
:: cmd /k 


:setESC
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (
  set ESC=%%b
  exit /B 0
)
exit /B 0