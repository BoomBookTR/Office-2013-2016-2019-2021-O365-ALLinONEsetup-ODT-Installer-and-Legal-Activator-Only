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

:: .... your code start ....
setlocal
call :setESC

cls
set homedir=%~dp0
:: Äalçümçyor>> set batdir="%homedir%kurulumayar"
@echo off & setlocal enableextensions disabledelayedexpansion


::@echo off & setlocal enableextensions disabledelayedexpansion
::for %%P in (%CD%) do set "OriginalDrive=%%~dP"
::if not !OriginalDrive!==%SystemDrive% %SystemDrive%
REM Rest of the original code here

(call;)
title Office KUR
mode con lines=35 cols=50
(set lf=^
%= DO NOT DELETE =%
)
set ^"nl=^^^%lf%%lf%^%lf%%lf%^"
set ^"\n=^^^%lf%%lf%^%lf%%lf%^^"

cls
echo(Seáiniz...%nl%%\n%
  %ESC%[93m1.%ESC%[0m ProPlus %ESC%[7;73m2013%ESC%[0m Retail%nl%%\n%
  %ESC%[93m2.%ESC%[0m ProPlus %ESC%[7;73m2013%ESC%[0m Volume%nl%%\n%
  %ESC%[93m3.%ESC%[0m ProPlus %ESC%[101;93m2016%ESC%[0m Retail%nl%%\n%
  %ESC%[93m4.%ESC%[0m ProPlus %ESC%[101;93m2016%ESC%[0m Volume%nl%%\n%
  %ESC%[93m5.%ESC%[0m ProPlus %ESC%[7;73m2019%ESC%[0m Retail%nl%%\n%
  %ESC%[93m6.%ESC%[0m ProPlus %ESC%[7;73m2019%ESC%[0m Volume%nl%%\n%
  %ESC%[93m7.%ESC%[0m ProPlus %ESC%[101;93m2021%ESC%[0m Retail%nl%%\n%
  %ESC%[93m8.%ESC%[0m ProPlus %ESC%[101;93m2021%ESC%[0m Volume%nl%%\n%
  %ESC%[93m9.%ESC%[0m ProPlus %ESC%[7;73mO365%ESC%[0m Retail%nl%%\n%
  %ESC%[93m10.%ESC%[0m Lisans_Bilgisi_Al%nl%%\n%
  %ESC%[93m11.%ESC%[0m Lisans_Sil_ve_Aktif_Et%nl%%\n%
  %ESC%[93m12.%ESC%[0m Office-Legal-Activation-Script-Menulu-v1.0%nl%%\n%
  %ESC%[93m13.%ESC%[0m Office-IID-CID-Checker-Tool-Online%nl%%\n%
  %ESC%[93m0.%ESC%[0m ÄIK

:readKey
echo ==================================================
echo.
echo.

set /p "opt=%ESC%[93mSeáiniz?(1,2,3,4,5,6,7,8,9,10,11,0):%ESC%[0m "


if %opt% lss 10 (
set opt | findstr /ix "opt=[0123456789]" >nul || goto readKey
) else if %opt% lss 14 (
set opt | findstr /ix "opt=[0123456789][0123]" >nul || goto readKey
) else (
echo ==================================================
echo Bîyle bir seáenek bulunmamaktadçr. LÅtfen seáeneßinizi dÅzeltin.
goto readKey
)

if %opt% equ 0 goto end

for /f "tokens=1,2 delims=:" %%A in (
^"1:2013retail%nl%2:2013volume%nl%3:2016retail%nl%4:2016volume%nl%5:2019retail%nl%6:2019volume%nl%7:2021retail%nl%8:2021volume%nl%9:365retail%nl%10:legalactivationscriptmenulu%nl%11:checkiid^"
) do if %opt% equ %%A (

::BATCH BAûLANGIÄ
start "Office Kurulumu Baülçyor... BU PENCEREYò KURULUM BòTENE KADAR KAPATMAYIN..!" "%homedir%kurulumayar\start%%B.cmd"
)
::BATCH BòTòû

goto end
) 


:end
endlocal & goto :EOF




:setESC
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (
  set ESC=%%b
  exit /B 0
)
exit /B 0