@echo off

title SENCOR-WFH Network Diagnostics
mode con: cols=100 lines=29

REM INITIALIZE
set vpn1A=203.147.107.250
set vpn1B=210.176.35.185
set vpn1Aname=telstra_sencor
set vpn1Bname=telstra_border
set vpn1Adomain=VPN.SENCOR.NET

set vpn2A=116.50.227.91
set vpn2B=116.50.227.65
set vpn2Aname=etpi_sencor
set vpn2Bname=etpi_border
set vpn2Adomain=ETPI-V.SENCOR.NET

set domain=sencor.net
set noddomain=sencor-mnl.net
set server=postoffice
set googDNS=8.8.8.8
set spdydomain=www.speedtest.net
set spdyconfig=speedtest-config
set spdyserver=speedtest-servers
set logonserver=192.168.1.61
set syncdir=%userprofile%\.sencor
set tempdir=%syncdir%\temp
set redo=false
set incr=1
set option=%1
set customdate=%date:~10,4%%date:~4,2%%date:~7,2%
set programname=syncore
set appname=nog.wfhmonitoring

if not [%option%]==[] goto :OPTIONS

:START
REM MAIN
for %%x in (%googDNS% %logonserver%) do (
    ping %%x -n 1 -w 1000
    cls
    if errorlevel 1 goto :offline
)

if not exist %tempdir% mkdir %tempdir%
if exist %userprofile%\.sencor\info.cmd GOTO :RELOADINFO

:: openfiles > NUL 2>&1 
:: if NOT %errorlevel% EQU 0 goto :NOTADMIN

:LOGIN
cls
echo.
echo	Welcome to SENCOR-NOD Network Diagnostic Tool. 
echo.
echo	Please key in your information honestly. This is a one time
echo	registration process. All information gathered is part of
echo	our troubleshooting and reporting.
echo.
echo. 
echo Name (ex. James C. Sacuan):
set /p FULLNAME=	
echo.
echo Department (ex. NOD):
set /p DEPT=
echo.
echo Username (ex. jecsacuan):
set /p USERNAME=SENCOR\
echo.
echo Group (leave blank if not applicable):
set /p GROUP=
echo.
echo ID Number (15-089):
set /p IDNO=
echo.
echo Main ISP at home (Globe/PLDT/Sky/Converge/others):
set /p ISP=
echo.

:REVIEW
cls
echo.
echo Please review your information:
echo.
echo Name:		%FULLNAME%
echo Username:	%USERNAME%
echo Department:	%DEPT%
echo Group:		%GROUP%
echo ID No:		%IDNO%
echo ISP:		%ISP%
echo.
echo Is above information correct? (yes/no):
echo *type-in yes or no (case-sensitive)
set /p REVOPT=
if %REVOPT% EQU no GOTO :LOGIN
if %REVOPT% EQU yes GOTO :SAVINGINFO
GOTO :REVIEW
pause

:OFFLINE
echo.
echo.
echo	You're not connected to the Internet or to GlobalProtect.
echo.
echo	Kindly check your connection settings and try running
echo	this utility tool again. Thanks.
echo.
echo.
pause
goto :end

:NOTADMIN
echo.
echo.
echo	You are not running as administrator.
echo.
echo	Right-click on this program and select 'Run as administrator'
echo	to run this tool in elevated mode.
echo.
echo.
pause
goto :end

:OPTIONS
if %option%==--uninstall (
    openfiles > NUL 2>&1 
    if NOT %errorlevel% EQU 0 goto :NOTADMIN
    cls
    echo.
    ECHO    UNINSTALLING...
    echo.
    :: del %windir%\System32\%programname%.bat
    DEL /F/Q/S %syncdir%\*.* > NUL
    RMDIR /Q/S %tempdir%
    RMDIR /Q/S %syncdir%
    timeout /t 2
    cls
    echo.
    echo.
    echo	%programname% has been uninstalled.
    echo.
    echo	Thanks for using!
    echo.
    echo.
    timeout /t 5
    goto :end
)
if %option%==--clear (
    cls
    echo.
    ECHO    Clearing Saved Data...
    echo.
    del %syncdir%\info.cmd
    timeout /t 2
    cls
    echo.
    ECHO    Data cleared. Please wait as we redirect you to
    echo    registration screen.
    echo.
    timeout /t 5
    goto :start
)
cls
echo.
ECHO INVALID OPTION
echo.
echo Usage: %programname% [--uninstall  OR  --clear]
echo.
echo Options:
echo    --uninstall     Delete all saved data and files associated with
echo                    this program. (Should be Run as Administrator)
echo    --clear         Command to clear saved data and redo registration
echo                    process.
echo.
pause
goto :end

:SAVINGINFO
REM SAVE INFO
@echo @ECHO OFF > "%syncdir%\info.cmd" 
@echo set FULLNAME=%FULLNAME% > "%syncdir%\info.cmd"
@echo set USERNAME=%USERNAME% >> "%syncdir%\info.cmd"
@echo set GROUP=%GROUP% >> "%syncdir%\info.cmd"
@echo set DEPT=%DEPT% >> "%syncdir%\info.cmd"
@echo set IDNO=%IDNO% >> "%syncdir%\info.cmd"
@echo set ISP=%ISP% >> "%syncdir%\info.cmd"
@echo set DATEREGISTERED=%DATE% %TIME% >> "%syncdir%\info.cmd"
echo.
:: copy %~dp0%~n0%~x0 %windir%\System32\%programname%.bat >NUL
echo saved.
pause


:RELOADINFO
REM RELOAD INFO
call "%syncdir%\info.cmd"
setlocal enabledelayedexpansion

for /f "tokens=* delims= " %%a in ("%USERNAME%") do set USERNAME=%%a
for /l %%a in (1,1,100) do if "!USERNAME:~-1!"==" " set USERNAME=!USERNAME:~0,-1!

for %%x in (vbs txt json png) do (
    if exist "%tempdir%\*.%%x" del "%tempdir%\*.%%x"
)
if exist "%syncdir%\sent.txt" del "%syncdir%\sent.txt"
if exist "%syncdir%\*.vbs" del "%syncdir%\*.vbs"
if exist "%syncdir%\*.ps1" del "%syncdir%\*.ps1"

:GETPENDING
if exist "%syncdir%\pack*.zip" (
    REM REDO SENDING

    cd %syncdir%
    for /f %%A in ('dir pack*.zip ^| find "File(s)"') do set redocnt=%%A

    set redo=true  
    set xdate=""
    set xtime=""
    set ydate=""
    set ytime=""
    
    setlocal ENABLEDELAYEDEXPANSION
    
    set /a c=0
    for /f %%Y in ('DIR /B /O:-D /A:-D pack*.zip') do (
        FOR %%i IN ("%%Y") DO (
            set filename=%%~ni
            set fileextn=%%~xi
        ) 
        set redoyear=!filename:~5,4!
        set redomonth=!filename:~9,2!
        set redoday=!filename:~11,2!
        set redovsn=!filename:~14,2!
        set redoname=!redoyear!!redomonth!!redoday!-!redovsn!
        set /a c=c+1

        goto :DIAGSTEP7
    )
    pause
    endlocal 
    goto :eof
) ELSE (
    REM START FRESH
    set redo=false
    if exist "%syncdir%\log-*.txt" del "%syncdir%\log-*.txt"
)

:DIAGNOSTICS
REM START DIAGNOSTIC TEST
SETLOCAL EnableDelayedExpansion

if not [!incr!]==[] (
    set incr=!incr!
)

if %incr% lss 10 set incr=0%incr%
set id=%customdate%-%incr%

FOR /F "skip=1" %%A IN ('WMIC OS GET LOCALDATETIME') DO (SET "t=%%A" & GOTO break_1)
:break_1

SET "m=%t:~10,2%" & SET "h=%t:~8,2%" & SET "d=%t:~6,2%" & SET "z=%t:~4,2%" & SET "y=%t:~0,4%"
IF !h! GTR 11 (SET /A "h-=12" & SET "ap=P" & IF "!h!"=="0" (SET "h=00") ELSE (IF !h! LEQ 9 (SET "h=0!h!"))) ELSE (SET "ap=A")

set xdate=%z%-%d%-%y%
set xtime=%h%:%m% %ap%M

powershell -command "wget https://api.ipify.org/ -OutFile '%tempdir%/ip.txt'"
set /p ipadd=<%userprofile%/.sencor/temp/ip.txt

set ydate=""
set ytime=""
set PANGPlog="%localappdata%\Palo Alto Networks\GlobalProtect\PanGPA.log"
set count=0

:DIAGBODY
cls
echo.
echo ******************************************************************
echo.
echo    Welcome, %fullname%^^!
echo.
echo    This is a three part series test and would take around
echo    5-8 minutes. Please DON'T CLOSE THIS WINDOW until our
echo    network diagnostic is complete.
echo.
echo    You may minimize this console and continue your work.   
echo.
echo ******************************************************************
echo.
if %redo%==false (
    echo    Time started: %xdate% %xtime%
    if %ydate% neq "" echo    Time finished: %ydate% %ytime%
    echo.
) 
echo.
echo    0. Information Gathering
if %count%==0 GOTO :BUSY
echo.
echo    1. PING Test
echo        a. %vpn1Adomain%
if %count%==1 GOTO :BUSY
echo        b. %vpn2Adomain%
if %count%==2 GOTO :BUSY
echo.
echo    2. TRACEROUTE Test
echo        a. %vpn1Adomain%
if %count%==3 GOTO :BUSY
echo        b. %vpn2Adomain%
if %count%==4 GOTO :BUSY
if %count%==5 GOTO :BUSY
echo.
echo    A. Generating report...
if %count%==6 GOTO :BUSY
echo.
echo    B. Sending report...
echo.
GOTO :DIAGSTEP7

:BUSY
echo.
powershell write-host -foregroundcolor YELLOW "[%count%/7] Working... Please Wait."
echo.
goto :DIAGSTEP%COUNT%


:DIAGSTEP0
REM 0.A INFORMATION GATHERING
powershell -command "wget http://ip-api.com/json/%ipadd% -OutFile '%tempdir%/net-%username%.json'"
del "%tempdir%\ip.txt"
if exist %PANGPlog% copy %PANGPlog% "%tempdir%"
if exist %PANGPlog% del %PANGPlog%
ipconfig > "%tempdir%\ipconfig-%username%.txt"
set clip="%syncdir%\clip.ps1"
type NUL > %clip%
@echo > %clip% [void][reflection.assembly]::loadwithpartialname("system.windows.forms")
@echo >>%clip% [system.windows.forms.sendkeys]::sendwait('{PRTSC}')
@echo >>%clip% Get-Clipboard -Format Image ^| ForEach-Object -MemberName Save -ArgumentList '%tempdir%\%random%.png'
powershell -executionpolicy remotesigned -File %clip%
del %clip%
set count=1

REM 0.B PREPARING PING FILE
set ping="%syncdir%\ping.ps1"
type NUL > %ping%
@echo > %ping% $hostname = $args[0]
@echo >> %ping% $descript = $args[1]
@echo. >> %ping% 
@echo >> %ping% ping.exe $hostname -n 15 ^|Foreach{"{0} - {1}" -f (Get-Date),$_} ^| Out-File -Filepath "%tempdir%\ping_$descript-%username%.txt"
GOTO :DIAGBODY

:DIAGSTEP1
REM 1.A PING - VPN
powershell %ping% %vpn1B% %vpn1Bname%
powershell %ping% %vpn1A% %vpn1Aname%
set count=2
GOTO :DIAGBODY

:DIAGSTEP2
REM 1.B PING - ETPI
powershell %ping% %vpn2B% %vpn2Bname%
powershell %ping% %vpn2A% %vpn2Aname%
del %ping%
set count=3
GOTO :DIAGBODY

:DIAGSTEP3
REM 2.A TRACEROUTE - VPN
tracert %vpn1B% > "%tempdir%\tracert_%vpn1Bname%-%username%.txt"
tracert %vpn1A% > "%tempdir%\tracert_%vpn1Aname%-%username%.txt"
set count=4
GOTO :DIAGBODY

:DIAGSTEP4
REM 2.B TRACEROUTE - ETPI
tracert %vpn2A% > "%tempdir%\tracert_%vpn2Aname%-%username%.txt"
tracert %vpn2B% > "%tempdir%\tracert_%vpn2Bname%-%username%.txt"
set count=5
GOTO :DIAGBODY

:DIAGSTEP5
REM 0.C WRAPPING THINGS UP
SETLOCAL EnableDelayedExpansion

FOR /F "skip=1" %%A IN ('WMIC OS GET LOCALDATETIME') DO (SET "t=%%A" & GOTO break_1)
:break_1

SET "m=%t:~10,2%" & SET "h=%t:~8,2%" & SET "d=%t:~6,2%" & SET "z=%t:~4,2%" & SET "y=%t:~0,4%"
IF !h! GTR 11 (SET /A "h-=12" & SET "ap=P" & IF "!h!"=="0" (SET "h=00") ELSE (IF !h! LEQ 9 (SET "h=0!h!"))) ELSE (SET "ap=A")

set ydate=%z%-%d%-%y%
set ytime=%h%:%m% %ap%M

echo We're done testing. Please wait as we attempt to send the files.
timeout /t 10
set count=6
GOTO :DIAGBODY

:DIAG8
REM 3.A SPEEDTEST
set %spdy%="%syncdir%\speed.ps1"
type nul > %spdy%

REM 3.B GET USER COORDINATES
REM 3.C GET LIST OF ISP
@echo >>%spdy% $objXmlHttp1 = New-Object -ComObject MSXML2.ServerXMLHTTP
@echo >>%spdy% $objXmlHttp1.Open("GET", "%spdydomain%/%spdyserver%.php", $False)
@echo >>%spdy% $objXmlHttp1.Send()
@echo. >>%spdy%
@echo >>%spdy% [xml]$ServerList = $objXmlHttp1.responseText
@echo >>%spdy% $cons = $ServerList.settings.servers.server
@echo. >>%spdy%

REM 3.D GET NEARBY ISP
@echo >>%spdy% foreach($val in $cons) { 
@echo >>%spdy%     $R = 6371;
@echo >>%spdy%     [float]$dlat = ([float]$oriLat - [float]$val.lat) * 3.14 / 180;
@echo >>%spdy%     [float]$dlon = ([float]$oriLon - [float]$val.lon) * 3.14 / 180;
@echo >>%spdy%     [float]$a = [math]::Sin([float]$dLat/2) * [math]::Sin([float]$dLat/2) + [math]::Cos([float]$oriLat * 3.14 / 180 ) * [math]::Cos([float]$val.lat * 3.14 / 180 ) * [math]::Sin([float]$dLon/2) * [math]::Sin([float]$dLon/2);
@echo >>%spdy%     [float]$c = 2 * [math]::Atan2([math]::Sqrt([float]$a ), [math]::Sqrt(1 - [float]$a));
@echo >>%spdy%     [float]$d = [float]$R * [float]$c;
@echo. >>%spdy% 
@echo >>%spdy%     $ServerInformation +=
@echo >>%spdy% @([pscustomobject]@{Distance = $d; Country = $val.country; Sponsor = $val.sponsor; Url = $val.url })
@echo >>%spdy% }

:DIAGSTEP6
REM GENERATING REPORT
type NUL > "%syncdir%\log-%id%.txt"
set log="%syncdir%\log-%id%.txt"
@echo >%log% Name: %FULLNAME%
@echo >>%log% Department: %DEPT%
@echo >>%log% Group: %GROUP%
@echo >>%log% Employee ID: %IDNO%
@echo. >>%log%
@echo >>%log% ISP: %ISP%
@echo >>%log% IP:  %ipadd%
@echo. >>%log%
@echo >>%log% Started: %xdate% %xtime%
@echo >>%log% Finished: %ydate% %ytime%
@echo. >> %log%
@echo. >> %log%
@echo PING TO %vpn1Adomain% >> %log% 
@echo >> %log% ----------------------------------
type "%tempdir%\ping_%vpn1Aname%-%username%.txt" >> %log%
@echo. >> %log%
@echo. >> %log%
@echo. >> %log%
@echo PING TO %vpn2Adomain% >> %log% 
@echo >> %log% ----------------------------------
type "%tempdir%\ping_%vpn2Aname%-%username%.txt" >> %log%
@echo. >> %log%
@echo. >> %log%
@echo >> %log% TRACEROUTE TO %vpn1Adomain%
@echo ---------------------------------- >> %log%
type "%tempdir%\tracert_%vpn1Aname%-%username%.txt" >> %log%
@echo. >> %log%
@echo. >> %log%
@echo. >> %log%
@echo >> %log% TRACEROUTE TO %vpn2Adomain%
@echo ---------------------------------- >> %log%
type "%tempdir%\tracert_%vpn2Aname%-%username%.txt" >> %log%
@echo. >> %log%
@echo. >> %log%
@echo. >> %log%
@echo >> %log% ADDITIONAL INFORMATION
@echo >> %log% ----------------------------------
type "%tempdir%\net-%username%.json" >> %log%
del "%tempdir%\net-%username%.json"

REM COMPRESSING REPORT
type NUL > "%syncdir%\_zipIt.vbs"
set zipit="%syncdir%\_zipIt.vbs"
echo >"%zipit%" Set objArgs = WScript.Arguments
echo >> %zipit% InputFolder = objArgs(0)
echo >> %zipit% ZipFile = objArgs(1)
echo >> %zipit% CreateObject("Scripting.FileSystemObject").CreateTextFile(ZipFile, True).Write "PK" ^& Chr(5) ^& Chr(6) ^& String(18, vbNullChar)
echo >> "%zipit%" Set objShell = CreateObject("Shell.Application")
echo >> "%zipit%" Set source = objShell.NameSpace(InputFolder).Items
echo >> %zipit% objShell.NameSpace(ZipFile).CopyHere(source)
echo >> %zipit% wScript.Sleep 2000
CScript %zipit% "%tempdir%" "%syncdir%\pack-%id%.zip"
del %zipit%
del "%tempdir%\*.txt"
del "%tempdir%\*.json"
del "%tempdir%\*.png"
if exist "%tempdir%\PanGPA.log" del "%tempdir%\PanGPA.log"
set count=7
GOTO :DIAGBODY

:DIAGSTEP7
REM SET MESSAGE HEADER
set from=notify
set cc=sencor.datacomm
set subj="SENCOR WFH NET DIAGNOSTICS - %username%"
if %redo%==true (
    cls
    echo.
    echo We found %redocnt% pending job. We will attempt to send this first.
    echo.
    echo.
    powershell write-host -foregroundcolor YELLOW "[%c%/%redocnt%] Processing !filename!. Please Wait."
    set id=!redoyear!!redomonth!!redoday!-!redovsn!
    if exist "%syncdir%\*.vbs" del "%syncdir%\*.vbs"
    if exist "%syncdir%\log-!id!.txt" (
        set log="%syncdir%\log-!id!.txt"
    ) else (
        echo This is an unsent log dated !redoyear!-!redomonth!-!redoday! >> "%syncdir%\log-!id!.txt"
        set log="%syncdir%\log-!id!.txt"
    )
    timeout /t 5
)

REM SENDING INFORMATION
call :MESSAGE "%syncdir%\mail.vbs"
call :SEND
if exist "%syncdir%\sent.txt" (
   cls
   echo.
   powershell write-host -foregroundcolor green "Logs have been sent to Datacomm [NOD]. Thanks."
   echo.
   del "%syncdir%\sent.txt"
   del "%syncdir%\*.vbs"
   timeout /t 10
   if %redo%==true (
       if %redocnt% gtr 1 (
           goto :GETPENDING
       ) else (
           goto :TRYAGAIN
       )
   ) else goto :END
) else (
   cls
   echo.
   powershell write-host -foregroundcolor red "We are unable to send the logs.`nWe will send it out on your next network diagnostics."
   echo.
   del "%syncdir%\*.vbs"
   timeout /t 10
   if %redo%==true (
       goto :TRYAGAIN
   ) else goto :END
)

REM LAST TRY
:TRYAGAIN
SETLOCAL EnableDelayedExpansion

set redo=false
if %customdate%==!redoyear!!redomonth!!redoday! (
    set /a incr=!redovsn:~1,1!+1
)
goto :DIAGNOSTICS

:SEND
cscript.exe /nologo "%blat%"
goto :EOF

:MESSAGE
set "blat=%~1"
if not [!id!]==[] (
    set fileattach="%syncdir%\pack-!id!.zip"
    set id=!id!
) else set fileattach="%syncdir%\pack-%id%.zip"

del %blat% 2>nul
set cdoSchema=http://schemas.microsoft.com/cdo/configuration
echo >>%blat% Set objMail = CreateObject("CDO.Message")
echo >>%blat% Set objConf = CreateObject("CDO.Configuration")
echo >>%blat% Set objFSO  = CreateObject("Scripting.FileSystemObject")
echo >>%blat% Set objFlds = objConf.Fields
echo >>%blat% Set f = objFSO.OpenTextFile(%log%, 1)
echo. >>%blat% 
echo >>%blat% objFlds.Item("%cdoSchema%/sendusing") = 2 'cdoSendUsingPort
echo >>%blat% objFlds.Item("%cdoSchema%/smtpserver") = "%server%.%domain%"
echo >>%blat% objFlds.Item("%cdoSchema%/smtpserverport") = 587
echo >>%blat% objFlds.Update
echo. >>%blat% 
echo >>%blat% objMail.Configuration = objConf
echo >>%blat% objMail.From = "%from%@%noddomain%"
echo >>%blat% objMail.To = "%appname%@%domain%"
echo >>%blat% objMail.CC = "%cc%@gmail.com"
echo >>%blat% objMail.Subject = %subj%
echo >>%blat% objMail.TextBody = f.ReadAll
echo >>%blat% f.Close
echo >>%blat% objMail.AddAttachment %fileattach%
echo >>%blat% objMail.Send
echo. >>%blat% 
echo >>%blat% file="" 
echo >>%blat% If err.number = 0 then 
echo >>%blat%    file="%syncdir%\sent.txt"
echo >>%blat%    objFSO.DeleteFile(%fileattach%)
echo >>%blat%    objFSO.DeleteFile("%syncdir%\log-%id%.txt")
echo >>%blat% Else
echo >>%blat%    file="%syncdir%\failed.txt"
echo >>%blat% End if
echo. >>%blat%
echo >>%blat% objFSO.CreateTextFile(file)
echo >>%blat% Set f = Nothing
echo >>%blat% Set objFlds = Nothing
echo >>%blat% Set objConf = Nothing
echo >>%blat% Set objMail = Nothing
echo >>%blat% Set objFSO = Nothing
:end