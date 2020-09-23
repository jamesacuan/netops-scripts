@echo off

title SENCOR-WFH Network Diagnostics
mode con: cols=100 lines=29

rem initialize
set incr=1
set domain=sencor.net
set noddomain=sencor-mnl.net
set server=postoffice
set googDNS=8.8.8.8
set ipvpnT=203.147.107.250
set ipvpnE=116.50.227.91
set ipvpnTborder=210.176.35.185
set ipvpnEborder=116.50.227.65
set logonserver=192.168.1.61
set ipvpnTname=telstra_sencor
set ipvpnEname=etpi_sencor
set ipvpnTbordername=telstra_border
set ipvpnEbordername=etpi_border
set syncdir=%userprofile%\.sencor
set tempdir=%syncdir%\temp
set redo=false
set spdydomain=www.speedtest.net
set spdyconfig=speedtest-config
set spdyserver=speedtest-servers
set customdate=%date:~10,4%%date:~4,2%%date:~7,2%

for %%x in (%googDNS% %logonserver%) do (
    ping %%x -n 1 -w 1000
    cls
    if errorlevel 1 goto :offline
)

if not exist %tempdir% mkdir %tempdir%
if exist %userprofile%\.sencor\info.cmd GOTO :RELOADINFO

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

:offline
echo.
echo.
echo	You're not connected to the Internet or to GlobalProtect.
echo.
echo	Kindly check your connection settings and try running
echo	this app again. Thanks.
echo.
echo.
echo.
echo.
pause
goto :end


:SAVINGINFO
@echo @ECHO OFF > "%syncdir%\info.cmd" 
@echo set FULLNAME=%FULLNAME% > "%syncdir%\info.cmd"
@echo set USERNAME=%USERNAME% >> "%syncdir%\info.cmd"
@echo set GROUP=%GROUP% >> "%syncdir%\info.cmd"
@echo set DEPT=%DEPT% >> "%syncdir%\info.cmd"
@echo set IDNO=%IDNO% >> "%syncdir%\info.cmd"
@echo set ISP=%ISP% >> "%syncdir%\info.cmd"
@echo set DATEREGISTERED=%DATE% %TIME% >> "%syncdir%\info.cmd"
echo.
echo saved.
pause


:RELOADINFO
call "%syncdir%\info.cmd"
setlocal enabledelayedexpansion

for /f "tokens=* delims= " %%a in ("%USERNAME%") do set USERNAME=%%a
for /l %%a in (1,1,100) do if "!USERNAME:~-1!"==" " set USERNAME=!USERNAME:~0,-1!

rem remove breadcrumbs
for %%x in (vbs txt json png) do (
    if exist "%tempdir%\*.%%x" del "%tempdir%\*.%%x"
)

if exist "%syncdir%\pack*.zip" (
    rem redo
    set redo=true
    call "%syncdir%\tmpPend.cmd"
    
    set xdate=""
    set xtime=""
    set ydate=""
    set ytime=""
    
    GOTO :DIAGSTEP7
) ELSE (
    rem Start Fresh
    set redo=false
    if exist "%syncdir%\tmpPend-*.cmd" del "%syncdir%\tmpPend-*.cmd"
    if exist "%syncdir%\tmpBody-*.txt" del "%syncdir%\tmpBody-*.txt"
    rem pause
    rem goto :DIAGSTEP6
)

:DIAGNOSTICS
REM initialize
SETLOCAL EnableDelayedExpansion
if %incr% lss 10 set incr=0%incr%
set id=%customdate%-%incr%

FOR /F "skip=1" %%A IN ('WMIC OS GET LOCALDATETIME') DO (SET "t=%%A" & GOTO break_1)
:break_1

SET "m=%t:~10,2%" & SET "h=%t:~8,2%" & SET "d=%t:~6,2%" & SET "z=%t:~4,2%" & SET "y=%t:~0,4%"
IF !h! GTR 11 (SET /A "h-=12" & SET "ap=P" & IF "!h!"=="0" (SET "h=00") ELSE (IF !h! LEQ 9 (SET "h=0!h!"))) ELSE (SET "ap=A")

set xdate=%z%-%d%-%y%
set xtime=%h%:%m% %ap%M

rem if exist "%syncdir%\temp.zip" del "%syncdir%\temp.zip"

powershell -command "wget https://api.ipify.org/ -OutFile '%tempdir%/ip.txt'"
set /p ipadd=<%userprofile%/.sencor/temp/ip.txt

set ydate=""
set ytime=""
set PANGPlog="%localappdata%\Palo Alto Networks\GlobalProtect\PanGPA.log"
set count=0

@echo @ECHO OFF > "%syncdir%\tmpPend-%id%.cmd
@echo set PENDING_JOB=TRUE >> "%syncdir%\tmpPend-%id%.cmd"
@echo set PENDING_DATETIME=%xdate% %xtime% >> "%syncdir%\tmpPend-%id%.cmd"
echo.

:DIAGBODY
cls
set horizontal=******************************************************************
echo.
echo %horizontal%
echo.
echo    Welcome, %fullname%^^!
echo.
echo    This is a three part series test and would take around
echo    5-8 minutes. Please DON'T CLOSE THIS WINDOW until our
echo    network diagnostic is complete.
echo.
echo    You may minimize this console and continue your work.   
echo.
echo %horizontal%
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
echo        a. vpn.sencor.net
if %count%==1 GOTO :BUSY
echo        b. etpi-v.sencor.net
if %count%==2 GOTO :BUSY
echo.
echo    2. TRACEROUTE Test
echo        a. vpn.sencor.net
if %count%==3 GOTO :BUSY
echo        b. etpi-v.sencor.net
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
rem powershell Write-Progress -Activity 'Network Diagnostics' -Status 'In progress' -PercentComplete ([math]::Round((%count%/7)*100))
goto :DIAGSTEP%COUNT%


:DIAGSTEP0
REM 0. INFORMATION GATHERING
powershell -command "wget http://ip-api.com/json/%ipadd% -OutFile '%tempdir%/netdetails-%username%.json'" 
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
REM prepare PING file
set ping="%syncdir%\ping.ps1"
type NUL > %ping%
@echo > %ping% $hostname = $args[0]
@echo >> %ping% $descript = $args[1]
@echo. >> %ping% 
@echo >> %ping% ping.exe $hostname -n 15 ^|Foreach{"{0} - {1}" -f (Get-Date),$_} ^| Out-File -Filepath "%tempdir%\ping_$descript-%username%.txt"
GOTO :DIAGBODY

:DIAGSTEP1
REM 1. PING / A. VPN
powershell %ping% %ipvpnTborder% %ipvpnTbordername%
powershell %ping% %ipvpnT% %ipvpnTname%
set count=2
GOTO :DIAGBODY

:DIAGSTEP2
REM 1. PING / B. ETPI
powershell %ping% %ipvpnEborder% %ipvpnEbordername%
powershell %ping% %ipvpnE% %ipvpnEname%
del %ping%
set count=3
GOTO :DIAGBODY

:DIAGSTEP3
REM 2. TRACEROUTE / A. VPN
tracert %ipvpnTborder% > "%tempdir%\tracert_%ipvpnTbordername%-%username%.txt"
tracert %ipvpnT% > "%tempdir%\tracert_%ipvpnTname%-%username%.txt"
set count=4
GOTO :DIAGBODY

:DIAGSTEP4
REM 2. TRACEROUTE / B. ETPI
tracert %ipvpnE% > "%tempdir%\tracert_%ipvpnEname%-%username%.txt"
tracert %ipvpnEborder% > "%tempdir%\tracert_%ipvpnEbordername%-%username%.txt"
set count=5
GOTO :DIAGBODY

:DIAGSTEP5
REM WRAPPING THINGS UP
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
REM 3. SPEEDTEST
set %spdy% = "%syncdir%\speed.ps1"
type nul > %spdy%
@echo >%spdy% $objXmlHttp = New-Object -ComObject MSXML2.ServerXMLHTTP
@echo >>%spdy% $objXmlHttp.Open("GET", "%spdydomain%/%spdyconfig%.php", $False)
@echo >>%spdy% $objXmlHttp.Send()
@echo. >>%spdy%
@echo >>%spdy% [xml]$content = $objXmlHttp.responseText
@echo >>%spdy% $oriLat = $content.settings.client.lat
@echo >>%spdy% $oriLon = $content.settings.client.lon
@echo. >>%spdy%
@echo >>%spdy% $objXmlHttp1 = New-Object -ComObject MSXML2.ServerXMLHTTP
@echo >>%spdy% $objXmlHttp1.Open("GET", "%spdydomain%/%spdyserver%.php", $False)
@echo >>%spdy% $objXmlHttp1.Send()
@echo. >>%spdy%
@echo >>%spdy% [xml]$ServerList = $objXmlHttp1.responseText
@echo >>%spdy% $cons = $ServerList.settings.servers.server
@echo. >>%spdy%
rem GET CLOSEST SERVER
@echo >>%spdy% foreach($val in $cons) { 
@echo >>%spdy%     $R = 6371;
@echo >>%spdy%     [float]$dlat = ([float]$oriLat - [float]$val.lat) * 3.14 / 180;
@echo >>%spdy%     [float]$dlon = ([float]$oriLon - [float]$val.lon) * 3.14 / 180;
@echo >>%spdy%     [float]$a = [math]::Sin([float]$dLat/2) * [math]::Sin([float]$dLat/2) + [math]::Cos([float]$oriLat * 3.14 / 180 ) * [math]::Cos([float]$val.lat * 3.14 / 180 ) * [math]::Sin([float]$dLon/2) * [math]::Sin([float]$dLon/2);
@echo >>%spdy%     [float]$c = 2 * [math]::Atan2([math]::Sqrt([float]$a ), [math]::Sqrt(1 - [float]$a));
@echo >>%spdy%     [float]$d = [float]$R * [float]$c;
@echo >>%spdy% 
@echo >>%spdy%     $ServerInformation +=
@echo >>%spdy% @([pscustomobject]@{Distance = $d; Country = $val.country; Sponsor = $val.sponsor; Url = $val.url })
@echo >>%spdy% }

:DIAGSTEP6
REM GENERATING REPORT

type NUL > "%syncdir%\tmpbody-%id%.txt"
set tmpbody="%syncdir%\tmpbody-%id%.txt"
@echo Name: %FULLNAME% > %tmpbody%
@echo Department: %DEPT% >> %tmpbody%
@echo Group: %GROUP% >> %tmpbody%
@echo Employee ID: %IDNO% >> %tmpbody%
@echo. >> %tmpbody%
@echo ISP: %ISP% >> %tmpbody%
@echo IP:  %ipadd% >> %tmpbody%
@echo. >> %tmpbody%
@echo Started: %xdate% %xtime%>> %tmpbody%
@echo Finished: %ydate% %ytime%>> %tmpbody%
rem important results to body
@echo. >> %tmpbody%
@echo. >> %tmpbody%
@echo PING TO VPN.SENCOR.NET >> %tmpbody% 
@echo ---------------------------------- >> %tmpbody%
type "%tempdir%\ping_telstra_sencor-%username%.txt" >> %tmpbody%
@echo. >> %tmpbody%
@echo. >> %tmpbody%
@echo. >> %tmpbody%
@echo PING TO ETPI-V.SENCOR.NET >> %tmpbody% 
@echo ---------------------------------- >> %tmpbody%
type "%tempdir%\ping_etpi_sencor-%username%.txt" >> %tmpbody%
@echo. >> %tmpbody%
@echo. >> %tmpbody%
@echo TRACEROUTE TO VPN.SENCOR.NET >> %tmpbody% 
@echo ---------------------------------- >> %tmpbody%
type "%tempdir%\tracert_telstra_sencor-%username%.txt" >> %tmpbody%
@echo. >> %tmpbody%
@echo. >> %tmpbody%
@echo. >> %tmpbody%
@echo TRACEROUTE TO ETPI-V.SENCOR.NET >> %tmpbody% 
@echo ---------------------------------- >> %tmpbody%
type "%tempdir%\tracert_etpi_sencor-%username%.txt" >> %tmpbody%
@echo. >> %tmpbody%
@echo. >> %tmpbody%
@echo. >> %tmpbody%
@echo ADDITIONAL INFORMATION >> %tmpbody% 
@echo ---------------------------------- >> %tmpbody%
type "%tempdir%\netdetails-%username%.json" >> %tmpbody%
del "%tempdir%\netdetails-%username%.json"

rem zipping report

echo Set objArgs = WScript.Arguments > "%syncdir%\_zipIt.vbs"
echo InputFolder = objArgs(0) >> "%syncdir%\_zipIt.vbs"
echo ZipFile = objArgs(1) >> "%syncdir%\_zipIt.vbs"
echo CreateObject("Scripting.FileSystemObject").CreateTextFile(ZipFile, True).Write "PK" ^& Chr(5) ^& Chr(6) ^& String(18, vbNullChar) >> "%syncdir%\_zipIt.vbs"
echo Set objShell = CreateObject("Shell.Application") >> "%syncdir%\_zipIt.vbs"
echo Set source = objShell.NameSpace(InputFolder).Items >> "%syncdir%\_zipIt.vbs"
echo objShell.NameSpace(ZipFile).CopyHere(source) >> "%syncdir%\_zipIt.vbs"
echo wScript.Sleep 2000 >> "%syncdir%\_zipIt.vbs"
CScript "%syncdir%\_zipIt.vbs"  "%tempdir%"  "%syncdir%\pack-%id%.zip"
del "%syncdir%\_zipIt.vbs"
del "%tempdir%\*.txt"
del "%tempdir%\*.json"
del "%tempdir%\*.png"
if exist "%tempdir%\PanGPA.log" del "%tempdir%\PanGPA.log"
set count=7
GOTO :DIAGBODY

:DIAGSTEP7
set from=notify
set to=nog.wfhmonitoring
set cc=networkoperations
set subj="SENCOR WFH NET DIAGNOSTICS - %username%"
if %redo%==true (
    cls
    echo.
    echo We found a pending job dated %PENDING_DATETIME%
    echo.
    echo.
    echo We will attempt to send this first.
    echo.
    if exist "%syncdir%\*.vbs" del "%syncdir%\*.vbs"
    if exist "%syncdir%\tmpbody.txt" (
        set tmpbody="%syncdir%\tmpbody.txt"
    ) else (
        echo "old unset log" >> "%syncdir%\tmpbody.txt"
        set tmpbody="%syncdir%\tmpbody.txt"
    )
    rem set count=5
    pause
    rem goto :DIAGBODY
)

REM SENDING INFORMATION

call :createVBS "%syncdir%\email-bat.vbs"
call :send
if exist "%syncdir%\sent.txt" (
   echo.
   echo "Logs have been sent to Datacomm (NOD). Thanks."
   echo.
   timeout /t 10
   del "%syncdir%\*.txt"
   del "%syncdir%\*.vbs"
) else (
   cls
   echo.
   powershell write-host -foregroundcolor red "We are unable to send the logs. We will send it out on your next network diagnostics."
   echo.
   timeout /t 10
)

if %redo%==true (
    rem SECOND RUN AFTER FAIL SENT
    %redo%=false
    SETLOCAL EnableDelayedExpansion

    FOR /F "skip=1" %%A IN ('WMIC OS GET LOCALDATETIME') DO (SET "t=%%A" & GOTO break_1)
    :break_1

    SET "m=%t:~10,2%" & SET "h=%t:~8,2%" & SET "d=%t:~6,2%" & SET "z=%t:~4,2%" & SET "y=%t:~0,4%"
    if !h! GTR 11 (SET /A "h-=12" & SET "ap=P" & IF "!h!"=="0" (SET "h=00") ELSE (IF !h! LEQ 9 (SET "h=0!h!"))) ELSE (SET "ap=A")

    set xdate=%z%-%d%-%y%
    set xtime=%h%:%m% %ap%M
    goto :DIAGNOSTICS
)
pause
goto :EOF

:send
cscript.exe /nologo "%vbsfile%"
goto :EOF

:createVBS
set "vbsfile=%~1"
set fileattach="%syncdir%\pack-%id%.zip"
del "%vbsfile%" 2>nul
set cdoSchema=http://schemas.microsoft.com/cdo/configuration
echo >>"%vbsfile%" Set objMail = CreateObject("CDO.Message")
echo >>"%vbsfile%" Set objConf = CreateObject("CDO.Configuration")
echo >>"%vbsfile%" Set objFSO  = CreateObject("Scripting.FileSystemObject")
echo >>"%vbsfile%" Set objFlds = objConf.Fields
echo >>"%vbsfile%" Set f = objFSO.OpenTextFile(%tmpbody%, 1)
echo. >>"%vbsfile%" 
echo >>"%vbsfile%" objFlds.Item("%cdoSchema%/sendusing") = 2 'cdoSendUsingPort
echo >>"%vbsfile%" objFlds.Item("%cdoSchema%/smtpserver") = "%server%.%domain%"
echo >>"%vbsfile%" objFlds.Item("%cdoSchema%/smtpserverport") = 587
echo >>"%vbsfile%" objFlds.Update
echo. >>"%vbsfile%" 
echo >>"%vbsfile%" objMail.Configuration = objConf
echo >>"%vbsfile%" objMail.From = "%from%@%noddomain%"
echo >>"%vbsfile%" objMail.To = "nog.wfhmonitoring@sencor.net"
echo >>"%vbsfile%" objMail.CC = "sencor.datacomm@gmail.com"
echo >>"%vbsfile%" objMail.Subject = %subj%
echo >>"%vbsfile%" objMail.TextBody = f.ReadAll
echo >>"%vbsfile%" f.Close
echo >>"%vbsfile%" objMail.AddAttachment %fileattach%
echo >>"%vbsfile%" objMail.Send
echo. >>"%vbsfile%" 
echo >>"%vbsfile%" file="" 
echo >>"%vbsfile%" If err.number = 0 then 
echo >>"%vbsfile%"    file="%syncdir%\sent.txt"
echo >>"%vbsfile%"    objFSO.DeleteFile("%syncdir%\tmpPend-%id%.cmd")
echo >>"%vbsfile%"    objFSO.DeleteFile(%fileattach%)
echo >>"%vbsfile%"    objFSO.DeleteFile("%syncdir%\tmpbody-%id%.txt")
echo >>"%vbsfile%" Else
echo >>"%vbsfile%"    file="%syncdir%\failed.txt"
echo >>"%vbsfile%" End if
echo. >>"%vbsfile%"
echo >>"%vbsfile%" objFSO.CreateTextFile(file)
echo >>"%vbsfile%" Set f = Nothing
echo >>"%vbsfile%" Set objFlds = Nothing
echo >>"%vbsfile%" Set objConf = Nothing
echo >>"%vbsfile%" Set objMail = Nothing
echo >>"%vbsfile%" Set objFSO = Nothing
:end