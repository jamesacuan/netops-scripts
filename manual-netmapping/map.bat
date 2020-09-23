@echo off
for %%x in (s g f l k h r) do net use %%x: /delete /y
cls

TITLE SENCOR Network Mapping
echo.
echo Welcome to SENCOR Domain
echo.
echo.
echo Please enter your username (ex. jecsacuan):
set /p USERID=SENCOR\
echo.
echo Please enter your Password:
set /p PASSWORD=
echo.

:CHOICE
echo SENCOR Department/Group:
echo [1] LCP-Acacia
echo [2] LCP-Mahogany
echo [3] LCP-Molave
echo [4] LCP-Narra
echo [5] CCG
echo.
set /p OPT=Select Group (1-4):
if not '%opt%'=='' set opt=%opt:~0,1%
if '%opt%'=='1' goto ACACIA
if '%opt%'=='2' goto MAHOGANY
if '%opt%'=='3' goto MOLAVE
if '%opt%'=='4' goto NARRA
if '%opt%'=='5' goto CCG

echo %opt% Invalid option, please try again.
ECHO.
goto CHOICE
pause
cls

:NARRA
net use G: "\\denmark\narra" %PASSWORD% /user:SENCOR\%USERID% /PERSISTENT:YES
net use F: "\\denmark\Narra\reports" %PASSWORD% /user:SENCOR\%USERID% /PERSISTENT:YES
goto EXIT

:MOLAVE
net use L: "\\denmark\molave" %PASSWORD% /user:SENCOR\%USERID% /PERSISTENT:YES
goto EXIT

:MAHOGANY
net use K: "\\denmark\mahogany" %PASSWORD% /user:SENCOR\%USERID% /PERSISTENT:YES
goto EXIT

:ACACIA
net use R: "\\denmark\acacia" %PASSWORD% /user:SENCOR\%USERID% /PERSISTENT:YES 
goto EXIT

:CCG
net use H: "\\uk\ccg-ongoing" %PASSWORD% /user:SENCOR\%USERID% /PERSISTENT:YES
net use K: "\\denmark\CCG_TR\Westlaw\tagcheck" %PASSWORD% /user:SENCOR\%USERID% /PERSISTENT:YES
net use L: "\\denmark\CCG_TR\Westlaw\OCR-EDIT" %PASSWORD% /user:SENCOR\%USERID% /PERSISTENT:YES
goto EXIT

:EXIT
net use S: "\\ireland\programs" %PASSWORD% /user:SENCOR\%USERID% /PERSISTENT:YES

pause

