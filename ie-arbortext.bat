@echo off
setlocal EnableDelayedExpansion

set "trusted_domains=HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains"

for %%G in (
	"http://uat.workbench.judicial.int.westgroup.com"
	"http://qa.bpms.jpath.judicial.int.westgroup.com"
	"http://test.bpms.jpath.judicial.thomsonreuters.com"
	"http://uat.bpms.jpath.judicial.thomsonreuters.com"
	"https://web2.westlaw.com"
	"http://prod.workflowui.int.thomsonreuters.com"
	"http://qa.cpp.securityservice.int.westgroup.com"
	"http://test.workbench.judicial.int.westgroup.com"
	"https://eg.tlrcitrix.thomsonreuters.com"
	"http://prod.bpms.jpath.judicial.ha.thomson.com:9083"
	"http://prod.judicial.multimediarepository.int.westgroup.com"
	"http://qa.workbench.judicial.int.westgroup.com"
	"http://prod.workbench.judicial.int.westgroup.com"
) do (
	set "site=%%~G"
	set "site=!site:*://=!"
	set "site=!site:*www.=!"
	>nul reg add "%trusted_domains%\!site!" /v * /t REG_DWORD /d 1 /f
)

pause