@echo off
(cd /d "%~dp0")&&(NET FILE||(powershell start-process -FilePath '%0' -verb runas)&&(exit /B)) >NUL 2>&1
title Office 2019 Activator r/Piracy
echo Converting... & mode 40,25
(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)
cscript //nologo ospp.vbs /unpkey:6MWKP >nul&cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul&set i=1
:server
if %i%==1 (
    set /p KMS_Sev=Please enter KMS server:
    )
if %i%==2  (
    set /p KMS_Sev=Activation failed, please re-type kms server:
    )
if %i%==3  (
    set /p KMS_Sev=Activation failed, please re-type kms server:
    )

cscript //nologo ospp.vbs /sethst:%KMS_Sev% >nul
echo Set KMS Server: %KMS_Sev% & echo Activating...
cscript //nologo ospp.vbs /act | find /i "successful" && (echo Complete) || (echo Acitvate Fail ! Trying another KMS Server & set /a i+=1 & goto server)
echo Start activating Windows......
slmgr /ipk WMN7B-Y7TKF-Y49QB-TMQ8T-GMT6T
slmgr /skms %KMS_Sev%
slmgr /ato 
slmgr /dlv
echo Finished
pause