@echo off
cd /d %~dp0
:admin
:: 7z unzip usage: 7z x -y file.7z -ofile/
:: use x command to unzip, -y switch to assume yes on all queries
::cd /d "%cpath%"
7za.exe x -y "BurntToast_0.8.5.zip" -o"%ProgramFiles%\WindowsPowerShell\Modules\BurntToast\0.8.5"
echo ����ֵΪ%errorlevel%
if "%errorlevel%" NEQ "0" (
    echo ����ʧ�ܣ������ļ�BurntToast_0.8.5.zip�Ƿ���ڣ������Ƿ��Թ���Ա������д˳���
) else (
    echo ������ɣ������Թرմ����ˡ�
)
echo.
pause