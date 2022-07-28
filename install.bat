@echo off
cd /d %~dp0
:admin
:: 7z unzip usage: 7z x -y file.7z -ofile/
:: use x command to unzip, -y switch to assume yes on all queries
::cd /d "%cpath%"
7za.exe x -y "BurntToast_0.8.5.zip" -o"%ProgramFiles%\WindowsPowerShell\Modules\BurntToast\0.8.5"
echo 返回值为%errorlevel%
if "%errorlevel%" NEQ "0" (
    echo 复制失败，请检查文件BurntToast_0.8.5.zip是否存在，或者是否以管理员身份运行此程序。
) else (
    echo 复制完成，您可以关闭窗口了。
)
echo.
pause