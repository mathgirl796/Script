@echo off
setlocal enabledelayedexpansion

:: Set console to UTF-8 encoding
chcp 936 >nul 2>&1

set "app_name=WORD套打"

echo ========================================
echo %app_name%构建脚本
echo ========================================

:: 获取当前时间戳（格式：YYYYMMDD_HHMMSS）
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "timestamp=%dt:~0,4%-%dt:~4,2%-%dt:~6,2%_%dt:~8,2%-%dt:~10,2%-%dt:~12,2%"

:: 创建版本目录
set "version_dir=version\%app_name%%timestamp%"
if not exist "version" mkdir version
if not exist "%version_dir%" mkdir "%version_dir%"
mkdir build_temp
copy icon.png build_temp

echo 版本目录: %version_dir%
echo 开始构建...

:: 使用PyInstaller构建exe
pyinstaller -F -w -i icon.png -n %app_name% word_mail_merge.py --distpath "%version_dir%" --workpath build_temp --specpath build_temp

:: 检查构建是否成功
if exist "%version_dir%\%app_name%.exe" (
    echo 构建成功！
    
    echo.
    echo ========================================
    echo 构建完成！
    echo 输出目录: %version_dir%
    echo 可执行文件: %version_dir%\%app_name%.exe
    echo ========================================
    
    :: 询问是否打开输出目录
    explorer "%version_dir%"
    
) else (
    echo 构建失败！请检查错误信息。
    if exist "build_temp" rmdir /S /Q build_temp >nul 2>&1
)

echo.
pause
