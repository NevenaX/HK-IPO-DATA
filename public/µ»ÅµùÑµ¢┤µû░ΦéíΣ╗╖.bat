@echo off
chcp 65001 >nul
cd /d "C:\Users\84517\Desktop\AI AUTO\hk_ipo_csdata"

echo [%date% %time%] 开始自动更新股价...

python scripts\fetch_company_info.py
if errorlevel 1 (
    echo [失败] 股价更新失败
    exit /b 1
)

python scripts\bundle_standalone.py
if errorlevel 1 (
    echo [失败] standalone 打包失败
    exit /b 1
)

echo [%date% %time%] 股价更新完成，standalone 已同步
