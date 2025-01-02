@echo off
chcp 65001
setlocal enabledelayedexpansion

set /p fileName=请输入文件名（默认使用 '新建 Microsoft Word 文档'）: 

if "%fileName%"=="" (
    set fileName=新建 Microsoft Word 文档
)

set "desktopPath=D:\我的文档\Desktop"
set "templatePath=C:\Users\DELL\AppData\Roaming\Microsoft\Templates\新建文档脚本专用（修改默认模板后请复制本文件名重新生成）.docx"

if not exist "%desktopPath%" (
    echo 目标路径不存在，请检查路径是否正确.
    pause
    exit /b
)

if not exist "%templatePath%" (
    echo 模板文件不存在，请检查模板路径是否正确.
    pause
    exit /b
)

set "fileName=%fileName: =_%"

set "newFileName=%fileName%.docx"
set "counter=1"
:checkFileName
if exist "%desktopPath%\%newFileName%" (
    set "newFileName=%fileName% (%counter%).docx"
    set /a counter+=1
    goto checkFileName
)

copy /Y "%templatePath%" "%desktopPath%\%newFileName%" >nul

if exist "%desktopPath%\%newFileName%" (
    set successMsg=创建成功，欢迎开启又一场牛马旅程！
    echo !successMsg!
    timeout /t 2 /nobreak >nul
    start "" "winword" "%desktopPath%\%newFileName%"
    timeout /t 2 /nobreak >nul
    exit
) else (
    echo 文件创建失败。
    pause
)

endlocal
