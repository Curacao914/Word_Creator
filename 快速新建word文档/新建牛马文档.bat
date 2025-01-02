@echo off
chcp 65001
setlocal enabledelayedexpansion

:: 弹出窗口让用户输入文件名
set /p fileName=请输入文件名（默认使用 '新建 Microsoft Word 文档'）: 

:: 如果没有输入文件名，则使用默认值
if "%fileName%"=="" (
    set fileName=新建 Microsoft Word 文档
)

:: 直接设置你的桌面路径（手动指定）
set "desktopPath=D:\我的文档\Desktop"

:: 设置模板路径
set "templatePath=C:\Users\DELL\AppData\Roaming\Microsoft\Templates\新建文档脚本专用（修改默认模板后请复制本文件名重新生成）.docx"

:: 检查桌面文件夹是否存在
if not exist "%desktopPath%" (
    echo 目标路径不存在，请检查路径是否正确.
    pause
    exit /b
)

:: 检查模板文件是否存在
if not exist "%templatePath%" (
    echo 模板文件不存在，请检查模板路径是否正确.
    pause
    exit /b
)

:: 将文件名中的空格替换为下划线（以防止路径问题）
set "fileName=%fileName: =_%"

:: 检查文件是否已存在，如果存在，添加 (n) 后缀
set "newFileName=%fileName%.docx"
set "counter=1"
:checkFileName
if exist "%desktopPath%\%newFileName%" (
    set "newFileName=%fileName% (%counter%).docx"
    set /a counter+=1
    goto checkFileName
)

:: 复制模板文件并重命名为目标文件
copy /Y "%templatePath%" "%desktopPath%\%newFileName%" >nul

:: 检查文件是否成功创建
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

:: 退出脚本，保持窗口打开
endlocal
