chcp 65001 >nul
@echo off
Title Xoa key Office - by Quang Huy Blog
mode con: cols=96 lines=35
chcp 65001 >nul
@echo.
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo  Run CMD as Administrator...
    goto goUAC 
) else (
 goto goADMIN )

:goUAC
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:goADMIN
    pushd "%CD%"
    CD /D "%~dp0"
	
:main
cls
color f0
@echo. 
echo         WELCOME TO QUANG HUY BLOG
echo              XÓA KEY OFFICE
@echo.
echo     Chọn phiên bản Office cần xóa key:
echo =========================================
echo [  1. Office 2010     : Nhấn phím số 1  ]
echo [  2. Office 2013     : Nhấn phím số 2  ]
echo [  3. Office 2016     : Nhấn phím số 3  ]
echo [  4. Office 2019     : Nhấn phím số 4  ]
echo [  5. Office 365      : Nhấn phím số 5  ]
echo =========================================
Choice /N /C 12345 /M "* Nhập lựa chọn : 
if %errorlevel% == 5 ( set "xx=16" & goto vogia)
if %errorlevel% == 4 ( set "xx=16" & goto vogia)
if %errorlevel% == 3 ( set "xx=16" & goto vogia)
if %errorlevel% == 2 ( set "xx=15" & goto vogia)
if %errorlevel% == 1 ( set "xx=14" & goto vogia)

:vogia
if exist "%ProgramFiles%\Microsoft Office\Office%xx%\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office%xx%"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office%xx%\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office%xx%"
cscript ospp.vbs /dstatus >dstatus.txt
start dstatus.txt
goto office
)

:office
set /p key= * Nhập 5 ký tự cuối của key : 
@echo  ...Đang xóa key Office...
cscript OSPP.VBS /unpkey:%key%
@echo =========================================
@echo      Đã xóa key Office thành công !
@echo =========================================
goto office
)