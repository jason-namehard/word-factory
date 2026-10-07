@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo  word工厂 打包（桌面版 exe）
echo ============================================================
if not exist ".buildenv\Scripts\python.exe" (
  echo [1/2] 第一次打包,正在建打包环境 .buildenv ...
  python -m venv ".buildenv" || goto :fail
  ".buildenv\Scripts\python.exe" -m pip install --upgrade pip || goto :fail
  rem  pywebview = 原生窗口(用系统自带 Edge WebView2 渲染,不弹浏览器)
  rem  pywin32   = 读 Word/WPS 真实页码 + 导 PDF(不装的话空白页删不掉)
  rem  pyinstaller = 打包器
  ".buildenv\Scripts\python.exe" -m pip install pywebview pywin32 pyinstaller || goto :fail
)
echo [2/2] 开始打包 ...
echo   a) 单文件版 -> dist_onefile\word工厂.exe
".buildenv\Scripts\python.exe" -m PyInstaller build_exe_onefile.spec --noconfirm --distpath dist_onefile --workpath build_onefile || goto :fail
echo   b) 文件夹版 -> dist\word工厂\word工厂.exe
".buildenv\Scripts\python.exe" -m PyInstaller build_exe.spec --noconfirm --distpath dist --workpath build || goto :fail
echo.
echo 打包完成:
echo   单文件: dist_onefile\word工厂.exe      (拷这一个文件到 U 盘即可)
echo   文件夹: dist\word工厂\                 (整个文件夹拷走; 启动更快)
echo 首次运行会在 exe 旁边自动建「word工厂数据」文件夹(规则/方案/临时文件都在那里)。
pause
exit /b 0

:fail
echo.
echo 打包失败,退出码 %ERRORLEVEL%
pause
exit /b %ERRORLEVEL%
