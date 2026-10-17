@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo  word工厂 打包（桌面版 exe，单文件）
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
".buildenv\Scripts\python.exe" -m PyInstaller build_exe_onefile.spec --noconfirm --distpath dist_onefile --workpath build_onefile || goto :fail
echo.
echo 打包完成: dist_onefile\word工厂.exe
echo 把这一个文件拷到 U 盘即可;首次运行会在它旁边自动建「word工厂数据」文件夹
echo (规则/方案/临时文件/清理审计日志都在那里)。
pause
exit /b 0

:fail
echo.
echo 打包失败,退出码 %ERRORLEVEL%
pause
exit /b %ERRORLEVEL%
