@echo off
set "CSC_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC_PATH%" set "CSC_PATH=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"

echo Target Compiler: %CSC_PATH%

"%CSC_PATH%" /target:winexe /out:WorkResume.exe /win32manifest:App.manifest /optimize+ /platform:anycpu /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll Program.cs FileDetector.cs SessionManager.cs Logger.cs

if %ERRORLEVEL% equ 0 (
    echo BUILD SUCCESS
    exit /b 0
) else (
    echo BUILD FAILED
    exit /b %ERRORLEVEL%
)
