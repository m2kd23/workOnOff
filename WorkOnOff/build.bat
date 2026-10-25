@echo off
set "CSC_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC_PATH%" set "CSC_PATH=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
set "WPF_REF_PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\WPF"
if not exist "%WPF_REF_PATH%\UIAutomationClient.dll" set "WPF_REF_PATH=C:\Windows\Microsoft.NET\Framework\v4.0.30319\WPF"

echo Target Compiler: %CSC_PATH%

"%CSC_PATH%" /target:winexe /out:WorkOnOff.exe /win32manifest:App.manifest /optimize+ /debug- /filealign:512 /platform:anycpu /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:"%WPF_REF_PATH%\UIAutomationClient.dll" /r:"%WPF_REF_PATH%\UIAutomationTypes.dll" Program.cs FileDetector.cs SessionManager.cs Logger.cs

if %ERRORLEVEL% equ 0 (
    echo BUILD SUCCESS
    exit /b 0
) else (
    echo BUILD FAILED
    exit /b %ERRORLEVEL%
)
