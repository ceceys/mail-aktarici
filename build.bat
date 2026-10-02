@echo off
rem Mail Aktarici derleme betigi. Windows'un icindeki .NET Framework derleyicisini kullanir,
rem baska hicbir sey kurmak gerekmez. Cikti: dist\MailAktarici.exe
setlocal
cd /d "%~dp0"
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
if not exist "%CSC%" (
  echo csc.exe bulunamadi. .NET Framework 4 gerekli.
  exit /b 2
)
if not exist dist mkdir dist
set REFS=/r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll
set COMMON=/nologo /codepage:65001 /platform:anycpu /optimize+ /warn:4 /win32manifest:app.manifest /win32icon:app.ico /resource:app.ico,MailAktarici.app.ico %REFS%

"%CSC%" %COMMON% /target:winexe /out:dist\MailAktarici.exe src\*.cs
if errorlevel 1 exit /b 1

rem Test ve betik kullanimi icin konsol surumu (ayni kod, pencere yerine konsola yazar)
if /i "%1"=="konsol" (
  if not exist tests\bin mkdir tests\bin
  "%CSC%" %COMMON% /target:exe /out:tests\bin\MailAktarici-konsol.exe src\*.cs
  if errorlevel 1 exit /b 1
)
echo Derleme tamam: dist\MailAktarici.exe
