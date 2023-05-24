@ECHO off

:: Go to the current directory where this batch file resides.
:: Execute VBScript from the 'scripts' subdirectory.
PUSHD %~dp0
CSCRIPT ./script/refresh.vbs

:: Check for the %ERRORLEVEL% value.
:: 17 - Neither Google Chrome nor Microsoft Edge are active.
IF %ERRORLEVEL% == 17 (
   START https://www.google.com/
) ELSE (
   EXIT 0
)
