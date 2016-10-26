@echo off


mode con: cols=100 lines=30

ECHO Hi there! I am the Matches copier.
ECHO.
ECHO I will find translations for each document in 'Matches against previous version' folder.
ECHO.
ECHO The only thing you need to do is to set a path to 'For publication' folder of the previous project!
ECHO.
ECHO Press any button to continue.
pause >nul
cls
ECHO Wait a second while I am trying to open the dialog box for selecting paths.

setlocal

set "psCommand="(new-object -COM 'Shell.Application')^
.BrowseForFolder(0,'Please, specify a path to ''For publication'' folder of the previous project. I will copy all the matches I find to ''For publication'' folder of the current project.',0,0).self.path""

for /f "usebackq delims=" %%I in (`powershell %psCommand%`) do set "folder=%%I"

set current_dir=%CD%

mkdir "%CD%\Temp"

copy "%CD%\KPD\# Matches against previous version\*.*" "%CD%\Temp" >nul


FOR %%G IN ("%CD%\Temp\*.*") do ECHO %%~nG>temp.txt
cls
ECHO You need to wait again, boy! 
ECHO.
ECHO This time the operation may take several minutes.

FOR /F "tokens=1,2,3 delims=-." %%G IN (temp.txt) DO (set one=%%G
set two=%%H
set three=%%I)

del temp.txt


IF "%three%"=="RU" (set any_ru=%one%-%two%-EN) ELSE (GOTO second_ru)

SET source_mask=%one%-%two%-%three%

CD "%cd%\Temp"

REN "%source_mask%*" "%any_ru%*"

GOTO copying_names

:second_ru

IF "%two%"=="RU" (set second_ru=%one%-EN-%three%)

SET source_mask=%one%-%two%-%three%

CD "%cd%\Temp"

REN "%source_mask%*" "%second_ru%*"

GOTO copying_names

:copying_names

CD "%current_dir%"

FOR %%G IN ("%CD%\Temp\*.*") do ECHO %%~nG>>temp.txt

setlocal enabledelayedexpansion

FOR /F "tokens=*" %%G in (temp.txt) DO COPY /-Y "!folder!\%%G.*" "%CD%\KPD\# For publication"

del temp.txt

rmdir /S /Q "%CD%\Temp" 

cls

ECHO Matches have been copied! I hope you enjoyed using me.

ECHO Say 'Goodbye' and press any button to close me :]

endlocal

pause
