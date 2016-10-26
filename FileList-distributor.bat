@echo off

mode con: cols=100 lines=30

:start

echo COPYBOT3000 GREETS YOU!

echo COPY 'FileList.csv' TO THIS FOLDER AND PRESS ANY BUTTON TO CONTINUE.

pause >nul

:filecheck

IF NOT EXIST FileList.csv (cls & ECHO Sweetheart, you forgot to copy 'FileList.csv' to the current directory. & ECHO Do it and press any button to continue. & pause >nul & GOTO filecheck) ELSE (GOTO processing)

:processing

cls

ECHO WAIT A SECOND WHILE I'M PROCESSING THE FILE.

FINDSTR /M /C:"Copy new item to left;create <-" FileList.csv >nul

IF %ERRORLEVEL% EQU 0 (FOR /F "tokens=1,3 delims=;" %%G IN (FileList.csv) DO (IF "%%H"=="'==" (ECHO %%G>>translated.txt) ELSE (ECHO %%G>>to_translate.txt))

FOR /F "tokens=1,2 delims=;" %%G IN (FileList.csv) DO (IF "%%G"=="create <-" (ECHO %%H>>to_translate.txt))

FOR /F "skip=13 tokens=*" %%G IN (to_translate.txt) DO (IF "%%G"=="create <-" (ECHO[>>to_translate_checked.txt) ELSE (ECHO %%G>>to_translate_checked.txt))

FOR /F "tokens=*" %%G IN (to_translate_checked.txt) DO (ECHO %%G>>to_translate_filtered.txt)) ELSE (FOR /F "tokens=1,3 delims=;" %%G IN (FileList.csv) DO (IF "%%H"=="'==" (ECHO %%G>>translated.txt) ELSE (ECHO %%G>>to_translate.txt))

FOR /F "tokens=1,2 delims=;" %%G IN (FileList.csv) DO (IF "%%G"=="only ->" (ECHO %%H>>to_translate.txt))

FOR /F "skip=11 tokens=*" %%G IN (to_translate.txt) DO (IF "%%G"=="only ->" (ECHO[>>to_translate_checked.txt) ELSE (ECHO %%G>>to_translate_checked.txt))

FOR /F "tokens=*" %%G IN (to_translate_checked.txt) DO (ECHO %%G>>to_translate_filtered.txt))

ECHO.

ECHO PROCESSING COMPLETED! PRESS ANY BUTTON TO CONTINUE!

pause >nul

cls

goto selection

:selection

echo Type 1 if you want to copy matches to 'Matches against previous version' folder

echo Type 2 if you want to copy non-matches to 'Source documents to be translated' folder

echo Type 3 if you want to copy matches and non-matches to appropriate folders

SET /p inputchoice=Make your choice: 

IF "%inputchoice%"=="" GOTO error

IF "%inputchoice%"=="2" (cls & GOTO nottranslated)

IF "%inputchoice%"=="3" (cls & GOTO bothfolders)

IF "%inputchoice%"=="1" (cls & GOTO translated) ELSE (GOTO error)

:translated

ECHO I am copying the files...

FOR /F "tokens=*" %%G in (translated.txt) DO copy "%cd%\KPD\# Source documents docx, doc, xls, xlsx\%%G" "%cd%\KPD\# Matches against previous version" >nul

SET /A counter_MAPV_both=0

FOR %%G IN ("%CD%\KPD\# Matches against previous version\*") DO (SET /A counter_MAPV_both+=1)

cls

echo Matches have been copied to 'Matches against previous version' folder!
ECHO.
ECHO Files copied to 'Matches against previous version': %counter_MAPV_both%

del to_translate.txt
del translated.txt
del to_translate_checked.txt
del to_translate_filtered.txt

GOTO filelist

:nottranslated

ECHO I am copying the files...

FOR /F "tokens=*" %%G in (to_translate_filtered.txt) DO copy "%cd%\KPD\# Source documents docx, doc, xls, xlsx\%%G" "%cd%\KPD\# Source documents to be translated" >nul

SET /A counter_SDTB_both=0

FOR %%G IN ("%CD%\KPD\# Source documents to be translated\*") DO (SET /A counter_SDTB_both+=1)

cls

echo Non-matches have been copied to 'Source documents to be translated' folder!
ECHO.
ECHO Files copied to 'Source documents to be translated': %counter_SDTB_both%
del to_translate.txt
del translated.txt
del to_translate_checked.txt
del to_translate_filtered.txt

GOTO filelist

:bothfolders

ECHO I am copying the files...

FOR /F "tokens=*" %%G in (translated.txt) DO copy "%cd%\KPD\# Source documents docx, doc, xls, xlsx\%%G" "%cd%\KPD\# Matches against previous version" >nul

SET /A counter_MAPV_both=0

FOR %%G IN ("%CD%\KPD\# Matches against previous version\*") DO (SET /A counter_MAPV_both+=1)

FOR /F "tokens=*" %%G in (to_translate_filtered.txt) DO copy "%cd%\KPD\# Source documents docx, doc, xls, xlsx\%%G" "%cd%\KPD\# Source documents to be translated" >nul

SET /A counter_SDTB_both=0

FOR %%G IN ("%CD%\KPD\# Source documents to be translated\*") DO (SET /A counter_SDTB_both+=1)

cls

SET /A counter_total_both=0

FOR %%G IN ("%CD%\KPD\# Source documents docx, doc, xls, xlsx\*") DO (SET /A counter_total_both+=1)

echo Matches and non-matches have been copied to appropriate folders!
ECHO.
ECHO Files copied to 'Matches against previous version': %counter_MAPV_both%
ECHO Files copied to 'Source documents to be translated': %counter_SDTB_both%
ECHO Total files in 'Source documents docx, doc, xls, xlsx': %counter_total_both%
del to_translate.txt
del translated.txt
del to_translate_checked.txt
del to_translate_filtered.txt

GOTO filelist

:filelist
ECHO.

ECHO Do you want to delete 'FileList.csv'?

ECHO Type 1 for Yes

ECHO Type 2 for No

SET /P fldeletion=Make your choice: 

IF "%fldeletion%"=="" GOTO error_deletion

IF "%fldeletion%"=="2" (Exit)

IF "%fldeletion%"=="1" (del FileList.csv & Exit) ELSE (GOTO error_deletion)

:error

cls

ECHO Type a number in the renage from 1 to 3.

GOTO selection

:error_deletion

cls

ECHO Type a number in the renage from 1 to 2.

GOTO filelist
