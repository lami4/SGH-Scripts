@echo off

mode con: cols=200 lines=80
SET root_directory=%CD%

ECHO HELLO THERE!
ECHO I KNOW WHY YOU ARE HERE, YOU CAME TO UPDATE SOME FILES IN '...\KPD\...' SUBFOLDERS.
ECHO PRESS ANY BUTTON TO CONTINUE.
PAUSE >nul

:list_file_creation
IF EXIST List.txt CLS & GOTO list_exists
ECHO INSTRUCTIONS:>>List.txt
ECHO      1) ONE LINE FOR ONE NAME;>>List.txt
ECHO      2) NO EXTENSIONS OR SPACES IN THE END OF A NAME (Example - PABKRF-GL-RU-07.01.00.dSTP.01.00);>>List.txt
ECHO      3) IF YOU WANT TO REMOVE A DOCUMENT FROM THE FOLDERS, TYPE 'REM' BEFORE A NAME (Example - REM PABKRF-GL-RU-07.01.00.dSTP.01.00);>>List.txt
ECHO      4) DO NOT FORGET TO SAVE THIS FILE!>>List.txt
ECHO.>>List.txt
ECHO.>>List.txt
ECHO Now, please, specify documents whose files will be updated.>>List.txt
ECHO.>>List.txt
ECHO START THE LIST FROM THE NEXT LINE:>>List.txt

:list_creation
FOR %%H IN (List.txt) DO (SET filesize=%%~zH)
START /W List.txt
FOR %%G IN (List.txt) DO (IF %%~zG EQU %filesize% (CLS & GOTO no_changes) ELSE (CLS & GOTO path_selection)) 
PAUSE

:no_changes
ECHO YOU HAVE NOT MADE ANY CHANGES TO 'List.txt' OR FORGOT TO SAVE IT!
ECHO.
ECHO TYPE 1 TO CLOSE THE SCRIPT
ECHO TYPE 2 TO OPEN 'List.txt' AGAIN AND HAVE ANOTHER GO
ECHO.
SET /P inputchoice_no_changes=Make your choice: 
IF "%inputchoice_no_changes%"=="" GOTO inputchoice_no_changes_wronginput
IF "%inputchoice_no_changes%"=="1" CLS & GOTO exit_from_no_changes
IF "%inputchoice_no_changes%"=="2" (cls & GOTO list_creation) ELSE (GOTO inputchoice_no_changes_wronginput)

:inputchoice_no_changes_wronginput
CLS
ECHO Please, type a number in the range from 1 to 2.
GOTO no_changes

:exit_script
EXIT

:exit_from_no_changes
ECHO ONE MORE QUESTION BEFORE CLOSING THE SCRIPT...
ECHO DO YOU WANT TO DELETE 'List.txt'?
ECHO.
ECHO TYPE 1 TO DELETE IT
ECHO TYPE 2 TO LEAVE IT AS IT IS
ECHO.
SET /P inputchoice_exit_from_no_changes=Make your choice: 
IF "%inputchoice_exit_from_no_changes%"=="" GOTO inputchoice_exit_from_no_changes_wronginput
IF "%inputchoice_exit_from_no_changes%"=="1" (DEL List.txt & GOTO exit_script)
IF "%inputchoice_exit_from_no_changes%"=="2" (GOTO exit_script) ELSE (GOTO inputchoice_exit_from_no_changes_wronginput)

:inputchoice_exit_from_no_changes_wronginput
CLS
ECHO Please, type a number in the range from 1 to 2.
GOTO exit_from_no_changes

:list_exists
ECHO 'List.txt' ALREADY EXISTS IN THE DIRECTORY!
ECHO.
ECHO TYPE 1 TO CLOSE THE SCRIPT
ECHO TYPE 2 TO OPEN THE EXISTING 'List.txt' AND EDIT IT
ECHO TYPE 3 TO START UPDATING FILES WITHOUT EDITING THE EXISTING 'List.txt'
ECHO TYPE 4 TO DELETE THE EXISTING 'List.txt' AND CREATE A NEW BLANK ONE
ECHO.
SET /P inputchoice_list_exists=Make your choice: 
IF "%inputchoice_list_exists%"=="" GOTO inputchoice_list_exists_wronginput
IF "%inputchoice_list_exists%"=="1" (CLS & GOTO exit_from_no_changes)
IF "%inputchoice_list_exists%"=="2" (GOTO list_editing_list_exists)
IF "%inputchoice_list_exists%"=="3" (CLS & GOTO path_selection)
IF "%inputchoice_list_exists%"=="4" (DEL List.txt & PAUSE & GOTO list_file_creation) ELSE (GOTO inputchoice_list_exists_wronginput)

:inputchoice_list_exists_wronginput
CLS
ECHO Please, type a number in the range from 1 to 4.
GOTO list_exists

:list_editing_list_exists
FOR %%H IN (List.txt) DO (SET filesize=%%~zH)
START /W List.txt
FOR %%G IN (List.txt) DO (IF %%~zG EQU %filesize% (CLS & GOTO no_changes_list_exists) ELSE (CLS & GOTO path_selection)) 

:no_changes_list_exists
ECHO YOU HAVE NOT MADE ANY CHANGES TO 'List.txt' OR FORGOT TO SAVE IT!
ECHO.
ECHO TYPE 1 TO CLOSE THE SCRIPT
ECHO TYPE 2 TO OPEN 'List.txt' AGAIN AND HAVE ANOTHER GO
ECHO TYPE 3 TO START UPDATING FILES WITH THE EXISTING 'List.txt'
ECHO.
SET /P inputchoice_no_changes_list_exists=Make your choice: 
IF "%inputchoice_no_changes_list_exists%"=="" GOTO inputchoice_no_changes_list_exists_wronginput
IF "%inputchoice_no_changes_list_exists%"=="1" CLS & GOTO exit_from_no_changes
IF "%inputchoice_no_changes_list_exists%"=="2" CLS & GOTO list_editing_list_exists
IF "%inputchoice_no_changes_list_exists%"=="3" (CLS & GOTO path_selection) ELSE (GOTO inputchoice_no_changes_list_exists_wronginput)

:inputchoice_no_changes_list_exists_wronginput
CLS
ECHO Please, type a number in the range from 1 to 3.
GOTO no_changes_list_exists

:path_selection
ECHO So far so good.
ECHO.
set /p notification=Please, enter the notification number: 
IF "%notification%"=="" (GOTO wronginput_notification) ELSE (GOTO going_on)

:wronginput_notification
CLS
ECHO Sir, you cannot leave this field empty.
GOTO path_selection

:going_on
MKDIR "MemoQ import"
CLS
ECHO WAIT A SECOND, I AM TRYING TO OPEN THE DIALOG BOX SO YOU CAN SPECIFY A PATH!
FOR /F "skip=10 tokens=1,2 delims= " %%G IN (List.txt) DO (IF "%%G"=="REM" (ECHO %%H>>Del_temporary.txt) ELSE (ECHO %%G>>Upd_temporary.txt))

SETLOCAL

set "psCommand="(new-object -COM 'Shell.Application')^
.BrowseForFolder(0,'Please choose a folder.',0,0).self.path""
for /f "usebackq delims=" %%I in (`powershell %psCommand%`) do set "folder=%%I"
setlocal enabledelayedexpansion
:files_found
ECHO.
ECHO BELOW ARE FILES I FOUND IN "!folder!" FOR EACH DOCUMENT SPECIFIED IN 'List.txt'.
ECHO.
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------
FOR /F "tokens=*" %%G in (Upd_temporary.txt) DO (ECHO Files for %%G: & XCOPY "!folder!\%%G.*" "%CD%\MemoQ import" /L & ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------)
ECHO.
ECHO MAKE SURE THAT *.pdf AND *.doc^(x^)/xls^(x^) FILES WERE FOUND FOR EACH DOCUMENT SPECIFIED IN 'List.txt'.
ECHO.
ECHO IF EVERYTHING SEEMS TO BE OK, TYPE 1 TO CONTINUE
ECHO IF ONE OR MORE DOCUMENTS DO NOT HAVE A *.pdf-*.doc^(x^)/xls^(x^) PAIR, TYPE 2 TO STOP THE SCRIPT AND FIGURE OUT WHAT IS WRONG
ECHO.
SET /P inputchoice_path_selection=Make your choice: 
IF "%inputchoice_path_selection%"=="" GOTO inputchoice_path_selection_wronginput
IF "%inputchoice_path_selection%"=="1" (GOTO copying_files)
IF "%inputchoice_path_selection%"=="2" (RMDIR MemoQ import & DEL Del_temporary.txt & DEL Upd_temporary.txt & GOTO exit_from_no_changes) ELSE (GOTO inputchoice_path_selection_wronginput)

:copying_files
FOR /F "tokens=*" %%G in (Upd_temporary.txt) DO XCOPY "!folder!\%%G.*" "%CD%\MemoQ import" >nul
CLS
GOTO processing_files

ENDLOCAL

:inputchoice_path_selection_wronginput
CLS
ECHO Please, type a number in the range from 1 to 2.
GOTO files_found

:processing_files
MKDIR "%CD%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%"
MKDIR "%CD%\KPD\# Source documents pdf\Files updated on %date% by notification No. %notification%"
FINDSTR /B "REM" List.txt >nul
IF %ERRORLEVEL% EQU 0 (GOTO removing_files) ELSE (GOTO processing_without_removing_files)

:removing_files
SET /a var_incr_del_pdf=1
SET /a var_incr_del_pdf_suc=0
SET /a var_incr_del_docs=1
SET /a var_incr_del_docs_suc=0
SET /a var_incr_del_both=1
SET /a var_incr_del_both_suc=0

SETLOCAL ENABLEDELAYEDEXPANSION
ECHO SOME DOCUMENTS IN 'List.txt' WERE MARKED AS 'REM' ^(=TO BE DELETED IN ALL "...\KPD\..." SUBFOLDERS^).
ECHO I WILL DO AS YOU SAY, BUT PUT THEIR COPIES IN THE APPROPRIATE 'Files updated on %date% by notification No. %notification%' FOLDERS IN CASE YOU NEED THEM AGAIN.
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
ECHO STATUS FOR "...KPD\Source documents pdf":
ECHO.
ECHO Файлы удаленные в "...KPD\Source documents pdf":>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
ECHO.>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
FOR /F "tokens=*" %%G IN (Del_temporary.txt) DO (IF EXIST "%CD%\KPD\# Source documents pdf\%%G.*" (ECHO !var_incr_del_pdf!^) *.pdf file for %%G found and deleted
MOVE "%CD%\KPD\# Source documents pdf\%%G.*" "%CD%\KPD\# Source documents pdf\Files updated on %date% by notification No. %notification%" >nul
ECHO !var_incr_del_pdf!^) *.pdf файл для документа %%G был найден>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
SET /a var_incr_del_pdf+=1
SET /a var_incr_del_pdf_suc+=1) ELSE (ECHO !var_incr_del_pdf!^) Failed to delete *.pdf file for %%G as it was not found in the folder
ECHO !var_incr_del_pdf!^) Не удалось удалить *.pdf файл для документа %%G, так как он не был найден>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
SET /a var_incr_del_pdf+=1))
ECHO.
ECHO Total files deleted: !var_incr_del_pdf_suc!
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
ECHO STATUS FOR "...KPD\Source documents docx, doc, xls, xlsx":
ECHO.
ECHO Файлы удаленные в папке "...KPD\Source documents docx, doc, xls, xlsx":>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
ECHO.>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
FOR /F "tokens=*" %%G IN (Del_temporary.txt) DO (IF EXIST "%CD%\KPD\# Source documents docx, doc, xls, xlsx\%%G.*" (ECHO !var_incr_del_docs!^) *.doc^(x^)/xls^(x^) file for %%G found and deleted
MOVE "%CD%\KPD\# Source documents docx, doc, xls, xlsx\%%G.*" "%CD%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%" >nul
ECHO !var_incr_del_docs!^) *.doc^(x^)/xls^(x^) файл для документа %%G был найден и удален>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
SET /a var_incr_del_docs+=1
SET /a var_incr_del_docs_suc+=1) ELSE (ECHO !var_incr_del_docs!^) Failed to delete *.doc^(x^)/xls^(x^) file for %%G as it was not found in the folder
ECHO !var_incr_del_docs!^) Не удалось удалить *.doc^(x^)/xls^(x^) файл для документа %%G, так как он не был найден>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
SET /a var_incr_del_docs+=1))
ECHO.
ECHO Total files deleted: !var_incr_del_docs_suc!
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
ECHO Файлы удаленные в папках "...KPD\Matches against previous version" и "...KPD\Source documents to be translated":>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
ECHO.>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
ECHO STATUS FOR "...KPD\Matches against previous version" and "...KPD\Source documents to be translated":
ECHO.
FOR /F "tokens=*" %%G IN (Del_temporary.txt) DO (IF EXIST "%CD%\KPD\# Matches against previous version\%%G.*" (ECHO !var_incr_del_both!^) *.doc^(x^)/xls^(x^) file for %%G found and deleted in '...KPD\Matches against previous version'
DEL "%CD%\KPD\# Matches against previous version\%%G.*"
ECHO !var_incr_del_both!^) *.doc^(x^)/xls^(x^) файл для документа %%G был найден и удален в папке "...KPD\Matches against previous version">>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
SET /a var_incr_del_both+=1
SET /a var_incr_del_both_suc+=1) ELSE (IF EXIST "%CD%\KPD\# Source documents to be translated\%%G.*" (ECHO !var_incr_del_both!^) *.doc^(x^)/xls^(x^) file for %%G found and deleted in '...\KPD\Source documents to be translated'
DEL "%CD%\KPD\# Source documents to be translated\%%G.*"
ECHO !var_incr_del_both!^) *.doc^(x^)/xls^(x^) файл для документа %%G был найден и удален в папке "...KPD\Source documents to be translated">>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
SET /a var_incr_del_both+=1
SET /a var_incr_del_both_suc+=1) ELSE (ECHO !var_incr_del_both!^) Failed to delete *.doc^(x^)/xls^(x^) file for %%G as it was not found in the folders
ECHO !var_incr_del_both!^) Не удалось удалить *.doc^(x^)/xls^(x^) файл для документа %%G, так как он не был найден ни в одной из папок>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Deletion.txt"
SET /a var_incr_del_both+=1)))
ECHO.
ECHO Total files deleted: !var_incr_del_both_suc!
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------
ENDLOCAL
ECHO.
ECHO DOCUMENTS HAVE BEEN DELETED^^! PRESS ANY BUTTON TO CONTINUE.
PAUSE >nul

:processing_without_removing_files
SET /a var_incr_upd_pdf=1
SET /a var_incr_upd_docs=1
SET /a var_incr_upd_both=1
SETLOCAL ENABLEDELAYEDEXPANSION
ECHO Обновление в папке Matches against previous version и Source documents to be translated:>"%CD%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\MAPV_and_SDTT.txt"
ECHO Обновление в папке Source documents pdf:>"%CD%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\PDF.txt"
ECHO Обновление в папке Source documents docx, doc, xls, xlsx:>"%CD%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\DOC.txt"
CLS
ECHO NOW, I WILL UPDATE DOCUMENTS SPECIFIED IN 'List.txt'.
ECHO COPIES OF THE OLD FILES WILL BE STORED IN "...KPD\Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%" IN CASE YOU NEED THEM.

ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------
ECHO STATUS FOR "...KPD\Source documents pdf":
ECHO.
CD "%CD%\KPD\# Source documents pdf"
FOR %%G IN ("%root_directory%\MemoQ import\*.pdf") DO IF EXIST %%~nxG (ECHO !var_incr_upd_pdf!^) %%~nxG was updated
MOVE "%CD%\%%~nxG" "%CD%\Files updated on %date% by notification No. %notification%" >nul
MOVE "%root_directory%\MemoQ import\%%~nxG" "%CD%" >nul
ECHO !var_incr_upd_pdf!^) Файл %%~nxG был найден и обновлен>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\PDF.txt"
SET /a var_incr_upd_pdf+=1) ELSE (ECHO !var_incr_upd_pdf!^) %%~nxG was just copied to the folder ^(=new file, released for the first time^)
ECHO This file was released for the first time.>"%CD%\Files updated on %date% by notification No. %notification%\%%~nG.txt"
ECHO As there was nothing to copy to the current folder, the script created this *.txt file so you know that this file is a new one.>>"%CD%\Files updated on %date% by notification No. %notification%\%%~nG.txt"
ECHO !var_incr_upd_pdf!^) Файл %%~nxG был просто скопирован в ...KPD\Source documents pdf, так как он был опубликован впервые>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\PDF.txt"
MOVE "%root_directory%\MemoQ import\%%~nxG" "%CD%" >nul
SET /a var_incr_upd_pdf+=1)
ECHO.
SET /a var_incr_upd_pdf-=1
ECHO Total files updated: !var_incr_upd_pdf!
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------

ECHO STATUS FOR "...KPD\Source documents docx, doc, xls, xlsx":
ECHO.
CD "%root_directory%\KPD\# Source documents docx, doc, xls, xlsx"
FOR %%G IN ("%root_directory%\MemoQ import\*") DO IF EXIST %%~nG.* (ECHO !var_incr_upd_docs!^) %%~nxG was updated
MOVE "%CD%\%%~nG.*" "%CD%\Files updated on %date% by notification No. %notification%" >nul
COPY "%root_directory%\MemoQ import\%%~nG.*" "%CD%" >nul
ECHO !var_incr_upd_docs!^) Файл %%~nxG был найден и обновлен>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\DOC.txt"
SET /a var_incr_upd_docs+=1) ELSE (ECHO !var_incr_upd_docs!^) %%~nxG was just copied to the folder ^(=new file, released for the first time^)
ECHO This file was released for the first time.>"%CD%\Files updated on %date% by notification No. %notification%\%%~nG.txt"
ECHO As there was nothing to copy to the current folder, the script created this *.txt file so you know that this file is a new one.>>"%CD%\Files updated on %date% by notification No. %notification%\%%~nG.txt"
ECHO !var_incr_upd_docs!^) Файл %%~nxG был просто скопирован в ...KPD\Source documents docx, doc, xls, xlsx, так как он был опубликован впервые>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\DOC.txt"
COPY "%root_directory%\MemoQ import\%%~nG.*" "%CD%" >nul
SET /a var_incr_upd_docs+=1)
ECHO.
SET /a var_incr_upd_docs-=1
ECHO Total files updated: !var_incr_upd_docs!
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------

ECHO STATUS FOR "...KPD\Matches against previous version" and "...KPD\Source documents to be translated":
ECHO.
CD "%root_directory%\KPD\# Matches against previous version"
FOR %%G IN ("%root_directory%\MemoQ import\*") DO IF EXIST %%~nG.* (ECHO !var_incr_upd_both!^) %%~nxG was found and updated in '...KPD\Matches against previous version'
DEL %%~nG.*
COPY "%root_directory%\MemoQ import\%%~nG.*" "%root_directory%\KPD\# Source documents to be translated" >nul
ECHO !var_incr_upd_both!^) Ôàéë %%~nxG áûë íàéäåí è îáíîâëåí â ïàïêå ...KPD\Matches against previous version>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\MAPV_and_SDTT.txt"
SET /a var_incr_upd_both+=1) ELSE (IF EXIST "%root_directory%\KPD\# Source documents to be translated\%%~nG.*" (ECHO !var_incr_upd_both!^) %%~nxG was found and updated in '...KPD\Source documents to be translated'
COPY "%root_directory%\MemoQ import\%%~nG.*" "%root_directory%\KPD\# Source documents to be translated" >nul
ECHO !var_incr_upd_both!^) Ôàéë %%~nxG áûë íàéäåí è îáíîâëåí â ïàïêå ...KPD\Source documents to be translated>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\MAPV_and_SDTT.txt"
SET /a var_incr_upd_both+=1) ELSE (ECHO !var_incr_upd_both!^) %%~nxG was just copied to '...KPD\Source documents to be translated' ^(=new file, released for the first time^)
ECHO !var_incr_upd_both!^) Ôàéë %%~nxG áûë ïðîñòî ñêîïèðîâàí â ...KPD\Source documents to be translated, òàê êàê îí áûë îïóáëèêîâàí âïåðâûå>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\MAPV_and_SDTT.txt"
COPY "%root_directory%\MemoQ import\%%~nG.*" "%root_directory%\KPD\# Source documents to be translated" >nul
SET /a var_incr_upd_both+=1))
ECHO.
SET /a var_incr_upd_both-=1
ECHO Total files updated: !var_incr_upd_both!
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------
ENDLOCAL


CD "%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%"
IF EXIST Deletion.txt (ECHO Íèæå ïðåäñòàâëåí ñïèñîê ôàéëîâ, êîòîðûå áûëè óäàëåíû âî âñåõ ïîäïàïêàõ "...\KPD\...".>Report.txt
ECHO Êîïèè óäàëåííûõ ôàéëîâ áûëè ïåðåìåùåíû â ñîîòâåòñâóþùèå "Files updated on %date% by notification No. %notification%" ïàïêè.>>Report.txt
ECHO.>>Report.txt
FOR /F "tokens=*" %%G IN (Deletion.txt) DO (ECHO %%G>>Report.txt)) ELSE (ECHO Íè îäèí èç ôàéëîâ â 'List.txt' íå áûë ïîìå÷åí êàê "REM" ^(=óäàëèòü èç âñåõ "...\KPD\..." ïîäïàïîê^).>Report.txt)
ECHO =====================================================================================================================================================================================================>>Report.txt
ECHO =====================================================================================================================================================================================================>>Report.txt
ECHO Íèæå ïðåäñòàâëåí ñïèñîê ôàéëîâ, êîòîðûå áûëè îáíîâëåíû âî âñåõ ïîäïàïêàõ "...\KPD\...".>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Report.txt"
ECHO Êîïèè îáíîâëåííûõ ôàéëîâ áûëè ïåðåìåùåíû â ñîîòâåòñâóþùèå "Files updated on %date% by notification No. %notification%" ïàïêè.>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Report.txt"
ECHO.>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Report.txt"
FINDSTR /M /C:"1)" PDF.txt >nul
IF %ERRORLEVEL% EQU 0 (FOR /F "tokens=*" %%G IN (PDF.txt) DO ECHO %%G>>Report.txt) ELSE (ECHO Ôàéëû äîêóìåíòîâ óêàçàííûõ â List.txt íå áûëè íàéäåíû â ïàïêå Source documents pdf>>Report.txt)
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Report.txt"
FINDSTR /M /C:"1)" DOC.txt >nul
IF %ERRORLEVEL% EQU 0 (FOR /F "tokens=*" %%G IN (DOC.txt) DO ECHO %%G>>Report.txt) ELSE (ECHO Ôàéëû äîêóìåíòîâ óêàçàííûõ â List.txt íå áûëè íàéäåíû â ïàïêå Source documents docx, doc, xls, xlsx>>Report.txt)
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------>>"%root_directory%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Report.txt"
FINDSTR /M /C:"1)" MAPV_and_SDTT.txt >nul
IF %ERRORLEVEL% EQU 0 (FOR /F "tokens=*" %%G IN (MAPV_and_SDTT.txt) DO ECHO %%G>>Report.txt) ELSE (ECHO Ôàéëû äîêóìåíòîâ óêàçàííûõ â List.txt íå áûëè íàéäåíû íè â ïàïêå Matches against previous version, íè â ïàïêå Source documents to be translated>>Report.txt)
ECHO =====================================================================================================================================================================================================>>Report.txt
ECHO =====================================================================================================================================================================================================>>Report.txt
DEL PDF.txt
DEL DOC.txt
DEL MAPV_and_SDTT.txt
IF EXIST Deletion.txt DEL Deletion.txt
CD %root_directory%
IF EXIST "%root_directory%\KPD\MemoQ import" (RMDIR /s /q "%root_directory%\KPD\MemoQ import"
MOVE "MemoQ import" "%root_directory%\KPD" >nul) ELSE (MOVE "MemoQ import" "%root_directory%\KPD" >nul)
FOR %%G IN ("%CD%\KPD\MemoQ import\*.*") DO ECHO %%~nG>>Upd_check.txt
FOR /F "delims=*" %%G in (Upd_temporary.txt) do call :doWork "%%G"
GOTO :end_of_script
:doWork
FINDSTR /C:"%~1" Upd_check.txt >nul
IF %ERRORLEVEL% EQU 1 (ECHO %~1>>Not_found.txt)
GOTO :EOF
:end_of_script
SETLOCAL ENABLEDELAYEDEXPANSION
SET /a var_incr_not_found=1
IF EXIST Not_found.txt (ECHO Äîêóìåíòû, êîòîðûå íå áûëè îáíîâëåíû, òàê êàê èõ ôàéëû îòñòóòñòâîâàëè íà ñåðâåðå:>>"%CD%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Report.txt"
ECHO Documents that were not updated as their files are missing on the server:
ECHO.
ECHO.>>"%CD%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Report.txt"
FOR /F "tokens=*" %%G IN (Not_found.txt) DO (ECHO !var_incr_not_found!^) %%G>>"%CD%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Report.txt"
ECHO !var_incr_not_found!^) %%G
SET /a var_incr_not_found+=1)
ECHO ------------------------------------------------------------------------------------------------------------------------------------------------------
ECHO =====================================================================================================================================================================================================>>"%CD%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Report.txt"
ECHO =====================================================================================================================================================================================================>>"%CD%\KPD\# Source documents docx, doc, xls, xlsx\Files updated on %date% by notification No. %notification%\Report.txt")
ENDLOCAL
IF EXIST Del_temporary.txt DEL Del_temporary.txt
DEL Upd_temporary.txt
DEL Upd_check.txt
IF EXIST Not_found.txt DEL Not_found.txt
ECHO.
ECHO I AM DONE^^!
ECHO.
ECHO I PUT NEW FILES IN A SEPARATE FOLDER CALLED 'MemoQ import' ^(see "...KPD\MemoQ import"^), SO YOU DON'T HAVE ANY PROBLEMS IMPORTING THEM.
ECHO.
ECHO PRESS ANY BOTTON TO EXIT.
PAUSE >nul
cls
:closing
ECHO ONE MORE QUESTION BEFORE CLOSING THE SCRIPT...
ECHO DO YOU WANT TO DELETE 'List.txt'?
ECHO.
ECHO TYPE 1 TO DELETE IT
ECHO TYPE 2 TO LEAVE IT AS IT IS
ECHO.
SET /P inputchoice_closing=Make your choice: 
IF "%inputchoice_closing%"=="" GOTO inputchoice_closing_wronginput
IF "%inputchoice_closing%"=="1" (DEL List.txt & GOTO exit_script)
IF "%inputchoice_closing%"=="2" (GOTO exit_script) ELSE (GOTO inputchoice_closing_wronginput)
:inputchoice_closing_wronginput
CLS
ECHO Please, type a number in the range from 1 to 2.
GOTO :closing
