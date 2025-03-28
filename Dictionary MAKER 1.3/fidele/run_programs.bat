@ECHO OFF
REM PROGRAM ONE --------------------------------------------------
ECHO Running program 1 ...
REM LIST ALL WORD DOCUMENTS IN FOLDER
ECHO Here are all the Word documents in the fidele folder ...
:start_program1
ECHO.
DIR *.docx /B
ECHO.
SET /P dups=What Word document (in the fidele folder) would you like to run Program 1 on (underline redundant lines and make frequency table)? Press the 'Tab' key to autocomplete. 
REM CHECK TO SEE IF %DUPS% IS A REAL FILE
IF NOT EXIST %dups% (
	ECHO %dups% is not in the fidele folder.
	GOTO :start_program1
	)
REM EXECUTE PROGRAM ONE
python program1.py %dups% redundantLines.docx
ECHO Opening Word file with redundant lines underlined ...
redundantLines.docx
ECHO Opening Excel file with frequency table ...
frequencyTable.xlsx
REM PROGRAM TWO --------------------------------------------------
ECHO Running program 2 ...
REM LIST ALL WORD DOCUMENTS IN FOLDER
ECHO Here are all the Word documents in the fidele folder ...
:start_program2
ECHO.
DIR *.docx /B
ECHO.
SET /P phrases=What Word document (in the fidele folder) would you like to run Program 2 on (identify and sort by reference words)? Press the 'Tab' key to autocomplete. 
REM CHECK TO SEE IF %PHRASES% IS A REAL FILE
IF NOT EXIST %phrases% (
	ECHO %phrases% is not in the fidele folder.
	GOTO :start_program2
	)
REM EXECUTE PROGRAM TWO
python program2.py %phrases% frequencyTable.xlsx
ECHO Opening Excel file with reference table ...
referenceTable.xlsx
