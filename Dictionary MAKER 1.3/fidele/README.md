# fidele
Code for Swahili Dictionary Project

# HOW TO DOWNLOAD ONTO A NEW COMPUTER:
1. open **first_download.bat** at https://github.com/noahmclean1/fidele/blob/master/first_download.bat
2. copy all the text in the file
3. open **Notepad** on your computer. you can find it by typing "notepad" into the search bar at the bottom left 
4. paste the text from **first_download.bat** into Notepad
5. click **File** in the top left
6. click **Save As**
7. where it says **File name:** at the bottom, type "first_download.bat"
8. click **Save** on the bottom right. this will create a batch file that will run the code for you
9. go to where you saved the file and double-click it

# HOW TO RUN THE PROGRAMS (AFTER CREATING **first_download.bat**):
1. put the files that you want to use (e.g. English_entries_with_DUPLICATES_WORD_LIST.docx, English_SERIES_alpha_TEXT.docx) inside the folder named **fidele**
2. open **fidele**
3. click on **run_programs.bat**. this will run the programs for you, and it will ask you the following questions:

`What Word document (in the fidele folder) would you like to run Program 1 on (underline redundant lines and make frequency table)?`

Enter the name of the document making sure to have the .docx extension (e.g. **English_entries_with_DUPLICATES_WORD_LIST.docx**).

This will start running Program 1. 

It will open the new Word document named **redundantLines.docx** with underlined words. To continue with the program, close the Word document. (If you need to find it, it is in the **fidele** folder.)

The program will then open the new Excel file named **frequencyTable.xlsx**. Again, to continue with the program, close the Excel file. (If you need to find it, it is in the **fidele** folder.)

`What Word document (in the fidele folder) would you like to run Program 2 on (identify and sort by reference words)?`

Enter the name of the document making sure to have the .docx extension (e.g. **English_SERIES_alpha_TEXT.docx**).

This will start running Program 2. 

It will open the new Word document named **referenceTable.xlsx**. (If you need to find it, it is in the **fidele** folder.)


# If you are working in Terminal:
Setup:
```
git clone https://github.com/noahmclean1/fidele.git
cd fidele\
pip3 install python-docx --user
pip3 install xlsxwriter --user
```

Program 1:
```
python3 program1.py English_entries_with_DUPLICATES_WORD_LIST.docx redundantLines.docx
open redundantLines.docx
open frequencyTable.docx
```

Program 2:
```
pip3 install xlrd --user
pip3 install xlwt --user
python3 program2.py English_SERIES_alpha_TEXT.docx frequencyTable.xlsx
open referenceTable.xlsx
```

To get permissions to execute shell file:
```
chmod u+x write_dictionary.sh
```
SOURCE: https://stackoverflow.com/questions/17015449/how-do-i-run-sh-or-bat-files-from-terminal
