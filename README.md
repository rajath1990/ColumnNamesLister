# VBA-Automation
List Column Headers 
This macro will list the column names/headers of by opening each excel file one-by-one as listed in the input .txt file. This is a generic macro and will work with MS Excel versions 2003,2007,2010,2013 and 2016.
General Rules
1)	To list all the excel file names (inside a folder) in excel, copy-paste the folder path in any web browsers. I would prefer chrome/firefox. Select All contents from browser window, open MS Excel and Paste as text. Remove the formatting (hyperlinks, .xlsx and .xls extensions) from the MS Excel path. Now copy-paste this content into a text file which is the input for this Excel VBA Macro
2)	This macro will open each excel file based on the name mentioned in the .txt file. If the name is not matching, the file will not be opened.
3)	The column headers can be listed either horizontally or vertically.   
a.	Sheet name : Horizontal List will list the column header names horizontally
b.	Sheet name: Vertical List will list the column header names vertically
Note : DO NOT CHANGE/RENAME THE SHEET NAMES or CONTENTS OF THE SHEETS. (eg : The first row contains ‘File Names’ and ‘Column Names’.
How to use the macro
1.	.txt File Path : Windows File dialog will open up to select your file path
2.	Excel Folder Path : Windows Folder dialog will open to browse your working folder path(where your physical excel files are present)
3.	Check-box – Vertical or Horizontal based on your requirement
4.	Execute: To Run the macro
5.	Cancel: To cancel operation and exit.

 

