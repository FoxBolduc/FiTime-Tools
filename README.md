# FiTime-Tools
Tools for Dr. Zumwalt's FiTime App

Purpose:
The python script jconvert.py is intended to convert data from Excel spreadsheets into a JSON format. This allows people unfamiliar with JSON
to contribute to the database in a way they find familiar.

Use Case:
Jconvert can be used to make any JSON structure that consists of an arbitrarily long list of elements, such that each element has a consistent 
number of attributes. In the example provided the list generated contains "programmers" as elements and each programmer has a first name, last name,
and a job.

Use:
***********************NOTE*******************************
DO NOT DELETE OR MOVE RED TEXT FROM THE EXCEL SPREADSHEET
**********************************************************
The Excel spreadsheet should be based on the template file provided in this repository named jcontemplate.xlsx. The file may be renamed anything
so long as it remains a .xlsx file. Once information is entered, run the python script from the command line with the command:

python jconvert.py <filename>

Note the filename should not contain the .xlsx extension, and that if no filename is entered, jconvert shall simply convert jcontemplate.xlsx.
