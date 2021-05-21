AUTO_HARDWARE_DOC
	Automates "Highstreet IT Computer/Software Asset Receipt" document creation. 

	Takes data from "input.xlsx" and generates documents using the "form.docx" file. 


PREREQ:
	-python installed 

	-pip packages "openpyxl" and "docx-mailmerge" installed

	-properly formated "input.xlsx" file
		-first line of data starts on row 3
		-data columns range from column B to F

	-properly formated "form.docx" file with mailmerge fields
		-NAME
		-DATE
		-ASSET
		-SERIAL_NUM
		-LIST

	-no output file in directory 


RUN:
	-check if python and pip packages are installed
		-go to "https://www.python.org/downloads/" to download python 
		-can run "setup.cmd" to install packages 	

	-enter data into "input.xlsx"
	
	-double click on main.py 

	-remove/delete/rename output file after use 