Prerequisites:
 Install and import all the packages listed:
	VaderSentiment - download from the attachment(Vader) / go to the downloaded location and type the code "python setup.py install" in cmd promt.
	NLTK - http://www.nltk.org/install.html
	Stanford NER - http://nlp.stanford.edu/software/CRF-NER.shtml
	Geocoder - https://pypi.python.org/pypi/geocoder
	RE - https://docs.python.org/2/library/re.html
	Counter - https://pypi.python.org/pypi/Counter/1.0.0
	XLRD - https://pypi.python.org/pypi/xlrd
	Xlutils - pip install Xlutils
	tomcat server - download from the attachment

Procedure:
Step 1: After installing all the packages run the tomcat server by double clicking on the "startup.bat" file
	for windows and "startup.sh" for IOS.

Step 2:	Open the code from the source file "iso.py" from the attachment.

Step 3: Change the sheet location in the source code to the location of the excel sheet which contains the fraudulent articles.
	example : file_location = "C:/Users/Textbook/Desktop/Class Notes/ISO.xlsx"

Step 4: Run the program and select the sheet number where 1 denotes the first sheet and 2 denotes the second sheet likewise.
	and enter the excel cell number which contains the article for testing. 

Step 5: The output of the result is provided in 3 locations:
		- Excel sheet -> The source code creats a new excel sheet with the name "Verisk iso" with
		  the same records as the original excel sheet but the output will be displayed from column B to Column E
		  where B represents the Fradulant Name, C represents the Location, D represents the government officials
		  or the innocent persons involved and the Column E represents the Category.
		- Text File -> The code also creats a text file by the name "Output.txt" which the case number and the
		  fraudulent name, location and the category.
		- Console -> The source code also prints the output in the console.
Note: For more information read the input-output word document.

	



