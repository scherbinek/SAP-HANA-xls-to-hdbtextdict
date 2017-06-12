SAP HANA *.xls to *.hdbtextdict

Simple Python Script to convert a more readable Excel Sheet (*.xls) to a SAP HANA Text Dictionary Object (*.hdbtextdict)

I created a very simple script for creating a *.hdbtextdict by a maintained Excel Sheet. Because it is really annoying to maintain a larger XML-like file by hand and Excel has some nice features (as filtering..).

Just a few simple steps to run the script:
1. Install Python and necessary dependencies (e.g. 'pip install xlrd')
2. Modify the XLS Template 'custom-dictionary-template.xlsx'
3. Run the Script from Command Line by 'py custom-dictionary-template.py'
4. Add the created Dictionary to your SAP HANA TA Configuration File and Upload the created file to your TA folders
