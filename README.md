# Word Doc Generator
This small script takes data from an Excel workbook, extract relevant information and generate a letter for printing in Word Doc format.

See "template.docx" to get a sense of what the final output should look like. For each "<< SomeDataHere >>" in the template, we get it from Excel workbook "input_database.xlsx"

Since there are 1015 rows in "input_database.xlsx", there are 1015 letters generated in Word doc format, ready for printing. This is why "output.docx" has 1015 pages.

For more details, see comments in "main.py". With the help of Python libraries "docx" and "pandas", this script was completed in less than 2 hours. 