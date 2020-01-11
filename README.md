# QBO_Bank
To find any missing transactions between the bank statements and what is in QuickBooks


IN ORDER TO RUN PLEASE DOWNLOAD IF YOU HAVE NOT AL READY:

Both are in order to read and write to an excel worksheet:

Please install the latest version of python. As so you will have the latest version of pip.

pip install XlsxWriter

pip install openpyxl

Also very useful:

pip install pandas

Once you have done that please create an Excel document that contains: the date, in general or short date format, the Deposit, and the Payment. For both in Quickbooks sheet and in the same document but in a different sheet have the same requirements in the Bank statement sheet in order for the program to search for any differenances between the two statements.

Once you have ran the program there will be a file next to the orginial file as to not destroy the original file, that will contian the same content as the original file but with an additional column at the end to signify "Match" as in it found the same payment in the other file within a 2 day search leway. But if not found it will display a "Not Match"


To run:

python findMissing.py <-Draw> <Filename.xlsx>

-Draw is opitional
