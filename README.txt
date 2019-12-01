Instructions to run our code:

the openpyxl library is needed to run the code!




Here is what you need to consider while giving inputs:

Do you want to create a new Sheet (N) or add Transactions to an old one (O)?
- Inputting "N" will create a new sheet
	-If the file "Monthly_Transactions.xlsx" doesn't yet exist in the same folder as the python file you have to choose "N" to create the file. If the file already exists, choosing "N" will add a new 	sheet to the existing file.
- Inputting "O" will allow you to add transactions to an existing sheet
	-"O" can only be chosen if the file "Monthly_Transacrtions.xlsx" already exists in the same folder as the python file. Choosing "O" allows you to open a sheet that already exists and add 		transactions to that sheet.



Naming a sheet or opening a sheet: 3 letters representing a month (JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OCT, NOV, DEC) and 2 numbers representing a year. For example December 2019 would be DEC19.

Starting Balance: must be an integer or a float

Date of the transaction: Enter only the day (Month and year are already known). The date must really exist. For example if your month is February and you give 30 as your input it will be deemed invalid.

Item Description: No restrictions. Can be anything.

Price: Must be an integer or a float. If you want to deduct from your balance insert a negative number. If a positive number is inserted it will be added to your balance.
