PO Detail Final.xlsm

	Purchase order containing a list of items bought and a database with information corresponding to those items.
	When printed to PDF, it creates a neatly and properly formatted purchase order.
    	Only the first sheet is printed, and the button is not included in the print.

Sheet 1 - Purchase Order

	Contains a list of all items purchased and all necessary information about those items.
	The first 22 lines are the header, which will be repeated on every page.
   		Header contains Brookfield Properties address, P.O. number, order date, Ship To and Bill To addresses, delivery instructions, among other information.
	Each entry contains a field for line number, item number, name, description, custom, unit/measurement, quantity, due date, price per unit, and total cost.
	Each new entry should be created by copying and pasting the previous entry and then modifying it to match the new entry.
    	Can clear the contents of the new entry after copying and pasting the old entry if that is easier. However, the Line and Total columns should never be cleared or manually modified.
	The Total column automatically calculates the value of the quantity multiplied by the price per item.
    	The last entry in the Total column calculates the sum of all the previous cost totals.
	The button in the header, when clicked, will fill in the Name, Description, Custom, and U/M fields with information from the Item Database based on the given Item Number (assuming that the item number is in the database).
    	The button will never overwrite a cell that already has data in it.
    	The button will fill every empty entry that contains an Item Number at once, allowing for even more efficiency.

Sheet 2 - Item Database

	Contains the information that is copied to the Purchase Order sheet when the script button is pressed.
	Formatted with Item Number, Name, Description, Custom, and Unit/Measure columns.
	Can be modified and added to by the user with any items that should be in the purchase order.


PO Autofill Script.bas

	Contains the VBA script that is run whenever the button in PO Detail Final.xlsm is pressed.


PO Output Final.pdf

	The final version of the purchase order that is generated automatically when printing the spreadsheet.
