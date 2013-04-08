PSExcel
=======
This is a powershell script to get values from excel file.


Get-Sheet
---------
The Get-Sheet cmdlet gets the sheets in Excel file by using COM Automation.
EXCEL.EXE process will create with cmdlet is executed, and quit.  


SYNTAX

	Get-Sheet [-File] <System.IO.FileInfo[]> [-Visible]


PARAMETERS

	-File <System.IO.FileInfo[]>
		Specifies a Excel file path or FileInfo object.
		This parameter can take values from incoming pipeline objects.

	-Visible <SwitchParameter>
		Determines whether the Excel application is visible.


Get-Range
---------
The Get-Sheet cmdlet gets the sheets in excel file by using COM Automation.
excel.exe process will create with cmdlet is executed, and quit.


SYNTAX

	Get-Range [-Sheet] <__ComObject> [-Range] <string> [-IncludeSheetName] [-HeaderRow <int>]


PARAMETERS

	-Sheet
		Specifies a Excel sheet object.
		This parameter can take values from incoming pipeline objects.

	-Range
		Specifies a string that represents a cell or a range of cells.
		This must be an A1-style.

	-IncludeSheetName
		Determines whether add the sheet name to the retrieved data.

	-HeaderRow
		Specifies a row number of headers, which is the property name of the retrieved data.


###Example###

	Get-ChildItem "*.xls" | Get-Sheet | ?{ $_.Name -eq "Sheet1" } | Get-Range "A1:C5,E1:F5"
