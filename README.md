PSExcel
=======
This is a powershell script to get values from excel file.


Get-Sheet
---------
Retrieve sheets from files that can be opened with Excel.


SYNTAX

	Get-Sheet [-File] <FileInfo> [-Visible] [<CommonParameters>]


PARAMETERS

	-File <System.IO.FileInfo>
		Specifies a Excel file path or FileInfo object.

	-Visible <SwitchParameter>
		Determines whether the Excel application is visible.


Get-Range
---------
Gets values of specified range from Excel sheets.


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
