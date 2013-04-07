PSExcel
=======

This is a powershell script to get values from excel file.

------------------------------------------------------------------
NAME
	Get-Sheet

SYNOPSIS
	Gets the sheets in Excel file.

SYNTAX
	Get-Sheet [-File] <System.IO.FileInfo[]> [-Visible]

DESCRIPTION
	The Get-Sheet cmdlet gets the sheets in excel file by using COM Automation.
	excel.exe process will create with cmdlet is executed, and quit.

PARAMETERS
	-File <System.IO.FileInfo[]>
		Specifies a Excel file path or FileInfo object.
		This parameter can take values from incoming pipeline objects.

	-Visible <SwitchParameter>
		Determines whether the Excel application is visible.

RELATED LINKS
	Get-Range


------------------------------------------------------------------
NAME
	Get-Range

SYNOPSIS
	Gets the values of range in Excel sheet.

SYNTAX
	Get-Range [-Sheet] <__ComObject> [-Range] <string> [-IncludeSheetName] [-HeaderRow <int>]

DESCRIPTION
	The Get-Sheet cmdlet gets the sheets in excel file by using COM Automation.
	excel.exe process will create with cmdlet is executed, and quit.

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

RELATED LINKS
	Get-Sheet
