function Get-Sheet
{
    <#
    .SYNOPSIS
    Retrieve sheets from files that can be opened with Excel.

    .DESCRIPTION
    The Get-Sheet cmdlet retrive sheets from files that can be opened with Excel.

    .PARAMETER File 
    Specifies a Excel file path or FileInfo object.

    .PARAMETER Visible
    Determines whether the Excel application is visible.

    .INPUTS
    System.IO.FileInfo

    .OUTPUTS
    __ComObject (Worksheet Object)

    .EXAMPLE
    Get-ChildItem "*.xls" | Get-Sheet | ?{ $_.Name -eq "Sheet1" } | Get-Range "A1:C5,E1:F5"
    
    .NOTES
    This cmdlet is using COM Automation, so EXCEL.EXE process will create with cmdlet is executed, and end with cmdlet is quit.

    .LINK
    Get-Range
    #>
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [System.IO.FileInfo] $File,
        [Parameter(Mandatory=$false)]
        [switch] $Visible
    )
    begin
    {
        $local:excel = New-Object -ComObject Excel.Application
        $local:excel.Visible = $Visible
        $local:excel.DisplayAlerts = $false
    }
    process
    {
        $book = $local:excel.Workbooks.Open($File.FullName, 0, $true)
        $book.Sheets
        $book.Close($false)
    }
    end
    {
        $local:excel.Quit()
    }
}


function Get-Range
{
    <#
    .SYNOPSIS
    Gets values of specified range from Excel sheets.

    .DESCRIPTION
    This cmdlet require Get-Sheet cmdlet's output.

    .PARAMETER Sheet 
    Specifies a Excel sheet object.

    .PARAMETER Range
    Specifies a string that represents a cell or a range of cells.
    This must be an A1-style.
    
    .PARAMETER IncludeSheetName
    Determines whether add the sheet name to the retrieved data.
    
    .PARAMETER HeaderRow
    Specifies a row number of headers, which is the property name of the retrieved data.

    .INPUTS
    __ComObject (Worksheet Object from Get-Sheet cmdlet)

    .OUTPUTS
    PSObject

    .EXAMPLE
    Get-ChildItem "*.xls" | Get-Sheet | ?{ $_.Name -eq "Sheet1" } | Get-Range "A1:C5,E1:F5"
    
    .LINK
    Get-Range
    #>
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [__ComObject] $Sheet,
        [Parameter(Mandatory=$true, Position=1)]
        [string] $Range,
        [Parameter(Mandatory=$false)]
        [switch] $IncludeSheetName,
        [Parameter(Mandatory=$false)]
        [int] $HeaderRow
    )
    process
    {
        foreach($sht in $Sheet)
        {
            $areas = $sht.Range($Range).Areas
            $lastRowOfSheet = $sht.Cells.Find("*", $sht.Range("A1"), -4163, 2, 1, 2).Row
            $lastColumnOfSheet = $sht.Cells.Find("*", $sht.Range("A1"), -4163, 2, 2, 2).Column
            
            $rowIndexes = @()
            $columnIndexes = @()
            $headers = @{}
            $psAreas = @()
            $returnValue = @()

            foreach($area in $areas)
            {
                $value = $null
                if($area.Rows.Count -eq 1 -and $area.Columns.Count -eq 1)
                {
                    $value = New-Object "object[,]" 2,2
                    $value[1,1] = $area.Value()
                }
                else
                {
                    $value = $area.Value()
                }

                $firstRow = $area.Row
                $lastRow = [Math]::Min($firstRow + $area.Rows.Count - 1, $lastRowOfSheet)
                $firstColumn = $area.Column
                $lastColumn = [Math]::Min($firstColumn + $area.Columns.Count - 1, $lastColumnOfSheet)

                $psArea = New-Object PSObject
                $psArea | Add-Member NoteProperty Value $value
                $psArea | Add-Member NoteProperty FirstRow $firstRow
                $psArea | Add-Member NoteProperty LastRow $lastRow
                $psArea | Add-Member NoteProperty FirstColumn $firstColumn
                $psArea | Add-Member NoteProperty LastColumn $lastColumn
                $psAreas += $psArea

                $rowIndexes += (@($firstRow .. $lastRow) -ne $headerRow)
                $columnIndexes += @($firstColumn .. $lastColumn)
            }
            
            foreach($c in $columnIndexes | sort -Unique)
            {
                $headers[$c] = ($sht.Cells.Item(1, $c).Address($true, $false) -split '\$')[0]

                if($HeaderRow -gt 0)
                {
                    $text = $sht.Cells.Item($HeaderRow, $c).Value()
                    if($text -ne $null)
                    {                    
                        if($headers[$headers.Keys -ne $c] -contains $text)
                        {
                            $text = "$text`_$c"
                        }
                        $headers[$c] = $text
                    }
                }
            }
 
            $rowData = @()

            foreach($r in $rowIndexes | sort -Unique)
            {
                $pso = New-Object PSObject
                if($IncludeSheetName)
                {
                    $pso | Add-Member NoteProperty "Sheet" $sht.Name
                }
                
                foreach($c in $columnIndexes | sort -Unique)
                {
                    $propertyName = $headers[$c]
                    $pso | Add-Member NoteProperty $propertyName $null
                    $psAreas |
                        ?{ $r -ge $_.FirstRow -and $r -le $_.LastRow -and $c -ge $_.FirstColumn -and $c -le $_.LastColumn } |
                        %{ $pso.$propertyName = $_.Value[($r - $_.FirstRow + 1), ($c - $_.FirstColumn + 1)] }
                }
                $rowData += $pso
            }
            $returnValue += ,$rowData
        }
        $returnValue
    }
}
