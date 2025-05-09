﻿#Requires -Modules Pester
Import-Module $PSScriptRoot\..\PSWriteExcel.psd1 -Force

### Preparing Data Start
$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason"; age = 42; info = "Food lover"}
)

$myitems1 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"}
)
$myitems2 = [PSCustomObject]@{
    name = "Joe"; age = 32; info = "Cat lover"
}

$InvoiceEntry1 = @{}
$InvoiceEntry1.Description = 'IT Services 1'
$InvoiceEntry1.Amount = '$200'

$InvoiceEntry2 = @{}
$InvoiceEntry2.Description = 'IT Services 2'
$InvoiceEntry2.Amount = '$300'

$InvoiceEntry3 = @{}
$InvoiceEntry3.Description = 'IT Services 3'
$InvoiceEntry3.Amount = '$288'

$InvoiceEntry4 = @{}
$InvoiceEntry4.Description = 'IT Services 4'
$InvoiceEntry4.Amount = '$301'

$InvoiceEntry5 = @{}
$InvoiceEntry5.Description = 'IT Services 5'
$InvoiceEntry5.Amount = '$299'

$InvoiceData1 = @()
$InvoiceData1 += $InvoiceEntry1
$InvoiceData1 += $InvoiceEntry2
$InvoiceData1 += $InvoiceEntry3
$InvoiceData1 += $InvoiceEntry4
$InvoiceData1 += $InvoiceEntry5

$InvoiceData2 = $InvoiceData1.ForEach( {[PSCustomObject]$_})

$InvoiceData3 = @()
$InvoiceData3 += $InvoiceEntry1

$InvoiceData4 = $InvoiceData3.ForEach( {[PSCustomObject]$_})
### Preparing Data End

$Object1 = Get-Process | Select-Object ProcessName, Handle, StartTime -First 5
$Object2 = Get-PSDrive | Where { $_.Provider -like '*Registry*' -or $_.Provider -like '*Environment*' -or $_.Provider -like '*FileSystem*' }
$Object3 = Get-PSDrive | Where { $_.Provider -like '*Registry*' -or $_.Provider -like '*Environment*' -or $_.Provider -like '*FileSystem*'} | Select-Object * -First 2
$Object4 = Get-PSDrive | Where { $_.Provider -like '*Registry*' -or $_.Provider -like '*Environment*' -or $_.Provider -like '*FileSystem*'} | Select-Object * -First 1

$obj = New-Object System.Object
$obj | Add-Member -type NoteProperty -name Name -Value "Ryan_PC"
$obj | Add-Member -type NoteProperty -name Manufacturer -Value "Dell"
$obj | Add-Member -type NoteProperty -name ProcessorSpeed -Value "3 Ghz"
$obj | Add-Member -type NoteProperty -name Memory -Value "6 GB"


$myObject2 = New-Object System.Object
$myObject2 | Add-Member -type NoteProperty -name Name -Value "Doug_PC"
$myObject2 | Add-Member -type NoteProperty -name Manufacturer -Value "HP"
$myObject2 | Add-Member -type NoteProperty -name ProcessorSpeed -Value "2.6 Ghz"
$myObject2 | Add-Member -type NoteProperty -name Memory -Value "4 GB"


$myObject3 = New-Object System.Object
$myObject3 | Add-Member -type NoteProperty -name Name -Value "Julie_PC"
$myObject3 | Add-Member -type NoteProperty -name Manufacturer -Value "Compaq"
$myObject3 | Add-Member -type NoteProperty -name ProcessorSpeed -Value "2.0 Ghz"
$myObject3 | Add-Member -type NoteProperty -name Memory -Value "2.5 GB"

$myArray1 = @($obj, $myobject2, $myObject3)
$myArray2 = @($obj)


$InvoiceEntry7 = [ordered]@{}
$InvoiceEntry7.Description = 'IT Services 4'
$InvoiceEntry7.Amount = '$301'

$InvoiceEntry8 = [ordered]@{}
$InvoiceEntry8.Description = 'IT Services 5'
$InvoiceEntry8.Amount = '$299'

$InvoiceDataOrdered1 = @()
$InvoiceDataOrdered1 += $InvoiceEntry7

$InvoiceDataOrdered2 = @()
$InvoiceDataOrdered2 += $InvoiceEntry7
$InvoiceDataOrdered2 += $InvoiceEntry8
<# Useful to display types
$Array = @()
$Array += Get-ObjectType -Object $myitems0  -ObjectName '$myitems0'
$Array += Get-ObjectType -Object $myitems1  -ObjectName '$myitems1'
$Array += Get-ObjectType -Object $myitems2 -ObjectName '$myitems2'
$Array += Get-ObjectType -Object $InvoiceEntry1 -ObjectName '$InvoiceEntry1'
$Array += Get-ObjectType -Object $InvoiceData1  -ObjectName '$InvoiceData1'
$Array += Get-ObjectType -Object $InvoiceData2  -ObjectName '$InvoiceData2'
$Array += Get-ObjectType -Object $InvoiceData3  -ObjectName '$InvoiceData3'
$Array += Get-ObjectType -Object $InvoiceData4  -ObjectName '$InvoiceData4'
$Array += Get-ObjectType -Object $Object1  -ObjectName '$Object1'
$Array += Get-ObjectType -Object $Object2  -ObjectName '$Object2'
$Array += Get-ObjectType -Object $Object3  -ObjectName '$Object3'
$Array += Get-ObjectType -Object $Object4  -ObjectName '$Object4'
$Array += Get-ObjectType -Object $obj -ObjectName '$obj'
$Array += Get-ObjectType -Object $myArray1 -ObjectName '$myArray1'
$Array += Get-ObjectType -Object $myArray2 -ObjectName '$myArray2'
$Array += Get-ObjectType -Object $InvoiceEntry7 -ObjectName '$InvoiceEntry7'
$Array += Get-ObjectType -Object $InvoiceDataOrdered1 -ObjectName '$InvoiceDataOrdered1'
$Array += Get-ObjectType -Object $InvoiceDataOrdered2 -ObjectName '$InvoiceDataOrdered2'
$Array | Format-Table -AutoSize
#>

Describe 'ConvertTo-Excel - Should deliver same results as Format-Table -Autosize (via pipeline)' {
    It 'Given (MyItems0) should have 3 columns, 4 rows, 3rd row 3rd column should be Food lover' {
        $Type = Get-ObjectType -Object $myitems0
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\17.xlsx"
        $myitems0 | ConvertTo-Excel -Path $Path -AutoFilter -AutoSize
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 4
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Joe'
        $pkg.Workbook.Worksheets[1].Cells['A3'].Value | Should -Be 'Sue'
        $pkg.Workbook.Worksheets[1].Cells['C4'].Value | Should -Be 'Food lover'
        $pkg.Dispose()
    }
    It 'Given (MyItems1) should have 3 columns, 2 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $myitems1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\18.xlsx"
        $myitems1 | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Joe'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'age'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'info'
        $pkg.Dispose()
    }
    It 'Given (MyItems2) should have 3 columns, 2 rows, data should be in proper columns' {
        $Type = Get-ObjectType -Object $MyItems2
        $Type.ObjectTypeName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\1.xlsx"
        $myitems1 | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Joe'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'age'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'info'
        $pkg.Dispose()

    }
    It 'Given (InvoiceEntry1) should have 2 columns, 3 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceEntry1
        $Type.ObjectTypeName | Should -Be 'Hashtable'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be ''
        $Type.ObjectTypeInsiderBaseName | Should -Be ''

        $Path = "$Env:TEMP\2.xlsx"
        $InvoiceEntry1 | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 3
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData1) should have 2 columns, 10 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceData1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Hashtable'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\3.xlsx"
        $InvoiceData1 | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 11
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData2) should have 2 columns, 6 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceData2
        $Type.ObjectTypeName | Should -Be 'Collection`1'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\4.xlsx"
        $InvoiceData2 | ConvertTo-Excel -Path $Path #-NoNumberConversion 'Amount'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 6
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'IT Services 1'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Amount'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be '$200'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData3) should have 2 columns, 3 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceData3
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Hashtable'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\5.xlsx"
        $InvoiceData3 | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 3
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData4) should have 2 columns, 2 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceData4
        $Type.ObjectTypeName | Should -Be 'Collection`1'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\6.xlsx"
        $InvoiceData4 | ConvertTo-Excel -Path $Path #-NoNumberConversion 'Amount'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'IT Services 1'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Amount'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be '$200'
        $pkg.Dispose()

    }
    It 'Given (Object1) should have 3 columns, 6 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $Object1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\7.xlsx"
        $Object1 | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 6
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'ProcessName'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Handle'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'StartTime'
        $pkg.Dispose()


    }

    It 'Given (Object2) should have 10 columns, Have more then 4 rows, data is in random order (unfortunately)' {

        $Type = Get-ObjectType -Object $Object2
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        #$Type.ObjectTypeInsiderName | Should -Be 'PSDriveInfo'
        #$Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\8.xlsx"
        $Object2 | ConvertTo-Excel -Path $Path #-Verbose
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -BeGreaterThan 4
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 10
        # Not sure yet how to predict thje order. Seems order of FT -a is differnt then FL and script takes FL for now
        #$pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'ProcessName'
        #$pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Handle'
        #$pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'StartTime'
        $pkg.Dispose()
    }

    It 'Given (Object3) should have 10 columns, Have more then 1 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $Object3
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\9.xlsx"
        $Object3 | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -BeGreaterThan 1
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 10
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Used'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Free'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'CurrentLocation'
        $pkg.Dispose()
    }

    It 'Given (Object4) should have 10 columns, Have more then 1 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $Object4
        $Type.ObjectTypeName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\10.xlsx"
        $Object4 | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -BeGreaterThan 1
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 10
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Used'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Free'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'CurrentLocation'
        $pkg.Dispose()
    }

    It 'Given (obj) should have 4 columns, Have 2 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $obj
        $Type.ObjectTypeName | Should -Be 'Object'
        $Type.ObjectTypeBaseName | Should -Be $null
        $Type.ObjectTypeInsiderName | Should -Be 'Object'
        $Type.ObjectTypeInsiderBaseName | Should -Be $null

        $Path = "$Env:TEMP\11.xlsx"
        $obj | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 4
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Manufacturer'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $pkg.Dispose()

    }

    It 'Given (myArray1) should have 4 columns, Have 4 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $myArray1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Object'
        $Type.ObjectTypeInsiderBaseName | Should -Be $null

        $Path = "$Env:TEMP\12.xlsx"
        $myArray1 | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 4
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 4
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Manufacturer'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $pkg.Workbook.Worksheets[1].Cells['E1'].Value | Should -Be $null
        $pkg.Dispose()

    }

    It 'Given (myArray2) should have 4 columns, Have 2 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $myArray2
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Object'
        $Type.ObjectTypeInsiderBaseName | Should -Be $null

        $Path = "$Env:TEMP\13.xlsx"
        $myArray2 | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 4
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Manufacturer'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $pkg.Workbook.Worksheets[1].Cells['E1'].Value | Should -Be $null
        $pkg.Dispose()
    }
    #>
    It 'Given (InvoiceEntry7) should have 2 columns, Have 3 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceEntry7
        $Type.ObjectTypeName | Should -Be 'OrderedDictionary'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'String'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'


        $Path = "$Env:TEMP\14.xlsx"
        $InvoiceEntry7 | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 3
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    It 'Given (InvoiceDataOrdered1) should have 2 columns, Have 3 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceDataOrdered1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'OrderedDictionary'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\15.xlsx"
        $InvoiceDataOrdered1 | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 3
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    It 'Given (InvoiceDataOrdered2) should have 2 columns, Have 5 rows, data should be in proper columns' {
        $Type = Get-ObjectType -Object $InvoiceDataOrdered2
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'OrderedDictionary'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\16.xlsx"
        $InvoiceDataOrdered2 | ConvertTo-Excel -Path $Path #-Show
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 5
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    ## Cleanup of tests
    for ($i = 1; $i -le 30; $i++) {
        $Path = "$($i).xlsx"
        Remove-Item $Path -Confirm:$False -ErrorAction SilentlyContinue
    }
}

Describe 'ConvertTo-Excel - Should deliver same results as Format-Table -Autosize (without pipeline)' {
    It 'Given (MyItems0) should have 3 columns, 4 rows, 3rd row 3rd column should be Food lover' {

        $Type = Get-ObjectType -Object $myitems0
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\17.xlsx"
        ConvertTo-Excel -Path $Path -AutoFilter -AutoSize -DataTable $myitems0
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 4
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Joe'
        $pkg.Workbook.Worksheets[1].Cells['A3'].Value | Should -Be 'Sue'
        $pkg.Workbook.Worksheets[1].Cells['C4'].Value | Should -Be 'Food lover'
        $pkg.Dispose()

    }
    It 'Given (MyItems1) should have 3 columns, 2 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $myitems1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\18.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $myitems1
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Joe'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'age'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'info'
        $pkg.Dispose()
    }
    It 'Given (MyItems2) should have 3 columns, 2 rows, data should be in proper columns' {
        $Type = Get-ObjectType -Object $MyItems2
        $Type.ObjectTypeName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\1.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $MyItems2
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Joe'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'age'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'info'
        $pkg.Dispose()

    }
    It 'Given (InvoiceEntry1) should have 2 columns, 3 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceEntry1
        $Type.ObjectTypeName | Should -Be 'Hashtable'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be ''
        $Type.ObjectTypeInsiderBaseName | Should -Be ''

        $Path = "$Env:TEMP\2.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $InvoiceEntry1
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 3
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData1) should have 2 columns, 10 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceData1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Hashtable'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\3.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $InvoiceData1
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 11
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }

    It 'Given (InvoiceData2) should have 2 columns, 6 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceData2
        $Type.ObjectTypeName | Should -Be 'Collection`1'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\4.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $InvoiceData2 #-NoNumberConversion 'Amount'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 6
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'IT Services 1'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Amount'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be '$200'
        $pkg.Dispose()

    }

    It 'Given (InvoiceData3) should have 2 columns, 3 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceData3
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Hashtable'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\5.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $InvoiceData3
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 3
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }

    It 'Given (InvoiceData4) should have 2 columns, 2 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceData4
        $Type.ObjectTypeName | Should -Be 'Collection`1'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\6.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $InvoiceData4 #-NoNumberConversion 'Amount'
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'IT Services 1'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Amount'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be '$200'
        $pkg.Dispose()

    }

    It 'Given (Object1) should have 3 columns, 6 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $Object1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\7.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $Object1
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 6
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'ProcessName'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Handle'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'StartTime'
        $pkg.Dispose()


    }

    It 'Given (Object2) should have 10 columns, Have more then 4 rows, data is in random order (unfortunately)' {

        $Type = Get-ObjectType -Object $Object2
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        #$Type.ObjectTypeInsiderName | Should -Be 'PSDriveInfo'
        #$Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\8.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $Object2
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -BeGreaterThan 4
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 10
        # Not sure yet how to predict thje order. Seems order of FT -a is differnt then FL and script takes FL for now
        #$pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'ProcessName'
        #$pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Handle'
        #$pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'StartTime'
        $pkg.Dispose()
    }

    It 'Given (Object3) should have 10 columns, Have more then 1 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $Object3
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\9.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $Object3
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -BeGreaterThan 1
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 10
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Used'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Free'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'CurrentLocation'
        $pkg.Dispose()
    }

    It 'Given (Object4) should have 10 columns, Have more then 1 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $Object4
        $Type.ObjectTypeName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\10.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $Object4
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -BeGreaterThan 1
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 10
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Used'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Free'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'CurrentLocation'
        $pkg.Dispose()
    }

    It 'Given (obj) should have 4 columns, Have 2 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $obj
        $Type.ObjectTypeName | Should -Be 'Object'
        $Type.ObjectTypeBaseName | Should -Be $null
        $Type.ObjectTypeInsiderName | Should -Be 'Object'
        $Type.ObjectTypeInsiderBaseName | Should -Be $null

        $Path = "$Env:TEMP\11.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $obj
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 4
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Manufacturer'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $pkg.Dispose()

    }

    It 'Given (myArray1) should have 4 columns, Have 4 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $myArray1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Object'
        $Type.ObjectTypeInsiderBaseName | Should -Be $null

        $Path = "$Env:TEMP\12.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $myArray1
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 4
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 4
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Manufacturer'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $pkg.Workbook.Worksheets[1].Cells['E1'].Value | Should -Be $null
        $pkg.Dispose()

    }

    It 'Given (myArray2) should have 4 columns, Have 2 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $myArray2
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Object'
        $Type.ObjectTypeInsiderBaseName | Should -Be $null

        $Path = "$Env:TEMP\13.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $myArray2
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 4
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Manufacturer'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $pkg.Workbook.Worksheets[1].Cells['E1'].Value | Should -Be $null
        $pkg.Dispose()
    }
    #>
    It 'Given (InvoiceEntry7) should have 2 columns, Have 3 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceEntry7
        $Type.ObjectTypeName | Should -Be 'OrderedDictionary'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'String'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'


        $Path = "$Env:TEMP\14.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $InvoiceEntry7
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 3
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    It 'Given (InvoiceDataOrdered1) should have 2 columns, Have 3 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceDataOrdered1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'OrderedDictionary'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\15.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $InvoiceDataOrdered1
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 3
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    It 'Given (InvoiceDataOrdered2) should have 2 columns, Have 5 rows, data should be in proper columns' {
        $Type = Get-ObjectType -Object $InvoiceDataOrdered2
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'OrderedDictionary'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = "$Env:TEMP\16.xlsx"
        ConvertTo-Excel -Path $Path -DataTable $InvoiceDataOrdered2
        $pkg = Get-ExcelDocument -Path $Path
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 5
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }

    ## Cleanup of tests
    for ($i = 1; $i -le 30; $i++) {
        $Path = "$($i).xlsx"
        Remove-Item $Path -Confirm:$False -ErrorAction SilentlyContinue
    }
}