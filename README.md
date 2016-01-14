ExcelHelper
===========

Library to help Excel VBA programming

Examples
-------

The following examples demonstrate using the Google Maps API to get directions between two locations.

### GetJSON Example
```VB.net
Const DIRECTORY_PATH As String = "C:\MY_PATH"
Const BD_SHEET_NAME As String = "BD"

Sub LoadFiles()
  Dim path As String, file As New OCFile, bdSheet
  
  Set bdSheet = Sheets(BD_SHEET_NAME)
  
  path = StringFormat("{0}\ANOTHER_PATH\{1}_PORTFOLIO.xls", DIRECTORY_PATH, Format(fileDate, "yyyyMMdd"))
  file.OpenNewFile (path)
  
  Call ClearFilter(file.newWorkbook.ActiveSheet)

  'Example 1

  With bdSheet
      file.newWorkbook.ActiveSheet.UsedRange.Copy bdSheet.[A1]
      ....
  End With
  
  'Example 2
  
  With file.newWorkbook.ActiveSheet.UsedRange
      .Columns(3).Copy bdSheet.[A2]
      .Columns(2).Copy bdSheet.[B2]
      .Columns(6).Copy bdSheet.[C2]
      .Columns(1).Copy bdSheet.[D2]
      .Columns(7).Copy bdSheet.[E2]
      .Columns(5).Copy bdSheet.[F2]
      .Columns(8).Copy bdSheet.[G2]
      .Columns(9).Copy bdSheet.[H2]
      .Columns(13).Copy bdSheet.[I2]
      
      colIndex = WorksheetFunction.Match("Price", .Rows(1), 0)
      .Columns(colIndex).Copy bdSheet.[K2]
  End With
  
  file.CloseFile
End sub
```
