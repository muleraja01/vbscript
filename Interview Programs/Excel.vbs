Dim objExcel, objWorkbook, objWorksheet, RowsCount
Set objExcel = CreateObject(“Excel.Application”)
Set objWorkbook = objExcel.Workbooks.Open(“C:\Users\G C REDDY\Desktop\January.xlsx”)
Set objWorksheet = objWorkbook.Worksheets(1)

RowsCount = objWorksheet.usedRange.Rows.Count

For i = 2 To RowsCount Step 1
SystemUtil.Run “C:\Program Files\HP\Unified Functional Testing\samples\flight\app\flight4a.exe”
Dialog(“Login”).Activate
Dialog(“Login”).WinEdit(“Agent Name:”).Set objWorksheet.Cells(i, “A”) ‘i for Row, A for Column
Dialog(“Login”).WinEdit(“Password:”).Set objWorksheet.Cells(i, 2)’i for Row, 2 for Column
Wait 2
Dialog(“Login”).WinButton(“OK”).Click
Window(“Flight Reservation”).Close
Next
objExcel.Quit
Set objWorksheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing