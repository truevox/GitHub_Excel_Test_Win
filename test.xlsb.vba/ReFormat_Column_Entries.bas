Sub ConvertColumntoGeneral_Currency()

'Run from any cell in the column but the column or column heading must start in row 1

'Converts numbers as text to either General or Currency depending on the celentry
' Selects all cells in the current column from the second cell, including blank cells, down to the last non-empty cell
' In an Excel table selects all cells down to the last row in the table, whether the last cell is blank or not

' The FieldInfo parameter Array(1,1) specifies that the 1st column in the selection will be parsed as General (or Currency if $ was include) (see XLColumnDataType web page

' Keyboard Shortcut: Ctrl+Shift+N
    Dim MyCol As Integer
    MyCol = ActiveCell.Column
    FinalRow = Cells(Rows.Count, MyCol).End(xlUp).Row
    Cells(2, MyCol).Resize(FinalRow - 1, 1).Select
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Cells(2, MyCol).Select
End Sub
Sub ConvertColumntoDate()

'Run from any cell in the column but the column or column heading must start in row 1

'Converts dates as text to Microsoft date format
' Selects all cells in the current column from the second cell, including blank cells, down to the last non-empty cell
' In an Excel table selects all cells down to the last row in the table, whether the last cell is blank or not

' The FieldInfo parameter Array(1,3) specifies that the 1st column in the selection will be parsed as MDY (see XLColumnDataType web page)
' It converts blank cells to General and may fix incorrect date formats, allowing grouping in pivot tables
' If a pivot table still won't group, sort for erroneous entries in the column, fix them, then refresh twice (CTRL ALT F5)
' If that doesn't work you have to delect the pivot table and recreate it.
' If you need individual dates in the pivot table rows but can only get grouped months, covert blank cells from General to Date


'Keyboard Shortcut: Ctrl+Shift+D

    Dim MyCol As Integer
    MyCol = ActiveCell.Column
    FinalRow = Cells(Rows.Count, MyCol).End(xlUp).Row
    Cells(2, MyCol).Resize(FinalRow - 1, 1).Select
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 3), TrailingMinusNumbers:=True
    Cells(2, MyCol).Select
End Sub
