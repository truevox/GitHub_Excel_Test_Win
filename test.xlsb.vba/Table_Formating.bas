    Option Explicit

Sub FormatSTDAnalysisTable()


'Variable TRows = Number of Rows in the Table
'Variable TCols = Number of Columns in the Table


    Dim TRows As Integer
    Dim TCols As Integer
    Dim FinalRow As Integer

    TRows = ActiveCell.CurrentRegion.Rows.Count
    TCols = ActiveCell.CurrentRegion.Columns.Count

    Selection.CurrentRegion.Select
    FinalRow = ActiveCell.Row + TRows - 1


'Selects the upper left corner of the table and formats the top row then the bottom row

    Selection.CurrentRegion.Select

    ActiveCell.Resize(1, 7).Select

    With Selection
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Interior.Color = RGB(221, 235, 247)
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With


    ActiveCell.Offset(TRows - 1, 0).Select

    With ActiveCell.Resize(1, 7)
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Interior.Color = RGB(221, 235, 247)
        .Font.Bold = True
    End With

'Formats the data cells in the columns


    Selection.CurrentRegion.Select
    ActiveCell.Offset(1, 1).Resize(TRows - 1, 3).Select
    Selection.NumberFormat = "#,##0"

    Selection.CurrentRegion.Select
    ActiveCell.Offset(1, 4).Resize(TRows - 1, 1).Select
    Selection.NumberFormat = "$#,##0_);($#,##0)"

    Selection.CurrentRegion.Select
    ActiveCell.Offset(1, 5).Resize(TRows - 1, 1).Select
    Selection.NumberFormat = "0.00%"

    Selection.CurrentRegion.Select
    ActiveCell.Offset(1, 6).Resize(TRows - 1, 1).Select
    Selection.NumberFormat = "$#,##0_);($#,##0)"

'Labels the headings

    ActiveCell.Offset(-1, 0).Value = "Average Gift"
    ActiveCell.Offset(-1, -1).Value = "Response Rate"
    ActiveCell.Offset(-1, -2).Value = "Gift Amount"
    ActiveCell.Offset(-1, -3).Value = "Number of Gifts"
    ActiveCell.Offset(-1, -4).Value = "Number of Last Gifts"
    ActiveCell.Offset(-1, -5).Value = "Number Mailed"

'Adds formulas to the last two columns

    Selection.CurrentRegion.Select
    ActiveCell.Offset(1, 5).Resize(TRows - 1, 1). _
        FormulaR1C1 = "=IF(OR(RC[-2]=0,RC[-4]=0),"""",RC[-2]/RC[-4])"
    ActiveCell.Offset(1, 6).Resize(TRows - 1, 1). _
        FormulaR1C1 = "=IF(OR(RC[-2]=0,RC[-3]=0),"""",RC[-2]/RC[-3])"


End Sub
Sub FormatCOST_RAISE_DOLLARAnalysisTable()

  
    Dim TRows As Integer
    Dim TCols As Integer
    Dim FinalRow As Integer
   
    
    TRows = ActiveCell.CurrentRegion.Rows.Count
    TCols = ActiveCell.CurrentRegion.Columns.Count
         
    Selection.CurrentRegion.Select
    FinalRow = ActiveCell.Row + TRows - 1
    
    MsgBox "If not already done so, delete the Number of Last Gifts column and rerun the macro. Starting columns 2 to 4 should be Number Mailed, Number of Gits, Gift Amount"

'Selects the upper left corner of the table and formats the top row then the bottom row

    Selection.CurrentRegion.Select
      
    ActiveCell.Resize(1, 7).Select
    
    With Selection
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Interior.Color = RGB(221, 235, 247)
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
    End With
         

    ActiveCell.Offset(TRows - 1, 0).Select
            
    With ActiveCell.Resize(1, 7)
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Interior.Color = RGB(221, 235, 247)
        .Font.Bold = True
    End With

'Formats the data cells in the columns


    Selection.CurrentRegion.Select
    ActiveCell.Offset(1, 1).Resize(TRows - 1, 2).Select
    Selection.NumberFormat = "#,##0"
    
    Selection.CurrentRegion.Select
    ActiveCell.Offset(1, 3).Resize(TRows - 1, 2).Select
    Selection.NumberFormat = "$#,##0_);($#,##0)"
        
    Selection.CurrentRegion.Select
    ActiveCell.Offset(1, 5).Resize(TRows - 1, 2).Select
    Selection.NumberFormat = "$#,##0.00_);($#,##0.00)"
    
'Labels the headings

    ActiveCell.Offset(-1, -5).Select
    ActiveCell.Offset(, 1).Value = "Number Mailed"
    ActiveCell.Offset(, 2).Value = "Number of Gifts"
    ActiveCell.Offset(, 3).Value = "Gift Amount"
    ActiveCell.Offset(, 4).Value = "Segment Cost"
    ActiveCell.Offset(, 5).Value = "Cost Each"
    ActiveCell.Offset(, 6).Value = "Cost to Raise A $"
    
'Adds formulas to the last three columns
   
    Selection.CurrentRegion.Select
        
    ActiveCell.Offset(1, 4).Resize(TRows - 2, 1). _
         FormulaR1C1 = "=if(RC[-3]=0,"""",R" & FinalRow & "C * RC[-3] / R" & FinalRow & "C[-3])"
    ActiveCell.Offset(1, 5).Resize(TRows - 1, 1). _
         FormulaR1C1 = "=if(OR(RC[-1]=0,RC[-4]=0),"""",RC[-1] / RC[-4])"
    ActiveCell.Offset(1, 6).Resize(TRows - 1, 1). _
         FormulaR1C1 = "=if(OR(RC[-2]=0,RC[-3]=0),"""",RC[-2] / RC[-3])"

    MsgBox "Enter the total cost in the Totals row of the Segment Cost column. If a segment(s) had special costs, for example a rental list, so that costs are not distributed proportionaly, overwirte the costs manually."
    
End Sub



Sub ColumnWidths18_11()
'run this from the cell in the upper left of the first table, row 1

    ActiveCell.Range("A:A").ColumnWidth = 18
    ActiveCell.Range("B:G").ColumnWidth = 11

End Sub



