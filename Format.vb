Sub FormatExcel(ByVal xlApp As Object)
    With xlApp
        .ScreenUpdating = True
        .ActiveWindow.Zoom = 90
        
        With .ActiveSheet.UsedRange.Rows(1)
            .Font.Bold = True
            .Interior.ThemeColor = xlThemeColorAccent6
            .Interior.TintAndShade = 0.51298390938963
        End With
    
        With .Cells
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Font.Name = "Arial"
            .Font.Size = 8
            .RowHeight = 34.5
        End With
        
        With .ActiveSheet.UsedRange
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
    
        Dim column As Range
        For Each column In .ActiveSheet.UsedRange.Rows(1).Cells
            column.EntireColumn.AutoFit
        If column.EntireColumn.ColumnWidth > 40 Then _
            column.EntireColumn.ColumnWidth = 40
            column.EntireColumn.WrapText = True
        Next column
        
        .Range("A1").Select
    
        .ScreenUpdating = True
    End With
End Sub

Sub Format()
    Call FormatExcel(Application)
End Sub
